use crate::{
    file_handling::{compress_content, decompress_content},
    files::{SqliteDatabases, XmlDeSerializer, XmlDocument, XmlSerializer},
    get_all_queries,
};
use anyhow::{anyhow, Context, Error as AnyError, Ok, Result as AnyResult};
use rusqlite::{params, Error, Row};
use std::{
    cell::RefCell,
    collections::HashMap,
    fs::{metadata, remove_file, File},
    io::{Cursor, Read, Write},
    rc::{Rc, Weak},
};
use zip::{write::SimpleFileOptions, ZipArchive, ZipWriter};

#[derive(Debug)]
pub struct ArchiveRecordModel {
    file_name: String,
    content_type: String,
    compressed_xml_file_size: i32,
    uncompressed_xml_file_size: i32,
    compression_level: i8,
    compression_type: String,
    file_content: Option<Vec<u8>>,
}

#[derive(Debug)]
pub struct ArchiveRecordFileModel {
    file_content: Option<Vec<u8>>,
}

#[derive(Debug)]
pub struct OfficeDocument {
    sqlite_database: SqliteDatabases,
    xml_document_collection: HashMap<String, Rc<RefCell<XmlDocument>>>,
    queries: HashMap<String, String>,
}

impl OfficeDocument {
    /// Create or Clone existing document to start with
    pub fn new(file_path: Option<String>, is_in_memory: bool) -> AnyResult<Self, AnyError> {
        let sqlite_database: SqliteDatabases =
            SqliteDatabases::new(is_in_memory).context("Sqlite database initialization failed")?;
        let queries = get_all_queries!("office_document.sql");
        Self::initialize_database(&sqlite_database, &queries)
            .context("Initialize Database Failed")?;
        if let Some(file_path) = file_path {
            // Load existing file to our system
            Self::load_archive_into_database(&sqlite_database, &queries, &file_path)
                .context("Load OpenXML Archive Into Database Failed")?;
        }
        Ok(Self {
            sqlite_database,
            xml_document_collection: HashMap::new(),
            queries,
        })
    }

    /// Get DB Options
    pub fn get_connection(&self) -> &SqliteDatabases {
        &self.sqlite_database
    }

    pub fn check_file_exist(&self, file_path: String) -> AnyResult<bool, AnyError> {
        if let Some(count) = self
            .get_connection()
            .get_count("count_archive_content", params![file_path])
            .context("Get count DB Execution Failed")?
        {
            Ok(count > 0)
        } else {
            Ok(false)
        }
    }

    pub fn get_xml_document_ref(
        &mut self,
        file_name: &str,
        xml_tree: XmlDocument,
    ) -> Weak<RefCell<XmlDocument>> {
        let ref_xml_document = Rc::new(RefCell::new(xml_tree));
        let weak_xml_document = Rc::downgrade(&ref_xml_document);
        self.xml_document_collection
            .insert(file_name.to_string(), ref_xml_document);
        weak_xml_document
    }

    /// Get Xml tree from content
    pub fn get_xml_tree(&self, file_path: &str) -> AnyResult<Option<XmlDocument>, AnyError> {
        let query = self
            .queries
            .get("select_archive_content")
            .ok_or_else(|| anyhow!("Expected Query Not Found"))?;
        fn row_mapper(row: &Row) -> AnyResult<ArchiveRecordFileModel, Error> {
            Result::Ok(ArchiveRecordFileModel {
                file_content: row.get(0)?,
            })
        }
        let result: Option<ArchiveRecordFileModel> = self
            .get_connection()
            .find_one(&query, params![file_path], row_mapper)
            .map_err(|e| anyhow!("Failed to execute the Find one Query : {}", e))?;
        if let Some(query_result) = result {
            let decompressed_data: Vec<u8> =
                decompress_content(&query_result.file_content.unwrap())
                    .context("Raw Content Decompression Failed")?;
            let xml_tree: XmlDocument = XmlSerializer::vec_to_xml_doc_tree(decompressed_data)
                .context("Xml Serializer Failed")?;
            Ok(Some(xml_tree))
        } else {
            Ok(None)
        }
    }

    /// Update the XML tree data to database and close the refCell
    pub fn close_xml_document(&mut self, file_path: &str) -> AnyResult<(), AnyError> {
        let xml_tree_option = self.xml_document_collection.remove(file_path);
        if let Some(xml_tree) = xml_tree_option {
            let mut xml_document = xml_tree.borrow_mut();
            let mut uncompressed_data = XmlDeSerializer::xml_tree_to_vec(&mut xml_document)
                .context("Xml Tree to String content")?;
            Self::insert_update_archive_record(
                &self.sqlite_database,
                &self.queries,
                file_path,
                &mut uncompressed_data,
            )
            .context("Create or update archive DB record Failed")?;
        }
        Ok(())
    }

    /// Save Current Document to final result
    pub fn save_as(&mut self, file_path: &str) -> AnyResult<(), AnyError> {
        // Save the live content update object to database
        let keys = self
            .xml_document_collection
            .keys()
            .cloned()
            .collect::<Vec<String>>();
        for key_file_path in keys {
            self.close_xml_document(&key_file_path)
                .context("Saving open object content failed")?;
        }
        let file_content: Vec<u8> = self
            .save_database_into_archive()
            .context("Save Archive Data into Database")?;
        if metadata(file_path).is_ok() {
            remove_file(file_path).map_err(|e| anyhow!("Remove Save File Target Failed. {}", e))?;
        }
        let mut file = File::create(file_path).context("Create Save File Failed")?;
        file.write_all(&file_content)
            .context("Save File Write Failed")
    }

    /// Initialize Local archive Database
    fn initialize_database(
        sqlite_databases: &SqliteDatabases,
        queries: &HashMap<String, String>,
    ) -> AnyResult<usize, AnyError> {
        let query = queries
            .get("create_archive_table")
            .ok_or_else(|| anyhow!("Expected Query Not Found"))?;
        sqlite_databases
            .create_table(&query)
            .context("Initialize Database Failed for Office Document")
    }

    /// Save the database content into file archive
    fn save_database_into_archive(&self) -> AnyResult<Vec<u8>, AnyError> {
        let query = self
            .queries
            .get("select_all_archive_rows")
            .ok_or_else(|| anyhow!("Expected Query Not Found"))?;
        let mut buffer = Cursor::new(Vec::new());
        fn row_mapper(row: &Row) -> Result<ArchiveRecordModel, Error> {
            Result::Ok(ArchiveRecordModel {
                file_name: row.get(0)?,
                content_type: row.get(1)?,
                compressed_xml_file_size: row.get(2)?,
                uncompressed_xml_file_size: row.get(3)?,
                compression_level: row.get(4)?,
                compression_type: row.get(5)?,
                file_content: row.get(6)?,
            })
        }
        let result = self
            .sqlite_database
            .find_many(&query, params![], row_mapper)
            .context("Query Get Many Failed")?;
        let mut zip_writer: ZipWriter<&mut Cursor<Vec<u8>>> = ZipWriter::new(&mut buffer);
        let zip_option = SimpleFileOptions::default().compression_level(Some(4));
        for ArchiveRecordModel {
            file_name,
            content_type: _,
            compressed_xml_file_size: _,
            uncompressed_xml_file_size: _,
            compression_level: _,
            compression_type: _,
            file_content,
        } in result
        {
            zip_writer
                .start_file(file_name, zip_option)
                .context("Zip File Write Start Fail")?;
            if let Some(xml_content_compressed) = file_content {
                let uncompressed =
                    decompress_content(&xml_content_compressed).context("Decompress Error")?;
                zip_writer
                    .write_all(&uncompressed)
                    .context("Writing compressed data to ZIp")?;
            }
        }
        zip_writer.finish().context("Zip Close Failed")?;
        Ok(buffer.into_inner())
    }

    /// Read Zip file and load it into database after compression
    fn load_archive_into_database(
        sqlite_database: &SqliteDatabases,
        queries: &HashMap<String, String>,
        file_path: &str,
    ) -> AnyResult<(), AnyError> {
        let file: File = File::open(file_path).context("Open Existing archive File")?;
        let mut zip_read: ZipArchive<File> =
            ZipArchive::new(file).context("Archive read Failed")?;
        let file_count: usize = zip_read.len();
        for i in 0..file_count {
            let mut file_content = zip_read.by_index(i).context("Zip file extract failed")?;
            let mut uncompressed_data = Vec::new();
            file_content
                .read_to_end(&mut uncompressed_data)
                .context("File Uncompressed failed")?;
            Self::insert_update_archive_record(
                sqlite_database,
                queries,
                file_content.name(),
                &mut uncompressed_data,
            )
            .context("Create or update archive DB record Failed")?;
        }
        Ok(())
    }

    fn insert_update_archive_record(
        sqlite_database: &SqliteDatabases,
        queries: &HashMap<String, String>,
        file_path: &str,
        uncompressed_data: &mut Vec<u8>,
    ) -> AnyResult<(), AnyError> {
        let query = queries
            .get("insert_archive_table")
            .ok_or_else(|| anyhow!("Expected Query Not Found"))?;
        let compression_level = 4;
        let compressed = compress_content(&uncompressed_data, compression_level)
            .context("Recompressing in GZip Failed")?;
        sqlite_database
            .insert_record(
                &query,
                params![
                    file_path,
                    "todo",
                    compressed.len(),
                    uncompressed_data.len(),
                    compression_level,
                    "gzip",
                    compressed
                ],
            )
            .context("Archive content load failed.")?;
        Ok(())
    }
}
