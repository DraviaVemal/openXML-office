use crate::{
    file_handling::{compress_content, decompress_content},
    files::{SqliteDatabases, XmlDeSerializer, XmlDocument, XmlSerializer},
    get_all_queries,
    global_2007::parts::ContentTypesPart,
};
use anyhow::{anyhow, Context, Error as AnyError, Result as AnyResult};
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
pub(crate) struct OfficeDocument {
    sqlite_database: SqliteDatabases,
    /// Key : File_Name -> Value : (File Handle, Content Type, File Extension, Extension Type)
    xml_document_collection: HashMap<
        String,
        (
            Rc<RefCell<XmlDocument>>,
            Option<String>,
            Option<String>,
            Option<String>,
        ),
    >,
    queries: HashMap<String, String>,
}

impl OfficeDocument {
    /// Create or Clone existing document to start with
    pub(crate) fn new(file_path: Option<String>, is_in_memory: bool) -> AnyResult<Self, AnyError> {
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
    pub(crate) fn get_connection(&self) -> &SqliteDatabases {
        &self.sqlite_database
    }

    pub(crate) fn check_file_exist(&self, file_path: String) -> AnyResult<bool, AnyError> {
        if let Some(count) = self
            .get_connection()
            .get_count(
                self.queries
                    .get("count_archive_content")
                    .ok_or(anyhow!("Reading Query Failed"))?,
                params![file_path],
            )
            .context("Get count DB Execution Failed")?
        {
            Ok(count > 0)
        } else {
            Ok(false)
        }
    }

    pub(crate) fn delete_document_mut(&mut self, file_name: &str) -> AnyResult<(), AnyError> {
        let file_path = if file_name.starts_with("/") {
            file_name.strip_prefix("/").unwrap()
        } else {
            file_name
        };
        self.xml_document_collection.remove(file_path);
        // Delete DB Data
        let query = self
            .queries
            .get("delete_archive_content")
            .unwrap()
            .to_owned();
        self.get_connection()
            .delete_record(&query, params![file_path])
            .context("Delete File Record Failed")?;
        Ok(())
    }

    pub(crate) fn get_xml_document_ref(
        &mut self,
        file_name: &str,
        content_type: Option<String>,
        file_extension: Option<String>,
        extension_type: Option<String>,
        xml_tree: XmlDocument,
    ) -> AnyResult<Weak<RefCell<XmlDocument>>, AnyError> {
        let ref_xml_document = Rc::new(RefCell::new(xml_tree));
        let weak_xml_document = Rc::downgrade(&ref_xml_document);
        // TODO : Did as quick fix to maintain the order. Refactor on next pass
        self.xml_document_collection.insert(
            file_name.to_string(),
            (
                ref_xml_document.clone(),
                content_type.clone(),
                file_extension.clone(),
                extension_type.clone(),
            ),
        );
        let mut vec: Vec<u8> = Vec::new();
        Self::insert_update_archive_record(
            &self.sqlite_database,
            &self.queries,
            file_name,
            file_extension,
            extension_type,
            content_type,
            &mut vec,
        )
        .context("Create or update archive DB record Failed")?;
        Ok(weak_xml_document)
    }

    /// Get Xml tree from content if not already open
    pub(crate) fn get_xml_tree(
        &self,
        file_path: &str,
    ) -> AnyResult<Option<(XmlDocument, Option<String>, Option<String>, Option<String>)>, AnyError>
    {
        // Check is the file path object already exist
        if self.xml_document_collection.get(file_path).is_some() {
            return Err(anyhow!(
                "Please close the Existing object before creating new handle"
            ));
        }
        let query = self
            .queries
            .get("select_archive_content")
            .ok_or_else(|| anyhow!("Expected Query Not Found"))?;
        fn row_mapper(
            row: &Row,
        ) -> AnyResult<(Vec<u8>, Option<String>, Option<String>, Option<String>), Error> {
            Result::Ok((row.get(0)?, row.get(1)?, row.get(2)?, row.get(3)?))
        }
        let result = self
            .get_connection()
            .find_one(&query, params![file_path], row_mapper)
            .map_err(|e| anyhow!("Failed to execute the Find one Query : {}", e))?;
        if let Some((filet_content, content_type, file_extension, extension_type)) = result {
            let decompressed_data: Vec<u8> =
                decompress_content(&filet_content).context("Raw Content Decompression Failed")?;
            let xml_tree: XmlDocument = XmlSerializer::vec_to_xml_doc_tree(decompressed_data)
                .context("Xml Serializer Failed")?;
            Ok(Some((
                xml_tree,
                content_type,
                file_extension,
                extension_type,
            )))
        } else {
            Ok(None)
        }
    }

    /// Update the XML tree data to database and close the refCell
    pub(crate) fn close_xml_document(&mut self, file_path: &str) -> AnyResult<(), AnyError> {
        let xml_tree_option = self.xml_document_collection.remove(file_path);
        if let Some((xml_tree, content_type, file_extension, extension_type)) = xml_tree_option {
            let mut xml_document = xml_tree
                .try_borrow_mut()
                .context("Failed to get document handle")?;
            let mut uncompressed_data = XmlDeSerializer::xml_tree_to_vec(&mut xml_document)
                .context("Xml Tree to String content")?;
            Self::insert_update_archive_record(
                &self.sqlite_database,
                &self.queries,
                file_path,
                file_extension,
                extension_type,
                content_type,
                &mut uncompressed_data,
            )
            .context("Create or update archive DB record Failed")?;
        }
        Ok(())
    }

    /// Save Current Document to final result
    pub(crate) fn save_as(&mut self, file_path: &str) -> AnyResult<(), AnyError> {
        // Save the live content update object to database
        let keys = self
            .xml_document_collection
            .keys()
            .cloned()
            .collect::<Vec<String>>();
        for key_file_path in keys {
            self.close_xml_document(&key_file_path)
                .context(" Saving open object content failed")?;
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
        let archive_create = queries
            .get("create_archive_table")
            .ok_or_else(|| anyhow!("Expected Query Not Found"))?;
        sqlite_databases
            .create_table(&archive_create)
            .context("Initialize Database Failed for Office Document")
    }

    /// Save the database content into file archive
    fn save_database_into_archive(&self) -> AnyResult<Vec<u8>, AnyError> {
        let extensions: Vec<(String, String)>;
        let mut overrides: Vec<(String, String)> = Vec::new();
        let mut buffer = Cursor::new(Vec::new());
        let mut zip_writer: ZipWriter<&mut Cursor<Vec<u8>>> = ZipWriter::new(&mut buffer);
        let zip_option = SimpleFileOptions::default().compression_level(Some(4));
        // Load Files into Archive and add Override for content types
        {
            let query = self
                .queries
                .get("select_archive_files")
                .ok_or_else(|| anyhow!("Expected Query Not Found"))?;
            fn row_mapper(
                row: &Row,
            ) -> AnyResult<(String, Option<String>, Option<Vec<u8>>), Error> {
                Ok((row.get(0)?, row.get(1)?, row.get(2)?))
            }
            let files = self
                .sqlite_database
                .find_many(&query, params![], row_mapper)
                .context("Query Get Many Failed")?;
            for (file_name, content_type, file_content) in files {
                if let Some(content_type) = content_type {
                    overrides.push((format!("/{}", file_name), content_type));
                }
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
        }
        // Load extensions and prepare for content type
        {
            let query = self
                .queries
                .get("select_extensions")
                .ok_or_else(|| anyhow!("Expected Query Not Found"))?;
            fn row_mapper(row: &Row) -> AnyResult<(String, String), Error> {
                Ok((row.get(0)?, row.get(1)?))
            }
            let rows = self
                .sqlite_database
                .find_many(&query, params![], row_mapper)
                .context("Query Get Many Failed")?;
            extensions = rows
                .iter()
                .map(|(file_extension, content_type)| {
                    (file_extension.to_string(), content_type.to_string())
                })
                .collect();
        }
        // Insert Content Type Details into Archive
        zip_writer
            .start_file("[Content_Types].xml", zip_option)
            .context("Zip File Write Start Fail")?;
        let content_type_file: Vec<u8> = ContentTypesPart::create_xml_file(extensions, overrides)
            .context("Creating Content Type XML Failed")?;
        zip_writer
            .write_all(&content_type_file)
            .context("Writing compressed data to ZIp")?;
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
        let mut uncompressed_file = Vec::new();
        zip_read
            .by_name("[Content_Types].xml")
            .context("Read [Content_Types].xml Failed")?
            .read_to_end(&mut uncompressed_file)
            .context("Failed to uncompress [Content_Types].xml")?;
        let file_names = zip_read
            .file_names()
            .filter(|name| "[Content_Types].xml" != *name)
            .map(|item| item.to_string())
            .collect::<Vec<_>>();
        let mut content_types_part =
            ContentTypesPart::new(uncompressed_file).context("Decoding Content Type Failed")?;
        let extension_collection = content_types_part
            .get_extensions()
            .context("Failed to pull extensions list")?;
        // Load File Content To DB
        for file_name in file_names {
            let mut file_extension = None;
            let mut extension_type = None;
            let mut uncompressed_data = Vec::new();
            let f_extension = file_name.rsplit(".").next();
            if let Some(f_extension) = f_extension {
                if let Some(extensions) = extension_collection.clone() {
                    let item = extensions.iter().find(|item| item.0 == f_extension);
                    if let Some(extension) = item {
                        file_extension = Some(extension.0.to_string());
                        extension_type = Some(extension.1.to_string());
                    }
                }
            }
            zip_read
                .by_name(&file_name)
                .context("Zip file extract failed")?
                .read_to_end(&mut uncompressed_data)
                .context("File Uncompressed failed")?;
            let content_type = content_types_part
                .get_override_content_type(&file_name)
                .context("Failed to extract Content Type")?;
            Self::insert_update_archive_record(
                sqlite_database,
                queries,
                &file_name,
                file_extension,
                extension_type,
                content_type,
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
        file_extension: Option<String>,
        extension_type: Option<String>,
        content_type: Option<String>,
        uncompressed_data: &mut Vec<u8>,
    ) -> AnyResult<(), AnyError> {
        let insert_archive_query = queries
            .get("insert_archive_table")
            .ok_or_else(|| anyhow!("Expected Query Not Found"))?;
        let compression_level = 4;
        let compressed = compress_content(&uncompressed_data, compression_level)
            .context("Recompressing in GZip Failed")?;
        sqlite_database
            .insert_record(
                &insert_archive_query,
                params![
                    file_path,
                    file_extension,
                    extension_type,
                    content_type,
                    compressed.len(),
                    uncompressed_data.len(),
                    compression_level,
                    compressed
                ],
            )
            .context("Archive content load failed.")?;
        Ok(())
    }
}
