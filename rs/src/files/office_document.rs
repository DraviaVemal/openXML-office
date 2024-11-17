use std::{
    fs::{metadata, remove_file, File},
    io::{Cursor, Read, Write},
    ptr::null,
};

use crate::{
    file_handling::{compress_content, decompress_content},
    files::SqliteDatabases,
    get_specific_queries,
};
use anyhow::{anyhow, Context, Error as AnyError, Result as AnyResult};
use rusqlite::{params, Row};
use zip::{write::SimpleFileOptions, ZipArchive, ZipWriter};

#[derive(Debug)]
pub struct ArchiveRecordModel {
    id: i16,
    file_name: String,
    content_type: String,
    compressed_file_size: i32,
    uncompressed_file_size: i32,
    compression_level: i8,
    compression_type: String,
    file_content: Vec<u8>,
    tree_content: Vec<u8>,
}

#[derive(Debug)]
pub struct OfficeDocument {
    sqlite_database: SqliteDatabases,
}

impl OfficeDocument {
    /// Create or Clone existing document to start with
    pub fn new(file_path: Option<String>, is_in_memory: bool) -> AnyResult<Self, AnyError> {
        let sqlite_database =
            SqliteDatabases::new(is_in_memory).context("Sqlite database initialization failed")?;
        Self::initialize_database(&sqlite_database).context("Initialize Database Failed")?;
        if let Some(file_path) = file_path {
            // Load existing file to our system
            Self::load_archive_into_database(&sqlite_database, &file_path)
                .context("Load OpenXML Archive Into Database Failed")?;
        } else { // Initialize new document
        }
        Ok(Self { sqlite_database })
    }

    /// Get DB Options
    pub fn get_connection(&self) -> &SqliteDatabases {
        &self.sqlite_database
    }

    /// Save Current Document to final result
    pub fn save_as(&self, file_path: &str) -> AnyResult<(), AnyError> {
        let file_content = self
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
    fn initialize_database(sqlite_databases: &SqliteDatabases) -> AnyResult<usize, AnyError> {
        let query: String = get_specific_queries!("open_xml_archive.sql", "create_archive_table")
            .map_err(|e| anyhow!("Query Macro Error : {}", e))?;
        sqlite_databases
            .create_table(&query)
            .context("Initialize Database Failed for Office Document")
    }

    /// Save the database content into file archive
    fn save_database_into_archive(&self) -> AnyResult<Vec<u8>, AnyError> {
        let query: String =
            get_specific_queries!("open_xml_archive.sql", "select_all_archive_rows")
                .map_err(|e| anyhow!("Query Macro Error : {}", e))?;
        fn row_mapper(row: &Row) -> Result<ArchiveRecordModel, rusqlite::Error> {
            Result::Ok(ArchiveRecordModel {
                id: row.get(0)?,
                file_name: row.get(1)?,
                content_type: row.get(2)?,
                compressed_file_size: row.get(3)?,
                uncompressed_file_size: row.get(4)?,
                compression_level: row.get(5)?,
                compression_type: row.get(6)?,
                file_content: row.get(7)?,
                tree_content: row.get(8)?,
            })
        }
        let result = self
            .sqlite_database
            .find_many(&query, params![], row_mapper)
            .context("Query Get Many Failed")?;
        let mut buffer = Cursor::new(Vec::new());
        let mut zip_writer: ZipWriter<&mut Cursor<Vec<u8>>> = ZipWriter::new(&mut buffer);
        let zip_option = SimpleFileOptions::default();
        for ArchiveRecordModel {
            id: _,
            file_name,
            content_type: _,
            compressed_file_size: _,
            uncompressed_file_size: _,
            compression_level: _,
            compression_type: _,
            file_content,
            tree_content: _,
        } in result
        {
            // TODO : Create New File Content for the Tree View
            let uncompressed = decompress_content(&file_content).context("Decompress Error")?;
            zip_writer
                .start_file(file_name, zip_option)
                .context("Zip File Write Start Fail")?;
            zip_writer
                .write_all(&uncompressed)
                .context("Writing compressed data to ZIp")?;
        }
        zip_writer.finish().context("Zip Close Failed")?;
        Ok(buffer.into_inner())
    }

    /// Read Zip file and load it into database after compression
    fn load_archive_into_database(
        &sqlite_database: &SqliteDatabases,
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
            let compressed =
                compress_content(&uncompressed_data).context("Recompressing in GZip Failed")?;
            let query = get_specific_queries!("open_xml_archive.sql", "insert_archive_table")
                .map_err(|e| anyhow!("Specific query pull fail. {}", e))?;
            sqlite_database
                .insert_record(
                    &query,
                    params![
                        file_content.name(),
                        "todo",
                        compressed.len(),
                        uncompressed_data.len(),
                        1,
                        "gzip",
                        compressed,
                        null()
                    ],
                )
                .context("Archive content load failed.")?;
        }
        Ok(())
    }
}
