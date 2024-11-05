use crate::{
    file_handling::compress_content, get_specific_queries, ArchiveRecordModel, OpenXmlFile,
};
use anyhow::{Error as AnyError, Result as AnyResult};
use rusqlite::{params, Connection, Row, ToSql};
use std::{
    fs::{metadata, remove_file, File},
    io::Read,
};
use zip::ZipArchive;

impl OpenXmlFile {
    /// Create Current file helper object from exiting source
    pub fn open(file_path: &str, is_editable: bool, is_in_memory: bool) -> Self {
        let archive_db: Connection = Self::common_initialization(is_in_memory);
        Self::load_archive_into_database(&archive_db, file_path);
        // Create a clone copy of master file to work with code
        Self {
            is_editable,
            archive_db,
        }
    }

    /// Create Current file helper object a new file to work with
    pub fn create(is_in_memory: bool) -> Self {
        let archive_db: Connection = Self::common_initialization(is_in_memory);
        // Default List of files common for all types
        Self::initialize_database(&archive_db);
        Self {
            is_editable: true,
            archive_db,
        }
    }

    /// Find one result and map the row to a specific type using the closure.
    pub fn find_one<F, T>(
        &self,
        query: &str,
        params: &[&(dyn ToSql)],
        row_mapper: F,
    ) -> AnyResult<Option<T>, AnyError>
    where
        F: Fn(&Row) -> AnyResult<T, rusqlite::Error>,
    {
        let mut stmt = self
            .archive_db
            .prepare(query)
            .expect("Failed to prepare query");
        return match stmt.query_row(params, row_mapper) {
            Ok(result) => Ok(Some(result)),
            Err(rusqlite::Error::QueryReturnedNoRows) => Ok(None),
            Err(e) => Err(e.into()),
        };
    }

    /// Find multiple results and map each row to a specific type using the closure.
    pub fn find_many<F, T>(
        &self,
        query: &str,
        params: &[&(dyn ToSql)],
        row_mapper: F,
    ) -> AnyResult<Vec<T>, AnyError>
    where
        F: Fn(&Row) -> AnyResult<T, rusqlite::Error>,
    {
        let mut stmt = self
            .archive_db
            .prepare(query)
            .expect("Failed to prepare query");
        let results = stmt
            .query_map(params, row_mapper)?
            .collect::<AnyResult<Vec<T>, _>>()
            .expect("Consolidating Result into vectors");
        Ok(results)
    }

    pub fn execute_query(
        &self,
        query: &str,
        params: &[&(dyn ToSql)],
    ) -> AnyResult<usize, AnyError> {
        return match self.archive_db.execute(&query, params) {
            Ok(result) => Ok(result),
            Err(e) => Err(e.into()),
        };
    }

    pub fn add_update_file_content(
        &self,
        file_name: &str,
        uncompressed_data: Vec<u8>,
    ) -> Result<usize, AnyError> {
        let query: String = get_specific_queries!("open_xml_archive.sql", "insert_archive_table")
            .expect("Specific query pull fail");
        let compressed_data: Vec<u8> =
            compress_content(&uncompressed_data).expect("Recompressing in GZip Failed");
        return match self.archive_db.execute(
            &query,
            params![
                file_name,
                compressed_data.len(),
                uncompressed_data.len(),
                1,
                "gzip",
                compressed_data
            ],
        ) {
            Ok(result) => Ok(result),
            Err(e) => Err(e.into()),
        };
    }

    /// Save the current temp directory state file into final result
    pub fn save(&self, save_file: &str) {
        if metadata(save_file).is_ok() {
            remove_file(save_file).expect("Failed to Remove existing file");
        }
        let query: String = get_specific_queries!("open_xml_archive.sql", "select_archive_table")
            .expect("Read Select Query Pull Failed");
        fn row_mapper(row: &Row) -> AnyResult<ArchiveRecordModel, rusqlite::Error> {
            Ok(ArchiveRecordModel {
                id: row.get(0)?,
                file_name: row.get(1)?,
                compressed_file_size: row.get(2)?,
                uncompressed_file_size: row.get(3)?,
                compression_level: row.get(4)?,
                compression_type: row.get(5)?,
                content: row.get(6)?,
            })
        }
        let result = self
            .find_many(&query, params![], row_mapper)
            .expect("Select Query Results");
        println!("{}", "Test")
    }
    /// Common initialization
    fn common_initialization(is_in_memory: bool) -> Connection {
        let archive_db;
        if is_in_memory {
            // In memory operation
            archive_db = Connection::open_in_memory().expect("Local Lite DB Failed");
        } else {
            match remove_file("./local.db") {
                Ok(_) => (),
                Err(e) if e.kind() == std::io::ErrorKind::NotFound => (),
                Err(e) => eprintln!("Failed to delete the file: {}", e),
            }
            archive_db = Connection::open("./local.db").expect("Local Lite DB Failed");
        }
        Self::initialize_database(&archive_db);
        archive_db
    }

    /// Initialize Local archive Database
    fn initialize_database(archive_db: &Connection) {
        let query: String = get_specific_queries!("open_xml_archive.sql", "create_archive_table")
            .expect("Target Query Pull Failed");
        archive_db
            .execute(&query, ())
            .expect("Base Database Struct Setup Failed");
    }

    /// Read Zip file and load it into database after compression
    fn load_archive_into_database(archive_db: &Connection, file_path: &str) {
        let file: File = File::open(file_path).expect("Read edit file failed.");
        let mut zip_read: ZipArchive<File> = ZipArchive::new(file).expect("Archive read Failed");
        let file_count: usize = zip_read.len();
        for i in 0..file_count {
            let mut file_content = zip_read.by_index(i).expect("Zip file extract failed");
            let mut uncompressed_data = Vec::new();
            file_content
                .read_to_end(&mut uncompressed_data)
                .expect("File Uncompressed failed");
            let compressed =
                compress_content(&uncompressed_data).expect("Recompressing in GZip Failed");
            let query = get_specific_queries!("open_xml_archive.sql", "insert_archive_table")
                .expect("Specific query pull fail");
            archive_db
                .execute(
                    &query,
                    params![
                        file_content.name(),
                        compressed.len(),
                        uncompressed_data.len(),
                        1,
                        "gzip",
                        compressed
                    ],
                )
                .expect("Archive content load failed.");
        }
    }
}
