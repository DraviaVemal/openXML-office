use crate::{file_handling::compress_content, get_all_queries, get_specific_queries, OpenXmlFile};
use anyhow::Result;
use rusqlite::{params, Connection, Error, ToSql};
use std::{
    fs::{metadata, remove_file, File},
    io::Read,
};
use zip::ZipArchive;

impl OpenXmlFile {
    /// Create Current file helper object from exiting source
    pub fn open(file_path: &str, is_editable: bool) -> Self {
        let archive_db = Connection::open_in_memory().expect("Local Lite DB Failed");
        Self::initialize_database(&archive_db);
        Self::load_archive_into_database(&archive_db, file_path);
        // Create a clone copy of master file to work with code
        Self {
            is_editable,
            archive_db,
        }
    }

    /// Create Current file helper object a new file to work with
    pub fn create() -> Self {
        let archive_db = Connection::open_in_memory().expect("Local Lite DB Failed");
        // Default List of files common for all types
        Self::initialize_database(&archive_db);
        Self {
            is_editable: true,
            archive_db,
        }
    }

    pub fn get_query_result(
        &self,
        query: &str,
        params: &[&(dyn ToSql)],
    ) -> Result<Option<Vec<u8>>, Error> {
        match self
            .get_database_connection()
            .query_row(&query, params, |row| row.get(0))
        {
            Ok(content) => Ok(Some(content)),
            Err(Error::QueryReturnedNoRows) => Ok(None), // Handle the "no rows" case.
            Err(e) => Err(e),
        }
    }

    pub fn execute_query(&self, query: &str, params: &[&(dyn ToSql)]) -> Result<usize, Error> {
        return self.get_database_connection().execute(&query, params);
    }

    pub fn add_update_file_content(&self, data: Vec<u8>, file_name: &str) -> Result<usize, Error> {
        let query = get_specific_queries!("open_xml_archive.sql", "insert_archive_table")
            .expect("Specific query pull fail");
        let gzip = compress_content(&data).expect("Recompressing in GZip Failed");
        return self.archive_db.execute(
            &query,
            params![file_name, gzip.len(), data.len(), 1, "gzip", gzip],
        );
    }

    /// Save the current temp directory state file into final result
    pub fn save(&self, save_file: &str) {
        if metadata(save_file).is_ok() {
            remove_file(save_file).expect("Failed to Remove existing file");
        }
    }

    /// Get Database connection
    fn get_database_connection(&self) -> &Connection {
        return &self.archive_db;
    }

    /// Initialize Local archive Database
    fn initialize_database(archive_db: &Connection) {
        let queries = get_all_queries!("open_xml_archive.sql");
        for query in queries {
            archive_db
                .execute(&query, ())
                .expect("Base Database Struct Setup Failed");
        }
    }

    fn load_archive_into_database(archive_db: &Connection, file_path: &str) {
        let file = File::open(file_path).expect("Read edit file failed.");
        let mut zip_read = ZipArchive::new(file).expect("Archive read Failed");
        let file_count = zip_read.len();
        for i in 0..file_count {
            let mut file_content = zip_read.by_index(i).expect("Zip file extract failed");
            let mut uncompressed_data = Vec::new();
            file_content
                .read_to_end(&mut uncompressed_data)
                .expect("File Uncompressed failed");
            let gzip = compress_content(&uncompressed_data).expect("Recompressing in GZip Failed");
            let query = get_specific_queries!("open_xml_archive.sql", "insert_archive_table")
                .expect("Specific query pull fail");
            archive_db
                .execute(
                    &query,
                    params![
                        file_content.name(),
                        gzip.len(),
                        uncompressed_data.len(),
                        1,
                        "gzip",
                        gzip
                    ],
                )
                .expect("Archive content load failed.");
        }
    }
}
