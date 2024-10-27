use crate::{
    get_all_queries, structs::open_xml_archive::OpenXmlFile, utils::file_handling::compress_content,
};
use rusqlite::Connection;
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

    pub fn get_database_connection(&self) -> &Connection {
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
            archive_db
                .execute("INSERT INTO archive (filename, compression_level, compression_type, content) VALUES (?1, ?2, ?3, ?4)", (file_content.name(),1,1,gzip))
                .expect("Archive content load failed.");
        }
    }

    /// Save the current temp directory state file into final result
    pub fn save(&self, save_file: &str) {
        if metadata(save_file).is_ok() {
            remove_file(save_file).expect("Failed to Remove existing file");
        }
    }
}
