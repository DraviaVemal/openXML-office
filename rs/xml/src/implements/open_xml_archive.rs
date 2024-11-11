// TODO : After the file gets shaped with multiple function. Reorganize it to xml, archive and database into separate implementation

use crate::{
    file_handling::{compress_content, decompress_content},
    get_specific_queries, ArchiveRecordModel, OpenXmlFile,
};
use anyhow::{anyhow, Context, Error as AnyError, Ok, Result as AnyResult};
use rusqlite::{params, Connection, Row, ToSql};
use std::{
    fs::{metadata, remove_file, File},
    io::{Cursor, Read, Write},
};
use tempfile::NamedTempFile;
use zip::{write::SimpleFileOptions, ZipArchive, ZipWriter};

impl OpenXmlFile {
    /// Create Current file helper object from exiting source
    pub fn open(
        file_path: &str,
        is_editable: bool,
        is_in_memory: bool,
    ) -> AnyResult<Self, AnyError> {
        let archive_db =
            Self::common_initialization(is_in_memory).context("Create Connection Fail")?;
        Self::load_archive_into_database(&archive_db, file_path);
        // Create a clone copy of master file to work with code
        Ok(Self {
            is_editable,
            archive_db,
        })
    }

    /// Create Current file helper object a new file to work with
    pub fn create(is_in_memory: bool) -> AnyResult<Self, AnyError> {
        let archive_db: Connection =
            Self::common_initialization(is_in_memory).context("Create Connection Fail")?;
        Ok(Self {
            is_editable: true,
            archive_db,
        })
    }

    /// Get File XML Content
    pub fn get_xml_content(&self, file_name: &str) -> AnyResult<Option<Vec<u8>>, AnyError> {
        let query = get_specific_queries!("open_xml_archive.sql", "select_archive_content")
            .map_err(|e| anyhow!("Query Macro Error : {}", e))?;
        fn row_mapper(row: &Row) -> Result<Vec<u8>, rusqlite::Error> {
            row.get(0)
        }
        let result = self
            .find_one(&query, params![file_name], row_mapper)
            .map_err(|e| anyhow!("Failed to execute the Find one Query : {}", e))?;
        match result {
            Some(data) => decompress_content(&data)
                .map(Some)
                .map_err(|e| anyhow!("Decompression failed: {}", e)),
            None => Ok(None),
        }
    }

    /// Set or Update XML Content of target file
    pub fn add_update_xml_content(
        &self,
        file_name: &str,
        uncompressed_data: &Vec<u8>,
    ) -> Result<(), AnyError> {
        let query: String = get_specific_queries!("open_xml_archive.sql", "insert_archive_table")
            .map_err(|e| anyhow!("Query Macro Error : {}", e))?;
        let compressed_data: Vec<u8> =
            compress_content(&uncompressed_data).context("Recompressing in GZip Failed")?;
        self.archive_db
            .execute(
                &query,
                params![
                    file_name,
                    "todo",
                    compressed_data.len(),
                    uncompressed_data.len(),
                    1,
                    "gzip",
                    compressed_data
                ],
            )
            .context("Failed To Add/Update Content")?;
        Ok(())
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
            .map_err(|e| anyhow!("Failed to Run Find One Query {}", e))?;
        match stmt.query_row(params, row_mapper) {
            Result::Ok(result) => Ok(Some(result)),
            Err(rusqlite::Error::QueryReturnedNoRows) => Ok(None),
            Err(e) => Err(e.into()),
        }
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
            .map_err(|e| anyhow!("Failed to Run Find Many Query {}", e))?;
        stmt.query_map(params, row_mapper)?
            .collect::<AnyResult<Vec<T>, _>>()
    }

    pub fn execute_query(
        &self,
        query: &str,
        params: &[&(dyn ToSql)],
    ) -> AnyResult<usize, AnyError> {
        match self.archive_db.execute(&query, params) {
            Result::Ok(result) => Ok(result),
            Err(e) => Err(e.into()),
        }
    }

    /// Save the current temp directory state file into final result
    pub fn save(&self, save_file: &str) -> AnyResult<(), AnyError> {
        let file_content = self
            .save_database_into_archive()
            .context("Save Archive Data into Database")?;
        if metadata(save_file).is_ok() {
            remove_file(save_file).map_err(|e| anyhow!("Remove Save File Target Failed. {}", e))?;
        }
        let mut file = File::create(save_file).context("Create Save File Failed")?;
        file.write_all(&file_content)
            .context("Save File Write Failed")
    }

    /// Common initialization
    fn common_initialization(is_in_memory: bool) -> AnyResult<Connection, AnyError> {
        let archive_db;
        if is_in_memory {
            // In memory operation
            archive_db = Connection::open_in_memory()
                .map_err(|e| anyhow!("Open in memory DB failed. {}", e))?;
        } else {
            let temp_file = NamedTempFile::new().context("Temp. File Creation Failed")?;
            match remove_file(&temp_file) {
                Result::Ok(_) => (),
                Err(e) if e.kind() == std::io::ErrorKind::NotFound => (),
                Err(e) => eprintln!("Failed to delete the file: {}", e),
            }
            archive_db =
                Connection::open(&temp_file).map_err(|e| anyhow!("Open file DB failed. {}", e))?;
        }
        Self::initialize_database(&archive_db);
        Ok(archive_db)
    }

    /// Initialize Local archive Database
    fn initialize_database(archive_db: &Connection) -> AnyResult<usize, AnyError> {
        let query: String = get_specific_queries!("open_xml_archive.sql", "create_archive_table")
            .map_err(|e| anyhow!("Query Macro Error : {}", e))?;
        archive_db
            .execute(&query, ())
            .context("Query Execution Failed")
    }

    /// Save the database content into file archive
    fn save_database_into_archive(&self) -> AnyResult<Vec<u8>, AnyError> {
        let query: String =
            get_specific_queries!("open_xml_archive.sql", "select_all_archive_rows")
                .map_err(|e| anyhow!("Query Macro Error : {}", e))?;
        fn row_mapper(row: &Row) -> AnyResult<ArchiveRecordModel, rusqlite::Error> {
            Ok(ArchiveRecordModel {
                id: row.get(0),
                file_name: row.get(1),
                content_type: row.get(2),
                compressed_file_size: row.get(3),
                uncompressed_file_size: row.get(4),
                compression_level: row.get(5),
                compression_type: row.get(6),
                content: row.get(7),
            })
        }
        let result = self.find_many(&query, params![], row_mapper)?;
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
            content,
        } in result
        {
            let uncompressed = decompress_content(&content).context("Decompress Error")?;
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
        archive_db: &Connection,
        file_path: &str,
    ) -> AnyResult<(), AnyError> {
        let file: File = File::open(file_path).context("Open Existing File")?;
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
            archive_db
                .execute(
                    &query,
                    params![
                        file_content.name(),
                        "todo",
                        compressed.len(),
                        uncompressed_data.len(),
                        1,
                        "gzip",
                        compressed
                    ],
                )
                .context("Archive content load failed.")?;
        }
        Ok(())
    }
}
