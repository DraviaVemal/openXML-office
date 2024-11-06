use rusqlite::Connection;

/**
 * This contains the root document to work with
 */
#[derive(Debug)]
pub struct OpenXmlFile {
    pub is_editable: bool,
    pub archive_db: Connection,
}

#[derive(Debug)]
pub struct ArchiveRecordModel {
    pub(crate) id: i16,
    pub(crate) file_name: String,
    pub(crate) compressed_file_size: i32,
    pub(crate) uncompressed_file_size: i32,
    pub(crate) compression_level: i8,
    pub(crate) compression_type: String,
    pub(crate) content: Vec<u8>,
}
