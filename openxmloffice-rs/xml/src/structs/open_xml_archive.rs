use rusqlite::Connection;

/**
 * This contains the root document to work with
 */
pub struct OpenXmlFile {
    pub is_editable: bool,
    pub archive_db: Connection,
}
