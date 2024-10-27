use rusqlite::Connection;

/**
 * This contains the root document to work with
 */
pub struct OpenXmlFile {
    pub(crate) is_editable: bool,
    pub(crate) archive_db: Connection,
}
