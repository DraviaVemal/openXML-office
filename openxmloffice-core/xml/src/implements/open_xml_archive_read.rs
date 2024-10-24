use crate::structs::open_xml_archive_read::OpenXmlRead;

impl OpenXmlRead {
    pub fn new() -> Self {
        return Self {};
    }
    /// Read target file from archive
    pub fn read_zip_archive() {}
    /// Read file content and parse it to XML object
    pub fn read_xml(&self, file_path: &str) {}
}
