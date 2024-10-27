use crate::structs::open_xml_archive_read::OpenXmlNonEditable;

impl OpenXmlNonEditable {
    ///
    pub fn new() -> Self {
        return Self {};
    }
    /// List all files in archive
    pub fn get_file_list(&self) -> Vec<String> {
        return vec![];
    }
    /// Read target file from archive
    pub fn read_zip_archive() {}
    /// Read file content and parse it to XML object
    pub fn read_xml(&self, file_path: &str) {}
}
