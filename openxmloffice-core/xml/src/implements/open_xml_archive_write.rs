use crate::structs::open_xml_archive_write::OpenXmlWrite;

impl OpenXmlWrite {
    pub fn new() -> Self {
        return Self {};
    }
    /// Add file with directory structure into the archive
    pub fn add_file(file_path: &str) {}
    /// Write the content to archive file
    pub fn write_zip_archive() {}
    /// Use the XML object to write it to file format
    pub fn write_xml(&self, file_path: &str) {}
}
