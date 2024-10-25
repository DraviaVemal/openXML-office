use std::fs::File;

use zip::ZipWriter;

use crate::{structs::open_xml_archive_write::OpenXmlEditable, OpenXmlFile};

impl<'file_handle> OpenXmlEditable<'file_handle> {
    pub fn new(file_handle: &'file_handle OpenXmlFile) -> Self {
        return Self {
            open_xml_file: file_handle,
        };
    }
    /// Add file with directory structure into the archive
    pub fn add_file(&self, file_path: &str) {
        let read_file = File::open(file_path).expect("Archive File Read Failed");
        let mut zip_archive = ZipWriter::new(read_file);
        
    }
    /// Use the XML object to write it to file format
    pub fn write_xml(&self, file_path: &str) {}
    /// Write the content to archive file
    pub fn write_zip_archive() {}
}
