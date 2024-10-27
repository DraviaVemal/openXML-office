use anyhow::{Ok, Result};

use crate::structs::open_xml_archive_write::OpenXmlEditable;

impl OpenXmlEditable {
    /// Create a Editor object for the current file
    pub fn new() -> Self {
        return Self {};
    }

    /// Add file with directory structure into the archive
    pub fn add_file(&self, archive_path: &str, content: &str) -> Result<()> {
        Ok(())
    }
    /// Use the XML object to write it to file format
    pub fn write_xml(&self, file_path: &str) {}
    /// Write the content to archive file
    pub fn write_zip_archive() {}
}
