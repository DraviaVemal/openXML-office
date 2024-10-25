use std::{cell::RefCell, rc::Rc};

use anyhow::{Ok, Result};
use zip::write::SimpleFileOptions;

use crate::structs::open_xml_archive_write::OpenXmlEditable;

impl<'buffer> OpenXmlEditable<'buffer> {
    /// Create a Editor object for the current file
    pub fn new(working_buffer: &'buffer Rc<RefCell<Vec<u8>>>) -> Self {
        return Self { working_buffer };
    }

    /// Add file with directory structure into the archive
    pub fn add_file(&self, archive_path: &str, content: &str) -> Result<()> {
        let zip_option = SimpleFileOptions::default();
        zip_write
            .start_file(archive_path, zip_option)
            .expect("Add file failed");
        Ok(())
    }
    /// Use the XML object to write it to file format
    pub fn write_xml(&self, file_path: &str) {}
    /// Write the content to archive file
    pub fn write_zip_archive() {}
}
