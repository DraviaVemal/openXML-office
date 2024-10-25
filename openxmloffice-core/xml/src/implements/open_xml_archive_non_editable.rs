use crate::structs::open_xml_archive_read::OpenXmlNonEditable;
use std::{cell::RefCell, rc::Rc};

impl<'buffer> OpenXmlNonEditable<'buffer> {
    pub fn new(working_buffer: &'buffer Rc<RefCell<Vec<u8>>>) -> Self {
        return Self { working_buffer };
    }
    /// Read target file from archive
    pub fn read_zip_archive() {}
    /// Read file content and parse it to XML object
    pub fn read_xml(&self, file_path: &str) {}
}
