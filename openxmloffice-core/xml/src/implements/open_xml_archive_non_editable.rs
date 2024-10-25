use crate::{structs::open_xml_archive_read::OpenXmlNonEditable, OpenXmlFile};

impl<'file_handle> OpenXmlNonEditable<'file_handle> {
    pub fn new(file_handle: &'file_handle OpenXmlFile) -> Self {
        return Self {
            open_xml_file: file_handle,
        };
    }
    /// Read target file from archive
    pub fn read_zip_archive() {}
    /// Read file content and parse it to XML object
    pub fn read_xml(&self, file_path: &str) {}
}
