use std::{cell::RefCell, rc::Rc};

use openxmloffice_xml::OpenXmlFile;

pub trait XmlElement {
    /// Create new object with file connector handle
    fn new(xml_fs: &Rc<RefCell<OpenXmlFile>>,file_name:Option<&str>) -> Self;
    /// Save the current file state
    fn flush(self);
    /// Get content of the current xml
    fn get_content_xml(xml_fs: &Rc<RefCell<OpenXmlFile>>, file_name: &str) -> Vec<u8>;
    /// Initialize the content if not already exist
    fn initialize_content_xml() -> Vec<u8>;
}
