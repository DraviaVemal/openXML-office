use openxmloffice_xml::OpenXmlFile;
use std::{cell::RefCell, rc::Rc};

#[derive(Debug)]
pub struct Worksheet {
    pub xml_fs: Rc<RefCell<OpenXmlFile>>,
    pub file_content: Vec<u8>,
    pub file_name: String,
}
