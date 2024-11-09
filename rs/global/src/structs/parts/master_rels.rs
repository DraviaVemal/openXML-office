use std::{cell::RefCell, rc::Rc};

use openxmloffice_xml::OpenXmlFile;

pub struct MasterRelsPart {
    pub xml_fs: Rc<RefCell<OpenXmlFile>>,
    pub file_content: Vec<u8>,
    pub file_name: String,
}
