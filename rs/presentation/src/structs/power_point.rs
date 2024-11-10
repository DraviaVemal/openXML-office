use std::{cell::RefCell, rc::Rc};

use openxmloffice_xml::OpenXmlFile;
#[derive(Debug)]
pub struct PowerPoint {
    pub(crate) xml_fs: Rc<RefCell<OpenXmlFile>>,
}

#[derive(Debug)]
pub struct PowerPointPropertiesModel {
    pub(crate) is_in_memory: bool,
}
