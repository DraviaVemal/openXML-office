use std::{cell::RefCell, rc::Rc};

use openxmloffice_xml::OpenXmlFile;
#[derive(Debug)]
pub struct Word {
    pub(crate) xml_fs: Rc<RefCell<OpenXmlFile>>,
}

#[derive(Debug)]
pub struct WordPropertiesModel {
    pub(crate) is_in_memory: bool,
}
