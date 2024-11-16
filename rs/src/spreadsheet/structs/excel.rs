use std::{cell::RefCell, rc::Rc};

use super::Workbook;
use openxmloffice_xml::OpenXmlFile;
#[derive(Debug)]
pub struct Excel {
    pub(crate) xml_fs: Rc<RefCell<OpenXmlFile>>,
    pub(crate) workbook: Workbook,
}

#[derive(Debug)]
pub struct ExcelPropertiesModel {
    pub is_in_memory: bool,
}