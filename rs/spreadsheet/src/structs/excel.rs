use std::{cell::RefCell, rc::Rc};

use super::Workbook;
use openxmloffice_xml::OpenXmlFile;
#[derive(Debug)]
pub struct Excel {
    pub(crate) xml_fs: Rc<RefCell<OpenXmlFile>>,
    pub(crate) workbook: Workbook,
}

pub struct ExcelPropertiesModel {
    pub(crate) is_in_memory: bool,
}
