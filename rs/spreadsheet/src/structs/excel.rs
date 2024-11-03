use super::Workbook;
use openxmloffice_xml::OpenXmlFile;

pub struct Excel {
    pub(crate) xml_fs: OpenXmlFile,
    pub(crate) workbook: Workbook,
}

pub struct ExcelPropertiesModel {
    pub(crate) is_in_memory: bool,
}
