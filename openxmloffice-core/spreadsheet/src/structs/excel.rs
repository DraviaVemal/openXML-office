use super::Workbook;
use openxmloffice_core_xml::OpenXmlFile;

pub struct Excel {
    pub(crate) xml_fs: OpenXmlFile,
    pub(crate) workbook: Workbook,
}
