use openxmloffice_core_xml::OpenXmlFile;

use super::workbook::Workbook;

pub struct Excel {
    pub(crate) xml_fs: OpenXmlFile,
    pub(crate) workbook: Workbook,
}
