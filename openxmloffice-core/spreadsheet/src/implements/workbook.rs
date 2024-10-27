use openxmloffice_core_xml::OpenXmlFile;

use crate::structs::workbook::Workbook;

impl Workbook {
    /// Create workbook
    pub fn new(xml_fs: &OpenXmlFile) -> Self {
        return Self {};
    }

    /// Read and load workbook xml to work with
    fn get_workbook_xml(xml_fs: &OpenXmlFile) {
        // xml_fs.get_database_connection().execute(sql, params)
    }

    /// Initialize workbook for new excel
    fn initialise_workbook_xml(xml_fs: &OpenXmlFile) {

    }
}
