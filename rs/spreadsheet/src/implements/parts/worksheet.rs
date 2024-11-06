use std::{cell::RefCell, rc::Rc};

use openxmloffice_global::xml_file::XmlElement;
use openxmloffice_xml::OpenXmlFile;

use crate::structs::worksheet::Worksheet;

impl XmlElement for Worksheet {
    /// Create New object for the group
    fn new(xml_fs: &Rc<RefCell<OpenXmlFile>>) -> Self {
        //TODO: Dynamically Update file name
        let file_name = "xl/workbook.xml".to_string();
        return Self {
            xml_fs: Rc::clone(xml_fs),
            file_content: Self::get_content_xml(&xml_fs, &file_name),
            file_name,
        };
    }

    fn flush(self) {}

    fn get_content_xml(xml_fs: &Rc<RefCell<OpenXmlFile>>, file_name: &str) -> Vec<u8> {
        let results = xml_fs.borrow().get_xml_content(file_name);
        if let Some(results) = results {
            return results;
        } else {
            return Self::initialize_content_xml(&xml_fs);
        }
    }

    /// Initialize workbook for new excel
    fn initialize_content_xml(xml_fs: &Rc<RefCell<OpenXmlFile>>) -> Vec<u8> {
        let content = Vec::new();
        return content;
    }
}

impl Worksheet {
    /// Set active status for the current worksheet
    pub fn set_active_sheet(&self, is_active: bool) {}

    /// Rename existing sheet name
    pub(crate) fn rename_sheet(&self, sheet_name: &str) {}
}
