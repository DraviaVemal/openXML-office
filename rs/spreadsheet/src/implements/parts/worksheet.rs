use crate::structs::worksheet::Worksheet;
use openxmloffice_global::xml_file::XmlElement;
use openxmloffice_xml::OpenXmlFile;
use std::{cell::RefCell, rc::Rc};

impl XmlElement for Worksheet {
    /// Create New object for the group
    fn new(xml_fs: &Rc<RefCell<OpenXmlFile>>, sheet_name: Option<&str>) -> Self {
        //TODO: Dynamically Update file name
        let mut file_name = "xl/worksheets/sheet1.xml".to_string();
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
            return Self::initialize_content_xml();
        }
    }

    /// Initialize xml content for this part from base template
    fn initialize_content_xml() -> Vec<u8> {
        let template_core_properties = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></worksheet>"#;
        return template_core_properties.as_bytes().to_vec();
    }
}

impl Worksheet {
    /// Set active status for the current worksheet
    pub fn set_active_sheet(&self, is_active: bool) {}

    /// Rename existing sheet name
    pub(crate) fn rename_sheet(&self, sheet_name: &str) {}
}
