use crate::structs::worksheet::Worksheet;
use openxmloffice_global::xml_file::XmlElement;
use openxmloffice_xml::OpenXmlFile;
use std::{cell::RefCell, rc::Rc};

impl Drop for Worksheet {
    fn drop(&mut self) {
        self.xml_fs
            .borrow()
            .add_update_xml_content(&self.file_name, &self.file_content)
    }
}

impl XmlElement for Worksheet {
    /// Create New object for the group
    fn new(xml_fs: &Rc<RefCell<OpenXmlFile>>, sheet_name: Option<&str>) -> Self {
        let mut file_name: String = "xl/worksheets/sheet1.xml".to_string();
        if let Some(sheet_name) = sheet_name {
            file_name = sheet_name.to_string();
        }
        return Self {
            xml_fs: Rc::clone(xml_fs),
            file_content: Self::get_content_xml(&xml_fs, &file_name),
            file_name,
        };
    }

    fn flush(self) {}

    /// Initialize xml content for this part from base template
    fn initialize_content_xml() -> Vec<u8> {
        let template_core_properties = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></worksheet>"#;
        return template_core_properties.as_bytes().to_vec();
    }
}

impl Worksheet {
    /// Set active status for the current worksheet
    pub fn set_active_sheet(&self, is_active: bool) {}

    /// Rename existing sheet name
    pub(crate) fn rename_sheet(&self, sheet_name: &str) {}
}
