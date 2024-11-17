use crate::{global_2007::traits::XmlDocument, files::OpenXmlFile};
use anyhow::{Error as AnyError, Result as AnyResult};
use std::{cell::RefCell, rc::Rc};

#[derive(Debug)]
pub struct WorkSheetPart {
    pub xml_fs: Rc<RefCell<OpenXmlFile>>,
    pub file_content: Vec<u8>,
    pub file_name: String,
}

impl Drop for WorkSheetPart {
    fn drop(&mut self) {
        let _ = self
            .xml_fs
            .borrow()
            .add_update_xml_content(&self.file_name, &self.file_content);
    }
}

impl XmlDocument for WorkSheetPart {
    /// Create New object for the group
    fn new(
        xml_fs: &Rc<RefCell<OpenXmlFile>>,
        sheet_name: Option<&str>,
    ) -> AnyResult<Self, AnyError> {
        let mut file_name: String = "xl/worksheets/sheet1.xml".to_string();
        if let Some(sheet_name) = sheet_name {
            file_name = sheet_name.to_string();
        }
        let file_content = Self::get_content_xml(&xml_fs, &file_name)?;
        return Ok(Self {
            xml_fs: Rc::clone(xml_fs),
            file_content,
            file_name,
        });
    }

    fn flush(self) {}

    /// Initialize xml content for this part from base template
    fn initialize_content_xml() -> Vec<u8> {
        let template_core_properties = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></worksheet>"#;
        return template_core_properties.as_bytes().to_vec();
    }
}

impl WorkSheetPart {
    /// Set active status for the current worksheet
    pub fn set_active_sheet(&self, is_active: bool) {}

    /// Rename existing sheet name
    pub(crate) fn rename_sheet(&self, sheet_name: &str) {}
}
