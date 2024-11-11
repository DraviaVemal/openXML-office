use crate::{CorePropertiesPart, XmlElement};
use anyhow::{Error as AnyError, Result as AnyResult};
use openxmloffice_xml::OpenXmlFile;
use std::{cell::RefCell, rc::Rc};

impl Drop for CorePropertiesPart {
    fn drop(&mut self) {
        self.update_last_modified();
        self.xml_fs
            .borrow()
            .add_update_xml_content(&self.file_name, &self.file_content);
    }
}

impl XmlElement for CorePropertiesPart {
    fn new(xml_fs: &Rc<RefCell<OpenXmlFile>>, _: Option<&str>) -> AnyResult<Self, AnyError> {
        let file_name = "docProps/core.xml".to_string();
        let file_content = Self::get_content_xml(&xml_fs, &file_name)?;
        Ok(Self {
            xml_fs: Rc::clone(xml_fs),
            file_content,
            file_name,
        })
    }

    /// Save the current file state
    fn flush(self) {}

    /// Initialize xml content for this part from base template
    fn initialize_content_xml() -> Vec<u8> {
        let template_core_properties = include_str!("core_properties.xml");
        return template_core_properties.as_bytes().to_vec();
    }
}

impl CorePropertiesPart {
    fn update_last_modified(&mut self) {}
}
