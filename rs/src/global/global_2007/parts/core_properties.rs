use crate::{global_2007::traits::XmlDocument, files::OfficeDocument};
use anyhow::{Error as AnyError, Result as AnyResult};
use std::{cell::RefCell, rc::Rc};

#[derive(Debug)]
pub struct CorePropertiesPart {
    pub office_document: Rc<RefCell<OfficeDocument>>,
    pub file_content: Vec<u8>,
    pub file_name: String,
}

impl Drop for CorePropertiesPart {
    fn drop(&mut self) {
        self.update_last_modified();
        let _ = self
            .office_document
            .borrow()
            .add_update_xml_content(&self.file_name, &self.file_content);
    }
}

impl XmlDocument for CorePropertiesPart {
    fn new(office_document: &Rc<RefCell<OfficeDocument>>, _: Option<&str>) -> AnyResult<Self, AnyError> {
        let file_name = "docProps/core.xml".to_string();
        let file_content = Self::get_content_xml(&office_document, &file_name)?;
        Ok(Self {
            office_document: Rc::clone(office_document),
            file_content,
            file_name,
        })
    }

    /// Save the current file state
    fn flush(self) {}

    /// Initialize xml content for this part from base template
    fn initialize_content_xml() -> Vec<u8> {
        let template_core_properties = include_str!("core_properties.xml");
        template_core_properties.as_bytes().to_vec()
    }
}

impl CorePropertiesPart {
    fn update_last_modified(&mut self) {}
}
