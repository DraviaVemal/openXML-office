use crate::{global_2007::traits::XmlElement, files::OpenXmlFile};
use anyhow::{Error as AnyError, Result as AnyResult};
use std::{cell::RefCell, rc::Rc};

#[derive(Debug)]
pub struct RelationsPart {
    pub xml_fs: Rc<RefCell<OpenXmlFile>>,
    pub file_content: Vec<u8>,
    pub file_name: String,
}

impl Drop for RelationsPart {
    fn drop(&mut self) {
        let _ = self
            .xml_fs
            .borrow()
            .add_update_xml_content(&self.file_name, &self.file_content);
    }
}

impl XmlElement for RelationsPart {
    fn new(
        xml_fs: &Rc<RefCell<OpenXmlFile>>,
        relation_file_name: Option<&str>,
    ) -> AnyResult<Self, AnyError> {
        let mut file_name = ".rels".to_string();
        if let Some(relation_file_name) = relation_file_name {
            file_name = relation_file_name.to_string();
        }
        let file_content = Self::get_content_xml(&xml_fs, &file_name)?;
        Ok(Self {
            xml_fs: Rc::clone(xml_fs),
            file_content,
            file_name,
        })
    }

    fn flush(self) {}

    /// Initialize xml content for this part from base template
    fn initialize_content_xml() -> Vec<u8> {
        let template_core_properties = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>"#;
        template_core_properties.as_bytes().to_vec()
    }
}

impl RelationsPart {}
