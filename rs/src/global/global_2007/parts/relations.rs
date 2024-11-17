use crate::{
    files::{OfficeDocument, XmlElement, XmlSerializer},
    global_2007::traits::XmlDocument,
};
use anyhow::{Error as AnyError, Result as AnyResult};
use std::{cell::RefCell, rc::Rc};

#[derive(Debug)]
pub struct RelationsPart {
    pub office_document: Rc<RefCell<OfficeDocument>>,
    pub file_content: XmlElement,
    pub file_name: String,
}

impl XmlDocument for RelationsPart {
    fn new(
        office_document: &Rc<RefCell<OfficeDocument>>,
        relation_file_name: Option<&str>,
    ) -> AnyResult<Self, AnyError> {
        let mut file_name = ".rels".to_string();
        if let Some(relation_file_name) = relation_file_name {
            file_name = relation_file_name.to_string();
        }
        let file_content = Self::get_xml_tree(&office_document, &file_name)?;
        Ok(Self {
            office_document: Rc::clone(office_document),
            file_content,
            file_name,
        })
    }

    fn flush(self) {}

    /// Initialize xml content for this part from base template
    fn initialize_content_xml() -> AnyResult<XmlElement, AnyError> {
        let template_core_properties = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>"#;
        XmlSerializer::xml_str_to_xml_tree(template_core_properties.as_bytes().to_vec())
    }
}

impl RelationsPart {}
