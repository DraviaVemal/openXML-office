use crate::{
    files::{OfficeDocument, XmlElement, XmlSerializer},
    global_2007::traits::XmlDocument,
};
use anyhow::{Error as AnyError, Result as AnyResult};
use std::{cell::RefCell, rc::Rc};

#[derive(Debug)]
pub struct CorePropertiesPart {
    pub office_document: Rc<RefCell<OfficeDocument>>,
    pub file_content: XmlElement,
    pub file_name: String,
}

impl XmlDocument for CorePropertiesPart {
    fn new(
        office_document: &Rc<RefCell<OfficeDocument>>,
        _: Option<&str>,
    ) -> AnyResult<Self, AnyError> {
        let file_name = "docProps/core.xml".to_string();
        let file_content = Self::get_xml_tree(&office_document, &file_name)?;
        Ok(Self {
            office_document: Rc::clone(office_document),
            file_content,
            file_name,
        })
    }

    /// Save the current file state
    fn flush(self) {}

    /// Initialize xml content for this part from base template
    fn initialize_content_xml() -> AnyResult<XmlElement, AnyError> {
        XmlSerializer::xml_str_to_xml_tree(include_str!("core_properties.xml").as_bytes().to_vec())
    }
}

impl CorePropertiesPart {}
