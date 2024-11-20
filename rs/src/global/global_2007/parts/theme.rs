use crate::{
    files::{OfficeDocument, XmlElement, XmlSerializer},
    global_2007::traits::XmlDocumentPart,
};
use anyhow::{Error as AnyError, Result as AnyResult};
use std::{cell::RefCell, rc::Rc};

#[derive(Debug)]
pub struct ThemePart {
    pub office_document: Rc<RefCell<OfficeDocument>>,
    pub file_content: XmlElement,
    pub file_name: String,
}

impl XmlDocumentPart for ThemePart {
    fn new(
        office_document: &Rc<RefCell<OfficeDocument>>,
        file_name: Option<&str>,
    ) -> AnyResult<Self, AnyError> {
        let mut local_file_name = "".to_string();
        if let Some(file_name) = file_name {
            local_file_name = file_name.to_string();
        }
        let file_content = Self::get_xml_tree(&office_document, &local_file_name)?;
        Ok(Self {
            office_document: Rc::clone(office_document),
            file_content,
            file_name: local_file_name.to_string(),
        })
    }

    fn flush(self) {}

    /// Initialize xml content for this part from base template
    fn initialize_content_xml() -> AnyResult<XmlElement, AnyError> {
        XmlSerializer::xml_str_to_xml_tree(include_str!("theme.xml").as_bytes().to_vec())
    }
}

impl ThemePart {}
