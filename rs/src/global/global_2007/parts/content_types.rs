use crate::global_2007::traits::XmlDocumentPartCommon;
use crate::{
    files::{OfficeDocument, XmlDocument, XmlSerializer},
    global_2007::traits::XmlDocumentPart,
};
use anyhow::{Error as AnyError, Result as AnyResult};
use std::{cell::RefCell, rc::Weak};

#[derive(Debug)]
pub struct ContentTypesPart {
    office_document: Weak<RefCell<OfficeDocument>>,
    xml_document: Weak<RefCell<XmlDocument>>,
    file_name: String,
}

impl Drop for ContentTypesPart {
    fn drop(&mut self) {
        let _ = self.close_document();
    }
}

impl XmlDocumentPartCommon for ContentTypesPart {
    /// Initialize xml content for this part from base template
    fn initialize_content_xml() -> AnyResult<XmlDocument, AnyError> {
        XmlSerializer::vec_to_xml_doc_tree(include_str!("content_types.xml").as_bytes().to_vec())
    }

    fn close_document(&mut self) -> AnyResult<(), AnyError>
    where
        Self: Sized,
    {
        if let Some(xml_tree) = self.office_document.upgrade() {
            xml_tree
                .try_borrow_mut()
                .unwrap()
                .close_xml_document(&self.file_name)?;
        }
        Ok(())
    }
}

/// ######################### Train implementation of XML Part - Only accessible within crate ##############
impl XmlDocumentPart for ContentTypesPart {
    fn new(
        office_document: Weak<RefCell<OfficeDocument>>,
        _: Option<String>,
    ) -> AnyResult<Self, AnyError> {
        let file_name = "[Content_Types].xml".to_string();
        let xml_document = Self::get_xml_document(&office_document, &file_name)?;
        Ok(Self {
            office_document,
            xml_document,
            file_name,
        })
    }
}

impl ContentTypesPart {}
