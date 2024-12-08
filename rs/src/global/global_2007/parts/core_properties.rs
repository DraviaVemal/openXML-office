use crate::global_2007::traits::XmlDocumentPartCommon;
use crate::{
    files::{OfficeDocument, XmlDocument, XmlSerializer},
    global_2007::traits::XmlDocumentPart,
};
use anyhow::{Error as AnyError, Result as AnyResult};
use chrono::Utc;
use std::{cell::RefCell, rc::Weak};

#[derive(Debug)]
pub struct CorePropertiesPart {
    office_document: Weak<RefCell<OfficeDocument>>,
    xml_document: Weak<RefCell<XmlDocument>>,
    file_name: String,
}

impl Drop for CorePropertiesPart {
    fn drop(&mut self) {
        let _ = self.close_document();
    }
}

impl XmlDocumentPartCommon for CorePropertiesPart {
    /// Initialize xml content for this part from base template
    fn initialize_content_xml() -> AnyResult<XmlDocument, AnyError> {
        XmlSerializer::vec_to_xml_doc_tree(include_str!("core_properties.xml").as_bytes().to_vec())
    }

    fn close_document(&mut self) -> AnyResult<(), AnyError>
    where
        Self: Sized,
    {
        // Update Last modified date part
        if let Some(xml_document_ref) = self.xml_document.upgrade() {
            let mut xml_document = xml_document_ref.borrow_mut();
            match xml_document
                .get_first_element_mut(vec!["cp:coreProperties", "dcterms:modified"], None)
            {
                Ok(result) => {
                    if let Some(element) = result {
                        element.set_value_mut(
                            Utc::now().to_rfc3339_opts(chrono::SecondsFormat::Secs, true),
                        );
                    }
                }
                Err(_) => (),
            }
        }
        // Update the current state to DB before dropping the object
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
impl XmlDocumentPart for CorePropertiesPart {
    fn new(
        office_document: Weak<RefCell<OfficeDocument>>,
        _: Option<String>,
    ) -> AnyResult<Self, AnyError> {
        let file_name = "docProps/core.xml".to_string();
        let xml_document = Self::get_xml_document(&office_document, &file_name)?;
        Ok(Self {
            office_document,
            xml_document,
            file_name,
        })
    }
}

impl CorePropertiesPart {}
