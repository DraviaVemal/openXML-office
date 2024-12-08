use crate::global_2007::traits::XmlDocumentPartCommon;
use crate::{
    files::{OfficeDocument, XmlDocument},
    global_2007::traits::XmlDocumentPart,
};
use anyhow::{Context, Error as AnyError, Result as AnyResult};
use std::{
    borrow::Borrow,
    cell::RefCell,
    collections::HashMap,
    rc::{Rc, Weak},
};

#[derive(Debug)]
pub struct RelationsPart {
    office_document: Weak<RefCell<OfficeDocument>>,
    xml_document: Weak<RefCell<XmlDocument>>,
    file_name: String,
}

impl Drop for RelationsPart {
    fn drop(&mut self) {
        let _ = self.close_document();
    }
}

impl XmlDocumentPartCommon for RelationsPart {
    /// Initialize xml content for this part from base template
    fn initialize_content_xml() -> AnyResult<XmlDocument, AnyError> {
        let mut attributes: HashMap<String, String> = HashMap::new();
        attributes.insert(
            "xmlns".to_string(),
            "http://schemas.openxmlformats.org/package/2006/relationships".to_string(),
        );
        let mut xml_document = XmlDocument::new();
        xml_document
            .create_root_mut("Relationships")
            .context("Create XML Root Element Failed")?
            .set_attribute_mut(attributes)
            .context("Updating Attribute Failed")?;
        Ok(xml_document)
    }
    fn close_document(&mut self) -> AnyResult<(), AnyError>
    where
        Self: Sized,
    {
        if let Some(xml_document) = self.office_document.upgrade() {
            xml_document
                .try_borrow_mut()
                .unwrap()
                .close_xml_document(&self.file_name)?;
        }
        Ok(())
    }
}

/// ######################### Train implementation of XML Part - Only accessible within crate ##############
impl XmlDocumentPart for RelationsPart {
    fn new(
        office_document: Weak<RefCell<OfficeDocument>>,
        dir_path: Option<String>,
    ) -> AnyResult<Self, AnyError> {
        let mut file_name = "_rels/.rels".to_string();
        if let Some(specific_file_name) = dir_path {
            file_name = specific_file_name;
        }
        let xml_document = Self::get_xml_document(&office_document, &file_name)?;
        Ok(Self {
            office_document,
            xml_document,
            file_name,
        })
    }
}

impl RelationsPart {
    /// Get Relation Target based on Type
    /// Note: This will get the first element match the criteria
    pub fn get_relationship_target_by_type(
        &self,
        content_type: &str,
    ) -> AnyResult<Option<String>, AnyError> {
        let xml_document_ref: Option<Rc<RefCell<XmlDocument>>> =
            self.xml_document.borrow().upgrade();
        if let Some(xml_document) = xml_document_ref {
            let xml_doc = xml_document
                .try_borrow_mut()
                .context("XML Document Borrow Failed")?;
            if let Some(result) = xml_doc.get_element_by_attribute("Type", content_type, None) {
                return Ok(Some(
                    result
                        .get_attribute()
                        .as_ref()
                        .unwrap()
                        .get("Target")
                        .unwrap()
                        .to_string(),
                ));
            }
        }
        Ok(None)
    }
}
