use crate::{
    files::{OfficeDocument, XmlDocument},
    global_2007::traits::{XmlDocumentPart, XmlDocumentPartCommon},
};
use anyhow::{anyhow, Context, Error as AnyError, Result as AnyResult};
use std::{cell::RefCell, collections::HashMap, rc::Weak};

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
        file_name: &str,
    ) -> AnyResult<Self, AnyError> {
        let xml_document = Self::get_xml_document(&office_document, &file_name)?;
        Ok(Self {
            office_document,
            xml_document,
            file_name: file_name.to_string(),
        })
    }
}
/// ####################### Im-Mut Access Functions ####################
impl RelationsPart {
    /// Get Relation Target based on Type
    /// Note: This will get the first element match the criteria
    pub fn get_relationship_target_by_type(
        &self,
        content_type: &str,
    ) -> AnyResult<Option<String>, AnyError> {
        if let Some(xml_document) = self.xml_document.upgrade() {
            let xml_doc = xml_document
                .try_borrow()
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

    fn get_next_relationship_id(&self) -> AnyResult<String, AnyError> {
        if let Some(xml_document) = self.xml_document.upgrade() {
            let xml_doc = xml_document
                .try_borrow()
                .context("XML Document Borrow Failed")?;
            if let Some(root) = xml_doc.get_root() {
                let mut children = root.get_child_count() + 1;
                loop {
                    if let Some(_) =
                        xml_doc.get_element_by_attribute("Id", &format!("rId{}", children), None)
                    {
                        children += 1;
                    } else {
                        break;
                    }
                }
                return Ok(format!("rId{}", children));
            }
        }
        Err(anyhow!("Generating Relationship Id Failed"))
    }
}
/// ####################### Mut Access Functions ####################
impl RelationsPart {
    pub fn set_new_relationship_mut(
        &mut self,
        content_type: &str,
        target: &str,
    ) -> AnyResult<(), AnyError> {
        if let Some(xml_document) = self.xml_document.upgrade() {
            let next_id = self
                .get_next_relationship_id()
                .context("Create New Relation Get next Id Failed")?;
            let mut xml_doc_mut = xml_document
                .try_borrow_mut()
                .context("XML Document Borrow Failed")?;
            let new_relationship = xml_doc_mut
                .append_child_mut("Relationship", None)
                .context("Creating Relationship Failed")?;
            let mut attributes: HashMap<String, String> = HashMap::new();
            attributes.insert("Id".to_string(), next_id);
            attributes.insert("Type".to_string(), content_type.to_string());
            attributes.insert("Target".to_string(), target.to_string());
            new_relationship
                .set_attribute_mut(attributes)
                .context("Setting Relationship Attributes Failed.")?;
        }
        Ok(())
    }
}
