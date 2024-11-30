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
    file_tree: Weak<RefCell<XmlDocument>>,
    file_name: String,
}

impl Drop for RelationsPart {
    fn drop(&mut self) {
        if let Some(xml_document) = self.office_document.upgrade() {
            let _ = xml_document
                .try_borrow_mut()
                .unwrap()
                .close_xml_document(&self.file_name);
        }
    }
}

/// ######################### Train implementation of XML Part - Only accessible within crate ##############
impl XmlDocumentPart for RelationsPart {
    fn new(
        office_document: Weak<RefCell<OfficeDocument>>,
        dir_path: Option<String>,
    ) -> AnyResult<Self, AnyError> {
        let mut file_name = "_rels/.rels".to_string();
        if let Some(dir_path) = dir_path {
            file_name = format!("{}/_rels/.rels", dir_path.to_string());
        }
        let file_tree = Self::get_xml_document(&office_document, &file_name)?;
        Ok(Self {
            office_document,
            file_tree,
            file_name,
        })
    }

    fn flush(self) {}

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
            .set_attribute(attributes);
        Ok(xml_document)
    }
}

impl RelationsPart {
    /// Get Relation Target based on Type
    /// Note: This will get the first element match the criteria
    pub fn get_relationship_target_by_type(
        &self,
        content_type: &str,
    ) -> AnyResult<Option<String>, AnyError> {
        let xml_document_ref: Option<Rc<RefCell<XmlDocument>>> = self.file_tree.borrow().upgrade();
        if let Some(xml_document) = xml_document_ref {
            let xml_doc = xml_document
                .try_borrow_mut()
                .context("XML Document Borrow Failed")?;
            if let Some(xml_root) = xml_doc.get_root() {
                if let Some(result) =
                    xml_doc.get_element_by_attribute(&xml_root.get_id(), "Type", content_type)
                {
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
        }
        Ok(None)
    }
}
