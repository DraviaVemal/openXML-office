use crate::{
    files::{OfficeDocument, XmlElement},
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
    file_tree: Weak<RefCell<XmlElement>>,
    file_name: String,
}

impl Drop for RelationsPart {
    fn drop(&mut self) {
        if let Some(xml_tree) = self.office_document.upgrade() {
            let _ = xml_tree
                .try_borrow_mut()
                .unwrap()
                .close_xml_tree(&self.file_name);
        }
    }
}

impl XmlDocumentPart for RelationsPart {
    fn new(
        office_document: &Rc<RefCell<OfficeDocument>>,
        dir_path: Option<String>,
    ) -> AnyResult<Self, AnyError> {
        let mut file_name = "_rels/.rels".to_string();
        if let Some(dir_path) = dir_path {
            file_name = format!("{}/_rels/.rels", dir_path.to_string());
        }
        let file_tree = Self::get_xml_tree(&office_document, &file_name)?;
        Ok(Self {
            office_document: Rc::downgrade(office_document),
            file_tree,
            file_name,
        })
    }

    fn flush(self) {}

    /// Initialize xml content for this part from base template
    fn initialize_content_xml() -> AnyResult<XmlElement, AnyError> {
        let mut xml_element: XmlElement = XmlElement::new("Relationships".to_string(), None);
        let mut attributes: HashMap<String, String> = HashMap::new();
        attributes.insert(
            "xmlns".to_string(),
            "http://schemas.openxmlformats.org/package/2006/relationships".to_string(),
        );
        xml_element.set_attribute(attributes);
        Ok(xml_element)
    }
}

impl RelationsPart {
    /// Get Relation Target based on Type
    /// Note: This will get the first element match the criteria
    pub fn get_relationship_target_by_type(
        &self,
        content_type: &str,
    ) -> AnyResult<Option<String>, AnyError> {
        let xml_tree_ref = self.file_tree.borrow().upgrade();
        if let Some(xml_tree) = xml_tree_ref {
            let xml_element = xml_tree
                .try_borrow()
                .context("XML Tree Borrow Failed")?
                .find_child_element_by_attribute("Type", content_type)
                .context("Find element by attribute Failed")?;
            if let Some(element) = xml_element {
                if let Some(attributes) = element.get_attribute() {
                    if let Some((_, value)) = attributes.get_key_value("Target") {
                        return Ok(Some(value.to_string()));
                    }
                }
            }
        }
        Ok(None)
    }
}
