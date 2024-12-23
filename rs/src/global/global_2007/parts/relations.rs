use crate::{
    files::{OfficeDocument, XmlDocument},
    global_2007::traits::XmlDocumentPartCommon,
    reference_dictionary::{Content, COMMON_TYPE_COLLECTION},
};
use anyhow::{anyhow, Context, Error as AnyError, Result as AnyResult};
use std::{cell::RefCell, collections::HashMap, rc::Weak};

#[derive(Debug)]
pub(crate) struct RelationsPart {
    /// Holds Id,Target,Type
    relationships: Vec<(String, String, String)>,
    office_document: Weak<RefCell<OfficeDocument>>,
    xml_document: Weak<RefCell<XmlDocument>>,
    file_path: String,
}

impl Drop for RelationsPart {
    fn drop(&mut self) {
        let _ = self.close_document();
    }
}

impl XmlDocumentPartCommon for RelationsPart {
    /// Initialize xml content for this part from base template
    fn initialize_content_xml(
    ) -> AnyResult<(XmlDocument, Option<String>, Option<String>, Option<String>), AnyError> {
        let relationship_content = COMMON_TYPE_COLLECTION.get("rels").unwrap();
        let mut attributes = HashMap::new();
        attributes.insert(
            "xmlns".to_string(),
            relationship_content.schemas_namespace.to_string(),
        );
        let mut xml_document = XmlDocument::new();
        xml_document
            .create_root_mut("Relationships")
            .context("Create XML Root Element Failed")?
            .set_attribute_mut(attributes)
            .context("Updating Attribute Failed")?;
        Ok((
            xml_document,
            None,
            Some(relationship_content.extension.to_string()),
            Some(relationship_content.extension_type.to_string()),
        ))
    }
    /// Close the Current Relation Document
    fn close_document(&mut self) -> AnyResult<(), AnyError>
    where
        Self: Sized,
    {
        if let Some(xml_document) = self.office_document.upgrade() {
            if self
                .save_relationship_to_doc()
                .context("Failed to Insert relationship to Relationships")?
            {
                // Remove Links that are not valid
                xml_document
                    .try_borrow_mut()
                    .context("Failed to Pull Open XML Handle")?
                    .close_xml_document(&self.file_path)?;
            } else {
                xml_document
                    .try_borrow_mut()
                    .context("Failed to Pull Open XML Handle")?
                    .delete_document_mut(&self.file_path)
                    .context("Failed to delete document from database")?;
            }
        }
        Ok(())
    }
}

/// ######################### Train implementation of XML Part - Only accessible within crate ##############
impl RelationsPart {
    pub(crate) fn new(
        office_document: Weak<RefCell<OfficeDocument>>,
        file_name: &str,
    ) -> AnyResult<Self, AnyError> {
        let mut xml_document = Self::get_xml_document(&office_document, &file_name)?;
        let relationships =
            Self::load_relations(&mut xml_document).context("Failed to decode Relations")?;
        Ok(Self {
            office_document,
            xml_document,
            relationships,
            file_path: file_name.to_string(),
        })
    }
}

/// ####################### Im-Mut Access Functions ####################
impl RelationsPart {
    pub(crate) fn load_relations(
        xml_document: &mut Weak<RefCell<XmlDocument>>,
    ) -> AnyResult<Vec<(String, String, String)>, AnyError> {
        let mut relationships = Vec::new();
        if let Some(xml_document) = xml_document.upgrade() {
            let mut xml_doc = xml_document
                .try_borrow_mut()
                .context("Failed for get XML Handle")?;
            if let Some(ids) = xml_doc.get_element_ids_by_tag("Relationship", None) {
                for element_id in ids {
                    let relationship = xml_doc
                        .pop_element_mut(&element_id)
                        .ok_or(anyhow!("Failed to Get Target Element"))?;
                    let attributes = relationship
                        .get_attribute()
                        .ok_or(anyhow!("Failed! Relationship attributes missing"))?;
                    relationships.push((
                        attributes
                            .get("Id")
                            .ok_or(anyhow!("Failed. Id in relationship Not Fount!"))?
                            .to_string(),
                        attributes
                            .get("Target")
                            .ok_or(anyhow!("Failed. Target in relationship Not Fount!"))?
                            .to_string(),
                        attributes
                            .get("Type")
                            .ok_or(anyhow!("Failed. Type in relationship Not Fount!"))?
                            .to_string(),
                    ));
                }
            }
        }
        Ok(relationships)
    }

    pub(crate) fn get_relative_path(&self) -> AnyResult<String, AnyError> {
        let rels_position = self
            .file_path
            .find("_rels")
            .ok_or(anyhow!("Failed to string Prefix path from relation"))?;
        Ok(self.file_path[..rels_position - 1].to_string())
    }

    pub(crate) fn get_target_by_id(&self, relationship_id: &str) -> Option<String> {
        if let Some(record) = self
            .relationships
            .iter()
            .find(|item| item.0 == relationship_id)
        {
            Some(record.1.clone())
        } else {
            None
        }
    }

    /// Get Relation Target based on Type
    /// Note: This will get the first element match the criteria
    pub(crate) fn get_relationship_target_by_type(&self, content_type: &str) -> Option<String> {
        if let Some(relationship) = self
            .relationships
            .iter()
            .find(|item| item.2 == content_type)
        {
            Some(relationship.1.to_string())
        } else {
            None
        }
    }

    /// Generate Next Relationship Id to add
    fn get_next_relationship_id(&self) -> String {
        let mut children = self.relationships.len() + 1;
        loop {
            if self
                .relationships
                .iter()
                .position(|item| item.0 == format!("rId{}", children))
                .is_some()
            {
                children += 1;
            } else {
                break;
            }
        }
        format!("rId{}", children)
    }
}

/// ####################### Mut Access Functions ####################
impl RelationsPart {
    /// Save the relationship detail to xml document
    pub(crate) fn save_relationship_to_doc(&mut self) -> AnyResult<bool, AnyError> {
        if let Some(xml_tree_ref) = self.xml_document.upgrade() {
            let mut xml_tree = xml_tree_ref
                .try_borrow_mut()
                .context("Failed to pull XML Handle")?;
            let child_count = xml_tree
                .get_root()
                .ok_or(anyhow!("No Root Relationship Element Found"))?
                .get_child_count();
            for relationship in self.relationships.clone() {
                let relationship_element = xml_tree
                    .append_child_mut("Relationship", None)
                    .context("Failed to add relationship element")?;
                let mut attributes = HashMap::new();
                attributes.insert("Id".to_string(), relationship.0);
                attributes.insert("Target".to_string(), relationship.1);
                attributes.insert("Type".to_string(), relationship.2);
                relationship_element
                    .set_attribute_mut(attributes)
                    .context("Failed to set Relationship attributes")?;
            }
            Ok(child_count > 0 || self.relationships.len() > 0)
        } else {
            Err(anyhow!("Failed to Get XML Handle"))
        }
    }

    /// Create new Relation
    pub(crate) fn set_new_relationship_mut(
        &mut self,
        content: &Content,
        file_path: Option<String>,
        file_name: Option<String>,
    ) -> AnyResult<String, AnyError> {
        let next_id = self.get_next_relationship_id();
        self.relationships.push((
            next_id.clone(),
            format!(
                "/{}/{}.{}",
                file_path.unwrap_or(content.default_path.to_string()),
                file_name.unwrap_or(content.default_name.to_string()),
                content.extension
            ),
            content.schemas_type.to_string(),
        ));
        Ok(next_id)
    }

    /// Delete the target file path
    pub(crate) fn delete_relationship_mut(&mut self, file_path: &str) {
        self.relationships.retain(|item| {
            item.1
                != if file_path.starts_with("/") {
                    file_path.to_string()
                } else {
                    format!("/{}", file_path)
                }
        })
    }
}
