use crate::{
    element_dictionary::COMMON_TYPE_COLLECTION,
    files::{OfficeDocument, XmlDocument, XmlSerializer},
    global_2007::{
        parts::RelationsPart,
        traits::{XmlDocumentPart, XmlDocumentPartCommon},
    },
};
use anyhow::{anyhow, Context, Error as AnyError, Result as AnyResult};
use chrono::Utc;
use std::{cell::RefCell, rc::Weak};

#[derive(Debug)]
pub struct CorePropertiesPart {
    office_document: Weak<RefCell<OfficeDocument>>,
    xml_document: Weak<RefCell<XmlDocument>>,
    file_path: String,
}

impl Drop for CorePropertiesPart {
    fn drop(&mut self) {
        let _ = self.close_document();
    }
}

impl XmlDocumentPartCommon for CorePropertiesPart {
    /// Initialize xml content for this part from base template
    fn initialize_content_xml() -> AnyResult<(XmlDocument, Option<String>, String, String), AnyError>
    {
        let content = COMMON_TYPE_COLLECTION.get("docProps_core").unwrap();
        Ok((
            XmlSerializer::vec_to_xml_doc_tree(
                include_str!("core_properties.xml").as_bytes().to_vec(),
            )
            .context("Initializing Core Property Failed")?,
            Some(content.content_type.to_string()),
            content.extension.to_string(),
            content.extension_type.to_string(),
        ))
    }

    fn close_document(&mut self) -> AnyResult<(), AnyError>
    where
        Self: Sized,
    {
        // Update Last modified date part
        if let Some(xml_document_ref) = self.xml_document.upgrade() {
            let mut xml_document = xml_document_ref
                .try_borrow_mut()
                .context("Failed to Pull Office document")?;
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
            match xml_document
                .get_first_element_mut(vec!["cp:coreProperties", "dcterms:created"], None)
            {
                Ok(result) => {
                    if let Some(element) = result {
                        if !element.has_value() {
                            element.set_value_mut(
                                Utc::now().to_rfc3339_opts(chrono::SecondsFormat::Secs, true),
                            );
                        }
                    }
                }
                Err(_) => (),
            }
        }
        // Update the current state to DB before dropping the object
        if let Some(xml_tree) = self.office_document.upgrade() {
            xml_tree
                .try_borrow_mut()
                .context("Failed to Pull XML Handle")?
                .close_xml_document(&self.file_path)?;
        }
        Ok(())
    }
}

/// ######################### Train implementation of XML Part - Only accessible within crate ##############
impl XmlDocumentPart for CorePropertiesPart {
    fn new(
        office_document: Weak<RefCell<OfficeDocument>>,
        parent_relationship_part: Weak<RefCell<RelationsPart>>,
    ) -> AnyResult<Self, AnyError> {
        let file_name = Self::get_core_properties_file_name(&parent_relationship_part)
            .context("Failed to pull Core Property file name")?
            .to_string();
        let xml_document = Self::get_xml_document(&office_document, &file_name)?;
        Ok(Self {
            office_document,
            xml_document,
            file_path: file_name,
        })
    }
}

impl CorePropertiesPart {
    fn get_core_properties_file_name(
        relations_part: &Weak<RefCell<RelationsPart>>,
    ) -> AnyResult<String, AnyError> {
        let relationship_content = COMMON_TYPE_COLLECTION.get("docProps_core").unwrap();
        if let Some(relations_part) = relations_part.upgrade() {
            relations_part
                .try_borrow_mut()
                .context("Failed to pull relationship connection")?
                .get_relationship_target_by_type_mut(
                    &relationship_content.schemas_type,
                    relationship_content,
                    None,
                    None,
                )
        } else {
            Err(anyhow!("Failed to upgrade relation part"))
        }
    }
}
