use crate::{
    files::{OfficeDocument, XmlDocument, XmlSerializer},
    global_2007::{
        parts::RelationsPart,
        traits::{XmlDocumentPart, XmlDocumentPartCommon},
    },
    reference_dictionary::COMMON_TYPE_COLLECTION,
};
use anyhow::{Context, Error as AnyError, Result as AnyResult};
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
    fn initialize_content_xml() -> AnyResult<(XmlDocument, Option<String>), AnyError> {
        Ok((
            XmlSerializer::vec_to_xml_doc_tree(
                include_str!("core_properties.xml").as_bytes().to_vec(),
            )
            .context("Initializing Core Property Failed")?,
            Some(
                COMMON_TYPE_COLLECTION
                    .get("docProps_core")
                    .unwrap()
                    .content_type
                    .to_string(),
            ),
        ))
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
                .context("Failed to Pull XML Handle")?
                .close_xml_document(&self.file_name)?;
        }
        Ok(())
    }
}

/// ######################### Train implementation of XML Part - Only accessible within crate ##############
impl XmlDocumentPart for CorePropertiesPart {
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

impl CorePropertiesPart {
    pub(crate) fn create_core_properties(
        relations_part: &mut RelationsPart,
        office_document: Weak<RefCell<OfficeDocument>>,
    ) -> AnyResult<CorePropertiesPart, AnyError> {
        let core_properties_path: Option<String> = relations_part.get_relationship_target_by_type("http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties").context("Parsing Core Properties path failed")?;
        Ok(if let Some(part_path) = core_properties_path {
            CorePropertiesPart::new(office_document, &part_path)
                .context("Load CorePart for Existing file failed")?
        } else {
            let relationship_content = COMMON_TYPE_COLLECTION.get("docProps_core").unwrap();
            relations_part
                .set_new_relationship_mut(relationship_content, None, None)
                .context("Setting New Theme Relationship Failed.")?;
            CorePropertiesPart::new(office_document, "docProps/core.xml")
                .context("Load CorePart for Existing file failed")?
        })
    }
}
