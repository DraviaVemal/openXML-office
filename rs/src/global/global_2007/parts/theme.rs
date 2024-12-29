use crate::element_dictionary::COMMON_TYPE_COLLECTION;
use crate::{
    files::{OfficeDocument, XmlDocument, XmlSerializer},
    global_2007::{
        parts::RelationsPart,
        traits::{XmlDocumentPart, XmlDocumentPartCommon},
    },
};
use anyhow::{anyhow, Context, Error as AnyError, Result as AnyResult};
use std::{cell::RefCell, rc::Weak};

#[derive(Debug)]
pub struct ThemePart {
    office_document: Weak<RefCell<OfficeDocument>>,
    _xml_document: Weak<RefCell<XmlDocument>>,
    file_path: String,
}

impl Drop for ThemePart {
    fn drop(&mut self) {
        let _ = self.close_document();
    }
}

impl XmlDocumentPartCommon for ThemePart {
    fn close_document(&mut self) -> AnyResult<(), AnyError>
    where
        Self: Sized,
    {
        if let Some(xml_tree) = self.office_document.upgrade() {
            xml_tree
                .try_borrow_mut()
                .context("Failed to pull XML Handle")?
                .close_xml_document(&self.file_path)?;
        }
        Ok(())
    }
    /// Initialize xml content for this part from base template
    fn initialize_content_xml(
    ) -> AnyResult<(XmlDocument, Option<String>, Option<String>, Option<String>), AnyError> {
        let content = COMMON_TYPE_COLLECTION.get("theme").unwrap();
        Ok((
            XmlSerializer::vec_to_xml_doc_tree(include_str!("theme.xml").as_bytes().to_vec())
                .context("Initializing Theme Failed")?,
            Some(content.content_type.to_string()),
            Some(content.extension.to_string()),
            Some(content.extension_type.to_string()),
        ))
    }
}

/// ######################### Train implementation of XML Part - Only accessible within crate ##############
impl XmlDocumentPart for ThemePart {
    fn new(
        office_document: Weak<RefCell<OfficeDocument>>,
        parent_relationship_part: Weak<RefCell<RelationsPart>>,
        _: Option<&str>,
    ) -> AnyResult<Self, AnyError> {
        let file_name = Self::get_theme_file_name(&parent_relationship_part)
            .context("Failed to pull theme file name")?
            .to_string();
        let xml_document = Self::get_xml_document(&office_document, &file_name)?;
        Ok(Self {
            office_document,
            _xml_document: xml_document,
            file_path: file_name.to_string(),
        })
    }
}

impl ThemePart {
    fn get_theme_file_name(
        relations_part: &Weak<RefCell<RelationsPart>>,
    ) -> AnyResult<String, AnyError> {
        let theme_content = COMMON_TYPE_COLLECTION.get("theme").unwrap();
        if let Some(relations_part) = relations_part.upgrade() {
            Ok(relations_part
                .try_borrow_mut()
                .context("Failed to pull relationship connection")?
                .get_relationship_target_by_type_mut(
                    &theme_content.schemas_type,
                    theme_content,
                    Some(format!("xl/{}", theme_content.default_path)),
                    None,
                )
                .context("Pull Path From Existing File Failed")?)
        } else {
            Err(anyhow!("Failed to upgrade relation part"))
        }
    }
}
