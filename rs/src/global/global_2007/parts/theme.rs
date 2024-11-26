use crate::{
    files::{OfficeDocument, XmlDocument, XmlSerializer},
    global_2007::traits::XmlDocumentPart,
};
use anyhow::{Error as AnyError, Result as AnyResult};
use std::{cell::RefCell, rc::Weak};

#[derive(Debug)]
pub struct ThemePart {
    office_document: Weak<RefCell<OfficeDocument>>,
    file_tree: Weak<RefCell<XmlDocument>>,
    file_name: String,
}

impl Drop for ThemePart {
    fn drop(&mut self) {
        if let Some(xml_tree) = self.office_document.upgrade() {
            let _ = xml_tree
                .try_borrow_mut()
                .unwrap()
                .close_xml_document(&self.file_name);
        }
    }
}

impl XmlDocumentPart for ThemePart {
    fn new(
        office_document: Weak<RefCell<OfficeDocument>>,
        file_name: Option<String>,
    ) -> AnyResult<Self, AnyError> {
        let mut local_file_name = "theme/theme1.xml".to_string();
        if let Some(file_name) = file_name {
            local_file_name = file_name.to_string();
        }
        let file_tree = Self::get_xml_document(&office_document, &local_file_name)?;
        Ok(Self {
            office_document,
            file_tree,
            file_name: local_file_name.to_string(),
        })
    }

    fn flush(self) {}

    /// Initialize xml content for this part from base template
    fn initialize_content_xml() -> AnyResult<XmlDocument, AnyError> {
        XmlSerializer::vec_to_xml_doc_tree(include_str!("theme.xml").as_bytes().to_vec())
    }
}

impl ThemePart {}
