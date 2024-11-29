use crate::{
    files::{OfficeDocument, XmlDocument},
    global_2007::traits::XmlDocumentPart,
};
use anyhow::{Context, Error as AnyError, Result as AnyResult};
use std::{cell::RefCell, collections::HashMap, rc::Weak};

#[derive(Debug)]
pub struct Style {
    office_document: Weak<RefCell<OfficeDocument>>,
    xml_document: Weak<RefCell<XmlDocument>>,
    file_path: String,
}

impl Drop for Style {
    fn drop(&mut self) {
        if let Some(xml_tree) = self.office_document.upgrade() {
            let _ = xml_tree
                .try_borrow_mut()
                .unwrap()
                .close_xml_document(&self.file_path);
        }
    }
}

impl XmlDocumentPart for Style {
    fn new(
        office_document: Weak<RefCell<OfficeDocument>>,
        file_path: Option<String>,
    ) -> AnyResult<Self, AnyError> {
        let file_path = file_path.unwrap_or("xl/styles.xml".to_string());
        let xml_document = Self::get_xml_document(&office_document, &file_path)?;
        Ok(Self {
            office_document,
            xml_document,
            file_path,
        })
    }

    fn flush(self) {}

    /// Initialize xml content for this part from base template
    fn initialize_content_xml() -> AnyResult<XmlDocument, AnyError> {
        let mut attributes: HashMap<String, String> = HashMap::new();
        attributes.insert(
            "xmlns".to_string(),
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main".to_string(),
        );
        attributes.insert(
            "xmlns:mc".to_string(),
            "http://schemas.openxmlformats.org/markup-compatibility/2006".to_string(),
        );
        let mut xml_document = XmlDocument::new();
        xml_document
            .create_root_mut("styleSheet")
            .context("Create XML Root Element Failed")?
            .set_attribute_mut(attributes)
            .context("Set Attribute Failed")?;
        Ok(xml_document)
    }
}
