use crate::{
    files::{OfficeDocument, XmlElement, XmlSerializer},
    global_2007::traits::XmlDocumentPart,
};
use anyhow::{Error as AnyError, Result as AnyResult};
use std::{
    cell::RefCell,
    rc::{Rc, Weak},
};

#[derive(Debug)]
pub struct CorePropertiesPart {
    office_document: Weak<RefCell<OfficeDocument>>,
    file_tree: Weak<RefCell<XmlElement>>,
    file_name: String,
}

impl Drop for CorePropertiesPart {
    fn drop(&mut self) {
        // Update Last modified date part
        
        // Update the current state to DB before dropping the object
        if let Some(xml_tree) = self.office_document.upgrade() {
            let _ = xml_tree
                .try_borrow_mut()
                .unwrap()
                .close_xml_tree(&self.file_name);
        }
    }
}

impl XmlDocumentPart for CorePropertiesPart {
    fn new(
        office_document: &Rc<RefCell<OfficeDocument>>,
        _: Option<String>,
    ) -> AnyResult<Self, AnyError> {
        let file_name = "docProps/core.xml".to_string();
        let file_tree = Self::get_xml_tree(&office_document, &file_name)?;
        Ok(Self {
            office_document: Rc::downgrade(office_document),
            file_tree,
            file_name,
        })
    }

    /// Save the current file state
    fn flush(self) {}

    /// Initialize xml content for this part from base template
    fn initialize_content_xml() -> AnyResult<XmlElement, AnyError> {
        XmlSerializer::xml_str_to_xml_tree(include_str!("core_properties.xml").as_bytes().to_vec())
    }
}

impl CorePropertiesPart {}
