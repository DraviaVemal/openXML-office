use crate::files::{OfficeDocument, XmlElement};
use anyhow::{Context, Error as AnyError, Ok, Result as AnyResult};
use std::{
    cell::RefCell,
    rc::{Rc, Weak},
};

pub trait XmlDocumentPart {
    /// Create new object with file connector handle
    fn new(
        office_document: &Rc<RefCell<OfficeDocument>>,
        file_name: Option<String>,
    ) -> AnyResult<Self, AnyError>
    where
        Self: Sized;
    /// Save the current file state
    fn flush(self);
    /// Get content of the current xml
    fn get_xml_tree(
        office_document: &Rc<RefCell<OfficeDocument>>,
        file_name: &str,
    ) -> AnyResult<Weak<RefCell<XmlElement>>, AnyError> {
        let xml_tree: XmlElement = if let Some(xml_tree) = office_document
            .borrow()
            .get_xml_tree(file_name)
            .context(format!("XML Tree Parsing Failed for File : {}", file_name))?
        {
            xml_tree
        } else {
            Self::initialize_content_xml().context("Initial XML element parsing failed")?
        };
        Ok(office_document
            .try_borrow_mut()
            .context("Getting XML Tree Handle Failed")?
            .get_xml_tree_ref(file_name, xml_tree))
    }
    /// Initialize the content if not already exist
    fn initialize_content_xml() -> AnyResult<XmlElement, AnyError>;
}
