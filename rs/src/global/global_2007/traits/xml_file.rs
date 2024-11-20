use crate::files::{OfficeDocument, XmlElement};
use anyhow::{Context, Error as AnyError, Ok, Result as AnyResult};
use std::{cell::RefCell, rc::Rc};

pub trait XmlDocumentPart {
    /// Create new object with file connector handle
    fn new(
        office_document: &Rc<RefCell<OfficeDocument>>,
        file_name: Option<&str>,
    ) -> AnyResult<Self, AnyError>
    where
        Self: Sized;
    /// Save the current file state
    fn flush(self);
    /// Get content of the current xml
    fn get_xml_tree(
        office_document: &Rc<RefCell<OfficeDocument>>,
        file_name: &str,
    ) -> AnyResult<XmlElement, AnyError> {
        let xml_tree: Option<XmlElement> = office_document
            .borrow()
            .get_xml_tree(file_name)
            .context(format!("XML Tree Parsing Failed for File : {}", file_name))?;
        if let Some(results) = xml_tree {
            Ok(results)
        } else {
            let tree =
                Self::initialize_content_xml().context("Initial XML element parsing failed")?;
            Ok(tree)
        }
    }
    /// Initialize the content if not already exist
    fn initialize_content_xml() -> AnyResult<XmlElement, AnyError>;
}
