use crate::files::{OfficeDocument, XmlDocument};
use crate::spreadsheet_2007::services::CommonServices;
use anyhow::{anyhow, Context, Error as AnyError, Ok, Result as AnyResult};
use std::{cell::RefCell, rc::Weak};

pub trait XmlDocumentPartCommon {
    /// Save the current file state
    fn flush(self) {}
    /// Get content of the current xml
    fn get_xml_document(
        office_document: &Weak<RefCell<OfficeDocument>>,
        file_name: &str,
    ) -> AnyResult<Weak<RefCell<XmlDocument>>, AnyError> {
        let xml_document: XmlDocument = if let Some(xml_document) = office_document
            .upgrade()
            .ok_or(anyhow!("Document Upgrade Handled Failed"))
            .context("XML Document Read Failed")?
            .borrow()
            .get_xml_tree(file_name)
            .context(format!("XML Tree Parsing Failed for File : {}", file_name))?
        {
            xml_document
        } else {
            Self::initialize_content_xml().context("Initial XML element parsing failed")?
        };
        Ok(office_document
            .upgrade()
            .ok_or(anyhow!("Document Upgrade Handled Failed"))
            .context("XML Document Read Failed")?
            .try_borrow_mut()
            .context("Getting XML Tree Handle Failed")?
            .get_xml_document_ref(file_name, xml_document))
    }
    /// Initialize the content if not already exist
    fn initialize_content_xml() -> AnyResult<XmlDocument, AnyError>;
}

pub trait XmlDocumentServicePart: XmlDocumentPartCommon {
    /// Create new object with file connector handle
    fn new(
        office_document: Weak<RefCell<OfficeDocument>>,
        common_service: Weak<RefCell<CommonServices>>,
        file_name: Option<String>,
    ) -> AnyResult<Self, AnyError>
    where
        Self: Sized;
}

pub trait XmlDocumentPart: XmlDocumentPartCommon + Drop {
    /// Create new object with file connector handle
    fn new(
        office_document: Weak<RefCell<OfficeDocument>>,
        file_name: Option<String>,
    ) -> AnyResult<Self, AnyError>
    where
        Self: Sized;
}
