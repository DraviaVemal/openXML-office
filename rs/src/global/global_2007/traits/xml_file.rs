use crate::files::{OfficeDocument, XmlDocument};
use crate::spreadsheet_2007::services::CommonServices;
use anyhow::{anyhow, Context, Error as AnyError, Result as AnyResult};
use std::{cell::RefCell, rc::Weak};

pub(crate) trait XmlDocumentPartCommon {
    /// Save the current file state
    fn flush(mut self) -> AnyResult<(), AnyError>
    where
        Self: Sized,
    {
        self.close_document()
    }

    fn delete_xml_document_mut(
        &mut self,
        office_document: &Weak<RefCell<OfficeDocument>>,
        file_name: &str,
    ) -> AnyResult<(), AnyError> {
        office_document
            .upgrade()
            .ok_or(anyhow!("Document Upgrade Handled Failed"))
            .context("XML Document Read Failed")?
            .try_borrow_mut()
            .context("Getting XML Tree Handle Failed")?
            .delete_document_mut(file_name)
    }

    fn close_document(&mut self) -> AnyResult<(), AnyError>
    where
        Self: Sized;
    /// Get content of the current xml
    fn get_xml_document(
        office_document: &Weak<RefCell<OfficeDocument>>,
        file_name: &str,
    ) -> AnyResult<Weak<RefCell<XmlDocument>>, AnyError> {
        let (xml_document, content_type) = if let Some((xml_document, content_type)) =
            office_document
                .upgrade()
                .ok_or(anyhow!("Document Upgrade Handled Failed"))
                .context("XML Document Read Failed")?
                .borrow()
                .get_xml_tree(file_name)
                .context(format!("XML Tree Parsing Failed for File : {}", file_name))?
        {
            (xml_document, content_type)
        } else {
            Self::initialize_content_xml().context("Initial XML element parsing failed")?
        };
        Ok(office_document
            .upgrade()
            .ok_or(anyhow!("Document Upgrade Handled Failed"))
            .context("XML Document Read Failed")?
            .try_borrow_mut()
            .context("Getting XML Tree Handle Failed")?
            .get_xml_document_ref(file_name, content_type, xml_document))
    }
    /// Initialize the content if not already exist
    fn initialize_content_xml() -> AnyResult<(XmlDocument, Option<String>), AnyError>;
}

pub(crate) trait XmlDocumentServicePart: XmlDocumentPartCommon {
    /// Create new object with file connector handle
    fn new(
        office_document: Weak<RefCell<OfficeDocument>>,
        common_service: Weak<RefCell<CommonServices>>,
        file_name: &str,
    ) -> AnyResult<Self, AnyError>
    where
        Self: Sized;
}
#[warn(drop_bounds)]
pub(crate) trait XmlDocumentPart: XmlDocumentPartCommon {
    /// Create new object with file connector handle
    fn new(
        office_document: Weak<RefCell<OfficeDocument>>,
        file_name: &str,
    ) -> AnyResult<Self, AnyError>
    where
        Self: Sized;
}
