use crate::{
    files::{OfficeDocument, XmlDocument, XmlSerializer},
    spreadsheet_2007::services::CommonServices,
};
use anyhow::{anyhow, Context, Error as AnyError, Result as AnyResult};
use std::{cell::RefCell, rc::Weak};

#[derive(Debug)]
pub struct WorkSheet {
    office_document: Weak<RefCell<OfficeDocument>>,
    file_tree: Weak<RefCell<XmlDocument>>,
    common_service: Weak<RefCell<CommonServices>>,
    file_name: String,
}

impl Drop for WorkSheet {
    fn drop(&mut self) {
        if let Some(xml_tree) = self.office_document.upgrade() {
            let _ = xml_tree
                .try_borrow_mut()
                .unwrap()
                .close_xml_document(&self.file_name);
        }
    }
}

// ############################# Internal Function ######################################
impl WorkSheet {
    /// Create New object for the group
    pub(crate) fn new(
        office_document: Weak<RefCell<OfficeDocument>>,
        common_service: Weak<RefCell<CommonServices>>,
        sheet_name: Option<String>,
    ) -> AnyResult<Self, AnyError> {
        let mut file_name: String = "xl/worksheets/sheet1.xml".to_string();
        if let Some(sheet_name) = sheet_name {
            file_name = sheet_name.to_string();
        }
        let file_tree = Self::get_xml_document(&office_document, &file_name)?;
        return Ok(Self {
            office_document,
            file_tree,
            common_service,
            file_name,
        });
    }

    pub(crate) fn flush(self) {}

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

    /// Initialize xml content for this part from base template
    fn initialize_content_xml() -> AnyResult<XmlDocument, AnyError> {
        let template_core_properties = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></worksheet>"#;
        XmlSerializer::vec_to_xml_doc_tree(template_core_properties.as_bytes().to_vec())
    }
}

// ##################################### Feature Function ################################
impl WorkSheet {
    
}
