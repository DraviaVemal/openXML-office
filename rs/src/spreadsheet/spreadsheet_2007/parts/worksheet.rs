use crate::global_2007::traits::{XmlDocumentPartCommon, XmlDocumentServicePart};
use crate::{
    files::{OfficeDocument, XmlDocument, XmlSerializer},
    spreadsheet_2007::{
        models::{ColumnCell, ColumnProperties, RowProperties},
        services::CommonServices,
    },
};
use anyhow::{Error as AnyError, Result as AnyResult};
use std::{cell::RefCell, rc::Weak};

#[derive(Debug)]
pub struct WorkSheet {
    office_document: Weak<RefCell<OfficeDocument>>,
    xml_document: Weak<RefCell<XmlDocument>>,
    common_service: Weak<RefCell<CommonServices>>,
    file_name: String,
}

impl Drop for WorkSheet {
    fn drop(&mut self) {
        let _ = self.close_document();
    }
}

impl XmlDocumentPartCommon for WorkSheet {
    /// Initialize xml content for this part from base template
    fn initialize_content_xml() -> AnyResult<XmlDocument, AnyError> {
        let template_core_properties = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></worksheet>"#;
        XmlSerializer::vec_to_xml_doc_tree(template_core_properties.as_bytes().to_vec())
    }
    fn close_document(&mut self) -> AnyResult<(), AnyError>
    where
        Self: Sized,
    {
        if let Some(xml_tree) = self.office_document.upgrade() {
            xml_tree
                .try_borrow_mut()
                .unwrap()
                .close_xml_document(&self.file_name)?;
        }
        Ok(())
    }
}

// ############################# Internal Function ######################################
impl XmlDocumentServicePart for WorkSheet {
    /// Create New object for the group
    fn new(
        office_document: Weak<RefCell<OfficeDocument>>,
        common_service: Weak<RefCell<CommonServices>>,
        sheet_name: Option<String>,
    ) -> AnyResult<Self, AnyError> {
        let mut file_name: String = "xl/worksheets/sheet1.xml".to_string();
        if let Some(sheet_name) = sheet_name {
            file_name = sheet_name.to_string();
        }
        let xml_document = Self::get_xml_document(&office_document, &file_name)?;
        Ok(Self {
            office_document,
            xml_document,
            common_service,
            file_name,
        })
    }
}

// ##################################### Feature Function ################################
impl WorkSheet {
    pub fn set_column_mut(&mut self, column_id: &usize, column_properties: ColumnProperties) -> () {
    }

    pub fn set_row_mut(
        &mut self,
        cell_id: &usize,
        column_cell: Vec<ColumnCell>,
        row_properties: RowProperties,
    ) -> () {
    }
}
