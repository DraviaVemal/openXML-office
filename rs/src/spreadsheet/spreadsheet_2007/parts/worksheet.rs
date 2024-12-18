use crate::{
    files::{OfficeDocument, XmlDocument, XmlSerializer},
    global_2007::{
        parts::RelationsPart,
        traits::{XmlDocumentPartCommon, XmlDocumentServicePart},
    },
    reference_dictionary::EXCEL_TYPE_COLLECTION,
    spreadsheet_2007::{
        models::{ColumnCell, ColumnProperties, RowProperties},
        services::CommonServices,
    },
};
use anyhow::{Context, Error as AnyError, Result as AnyResult};
use std::{
    cell::RefCell,
    rc::{Rc, Weak},
};

#[derive(Debug)]
pub struct WorkSheet {
    office_document: Weak<RefCell<OfficeDocument>>,
    xml_document: Weak<RefCell<XmlDocument>>,
    common_service: Weak<RefCell<CommonServices>>,
    sheet_relationship_part: Rc<RefCell<RelationsPart>>,
    file_path: String,
}

impl Drop for WorkSheet {
    fn drop(&mut self) {
        let _ = self.close_document();
    }
}

impl XmlDocumentPartCommon for WorkSheet {
    /// Initialize xml content for this part from base template
    fn initialize_content_xml(
    ) -> AnyResult<(XmlDocument, Option<String>, Option<String>, Option<String>), AnyError> {
        let content = EXCEL_TYPE_COLLECTION.get("worksheet").unwrap();
        let template_core_properties = include_str!("worksheet.xml");
        Ok((
            XmlSerializer::vec_to_xml_doc_tree(template_core_properties.as_bytes().to_vec())
                .context("Initializing Worksheet Failed")?,
            Some(content.content_type.to_string()),
            Some(content.extension.to_string()),
            Some(content.extension_type.to_string()),
        ))
    }
    fn close_document(&mut self) -> AnyResult<(), AnyError>
    where
        Self: Sized,
    {
        if let Some(xml_tree) = self.office_document.upgrade() {
            xml_tree
                .try_borrow_mut()
                .context("Failed to Pull XML Handle")?
                .close_xml_document(&self.file_path)?;
        }
        Ok(())
    }
}

// ############################# Internal Function ######################################
impl XmlDocumentServicePart for WorkSheet {
    /// Create New object for the group
    fn new(
        office_document: Weak<RefCell<OfficeDocument>>,
        parent_relationship_part: Weak<RefCell<RelationsPart>>,
        common_service: Weak<RefCell<CommonServices>>,
        file_path: &str,
    ) -> AnyResult<Self, AnyError> {
        let xml_document = Self::get_xml_document(&office_document, &file_path)?;
        let sheet_relationship_part = Rc::new(RefCell::new(
            RelationsPart::new(
                office_document.clone(),
                &format!(
                    "{}/_rels/{}.rels",
                    &file_path[..file_path.rfind("/").unwrap()],
                    file_path.rsplit("/").next().unwrap()
                ),
            )
            .context("Creating Relation ship part for workbook failed.")?,
        ));
        Ok(Self {
            office_document,
            xml_document,
            common_service,
            sheet_relationship_part,
            file_path: file_path.to_string(),
        })
    }
}

// ##################################### Feature Function ################################
impl WorkSheet {
    /// Set Column property
    pub fn set_column_mut(&mut self, column_id: &usize, column_properties: ColumnProperties) -> () {
    }

    /// Set Active sheet internal
    pub(crate) fn set_active_sheet_mut(&mut self) {}

    /// Set Active cell of the current sheet
    pub fn set_active_cell_mut(&mut self) {}

    /// Set data for same row multiple columns along with row property
    pub fn set_row_mut(
        &mut self,
        cell_id: &usize,
        column_cell: Vec<ColumnCell>,
        row_properties: RowProperties,
    ) -> () {
    }

    /// Add hyper link to current document
    pub fn add_hyperlink_mut(&mut self) {}

    /// Set Cell Range to merge
    pub fn set_merge_cell_mut(&mut self) {}

    /// List all Cell Range merged
    pub fn list_merge_cell_(&mut self) {}

    /// Remove merged cell range
    pub fn remove_merge_cell_mut(&mut self) {}

    /// Delete Current sheet and all its components
    pub fn delete_sheet_mut(&mut self) -> AnyResult<(), AnyError> {
        // TODO : Delete all items in relationship
        if let Some(xml_tree) = self.office_document.upgrade() {
            xml_tree
                .try_borrow_mut()
                .context("Failed to Pull XML Handle")?
                .delete_document_mut(&self.file_path)?;
        }
        Ok(())
    }
}
