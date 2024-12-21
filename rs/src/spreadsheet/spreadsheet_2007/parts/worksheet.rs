use crate::{
    files::{OfficeDocument, XmlDocument, XmlSerializer},
    global_2007::{parts::RelationsPart, traits::XmlDocumentPartCommon},
    reference_dictionary::EXCEL_TYPE_COLLECTION,
    spreadsheet_2007::{
        models::{ColumnCell, ColumnProperties, RowProperties},
        services::CommonServices,
    },
};
use anyhow::{anyhow, Context, Error as AnyError, Result as AnyResult};
use std::{
    cell::RefCell,
    rc::{Rc, Weak},
};

#[derive(Debug)]
pub struct WorkSheet {
    office_document: Weak<RefCell<OfficeDocument>>,
    xml_document: Weak<RefCell<XmlDocument>>,
    common_service: Weak<RefCell<CommonServices>>,
    workbook_relationship_part: Weak<RefCell<RelationsPart>>,
    sheet_collection: Weak<RefCell<Vec<(String, String)>>>,
    sheet_relationship_part: Rc<RefCell<RelationsPart>>,
    file_path: String,
    sheet_name: String,
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
impl WorkSheet {
    /// Create New object for the group
    pub(crate) fn new(
        office_document: Weak<RefCell<OfficeDocument>>,
        sheet_collection: Weak<RefCell<Vec<(String, String)>>>,
        workbook_relationship_part: Weak<RefCell<RelationsPart>>,
        common_service: Weak<RefCell<CommonServices>>,
        sheet_name: Option<String>,
    ) -> AnyResult<Self, AnyError> {
        let (file_path, sheet_name) = Self::get_sheet_file_name(
            sheet_name,
            &office_document,
            &sheet_collection,
            &workbook_relationship_part,
        )
        .context("Failed to pull calc chain file name")?;
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
            workbook_relationship_part,
            sheet_relationship_part,
            sheet_collection,
            file_path: file_path.to_string(),
            sheet_name,
        })
    }
}

impl WorkSheet {
    fn get_sheet_file_name(
        sheet_name: Option<String>,
        office_document: &Weak<RefCell<OfficeDocument>>,
        sheet_collection: &Weak<RefCell<Vec<(String, String)>>>,
        workbook_relationship_part: &Weak<RefCell<RelationsPart>>,
    ) -> AnyResult<(String, String), AnyError> {
        let worksheet_content = EXCEL_TYPE_COLLECTION.get("worksheet").unwrap();
        if let Some(sheet_collection) = sheet_collection.upgrade() {
            if let Some(workbook_relationship_part) = workbook_relationship_part.upgrade() {
                let mut sheet_count = sheet_collection
                    .try_borrow()
                    .context("Failed to pull Sheet Name Collection")?
                    .len()
                    + 1;
                let relative_path = workbook_relationship_part
                    .try_borrow_mut()
                    .context("Failed to pull relationship connection")?
                    .get_relative_path()
                    .context("Get Relative Path for Part File")?;
                if let Some(office_doc) = office_document.upgrade() {
                    let document = office_doc
                        .try_borrow()
                        .context("Failed to Borrow Document")?;
                    loop {
                        if document
                            .check_file_exist(format!(
                                "{}/{}/{}{}.{}",
                                relative_path,
                                worksheet_content.default_path,
                                worksheet_content.default_name,
                                sheet_count,
                                worksheet_content.extension
                            ))
                            .context("Failed to Check the File Exist")?
                        {
                            sheet_count += 1;
                        } else {
                            break;
                        }
                    }
                }
                let file_path = format!("{}/{}", relative_path, worksheet_content.default_path);
                let sheet_name = format!(
                    "{}",
                    sheet_name.clone().unwrap_or(format!(
                        "{}{}",
                        worksheet_content.default_name, &sheet_count
                    ))
                );
                let relationship_id = workbook_relationship_part
                    .try_borrow_mut()
                    .context("Failed to Get Relationship Handle")?
                    .set_new_relationship_mut(
                        worksheet_content,
                        Some(file_path.clone()),
                        Some(format!(
                            "{}{}",
                            worksheet_content.default_name, &sheet_count
                        )),
                    )
                    .context("Setting New Calculation Chain Relationship Failed.")?;
                sheet_collection
                    .try_borrow_mut()
                    .context("Failed To pull Sheet Collection Handle")?
                    .push((sheet_name.clone(), relationship_id));
                return Ok((
                    format!(
                        "{}/{}{}.{}",
                        file_path,
                        worksheet_content.default_name,
                        &sheet_count,
                        worksheet_content.extension
                    ),
                    sheet_name,
                ));
            }
        }
        Err(anyhow!("Failed to upgrade relation part"))
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
    pub fn delete_sheet_mut(self) -> AnyResult<(), AnyError> {
        if let Some(sheet_collection) = self.sheet_collection.upgrade() {
            sheet_collection
                .try_borrow_mut()
                .context("Failed to pull Sheets Collection")?
                .retain(|item| item.0 != self.sheet_name);
        }
        if let Some(workbook_relationship_part) = self.workbook_relationship_part.upgrade() {
            workbook_relationship_part
                .try_borrow_mut()
                .context("Failed to pull workbook relationship handle")?
                .delete_relationship_mut(&self.file_path);
        }
        if let Some(xml_tree) = self.office_document.upgrade() {
            xml_tree
                .try_borrow_mut()
                .context("Failed to Pull XML Handle")?
                .delete_document_mut(&self.file_path)
                .context("Failed to delete the document from database")?;
        }
        self.flush().context("Failed to flush the worksheet")?;
        Ok(())
    }
}
