use crate::{
    files::{OfficeDocument, XmlDocument, XmlSerializer},
    global_2007::{
        parts::{RelationsPart, ThemePart},
        traits::{XmlDocumentPart, XmlDocumentPartCommon, XmlDocumentServicePart},
    },
    reference_dictionary::{COMMON_TYPE_COLLECTION, EXCEL_TYPE_COLLECTION},
    spreadsheet_2007::{
        parts::WorkSheet,
        services::{CalculationChain, CommonServices, ShareString, Style},
    },
};
use anyhow::{anyhow, Context, Error as AnyError, Result as AnyResult};
use std::{
    cell::RefCell,
    collections::HashMap,
    path::Path,
    rc::{Rc, Weak},
};

#[derive(Debug)]
pub struct WorkbookPart {
    office_document: Weak<RefCell<OfficeDocument>>,
    xml_document: Weak<RefCell<XmlDocument>>,
    file_path: String,
    common_service: Rc<RefCell<CommonServices>>,
    relations_part: RelationsPart,
    theme_part: ThemePart,
    relation_path: String,
    /// This contain the sheet name and order along with relationId
    sheet_names: Vec<(String, String)>,
}

impl Drop for WorkbookPart {
    fn drop(&mut self) {
        let _ = self.close_document();
    }
}

impl XmlDocumentPartCommon for WorkbookPart {
    /// Initialize xml content for this part from base template
    fn initialize_content_xml() -> AnyResult<(XmlDocument, Option<String>), AnyError> {
        let template_core_properties = include_str!("workbook.xml");
        Ok((
            XmlSerializer::vec_to_xml_doc_tree(template_core_properties.as_bytes().to_vec())
                .context("Initializing Workbook Failed")?,
            Some(
                EXCEL_TYPE_COLLECTION
                    .get("workbook")
                    .unwrap()
                    .content_type
                    .to_string(),
            ),
        ))
    }

    fn close_document(&mut self) -> AnyResult<(), AnyError>
    where
        Self: Sized,
    {
        // Write Sheet Records to Workbook
        if let Some(xml_document_mut) = self.xml_document.upgrade() {
            let mut xml_doc_mut = xml_document_mut
                .try_borrow_mut()
                .context("Borrow XML Document Failed")?;
            let sheets_id = xml_doc_mut
                .insert_children_after_tag_mut("sheets", "bookViews", None)
                .context("Create Sheets Node Failed")?
                .get_id();
            let mut sheet_count = 1;
            for (sheet_display_name, relationship_id) in &self.sheet_names {
                let sheet = xml_doc_mut
                    .append_child_mut("sheet", Some(&sheets_id))
                    .context("Create Sheet Node Failed")?;
                let mut attributes: HashMap<String, String> = HashMap::new();
                attributes.insert("name".to_string(), sheet_display_name.to_string());
                attributes.insert("sheetId".to_string(), sheet_count.to_string());
                attributes.insert("r:id".to_string(), relationship_id.to_string());
                sheet
                    .set_attribute_mut(attributes)
                    .context("Sheet Attributes Failed")?;
                sheet_count += 1;
            }
        }
        if let Some(xml_tree) = self.office_document.upgrade() {
            xml_tree
                .try_borrow_mut()
                .unwrap()
                .close_xml_document(&self.file_path)?;
        }
        Ok(())
    }
}

/// ######################### Train implementation of XML Part - Only accessible within crate ##############
impl XmlDocumentPart for WorkbookPart {
    /// Create workbook
    fn new(
        office_document: Weak<RefCell<OfficeDocument>>,
        file_path: &str,
    ) -> AnyResult<Self, AnyError> {
        let mut file_tree = Self::get_xml_document(&office_document, file_path)?;
        let relation_path = Path::new(&file_path)
            .parent()
            .unwrap_or(Path::new(""))
            .display()
            .to_string();
        let mut relations_part = RelationsPart::new(
            office_document.clone(),
            &format!("{}/_rels/workbook.xml.rels", relation_path),
        )
        .context("Creating Relation ship part for workbook failed.")?;
        let theme_part =
            Self::create_theme_part(&mut relations_part, office_document.clone(), &relation_path)
                .context("Loading Theme Part Failed")?;
        let share_string =
            Self::create_share_string(&mut relations_part, office_document.clone(), &relation_path)
                .context("Loading Share String Failed")?;
        let calculation_chain = Self::create_calculation_chain(
            &mut relations_part,
            office_document.clone(),
            &relation_path,
        )
        .context("Loading Calculation Chain Failed")?;
        let style =
            Self::create_style_part(&mut relations_part, office_document.clone(), &relation_path)
                .context("Loading Style Part Failed")?;
        let common_service = Rc::new(RefCell::new(CommonServices::new(
            calculation_chain,
            share_string,
            style,
        )));
        let sheet_names =
            Self::load_sheet_names(&mut file_tree).context("Loading Sheet Names Failed")?;
        Ok(Self {
            office_document,
            xml_document: file_tree,
            file_path: file_path.to_string(),
            common_service,
            relations_part,
            theme_part,
            sheet_names,
            relation_path,
        })
    }
}

// ############################# Internal Function ######################################
// ############################# mut Function ######################################
impl WorkbookPart {
    fn create_theme_part(
        relations_part: &mut RelationsPart,
        office_document: Weak<RefCell<OfficeDocument>>,
        relation_path: &str,
    ) -> AnyResult<ThemePart, AnyError> {
        let theme_content = COMMON_TYPE_COLLECTION.get("theme").unwrap();
        let theme_part_path = relations_part
            .get_relationship_target_by_type(&theme_content.schemas_type)
            .context("Parsing Theme part path failed")?;
        Ok(if let Some(part_path) = theme_part_path {
            ThemePart::new(
                office_document.clone(),
                &format!("{}/{}", relation_path, &part_path),
            )
            .context("Creating Theme part for workbook failed")?
        } else {
            relations_part
                .set_new_relationship_mut(theme_content, None)
                .context("Setting New Theme Relationship Failed.")?;
            ThemePart::new(
                office_document.clone(),
                &format!(
                    "{}/{}.{}",
                    theme_content.default_path, theme_content.default_name, theme_content.extension
                ),
            )
            .context("Creating Theme part for workbook failed")?
        })
    }
    fn create_share_string(
        relations_part: &mut RelationsPart,
        office_document: Weak<RefCell<OfficeDocument>>,
        relation_path: &str,
    ) -> AnyResult<ShareString, AnyError> {
        let share_string_content = EXCEL_TYPE_COLLECTION.get("share_string").unwrap();
        let share_string_path = relations_part
            .get_relationship_target_by_type(&share_string_content.schemas_type)
            .context("Parsing Theme part path failed")?;
        Ok(if let Some(part_path) = share_string_path {
            ShareString::new(
                office_document.clone(),
                &format!("{}/{}", relation_path, &part_path),
            )
            .context("Share String Service Object Creation Failure")?
        } else {
            relations_part
                .set_new_relationship_mut(share_string_content, None)
                .context("Setting New Theme Relationship Failed.")?;
            ShareString::new(
                office_document.clone(),
                &format!(
                    "{}/{}.{}",
                    share_string_content.default_path,
                    share_string_content.default_name,
                    share_string_content.extension
                ),
            )
            .context("Share String Service Object Creation Failure")?
        })
    }
    fn create_calculation_chain(
        relations_part: &mut RelationsPart,
        office_document: Weak<RefCell<OfficeDocument>>,
        relation_path: &str,
    ) -> AnyResult<CalculationChain, AnyError> {
        let calc_chain_content = EXCEL_TYPE_COLLECTION.get("calc_chain").unwrap();
        let calculation_chain_path = relations_part
            .get_relationship_target_by_type(&calc_chain_content.schemas_type)
            .context("Parsing Theme part path failed")?;
        Ok(if let Some(part_path) = calculation_chain_path {
            CalculationChain::new(
                office_document.clone(),
                &format!("{}/{}", relation_path, &part_path),
            )
            .context("Calculation Chain Service Object Creation Failure")?
        } else {
            relations_part
                .set_new_relationship_mut(calc_chain_content, None)
                .context("Setting New Theme Relationship Failed.")?;
            CalculationChain::new(
                office_document.clone(),
                &format!(
                    "{}/{}.{}",
                    calc_chain_content.default_path,
                    calc_chain_content.default_name,
                    calc_chain_content.extension
                ),
            )
            .context("Calculation Chain Service Object Creation Failure")?
        })
    }
    fn create_style_part(
        relations_part: &mut RelationsPart,
        office_document: Weak<RefCell<OfficeDocument>>,
        relation_path: &str,
    ) -> AnyResult<Style, AnyError> {
        let style_content = EXCEL_TYPE_COLLECTION.get("style").unwrap();
        let style_path = relations_part
            .get_relationship_target_by_type(&style_content.schemas_type)
            .context("Parsing Theme part path failed")?;
        Ok(if let Some(part_path) = style_path {
            Style::new(
                office_document.clone(),
                &format!("{}/{}", relation_path, &part_path),
            )
            .context("Style Service Object Creation Failure")?
        } else {
            relations_part
                .set_new_relationship_mut(style_content, None)
                .context("Setting New Theme Relationship Failed.")?;
            Style::new(
                office_document.clone(),
                &format!(
                    "{}/{}.{}",
                    style_content.default_path, style_content.default_name, style_content.extension
                ),
            )
            .context("Style Service Object Creation Failure")?
        })
    }
    fn load_sheet_names(
        xml_document: &mut Weak<RefCell<XmlDocument>>,
    ) -> AnyResult<Vec<(String, String)>, AnyError> {
        let mut sheet_names: Vec<(String, String)> = Vec::new();
        if let Some(xml_doc) = xml_document.upgrade() {
            let mut xml_doc_mut = xml_doc.try_borrow_mut().context("xml doc borrow failed")?;
            if let Some(mut sheets_vec) = xml_doc_mut.pop_elements_by_tag_mut("sheets", None) {
                if let Some(sheets) = sheets_vec.pop() {
                    // Load Sheet from File if exist
                    loop {
                        if let Some(sheet_id) = sheets.pop_child_id_mut() {
                            if let Some(sheet) = xml_doc_mut.pop_element_mut(&sheet_id) {
                                if let Some(attributes) = sheet.get_attribute() {
                                    let name = attributes.get("name").ok_or(anyhow!(
                                        "Error When Trying to read Sheet Details."
                                    ))?;
                                    let r_id = attributes.get("r:id").ok_or(anyhow!(
                                        "Error When Trying to read Sheet Details."
                                    ))?;
                                    sheet_names.push((name.to_string(), r_id.to_string()));
                                }
                            }
                        } else {
                            break;
                        }
                    }
                }
            }
        }
        Ok(sheet_names)
    }
}

// ############################# Feature Function ######################################
// ############################# mut Function ######################################
impl WorkbookPart {
    pub(crate) fn add_sheet(
        &mut self,
        sheet_name: Option<String>,
    ) -> AnyResult<WorkSheet, AnyError> {
        let mut sheet_count = self.sheet_names.len() + 1;
        loop {
            if let Some(office_doc) = self.office_document.upgrade() {
                let document = office_doc
                    .try_borrow()
                    .context("Failed to Borrow Document")?;
                if document
                    .check_file_exist(format!(
                        "{}/worksheets/sheet{}.xml",
                        self.relation_path, sheet_count
                    ))
                    .context("Failed to Check the File Exist")?
                {
                    sheet_count += 1;
                } else {
                    break;
                }
            }
            break;
        }
        let sheet_name = sheet_name.unwrap_or(format!("sheet{}", sheet_count));
        let worksheet_content = EXCEL_TYPE_COLLECTION.get("worksheet").unwrap();
        let sheet_relationship_id = self
            .relations_part
            .set_new_relationship_mut(&worksheet_content, Some(format!("sheet{}", sheet_count)))
            .context("Failed to Create Sheet Relationship")?;
        self.sheet_names.push((sheet_name, sheet_relationship_id));
        Ok(WorkSheet::new(
            self.office_document.clone(),
            Rc::downgrade(&self.common_service),
            &format!("{}/worksheets/sheet{}.xml", self.relation_path, sheet_count),
        )
        .context("Worksheet Creation Failed")?)
    }

    pub(crate) fn set_active_sheet(&mut self, sheet_name: &str) {}

    pub(crate) fn rename_sheet_name(&mut self, sheet_name: &str, new_sheet_name: &str) {}
}

// ############################# im-mut Function ######################################
impl WorkbookPart {
    pub(crate) fn list_sheet_names(&self) -> Vec<String> {
        self.sheet_names
            .iter()
            .map(|(sheet_name, _)| sheet_name.to_string())
            .collect::<Vec<String>>()
    }
}
