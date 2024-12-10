use crate::global_2007::traits::{XmlDocumentPartCommon, XmlDocumentServicePart};
use crate::{
    files::{OfficeDocument, XmlDocument, XmlSerializer},
    global_2007::{
        parts::{RelationsPart, ThemePart},
        traits::XmlDocumentPart,
    },
    spreadsheet_2007::{
        parts::WorkSheet,
        services::{CalculationChain, CommonServices, ShareString, Style},
    },
};
use anyhow::{Context, Error as AnyError, Ok, Result as AnyResult};
use std::{
    cell::RefCell,
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
}

impl Drop for WorkbookPart {
    fn drop(&mut self) {
        let _ = self.close_document();
    }
}

impl XmlDocumentPartCommon for WorkbookPart {
    /// Initialize xml content for this part from base template
    fn initialize_content_xml() -> AnyResult<XmlDocument, AnyError> {
        let template_core_properties = r#"
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <fileVersion appName="openxml-office" lastEdited="7" lowestEdited="7"/>
        </workbook>"#;
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
        let file_tree = Self::get_xml_document(&office_document, file_path)?;
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
        let theme_part = WorkbookPart::create_theme_part(
            &mut relations_part,
            office_document.clone(),
            &relation_path,
        )
        .context("Loading Theme Part Failed")?;
        let share_string = WorkbookPart::create_share_string(
            &mut relations_part,
            office_document.clone(),
            &relation_path,
        )
        .context("Loading Share String Failed")?;
        let calculation_chain = WorkbookPart::create_calculation_chain(
            &mut relations_part,
            office_document.clone(),
            &relation_path,
        )
        .context("Loading Calculation Chain Failed")?;
        let style = WorkbookPart::create_style_part(
            &mut relations_part,
            office_document.clone(),
            &relation_path,
        )
        .context("Loading Style Part Failed")?;
        let common_service = Rc::new(RefCell::new(CommonServices::new(
            calculation_chain,
            share_string,
            style,
        )));
        Ok(Self {
            office_document,
            xml_document: file_tree,
            file_path: file_path.to_string(),
            common_service,
            relations_part,
            theme_part,
        })
    }
}

// ############################# Internal Function ######################################
impl WorkbookPart {
    fn create_theme_part(
        relations_part: &mut RelationsPart,
        office_document: Weak<RefCell<OfficeDocument>>,
        relation_path: &str,
    ) -> AnyResult<ThemePart, AnyError> {
        let theme_part_path = relations_part
            .get_relationship_target_by_type(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
            )
            .context("Parsing Theme part path failed")?;
        Ok(if let Some(part_path) = theme_part_path {
            ThemePart::new(office_document.clone(), &part_path)
                .context("Creating Theme part for workbook failed")?
        } else {
            relations_part
                .set_new_relationship_mut(
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
                    &format!("{}/theme/theme1.xml", relation_path),
                )
                .context("Setting New Theme Relationship Failed.")?;
            ThemePart::new(
                office_document.clone(),
                &format!("{}/theme/theme1.xml", relation_path),
            )
            .context("Creating Theme part for workbook failed")?
        })
    }
    fn create_share_string(
        relations_part: &mut RelationsPart,
        office_document: Weak<RefCell<OfficeDocument>>,
        relation_path: &str,
    ) -> AnyResult<ShareString, AnyError> {
        let share_string_path = relations_part
            .get_relationship_target_by_type(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
            )
            .context("Parsing Theme part path failed")?;
        Ok(if let Some(part_path) = share_string_path {
            ShareString::new(office_document.clone(), &part_path)
                .context("Share String Service Object Creation Failure")?
        } else {
            relations_part
            .set_new_relationship_mut(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
                &format!("{}/sharedStrings.xml", relation_path),
            )
            .context("Setting New Theme Relationship Failed.")?;
            ShareString::new(
                office_document.clone(),
                &format!("{}/sharedStrings.xml", relation_path),
            )
            .context("Share String Service Object Creation Failure")?
        })
    }
    fn create_calculation_chain(
        relations_part: &mut RelationsPart,
        office_document: Weak<RefCell<OfficeDocument>>,
        relation_path: &str,
    ) -> AnyResult<CalculationChain, AnyError> {
        let calculation_chain_path = relations_part
            .get_relationship_target_by_type(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain",
            )
            .context("Parsing Theme part path failed")?;
        Ok(if let Some(part_path) = calculation_chain_path {
            CalculationChain::new(office_document.clone(), &part_path)
                .context("Calculation Chain Service Object Creation Failure")?
        } else {
            relations_part
                .set_new_relationship_mut(
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain",
                    &format!("{}/calcChain.xml", relation_path),
                )
                .context("Setting New Theme Relationship Failed.")?;
            CalculationChain::new(
                office_document.clone(),
                &format!("{}/calcChain.xml", relation_path),
            )
            .context("Calculation Chain Service Object Creation Failure")?
        })
    }

    fn create_style_part(
        relations_part: &mut RelationsPart,
        office_document: Weak<RefCell<OfficeDocument>>,
        relation_path: &str,
    ) -> AnyResult<Style, AnyError> {
        let style_path = relations_part
            .get_relationship_target_by_type(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
            )
            .context("Parsing Theme part path failed")?;
        Ok(if let Some(part_path) = style_path {
            Style::new(office_document.clone(), &part_path)
                .context("Style Service Object Creation Failure")?
        } else {
            relations_part
                .set_new_relationship_mut(
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
                    &format!("{}/styles.xml", relation_path),
                )
                .context("Setting New Theme Relationship Failed.")?;
            Style::new(
                office_document.clone(),
                &format!("{}/styles.xml", relation_path),
            )
            .context("Style Service Object Creation Failure")?
        })
    }
}

// ############################# Feature Function ######################################
impl WorkbookPart {
    pub(crate) fn add_sheet(
        &mut self,
        file_name: Option<String>,
    ) -> AnyResult<WorkSheet, AnyError> {
        Ok(WorkSheet::new(
            self.office_document.clone(),
            Rc::downgrade(&self.common_service),
            file_name,
        )
        .context("Worksheet Creation Failed")?)
    }

    pub(crate) fn set_active_sheet(&mut self, sheet_name: &str) {}

    pub(crate) fn get_sheet_name(&self, sheet_name: &str) {}

    pub(crate) fn rename_sheet_name(&self, sheet_name: &str, new_sheet_name: &str) {}
}
