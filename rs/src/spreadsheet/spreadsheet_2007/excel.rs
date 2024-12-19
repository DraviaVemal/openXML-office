use crate::{
    files::OfficeDocument,
    global_2007::{
        parts::{CorePropertiesPart, RelationsPart},
        traits::{XmlDocumentPart, XmlDocumentPartCommon},
    },
    spreadsheet_2007::parts::{WorkSheet, WorkbookPart},
};
use anyhow::{Context, Error as AnyError, Result as AnyResult};
use std::{cell::RefCell, rc::Rc};

#[derive(Debug)]
pub struct Excel {
    office_document: Rc<RefCell<OfficeDocument>>,
    root_relations: Rc<RefCell<RelationsPart>>,
    core_properties: CorePropertiesPart,
    workbook: WorkbookPart,
}

#[derive(Debug)]
pub struct ExcelPropertiesModel {
    pub is_in_memory: bool,
    pub is_editable: bool,
}

impl ExcelPropertiesModel {
    pub fn default() -> ExcelPropertiesModel {
        ExcelPropertiesModel {
            is_in_memory: true,
            is_editable: true,
        }
    }
}

// ##################################### Feature Function ################################
impl Excel {
    /// Create new or clone source file to start working on Excel
    pub fn new(
        file_name: Option<String>,
        excel_setting: ExcelPropertiesModel,
    ) -> AnyResult<Self, AnyError> {
        let office_document = Rc::new(RefCell::new(
            OfficeDocument::new(file_name.clone(), excel_setting.is_in_memory)
                .context("Creating Office Document Struct Failed")?,
        ));
        let root_relations = Rc::new(RefCell::new(
            RelationsPart::new(Rc::downgrade(&office_document), "_rels/.rels")
                .context("Initialize Root Relation Part failed")?,
        ));
        // Load relevant parts from root relations part
        let core_properties = CorePropertiesPart::new(
            Rc::downgrade(&office_document),
            Rc::downgrade(&root_relations),
            None,
        )
        .context("Creating Core Property Part Failed.")?;
        let workbook = WorkbookPart::new(
            Rc::downgrade(&office_document),
            Rc::downgrade(&root_relations),
            None,
        )
        .context("Creating Workbook part Failed")?;
        let mut excel = Self {
            office_document,
            root_relations,
            core_properties,
            workbook,
        };
        if file_name.is_none() {
            excel
                .add_sheet(None)
                .context("Failed To Add Default Sheet to excel")?;
        }
        Ok(excel)
    }

    /// Add sheet to the current excel
    pub fn add_sheet(&mut self, sheet_name: Option<String>) -> AnyResult<WorkSheet, AnyError> {
        self.get_workbook_mut().add_sheet(sheet_name)
    }

    /// Add sheet to the current excel
    pub fn rename_sheet_name(
        &mut self,
        old_sheet_name: String,
        new_sheet_name: String,
    ) -> AnyResult<(), AnyError> {
        self.get_workbook_mut()
            .rename_sheet_name(&old_sheet_name, &new_sheet_name)
    }

    /// Get Worksheet handle by sheet name
    pub fn get_worksheet(&mut self, sheet_name: String) -> AnyResult<WorkSheet, AnyError> {
        self.get_workbook_mut().get_worksheet(&sheet_name)
    }

    /// Save/Replace the current file into target destination
    pub fn save_as(self, file_name: &str) -> AnyResult<(), AnyError> {
        self.workbook.flush()?;
        self.core_properties.flush()?;
        self.root_relations
            .try_borrow_mut()
            .context("Failed To Pull Relation Handle")?
            .close_document()?;
        self.office_document
            .try_borrow_mut()
            .context("Save Office Document handle Failed")?
            .save_as(file_name)
            .context("File Save Failed for the target path.")?;
        Ok(())
    }
}

// ############################# Internal Function ######################################
// ############################# Mut Function      ######################################
impl Excel {
    fn get_workbook_mut(&mut self) -> &mut WorkbookPart {
        &mut self.workbook
    }
}

// ############################# Im-Mut Function   ######################################
impl Excel {
    fn get_workbook(&self) -> &WorkbookPart {
        &self.workbook
    }
    pub fn list_sheet_names(&self) -> AnyResult<Vec<String>, AnyError> {
        self.get_workbook().list_sheet_names()
    }
}
