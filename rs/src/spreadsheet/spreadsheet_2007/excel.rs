use crate::global_2007::traits::XmlDocumentPartCommon;
use crate::reference_dictionary::EXCEL_TYPE_COLLECTION;
use crate::{
    files::OfficeDocument,
    global_2007::{
        parts::{ContentTypesPart, CorePropertiesPart, RelationsPart},
        traits::XmlDocumentPart,
    },
    spreadsheet_2007::parts::{WorkSheet, WorkbookPart},
};
use anyhow::{Context, Error as AnyError, Ok, Result as AnyResult};
use std::rc::Weak;
use std::{cell::RefCell, rc::Rc};

#[derive(Debug)]
pub struct Excel {
    office_document: Rc<RefCell<OfficeDocument>>,
    root_relations: RelationsPart,
    content_type: ContentTypesPart,
    core_properties: CorePropertiesPart,
    workbook: WorkbookPart,
}

#[derive(Debug)]
pub struct ExcelPropertiesModel {
    pub is_in_memory: bool,
}

// ##################################### Feature Function ################################
impl Excel {
    /// Default Excel Setting
    pub fn default() -> ExcelPropertiesModel {
        ExcelPropertiesModel { is_in_memory: true }
    }
    /// Create new or clone source file to start working on Excel
    pub fn new(
        file_name: Option<String>,
        excel_setting: ExcelPropertiesModel,
    ) -> AnyResult<Self, AnyError> {
        let office_document: OfficeDocument =
            OfficeDocument::new(file_name.clone(), excel_setting.is_in_memory)
                .context("Creating Office Document Struct Failed")?;
        let rc_office_document: Rc<RefCell<OfficeDocument>> =
            Rc::new(RefCell::new(office_document));
        let mut root_relations =
            RelationsPart::new(Rc::downgrade(&rc_office_document), "_rels/.rels")
                .context("Initialize Root Relation Part failed")?;
        let content_type =
            ContentTypesPart::new(Rc::downgrade(&rc_office_document), "[Content_Types].xml")
                .context("Initializing Content Type Part Failed")?;
        // Load relevant parts from root relations part
        let core_properties = CorePropertiesPart::create_core_properties(
            &mut root_relations,
            Rc::downgrade(&rc_office_document),
        )
        .context("Creating Core Property Part Failed.")?;
        let workbook =
            Excel::create_workbook_part(&mut root_relations, Rc::downgrade(&rc_office_document))
                .context("Creating Workbook part Failed")?;
        let mut excel = Self {
            office_document: rc_office_document,
            root_relations,
            content_type,
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

    /// Save/Replace the current file into target destination
    pub fn save_as(self, file_name: &str) -> AnyResult<(), AnyError> {
        self.workbook.flush()?;
        self.content_type.flush()?;
        self.core_properties.flush()?;
        self.root_relations.flush()?;
        self.office_document
            .try_borrow_mut()
            .context("Save Office Document handle Failed")?
            .save_as(file_name)
            .context("File Save Failed for the target path.")?;
        Ok(())
    }
}

// ############################# Internal Function ######################################
impl Excel {
    fn create_workbook_part(
        relations_part: &mut RelationsPart,
        office_document: Weak<RefCell<OfficeDocument>>,
    ) -> AnyResult<WorkbookPart, AnyError> {
        let workbook_content = EXCEL_TYPE_COLLECTION.get("workbook").unwrap();
        let workbook_path: Option<String> = relations_part
            .get_relationship_target_by_type(&workbook_content.schemas_type)
            .context("Parsing workbook path failed")?;
        Ok(if let Some(part_path) = workbook_path {
            WorkbookPart::new(office_document, &part_path).context("Workbook Creation Failed")?
        } else {
            relations_part
                .set_new_relationship_mut(workbook_content, None)
                .context("Setting New Theme Relationship Failed.")?;
            WorkbookPart::new(office_document, "xl/workbook.xml")
                .context("Workbook Creation Failed")?
        })
    }
}
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
    pub fn list_sheet_names(&self) -> Vec<String> {
        self.get_workbook().list_sheet_names()
    }
}
