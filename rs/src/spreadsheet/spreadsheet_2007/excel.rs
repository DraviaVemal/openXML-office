use crate::global_2007::traits::XmlDocumentPartCommon;
use crate::{
    files::OfficeDocument,
    global_2007::{
        parts::{ContentTypesPart, CorePropertiesPart, RelationsPart},
        traits::XmlDocumentPart,
    },
    spreadsheet_2007::parts::{WorkSheet, WorkbookPart},
};
use anyhow::{Context, Error as AnyError, Ok, Result as AnyResult};
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
            OfficeDocument::new(file_name, excel_setting.is_in_memory)
                .context("Creating Office Document Struct Failed")?;
        let rc_office_document: Rc<RefCell<OfficeDocument>> =
            Rc::new(RefCell::new(office_document));
        let root_relations = RelationsPart::new(Rc::downgrade(&rc_office_document), None)
            .context("Initialize Root Relation Part failed")?;
        let content_type = ContentTypesPart::new(Rc::downgrade(&rc_office_document), None)
            .context("Initializing Content Type Part Failed")?;
        // Load relevant parts from root relations part
        let core_properties_path: Option<String> = root_relations.get_relationship_target_by_type("http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties").context("Parsing Core Properties path failed")?;
        let workbook_path: Option<String> = root_relations.get_relationship_target_by_type("http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument").context("Parsing workbook path failed")?;
        let core_properties =
            CorePropertiesPart::new(Rc::downgrade(&rc_office_document), core_properties_path)
                .context("Load CorePart for Existing file failed")?;
        let workbook: WorkbookPart =
            WorkbookPart::new(Rc::downgrade(&rc_office_document), workbook_path)
                .context("Workbook Creation Failed")?;
        Ok(Self {
            office_document: rc_office_document,
            root_relations,
            content_type,
            core_properties,
            workbook,
        })
    }

    /// Add sheet to the current excel
    pub fn add_sheet(&mut self, sheet_name: Option<String>) -> AnyResult<WorkSheet, AnyError> {
        self.get_workbook().add_sheet(sheet_name)
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
    fn get_workbook(&mut self) -> &mut WorkbookPart {
        &mut self.workbook
    }
}
