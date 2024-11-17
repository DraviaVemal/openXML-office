use crate::{
    files::OfficeDocument,
    get_all_queries,
    global_2007::{
        parts::{CorePropertiesPart, RelationsPart, ThemePart},
        traits::XmlDocument,
    },
    spreadsheet_2007::parts::{WorkSheetPart, WorkbookPart},
};
use anyhow::{Context, Error as AnyError, Ok, Result as AnyResult};
use std::{cell::RefCell, rc::Rc};

#[derive(Debug)]
pub struct Excel {
    pub(crate) office_document: Rc<RefCell<OfficeDocument>>,
    workbook: WorkbookPart,
}

#[derive(Debug)]
pub struct ExcelPropertiesModel {
    pub is_in_memory: bool,
}

impl Excel {
    /// Default Excel Setting
    pub fn default() -> ExcelPropertiesModel {
        ExcelPropertiesModel { is_in_memory: true }
    }
    /// Create new or clone source file to start working on excel
    pub fn new(
        file_name: Option<String>,
        excel_setting: ExcelPropertiesModel,
    ) -> AnyResult<Self, AnyError> {
        let is_file_exist = file_name.is_some();
        let office_document = OfficeDocument::new(file_name, excel_setting.is_in_memory)
            .context("Creating Office Document Struct Failed")?;
        let rc_office_document: Rc<RefCell<OfficeDocument>> =
            Rc::new(RefCell::new(office_document));
        Self::setup_database_schema(&rc_office_document).context("Excel Schema Setup Failed")?;
        if is_file_exist {
            CorePropertiesPart::new(&rc_office_document, None)
                .context("Load CorePart for Existing file failed")?;
        } else {
            RelationsPart::new(&rc_office_document, None)
                .context("Initialize Relation Part failed")?;
            CorePropertiesPart::new(&rc_office_document, None)
                .context("Create CorePart for new file failed");
            ThemePart::new(&rc_office_document, Some("doc/theme/theme1.xml"))
                .context("Initializing new theme part failed");
        }
        let workbook =
            WorkbookPart::new(&rc_office_document, None).context("Workbook Creation Failed")?;
        Ok(Self {
            office_document: rc_office_document,
            workbook,
        })
    }

    /// Add sheet to the current excel
    pub fn add_sheet(&self, sheet_name: &str) -> AnyResult<WorkSheetPart, AnyError> {
        let worksheet = WorkSheetPart::new(&self.office_document, Some(sheet_name));
        worksheet
    }

    /// Save/Replace the current file into target destination
    pub fn save_as(self, file_name: &str) -> AnyResult<(), AnyError> {
        self.workbook.flush();
        self.office_document
            .borrow()
            .save_as(file_name)
            .context("File Save Failed for the target path.")?;
        Ok(())
    }

    /// Initialism table schema for Excel
    fn setup_database_schema(xml_fs: &Rc<RefCell<OfficeDocument>>) -> AnyResult<(), AnyError> {
        let scheme = get_all_queries!("excel.sql");
        for query in scheme {
            xml_fs
                .borrow()
                .get_connection()
                .create_table(&query)
                .context("Excel Schema Initialization Failed")?;
        }
        Ok(())
    }
}
