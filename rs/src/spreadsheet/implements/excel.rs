use super::{worksheet::Worksheet, Workbook};
use crate::{
    core_properties::CorePropertiesPart, get_all_queries, relations::RelationsPart,
    theme::ThemePart, OpenXmlFile, XmlElement,
};
use anyhow::{Context, Error as AnyError, Ok, Result as AnyResult};
use rusqlite::params;
use std::{cell::RefCell, rc::Rc};

#[derive(Debug)]
pub struct Excel {
    pub(crate) xml_fs: Rc<RefCell<OpenXmlFile>>,
    pub(crate) workbook: Workbook,
}

#[derive(Debug)]
pub struct ExcelPropertiesModel {
    pub is_in_memory: bool,
}

impl Excel {
    /// Default Excel Setting
    pub fn default() -> ExcelPropertiesModel {
        return ExcelPropertiesModel { is_in_memory: true };
    }
    /// Create new or clone source file to start working on excel
    pub fn new(
        file_name: Option<String>,
        excel_setting: ExcelPropertiesModel,
    ) -> AnyResult<Self, AnyError> {
        let xml_fs;
        //
        if let Some(file_name) = file_name {
            let open_xml_file = OpenXmlFile::open(&file_name, true, excel_setting.is_in_memory)
                .context("Open Existing File Failed")?;
            xml_fs = Rc::new(RefCell::new(open_xml_file));
            Self::setup_database_schema(&xml_fs)?;
            Self::load_common_reference(&xml_fs);
            CorePropertiesPart::new(&xml_fs, None)?;
        } else {
            let open_xml_file = OpenXmlFile::create(excel_setting.is_in_memory)
                .context("Create New File Failed")?;
            xml_fs = Rc::new(RefCell::new(open_xml_file));
            Self::setup_database_schema(&xml_fs)?;
            Self::initialize_common_reference(&xml_fs);
            RelationsPart::new(&xml_fs, None)?;
            CorePropertiesPart::new(&xml_fs, None)?;
            ThemePart::new(&xml_fs, Some("xl/theme/theme1.xml"))?;
        }
        let workbook = Workbook::new(&xml_fs, None).context("Workbook Creation Failed")?;
        return Ok(Self { xml_fs, workbook });
    }

    /// Add sheet to the current excel
    pub fn add_sheet(&self, sheet_name: &str) -> AnyResult<Worksheet, AnyError> {
        let worksheet = Worksheet::new(&self.xml_fs, Some(sheet_name));
        worksheet
    }

    /// Save/Replace the current file into target destination
    pub fn save_as(self, file_name: &str) -> AnyResult<(), AnyError> {
        self.workbook.flush();
        self.xml_fs
            .borrow()
            .save(file_name)
            .context("File Save Failed for the target path.")?;
        Ok(())
    }

    /// Initialism table schema for Excel
    fn setup_database_schema(xml_fs: &Rc<RefCell<OpenXmlFile>>) -> AnyResult<(), AnyError> {
        let scheme = get_all_queries!("excel.sql");
        for query in scheme {
            xml_fs
                .borrow()
                .execute_query(&query, params![])
                .context(format!("Failed Executing Query : {}", &query))?;
        }
        Ok(())
    }

    /// For new file initialize the default reference
    fn initialize_common_reference(xml_fs: &Rc<RefCell<OpenXmlFile>>) {
        // Share String Start
        // Style Start
    }

    /// Load existing data from excel to database
    fn load_common_reference(xml_fs: &Rc<RefCell<OpenXmlFile>>) {
        // xml_fs.get_database_connection().execute(sql, params)
        // Ok(());
    }
}
