use crate::{
    structs::{Workbook, Worksheet},
    Excel, ExcelPropertiesModel,
};
use anyhow::{Ok, Result};
use openxmloffice_global::CoreProperties;
use openxmloffice_xml::{get_all_queries, OpenXmlFile};
use rusqlite::params;

impl Excel {
    /// Create new or clone source file to start working on excel
    pub fn new(file_name: Option<String>, excel_setting: ExcelPropertiesModel) -> Self {
        let workbook;
        let xml_fs;
        //
        if let Some(file_name) = file_name {
            xml_fs = OpenXmlFile::open(&file_name, true, excel_setting.is_in_memory);
            Self::setup_database_schema(&xml_fs).expect("Initial schema setup Failed");
            Self::load_common_reference(&xml_fs);
            CoreProperties::update_core_properties(&xml_fs);
            workbook = Workbook::new(&xml_fs);
        } else {
            xml_fs = OpenXmlFile::create(excel_setting.is_in_memory);
            Self::setup_database_schema(&xml_fs).expect("Initial schema setup Failed");
            Self::initialize_common_reference(&xml_fs);
            CoreProperties::initialize_core_properties(&xml_fs);
            workbook = Workbook::new(&xml_fs);
        }
        return Self { xml_fs, workbook };
    }

    /// Add sheet to the current excel
    pub fn add_sheet(&self, sheet_name: &str) -> Worksheet {
        return Worksheet::new(&self, Some(sheet_name));
    }

    /// Save/Replace the current file into target destination
    pub fn save_as(&self, file_name: &str) {
        self.xml_fs.save(file_name);
    }

    /// Initialism table schema for Excel
    fn setup_database_schema(xml_fs: &OpenXmlFile) -> Result<()> {
        let scheme = get_all_queries!("excel.sql");
        for query in scheme {
            xml_fs
                .execute_query(&query, params![])
                .expect("Share string table failed");
        }
        Ok(())
    }

    /// For new file initialize the default reference
    fn initialize_common_reference(xml_fs: &OpenXmlFile) {
        // Share String Start
        // Style Start
    }

    /// Load existing data from excel to database
    fn load_common_reference(xml_fs: &OpenXmlFile) {
        // xml_fs.get_database_connection().execute(sql, params)
        // Ok(());
    }
}
