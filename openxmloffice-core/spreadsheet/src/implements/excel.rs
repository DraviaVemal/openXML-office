use crate::{
    structs::{Workbook, Worksheet},
    Excel,
};
use anyhow::{Ok, Result};
use openxmloffice_core_global::*;
use openxmloffice_core_xml::{get_all_queries, OpenXmlFile};
use rusqlite::params;

impl Excel {
    pub fn new(file_name: Option<String>) -> Self {
        if let Some(file_name) = file_name {
            let xml_fs = OpenXmlFile::open(&file_name, true);
            Self::setup_database_schema(&xml_fs).expect("Initial schema setup Failed");
            Self::load_common_reference(&xml_fs);
            CoreProperties::update_core_properties(&xml_fs);
            let workbook = Workbook::new(&xml_fs);
            return Self { xml_fs, workbook };
        } else {
            let xml_fs = OpenXmlFile::create();
            Self::setup_database_schema(&xml_fs).expect("Initial schema setup Failed");
            Self::initialize_common_reference(&xml_fs);
            CoreProperties::initialize_core_properties(&xml_fs);
            let workbook = Workbook::new(&xml_fs);
            return Self { xml_fs, workbook };
        }
    }
    pub fn add_sheet(&self, sheet_name: &str) -> Worksheet {
        return Worksheet::new(&self, Some(sheet_name));
    }

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
