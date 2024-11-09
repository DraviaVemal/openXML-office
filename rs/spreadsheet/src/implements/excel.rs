use crate::{
    structs::{Workbook, Worksheet},
    Excel, ExcelPropertiesModel,
};
use anyhow::{Ok, Result};
use openxmloffice_global::{xml_file::XmlElement, CorePropertiesPart, RelationsPart, ThemePart};
use openxmloffice_xml::{get_all_queries, OpenXmlFile};
use rusqlite::params;
use std::{cell::RefCell, rc::Rc};

impl Excel {
    /// Default Excel Setting
    pub fn default() -> ExcelPropertiesModel {
        return ExcelPropertiesModel { is_in_memory: true };
    }
    /// Create new or clone source file to start working on excel
    pub fn new(file_name: Option<String>, excel_setting: ExcelPropertiesModel) -> Self {
        let workbook;
        let xml_fs;
        //
        if let Some(file_name) = file_name {
            let open_xml_file = OpenXmlFile::open(&file_name, true, excel_setting.is_in_memory);
            xml_fs = Rc::new(RefCell::new(open_xml_file));
            Self::setup_database_schema(&xml_fs).expect("Initial schema setup Failed");
            Self::load_common_reference(&xml_fs);
            workbook = Workbook::new(&xml_fs, None);
        } else {
            let open_xml_file = OpenXmlFile::create(excel_setting.is_in_memory);
            xml_fs = Rc::new(RefCell::new(open_xml_file));
            Self::setup_database_schema(&xml_fs).expect("Initial schema setup Failed");
            Self::initialize_common_reference(&xml_fs);
            RelationsPart::new(&xml_fs, None);
            CorePropertiesPart::new(&xml_fs, None);
            ThemePart::new(&xml_fs, Some("xl/theme/theme1.xml"));
            workbook = Workbook::new(&xml_fs, None);
        }
        return Self { xml_fs, workbook };
    }

    /// Add sheet to the current excel
    pub fn add_sheet(&self, sheet_name: &str) -> Worksheet {
        let worksheet = Worksheet::new(&self.xml_fs, Some(sheet_name));
        return worksheet;
    }

    /// Save/Replace the current file into target destination
    pub fn save_as(self, file_name: &str) {
        self.workbook.flush();
        self.xml_fs.borrow().save(file_name);
    }

    /// Initialism table schema for Excel
    fn setup_database_schema(xml_fs: &Rc<RefCell<OpenXmlFile>>) -> Result<()> {
        let scheme = get_all_queries!("excel.sql");
        for query in scheme {
            xml_fs
                .borrow()
                .execute_query(&query, params![])
                .expect("Share string table failed");
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
