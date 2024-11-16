use crate::{
    get_all_queries,
    global_2007::{
        parts::{CorePropertiesPart, RelationsPart, ThemePart},
        traits::XmlElement,
    },
    files::OpenXmlFile,
};
use anyhow::{Error as AnyError, Ok, Result as AnyResult};
use rusqlite::params;
use std::{cell::RefCell, rc::Rc};

#[derive(Debug)]
pub struct PowerPoint {
    pub(crate) xml_fs: Rc<RefCell<OpenXmlFile>>,
}

#[derive(Debug)]
pub struct PowerPointPropertiesModel {
    pub is_in_memory: bool,
}

impl PowerPoint {
    /// Default Power Point Setting
    pub fn default() -> PowerPointPropertiesModel {
        return PowerPointPropertiesModel { is_in_memory: true };
    }
    /// Create new or clone source file to start working on Power Point
    pub fn new(
        file_name: Option<String>,
        power_point_setting: PowerPointPropertiesModel,
    ) -> AnyResult<Self, AnyError> {
        let xml_fs;
        //
        if let Some(file_name) = file_name {
            let open_xml_file =
                OpenXmlFile::open(&file_name, true, power_point_setting.is_in_memory)?;
            xml_fs = Rc::new(RefCell::new(open_xml_file));
            Self::setup_database_schema(&xml_fs)?;
            Self::load_common_reference(&xml_fs);
            CorePropertiesPart::new(&xml_fs, None)?;
        } else {
            let open_xml_file = OpenXmlFile::create(power_point_setting.is_in_memory)?;
            xml_fs = Rc::new(RefCell::new(open_xml_file));
            Self::setup_database_schema(&xml_fs)?;
            Self::initialize_common_reference(&xml_fs);
            RelationsPart::new(&xml_fs, None)?;
            CorePropertiesPart::new(&xml_fs, None)?;
            ThemePart::new(&xml_fs, Some("ppt/theme/theme1.xml"))?;
        }
        return Ok(Self { xml_fs });
    }

    /// Save/Replace the current file into target destination
    pub fn save_as(self, file_name: &str) -> Result<(), anyhow::Error> {
        self.xml_fs.borrow().save(file_name)
    }

    /// Initialism table schema for PowerPoint
    fn setup_database_schema(xml_fs: &Rc<RefCell<OpenXmlFile>>) -> AnyResult<(), AnyError> {
        let scheme = get_all_queries!("power_point.sql");
        for query in scheme {
            xml_fs.borrow().execute_query(&query, params![])?;
        }
        Ok(())
    }

    /// For new file initialize the default reference
    fn initialize_common_reference(xml_fs: &Rc<RefCell<OpenXmlFile>>) {
        // Share String Start
        // Style Start
    }

    /// Load existing data from power point to database
    fn load_common_reference(xml_fs: &Rc<RefCell<OpenXmlFile>>) {
        // xml_fs.get_database_connection().execute(sql, params)
        // Ok(());
    }
}
