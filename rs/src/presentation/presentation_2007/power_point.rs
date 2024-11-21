use crate::{
    files::OfficeDocument,
    get_all_queries,
    global_2007::{
        parts::{CorePropertiesPart, RelationsPart, ThemePart},
        traits::XmlDocumentPart,
    },
};
use anyhow::{Context, Error as AnyError, Ok, Result as AnyResult};
use std::{cell::RefCell, rc::Rc};

#[derive(Debug)]
pub struct PowerPoint {
    pub(crate) office_document: Rc<RefCell<OfficeDocument>>,
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
        let is_file_exist = file_name.is_some();
        let office_document = OfficeDocument::new(file_name, power_point_setting.is_in_memory)
            .context("Creating Office Document Struct Failed")?;
        let rc_office_document: Rc<RefCell<OfficeDocument>> =
            Rc::new(RefCell::new(office_document));
        Self::setup_database_schema(&rc_office_document).context("Word Schema Setup Failed")?;
        if is_file_exist {
            CorePropertiesPart::new(&rc_office_document, None)
                .context("Load CorePart for Existing file failed")?;
        } else {
            RelationsPart::new(&rc_office_document, None)
                .context("Initialize Relation Part failed")?;
            CorePropertiesPart::new(&rc_office_document, None)
                .context("Create CorePart for new file failed")?;
            ThemePart::new(
                &rc_office_document,
                Some("ppt/theme/theme1.xml".to_string()),
            )
            .context("Initializing new theme part failed")?;
        }
        Ok(Self {
            office_document: rc_office_document,
        })
    }

    /// Save/Replace the current file into target destination
    pub fn save_as(self, file_name: &str) -> AnyResult<(), AnyError> {
        self.office_document
            .try_borrow_mut()
            .context("Save Office Document handle Failed")?
            .save_as(file_name)
            .context("File Save Failed for the target path.")
    }

    /// Initialism table schema for PowerPoint
    fn setup_database_schema(xml_fs: &Rc<RefCell<OfficeDocument>>) -> AnyResult<(), AnyError> {
        let scheme = get_all_queries!("power_point.sql");
        for query in scheme {
            xml_fs
                .borrow()
                .get_connection()
                .create_table(&query)
                .context("Power Point Schema Initialization Failed")?;
        }
        Ok(())
    }
}
