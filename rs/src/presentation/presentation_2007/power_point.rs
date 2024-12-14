use crate::{
    files::OfficeDocument,
    get_all_queries,
    global_2007::{
        parts::{CorePropertiesPart, RelationsPart},
        traits::{XmlDocumentPart, XmlDocumentPartCommon},
    },
};
use anyhow::{Context, Error as AnyError, Ok, Result as AnyResult};
use std::{cell::RefCell, rc::Rc};

#[derive(Debug)]
pub struct PowerPoint {
    office_document: Rc<RefCell<OfficeDocument>>,
    root_relations: RelationsPart,
    core_properties: CorePropertiesPart,
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
        let office_document = OfficeDocument::new(file_name, power_point_setting.is_in_memory)
            .context("Creating Office Document Struct Failed")?;
        let rc_office_document: Rc<RefCell<OfficeDocument>> =
            Rc::new(RefCell::new(office_document));
        Self::setup_database_schema(&rc_office_document).context("Word Schema Setup Failed")?;
        let mut root_relations =
            RelationsPart::new(Rc::downgrade(&rc_office_document), "_rels/.rels")
                .context("Initialize Root Relation Part failed")?;
        let core_properties = CorePropertiesPart::create_core_properties(
            &mut root_relations,
            Rc::downgrade(&rc_office_document),
        )
        .context("Creating Core Property Part Failed.")?;
        Ok(Self {
            office_document: rc_office_document,
            root_relations,
            core_properties,
        })
    }

    /// Save/Replace the current file into target destination
    pub fn save_as(self, file_name: &str) -> AnyResult<(), AnyError> {
        self.core_properties.flush()?;
        self.root_relations.flush()?;
        self.office_document
            .try_borrow_mut()
            .context("Save Office Document handle Failed")?
            .save_as(file_name)
            .context("File Save Failed for the target path.")
    }

    /// Initialism table schema for PowerPoint
    fn setup_database_schema(xml_fs: &Rc<RefCell<OfficeDocument>>) -> AnyResult<(), AnyError> {
        for query in get_all_queries!("power_point.sql").values() {
            xml_fs
                .borrow()
                .get_connection()
                .create_table(&query)
                .context("Power Point Schema Initialization Failed")?;
        }
        Ok(())
    }
}
