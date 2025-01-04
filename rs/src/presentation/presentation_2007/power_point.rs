use crate::{
    files::OfficeDocument,
    global_2007::{
        parts::{CorePropertiesPart, RelationsPart},
        traits::{XmlDocumentPart, XmlDocumentPartCommon},
    },
};
use anyhow::{Context, Error as AnyError, Result as AnyResult};
use std::{cell::RefCell, rc::Rc};

#[derive(Debug)]
pub struct PowerPoint {
    office_document: Rc<RefCell<OfficeDocument>>,
    root_relations: Rc<RefCell<RelationsPart>>,
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
        let office_document = Rc::new(RefCell::new(
            OfficeDocument::new(file_name.clone(), power_point_setting.is_in_memory)
                .context("Creating Office Document Struct Failed")?,
        ));
        let root_relations = Rc::new(RefCell::new(
            RelationsPart::new(Rc::downgrade(&office_document), "_rels/.rels")
                .context("Initialize Root Relation Part failed")?,
        ));
        let core_properties = CorePropertiesPart::new(
            Rc::downgrade(&office_document),
            Rc::downgrade(&root_relations),
        )
        .context("Creating Core Property Part Failed.")?;
        Ok(Self {
            office_document,
            root_relations,
            core_properties,
        })
    }

    /// Save/Replace the current file into target destination
    pub fn save_as(self, file_name: &str) -> AnyResult<(), AnyError> {
        self.core_properties.flush()?;
        self.root_relations
            .try_borrow_mut()
            .context("Failed To Pull Relation Handle")?
            .close_document()?;
        self.office_document
            .try_borrow_mut()
            .context("Save Office Document handle Failed")?
            .save_as(file_name)
            .context("File Save Failed for the target path.")
    }
}
