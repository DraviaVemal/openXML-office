use crate::global_2007::traits::XmlDocumentPartCommon;
use crate::reference_dictionary::EXCEL_TYPE_COLLECTION;
use crate::{
    files::{OfficeDocument, XmlDocument},
    get_all_queries,
    global_2007::traits::XmlDocumentPart,
};
use anyhow::{anyhow, Context, Error as AnyError, Result as AnyResult};
use rusqlite::{params, Row};
use std::{cell::RefCell, collections::HashMap, rc::Weak};

#[derive(Debug)]
pub struct CalculationChain {
    office_document: Weak<RefCell<OfficeDocument>>,
    xml_document: Weak<RefCell<XmlDocument>>,
    file_path: String,
}

impl Drop for CalculationChain {
    fn drop(&mut self) {
        let _ = self.close_document();
    }
}

impl XmlDocumentPartCommon for CalculationChain {
    /// Initialize xml content for this part from base template
    fn initialize_content_xml() -> AnyResult<(XmlDocument, Option<String>), AnyError> {
        let mut attributes: HashMap<String, String> = HashMap::new();
        attributes.insert(
            "xmlns".to_string(),
            EXCEL_TYPE_COLLECTION
                .get("calc_chain")
                .unwrap()
                .schemas_namespace
                .to_string(),
        );
        let mut xml_document = XmlDocument::new();
        xml_document
            .create_root_mut("calcChain")
            .context("Create XML Root Element Failed")?
            .set_attribute_mut(attributes)
            .context("Set Attribute Failed")?;
        Ok((
            xml_document,
            Some(
                EXCEL_TYPE_COLLECTION
                    .get("calc_chain")
                    .unwrap()
                    .content_type
                    .to_string(),
            ),
        ))
    }
    fn close_document(&mut self) -> AnyResult<(), AnyError>
    where
        Self: Sized,
    {
        if let Some(office_doc_ref) = self.office_document.upgrade() {
            let select_query = get_all_queries!("calculation_chain.sql").get("select_calculation_chain_table").unwrap().to_owned();
            fn row_mapper(row: &Row) -> AnyResult<(String, String), rusqlite::Error> {
                Result::Ok((row.get(0)?, row.get(1)?))
            }
            let string_collection = office_doc_ref
                .try_borrow()
                .context("Failed to get office handle")?
                .get_connection()
                .find_many(&select_query, params![], row_mapper)
                .context("Failed to Pull All Calculation Chain Items")?;
            if string_collection.len() > 0 {
                if let Some(xml_doc) = self.xml_document.upgrade() {
                    let mut doc = xml_doc.borrow_mut();
                    for (cell_id, sheet_id) in string_collection {
                        let mut attributes = HashMap::new();
                        attributes.insert("r".to_string(), cell_id);
                        attributes.insert("i".to_string(), sheet_id);
                        doc.append_child_mut("c", None)
                            .context("Failed To Add Child Item")?
                            .set_attribute_mut(attributes)?;
                    }
                }
                office_doc_ref
                    .try_borrow_mut()
                    .context("Failed to Borrow Share Tree")?
                    .close_xml_document(&self.file_path)?;
            } else {
                office_doc_ref
                    .try_borrow_mut()
                    .context("Failed to Borrow Share Tree")?
                    .delete_document_mut(&self.file_path)?;
            }
        }
        Ok(())
    }
}

impl XmlDocumentPart for CalculationChain {
    fn new(
        office_document: Weak<RefCell<OfficeDocument>>,
        file_path: &str,
    ) -> AnyResult<Self, AnyError> {
        let mut xml_document = Self::get_xml_document(&office_document, &file_path)?;
        Self::load_content_to_database(&office_document, &mut xml_document)
            .context("Load Calculation Chain To DB Failed")?;
        Ok(Self {
            office_document,
            xml_document,
            file_path: file_path.to_string(),
        })
    }
}

impl CalculationChain {
    fn load_content_to_database(
        office_document: &Weak<RefCell<OfficeDocument>>,
        xml_document: &mut Weak<RefCell<XmlDocument>>,
    ) -> AnyResult<(), AnyError> {
        if let Some(office_doc_ref) = office_document.upgrade() {
            let queries = get_all_queries!("calculation_chain.sql");
            if let Some(create_query) = queries.get("create_calculation_chain_table") {
                if let Some(insert_query) = queries.get("insert_calculation_chain_table") {
                    let office_doc = office_doc_ref
                        .try_borrow()
                        .context("Pulling Office Doc Failed")?;
                    office_doc
                        .get_connection()
                        .create_table(&create_query)
                        .context("Create Share String Table Failed")?;
                    if let Some(xml_doc) = xml_document.upgrade() {
                        let mut xml_doc_mut =
                            xml_doc.try_borrow_mut().context("xml doc borrow failed")?;
                        if let Some(elements) = xml_doc_mut.pop_elements_by_tag_mut("c", None) {
                            for element in elements {
                                if let Some(attributes) = element.get_attribute() {
                                    office_doc
                                        .get_connection()
                                        .insert_record(
                                            &insert_query,
                                            params![attributes["r"], attributes["i"]],
                                        )
                                        .context("Create Share String Table Failed")?;
                                }
                            }
                        }
                    }
                } else {
                    return Err(anyhow!("Insert Query Failed"));
                }
            } else {
                return Err(anyhow!("Create Table Failed"));
            }
        }
        Ok(())
    }
}
