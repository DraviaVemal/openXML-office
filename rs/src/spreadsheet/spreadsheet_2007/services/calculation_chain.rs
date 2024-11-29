use crate::{
    files::{OfficeDocument, XmlDocument},
    get_specific_queries,
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
        if let Some(office_doc_ref) = self.office_document.upgrade() {
            if let Some(select_query) =
                get_specific_queries!("calculation_chain.sql", "select_calculation_chain_table")
            {
                fn row_mapper(row: &Row) -> AnyResult<(String, String), rusqlite::Error> {
                    Result::Ok((row.get(0)?, row.get(1)?))
                }
                let string_collection = office_doc_ref
                    .borrow()
                    .get_connection()
                    .find_many(&select_query, params![], row_mapper)
                    .unwrap();
                if let Some(xml_doc) = self.xml_document.upgrade() {
                    let mut doc = xml_doc.borrow_mut();
                    for (cell_id, sheet_id) in string_collection {
                        let mut attributes = HashMap::new();
                        attributes.insert("r".to_string(), cell_id);
                        attributes.insert("i".to_string(), sheet_id);
                        let _ = doc
                            .append_child_mut(&0, "c")
                            .unwrap()
                            .set_attribute_mut(attributes);
                    }
                }
            }
        }

        if let Some(xml_tree) = self.office_document.upgrade() {
            let _ = xml_tree
                .try_borrow_mut()
                .unwrap()
                .close_xml_document(&self.file_path);
        }
    }
}

impl XmlDocumentPart for CalculationChain {
    fn new(
        office_document: Weak<RefCell<OfficeDocument>>,
        file_path: Option<String>,
    ) -> AnyResult<Self, AnyError> {
        let file_path = file_path.unwrap_or("xl/calcChain.xml".to_string());
        let mut xml_document = Self::get_xml_document(&office_document, &file_path)?;
        Self::load_content_to_database(&office_document, &mut xml_document)
            .context("Load Calculation Chain To DB Failed")?;
        Ok(Self {
            office_document,
            xml_document,
            file_path,
        })
    }

    fn flush(self) {}

    /// Initialize xml content for this part from base template
    fn initialize_content_xml() -> AnyResult<XmlDocument, AnyError> {
        let mut attributes: HashMap<String, String> = HashMap::new();
        attributes.insert(
            "xmlns".to_string(),
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main".to_string(),
        );
        let mut xml_document = XmlDocument::new();
        xml_document
            .create_root_mut("calcChain")
            .context("Create XML Root Element Failed")?
            .set_attribute_mut(attributes)
            .context("Set Attribute Failed")?;
        Ok(xml_document)
    }
}

impl CalculationChain {
    fn load_content_to_database(
        office_document: &Weak<RefCell<OfficeDocument>>,
        xml_document: &mut Weak<RefCell<XmlDocument>>,
    ) -> AnyResult<(), AnyError> {
        if let Some(office_doc_ref) = office_document.upgrade() {
            if let Some(create_query) =
                get_specific_queries!("calculation_chain.sql", "create_calculation_chain_table")
            {
                if let Some(insert_query) =
                    get_specific_queries!("calculation_chain.sql", "insert_calculation_chain_table")
                {
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
                        if let Some(elements) = xml_doc_mut.pop_elements_by_tag_mut(&0, "c") {
                            for element in elements {
                                if let Some(attributes) = element.get_attribute() {
                                    let _ = office_doc
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
