use crate::{
    files::{OfficeDocument, XmlDocument},
    get_specific_queries,
    global_2007::traits::XmlDocumentPart,
};
use anyhow::{anyhow, Context, Error as AnyError, Result as AnyResult};
use rusqlite::{params, Row};
use std::{
    cell::RefCell,
    collections::{HashMap, HashSet},
    rc::Weak,
};

#[derive(Debug)]
pub struct ShareString {
    office_document: Weak<RefCell<OfficeDocument>>,
    xml_document: Weak<RefCell<XmlDocument>>,
    file_path: String,
}

impl Drop for ShareString {
    fn drop(&mut self) {
        if let Some(office_doc_ref) = self.office_document.upgrade() {
            if let Some(select_query) =
                get_specific_queries!("share_string.sql", "select_share_string_table")
            {
                fn row_mapper(row: &Row) -> AnyResult<String, rusqlite::Error> {
                    Result::Ok(row.get(0)?)
                }
                let string_collection = office_doc_ref
                    .borrow()
                    .get_connection()
                    .find_many(&select_query, params![], row_mapper)
                    .unwrap();
                if let Some(xml_doc) = self.xml_document.upgrade() {
                    let mut doc = xml_doc.borrow_mut();
                    // Update count & uniqueCount in root
                    if let Some(root) = doc.get_root_mut() {
                        if let Some(attributes) = root.get_attribute_mut() {
                            attributes
                                .insert("count".to_string(), string_collection.len().to_string());
                            attributes.insert(
                                "uniqueCount".to_string(),
                                string_collection
                                    .iter()
                                    .map(|s| s.to_string())
                                    .collect::<HashSet<String>>()
                                    .len()
                                    .to_string(),
                            );
                        }
                    }
                    for string in string_collection {
                        let parent_id = doc.append_child_mut(&0, "si").unwrap().get_id();
                        doc.append_child_mut(&parent_id, "t")
                            .unwrap()
                            .set_value(string);
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

impl XmlDocumentPart for ShareString {
    fn new(
        office_document: Weak<RefCell<OfficeDocument>>,
        file_path: Option<String>,
    ) -> AnyResult<Self, AnyError> {
        let file_path = file_path.unwrap_or("sharedStrings.xml".to_string());
        let mut xml_document = Self::get_xml_document(&office_document, &file_path)?;
        Self::load_content_to_database(&office_document, &mut xml_document)
            .context("Load Share String To DB Failed")?;
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
            .create_root_mut("sst")
            .context("Create Root Element Failed")?
            .set_attribute_mut(attributes)
            .context("Set Attribute Failed")?;
        Ok(xml_document)
    }
}

impl ShareString {
    fn load_content_to_database(
        office_document: &Weak<RefCell<OfficeDocument>>,
        xml_document: &mut Weak<RefCell<XmlDocument>>,
    ) -> AnyResult<(), AnyError> {
        if let Some(office_doc_ref) = office_document.upgrade() {
            if let Some(create_query) =
                get_specific_queries!("share_string.sql", "create_share_string_table")
            {
                if let Some(insert_query) =
                    get_specific_queries!("share_string.sql", "insert_share_string_table")
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
                        if let Some(elements) = xml_doc_mut.pop_elements_by_tag_mut(&0, "si") {
                            for element in elements {
                                if let Some(child_id) = element.get_first_child_id() {
                                    if let Some(text_element) =
                                        xml_doc_mut.pop_element_mut(&child_id)
                                    {
                                        let value = text_element
                                            .get_value()
                                            .clone()
                                            .unwrap_or("".to_string());
                                        let _ = office_doc
                                            .get_connection()
                                            .insert_record(&insert_query, params![value])
                                            .context("Create Share String Table Failed")?;
                                    }
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
