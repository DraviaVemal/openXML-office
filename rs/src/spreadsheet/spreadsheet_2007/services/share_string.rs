use crate::element_dictionary::EXCEL_TYPE_COLLECTION;
use crate::global_2007::parts::RelationsPart;
use crate::global_2007::traits::XmlDocumentPartCommon;
use crate::{
    files::{OfficeDocument, XmlDocument},
    get_all_queries,
    global_2007::traits::XmlDocumentPart,
};
use anyhow::{anyhow, Context, Error as AnyError, Result as AnyResult};
use rusqlite::{params, Row, ToSql};
use std::{
    cell::RefCell,
    collections::{HashMap, HashSet},
    rc::Weak,
};

#[derive(Debug)]
pub struct ShareStringPart {
    office_document: Weak<RefCell<OfficeDocument>>,
    parent_relationship_part: Weak<RefCell<RelationsPart>>,
    xml_document: Weak<RefCell<XmlDocument>>,
    queries: HashMap<String, String>,
    file_path: String,
}

impl Drop for ShareStringPart {
    fn drop(&mut self) {
        let _ = self.close_document();
    }
}

impl XmlDocumentPartCommon for ShareStringPart {
    fn close_document(&mut self) -> AnyResult<(), AnyError>
    where
        Self: Sized,
    {
        if let Some(office_doc_ref) = self.office_document.upgrade() {
            let select_query = self
                .queries
                .get("select_share_string_table")
                .unwrap()
                .to_owned();
            fn row_mapper(row: &Row) -> AnyResult<String, rusqlite::Error> {
                Ok(row.get(0)?)
            }
            let string_collection = office_doc_ref
                .try_borrow()
                .context("Failed to borrow Doc Handle")?
                .get_database()
                .find_many(&select_query, params![], row_mapper, None)
                .context("Getting Share String Records Failed")?;
            if string_collection.len() > 0 {
                if let Some(xml_document) = self.xml_document.upgrade() {
                    let mut xml_doc_mut = xml_document
                        .try_borrow_mut()
                        .context("Failed to Pull Doc Reference")?;
                    // Update count & uniqueCount in root
                    if let Some(root) = xml_doc_mut.get_root_mut() {
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
                        let parent_id = xml_doc_mut
                            .append_child_mut("si", None)
                            .context("Failed to Add Child")?
                            .get_id();
                        xml_doc_mut
                            .append_child_mut("t", Some(&parent_id))
                            .context("Creating Share String Child Failed")?
                            .set_value_mut(string);
                    }
                }
                office_doc_ref
                    .try_borrow_mut()
                    .context("Failed To pull XML Handle")?
                    .close_xml_document(&self.file_path)
                    .context("Failed to close XML Document Share String")?;
            } else {
                if let Some(relationship_part) = self.parent_relationship_part.upgrade() {
                    relationship_part
                        .try_borrow_mut()
                        .context("Failed To pull parent relation ship part of Share String")?
                        .delete_relationship_mut(&self.file_path);
                    office_doc_ref
                        .try_borrow_mut()
                        .context("Failed To pull XML Handle")?
                        .delete_document_mut(&self.file_path)
                        .context("Failed to delete XML Document Share String")?;
                }
            }
        }
        Ok(())
    }
    /// Initialize xml content for this part from base template
    fn initialize_content_xml(
    ) -> AnyResult<(XmlDocument, Option<String>, Option<String>, Option<String>), AnyError> {
        let content = EXCEL_TYPE_COLLECTION.get("share_string").unwrap();
        let mut attributes: HashMap<String, String> = HashMap::new();
        attributes.insert(
            "xmlns".to_string(),
            EXCEL_TYPE_COLLECTION
                .get("share_string")
                .unwrap()
                .schemas_namespace
                .to_string(),
        );
        let mut xml_document = XmlDocument::new();
        xml_document
            .create_root_mut("sst")
            .context("Create Root Element Failed")?
            .set_attribute_mut(attributes)
            .context("Set Attribute Failed")?;
        Ok((
            xml_document,
            Some(content.content_type.to_string()),
            Some(content.extension.to_string()),
            Some(content.extension_type.to_string()),
        ))
    }
}

impl XmlDocumentPart for ShareStringPart {
    fn new(
        office_document: Weak<RefCell<OfficeDocument>>,
        parent_relationship_part: Weak<RefCell<RelationsPart>>,
    ) -> AnyResult<Self, AnyError> {
        let queries = get_all_queries!("share_string.sql");
        let file_name = Self::get_share_string_file_name(&parent_relationship_part)
            .context("Failed to pull share string file name")?
            .to_string();
        let mut xml_document = Self::get_xml_document(&office_document, &file_name)?;
        Self::load_content_to_database(&office_document, &mut xml_document, &queries)
            .context("Load Share String To DB Failed")?;
        Ok(Self {
            office_document,
            parent_relationship_part,
            xml_document,
            queries,
            file_path: file_name,
        })
    }
}

impl ShareStringPart {
    fn get_share_string_file_name(
        relations_part: &Weak<RefCell<RelationsPart>>,
    ) -> AnyResult<String, AnyError> {
        let share_string_content = EXCEL_TYPE_COLLECTION.get("share_string").unwrap();
        if let Some(relations_part) = relations_part.upgrade() {
            Ok(relations_part
                .try_borrow_mut()
                .context("Failed to pull relationship connection")?
                .get_relationship_target_by_type_mut(
                    &share_string_content.schemas_type,
                    share_string_content,
                    None,
                    None,
                )
                .context("Pull Path From Existing File Failed")?)
        } else {
            Err(anyhow!("Failed to upgrade relation part"))
        }
    }

    fn load_content_to_database(
        office_document: &Weak<RefCell<OfficeDocument>>,
        xml_document: &mut Weak<RefCell<XmlDocument>>,
        queries: &HashMap<String, String>,
    ) -> AnyResult<(), AnyError> {
        if let Some(office_doc_ref) = office_document.upgrade() {
            let mut share_string_record = Vec::new();
            let create_query = queries
                .get("create_share_string_table")
                .ok_or_else(|| anyhow!("Expected Query Not Found"))?;
            let insert_query = queries
                .get("insert_share_string_table")
                .ok_or_else(|| anyhow!("Expected Query Not Found"))?;
            let office_doc = office_doc_ref
                .try_borrow()
                .context("Pulling Office Doc Failed")?;
            office_doc
                .get_database()
                .create_table(&create_query, None)
                .context("Create Share String Table Failed")?;
            if let Some(xml_document) = xml_document.upgrade() {
                let mut xml_doc_mut = xml_document
                    .try_borrow_mut()
                    .context("xml doc borrow failed")?;
                if let Some(elements) = xml_doc_mut.pop_elements_by_tag_mut("si", None) {
                    for element in elements {
                        if let Some((child_id, _)) = element.pop_child_mut() {
                            if let Some(text_element) = xml_doc_mut.pop_element_mut(&child_id) {
                                let value =
                                    text_element.get_value().clone().unwrap_or("".to_string());
                                share_string_record.push(value);
                            }
                        }
                    }
                }
            }
            fn row_parser(share_string: String) -> Vec<Box<dyn ToSql>> {
                vec![Box::new(share_string)]
            }
            office_doc
                .get_database()
                .insert_records(&insert_query, share_string_record, row_parser, None)
                .context("Create Share String Table Failed")?;
        }
        Ok(())
    }
}

impl ShareStringPart {
    pub(crate) fn get_string_id(&self, value: String) -> AnyResult<String, AnyError> {
        if let Some(office_doc_ref) = self.office_document.upgrade() {
            let office_doc = office_doc_ref
                .try_borrow_mut()
                .context("Failed to get office handle Share String")?;
            let insert_query = self
                .queries
                .get("insert_ignore_share_string_table")
                .ok_or_else(|| anyhow!("Expected Query Not Found"))?;
            office_doc
                .get_database()
                .insert_record(&insert_query, params![value], None)
                .context("Create Share String Table Failed")?;
            let find_query = self
                .queries
                .get("select_find_share_string_table")
                .ok_or_else(|| anyhow!("Expected Query Not Found"))?;
            fn row_mapper(row: &Row) -> AnyResult<usize, rusqlite::Error> {
                Ok(row.get(0)?)
            }
            let result = office_doc.get_database().find_one(
                &find_query,
                params![value],
                row_mapper,
                None,
            )?;
            if let Some(id) = result {
                Ok((id - 1).to_string())
            } else {
                Err(anyhow!("Failed to Get Share String id"))
            }
        } else {
            Err(anyhow!("Failed to Get Office Doc Handle"))
        }
    }
}
