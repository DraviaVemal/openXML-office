use crate::files::XmlElement;
use crate::global_2007::traits::XmlDocumentPartCommon;
use crate::{
    files::{OfficeDocument, XmlDocument},
    get_all_queries,
    global_2007::traits::XmlDocumentPart,
};
use anyhow::{anyhow, Context, Error as AnyError, Result as AnyResult};
use rusqlite::params;
use std::{cell::RefCell, collections::HashMap, rc::Weak};

#[derive(Debug)]
pub struct Style {
    office_document: Weak<RefCell<OfficeDocument>>,
    xml_document: Weak<RefCell<XmlDocument>>,
    file_path: String,
}

impl Drop for Style {
    fn drop(&mut self) {
        if let Some(xml_tree) = self.office_document.upgrade() {
            let _ = xml_tree
                .try_borrow_mut()
                .unwrap()
                .close_xml_document(&self.file_path);
        }
    }
}

impl XmlDocumentPartCommon for Style {
    /// Initialize xml content for this part from base template
    fn initialize_content_xml() -> AnyResult<XmlDocument, AnyError> {
        let mut attributes: HashMap<String, String> = HashMap::new();
        attributes.insert(
            "xmlns".to_string(),
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main".to_string(),
        );
        attributes.insert(
            "xmlns:mc".to_string(),
            "http://schemas.openxmlformats.org/markup-compatibility/2006".to_string(),
        );
        let mut xml_document = XmlDocument::new();
        xml_document
            .create_root_mut("styleSheet")
            .context("Create XML Root Element Failed")?
            .set_attribute_mut(attributes)
            .context("Set Attribute Failed")?;
        Ok(xml_document)
    }
}

impl XmlDocumentPart for Style {
    fn new(
        office_document: Weak<RefCell<OfficeDocument>>,
        file_path: Option<String>,
    ) -> AnyResult<Self, AnyError> {
        let file_path = file_path.unwrap_or("xl/styles.xml".to_string());
        let mut xml_document = Self::get_xml_document(&office_document, &file_path)?;
        Self::load_content_to_database(&office_document, &mut xml_document)
            .context("Load Share String To DB Failed")?;
        Ok(Self {
            office_document,
            xml_document,
            file_path,
        })
    }
}

impl Style {
    fn initialize_database(
        office_document: &Weak<RefCell<OfficeDocument>>,
        queries: &HashMap<String, String>,
    ) -> AnyResult<(), AnyError> {
        if let Some(office_doc_ref) = office_document.upgrade() {
            // Decode XML to load in DB
            let office_doc = office_doc_ref
                .try_borrow()
                .context("Pulling Office Doc Failed")?;
            let create_query_num_format = queries
                .get("create_number_format_table")
                .ok_or_else(|| anyhow!("Expected Query Not Found"))?;
            let create_query_font_style = queries
                .get("create_font_style_table")
                .ok_or_else(|| anyhow!("Expected Query Not Found"))?;
            let create_query_fill_style = queries
                .get("create_fill_style_table")
                .ok_or_else(|| anyhow!("Expected Query Not Found"))?;
            let create_query_border_style = queries
                .get("create_border_style_table")
                .ok_or_else(|| anyhow!("Expected Query Not Found"))?;
            let create_query_cell_style = queries
                .get("create_cell_style_table")
                .ok_or_else(|| anyhow!("Expected Query Not Found"))?;
            office_doc
                .get_connection()
                .create_table(&create_query_num_format)
                .context("Create Number Format Table Failed")?;
            office_doc
                .get_connection()
                .create_table(&create_query_font_style)
                .context("Create Font Style Table Failed")?;
            office_doc
                .get_connection()
                .create_table(&create_query_fill_style)
                .context("Create Query Fill Table Failed")?;
            office_doc
                .get_connection()
                .create_table(&create_query_border_style)
                .context("Create Border Style Table Failed")?;
            office_doc
                .get_connection()
                .create_table(&create_query_cell_style)
                .context("Create Cell Style Table Failed")?;
        }
        Ok(())
    }

    fn load_content_to_database(
        office_document: &Weak<RefCell<OfficeDocument>>,
        xml_document: &mut Weak<RefCell<XmlDocument>>,
    ) -> AnyResult<(), AnyError> {
        if let Some(office_doc_ref) = office_document.upgrade() {
            fn extract_attribute_value(
                attribute_keys: Vec<&str>,
                row: &mut Vec<String>,
                element: &mut Vec<XmlElement>,
            ) {
                if let Some(scheme) = element.pop() {
                    if let Some(attributes) = scheme.get_attribute() {
                        for attribute_key in attribute_keys {
                            row.push(
                                attributes
                                    .get(attribute_key)
                                    .unwrap_or(&"".to_string())
                                    .to_string(),
                            );
                        }
                    }
                }
            }

            // Decode XML to load in DB
            let office_doc = office_doc_ref
                .try_borrow()
                .context("Pulling Office Doc Failed")?;
            // Load Required Queries
            let queries = get_all_queries!("style.sql");
            Self::initialize_database(office_document, &queries)
                .context("Database Initialization Failed")?;

            if let Some(xml_doc) = xml_document.upgrade() {
                let mut xml_doc_mut = xml_doc.try_borrow_mut().context("xml doc borrow failed")?;
                // Load Number Format Region
                if let Some(number_formats) = xml_doc_mut
                    .pop_elements_by_tag_mut("numFmts", None)
                    .context("Failed find the Target node")?
                    .pop()
                {
                    loop {
                        if let Some(element_id) = number_formats.pop_child_id_mut() {
                            let num_fmt = xml_doc_mut
                                .pop_element_mut(&element_id)
                                .ok_or(anyhow!("Element not Found Error"))?;
                            if let Some(attributes) = num_fmt.get_attribute() {
                                let insert_query_num_format = queries
                                    .get("insert_number_format_table")
                                    .ok_or_else(|| anyhow!("Expected Query Not Found"))?;
                                office_doc
                                    .get_connection()
                                    .insert_record(
                                        &insert_query_num_format,
                                        params![
                                            attributes
                                                .get("numFmtId")
                                                .ok_or(anyhow!("numFmtId Attribute Not Found!"))?,
                                            attributes.get("formatCode").ok_or(anyhow!(
                                                "formatCode Attribute Not Found!"
                                            ))?
                                        ],
                                    )
                                    .context("Number Format Data Insert Failed")?;
                            }
                        } else {
                            break;
                        }
                    }
                }
                // Load Font and its child into single row
                let mut font_row: Vec<String> = Vec::new();
                if let Some(fonts) = xml_doc_mut
                    .pop_elements_by_tag_mut("fonts", None)
                    .context("Failed find the Target node")?
                    .pop()
                {
                    // fonts
                    loop {
                        if let Some(font_id) = fonts.pop_child_id_mut() {
                            // font
                            let font = xml_doc_mut
                                .pop_element_mut(&font_id)
                                .ok_or(anyhow!("Element not Found Error"))?;
                            if let Some(mut sz) =
                                xml_doc_mut.pop_elements_by_tag_mut("sz", Some(&font.get_id()))
                            {
                                extract_attribute_value(vec!["val"], &mut font_row, &mut sz);
                            }
                            if let Some(mut color) =
                                xml_doc_mut.pop_elements_by_tag_mut("color", Some(&font.get_id()))
                            {
                                extract_attribute_value(
                                    vec!["rgb", "theme"],
                                    &mut font_row,
                                    &mut color,
                                );
                            }
                            if let Some(mut name) =
                                xml_doc_mut.pop_elements_by_tag_mut("name", Some(&font.get_id()))
                            {
                                extract_attribute_value(vec!["val"], &mut font_row, &mut name);
                            }
                            if let Some(mut family) =
                                xml_doc_mut.pop_elements_by_tag_mut("family", Some(&font.get_id()))
                            {
                                extract_attribute_value(vec!["val"], &mut font_row, &mut family);
                            }
                            if let Some(mut scheme) =
                                xml_doc_mut.pop_elements_by_tag_mut("scheme", Some(&font.get_id()))
                            {
                                extract_attribute_value(vec!["val"], &mut font_row, &mut scheme);
                            }
                            // Insert Data into Database
                            let insert_query_font_style = queries
                                .get("insert_number_format_table")
                                .ok_or_else(|| anyhow!("Expected Query Not Found"))?;
                            office_doc
                                .get_connection()
                                .insert_record(&insert_query_font_style, params![])
                                .context("Insert Font Style Failed")?;
                        } else {
                            break;
                        }
                    }
                }

                // Insert Tables Queries
                let insert_query_fill_style = queries
                    .get("insert_number_format_table")
                    .ok_or_else(|| anyhow!("Expected Query Not Found"))?;
                let insert_query_border_style = queries
                    .get("insert_number_format_table")
                    .ok_or_else(|| anyhow!("Expected Query Not Found"))?;
                let insert_query_cell_style = queries
                    .get("insert_number_format_table")
                    .ok_or_else(|| anyhow!("Expected Query Not Found"))?;
                if let Some(fill) = xml_doc_mut
                    .pop_elements_by_tag_mut("fills", None)
                    .context("Failed find the Target node")?
                    .pop()
                {}
                if let Some(border) = xml_doc_mut
                    .pop_elements_by_tag_mut("borders", None)
                    .context("Failed find the Target node")?
                    .pop()
                {}
                if let Some(cell_xf) = xml_doc_mut
                    .pop_elements_by_tag_mut("cell_xfs", None)
                    .context("Failed find the Target node")?
                    .pop()
                {}
                if let Some(cell_style_xf) = xml_doc_mut
                    .pop_elements_by_tag_mut("cell_style_xfs", None)
                    .context("Failed find the Target node")?
                    .pop()
                {}
                if let Some(cell_style) = xml_doc_mut
                    .pop_elements_by_tag_mut("cell_styles", None)
                    .context("Failed find the Target node")?
                    .pop()
                {}
                if let Some(dxf) = xml_doc_mut
                    .pop_elements_by_tag_mut("dxfs", None)
                    .context("Failed find the Target node")?
                    .pop()
                {}
                if let Some(table_style) = xml_doc_mut
                    .pop_elements_by_tag_mut("table_styles", None)
                    .context("Failed find the Target node")?
                    .pop()
                {}
            }
        }
        Ok(())
    }
}
