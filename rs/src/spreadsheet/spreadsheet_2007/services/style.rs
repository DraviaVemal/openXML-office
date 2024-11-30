use crate::global_2007::traits::XmlDocumentPartCommon;
use crate::{
    files::{OfficeDocument, XmlDocument},
    get_all_queries,
    global_2007::traits::XmlDocumentPart,
};
use anyhow::{anyhow, Context, Error as AnyError, Result as AnyResult};
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
    fn load_content_to_database(
        office_document: &Weak<RefCell<OfficeDocument>>,
        xml_document: &mut Weak<RefCell<XmlDocument>>,
    ) -> AnyResult<(), AnyError> {
        if let Some(office_doc_ref) = office_document.upgrade() {
            // Decode XML to load in DB
            let office_doc = office_doc_ref
                .try_borrow()
                .context("Pulling Office Doc Failed")?;
            // Load Required Queries
            let queries = get_all_queries!("style.sql");
            // Create Required Tables Queries
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
            // Insert Tables Queries
            let insert_query_num_format = queries
                .get("insert_number_format_table")
                .ok_or_else(|| anyhow!("Expected Query Not Found"))?;
            let insert_query_font_style = queries
                .get("insert_number_format_table")
                .ok_or_else(|| anyhow!("Expected Query Not Found"))?;
            let insert_query_fill_style = queries
                .get("insert_number_format_table")
                .ok_or_else(|| anyhow!("Expected Query Not Found"))?;
            let insert_query_border_style = queries
                .get("insert_number_format_table")
                .ok_or_else(|| anyhow!("Expected Query Not Found"))?;
            let insert_query_cell_style = queries
                .get("insert_number_format_table")
                .ok_or_else(|| anyhow!("Expected Query Not Found"))?;

            if let Some(xml_doc) = xml_document.upgrade() {
                let mut xml_doc_mut = xml_doc.try_borrow_mut().context("xml doc borrow failed")?;
                let number_formats = xml_doc_mut
                    .pop_elements_by_tag_mut("numFmts", None)
                    .context("Failed find the Target node")?;
                if let Some(number_formats_id) = number_formats {}
                let fonts = xml_doc_mut
                    .get_first_element_id(vec!["styleSheet", "fonts"], None)
                    .context("Failed find the Target node")?;
                if let Some(fonts_id) = fonts {}
                let fills = xml_doc_mut
                    .get_first_element_id(vec!["styleSheet", "fills"], None)
                    .context("Failed find the Target node")?;
                if let Some(fills_id) = fills {}
                let borders = xml_doc_mut
                    .get_first_element_id(vec!["styleSheet", "borders"], None)
                    .context("Failed find the Target node")?;
                if let Some(borders_id) = borders {}
                let cell_xfs = xml_doc_mut
                    .get_first_element_id(vec!["styleSheet", "cell_xfs"], None)
                    .context("Failed find the Target node")?;
                if let Some(cell_xfs_id) = cell_xfs {}
                let cell_style_xfs = xml_doc_mut
                    .get_first_element_id(vec!["styleSheet", "cell_style_xfs"], None)
                    .context("Failed find the Target node")?;
                if let Some(cell_style_xfs_id) = cell_style_xfs {}
                let cell_styles = xml_doc_mut
                    .get_first_element_id(vec!["styleSheet", "cell_styles"], None)
                    .context("Failed find the Target node")?;
                if let Some(cell_styles_id) = cell_styles {}
                let dxfs = xml_doc_mut
                    .get_first_element_id(vec!["styleSheet", "dxfs"], None)
                    .context("Failed find the Target node")?;
                if let Some(dxfs_id) = dxfs {}
                let table_styles = xml_doc_mut
                    .get_first_element_id(vec!["styleSheet", "table_styles"], None)
                    .context("Failed find the Target node")?;
                if let Some(table_styles_id) = table_styles {}
            }
        }
        Ok(())
    }
}
