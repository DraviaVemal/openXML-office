use crate::{
    converters::ConverterUtil,
    element_dictionary::EXCEL_TYPE_COLLECTION,
    files::{OfficeDocument, XmlDocument, XmlSerializer},
    get_all_queries,
    global_2007::{
        parts::RelationsPart,
        traits::{Enum, XmlDocumentPartCommon},
    },
    order_dictionary::EXCEL_ORDER_COLLECTION,
    spreadsheet_2007::{
        models::{
            CellDataType, CellRecord, ColumnCell, ColumnProperties, RowProperties, RowRecord,
        },
        services::CommonServices,
    },
};
use anyhow::{anyhow, Context, Error as AnyError, Result as AnyResult};
use rusqlite::{params, Row};
use std::{
    cell::RefCell,
    collections::HashMap,
    rc::{Rc, Weak},
};

#[derive(Debug)]
pub struct WorkSheet {
    queries: HashMap<String, String>,
    office_document: Weak<RefCell<OfficeDocument>>,
    xml_document: Weak<RefCell<XmlDocument>>,
    common_service: Weak<RefCell<CommonServices>>,
    workbook_relationship_part: Weak<RefCell<RelationsPart>>,
    sheet_collection: Weak<RefCell<Vec<(String, String, bool, bool)>>>,
    sheet_relationship_part: Rc<RefCell<RelationsPart>>,
    // sheet_property: Option<_>,
    // sheet_view: Option<_>,
    // sheet_format_property: Option<_>,
    column_collection: Option<Vec<ColumnProperties>>,
    // sheet_calculation_property:Option<_>
    // protected_range:Option<_>
    // merge_cells:Option<_>
    // hyperlinks:Option<_>
    file_path: String,
    sheet_name: String,
}

impl Drop for WorkSheet {
    fn drop(&mut self) {
        let _ = self.close_document();
    }
}

impl XmlDocumentPartCommon for WorkSheet {
    /// Initialize xml content for this part from base template
    fn initialize_content_xml(
    ) -> AnyResult<(XmlDocument, Option<String>, Option<String>, Option<String>), AnyError> {
        let content = EXCEL_TYPE_COLLECTION.get("worksheet").unwrap();
        let template_core_properties = include_str!("worksheet.xml");
        Ok((
            XmlSerializer::vec_to_xml_doc_tree(template_core_properties.as_bytes().to_vec())
                .context("Initializing Worksheet Failed")?,
            Some(content.content_type.to_string()),
            Some(content.extension.to_string()),
            Some(content.extension_type.to_string()),
        ))
    }
    fn close_document(&mut self) -> AnyResult<(), AnyError>
    where
        Self: Sized,
    {
        if let Some(office_document) = self.office_document.upgrade() {
            let mut office_doc_mut = office_document
                .try_borrow_mut()
                .context("Failed to pull office document")?;
            if let Some(xml_document) = self.xml_document.upgrade() {
                let mut xml_doc_mut = xml_document
                    .try_borrow_mut()
                    .context("Failed to Pull XML Handle")?;
                // Add Cols Record to Document
                {
                    if let Some(mut column_collection) = self.column_collection.take() {
                        let cols_id = xml_doc_mut
                            .insert_children_after_tag_mut("cols", "sheetViews", None)
                            .context("Failed to Insert Cols Element")?
                            .get_id();
                        loop {
                            if let Some(item) = column_collection.pop() {
                                let col_element = xml_doc_mut
                                    .append_child_mut("col", Some(&cols_id))
                                    .context("Failed to insert col record")?;
                                let mut attribute = HashMap::new();
                                attribute.insert("min".to_string(), item.min.to_string());
                                attribute.insert("max".to_string(), item.max.to_string());
                                if let Some(width) = item.width {
                                    attribute.insert("customWidth".to_string(), "1".to_string());
                                    attribute.insert("width".to_string(), width.to_string());
                                }
                                if let Some(style_id) = item.style_id {
                                    attribute.insert("style".to_string(), style_id.to_string());
                                }
                                if let Some(_) = item.best_fit {
                                    attribute.insert("bestFit".to_string(), "1".to_string());
                                }
                                col_element
                                    .set_attribute_mut(attribute)
                                    .context("Failed to Add Attribute to col element")?;
                            } else {
                                break;
                            }
                        }
                    }
                }
                // Add Sheet Data to Document
                {
                    let select_all_query = self
                        .queries
                        .get("select_all_dynamic_sheet")
                        .ok_or(anyhow!("Failed to Get Select All Query"))?;
                    fn row_mapper(
                        row: &Row,
                    ) -> AnyResult<(RowRecord, CellRecord), rusqlite::Error> {
                        Ok((
                            RowRecord {
                                row_id: row.get(0)?,
                                row_hide: row.get(1)?,
                                row_span: row.get(2)?,
                                row_height: row.get(3)?,
                                row_style_id: row.get(4)?,
                                row_thick_top: row.get(5)?,
                                row_thick_bottom: row.get(6)?,
                                row_group_level: row.get(7)?,
                                row_collapsed: row.get(8)?,
                                row_place_holder: row.get(9)?,
                            },
                            CellRecord {
                                row_id: row.get(0)?,
                                col_id: row.get(10)?,
                                cell_style_id: row.get(11)?,
                                cell_value: row.get(12)?,
                                cell_formula: row.get(13)?,
                                cell_type: if let Some(value) = row.get(14)? {
                                    let type_text: String = value;
                                    Some(CellDataType::get_enum(&type_text))
                                } else {
                                    None
                                },
                                cell_metadata: row.get(15)?,
                                cell_place_holder: row.get(16)?,
                                cell_comment_id: row.get(17)?,
                            },
                        ))
                    }
                    let mut sheet_data = office_doc_mut
                        .get_connection()
                        .find_many(
                            &select_all_query.replace("{0}", "{}_row"),
                            params![],
                            row_mapper,
                            Some(
                                self.file_path
                                    .rsplit("/")
                                    .next()
                                    .unwrap()
                                    .to_string()
                                    .replace(".xml", ""),
                            ),
                        )
                        .context("Failed to Get Sheet Data Records")?;
                    let sheet_data_id = xml_doc_mut
                        .insert_children_after_tag_mut("sheetData", "cols", None)
                        .context("Failed to Insert Cols Element")?
                        .get_id();
                    if sheet_data.len() > 0 {
                        let mut row_element_id = 0;
                        let mut row_id = 0;
                        loop {
                            if let Some((db_row, db_cell)) = sheet_data.pop() {
                                // Create row element
                                if row_id != db_row.row_id {
                                    let row_element = xml_doc_mut
                                        .append_child_mut("row", Some(&sheet_data_id))
                                        .context("Failed to insert row element")?;
                                    row_element_id = row_element.get_id();
                                    row_id = db_row.row_id;
                                    let mut row_attribute = HashMap::new();
                                    row_attribute
                                        .insert("r".to_string(), db_row.row_id.to_string());
                                    if let Some(row_span) = db_row.row_span {
                                        row_attribute.insert("spans".to_string(), row_span);
                                    }
                                    if let Some(row_style_id) = db_row.row_style_id {
                                        row_attribute
                                            .insert("customFormat".to_string(), "1".to_string());
                                        row_attribute
                                            .insert("s".to_string(), row_style_id.to_string());
                                    }
                                    if let Some(row_height) = db_row.row_height {
                                        row_attribute
                                            .insert("customHeight".to_string(), "1".to_string());
                                        row_attribute
                                            .insert("ht".to_string(), row_height.to_string());
                                    }
                                    if let Some(_) = db_row.row_hide {
                                        row_attribute.insert("hidden".to_string(), "1".to_string());
                                    }
                                    if let Some(row_group_level) = db_row.row_group_level {
                                        row_attribute.insert(
                                            "outlineLevel".to_string(),
                                            row_group_level.to_string(),
                                        );
                                    }
                                    if let Some(_) = db_row.row_collapsed {
                                        row_attribute
                                            .insert("collapsed".to_string(), "1".to_string());
                                    }
                                    if let Some(_) = db_row.row_thick_top {
                                        row_attribute
                                            .insert("thickTop".to_string(), "1".to_string());
                                    }
                                    if let Some(_) = db_row.row_thick_bottom {
                                        row_attribute
                                            .insert("thickBot".to_string(), "1".to_string());
                                    }
                                    if let Some(_) = db_row.row_place_holder {
                                        row_attribute.insert("ph".to_string(), "1".to_string());
                                    }
                                    row_element
                                        .set_attribute_mut(row_attribute)
                                        .context("Failed to set attribute for row")?;
                                }
                                let cell_id;
                                // Create cell element
                                {
                                    let cell_element = xml_doc_mut
                                        .append_child_mut("c", Some(&row_element_id))
                                        .context("Failed to insert row element")?;
                                    cell_id = cell_element.get_id();
                                    let mut cell_attribute = HashMap::new();
                                    cell_attribute.insert(
                                        "r".to_string(),
                                        format!(
                                            "{}{}",
                                            ConverterUtil::get_column_key(db_cell.col_id)
                                                .context("Failed to get Char Id from Int")?,
                                            db_row.row_id
                                        ),
                                    );
                                    if let Some(cell_style_id) = db_cell.cell_style_id {
                                        cell_attribute
                                            .insert("s".to_string(), cell_style_id.to_string());
                                    }
                                    if let Some(cell_type) = db_cell.cell_type {
                                        cell_attribute.insert(
                                            "t".to_string(),
                                            CellDataType::get_string(cell_type),
                                        );
                                    }
                                    if let Some(cell_comment_id) = db_cell.cell_comment_id {
                                        cell_attribute
                                            .insert("cm".to_string(), cell_comment_id.to_string());
                                    }
                                    if let Some(cell_metadata) = db_cell.cell_metadata {
                                        cell_attribute
                                            .insert("vm".to_string(), cell_metadata.to_string());
                                    }
                                    if let Some(_) = db_cell.cell_place_holder {
                                        cell_attribute.insert("ph".to_string(), "1".to_string());
                                    }
                                    cell_element
                                        .set_attribute_mut(cell_attribute)
                                        .context("Failed to Set Attribute for cell")?;
                                }
                                // Create cell's child element
                                match db_cell.cell_type.unwrap_or(CellDataType::Number) {
                                    CellDataType::InlineString => {
                                        let inline_string_id = xml_doc_mut
                                            .append_child_mut("is", Some(&cell_id))
                                            .context("Failed to insert Inline string element")?
                                            .get_id();
                                        let text_element = xml_doc_mut
                                            .append_child_mut("t", Some(&inline_string_id))
                                            .context(
                                                "Failed To insert Text Value to inline string",
                                            )?;
                                        text_element.set_value_mut(
                                            if let Some(value) = db_cell.cell_value {
                                                value
                                            } else {
                                                "".to_string()
                                            },
                                        );
                                    }
                                    _ => {
                                        if let Some(formula) = db_cell.cell_formula {
                                            let formula_element = xml_doc_mut
                                                .append_child_mut("f", Some(&cell_id))
                                                .context(
                                                    "Failed to insert Inline string element",
                                                )?;
                                            formula_element.set_value_mut(formula);
                                        }
                                        let value_element = xml_doc_mut
                                            .append_child_mut("v", Some(&cell_id))
                                            .context("Failed to insert Inline string element")?;
                                        value_element.set_value_mut(
                                            if let Some(value) = db_cell.cell_value {
                                                value
                                            } else {
                                                "".to_string()
                                            },
                                        );
                                    }
                                }
                            } else {
                                break;
                            }
                        }
                    }
                }
                // Add dimension
                {
                    let root_id = xml_doc_mut
                        .get_root()
                        .context("Failed to Get Root Element")?
                        .get_id();
                    let dim_query = self
                        .queries
                        .get("select_dimension_dynamic_sheet")
                        .ok_or(anyhow!("Failed to get query"))?;
                    fn row_mapper(
                        row: &Row,
                    ) -> Result<(usize, usize, usize, usize), rusqlite::Error> {
                        Ok((row.get(0)?, row.get(1)?, row.get(2)?, row.get(3)?))
                    }
                    let mut dim_range = office_doc_mut
                        .get_connection()
                        .find_many(
                            dim_query,
                            params![],
                            row_mapper,
                            Some(
                                self.file_path
                                    .rsplit("/")
                                    .next()
                                    .unwrap()
                                    .to_string()
                                    .replace(".xml", ""),
                            ),
                        )
                        .context("Failed to pull Dim Range Query Results")?;
                    if let Some((start_row_id, start_col_id, end_row_id, end_col_id)) =
                        dim_range.pop()
                    {
                        let dim_element = xml_doc_mut
                            .append_child_mut("dimension", Some(&root_id))
                            .context("Failed to Add Dimension node to worksheet")?;
                        let mut dimension_attribute = HashMap::new();
                        dimension_attribute.insert(
                            "ref".to_string(),
                            format!(
                                "{}{}:{}{}",
                                ConverterUtil::get_column_key(start_col_id)
                                    .context("Failed to convert dim col start")?,
                                start_row_id,
                                ConverterUtil::get_column_key(end_col_id)
                                    .context("Failed to convert dim col end")?,
                                end_row_id
                            ),
                        );
                        dim_element
                            .set_attribute_mut(dimension_attribute)
                            .context("Failed to set attribute value to dimension")?;
                    }
                }
                if let Some(root_element) = xml_doc_mut.get_root_mut() {
                    root_element
                        .order_child_mut(
                            EXCEL_ORDER_COLLECTION
                                .get("worksheet")
                                .ok_or(anyhow!("Failed to get worksheet default order"))?,
                        )
                        .context("Failed Reorder the element child's")?;
                }
            }
            office_doc_mut
                .close_xml_document(&self.file_path)
                .context("Failed to close the current tree document")?;
            office_doc_mut.get_connection().drop_table(
                self.file_path
                    .rsplit("/")
                    .next()
                    .unwrap()
                    .to_string()
                    .replace(".xml", ""),
            )?;
            office_doc_mut.get_connection().drop_table(format!(
                "{}_row",
                self.file_path
                    .rsplit("/")
                    .next()
                    .unwrap()
                    .to_string()
                    .replace(".xml", "")
            ))?;
        }
        Ok(())
    }
}

// ############################# Internal Function ######################################
impl WorkSheet {
    /// Create New object for the group
    pub(crate) fn new(
        office_document: Weak<RefCell<OfficeDocument>>,
        sheet_collection: Weak<RefCell<Vec<(String, String, bool, bool)>>>,
        workbook_relationship_part: Weak<RefCell<RelationsPart>>,
        common_service: Weak<RefCell<CommonServices>>,
        sheet_name: Option<String>,
    ) -> AnyResult<Self, AnyError> {
        let (file_path, sheet_name) = Self::get_sheet_file_name(
            sheet_name,
            &office_document,
            &sheet_collection,
            &workbook_relationship_part,
        )
        .context("Failed to pull calc chain file name")?;
        let xml_document = Self::get_xml_document(&office_document, &file_path)?;
        let sheet_relationship_part = Rc::new(RefCell::new(
            RelationsPart::new(
                office_document.clone(),
                &format!(
                    "{}/_rels/{}.rels",
                    &file_path[..file_path.rfind("/").unwrap()],
                    file_path.rsplit("/").next().unwrap()
                ),
            )
            .context("Creating Relation ship part for workbook failed.")?,
        ));
        let queries = get_all_queries!("worksheet.sql");
        let column_collection = Self::initialize_worksheet(
            file_path
                .rsplit("/")
                .next()
                .unwrap()
                .to_string()
                .replace(".xml", ""),
            &queries,
            &office_document,
            &xml_document,
        )
        .context("Failed to open Worksheet")?;
        Ok(Self {
            queries,
            office_document,
            xml_document,
            common_service,
            workbook_relationship_part,
            sheet_relationship_part,
            sheet_collection,
            column_collection,
            file_path: file_path.to_string(),
            sheet_name,
        })
    }

    fn initialize_worksheet(
        table_name: String,
        queries: &HashMap<String, String>,
        office_document: &Weak<RefCell<OfficeDocument>>,
        xml_document: &Weak<RefCell<XmlDocument>>,
    ) -> AnyResult<Option<Vec<ColumnProperties>>, AnyError> {
        let mut column_collection = Vec::new();
        if let Some(office_document) = office_document.upgrade() {
            if let Some(xml_document) = xml_document.upgrade() {
                let mut xml_doc_mut = xml_document
                    .try_borrow_mut()
                    .context("Failed to get XML doc handle")?;
                let office_doc_mut = office_document
                    .try_borrow_mut()
                    .context("Failed to Get office doc handle")?;
                let create_query = queries.get("create_dynamic_sheet").ok_or(anyhow!(
                    "Failed to Get the create query at sheet data parser"
                ))?;
                let create_query_row = queries.get("create_dynamic_sheet_row").ok_or(anyhow!(
                    "Failed to Get the create query at sheet data parser"
                ))?;
                office_doc_mut
                    .get_connection()
                    .create_table(&create_query, Some(table_name.clone()))
                    .context("Failed to create sheet data table")?;
                office_doc_mut
                    .get_connection()
                    .create_table(&create_query_row, Some(format!("{}_row", table_name)))
                    .context("Failed to create sheet data row table")?;
                // unwrap columns to local collection
                {
                    if let Some(mut cols_element) =
                        xml_doc_mut.pop_elements_by_tag_mut("cols", None)
                    {
                        // Process the columns record if parent node exist
                        if let Some(cols) = cols_element.pop() {
                            loop {
                                if let Some((col_elements, _)) = cols.pop_child_mut() {
                                    if let Some(col) = xml_doc_mut.pop_element_mut(&col_elements) {
                                        let mut column_properties = ColumnProperties::default();
                                        let attributes = col
                                            .get_attribute()
                                            .ok_or(anyhow!("Error Getting col attribute"))?;
                                        if let Some(min) = attributes.get("min") {
                                            column_properties.min =
                                                min.parse().context("Failed to parse min value")?;
                                        }
                                        if let Some(max) = attributes.get("max") {
                                            column_properties.max =
                                                max.parse().context("Failed to parse min value")?;
                                        }
                                        if let Some(best_fit) = attributes.get("bestFit") {
                                            column_properties.best_fit =
                                                if best_fit == "1" { Some(()) } else { None }
                                        }
                                        if let Some(hidden) = attributes.get("hidden") {
                                            column_properties.hide =
                                                if hidden == "1" { Some(()) } else { None }
                                        }
                                        if let Some(style) = attributes.get("style") {
                                            column_properties.style_id = Some(
                                                style
                                                    .parse()
                                                    .context("Failed to parse style ID")?,
                                            );
                                        }
                                        if let Some(outline_level) = attributes.get("outlineLevel")
                                        {
                                            column_properties.group_level =
                                                outline_level
                                                    .parse()
                                                    .context("Failed to parse style ID")?;
                                        }
                                        if let Some(custom_width) = attributes.get("customWidth") {
                                            if custom_width == "1" {
                                                column_properties.width = Some(
                                                    attributes
                                                        .get("width")
                                                        .ok_or(anyhow!(
                                                            "Failed to get custom width"
                                                        ))?
                                                        .parse()
                                                        .context("Failed to parse custom width")?,
                                                );
                                            }
                                        }
                                        if let Some(collapsed) = attributes.get("collapsed") {
                                            column_properties.collapsed =
                                                if collapsed == "1" { Some(()) } else { None }
                                        }
                                        column_collection.push(column_properties);
                                    }
                                } else {
                                    break;
                                }
                            }
                        }
                    }
                }
                // unwrap sheet data into database
                {
                    if let Some(mut sheet_data_element) =
                        xml_doc_mut.pop_elements_by_tag_mut("sheetData", None)
                    {
                        if let Some(sheet_data) = sheet_data_element.pop() {
                            // Loop All rows of sheet data
                            loop {
                                if let Some((row_id, _)) = sheet_data.pop_child_mut() {
                                    if let Some(row_element) = xml_doc_mut.pop_element_mut(&row_id)
                                    {
                                        let mut db_row_record = RowRecord::default();
                                        let row_attribute = row_element
                                            .get_attribute()
                                            .ok_or(anyhow!("Failed to pull Row Attribute."))?;
                                        // Get Row Id
                                        db_row_record.row_id = row_attribute
                                            .get("r")
                                            .ok_or(anyhow!("Missing mandatory row id attribute"))?
                                            .parse()
                                            .context("Failed to parse row id")?;
                                        if let Some(row_span) = row_attribute.get("spans") {
                                            db_row_record.row_span = Some(row_span.to_string());
                                        }
                                        if let Some(style_id) = row_attribute.get("s") {
                                            if let Some(custom_formant) =
                                                row_attribute.get("customFormat")
                                            {
                                                db_row_record.row_style_id =
                                                    if custom_formant == "1" {
                                                        Some(style_id.parse().context(
                                                            "Failed to parse the row style id",
                                                        )?)
                                                    } else {
                                                        None
                                                    };
                                            }
                                        }
                                        if let Some(hidden) = row_attribute.get("hidden") {
                                            db_row_record.row_hide =
                                                if hidden == "1" { Some(true) } else { None };
                                        }
                                        if let Some(height) = row_attribute.get("ht") {
                                            if let Some(custom_height) =
                                                row_attribute.get("customHeight")
                                            {
                                                db_row_record.row_height = if custom_height == "1" {
                                                    Some(height.parse().context(
                                                        "Failed to parse the row height",
                                                    )?)
                                                } else {
                                                    None
                                                };
                                            }
                                        }
                                        if let Some(row_group_level) =
                                            row_attribute.get("outlineLevel")
                                        {
                                            let outline_level = row_group_level
                                                .parse()
                                                .context("Failed to parse the row group level")?;
                                            db_row_record.row_group_level = if outline_level > 0 {
                                                Some(outline_level)
                                            } else {
                                                None
                                            };
                                        }
                                        if let Some(collapsed) = row_attribute.get("collapsed") {
                                            db_row_record.row_collapsed =
                                                if collapsed == "1" { Some(true) } else { None };
                                        }
                                        if let Some(thick_top) = row_attribute.get("thickTop") {
                                            db_row_record.row_thick_top =
                                                if thick_top == "1" { Some(true) } else { None };
                                        }
                                        if let Some(thick_bottom) = row_attribute.get("thickBot") {
                                            db_row_record.row_thick_bottom = if thick_bottom == "1"
                                            {
                                                Some(true)
                                            } else {
                                                None
                                            };
                                        }
                                        if let Some(place_holder) = row_attribute.get("ph") {
                                            db_row_record.row_place_holder = if place_holder == "1"
                                            {
                                                Some(true)
                                            } else {
                                                None
                                            };
                                        }
                                        // Insert record to the Database
                                        let insert_query_row = queries.get("insert_dynamic_sheet_row")
                                                    .ok_or(anyhow!("Failed to Get the insert Row Query at sheet data parser"))?;
                                        office_doc_mut
                                            .get_connection()
                                            .insert_record(
                                                insert_query_row,
                                                params![
                                                    db_row_record.row_id,
                                                    db_row_record.row_hide,
                                                    db_row_record.row_span,
                                                    db_row_record.row_height,
                                                    db_row_record.row_style_id,
                                                    db_row_record.row_thick_top,
                                                    db_row_record.row_thick_bottom,
                                                    db_row_record.row_group_level,
                                                    db_row_record.row_collapsed,
                                                    db_row_record.row_place_holder,
                                                ],
                                                Some(format!("{}_row", table_name)),
                                            )
                                            .context(
                                                "Failed to insert Cell Data record into Sheet DB",
                                            )?;
                                        // Loop All Columns of row
                                        loop {
                                            let mut db_cell_record = CellRecord::default();
                                            if let Some((col_id, _)) = row_element.pop_child_mut() {
                                                if let Some(col_element) =
                                                    xml_doc_mut.pop_element_mut(&col_id)
                                                {
                                                    let cell_attribute = col_element
                                                        .get_attribute()
                                                        .ok_or(anyhow!(
                                                            "Failed to pull attribute for Column"
                                                        ))?;
                                                    db_cell_record.row_id =
                                                        db_row_record.row_id.clone();
                                                    // Get Col Id
                                                    db_cell_record.col_id =
                                                        ConverterUtil::get_column_int(
                                                            cell_attribute.get("r").ok_or(
                                                                anyhow!("Missing mandatory col id attribute"),
                                                            )?,
                                                        ).context("Failed to Convert col worksheet initialize")?;
                                                    if let Some(style_id) = cell_attribute.get("s")
                                                    {
                                                        db_cell_record.cell_style_id =
                                                            Some(style_id.parse().context(
                                                                "Failed to parse the col style id",
                                                            )?);
                                                    }
                                                    if let Some(cell_type) = cell_attribute.get("t")
                                                    {
                                                        db_cell_record.cell_type =
                                                            Some(CellDataType::get_enum(&cell_type))
                                                    };
                                                    if let Some(comment_id) =
                                                        cell_attribute.get("cm")
                                                    {
                                                        db_cell_record.cell_comment_id =
                                                            Some(comment_id.parse().context(
                                                                "Failed to parse the col comment id",
                                                            )?);
                                                    };
                                                    if let Some(value_meta_id) =
                                                        cell_attribute.get("vm")
                                                    {
                                                        db_cell_record.cell_metadata =
                                                            Some(value_meta_id.parse().context(
                                                                "Failed to parse the col value meta id",
                                                            )?);
                                                    };
                                                    if let Some(place_holder) =
                                                        cell_attribute.get("ph")
                                                    {
                                                        db_cell_record.cell_place_holder =
                                                            if place_holder == "1" {
                                                                Some(true)
                                                            } else {
                                                                None
                                                            };
                                                    };
                                                    loop {
                                                        if let Some((cell_child_id, _)) =
                                                            col_element.pop_child_mut()
                                                        {
                                                            if let Some(element) = xml_doc_mut
                                                                .pop_element_mut(&cell_child_id)
                                                            {
                                                                match element.get_tag() {
                                                                    "v" => {
                                                                        db_cell_record.cell_value =
                                                                            element
                                                                                .get_value()
                                                                                .clone();
                                                                    }
                                                                    "f" => {
                                                                        db_cell_record
                                                                            .cell_formula = element
                                                                            .get_value()
                                                                            .clone();
                                                                    }
                                                                    "is" => {
                                                                        if let Some((text_id, _)) =
                                                                            element.pop_child_mut()
                                                                        {
                                                                            if let Some(
                                                                                text_element,
                                                                            ) = xml_doc_mut
                                                                                .pop_element_mut(
                                                                                    &text_id,
                                                                                )
                                                                            {
                                                                                db_cell_record
                                                                                    .cell_value =
                                                                                    text_element
                                                                                        .get_value()
                                                                                        .clone();
                                                                            }
                                                                        }
                                                                    }
                                                                    _ => {
                                                                        return Err(anyhow!("Found un-know element cell child"));
                                                                    }
                                                                }
                                                            }
                                                        } else {
                                                            break;
                                                        }
                                                    }
                                                }
                                                // Insert record to the Database
                                                let insert_query = queries.get("insert_dynamic_sheet")
                                                    .ok_or(anyhow!("Failed to Get the insert Query at sheet data parser"))?;
                                                office_doc_mut
                                                    .get_connection()
                                                    .insert_record(insert_query, params![
                                                        db_cell_record.row_id,
                                                        db_cell_record.col_id,
                                                        db_cell_record.cell_style_id,
                                                        db_cell_record.cell_value,
                                                        db_cell_record.cell_formula,
                                                        if let Some(cell_type) = db_cell_record.cell_type{Some(CellDataType::get_string(cell_type))}else{None},
                                                        db_cell_record.cell_metadata,
                                                        db_cell_record.cell_place_holder,
                                                        db_cell_record.cell_comment_id
                                                    ],Some(table_name.clone()))
                                                    .context("Failed to insert row Data record into Sheet DB")?;
                                            } else {
                                                break;
                                            }
                                        }
                                    }
                                } else {
                                    break;
                                }
                            }
                        }
                    }
                }
                // unwrap dimension
                {
                    xml_doc_mut.pop_elements_by_tag_mut("dimension", None);
                }
            }
        }
        if column_collection.len() > 0 {
            column_collection.reverse();
            Ok(Some(column_collection))
        } else {
            Ok(None)
        }
    }
}

impl WorkSheet {
    fn get_sheet_file_name(
        sheet_name: Option<String>,
        office_document: &Weak<RefCell<OfficeDocument>>,
        sheet_collection: &Weak<RefCell<Vec<(String, String, bool, bool)>>>,
        workbook_relationship_part: &Weak<RefCell<RelationsPart>>,
    ) -> AnyResult<(String, String), AnyError> {
        let worksheet_content = EXCEL_TYPE_COLLECTION.get("worksheet").unwrap();
        if let Some(sheet_collection) = sheet_collection.upgrade() {
            if let Some(workbook_relationship_part) = workbook_relationship_part.upgrade() {
                if let Some(sheet_name) = sheet_name.clone() {
                    // If the Sheet name already exist get the path of sheet name
                    if let Some((_, rel_id, _, _)) = sheet_collection
                        .try_borrow()
                        .context("Failed to Get Sheet Collection")?
                        .iter()
                        .find(|item| item.0 == sheet_name)
                    {
                        return Ok((
                            workbook_relationship_part
                                .try_borrow()
                                .context("Failed to Get Workbook relationship")?
                                .get_target_by_id(&rel_id)
                                .context("Failed to Get Target Path")?
                                .ok_or(anyhow!("Failed to Get Relationship path"))?,
                            sheet_name,
                        ));
                    }
                }
                let mut sheet_count = sheet_collection
                    .try_borrow()
                    .context("Failed to pull Sheet Name Collection")?
                    .len()
                    + 1;
                let relative_path = workbook_relationship_part
                    .try_borrow_mut()
                    .context("Failed to pull relationship connection")?
                    .get_relative_path()
                    .context("Get Relative Path for Part File")?;
                if let Some(office_doc) = office_document.upgrade() {
                    let document = office_doc
                        .try_borrow()
                        .context("Failed to Borrow Document")?;
                    loop {
                        if document
                            .check_file_exist(format!(
                                "{}{}/{}{}.{}",
                                relative_path,
                                worksheet_content.default_path,
                                worksheet_content.default_name,
                                sheet_count,
                                worksheet_content.extension
                            ))
                            .context("Failed to Check the File Exist")?
                        {
                            sheet_count += 1;
                        } else {
                            break;
                        }
                    }
                }
                let file_path = format!("{}/{}", relative_path, worksheet_content.default_path);
                let sheet_name = format!(
                    "{}",
                    sheet_name.clone().unwrap_or(format!(
                        "{}{}",
                        worksheet_content.default_name, &sheet_count
                    ))
                );
                let relationship_id = workbook_relationship_part
                    .try_borrow_mut()
                    .context("Failed to Get Relationship Handle")?
                    .set_new_relationship_mut(
                        worksheet_content,
                        Some(file_path.clone()),
                        Some(format!(
                            "{}{}",
                            worksheet_content.default_name, &sheet_count
                        )),
                    )
                    .context("Setting New Calculation Chain Relationship Failed.")?;
                sheet_collection
                    .try_borrow_mut()
                    .context("Failed To pull Sheet Collection Handle")?
                    .push((sheet_name.clone(), relationship_id, false, false));
                return Ok((
                    format!(
                        "{}/{}{}.{}",
                        file_path,
                        worksheet_content.default_name,
                        &sheet_count,
                        worksheet_content.extension
                    ),
                    sheet_name,
                ));
            }
        }
        Err(anyhow!("Failed to upgrade relation part"))
    }
}

// ##################################### Feature Function ################################
impl WorkSheet {
    /// Set Active cell of the current sheet
    pub fn set_active_cell_mut(&mut self, cell_id: &str) {}
    /// Set Column property
    pub fn set_column_properties_mut(
        &mut self,
        column_id: &usize,
        column_properties: ColumnProperties,
    ) -> () {
        // Check if the target column has setting
        // Create new record
        // Update old record
    }
    pub fn set_row_properties_mut(
        &mut self,
        column_id: &usize,
        row_properties: RowProperties,
    ) -> () {
        // Check if the target column has setting
        // Create new record
        // Update old record
    }

    /// Set data for same row multiple columns along with row property
    pub fn set_row_value_mut(
        &mut self,
        cell_key: &str,
        column_cell: Vec<ColumnCell>,
    ) -> AnyResult<(), AnyError> {
        let (row_id, col_id) =
            ConverterUtil::get_cell_int(cell_key).context("Failed to extract cell key")?;
        self.set_row_value_id_mut(row_id, col_id, column_cell)
    }

    /// Set data for same row multiple columns along with row property
    pub fn set_row_value_id_mut(
        &mut self,
        row_id: usize,
        mut col_id: usize,
        mut column_cell: Vec<ColumnCell>,
    ) -> AnyResult<(), AnyError> {
        if let Some(office_document) = self.office_document.upgrade() {
            let office_doc_mut = office_document
                .try_borrow_mut()
                .context("Failed to get office doc handle")?;
            let insert_query = self
                .queries
                .get("insert_conflict_dynamic_sheet")
                .ok_or(anyhow!("Failed to Get Insert Query"))?;
            loop {
                if let Some(cell_data) = column_cell.pop() {
                    office_doc_mut
                        .get_connection()
                        .insert_record(
                            insert_query,
                            params![
                                row_id,
                                col_id,
                                cell_data.style_id,
                                cell_data.value,
                                cell_data.formula,
                                CellDataType::get_string(cell_data.data_type),
                                None::<String>,
                                None::<bool>,
                                None::<usize>
                            ],
                            Some(
                                self.file_path
                                    .rsplit("/")
                                    .next()
                                    .unwrap()
                                    .to_string()
                                    .replace(".xml", ""),
                            ),
                        )
                        .context("Failed to insert New record to the DB")?;
                    col_id += 1;
                } else {
                    break;
                }
            }
        }
        Ok(())
    }

    /// Add hyper link to current document
    pub fn add_hyperlink_mut(&mut self) {}

    /// Set Cell Range to merge
    pub fn set_merge_cell_mut(&mut self) {}

    /// List all Cell Range merged
    pub fn list_merge_cell_(&mut self) {}

    /// Remove merged cell range
    pub fn remove_merge_cell_mut(&mut self) {}

    /// Delete Current sheet and all its components
    pub fn delete_sheet_mut(self) -> AnyResult<(), AnyError> {
        if let Some(sheet_collection) = self.sheet_collection.upgrade() {
            sheet_collection
                .try_borrow_mut()
                .context("Failed to pull Sheets Collection")?
                .retain(|item| item.0 != self.sheet_name);
        }
        if let Some(workbook_relationship_part) = self.workbook_relationship_part.upgrade() {
            workbook_relationship_part
                .try_borrow_mut()
                .context("Failed to pull workbook relationship handle")?
                .delete_relationship_mut(&self.file_path);
        }
        if let Some(xml_tree) = self.office_document.upgrade() {
            xml_tree
                .try_borrow_mut()
                .context("Failed to Pull XML Handle")?
                .delete_document_mut(&self.file_path)
                .context("Failed to delete the document from database")?;
        }
        self.flush().context("Failed to flush the worksheet")?;
        Ok(())
    }
}
