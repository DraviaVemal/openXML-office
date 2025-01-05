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
            StyleId,
        },
        services::CommonServices,
    },
};
use anyhow::{anyhow, Context, Error as AnyError, Result as AnyResult};
use rusqlite::{params, Row, ToSql};
use std::{
    cell::RefCell,
    collections::{HashMap, VecDeque},
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
    column_collection: Option<VecDeque<ColumnProperties>>,
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
                self.serialize_cols(&mut xml_doc_mut)?;
                // Add Sheet Data to Document
                self.serialize_row_cell(&office_doc_mut, &mut xml_doc_mut)?;
                // Add dimension
                self.serialize_dimension(&office_doc_mut, &mut xml_doc_mut)?;
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
            office_doc_mut.get_database().drop_table(
                self.file_path
                    .rsplit("/")
                    .next()
                    .unwrap()
                    .to_string()
                    .replace(".xml", ""),
            )?;
            office_doc_mut.get_database().drop_table(format!(
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
    ) -> AnyResult<Option<VecDeque<ColumnProperties>>, AnyError> {
        let mut column_collection = VecDeque::new();
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
                    .get_database()
                    .create_table(&create_query, Some(table_name.clone()))
                    .context("Failed to create sheet data table")?;
                office_doc_mut
                    .get_database()
                    .create_table(&create_query_row, Some(format!("{}_row", table_name)))
                    .context("Failed to create sheet data row table")?;
                // unwrap columns to local collection
                deserialize_cols(&mut column_collection, &mut xml_doc_mut)
                    .context("Failed To Deserialize Cols")?;
                // unwrap sheet data into database
                deserialize_sheet_data(table_name, queries, &mut xml_doc_mut, office_doc_mut)?;
                // unwrap dimension
                xml_doc_mut.pop_elements_by_tag_mut("dimension", None);
            }
        }
        Ok(Some(column_collection))
    }

    fn serialize_cols(
        &mut self,
        xml_doc_mut: &mut std::cell::RefMut<'_, XmlDocument>,
    ) -> Result<(), AnyError> {
        Ok(
            if let Some(mut column_collection) = self.column_collection.take() {
                if column_collection.len() > 0 {
                    let cols_id = xml_doc_mut
                        .insert_children_after_tag_mut("cols", "sheetViews", None)
                        .context("Failed to Insert Cols Element")?
                        .get_id();
                    loop {
                        if let Some(item) = column_collection.pop_front() {
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
                                attribute.insert("style".to_string(), style_id.id.to_string());
                            }
                            if let Some(_) = item.hidden {
                                attribute.insert("hidden".to_string(), "1".to_string());
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
            },
        )
    }

    fn serialize_row_cell(
        &mut self,
        office_doc_mut: &std::cell::RefMut<'_, OfficeDocument>,
        xml_doc_mut: &mut std::cell::RefMut<'_, XmlDocument>,
    ) -> Result<(), AnyError> {
        let select_all_query = self
            .queries
            .get("select_all_dynamic_sheet")
            .ok_or(anyhow!("Failed to Get Select All Query"))?;
        fn row_mapper(row: &Row) -> AnyResult<(RowRecord, CellRecord), rusqlite::Error> {
            Ok((
                RowRecord {
                    index: row.get(0)?,
                    hide: row.get(1)?,
                    span: row.get(2)?,
                    height: row.get(3)?,
                    style_id: row.get(4)?,
                    thick_top: row.get(5)?,
                    thick_bottom: row.get(6)?,
                    group_level: row.get(7)?,
                    collapsed: row.get(8)?,
                    place_holder: row.get(9)?,
                },
                CellRecord {
                    row_index: row.get(0)?,
                    col_index: row.get(10)?,
                    style_id: row.get(11)?,
                    value: row.get(12)?,
                    formula: row.get(13)?,
                    data_type: if let Some(value) = row.get(14)? {
                        let type_text: String = value;
                        Some(CellDataType::get_enum(&type_text))
                    } else {
                        None
                    },
                    metadata: row.get(15)?,
                    place_holder: row.get(16)?,
                    comment_id: row.get(17)?,
                },
            ))
        }
        let mut sheet_data = office_doc_mut
            .get_database()
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
        Ok(if sheet_data.len() > 0 {
            let mut row_element_id = 0;
            let mut row_index = 0;
            loop {
                if let Some((db_row, db_cell)) = sheet_data.pop() {
                    // Create row element
                    if row_index != db_row.index {
                        let row_element = xml_doc_mut
                            .append_child_mut("row", Some(&sheet_data_id))
                            .context("Failed to insert row element")?;
                        row_element_id = row_element.get_id();
                        row_index = db_row.index;
                        let mut row_attribute = HashMap::new();
                        row_attribute.insert("r".to_string(), db_row.index.to_string());
                        if let Some(row_span) = db_row.span {
                            row_attribute.insert("spans".to_string(), row_span);
                        }
                        if let Some(row_style_id) = db_row.style_id {
                            row_attribute.insert("customFormat".to_string(), "1".to_string());
                            row_attribute.insert("s".to_string(), row_style_id.id.to_string());
                        }
                        if let Some(row_height) = db_row.height {
                            row_attribute.insert("customHeight".to_string(), "1".to_string());
                            row_attribute.insert("ht".to_string(), row_height.to_string());
                        }
                        if let Some(_) = db_row.hide {
                            row_attribute.insert("hidden".to_string(), "1".to_string());
                        }
                        if let Some(row_group_level) = db_row.group_level {
                            row_attribute
                                .insert("outlineLevel".to_string(), row_group_level.to_string());
                        }
                        if let Some(_) = db_row.collapsed {
                            row_attribute.insert("collapsed".to_string(), "1".to_string());
                        }
                        if let Some(_) = db_row.thick_top {
                            row_attribute.insert("thickTop".to_string(), "1".to_string());
                        }
                        if let Some(_) = db_row.thick_bottom {
                            row_attribute.insert("thickBot".to_string(), "1".to_string());
                        }
                        if let Some(_) = db_row.place_holder {
                            row_attribute.insert("ph".to_string(), "1".to_string());
                        }
                        row_element
                            .set_attribute_mut(row_attribute)
                            .context("Failed to set attribute for row")?;
                    }
                    // Create cell element
                    if let Some(col_index) = db_cell.col_index {
                        let cell_id;
                        let cell_element = xml_doc_mut
                            .append_child_mut("c", Some(&row_element_id))
                            .context("Failed to insert row element")?;
                        cell_id = cell_element.get_id();
                        let mut cell_attribute = HashMap::new();
                        cell_attribute.insert(
                            "r".to_string(),
                            format!(
                                "{}{}",
                                ConverterUtil::get_column_ref(col_index)
                                    .context("Failed to get Char Id from Int")?,
                                db_row.index
                            ),
                        );
                        if let Some(cell_style_id) = db_cell.style_id {
                            cell_attribute.insert("s".to_string(), cell_style_id.id.to_string());
                        }
                        if let Some(cell_type) = db_cell.data_type {
                            cell_attribute
                                .insert("t".to_string(), CellDataType::get_string(cell_type));
                        }
                        if let Some(cell_comment_id) = db_cell.comment_id {
                            cell_attribute.insert("cm".to_string(), cell_comment_id.to_string());
                        }
                        if let Some(cell_metadata) = db_cell.metadata {
                            cell_attribute.insert("vm".to_string(), cell_metadata.to_string());
                        }
                        if let Some(_) = db_cell.place_holder {
                            cell_attribute.insert("ph".to_string(), "1".to_string());
                        }
                        cell_element
                            .set_attribute_mut(cell_attribute)
                            .context("Failed to Set Attribute for cell")?;
                        // Create cell's child element
                        match db_cell.data_type.unwrap_or(CellDataType::Number) {
                            CellDataType::InlineString => {
                                let inline_string_id = xml_doc_mut
                                    .append_child_mut("is", Some(&cell_id))
                                    .context("Failed to insert Inline string element")?
                                    .get_id();
                                let text_element = xml_doc_mut
                                    .append_child_mut("t", Some(&inline_string_id))
                                    .context("Failed To insert Text Value to inline string")?;
                                text_element.set_value_mut(if let Some(value) = db_cell.value {
                                    value
                                } else {
                                    "".to_string()
                                });
                            }
                            _ => {
                                if let Some(formula) = db_cell.formula {
                                    let formula_element = xml_doc_mut
                                        .append_child_mut("f", Some(&cell_id))
                                        .context("Failed to insert Inline string element")?;
                                    formula_element.set_value_mut(formula);
                                }
                                let value_element = xml_doc_mut
                                    .append_child_mut("v", Some(&cell_id))
                                    .context("Failed to insert Inline string element")?;
                                value_element.set_value_mut(if let Some(value) = db_cell.value {
                                    value
                                } else {
                                    "".to_string()
                                });
                            }
                        }
                    }
                } else {
                    break;
                }
            }
        })
    }

    fn serialize_dimension(
        &mut self,
        office_doc_mut: &std::cell::RefMut<'_, OfficeDocument>,
        xml_doc_mut: &mut std::cell::RefMut<'_, XmlDocument>,
    ) -> Result<(), AnyError> {
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
        ) -> Result<(Option<usize>, Option<usize>, Option<usize>, Option<usize>), rusqlite::Error>
        {
            Ok((row.get(0)?, row.get(1)?, row.get(2)?, row.get(3)?))
        }
        let mut dim_range = office_doc_mut
            .get_database()
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
        Ok(
            if let Some((start_row_id, start_col_id, end_row_id, end_col_id)) = dim_range.pop() {
                if start_row_id.is_some() && start_col_id.is_some() {
                    let dim_element = xml_doc_mut
                        .append_child_mut("dimension", Some(&root_id))
                        .context("Failed to Add Dimension node to worksheet")?;
                    let mut dimension_attribute = HashMap::new();
                    dimension_attribute.insert(
                        "ref".to_string(),
                        format!(
                            "{}{}:{}{}",
                            ConverterUtil::get_column_ref(start_col_id.unwrap_or(1))
                                .context("Failed to convert dim col start")?,
                            start_row_id.unwrap_or(1),
                            ConverterUtil::get_column_ref(end_col_id.unwrap_or(1))
                                .context("Failed to convert dim col end")?,
                            end_row_id.unwrap_or(1)
                        ),
                    );
                    dim_element
                        .set_attribute_mut(dimension_attribute)
                        .context("Failed to set attribute value to dimension")?;
                }
            },
        )
    }
}

fn deserialize_sheet_data(
    table_name: String,
    queries: &HashMap<String, String>,
    xml_doc_mut: &mut std::cell::RefMut<'_, XmlDocument>,
    office_doc_mut: std::cell::RefMut<'_, OfficeDocument>,
) -> Result<(), AnyError> {
    Ok(
        if let Some(mut sheet_data_element) = xml_doc_mut.pop_elements_by_tag_mut("sheetData", None)
        {
            if let Some(sheet_data) = sheet_data_element.pop() {
                // Loop All rows of sheet data
                loop {
                    if let Some((row_index, _)) = sheet_data.pop_child_mut() {
                        if let Some(row_element) = xml_doc_mut.pop_element_mut(&row_index) {
                            let mut db_row_record = RowRecord::default();
                            let row_attribute = row_element
                                .get_attribute()
                                .ok_or(anyhow!("Failed to pull Row Attribute."))?;
                            // Get Row Id
                            db_row_record.index = row_attribute
                                .get("r")
                                .ok_or(anyhow!("Missing mandatory row id attribute"))?
                                .parse()
                                .context("Failed to parse row id")?;
                            if let Some(row_span) = row_attribute.get("spans") {
                                db_row_record.span = Some(row_span.to_string());
                            }
                            if let Some(style_id) = row_attribute.get("s") {
                                if let Some(custom_formant) = row_attribute.get("customFormat") {
                                    db_row_record.style_id = if custom_formant == "1" {
                                        Some(StyleId::new(
                                            style_id
                                                .parse()
                                                .context("Failed to parse the row style id")?,
                                        ))
                                    } else {
                                        None
                                    };
                                }
                            }
                            if let Some(hidden) = row_attribute.get("hidden") {
                                db_row_record.hide = if hidden == "1" { Some(true) } else { None };
                            }
                            if let Some(height) = row_attribute.get("ht") {
                                if let Some(custom_height) = row_attribute.get("customHeight") {
                                    db_row_record.height = if custom_height == "1" {
                                        Some(
                                            height
                                                .parse()
                                                .context("Failed to parse the row height")?,
                                        )
                                    } else {
                                        None
                                    };
                                }
                            }
                            if let Some(row_group_level) = row_attribute.get("outlineLevel") {
                                let outline_level = row_group_level
                                    .parse()
                                    .context("Failed to parse the row group level")?;
                                db_row_record.group_level = if outline_level > 0 {
                                    Some(outline_level)
                                } else {
                                    None
                                };
                            }
                            if let Some(collapsed) = row_attribute.get("collapsed") {
                                db_row_record.collapsed =
                                    if collapsed == "1" { Some(true) } else { None };
                            }
                            if let Some(thick_top) = row_attribute.get("thickTop") {
                                db_row_record.thick_top =
                                    if thick_top == "1" { Some(true) } else { None };
                            }
                            if let Some(thick_bottom) = row_attribute.get("thickBot") {
                                db_row_record.thick_bottom = if thick_bottom == "1" {
                                    Some(true)
                                } else {
                                    None
                                };
                            }
                            if let Some(place_holder) = row_attribute.get("ph") {
                                db_row_record.place_holder = if place_holder == "1" {
                                    Some(true)
                                } else {
                                    None
                                };
                            }
                            // Insert record to the Database
                            let insert_query_row = queries.get("insert_dynamic_sheet_row").ok_or(
                                anyhow!("Failed to Get the insert Row Query at sheet data parser"),
                            )?;
                            office_doc_mut
                                .get_database()
                                .insert_record(
                                    insert_query_row,
                                    params![
                                        db_row_record.index,
                                        db_row_record.hide,
                                        db_row_record.span,
                                        db_row_record.height,
                                        db_row_record.style_id,
                                        db_row_record.thick_top,
                                        db_row_record.thick_bottom,
                                        db_row_record.group_level,
                                        db_row_record.collapsed,
                                        db_row_record.place_holder,
                                    ],
                                    Some(format!("{}_row", table_name)),
                                )
                                .context("Failed to insert Cell Data record into Sheet DB")?;
                            // Loop All Columns of row
                            loop {
                                let mut db_cell_record = CellRecord::default();
                                if let Some((col_index, _)) = row_element.pop_child_mut() {
                                    if let Some(col_element) =
                                        xml_doc_mut.pop_element_mut(&col_index)
                                    {
                                        let cell_attribute = col_element.get_attribute().ok_or(
                                            anyhow!("Failed to pull attribute for Column"),
                                        )?;
                                        db_cell_record.row_index = db_row_record.index.clone();
                                        // Get Col Id
                                        db_cell_record.col_index = Some(
                                            ConverterUtil::get_column_index(
                                                cell_attribute.get("r").ok_or(anyhow!(
                                                    "Missing mandatory col id attribute"
                                                ))?,
                                            )
                                            .context(
                                                "Failed to Convert col worksheet initialize",
                                            )?,
                                        );
                                        if let Some(style_id) = cell_attribute.get("s") {
                                            db_cell_record.style_id =
                                                Some(StyleId::new(style_id.parse().context(
                                                    "Failed to parse the col style id",
                                                )?));
                                        }
                                        if let Some(cell_type) = cell_attribute.get("t") {
                                            db_cell_record.data_type =
                                                Some(CellDataType::get_enum(&cell_type))
                                        };
                                        if let Some(comment_id) = cell_attribute.get("cm") {
                                            db_cell_record.comment_id =
                                                Some(comment_id.parse().context(
                                                    "Failed to parse the col comment id",
                                                )?);
                                        };
                                        if let Some(value_meta_id) = cell_attribute.get("vm") {
                                            db_cell_record.metadata =
                                                Some(value_meta_id.parse().context(
                                                    "Failed to parse the col value meta id",
                                                )?);
                                        };
                                        if let Some(place_holder) = cell_attribute.get("ph") {
                                            db_cell_record.place_holder = if place_holder == "1" {
                                                Some(true)
                                            } else {
                                                None
                                            };
                                        };
                                        loop {
                                            if let Some((cell_child_id, _)) =
                                                col_element.pop_child_mut()
                                            {
                                                if let Some(element) =
                                                    xml_doc_mut.pop_element_mut(&cell_child_id)
                                                {
                                                    match element.get_tag() {
                                                        "v" => {
                                                            db_cell_record.value =
                                                                element.get_value().clone();
                                                        }
                                                        "f" => {
                                                            db_cell_record.formula =
                                                                element.get_value().clone();
                                                        }
                                                        "is" => {
                                                            if let Some((text_id, _)) =
                                                                element.pop_child_mut()
                                                            {
                                                                if let Some(text_element) =
                                                                    xml_doc_mut
                                                                        .pop_element_mut(&text_id)
                                                                {
                                                                    db_cell_record.value =
                                                                        text_element
                                                                            .get_value()
                                                                            .clone();
                                                                }
                                                            }
                                                        }
                                                        _ => {
                                                            return Err(anyhow!(
                                                                "Found un-know element cell child"
                                                            ));
                                                        }
                                                    }
                                                }
                                            } else {
                                                break;
                                            }
                                        }
                                    }
                                    // Insert record to the Database
                                    let insert_query =
                                        queries.get("insert_dynamic_sheet").ok_or(anyhow!(
                                            "Failed to Get the insert Query at sheet data parser"
                                        ))?;
                                    office_doc_mut
                                        .get_database()
                                        .insert_record(
                                            insert_query,
                                            params![
                                                db_cell_record.row_index,
                                                db_cell_record.col_index,
                                                db_cell_record.style_id,
                                                db_cell_record.value,
                                                db_cell_record.formula,
                                                if let Some(cell_type) = db_cell_record.data_type {
                                                    Some(CellDataType::get_string(cell_type))
                                                } else {
                                                    None
                                                },
                                                db_cell_record.metadata,
                                                db_cell_record.place_holder,
                                                db_cell_record.comment_id
                                            ],
                                            Some(table_name.clone()),
                                        )
                                        .context(
                                            "Failed to insert row Data record into Sheet DB",
                                        )?;
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
        },
    )
}

fn deserialize_cols(
    column_collection: &mut VecDeque<ColumnProperties>,
    xml_doc_mut: &mut std::cell::RefMut<'_, XmlDocument>,
) -> Result<(), AnyError> {
    Ok(
        if let Some(mut cols_element) = xml_doc_mut.pop_elements_by_tag_mut("cols", None) {
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
                                    if best_fit == "1" { Some(true) } else { None }
                            }
                            if let Some(hidden) = attributes.get("hidden") {
                                column_properties.hidden =
                                    if hidden == "1" { Some(true) } else { None }
                            }
                            if let Some(style) = attributes.get("style") {
                                column_properties.style_id = Some(StyleId::new(
                                    style.parse().context("Failed to parse style ID")?,
                                ));
                            }
                            if let Some(outline_level) = attributes.get("outlineLevel") {
                                column_properties.group_level =
                                    outline_level.parse().context("Failed to parse style ID")?;
                            }
                            if let Some(custom_width) = attributes.get("customWidth") {
                                if custom_width == "1" {
                                    column_properties.width = Some(
                                        attributes
                                            .get("width")
                                            .ok_or(anyhow!("Failed to get custom width"))?
                                            .parse()
                                            .context("Failed to parse custom width")?,
                                    );
                                }
                            }
                            if let Some(collapsed) = attributes.get("collapsed") {
                                column_properties.collapsed =
                                    if collapsed == "1" { Some(true) } else { None }
                            }
                            column_collection.push_back(column_properties);
                        }
                    } else {
                        break;
                    }
                }
            }
        },
    )
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
                let file_path = format!("{}{}", relative_path, worksheet_content.default_path);
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
    pub fn set_active_cell_mut(&mut self, cell_ref: &str, selected_range: Vec<&str>) {}

    /// Set Column property
    pub fn set_column_ref_properties_mut(
        &mut self,
        cell_ref: &str,
        column_properties: Option<ColumnProperties>,
    ) -> AnyResult<(), AnyError> {
        let col_index = ConverterUtil::get_column_index(cell_ref)
            .context("Failed to Get Column index from reference")?;
        self.set_column_index_properties_mut(&col_index, column_properties)
    }

    /// Set Column property
    pub fn set_column_index_properties_mut(
        &mut self,
        col_index: &usize,
        column_properties: Option<ColumnProperties>,
    ) -> AnyResult<(), AnyError> {
        if let Some(column_collection) = self.column_collection.as_mut() {
            let mut new_ranges = VecDeque::new();
            // Delete Old Record
            column_collection.retain_mut(|range| {
                if range.min == *col_index && range.max == *col_index {
                    // Fully matched range, remove it
                    return false;
                } else if range.min <= *col_index && *col_index <= range.max {
                    // Value lies within the range
                    if range.min == *col_index {
                        // Trim the start
                        range.min = col_index + 1;
                    } else if range.max == *col_index {
                        // Trim the end
                        range.max = col_index - 1;
                    } else {
                        // Split the range
                        new_ranges.push_back(ColumnProperties {
                            min: col_index + 1,
                            max: range.max,
                            ..ColumnProperties::default()
                        });
                        range.max = col_index - 1;
                    }
                }
                true
            });
            if let Some(column_properties) = column_properties {
                column_collection.push_back(ColumnProperties {
                    max: *col_index,
                    min: *col_index,
                    ..column_properties
                });
            }
            column_collection.append(&mut new_ranges);
        }
        Ok(())
    }

    /// Set/Reset Row property
    pub fn set_row_index_properties_mut(
        &mut self,
        row_index: &usize,
        row_properties: Option<RowProperties>,
    ) -> AnyResult<(), AnyError> {
        if let Some(office_document) = self.office_document.upgrade() {
            let office_doc = office_document
                .try_borrow_mut()
                .context("Failed To pull Office Handle WorkSheet row property")?;
            let query = self
                .queries
                .get("insert_conflict_dynamic_sheet_row")
                .ok_or(anyhow!("Failed to Get Expected Query"))?;
            let row_prop = if let Some(row_prop) = row_properties {
                row_prop
            } else {
                RowProperties::default()
            };
            office_doc
                .get_database()
                .insert_record(
                    query,
                    params![
                        row_index,
                        row_prop.hidden,
                        row_prop.span,
                        row_prop.height,
                        row_prop.style_id,
                        row_prop.thick_top,
                        row_prop.thick_bottom,
                        row_prop.group_level,
                        row_prop.collapsed,
                        row_prop.place_holder
                    ],
                    Some(format!(
                        "{}_row",
                        self.file_path
                            .rsplit("/")
                            .next()
                            .unwrap()
                            .to_string()
                            .replace(".xml", "")
                    )),
                )
                .context("Failed to update Row properties")?;

            Ok(())
        } else {
            Err(anyhow!("Failed to Get Office Handle"))
        }
    }

    /// Set data for same row multiple columns along with row property
    pub fn set_row_value_ref_mut(
        &mut self,
        cell_ref: &str,
        column_cell: Vec<ColumnCell>,
    ) -> AnyResult<(), AnyError> {
        let (row_index, col_index) =
            ConverterUtil::get_cell_index(cell_ref).context("Failed to extract cell key")?;
        self.set_row_value_index_mut(row_index, col_index, column_cell)
    }

    /// Set data for same row multiple columns along with row property
    pub fn set_row_value_index_mut(
        &mut self,
        row_index: usize,
        mut col_index: usize,
        mut column_cell: Vec<ColumnCell>,
    ) -> AnyResult<(), AnyError> {
        if let Some(office_document) = self.office_document.upgrade() {
            for cell_data in column_cell.iter_mut() {
                if let Some(cell_value) = cell_data.value.as_ref() {
                    match cell_data.data_type {
                        CellDataType::Auto => {
                            if cell_value.parse::<f64>().is_ok() {
                                cell_data.data_type = CellDataType::Number;
                            } else if cell_value.parse::<bool>().is_ok() {
                                cell_data.data_type = CellDataType::Boolean;
                                if cell_value.parse::<bool>().context("Parse Fail")? {
                                    cell_data.value = Some("1".to_string());
                                } else {
                                    cell_data.value = Some("0".to_string());
                                }
                            } else {
                                cell_data.data_type = CellDataType::ShareString;
                                cell_data.value = Some(self.update_share_string(cell_value)?);
                            }
                        }
                        CellDataType::ShareString => {
                            cell_data.value = Some(self.update_share_string(cell_value)?);
                        }
                        CellDataType::Boolean => {
                            cell_data.value = match cell_value.to_lowercase().as_str() {
                                "false" | "0" | "" => Some("0".to_string()),
                                _ => Some("1".to_string()),
                            }
                        }
                        _ => {}
                    }
                } else {
                    cell_data.data_type = CellDataType::Number;
                }
                cell_data.row_index = row_index;
                cell_data.col_index = col_index;
                col_index += 1;
            }
            let insert_query = self
                .queries
                .get("insert_conflict_dynamic_sheet")
                .ok_or(anyhow!("Failed to Get Insert Query"))?
                .to_owned();
            let office_doc_mut = office_document
                .try_borrow_mut()
                .context("Failed to get office doc handle")?;
            fn row_parser(cell_data: ColumnCell) -> Vec<Box<dyn ToSql>> {
                vec![
                    Box::new(cell_data.row_index),
                    Box::new(cell_data.col_index),
                    Box::new(cell_data.style_id),
                    Box::new(cell_data.value),
                    Box::new(cell_data.formula),
                    Box::new(CellDataType::get_string(cell_data.data_type)),
                    Box::new(cell_data.metadata),
                    Box::new(cell_data.place_holder),
                    Box::new(cell_data.comment_id),
                ]
            }
            office_doc_mut
                .get_database()
                .insert_records(
                    &insert_query,
                    column_cell,
                    row_parser,
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
        }
        Ok(())
    }

    fn update_share_string(&mut self, cell_value: &String) -> AnyResult<String, AnyError> {
        if let Some(common_service) = self.common_service.upgrade() {
            common_service
                .try_borrow_mut()
                .context("Failed to Get Share String Handle")?
                .get_string_id(cell_value.to_owned())
                .context("Failed to get share string id")
        } else {
            Err(anyhow!("Failed to update Share String Record"))
        }
    }

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
