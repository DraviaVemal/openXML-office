use crate::{
    converters::ConverterUtil,
    element_dictionary::EXCEL_TYPE_COLLECTION,
    files::{OfficeDocument, XmlDocument, XmlSerializer},
    global_2007::{
        parts::RelationsPart,
        traits::{Enum, XmlDocumentPartCommon},
    },
    log_elapsed,
    order_dictionary::EXCEL_ORDER_COLLECTION,
    spreadsheet_2007::{
        models::{CellDataType, CellProperties, ColumnProperties, RowProperties, StyleId},
        services::CommonServices,
    },
};
use anyhow::{anyhow, Context, Error as AnyError, Result as AnyResult};
use std::{
    cell::RefCell,
    cmp::{max, min},
    collections::{BTreeMap, HashMap, VecDeque},
    rc::{Rc, Weak},
};

#[derive(Debug)]
pub(crate) struct RowData {
    row_record: RowProperties,
    cell_records: Option<BTreeMap<u16, CellProperties>>,
}

#[derive(Debug)]
pub(crate) struct Dimension {
    start_col: u16,
    end_col: u16,
}

impl Default for Dimension {
    fn default() -> Self {
        Self {
            start_col: 16384,
            end_col: 1,
        }
    }
}

#[derive(Debug)]
pub(crate) struct WorkSheetView {
    tab_color: Option<i16>,
    default_grid_color: Option<bool>,
    view_right_to_left: Option<bool>,
    show_formula_bar: Option<bool>,
    show_grid_line: Option<bool>,
    show_outline_symbol: Option<bool>,
    show_row_col_header: Option<bool>,
    show_ruler: Option<bool>,
    show_white_space: Option<bool>,
    show_zero: Option<bool>,
    tab_selected: Option<bool>,
    top_left_cell: Option<String>,
    view: Option<String>,
    window_protection: Option<bool>,
    zoom_scale: Option<i16>,
    zoom_scale_normal: Option<i16>,
    zoom_scale_page_layout: Option<i16>,
    zoom_scale_sheet_layout: Option<i16>,
}

#[derive(Debug)]
pub struct WorkSheet {
    office_document: Weak<RefCell<OfficeDocument>>,
    xml_document: Weak<RefCell<XmlDocument>>,
    common_service: Weak<RefCell<CommonServices>>,
    workbook_relationship_part: Weak<RefCell<RelationsPart>>,
    sheet_collection: Weak<RefCell<Vec<(String, String, bool, bool)>>>,
    sheet_relationship_part: Rc<RefCell<RelationsPart>>,
    dimension: Dimension,
    // sheet_property: Option<_>,
    sheet_view: Option<WorkSheetView>,
    // sheet_format_property: Option<_>,
    column_collection: Option<VecDeque<ColumnProperties>>,
    sheet_data: Option<BTreeMap<u32, RowData>>,
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
    fn initialize_content_xml() -> AnyResult<(XmlDocument, Option<String>, String, String), AnyError>
    {
        let content = EXCEL_TYPE_COLLECTION.get("worksheet").unwrap();
        let template_core_properties = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <sheetData />
</worksheet>"#;
        Ok((
            XmlSerializer::vec_to_xml_doc_tree(
                template_core_properties.as_bytes().to_vec(),
                "Default Worksheet",
            )
            .context("Initializing Worksheet Failed")?,
            Some(content.content_type.to_string()),
            content.extension.to_string(),
            content.extension_type.to_string(),
        ))
    }
    fn close_document(&mut self) -> AnyResult<(), AnyError>
    where
        Self: Sized,
    {
        log_elapsed!(
            || {
                if let Some(office_document) = self.office_document.upgrade() {
                    let mut office_doc_mut = office_document
                        .try_borrow_mut()
                        .context("Failed to pull office document")?;
                    if let Some(xml_document) = self.xml_document.upgrade() {
                        let mut xml_doc_mut = xml_document
                            .try_borrow_mut()
                            .context("Failed to Pull XML Handle")?;
                        // Add dimension
                        log_elapsed!(self.serialize_dimension(&mut xml_doc_mut))?;
                        // Add Cols Record to Document
                        log_elapsed!(self.serialize_cols(&mut xml_doc_mut))?;
                        // Add Sheet Data to Document
                        log_elapsed!(self.serialize_sheet_data(&mut xml_doc_mut))?;
                        if let Some(root_element) = xml_doc_mut.get_root_mut() {
                            log_elapsed!(root_element
                                .order_child_mut(
                                    EXCEL_ORDER_COLLECTION
                                        .get("worksheet")
                                        .ok_or(anyhow!("Failed to get worksheet default order"),)?,
                                )
                                .context("Failed Reorder the element child's"))?;
                        }
                    }
                    log_elapsed!(
                        || {
                            office_doc_mut
                                .close_xml_document(&self.file_path)
                                .context("Failed to close the current tree document")
                        },
                        "Close worksheet document"
                    )?;
                }
                Ok(())
            },
            "Close Worksheet"
        )
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
        let (column_collection, sheet_data, sheet_view, dimension) = log_elapsed!(
            || { Self::initialize_worksheet(&xml_document).context("Failed to open Worksheet") },
            "Worksheet Initialize Time"
        )?;
        Ok(Self {
            office_document,
            xml_document,
            common_service,
            workbook_relationship_part,
            sheet_relationship_part,
            dimension,
            sheet_view,
            sheet_collection,
            column_collection,
            sheet_data,
            file_path: file_path.to_string(),
            sheet_name,
        })
    }

    fn initialize_worksheet(
        xml_document: &Weak<RefCell<XmlDocument>>,
    ) -> AnyResult<
        (
            Option<VecDeque<ColumnProperties>>,
            Option<BTreeMap<u32, RowData>>,
            Option<WorkSheetView>,
            Dimension,
        ),
        AnyError,
    > {
        if let Some(xml_document) = xml_document.upgrade() {
            let mut xml_doc_mut = xml_document
                .try_borrow_mut()
                .context("Failed to get XML doc handle")?;
            // unwrap dimension
            xml_doc_mut.pop_elements_by_tag_mut("dimension", None);
            // unwrap columns to local collection
            let column_collection = log_elapsed!(
                || { deserialize_cols(&mut xml_doc_mut).context("Failed To Deserialize Cols") },
                "Column deserialize"
            )?;
            // unwrap sheet data into database
            let (sheet_data, dimension) = log_elapsed!(
                || {
                    deserialize_sheet_data(&mut xml_doc_mut)
                        .context("Failed To Deserialize Sheet Data")
                },
                "Sheet Data Deserialize"
            )?;
            let worksheet_view = log_elapsed!(
                || {
                    deserialize_worksheet_view(&mut xml_doc_mut)
                        .context("Failed to deserialize Worksheet View")
                },
                "Worksheet View Deserialization"
            )?;
            Ok((column_collection, sheet_data, worksheet_view, dimension))
        } else {
            Ok((None, None, None, Dimension::default()))
        }
    }

    fn serialize_dimension(&mut self, xml_doc_mut: &mut XmlDocument) -> Result<(), AnyError> {
        fn set_default(xml_doc_mut: &mut XmlDocument) -> AnyResult<(), AnyError> {
            let mut dimension_attribute = HashMap::new();
            dimension_attribute.insert("ref".to_string(), "A1".to_string());
            xml_doc_mut
                .append_child_mut("dimension", None)
                .context("Failed to Add Dimension node to worksheet")?
                .set_attribute_mut(dimension_attribute)
                .context("Failed to set attribute value to dimension")?;
            Ok(())
        }
        if let Some(sheet_data) = self.sheet_data.as_ref() {
            if let Some(first_item) = sheet_data.first_key_value() {
                let mut dimension_attribute = HashMap::new();
                dimension_attribute.insert(
                    "ref".to_string(),
                    format!(
                        "{}{}:{}{}",
                        ConverterUtil::get_column_ref(self.dimension.start_col)
                            .context("Failed to convert dim col start")?,
                        first_item.0,
                        ConverterUtil::get_column_ref(self.dimension.end_col)
                            .context("Failed to convert dim col end")?,
                        if let Some(row_end) = sheet_data.last_key_value() {
                            row_end.0
                        } else {
                            first_item.0
                        }
                    ),
                );
                xml_doc_mut
                    .append_child_mut("dimension", None)
                    .context("Failed to Add Dimension node to worksheet")?
                    .set_attribute_mut(dimension_attribute)
                    .context("Failed to set attribute value to dimension")?;
            } else {
                set_default(xml_doc_mut)?;
            }
        } else {
            set_default(xml_doc_mut)?;
        }
        Ok(())
    }

    fn serialize_cols(&mut self, xml_doc_mut: &mut XmlDocument) -> AnyResult<(), AnyError> {
        Ok(
            if let Some(mut column_collection) = self.column_collection.take() {
                if column_collection.len() > 0 {
                    let cols_id = xml_doc_mut
                        .insert_children_after_tag_mut("cols", "sheetViews", None)
                        .context("Failed to Insert Cols Element")?
                        .get_id();
                    loop {
                        if let Some(item) = column_collection.pop_front() {
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
                            xml_doc_mut
                                .append_child_mut("col", Some(&cols_id))
                                .context("Failed to insert col record")?
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

    fn serialize_sheet_data(&mut self, xml_doc_mut: &mut XmlDocument) -> AnyResult<(), AnyError> {
        if let Some(sheet_data) = self.sheet_data.take() {
            let sheet_data_id = xml_doc_mut
                .insert_children_after_tag_mut("sheetData", "cols", None)
                .context("Failed to Insert Cols Element")?
                .get_id();
            for (row_index, db_row) in sheet_data {
                let row_element = xml_doc_mut
                    .append_child_mut("row", Some(&sheet_data_id))
                    .context("Failed to insert row element")?;
                let row_element_id = row_element.get_id();
                let mut row_attribute = HashMap::new();
                row_attribute.insert("r".to_string(), row_index.to_string());
                if let Some(row_span) = db_row.row_record.span {
                    row_attribute.insert("spans".to_string(), row_span);
                }
                if let Some(row_style_id) = db_row.row_record.style_id {
                    row_attribute.insert("customFormat".to_string(), "1".to_string());
                    row_attribute.insert("s".to_string(), row_style_id.id.to_string());
                }
                if let Some(row_height) = db_row.row_record.height {
                    row_attribute.insert("customHeight".to_string(), "1".to_string());
                    row_attribute.insert("ht".to_string(), row_height.to_string());
                }
                if let Some(_) = db_row.row_record.hidden {
                    row_attribute.insert("hidden".to_string(), "1".to_string());
                }
                if let Some(row_group_level) = db_row.row_record.group_level {
                    row_attribute.insert("outlineLevel".to_string(), row_group_level.to_string());
                }
                if let Some(_) = db_row.row_record.collapsed {
                    row_attribute.insert("collapsed".to_string(), "1".to_string());
                }
                if let Some(_) = db_row.row_record.thick_top {
                    row_attribute.insert("thickTop".to_string(), "1".to_string());
                }
                if let Some(_) = db_row.row_record.thick_bottom {
                    row_attribute.insert("thickBot".to_string(), "1".to_string());
                }
                if let Some(_) = db_row.row_record.place_holder {
                    row_attribute.insert("ph".to_string(), "1".to_string());
                }
                row_element
                    .set_attribute_mut(row_attribute)
                    .context("Failed to set attribute for row")?;
                if let Some(cols) = db_row.cell_records {
                    for (col_index, cell_record) in cols {
                        // Create cell element
                        let cell_element = xml_doc_mut
                            .append_child_mut("c", Some(&row_element_id))
                            .context("Failed to insert row element")?;
                        let cell_id = cell_element.get_id();
                        let mut cell_attribute = HashMap::new();
                        cell_attribute.insert(
                            "r".to_string(),
                            format!(
                                "{}{}",
                                ConverterUtil::get_column_ref(col_index)
                                    .context("Failed to get Char Id from Int")?,
                                row_index
                            ),
                        );
                        if let Some(cell_style_id) = cell_record.style_id {
                            cell_attribute.insert("s".to_string(), cell_style_id.id.to_string());
                        }
                        if cell_record.data_type != CellDataType::Number {
                            cell_attribute.insert(
                                "t".to_string(),
                                CellDataType::get_string(cell_record.data_type),
                            );
                        }
                        if let Some(cell_comment_id) = cell_record.comment_id {
                            cell_attribute.insert("cm".to_string(), cell_comment_id.to_string());
                        }
                        if let Some(cell_metadata) = cell_record.metadata {
                            cell_attribute.insert("vm".to_string(), cell_metadata.to_string());
                        }
                        if let Some(_) = cell_record.place_holder {
                            cell_attribute.insert("ph".to_string(), "1".to_string());
                        }
                        cell_element
                            .set_attribute_mut(cell_attribute)
                            .context("Failed to Set Attribute for cell")?;
                        // Create cell's child element
                        match cell_record.data_type {
                            CellDataType::InlineString => {
                                let inline_string_id = xml_doc_mut
                                    .append_child_mut("is", Some(&cell_id))
                                    .context("Failed to insert Inline string element")?
                                    .get_id();
                                let text_element = xml_doc_mut
                                    .append_child_mut("t", Some(&inline_string_id))
                                    .context("Failed To insert Text Value to inline string")?;
                                text_element.set_value_mut(
                                    if let Some(value) = cell_record.value {
                                        value
                                    } else {
                                        "".to_string()
                                    },
                                );
                            }
                            _ => {
                                if let Some(formula) = cell_record.formula {
                                    let formula_element = xml_doc_mut
                                        .append_child_mut("f", Some(&cell_id))
                                        .context("Failed to insert Inline string element")?;
                                    formula_element.set_value_mut(formula);
                                }
                                xml_doc_mut
                                    .append_child_mut("v", Some(&cell_id))
                                    .context("Failed to insert Inline string element")?
                                    .set_value_mut(if let Some(value) = cell_record.value {
                                        value
                                    } else {
                                        "".to_string()
                                    });
                            }
                        }
                    }
                }
            }
        }
        Ok(())
    }
}

fn deserialize_cols(
    xml_doc_mut: &mut XmlDocument,
) -> AnyResult<Option<VecDeque<ColumnProperties>>, AnyError> {
    if let Some(mut cols_element) = xml_doc_mut.pop_elements_by_tag_mut("cols", None) {
        // Process the columns record if parent node exist
        if let Some(cols) = cols_element.pop() {
            let mut column_collection = VecDeque::with_capacity(cols.get_child_count());
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
                            column_properties.hidden = if hidden == "1" { Some(true) } else { None }
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
            return Ok(Some(column_collection));
        }
    }
    Ok(None)
}

fn deserialize_worksheet_view(
    xml_doc_mut: &mut XmlDocument,
) -> AnyResult<Option<WorkSheetView>, AnyError> {
    Ok(None)
}
fn deserialize_sheet_data(
    xml_doc_mut: &mut XmlDocument,
) -> AnyResult<(Option<BTreeMap<u32, RowData>>, Dimension), AnyError> {
    let mut dimension = Dimension::default();
    if let Some(mut sheet_data_element) = xml_doc_mut.pop_elements_by_tag_mut("sheetData", None) {
        if let Some(sheet_data) = sheet_data_element.pop() {
            let mut sheet_data_collection: BTreeMap<u32, RowData> = BTreeMap::new();
            // Loop All rows of sheet data
            loop {
                if let Some((row_element_id, _)) = sheet_data.pop_child_mut() {
                    if let Some(row_element) = xml_doc_mut.pop_element_mut(&row_element_id) {
                        let mut row_record = RowProperties::default();
                        let row_attribute = row_element
                            .get_attribute()
                            .ok_or(anyhow!("Failed to pull Row Attribute."))?;
                        // Get Row Id
                        let row_index = row_attribute
                            .get("r")
                            .ok_or(anyhow!("Missing mandatory row id attribute"))?
                            .parse()
                            .context("Failed to parse row id")?;
                        if let Some(row_span) = row_attribute.get("spans") {
                            row_record.span = Some(row_span.to_string());
                        }
                        if let Some(style_id) = row_attribute.get("s") {
                            if let Some(custom_formant) = row_attribute.get("customFormat") {
                                row_record.style_id = if custom_formant == "1" {
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
                            row_record.hidden = if hidden == "1" { Some(true) } else { None };
                        }
                        if let Some(height) = row_attribute.get("ht") {
                            if let Some(custom_height) = row_attribute.get("customHeight") {
                                row_record.height = if custom_height == "1" {
                                    Some(height.parse().context("Failed to parse the row height")?)
                                } else {
                                    None
                                };
                            }
                        }
                        if let Some(row_group_level) = row_attribute.get("outlineLevel") {
                            let outline_level = row_group_level
                                .parse()
                                .context("Failed to parse the row group level")?;
                            row_record.group_level = if outline_level > 0 {
                                Some(outline_level)
                            } else {
                                None
                            };
                        }
                        if let Some(collapsed) = row_attribute.get("collapsed") {
                            row_record.collapsed = if collapsed == "1" { Some(true) } else { None };
                        }
                        if let Some(thick_top) = row_attribute.get("thickTop") {
                            row_record.thick_top = if thick_top == "1" { Some(true) } else { None };
                        }
                        if let Some(thick_bottom) = row_attribute.get("thickBot") {
                            row_record.thick_bottom = if thick_bottom == "1" {
                                Some(true)
                            } else {
                                None
                            };
                        }
                        if let Some(place_holder) = row_attribute.get("ph") {
                            row_record.place_holder = if place_holder == "1" {
                                Some(true)
                            } else {
                                None
                            };
                        }
                        let mut cell_records: BTreeMap<u16, CellProperties> = BTreeMap::new();
                        // Loop All Columns of row
                        loop {
                            let mut cell_record = CellProperties::default();
                            if let Some((col_element_id, _)) = row_element.pop_child_mut() {
                                if let Some(col_element) =
                                    xml_doc_mut.pop_element_mut(&col_element_id)
                                {
                                    let cell_attribute = col_element
                                        .get_attribute()
                                        .ok_or(anyhow!("Failed to pull attribute for Column"))?;
                                    // Get Col Id
                                    let col_index = ConverterUtil::get_column_index(
                                        cell_attribute
                                            .get("r")
                                            .ok_or(anyhow!("Missing mandatory col id attribute"))?,
                                    )
                                    .context("Failed to Convert col worksheet initialize")?;
                                    if let Some(style_id) = cell_attribute.get("s") {
                                        cell_record.style_id = Some(StyleId::new(
                                            style_id
                                                .parse()
                                                .context("Failed to parse the col style id")?,
                                        ));
                                    }
                                    if let Some(cell_type) = cell_attribute.get("t") {
                                        cell_record.data_type = CellDataType::get_enum(&cell_type);
                                    } else {
                                        cell_record.data_type = CellDataType::Number;
                                    }
                                    if let Some(comment_id) = cell_attribute.get("cm") {
                                        cell_record.comment_id = Some(
                                            comment_id
                                                .parse()
                                                .context("Failed to parse the col comment id")?,
                                        );
                                    };
                                    if let Some(value_meta_id) = cell_attribute.get("vm") {
                                        cell_record.metadata =
                                            Some(value_meta_id.parse().context(
                                                "Failed to parse the col value meta id",
                                            )?);
                                    };
                                    if let Some(place_holder) = cell_attribute.get("ph") {
                                        cell_record.place_holder = if place_holder == "1" {
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
                                                        cell_record.value =
                                                            element.get_value().clone();
                                                    }
                                                    "f" => {
                                                        cell_record.formula =
                                                            element.get_value().clone();
                                                    }
                                                    "is" => {
                                                        if let Some((text_id, _)) =
                                                            element.pop_child_mut()
                                                        {
                                                            if let Some(text_element) = xml_doc_mut
                                                                .pop_element_mut(&text_id)
                                                            {
                                                                cell_record.value = text_element
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
                                    dimension.start_col = min(dimension.start_col, col_index);
                                    dimension.end_col = max(dimension.end_col, col_index);
                                    cell_records.insert(col_index, cell_record);
                                }
                            } else {
                                break;
                            }
                        }
                        sheet_data_collection.insert(
                            row_index,
                            RowData {
                                row_record,
                                cell_records: if cell_records.len() > 0 {
                                    Some(cell_records)
                                } else {
                                    None
                                },
                            },
                        );
                    }
                } else {
                    break;
                }
            }
            return Ok((Some(sheet_data_collection), dimension));
        }
    }
    Ok((None, dimension))
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
                        if document.check_file_exist(format!(
                            "{}{}/{}{}.{}",
                            relative_path,
                            worksheet_content.default_path,
                            worksheet_content.default_name,
                            sheet_count,
                            worksheet_content.extension
                        )) {
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
        col_index: &u16,
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
        row_index: &u32,
        row_properties: RowProperties,
    ) -> AnyResult<(), AnyError> {
        if let Some(sheet_data) = self.sheet_data.as_mut() {
            if let Some(row) = sheet_data.get_mut(row_index) {
                row.row_record = row_properties;
            } else {
                sheet_data.insert(
                    *row_index,
                    RowData {
                        row_record: row_properties,
                        cell_records: None,
                    },
                );
            }
        } else {
            let mut map = BTreeMap::new();
            map.insert(
                *row_index,
                RowData {
                    row_record: row_properties,
                    cell_records: None,
                },
            );
            self.sheet_data = Some(map);
        }
        Ok(())
    }

    /// Set data for same row multiple columns along with row property
    pub fn set_row_value_ref_mut(
        &mut self,
        cell_ref: &str,
        column_cell: Vec<CellProperties>,
    ) -> AnyResult<(), AnyError> {
        let (row_index, col_index) =
            ConverterUtil::get_cell_index(cell_ref).context("Failed to extract cell key")?;
        self.set_row_value_index_mut(row_index, col_index, column_cell)
    }

    /// Set data for same row multiple columns along with row property
    pub fn set_row_value_index_mut(
        &mut self,
        row_index: u32,
        mut col_index: u16,
        mut column_cell: Vec<CellProperties>,
    ) -> AnyResult<(), AnyError> {
        // Map Start Normalization
        col_index -= 1;
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
        }
        // Load If Sheet Data Exist
        if let Some(sheet_data) = self.sheet_data.as_mut() {
            // Load If Row Exits
            if let Some(row) = sheet_data.get_mut(&row_index) {
                // Check if the row already has cell data
                if let Some(cell_records) = row.cell_records.as_mut() {
                    cell_records.extend(column_cell.iter_mut().map(|item| {
                        col_index += 1;
                        self.dimension.start_col = min(self.dimension.start_col, col_index);
                        self.dimension.end_col = max(self.dimension.end_col, col_index);
                        (col_index, item.clone())
                    }));
                } else {
                    //Create cells If new
                    let mut cell_records = BTreeMap::new();
                    cell_records.extend(column_cell.iter_mut().map(|item| {
                        col_index += 1;
                        self.dimension.start_col = min(self.dimension.start_col, col_index);
                        self.dimension.end_col = max(self.dimension.end_col, col_index);
                        (col_index, item.clone())
                    }));
                    row.cell_records = Some(cell_records);
                }
            } else {
                // Create If new Row
                let mut cell_records = BTreeMap::new();
                cell_records.extend(column_cell.iter_mut().map(|item| {
                    col_index += 1;
                    self.dimension.start_col = min(self.dimension.start_col, col_index);
                    self.dimension.end_col = max(self.dimension.end_col, col_index);
                    (col_index, item.clone())
                }));
                sheet_data.insert(
                    row_index,
                    RowData {
                        row_record: RowProperties::default(),
                        cell_records: Some(cell_records),
                    },
                );
            }
        } else {
            // Create Sheet Data
            let mut sheet_data = BTreeMap::new();
            let mut cell_records = BTreeMap::new();
            cell_records.extend(column_cell.iter_mut().map(|item| {
                col_index += 1;
                self.dimension.start_col = min(self.dimension.start_col, col_index);
                self.dimension.end_col = max(self.dimension.end_col, col_index);
                (col_index, item.clone())
            }));
            sheet_data.insert(
                row_index,
                RowData {
                    row_record: RowProperties::default(),
                    cell_records: Some(cell_records),
                },
            );
            self.sheet_data = Some(sheet_data);
        }
        Ok(())
    }

    fn update_share_string(&mut self, cell_value: &String) -> AnyResult<String, AnyError> {
        if let Some(common_service) = self.common_service.upgrade() {
            common_service
                .try_borrow_mut()
                .context("Failed to Get Share String Handle")?
                .get_string_id_mut(cell_value.to_owned())
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
                .delete_document_mut(&self.file_path);
        }
        self.flush().context("Failed to flush the worksheet")?;
        Ok(())
    }
}
