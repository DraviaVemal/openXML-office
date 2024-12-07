use crate::files::{OfficeDocument, XmlDocument, XmlElement};
use crate::get_all_queries;
use crate::global_2007::traits::{Enum, XmlDocumentPart, XmlDocumentPartCommon};
use crate::spreadsheet_2007::models::{
    BorderSetting, BorderStyleValues, CellXfs, ColorSetting, ColorSettingTypeValues, FillStyle,
    FontStyle, HorizontalAlignmentValues, NumberFormats, PatternTypeValues, SchemeValues,
    VerticalAlignmentValues,
};
use anyhow::{anyhow, Context, Error as AnyError, Result as AnyResult};
use rusqlite::{params, Row};
use serde_json::to_string;
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
            let _ = Self::save_content_to_tree(&self.office_document, &mut self.xml_document)
                .context("Load Share String To DB Failed");
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
            let create_query_cell_style_xfs = queries
                .get("create_cell_style_xfs_table")
                .ok_or_else(|| anyhow!("Expected Query Not Found"))?;
            let create_query_cell_xfs = queries
                .get("create_cell_xfs_table")
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
                .create_table(&create_query_cell_style_xfs)
                .context("Create Cell Style Table Failed")?;
            office_doc
                .get_connection()
                .create_table(&create_query_cell_xfs)
                .context("Create Cell Style Table Failed")?;
        }
        Ok(())
    }

    fn save_content_to_tree(
        office_document: &Weak<RefCell<OfficeDocument>>,
        xml_document: &mut Weak<RefCell<XmlDocument>>,
    ) -> AnyResult<(), AnyError> {
        if let Some(office_doc_ref) = office_document.upgrade() {
            // Decode XML to load in DB
            let office_doc = office_doc_ref
                .try_borrow()
                .context("Pulling Office Doc Failed")?;
            let queries = get_all_queries!("style.sql");
            if let Some(xml_doc) = xml_document.upgrade() {
                let mut xml_doc_mut = xml_doc.try_borrow_mut().context("xml doc borrow failed")?;
                // Create Number Formats Element
                let num_formats = xml_doc_mut
                    .append_child_mut("numFmts", None)
                    .context("Create Number Formats Parent Failed.")?;
                if let Some(num_format_query) = queries.get("select_number_formats_table") {
                    fn row_mapper(row: &Row) -> AnyResult<NumberFormats, rusqlite::Error> {
                        Ok(NumberFormats {
                            id: 0,
                            format_id: row.get(0)?,
                            format_code: row.get(1)?,
                        })
                    }
                    let num_format_data = office_doc
                        .get_connection()
                        .find_many(num_format_query, params![], row_mapper)
                        .context("Num Format Query Results")?;
                    let mut attributes = HashMap::new();
                    attributes.insert("count".to_string(), num_format_data.len().to_string());
                    num_formats
                        .set_attribute_mut(attributes)
                        .context("Updating Number Formats Element Attributes Failed")?;
                    let num_formats_id = num_formats.get_id();
                    for num_format in num_format_data {
                        let num_format_element = xml_doc_mut
                            .append_child_mut("numFmt ", Some(&num_formats_id))
                            .context("Create Number Format Element Failed")?;
                        let mut attributes = HashMap::new();
                        attributes.insert("numFmtId".to_string(), num_format.format_id.to_string());
                        attributes.insert("formatCode".to_string(), num_format.format_code);
                        num_format_element
                            .set_attribute_mut(attributes)
                            .context("Updating Number Format Element Attributes Failed")?;
                    }
                }
                // Create Fonts Element
                let fonts = xml_doc_mut
                    .append_child_mut("fonts", None)
                    .context("Create Fonts Parent Failed.")?;
                let fonts_id = fonts.get_id();
                if let Some(fonts_query) = queries.get("select_fonts_table") {
                    fn row_mapper(row: &Row) -> AnyResult<FontStyle, rusqlite::Error> {
                        let color_type: String = row.get(1)?;
                        let font_scheme: String = row.get(5)?;
                        Ok(FontStyle {
                            id: 0,
                            name: row.get(0)?,
                            color: ColorSetting {
                                color_setting_type: ColorSettingTypeValues::get_enum(&color_type),
                                value: row.get(2)?,
                            },
                            family: row.get(3)?,
                            size: row.get(4)?,
                            font_scheme: SchemeValues::get_enum(&font_scheme),
                            is_bold: row.get(6)?,
                            is_italic: row.get(7)?,
                            is_underline: row.get(8)?,
                            is_double_underline: row.get(9)?,
                        })
                    }
                    let fonts_data = office_doc
                        .get_connection()
                        .find_many(fonts_query, params![], row_mapper)
                        .context("Num Format Query Results")?;
                    for font_style in fonts_data {
                        let font = xml_doc_mut
                            .append_child_mut("font", Some(&fonts_id))
                            .context("Adding Font to Fonts Failed")?;
                        
                    }
                }
            }
        }
        Ok(())
    }

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
            Self::initialize_database(office_document, &queries)
                .context("Database Initialization Failed")?;
            if let Some(xml_doc) = xml_document.upgrade() {
                let mut xml_doc_mut = xml_doc.try_borrow_mut().context("xml doc borrow failed")?;
                // Load Number Format Region
                if let Some(number_formats) = xml_doc_mut
                    .pop_elements_by_tag_mut("numFmts", None)
                    .context("Failed find the Target number format node")?
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
                if let Some(fonts) = xml_doc_mut
                    .pop_elements_by_tag_mut("fonts", None)
                    .context("Failed find the Target font node")?
                    .pop()
                {
                    // fonts
                    loop {
                        // Loop every font element
                        if let Some(font_id) = fonts.pop_child_id_mut() {
                            // font
                            let font = xml_doc_mut
                                .pop_element_mut(&font_id)
                                .ok_or(anyhow!("Element not Found Error"))?;
                            let mut font_style = FontStyle::default();
                            loop {
                                if let Some(item_id) = font.pop_child_id_mut() {
                                    let current_element = xml_doc_mut
                                        .pop_element_mut(&item_id)
                                        .ok_or(anyhow!("Failed to pull child element"))?;
                                    match current_element.get_tag() {
                                        "b" => font_style.is_bold = true,
                                        "u" => {
                                            if let Some(attributes) =
                                                current_element.get_attribute()
                                            {
                                                if let Some(double) = attributes.get("val") {
                                                    if double == "double" {
                                                        font_style.is_double_underline = true;
                                                    }
                                                }
                                            }
                                            font_style.is_underline = false;
                                        }
                                        "i" => font_style.is_italic = false,
                                        "sz" => {
                                            if let Some(attributes) =
                                                current_element.get_attribute()
                                            {
                                                if let Some(val) = attributes.get("val") {
                                                    font_style.size = val
                                                        .parse()
                                                        .context("Font Size Parse Failed")?
                                                }
                                            }
                                        }
                                        "color" => {
                                            if let Some(attributes) =
                                                current_element.get_attribute()
                                            {
                                                if let Some(theme) = attributes.get("theme") {
                                                    font_style.color.color_setting_type =
                                                        ColorSettingTypeValues::Theme;
                                                    font_style.color.value = theme
                                                        .parse()
                                                        .context("Font color theme parse failed")?
                                                } else if let Some(rgb) = attributes.get("rgb") {
                                                    font_style.color.color_setting_type =
                                                        ColorSettingTypeValues::Rgb;
                                                    let rgb_string = rgb.to_string();
                                                    font_style.color.value = rgb_string;
                                                } else if let Some(indexed) =
                                                    attributes.get("indexed")
                                                {
                                                    font_style.color.color_setting_type =
                                                        ColorSettingTypeValues::Indexed;
                                                    let indexed_string = indexed.to_string();
                                                    font_style.color.value = indexed_string;
                                                }
                                            }
                                        }
                                        "name" => {
                                            if let Some(attributes) =
                                                current_element.get_attribute()
                                            {
                                                if let Some(val) = attributes.get("val") {
                                                    font_style.name = val.to_string()
                                                }
                                            }
                                        }
                                        "family" => {
                                            if let Some(attributes) =
                                                current_element.get_attribute()
                                            {
                                                if let Some(val) = attributes.get("val") {
                                                    font_style.family = val
                                                        .parse()
                                                        .context("Font Size Parse Failed")?
                                                }
                                            }
                                        }
                                        "scheme" => {
                                            if let Some(attributes) =
                                                current_element.get_attribute()
                                            {
                                                if let Some(val) = attributes.get("val") {
                                                    font_style.font_scheme =
                                                        SchemeValues::get_enum(val)
                                                }
                                            }
                                        }
                                        _ => {
                                            return Err(anyhow!("Unknown Font Style Found!"));
                                        }
                                    }
                                } else {
                                    break;
                                }
                            }
                            // Insert Data into Database
                            let insert_query_font_style = queries
                                .get("insert_font_style_table")
                                .ok_or_else(|| anyhow!("Expected Query Not Found"))?;

                            office_doc
                                .get_connection()
                                .insert_record(
                                    &insert_query_font_style,
                                    params![
                                        font_style.name,
                                        ColorSettingTypeValues::get_string(
                                            font_style.color.color_setting_type
                                        ),
                                        font_style.color.value,
                                        font_style.family,
                                        font_style.size,
                                        SchemeValues::get_string(font_style.font_scheme),
                                        font_style.is_bold,
                                        font_style.is_italic,
                                        font_style.is_underline,
                                        font_style.is_double_underline
                                    ],
                                )
                                .context("Insert Font Style Failed")?;
                        } else {
                            break;
                        }
                    }
                }
                if let Some(fills) = xml_doc_mut
                    .pop_elements_by_tag_mut("fills", None)
                    .context("Failed find the Target fill node")?
                    .pop()
                {
                    loop {
                        if let Some(fill_id) = fills.pop_child_id_mut() {
                            let current_element = xml_doc_mut
                                .pop_element_mut(&fill_id)
                                .ok_or(anyhow!("Failed to pull child element"))?;
                            let mut fill_style = FillStyle::default();
                            if let Some(pattern_fill_id) = current_element.pop_child_id_mut() {
                                if let Some(pattern_fill) =
                                    xml_doc_mut.pop_element_mut(&pattern_fill_id)
                                {
                                    if let Some(pattern_attributes) = pattern_fill.get_attribute() {
                                        if let Some(pattern_type) =
                                            pattern_attributes.get("patternType")
                                        {
                                            fill_style.pattern_type =
                                                PatternTypeValues::get_enum(pattern_type);
                                            loop {
                                                if let Some(child_id) =
                                                    pattern_fill.pop_child_id_mut()
                                                {
                                                    if let Some(pop_child) =
                                                        xml_doc_mut.pop_element_mut(&child_id)
                                                    {
                                                        if let Some(attributes) =
                                                            pop_child.get_attribute()
                                                        {
                                                            match pop_child.get_tag() {
                                                                "fgColor" => {
                                                                    if let Some(theme) =
                                                                        attributes.get("theme")
                                                                    {
                                                                        fill_style
                                                                            .foreground_color =
                                                                            Some(ColorSetting {
                                                                                color_setting_type:ColorSettingTypeValues::Theme,
                                                                                value:theme
                                                                                .parse()
                                                                                .context("color theme parse failed")?
                                                                            });
                                                                    } else if let Some(rgb) =
                                                                        attributes.get("rgb")
                                                                    {
                                                                        let rgb_string =
                                                                            rgb.to_string();
                                                                        fill_style
                                                                            .foreground_color =
                                                                            Some(ColorSetting {
                                                                                color_setting_type:ColorSettingTypeValues::Rgb,
                                                                                value:rgb_string
                                                                            });
                                                                    } else if let Some(indexed) =
                                                                        attributes.get("indexed")
                                                                    {
                                                                        fill_style
                                                                            .foreground_color =
                                                                            Some(ColorSetting {
                                                                                color_setting_type:ColorSettingTypeValues::Indexed,
                                                                                value:indexed
                                                                                .parse()
                                                                                .context("color color index parse failed")?
                                                                            });
                                                                    }
                                                                }
                                                                "bgColor" => {
                                                                    if let Some(theme) =
                                                                        attributes.get("theme")
                                                                    {
                                                                        fill_style
                                                                            .background_color =
                                                                            Some(ColorSetting {
                                                                                color_setting_type:ColorSettingTypeValues::Theme,
                                                                                value:theme
                                                                                .parse()
                                                                                .context("color theme parse failed")?
                                                                            });
                                                                    } else if let Some(rgb) =
                                                                        attributes.get("rgb")
                                                                    {
                                                                        let rgb_string =
                                                                            rgb.to_string();
                                                                        fill_style
                                                                            .background_color =
                                                                            Some(ColorSetting {
                                                                                color_setting_type:ColorSettingTypeValues::Rgb,
                                                                                value:rgb_string
                                                                            });
                                                                    } else if let Some(indexed) =
                                                                        attributes.get("indexed")
                                                                    {
                                                                        fill_style
                                                                            .background_color =
                                                                            Some(ColorSetting {
                                                                                color_setting_type:ColorSettingTypeValues::Indexed,
                                                                                value:indexed
                                                                                .parse()
                                                                                .context("color color index parse failed")?
                                                                            });
                                                                    }
                                                                }
                                                                _ => {
                                                                    return Err(anyhow!(
                                                                        "Unknown Color patter found"
                                                                    ));
                                                                }
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
                            }
                            // Insert Tables Queries
                            let insert_query_fill_style = queries
                                .get("insert_fill_style_table")
                                .ok_or_else(|| anyhow!("Expected Query Not Found"))?;
                            office_doc
                                .get_connection()
                                .insert_record(
                                    &insert_query_fill_style,
                                    params![
                                        to_string(&fill_style.background_color)
                                            .context("Background Fill Parse Failed")?,
                                        to_string(&fill_style.foreground_color)
                                            .context("Foreground Fill Parse Failed")?,
                                        PatternTypeValues::get_string(fill_style.pattern_type)
                                    ],
                                )
                                .context("Insert Fill Style Failed")?;
                        } else {
                            break;
                        }
                    }
                }
                if let Some(borders) = xml_doc_mut
                    .pop_elements_by_tag_mut("borders", None)
                    .context("Failed find the Target border node")?
                    .pop()
                {
                    loop {
                        if let Some(border_id) = borders.pop_child_id_mut() {
                            if let Some(border) = xml_doc_mut.pop_element_mut(&border_id) {
                                let mut left_border = BorderSetting::default();
                                let mut right_border = BorderSetting::default();
                                let mut top_border = BorderSetting::default();
                                let mut bottom_border = BorderSetting::default();
                                let mut diagonal_border = BorderSetting::default();
                                loop {
                                    if let Some(border_child_id) = border.pop_child_id_mut() {
                                        if let Some(current_element) =
                                            xml_doc_mut.pop_element_mut(&border_child_id)
                                        {
                                            match current_element.get_tag() {
                                                "left" => {
                                                    Style::deserialize_border_setting(
                                                        &current_element,
                                                        &mut left_border,
                                                        &mut xml_doc_mut,
                                                    )
                                                    .context("Left Border Decode Failed")?;
                                                }
                                                "right" => {
                                                    Style::deserialize_border_setting(
                                                        &current_element,
                                                        &mut right_border,
                                                        &mut xml_doc_mut,
                                                    )
                                                    .context("Left Border Decode Failed")?;
                                                }
                                                "top" => {
                                                    Style::deserialize_border_setting(
                                                        &current_element,
                                                        &mut top_border,
                                                        &mut xml_doc_mut,
                                                    )
                                                    .context("Left Border Decode Failed")?;
                                                }
                                                "bottom" => {
                                                    Style::deserialize_border_setting(
                                                        &current_element,
                                                        &mut bottom_border,
                                                        &mut xml_doc_mut,
                                                    )
                                                    .context("Left Border Decode Failed")?;
                                                }
                                                "diagonal" => {
                                                    Style::deserialize_border_setting(
                                                        &current_element,
                                                        &mut diagonal_border,
                                                        &mut xml_doc_mut,
                                                    )
                                                    .context("Left Border Decode Failed")?;
                                                }
                                                _ => {
                                                    return Err(anyhow!(
                                                        "Unknown border style found"
                                                    ));
                                                }
                                            }
                                        }
                                    } else {
                                        break;
                                    }
                                }
                                let insert_query_border_style = queries
                                    .get("insert_border_style_table")
                                    .ok_or_else(|| anyhow!("Expected Query Not Found"))?;
                                office_doc
                                    .get_connection()
                                    .insert_record(
                                        &insert_query_border_style,
                                        params![
                                            to_string(&left_border)
                                                .context("Left Border Parsing Failed")?,
                                            to_string(&top_border)
                                                .context("Top Border Parsing Failed")?,
                                            to_string(&right_border)
                                                .context("Right Border Parsing Failed")?,
                                            to_string(&bottom_border)
                                                .context("Bottom Border Parsing Failed")?,
                                            to_string(&diagonal_border)
                                                .context("Bottom Border Parsing Failed")?
                                        ],
                                    )
                                    .context("Insert border Style Failed")?;
                            }
                        } else {
                            break;
                        }
                    }
                }
                if let Some(cell_style_xfs) = xml_doc_mut
                    .pop_elements_by_tag_mut("cellStyleXfs", None)
                    .context("Failed find the Target cell style xfs node")?
                    .pop()
                {
                    Style::deserialize_cell_style(
                        cell_style_xfs,
                        &mut xml_doc_mut,
                        &office_doc,
                        &queries,
                        "insert_cell_style_xfs_table",
                    )
                    .context("Deserializing Cell Style Xfs Failed")?;
                }
                if let Some(cell_xfs) = xml_doc_mut
                    .pop_elements_by_tag_mut("cellXfs", None)
                    .context("Failed find the Target cell xfs node")?
                    .pop()
                {
                    Style::deserialize_cell_style(
                        cell_xfs,
                        &mut xml_doc_mut,
                        &office_doc,
                        &queries,
                        "insert_cell_xfs_table",
                    )
                    .context("Deserializing Cell Xfs Failed")?;
                }
            }
        }
        Ok(())
    }
    pub(crate) fn deserialize_cell_style(
        style_xfs: XmlElement,
        xml_doc_mut: &mut std::cell::RefMut<'_, XmlDocument>,
        office_doc: &std::cell::Ref<'_, OfficeDocument>,
        queries: &HashMap<String, String>,
        query_target: &str,
    ) -> Result<(), AnyError> {
        loop {
            if let Some(xf_id) = style_xfs.pop_child_id_mut() {
                let mut cell_xf = CellXfs::default();
                if let Some(current_element) = xml_doc_mut.pop_element_mut(&xf_id) {
                    if let Some(attributes) = current_element.get_attribute() {
                        if let Some(number_format_id) = attributes.get("numFmtId") {
                            cell_xf.number_format_id = number_format_id
                                .parse()
                                .context("Number Number Format Id Parse Failed")?;
                        }
                        if let Some(font_id) = attributes.get("fontId") {
                            cell_xf.font_id =
                                font_id.parse().context("Number Font Id Parse Failed")?;
                        }
                        if let Some(fill_id) = attributes.get("fillId") {
                            cell_xf.fill_id =
                                fill_id.parse().context("Number Fill Id Parse Failed")?;
                        }
                        if let Some(border_id) = attributes.get("borderId") {
                            cell_xf.border_id =
                                border_id.parse().context("Number Border Id Parse Failed")?;
                        }
                        if let Some(format_id) = attributes.get("xfId") {
                            cell_xf.format_id =
                                format_id.parse().context("Number Format Id Parse Failed")?;
                        }
                        if let Some(apply_protection) = attributes.get("applyProtection") {
                            cell_xf.apply_protection = apply_protection
                                .parse()
                                .context("Number Apply Protection Parse Failed")?;
                        }
                        if let Some(apply_alignment) = attributes.get("applyAlignment") {
                            cell_xf.apply_alignment = apply_alignment
                                .parse()
                                .context("Number Apply Alignment Parse Failed")?;
                        }
                        if let Some(apply_border) = attributes.get("applyBorder") {
                            cell_xf.apply_border = apply_border
                                .parse()
                                .context("Number Apply Border Parse Failed")?;
                        }
                        if let Some(apply_fill) = attributes.get("applyFill") {
                            cell_xf.apply_fill = apply_fill
                                .parse()
                                .context("Number Apply Fill Parse Failed")?;
                        }
                        if let Some(apply_font) = attributes.get("applyFont") {
                            cell_xf.apply_font = apply_font
                                .parse()
                                .context("Number Apply Font Parse Failed")?;
                        }
                        if let Some(apply_number_format) = attributes.get("applyNumberFormat") {
                            cell_xf.apply_number_format = apply_number_format
                                .parse()
                                .context("Number Apply Number Format Parse Failed")?;
                        }
                        // Load Alignment Values if exist
                        if let Some(alignment_id) = current_element.pop_child_id_mut() {
                            if let Some(alignment_element) =
                                xml_doc_mut.pop_element_mut(&alignment_id)
                            {
                                if let Some(alignment_attributes) =
                                    alignment_element.get_attribute()
                                {
                                    if let Some(is_wrap_text) = alignment_attributes.get("wrapText")
                                    {
                                        cell_xf.is_wrap_text = is_wrap_text
                                            .parse()
                                            .context("Number Wrap Text Parse Failed")?;
                                    }
                                    if let Some(vertical_alignment) =
                                        alignment_attributes.get("horizontal")
                                    {
                                        cell_xf.vertical_alignment =
                                            VerticalAlignmentValues::get_enum(vertical_alignment);
                                    }
                                    if let Some(horizontal_alignment) =
                                        alignment_attributes.get("horizontal")
                                    {
                                        cell_xf.horizontal_alignment =
                                            HorizontalAlignmentValues::get_enum(
                                                horizontal_alignment,
                                            );
                                    }
                                }
                            }
                        }
                    }
                    let insert_query_style_xfs = queries
                        .get(query_target)
                        .ok_or_else(|| anyhow!("Expected Query Not Found"))?;
                    office_doc
                        .get_connection()
                        .insert_record(
                            &insert_query_style_xfs,
                            params![
                                cell_xf.format_id,
                                cell_xf.number_format_id,
                                cell_xf.font_id,
                                cell_xf.fill_id,
                                cell_xf.border_id,
                                cell_xf.apply_font,
                                cell_xf.apply_alignment,
                                cell_xf.apply_fill,
                                cell_xf.apply_border,
                                cell_xf.apply_number_format,
                                cell_xf.apply_protection,
                                cell_xf.is_wrap_text,
                                HorizontalAlignmentValues::get_string(cell_xf.horizontal_alignment),
                                VerticalAlignmentValues::get_string(cell_xf.vertical_alignment),
                            ],
                        )
                        .context("Insert border Style Failed")?;
                }
            } else {
                break;
            }
        }
        Ok(())
    }

    pub(crate) fn deserialize_border_setting(
        current_element: &XmlElement,
        border: &mut BorderSetting,
        xml_doc_mut: &mut std::cell::RefMut<'_, XmlDocument>,
    ) -> Result<(), AnyError> {
        Ok(if let Some(attributes) = current_element.get_attribute() {
            if let Some(style) = attributes.get("style") {
                border.style = BorderStyleValues::get_enum(&style);
                if border.style != BorderStyleValues::None {
                    if let Some(color_id) = current_element.pop_child_id_mut() {
                        if let Some(color_element) = xml_doc_mut.pop_element_mut(&color_id) {
                            if let Some(attributes) = color_element.get_attribute() {
                                if let Some(theme) = attributes.get("theme") {
                                    border.border_color = Some(ColorSetting {
                                        color_setting_type: ColorSettingTypeValues::Theme,
                                        value: theme.parse().context("color theme parse failed")?,
                                    });
                                } else if let Some(rgb) = attributes.get("rgb") {
                                    let rgb_string = rgb.to_string();
                                    border.border_color = Some(ColorSetting {
                                        color_setting_type: ColorSettingTypeValues::Rgb,
                                        value: rgb_string,
                                    });
                                } else if let Some(indexed) = attributes.get("indexed") {
                                    border.border_color = Some(ColorSetting {
                                        color_setting_type: ColorSettingTypeValues::Indexed,
                                        value: indexed
                                            .parse()
                                            .context("color color index parse failed")?,
                                    });
                                }
                            }
                        }
                    }
                }
            }
        })
    }
}
