use crate::files::{OfficeDocument, XmlDocument, XmlElement};
use crate::get_all_queries;
use crate::global_2007::traits::{Enum, XmlDocumentPart, XmlDocumentPartCommon};
use crate::spreadsheet_2007::models::{
    BorderSetting, BorderStyleValues, ColorSetting, ColorSettingTypeValues, FillStyle, FontStyle,
    PatternTypeValues, SchemeValues,
};
use anyhow::{anyhow, Context, Error as AnyError, Result as AnyResult};
use rusqlite::params;
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
                if let Some(fonts) = xml_doc_mut
                    .pop_elements_by_tag_mut("fonts", None)
                    .context("Failed find the Target node")?
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
                    .context("Failed find the Target node")?
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
                    .context("Failed find the Target node")?
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
                                                    deserialize_border_setting(
                                                        &current_element,
                                                        &mut left_border,
                                                        &mut xml_doc_mut,
                                                    )
                                                    .context("Left Border Decode Failed")?;
                                                }
                                                "right" => {
                                                    deserialize_border_setting(
                                                        &current_element,
                                                        &mut right_border,
                                                        &mut xml_doc_mut,
                                                    )
                                                    .context("Left Border Decode Failed")?;
                                                }
                                                "top" => {
                                                    deserialize_border_setting(
                                                        &current_element,
                                                        &mut top_border,
                                                        &mut xml_doc_mut,
                                                    )
                                                    .context("Left Border Decode Failed")?;
                                                }
                                                "bottom" => {
                                                    deserialize_border_setting(
                                                        &current_element,
                                                        &mut bottom_border,
                                                        &mut xml_doc_mut,
                                                    )
                                                    .context("Left Border Decode Failed")?;
                                                }
                                                "diagonal" => {
                                                    deserialize_border_setting(
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

                if let Some(cell_style_xf) = xml_doc_mut
                    .pop_elements_by_tag_mut("cell_style_xfs", None)
                    .context("Failed find the Target node")?
                    .pop()
                {}
                if let Some(cell_xfs) = xml_doc_mut
                    .pop_elements_by_tag_mut("cell_xfs", None)
                    .context("Failed find the Target node")?
                    .pop()
                {}
                if let Some(cell_styles) = xml_doc_mut
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

fn deserialize_border_setting(
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
