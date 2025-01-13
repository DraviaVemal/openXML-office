use crate::{
    converters::ConverterUtil,
    element_dictionary::EXCEL_TYPE_COLLECTION,
    files::{OfficeDocument, XmlDocument, XmlElement, XmlSerializer},
    global_2007::{
        parts::RelationsPart,
        traits::{Enum, XmlDocumentPart, XmlDocumentPartCommon},
    },
    log_elapsed,
    spreadsheet_2007::models::{
        BorderSetting, BorderStyle, BorderStyleValues, CellXfs, ColorSetting,
        ColorSettingTypeValues, FillStyle, FontSchemeValues, FontStyle, HorizontalAlignmentValues,
        NumberFormat, NumberFormatValues, PatternTypeValues, StyleId, StyleSetting,
        VerticalAlignmentValues,
    },
};
use anyhow::{anyhow, Context, Error as AnyError, Result as AnyResult};
use std::{
    cell::RefCell,
    collections::{HashMap, VecDeque},
    hash::{DefaultHasher, Hash, Hasher},
    rc::Weak,
};

#[derive(Debug)]
pub struct StylePart {
    office_document: Weak<RefCell<OfficeDocument>>,
    xml_document: Weak<RefCell<XmlDocument>>,
    file_path: String,
    cache_id: HashMap<u64, u32>,
    cache_order: VecDeque<u64>,
    cache_capacity: u8,
    number_format_collection: Vec<NumberFormat>,
    font_collection: Vec<FontStyle>,
    fill_collection: Vec<FillStyle>,
    border_collection: Vec<BorderStyle>,
    cell_style_collection: Vec<CellXfs>,
    cell_collection: Vec<CellXfs>,
}

impl Drop for StylePart {
    fn drop(&mut self) {
        let _ = self.close_document();
    }
}

impl XmlDocumentPartCommon for StylePart {
    fn close_document(&mut self) -> AnyResult<(), AnyError>
    where
        Self: Sized,
    {
        log_elapsed!(
            || {
                log_elapsed!(
                    || {
                        {
                            self.save_content_to_tree_mut()
                                .context("Style Save Content Failed")
                        }
                    },
                    "Save Tree"
                )?;
                if let Some(xml_tree) = self.office_document.upgrade() {
                    log_elapsed!(
                        || {
                            xml_tree
                                .try_borrow_mut()
                                .context("Failed To pull XML Handle")?
                                .close_xml_document(&self.file_path)
                        },
                        "Close Document"
                    )?;
                }
                Ok(())
            },
            "Close Style Service"
        )
    }
    /// Initialize xml content for this part from base template
    fn initialize_content_xml() -> AnyResult<(XmlDocument, Option<String>, String, String), AnyError>
    {
        let content = EXCEL_TYPE_COLLECTION.get("style").unwrap();
        Ok((
            XmlSerializer::vec_to_xml_doc_tree(
                include_str!("style.xml").as_bytes().to_vec(),
                "Default Style",
            )
            .context("Initializing Theme Failed")?,
            Some(content.content_type.to_string()),
            content.extension.to_string(),
            content.extension_type.to_string(),
        ))
    }
}

impl XmlDocumentPart for StylePart {
    fn new(
        office_document: Weak<RefCell<OfficeDocument>>,
        parent_relationship_part: Weak<RefCell<RelationsPart>>,
    ) -> AnyResult<Self, AnyError> {
        let file_name = Self::get_style_file_name(&parent_relationship_part)
            .context("Failed to pull style file name")?
            .to_string();
        let mut xml_document = Self::get_xml_document(&office_document, &file_name)?;
        let (
            number_format_collection,
            font_collection,
            fill_collection,
            border_collection,
            cell_style_collection,
            cell_collection,
        ) = Self::load_content(&mut xml_document).context("Load Share String To DB Failed")?;
        Ok(Self {
            office_document,
            xml_document,
            file_path: file_name,
            cache_id: HashMap::new(),
            cache_order: VecDeque::new(),
            cache_capacity: 25,
            number_format_collection,
            font_collection,
            fill_collection,
            border_collection,
            cell_style_collection,
            cell_collection,
        })
    }
}

// ################################# Load / Save Functions ################
impl StylePart {
    fn get_style_file_name(
        relations_part: &Weak<RefCell<RelationsPart>>,
    ) -> AnyResult<String, AnyError> {
        let style_content = EXCEL_TYPE_COLLECTION.get("style").unwrap();
        if let Some(relations_part) = relations_part.upgrade() {
            Ok(relations_part
                .try_borrow_mut()
                .context("Failed to pull relationship connection")?
                .get_relationship_target_by_type_mut(
                    &style_content.schemas_type,
                    style_content,
                    None,
                    None,
                )
                .context("Pull Path From Existing File Failed")?)
        } else {
            Err(anyhow!("Failed to upgrade relation part"))
        }
    }

    /// Load existing file style to database
    fn load_content(
        xml_document: &mut Weak<RefCell<XmlDocument>>,
    ) -> AnyResult<
        (
            Vec<NumberFormat>,
            Vec<FontStyle>,
            Vec<FillStyle>,
            Vec<BorderStyle>,
            Vec<CellXfs>,
            Vec<CellXfs>,
        ),
        AnyError,
    > {
        let mut num_format_records = Vec::new();
        let mut font_records = Vec::new();
        let mut fill_records = Vec::new();
        let mut border_records = Vec::new();
        let mut style_collection = Vec::new();
        let mut xfs_collection = Vec::new();
        if let Some(xml_document) = xml_document.upgrade() {
            let mut xml_doc_mut = xml_document
                .try_borrow_mut()
                .context("xml doc borrow failed")?;
            // Load Number Format Region
            if let Some(mut number_formats_vec) =
                xml_doc_mut.pop_elements_by_tag_mut("numFmts", None)
            {
                if let Some(number_formats) = number_formats_vec.pop() {
                    // Load Number Format from File if exist
                    loop {
                        if let Some((element_id, _)) = number_formats.pop_child_mut() {
                            let num_fmt = xml_doc_mut
                                .pop_element_mut(&element_id)
                                .ok_or(anyhow!("Element not Found Error"))?;
                            if let Some(attributes) = num_fmt.get_attribute() {
                                let mut number_format = NumberFormat::default();
                                number_format.format_id = attributes
                                    .get("numFmtId")
                                    .ok_or(anyhow!("numFmtId Attribute Not Found!"))?
                                    .parse()
                                    .context("Number format ID parsing Failed")?;
                                number_format.format_code = attributes
                                    .get("formatCode")
                                    .ok_or(anyhow!("formatCode Attribute Not Found!"))?
                                    .to_string();
                                num_format_records.push(number_format);
                            }
                        } else {
                            break;
                        }
                    }
                }
            }
            if let Some(mut fonts_vec) = xml_doc_mut.pop_elements_by_tag_mut("fonts", None) {
                if let Some(fonts) = fonts_vec.pop() {
                    // fonts
                    loop {
                        // Loop every font element
                        if let Some((font_id, _)) = fonts.pop_child_mut() {
                            // font
                            let font = xml_doc_mut
                                .pop_element_mut(&font_id)
                                .ok_or(anyhow!("Element not Found Error"))?;
                            let mut font_style = FontStyle::default();
                            loop {
                                if let Some((item_id, _)) = font.pop_child_mut() {
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
                                                        FontSchemeValues::get_enum(val)
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
                            font_records.push(font_style);
                        } else {
                            break;
                        }
                    }
                }
            }
            if let Some(mut fills_vec) = xml_doc_mut.pop_elements_by_tag_mut("fills", None) {
                if let Some(fills) = fills_vec.pop() {
                    loop {
                        if let Some((fill_id, _)) = fills.pop_child_mut() {
                            let current_element = xml_doc_mut
                                .pop_element_mut(&fill_id)
                                .ok_or(anyhow!("Failed to pull child element"))?;
                            let mut fill_style = FillStyle::default();
                            if let Some((pattern_fill_id, _)) = current_element.pop_child_mut() {
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
                                                if let Some((child_id, _)) =
                                                    pattern_fill.pop_child_mut()
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
                            fill_records.push(fill_style);
                        } else {
                            break;
                        }
                    }
                }
            }
            if let Some(mut borders_vec) = xml_doc_mut.pop_elements_by_tag_mut("borders", None) {
                if let Some(borders) = borders_vec.pop() {
                    // Loop Each Border Style
                    loop {
                        if let Some((border_id, _)) = borders.pop_child_mut() {
                            if let Some(border) = xml_doc_mut.pop_element_mut(&border_id) {
                                let mut border_style = BorderStyle::default();
                                // Loop Details of current border
                                loop {
                                    if let Some((border_child_id, _)) = border.pop_child_mut() {
                                        if let Some(current_element) =
                                            xml_doc_mut.pop_element_mut(&border_child_id)
                                        {
                                            match current_element.get_tag() {
                                                "left" => {
                                                    StylePart::deserialize_border_setting(
                                                        &current_element,
                                                        &mut border_style.left,
                                                        &mut xml_doc_mut,
                                                    )
                                                    .context("Left Border Decode Failed")?;
                                                }
                                                "right" => {
                                                    StylePart::deserialize_border_setting(
                                                        &current_element,
                                                        &mut border_style.right,
                                                        &mut xml_doc_mut,
                                                    )
                                                    .context("Left Border Decode Failed")?;
                                                }
                                                "top" => {
                                                    StylePart::deserialize_border_setting(
                                                        &current_element,
                                                        &mut border_style.top,
                                                        &mut xml_doc_mut,
                                                    )
                                                    .context("Left Border Decode Failed")?;
                                                }
                                                "bottom" => {
                                                    StylePart::deserialize_border_setting(
                                                        &current_element,
                                                        &mut border_style.bottom,
                                                        &mut xml_doc_mut,
                                                    )
                                                    .context("Left Border Decode Failed")?;
                                                }
                                                "diagonal" => {
                                                    StylePart::deserialize_border_setting(
                                                        &current_element,
                                                        &mut border_style.diagonal,
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
                                border_records.push(border_style);
                            }
                        } else {
                            break;
                        }
                    }
                }
            }
            if let Some(mut cell_style_xfs_vec) =
                xml_doc_mut.pop_elements_by_tag_mut("cellStyleXfs", None)
            {
                if let Some(cell_style_xfs) = cell_style_xfs_vec.pop() {
                    style_collection =
                        StylePart::deserialize_cell_style(cell_style_xfs, &mut xml_doc_mut)
                            .context("Deserializing Cell Style Xfs Failed")?;
                }
            }
            if let Some(mut cell_xfs_vec) = xml_doc_mut.pop_elements_by_tag_mut("cellXfs", None) {
                if let Some(cell_xfs) = cell_xfs_vec.pop() {
                    xfs_collection = StylePart::deserialize_cell_style(cell_xfs, &mut xml_doc_mut)
                        .context("Deserializing Cell Xfs Failed")?;
                }
            }
        }
        Ok((
            num_format_records,
            font_records,
            fill_records,
            border_records,
            style_collection,
            xfs_collection,
        ))
    }

    /// Save Database record back to XML File
    fn save_content_to_tree_mut(&mut self) -> AnyResult<(), AnyError> {
        if let Some(xml_document) = self.xml_document.upgrade() {
            let mut xml_doc_mut = xml_document
                .try_borrow_mut()
                .context("xml doc borrow failed")?;
            // Create Number Formats Elements
            {
                let num_formats = xml_doc_mut
                    .insert_child_at_mut("numFmts", &0, None)
                    .context("Create Number Formats Parent Failed.")?;
                let num_formats_id = num_formats.get_id();
                let mut attributes = HashMap::new();
                attributes.insert(
                    "count".to_string(),
                    self.number_format_collection.len().to_string(),
                );
                num_formats
                    .set_attribute_mut(attributes)
                    .context("Updating Number Formats Element Attributes Failed")?;
                for num_format in self.number_format_collection.as_slice() {
                    let num_format_element = xml_doc_mut
                        .append_child_mut("numFmt", Some(&num_formats_id))
                        .context("Create Number Format Element Failed")?;
                    let mut attributes = HashMap::new();
                    attributes.insert("numFmtId".to_string(), num_format.format_id.to_string());
                    attributes.insert("formatCode".to_string(), num_format.format_code.clone());
                    num_format_element
                        .set_attribute_mut(attributes)
                        .context("Updating Number Format Element Attributes Failed")?;
                }
            }
            // Create Fonts Elements
            {
                let fonts = xml_doc_mut
                    .insert_children_after_tag_mut("fonts", "numFmts", None)
                    .context("Create Fonts Parent Failed.")?;
                let fonts_id = fonts.get_id();
                let mut attributes = HashMap::new();
                attributes.insert("count".to_string(), self.font_collection.len().to_string());
                fonts
                    .set_attribute_mut(attributes)
                    .context("Set attribute failed for fonts style")?;
                for font_style in self.font_collection.as_slice() {
                    let font_id = xml_doc_mut
                        .append_child_mut("font", Some(&fonts_id))
                        .context("Adding Font to Fonts Failed")?
                        .get_id();
                    if font_style.is_bold {
                        xml_doc_mut
                            .append_child_mut("b", Some(&font_id))
                            .context("Create Size Failed")?;
                    }
                    if font_style.is_italic {
                        xml_doc_mut
                            .append_child_mut("i", Some(&font_id))
                            .context("Create Size Failed")?;
                    }
                    if font_style.is_underline {
                        xml_doc_mut
                            .append_child_mut("u", Some(&font_id))
                            .context("Create Size Failed")?;
                    }
                    if font_style.is_double_underline {
                        let double_underline = xml_doc_mut
                            .append_child_mut("u", Some(&font_id))
                            .context("Create Size Failed")?;
                        let mut double_underline_attributes: HashMap<String, String> =
                            HashMap::new();
                        double_underline_attributes.insert("val".to_string(), "double".to_string());
                        double_underline
                            .set_attribute_mut(double_underline_attributes)
                            .context("Setting Size Attribute Failing")?;
                    }
                    let size = xml_doc_mut
                        .append_child_mut("sz", Some(&font_id))
                        .context("Create Size Failed")?;
                    let mut size_attributes: HashMap<String, String> = HashMap::new();
                    size_attributes.insert("val".to_string(), font_style.size.to_string());
                    size.set_attribute_mut(size_attributes)
                        .context("Setting Size Attribute Failing")?;
                    StylePart::add_color_element(
                        Some(font_style.color.clone()),
                        &mut xml_doc_mut,
                        font_id,
                    )?;
                    let name = xml_doc_mut
                        .append_child_mut("name", Some(&font_id))
                        .context("Create Name Failed")?;
                    let mut name_attributes: HashMap<String, String> = HashMap::new();
                    name_attributes.insert("val".to_string(), font_style.name.clone());
                    name.set_attribute_mut(name_attributes)
                        .context("Setting name Attribute Failing")?;
                    let family = xml_doc_mut
                        .append_child_mut("family", Some(&font_id))
                        .context("Create Name Failed")?;
                    let mut family_attributes: HashMap<String, String> = HashMap::new();
                    family_attributes.insert("val".to_string(), font_style.family.to_string());
                    family
                        .set_attribute_mut(family_attributes)
                        .context("Setting family attribute failing")?;
                    let scheme = xml_doc_mut
                        .append_child_mut("scheme", Some(&font_id))
                        .context("Create scheme Failed")?;
                    let mut scheme_attributes: HashMap<String, String> = HashMap::new();
                    scheme_attributes.insert(
                        "val".to_string(),
                        FontSchemeValues::get_string(font_style.font_scheme.clone()),
                    );
                    scheme
                        .set_attribute_mut(scheme_attributes)
                        .context("Setting scheme attribute failing")?;
                }
            }
            // Create Fills Elements
            {
                let fills = xml_doc_mut
                    .insert_children_after_tag_mut("fills", "fonts", None)
                    .context("Create Fills parents Failed.")?;
                let fills_id = fills.get_id();
                let mut attributes = HashMap::new();
                attributes.insert("count".to_string(), self.fill_collection.len().to_string());
                fills
                    .set_attribute_mut(attributes)
                    .context("Set Fill Attribute Failed")?;
                for fill_data in self.fill_collection.as_slice() {
                    let fill_id = xml_doc_mut
                        .append_child_mut("fill", Some(&fills_id))
                        .context("Adding Fill Element Failed")?
                        .get_id();
                    let pattern_fill_element = xml_doc_mut
                        .append_child_mut("patternFill", Some(&fill_id))
                        .context("Pattern Fill Element Failed")?;
                    let mut pattern_attribute: HashMap<String, String> = HashMap::new();
                    pattern_attribute.insert(
                        "patternType".to_string(),
                        PatternTypeValues::get_string(fill_data.pattern_type.clone()),
                    );
                    pattern_fill_element
                        .set_attribute_mut(pattern_attribute)
                        .context("Set Pattern Fill Attribute Failed")?;
                    let pattern_fill_id = pattern_fill_element.get_id();
                    if let Some(fg_setting) = fill_data.foreground_color.clone() {
                        let fg_element = xml_doc_mut
                            .append_child_mut("fgColor", Some(&pattern_fill_id))
                            .context("Pattern Fill Foreground Element Failed")?;
                        let mut fg_attributes: HashMap<String, String> = HashMap::new();
                        fg_attributes.insert(
                            ColorSettingTypeValues::get_string(fg_setting.color_setting_type),
                            fg_setting.value,
                        );
                        fg_element
                            .set_attribute_mut(fg_attributes)
                            .context("Set Foreground attribute Failed")?;
                    }
                    if let Some(bg_setting) = fill_data.background_color.clone() {
                        let bg_element = xml_doc_mut
                            .append_child_mut("bgColor", Some(&pattern_fill_id))
                            .context("Pattern Fill Background Element Failed")?;
                        let mut bg_attributes: HashMap<String, String> = HashMap::new();
                        bg_attributes.insert(
                            ColorSettingTypeValues::get_string(bg_setting.color_setting_type),
                            bg_setting.value,
                        );
                        bg_element
                            .set_attribute_mut(bg_attributes)
                            .context("Set Background attribute Failed")?;
                    }
                }
            }
            // Create Border Elements
            {
                let borders = xml_doc_mut
                    .insert_children_after_tag_mut("borders", "fills", None)
                    .context("Create borders parents Failed.")?;
                let borders_id = borders.get_id();
                let mut attributes = HashMap::new();
                attributes.insert(
                    "count".to_string(),
                    self.border_collection.len().to_string(),
                );
                borders
                    .set_attribute_mut(attributes)
                    .context("Updating Number Formats Element Attributes Failed")?;
                for border_data in self.border_collection.as_slice() {
                    let border_id = xml_doc_mut
                        .append_child_mut("border", Some(&borders_id))
                        .context("Create Border Failed")?
                        .get_id();
                    // Left Border Setting
                    StylePart::add_border_element(
                        "left",
                        &mut xml_doc_mut,
                        &border_id,
                        border_data.left.clone(),
                    )?;
                    //Right Border Setting
                    StylePart::add_border_element(
                        "right",
                        &mut xml_doc_mut,
                        &border_id,
                        border_data.right.clone(),
                    )?;
                    // Top Border Setting
                    StylePart::add_border_element(
                        "top",
                        &mut xml_doc_mut,
                        &border_id,
                        border_data.top.clone(),
                    )?;
                    // Bottom Border Setting
                    StylePart::add_border_element(
                        "bottom",
                        &mut xml_doc_mut,
                        &border_id,
                        border_data.bottom.clone(),
                    )?;
                    // Diagonal Border Setting
                    StylePart::add_border_element(
                        "diagonal",
                        &mut xml_doc_mut,
                        &border_id,
                        border_data.diagonal.clone(),
                    )?;
                }
            }
            // Create Cell Style Elements
            {
                StylePart::add_cell_style(&mut xml_doc_mut, &mut self.cell_collection, true)?;
            }
            // Create Defined Cell Style Elements
            {
                StylePart::add_cell_style(
                    &mut xml_doc_mut,
                    &mut self.cell_style_collection,
                    false,
                )?;
            }
        }
        Ok(())
    }

    pub(crate) fn deserialize_cell_style(
        style_xfs: XmlElement,
        xml_doc_mut: &mut XmlDocument,
    ) -> AnyResult<Vec<CellXfs>, AnyError> {
        let mut style_records = Vec::new();
        loop {
            if let Some((xf_id, _)) = style_xfs.pop_child_mut() {
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
                            cell_xf.apply_protection =
                                ConverterUtil::normalize_bool_property(apply_protection);
                        }
                        if let Some(apply_alignment) = attributes.get("applyAlignment") {
                            cell_xf.apply_alignment =
                                ConverterUtil::normalize_bool_property(apply_alignment);
                        }
                        if let Some(apply_border) = attributes.get("applyBorder") {
                            cell_xf.apply_border =
                                ConverterUtil::normalize_bool_property(apply_border);
                        }
                        if let Some(apply_fill) = attributes.get("applyFill") {
                            cell_xf.apply_fill = ConverterUtil::normalize_bool_property(apply_fill);
                        }
                        if let Some(apply_font) = attributes.get("applyFont") {
                            cell_xf.apply_font = ConverterUtil::normalize_bool_property(apply_font);
                        }
                        if let Some(apply_number_format) = attributes.get("applyNumberFormat") {
                            cell_xf.apply_number_format =
                                ConverterUtil::normalize_bool_property(apply_number_format);
                        }
                        // Load Alignment Values if exist
                        if let Some((alignment_id, _)) = current_element.pop_child_mut() {
                            if let Some(alignment_element) =
                                xml_doc_mut.pop_element_mut(&alignment_id)
                            {
                                if let Some(alignment_attributes) =
                                    alignment_element.get_attribute()
                                {
                                    if let Some(is_wrap_text) = alignment_attributes.get("wrapText")
                                    {
                                        cell_xf.is_wrap_text =
                                            ConverterUtil::normalize_bool_property(is_wrap_text);
                                    }
                                    if let Some(vertical_alignment) =
                                        alignment_attributes.get("vertical")
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
                    style_records.push(cell_xf);
                }
            } else {
                break;
            }
        }
        Ok(style_records)
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
                    if let Some((color_id, _)) = current_element.pop_child_mut() {
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

    /// Add Color Element Node To XML
    fn add_color_element(
        color_setting: Option<ColorSetting>,
        xml_doc_mut: &mut std::cell::RefMut<'_, XmlDocument>,
        parent_id: usize,
    ) -> Result<(), AnyError> {
        Ok(if let Some(border_color_setting) = color_setting {
            let colors = xml_doc_mut
                .append_child_mut("color", Some(&parent_id))
                .context("Create Color Element Failed")?;
            let mut color_attribute: HashMap<String, String> = HashMap::new();
            color_attribute.insert(
                ColorSettingTypeValues::get_string(border_color_setting.color_setting_type),
                border_color_setting.value,
            );
            colors
                .set_attribute_mut(color_attribute)
                .context("Setting Color Attribute Failed")?;
        })
    }

    /// Add Border Element Node to XML
    fn add_border_element(
        border_side: &str,
        xml_doc_mut: &mut std::cell::RefMut<'_, XmlDocument>,
        border_id: &usize,
        border_data: BorderSetting,
    ) -> Result<(), AnyError> {
        let border_direction = xml_doc_mut
            .append_child_mut(border_side, Some(&border_id))
            .context(format!("{} Border Element Creation Failed", border_side))?;
        let id = border_direction.get_id();
        if border_data.style != BorderStyleValues::None {
            let mut left_attribute: HashMap<String, String> = HashMap::new();
            left_attribute.insert(
                "style".to_string(),
                BorderStyleValues::get_string(border_data.style),
            );
            border_direction
                .set_attribute_mut(left_attribute)
                .context(format!("Set {} Attribute Failed", border_side))?;
            StylePart::add_color_element(border_data.border_color, xml_doc_mut, id)?;
        }
        Ok(())
    }

    fn add_cell_style(
        xml_doc_mut: &mut XmlDocument,
        cell_style_xfs_data: &mut Vec<CellXfs>,
        enable_format_id: bool,
    ) -> Result<(), AnyError> {
        let cell_style_xfs = xml_doc_mut
            .insert_children_after_tag_mut(
                if enable_format_id {
                    "cellXfs"
                } else {
                    "cellStyleXfs"
                },
                "borders",
                None,
            )
            .context("Create Cell Style parents Failed.")?;
        let cell_style_xfs_id = cell_style_xfs.get_id();
        let mut attributes = HashMap::new();
        attributes.insert("count".to_string(), cell_style_xfs_data.len().to_string());
        cell_style_xfs
            .set_attribute_mut(attributes)
            .context("Updating Number Formats Element Attributes Failed")?;
        for cell_style_xfs in cell_style_xfs_data {
            let xf = xml_doc_mut
                .append_child_mut("xf", Some(&cell_style_xfs_id))
                .context("Create Cell Style Config Failed")?;
            let xf_id = xf.get_id();
            let mut attributes: HashMap<String, String> = HashMap::new();
            attributes.insert(
                "numFmtId".to_string(),
                cell_style_xfs.number_format_id.to_string(),
            );
            attributes.insert("fontId".to_string(), cell_style_xfs.font_id.to_string());
            attributes.insert("fillId".to_string(), cell_style_xfs.fill_id.to_string());
            attributes.insert("borderId".to_string(), cell_style_xfs.border_id.to_string());
            if enable_format_id {
                attributes.insert("xfId".to_string(), cell_style_xfs.format_id.to_string());
            }
            if cell_style_xfs.apply_font > 0 {
                attributes.insert(
                    "applyFont".to_string(),
                    cell_style_xfs.apply_font.to_string(),
                );
            }
            if cell_style_xfs.apply_alignment > 0 {
                attributes.insert(
                    "applyAlignment".to_string(),
                    cell_style_xfs.apply_alignment.to_string(),
                );
            }
            if cell_style_xfs.apply_fill > 0 {
                attributes.insert(
                    "applyFill".to_string(),
                    cell_style_xfs.apply_fill.to_string(),
                );
            }
            if cell_style_xfs.apply_border > 0 {
                attributes.insert(
                    "applyBorder".to_string(),
                    cell_style_xfs.apply_border.to_string(),
                );
            }
            if cell_style_xfs.apply_number_format > 0 {
                attributes.insert(
                    "applyNumberFormat".to_string(),
                    cell_style_xfs.apply_number_format.to_string(),
                );
            }
            if cell_style_xfs.apply_protection > 0 {
                attributes.insert(
                    "applyProtection".to_string(),
                    cell_style_xfs.apply_protection.to_string(),
                );
            }
            xf.set_attribute_mut(attributes)
                .context("Setting Attributes Failed")?;
            if cell_style_xfs.is_wrap_text > 0
                || cell_style_xfs.vertical_alignment != VerticalAlignmentValues::None
                || cell_style_xfs.horizontal_alignment != HorizontalAlignmentValues::None
            {
                let alignment = xml_doc_mut
                    .append_child_mut("alignment", Some(&xf_id))
                    .context("Create Cell Alignment Style Config Failed")?;
                let mut alignment_attributes: HashMap<String, String> = HashMap::new();
                if cell_style_xfs.is_wrap_text > 0 {
                    alignment_attributes.insert(
                        "wrapText".to_string(),
                        cell_style_xfs.is_wrap_text.to_string(),
                    );
                }
                if cell_style_xfs.vertical_alignment != VerticalAlignmentValues::None {
                    alignment_attributes.insert(
                        "vertical".to_string(),
                        VerticalAlignmentValues::get_string(
                            cell_style_xfs.vertical_alignment.clone(),
                        ),
                    );
                }
                if cell_style_xfs.horizontal_alignment != HorizontalAlignmentValues::None {
                    alignment_attributes.insert(
                        "horizontal".to_string(),
                        HorizontalAlignmentValues::get_string(
                            cell_style_xfs.horizontal_alignment.clone(),
                        ),
                    );
                }
                alignment
                    .set_attribute_mut(alignment_attributes)
                    .context("Setting Alignment Attribute Failed")?;
            }
        }
        Ok(())
    }
}

// ################################## mut feature ########################
impl StylePart {
    fn generate_setting_hash<T: Hash>(&mut self, item: &T) -> u64 {
        let mut hasher = DefaultHasher::new();
        item.hash(&mut hasher);
        hasher.finish()
    }

    pub(crate) fn get_style_id_mut(
        &mut self,
        style_setting: StyleSetting,
    ) -> AnyResult<StyleId, AnyError> {
        let style_hash = self.generate_setting_hash(&style_setting);
        if let Some((_, id)) = self.cache_id.get_key_value(&style_hash) {
            Ok(StyleId::new(*id))
        } else {
            if let Some(office_document) = self.office_document.upgrade() {
                let office_doc = office_document
                    .try_borrow_mut()
                    .context("Failed to get Office Handle")?;
                if style_setting.number_format == NumberFormatValues::Custom {
                    // Get Number Format ID
                }
                // Get Font Style ID
                {}
                // Get Fill Style ID
                {}
                // Get Border Style ID
                {}
                // Get Cell Style xfs
                {}
                // Get Cell xfs
                {}
            }
            Ok(StyleId::new(0))
        }
    }
}
