use crate::{
    converters::ConverterUtil,
    element_dictionary::EXCEL_TYPE_COLLECTION,
    files::{office_document, OfficeDocument, XmlDocument, XmlElement, XmlSerializer},
    get_all_queries,
    global_2007::{
        parts::RelationsPart,
        traits::{Enum, XmlDocumentPart, XmlDocumentPartCommon},
    },
    spreadsheet_2007::models::{
        BorderSetting, BorderStyle, BorderStyleValues, CellXfs, ColorSetting,
        ColorSettingTypeValues, FillStyle, FontSchemeValues, FontStyle, HorizontalAlignmentValues,
        NumberFormat, PatternTypeValues, StyleId, StyleSetting, VerticalAlignmentValues,
    },
};
use anyhow::{anyhow, Context, Error as AnyError, Result as AnyResult};
use rusqlite::{params, Row, ToSql};
use serde_json::{from_str, to_string};
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
    cache_id: HashMap<u64, usize>,
    cache_order: VecDeque<usize>,
    cache_capacity: usize,
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
        if let Some(xml_tree) = self.office_document.upgrade() {
            Self::save_content_to_tree(&self.office_document, &mut self.xml_document)
                .context("Style Save Content Failed")?;
            xml_tree
                .try_borrow_mut()
                .context("Failed To pull XML Handle")?
                .close_xml_document(&self.file_path)?;
        }
        Ok(())
    }
    /// Initialize xml content for this part from base template
    fn initialize_content_xml(
    ) -> AnyResult<(XmlDocument, Option<String>, Option<String>, Option<String>), AnyError> {
        let content = EXCEL_TYPE_COLLECTION.get("style").unwrap();
        Ok((
            XmlSerializer::vec_to_xml_doc_tree(include_str!("style.xml").as_bytes().to_vec())
                .context("Initializing Theme Failed")?,
            Some(content.content_type.to_string()),
            Some(content.extension.to_string()),
            Some(content.extension_type.to_string()),
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
        Self::load_content_to_database(&office_document, &mut xml_document)
            .context("Load Share String To DB Failed")?;
        Ok(Self {
            office_document,
            xml_document,
            file_path: file_name,
            cache_id: HashMap::new(),
            cache_order: VecDeque::new(),
            cache_capacity: 25,
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

    /// Create the new Table Required to Load the data
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
                .get_database()
                .create_table(&create_query_num_format, None)
                .context("Create Number Format Table Failed")?;
            office_doc
                .get_database()
                .create_table(&create_query_font_style, None)
                .context("Create Font Style Table Failed")?;
            office_doc
                .get_database()
                .create_table(&create_query_fill_style, None)
                .context("Create Query Fill Table Failed")?;
            office_doc
                .get_database()
                .create_table(&create_query_border_style, None)
                .context("Create Border Style Table Failed")?;
            office_doc
                .get_database()
                .create_table(&create_query_cell_style_xfs, None)
                .context("Create Cell Style Table Failed")?;
            office_doc
                .get_database()
                .create_table(&create_query_cell_xfs, None)
                .context("Create Cell Style Table Failed")?;
        }
        Ok(())
    }

    /// Load existing file style to database
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
            if let Some(xml_document) = xml_document.upgrade() {
                let mut xml_doc_mut = xml_document
                    .try_borrow_mut()
                    .context("xml doc borrow failed")?;
                // Load Number Format Region
                if let Some(mut number_formats_vec) =
                    xml_doc_mut.pop_elements_by_tag_mut("numFmts", None)
                {
                    let insert_query_num_format = queries
                        .get("insert_number_format_table")
                        .ok_or_else(|| anyhow!("Expected Query Not Found"))?;
                    let mut num_format_records = Vec::new();
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
                    fn row_parser(number_format: NumberFormat) -> Vec<Box<dyn ToSql>> {
                        vec![
                            Box::new(number_format.format_id),
                            Box::new(number_format.format_code),
                        ]
                    }
                    office_doc
                        .get_database()
                        .insert_records(
                            &insert_query_num_format,
                            num_format_records,
                            row_parser,
                            None,
                        )
                        .context("Number Format Data Insert Failed")?;
                }
                if let Some(mut fonts_vec) = xml_doc_mut.pop_elements_by_tag_mut("fonts", None) {
                    let mut font_records = Vec::new();
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
                                                        font_style.color.value =
                                                            theme.parse().context(
                                                                "Font color theme parse failed",
                                                            )?
                                                    } else if let Some(rgb) = attributes.get("rgb")
                                                    {
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
                    // Insert Data into Database
                    let insert_query_font_style = queries
                        .get("insert_font_style_table")
                        .ok_or_else(|| anyhow!("Expected Query Not Found"))?;
                    fn row_parser(font_style: FontStyle) -> Vec<Box<dyn ToSql>> {
                        vec![
                            Box::new(font_style.name),
                            Box::new(ColorSettingTypeValues::get_string(
                                font_style.color.color_setting_type,
                            )),
                            Box::new(font_style.color.value),
                            Box::new(font_style.family),
                            Box::new(font_style.size),
                            Box::new(FontSchemeValues::get_string(font_style.font_scheme)),
                            Box::new(font_style.is_bold),
                            Box::new(font_style.is_italic),
                            Box::new(font_style.is_underline),
                            Box::new(font_style.is_double_underline),
                        ]
                    }
                    office_doc
                        .get_database()
                        .insert_records(&insert_query_font_style, font_records, row_parser, None)
                        .context("Insert Font Style Failed")?;
                }
                if let Some(mut fills_vec) = xml_doc_mut.pop_elements_by_tag_mut("fills", None) {
                    let mut fill_records = Vec::new();
                    if let Some(fills) = fills_vec.pop() {
                        loop {
                            if let Some((fill_id, _)) = fills.pop_child_mut() {
                                let current_element = xml_doc_mut
                                    .pop_element_mut(&fill_id)
                                    .ok_or(anyhow!("Failed to pull child element"))?;
                                let mut fill_style = FillStyle::default();
                                if let Some((pattern_fill_id, _)) = current_element.pop_child_mut()
                                {
                                    if let Some(pattern_fill) =
                                        xml_doc_mut.pop_element_mut(&pattern_fill_id)
                                    {
                                        if let Some(pattern_attributes) =
                                            pattern_fill.get_attribute()
                                        {
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
                                                                        } else if let Some(
                                                                            indexed,
                                                                        ) = attributes
                                                                            .get("indexed")
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
                                                                        } else if let Some(
                                                                            indexed,
                                                                        ) = attributes
                                                                            .get("indexed")
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
                                fill_records.push((
                                    to_string(&fill_style.background_color)
                                        .context("Background Fill Parse Failed")?,
                                    to_string(&fill_style.foreground_color)
                                        .context("Foreground Fill Parse Failed")?,
                                    PatternTypeValues::get_string(fill_style.pattern_type),
                                ));
                            } else {
                                break;
                            }
                        }
                    }
                    // Insert Tables Queries
                    let insert_query_fill_style = queries
                        .get("insert_fill_style_table")
                        .ok_or_else(|| anyhow!("Expected Query Not Found"))?;
                    fn row_parser(fill_style: (String, String, String)) -> Vec<Box<dyn ToSql>> {
                        vec![
                            Box::new(fill_style.0),
                            Box::new(fill_style.1),
                            Box::new(fill_style.2),
                        ]
                    }
                    office_doc
                        .get_database()
                        .insert_records(&insert_query_fill_style, fill_records, row_parser, None)
                        .context("Insert Fill Style Failed")?;
                }
                if let Some(mut borders_vec) = xml_doc_mut.pop_elements_by_tag_mut("borders", None)
                {
                    let mut border_records = Vec::new();
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
                                    border_records.push((
                                        to_string(&border_style.left)
                                            .context("Left Border Parsing Failed")?,
                                        to_string(&border_style.top)
                                            .context("Top Border Parsing Failed")?,
                                        to_string(&border_style.right)
                                            .context("Right Border Parsing Failed")?,
                                        to_string(&border_style.bottom)
                                            .context("Bottom Border Parsing Failed")?,
                                        to_string(&border_style.diagonal)
                                            .context("diagonal Border Parsing Failed")?,
                                    ));
                                }
                            } else {
                                break;
                            }
                        }
                    }
                    let insert_query_border_style = queries
                        .get("insert_border_style_table")
                        .ok_or_else(|| anyhow!("Expected Query Not Found"))?;
                    fn row_parser(
                        fill_style: (String, String, String, String, String),
                    ) -> Vec<Box<dyn ToSql>> {
                        vec![
                            Box::new(fill_style.0),
                            Box::new(fill_style.1),
                            Box::new(fill_style.2),
                            Box::new(fill_style.3),
                            Box::new(fill_style.4),
                        ]
                    }
                    office_doc
                        .get_database()
                        .insert_records(
                            &insert_query_border_style,
                            border_records,
                            row_parser,
                            None,
                        )
                        .context("Insert border Style Failed")?;
                }
                if let Some(mut cell_style_xfs_vec) =
                    xml_doc_mut.pop_elements_by_tag_mut("cellStyleXfs", None)
                {
                    if let Some(cell_style_xfs) = cell_style_xfs_vec.pop() {
                        StylePart::deserialize_cell_style(
                            cell_style_xfs,
                            &mut xml_doc_mut,
                            &office_doc,
                            &queries,
                            "insert_cell_style_xfs_table",
                        )
                        .context("Deserializing Cell Style Xfs Failed")?;
                    }
                }
                if let Some(mut cell_xfs_vec) = xml_doc_mut.pop_elements_by_tag_mut("cellXfs", None)
                {
                    if let Some(cell_xfs) = cell_xfs_vec.pop() {
                        StylePart::deserialize_cell_style(
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
        }
        Ok(())
    }

    /// Save Database record back to XML File
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
            if let Some(xml_document) = xml_document.upgrade() {
                let mut xml_doc_mut = xml_document
                    .try_borrow_mut()
                    .context("xml doc borrow failed")?;
                // Create Number Formats Elements
                {
                    let num_formats = xml_doc_mut
                        .insert_child_at_mut("numFmts", &0, None)
                        .context("Create Number Formats Parent Failed.")?;
                    let num_formats_id = num_formats.get_id();
                    if let Some(num_format_query) = queries.get("select_number_formats_table") {
                        fn row_mapper(row: &Row) -> AnyResult<NumberFormat, rusqlite::Error> {
                            Ok(NumberFormat {
                                id: 0,
                                format_id: row.get(0)?,
                                format_code: row.get(1)?,
                            })
                        }
                        let num_format_data = office_doc
                            .get_database()
                            .find_many(num_format_query, params![], row_mapper, None)
                            .context("Num Format Query Results Failed")?;
                        let mut attributes = HashMap::new();
                        attributes.insert("count".to_string(), num_format_data.len().to_string());
                        num_formats
                            .set_attribute_mut(attributes)
                            .context("Updating Number Formats Element Attributes Failed")?;
                        for num_format in num_format_data {
                            let num_format_element = xml_doc_mut
                                .append_child_mut("numFmt", Some(&num_formats_id))
                                .context("Create Number Format Element Failed")?;
                            let mut attributes = HashMap::new();
                            attributes
                                .insert("numFmtId".to_string(), num_format.format_id.to_string());
                            attributes.insert("formatCode".to_string(), num_format.format_code);
                            num_format_element
                                .set_attribute_mut(attributes)
                                .context("Updating Number Format Element Attributes Failed")?;
                        }
                    }
                }
                // Create Fonts Elements
                {
                    let fonts = xml_doc_mut
                        .insert_children_after_tag_mut("fonts", "numFmts", None)
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
                                    color_setting_type: ColorSettingTypeValues::get_enum(
                                        &color_type,
                                    ),
                                    value: row.get(2)?,
                                },
                                family: row.get(3)?,
                                size: row.get(4)?,
                                font_scheme: FontSchemeValues::get_enum(&font_scheme),
                                is_bold: row.get(6)?,
                                is_italic: row.get(7)?,
                                is_underline: row.get(8)?,
                                is_double_underline: row.get(9)?,
                            })
                        }
                        let fonts_data = office_doc
                            .get_database()
                            .find_many(fonts_query, params![], row_mapper, None)
                            .context("Fonts Query Results Failed")?;
                        let mut attributes = HashMap::new();
                        attributes.insert("count".to_string(), fonts_data.len().to_string());
                        fonts
                            .set_attribute_mut(attributes)
                            .context("Set attribute failed for fonts style")?;
                        for font_style in fonts_data {
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
                                double_underline_attributes
                                    .insert("val".to_string(), "double".to_string());
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
                                Some(font_style.color),
                                &mut xml_doc_mut,
                                font_id,
                            )?;
                            let name = xml_doc_mut
                                .append_child_mut("name", Some(&font_id))
                                .context("Create Name Failed")?;
                            let mut name_attributes: HashMap<String, String> = HashMap::new();
                            name_attributes.insert("val".to_string(), font_style.name);
                            name.set_attribute_mut(name_attributes)
                                .context("Setting name Attribute Failing")?;
                            let family = xml_doc_mut
                                .append_child_mut("family", Some(&font_id))
                                .context("Create Name Failed")?;
                            let mut family_attributes: HashMap<String, String> = HashMap::new();
                            family_attributes
                                .insert("val".to_string(), font_style.family.to_string());
                            family
                                .set_attribute_mut(family_attributes)
                                .context("Setting family attribute failing")?;
                            let scheme = xml_doc_mut
                                .append_child_mut("scheme", Some(&font_id))
                                .context("Create scheme Failed")?;
                            let mut scheme_attributes: HashMap<String, String> = HashMap::new();
                            scheme_attributes.insert(
                                "val".to_string(),
                                FontSchemeValues::get_string(font_style.font_scheme),
                            );
                            scheme
                                .set_attribute_mut(scheme_attributes)
                                .context("Setting scheme attribute failing")?;
                        }
                    }
                }
                // Create Fills Elements
                {
                    let fills = xml_doc_mut
                        .insert_children_after_tag_mut("fills", "fonts", None)
                        .context("Create Fills parents Failed.")?;
                    let fills_id = fills.get_id();
                    if let Some(fills_query) = queries.get("select_fill_style_table") {
                        fn row_mapper(row: &Row) -> AnyResult<FillStyle, rusqlite::Error> {
                            let mut fill_style = FillStyle::default();
                            let background_color: String = row.get(0)?;
                            let foreground_color: String = row.get(1)?;
                            let pattern_type: String = row.get(2)?;
                            fill_style.background_color =
                                from_str(&background_color).map_err(|err| {
                                    rusqlite::Error::InvalidColumnName(
                                        format!("Column : {} Parsing Failed", err.column())
                                            .to_string(),
                                    )
                                })?;
                            fill_style.foreground_color =
                                from_str(&foreground_color).map_err(|err| {
                                    rusqlite::Error::InvalidColumnName(
                                        format!("Column : {} Parsing Failed", err.column())
                                            .to_string(),
                                    )
                                })?;
                            fill_style.pattern_type = PatternTypeValues::get_enum(&pattern_type);
                            Ok(fill_style)
                        }
                        let fills_data = office_doc
                            .get_database()
                            .find_many(fills_query, params![], row_mapper, None)
                            .context("Fills Data Pull Failed")?;
                        // .context("Fills Query Results Failed")?;
                        let mut attributes = HashMap::new();
                        attributes.insert("count".to_string(), fills_data.len().to_string());
                        fills
                            .set_attribute_mut(attributes)
                            .context("Set Fill Attribute Failed")?;
                        for fill_data in fills_data {
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
                                PatternTypeValues::get_string(fill_data.pattern_type),
                            );
                            pattern_fill_element
                                .set_attribute_mut(pattern_attribute)
                                .context("Set Pattern Fill Attribute Failed")?;
                            let pattern_fill_id = pattern_fill_element.get_id();
                            if let Some(fg_setting) = fill_data.foreground_color {
                                let fg_element = xml_doc_mut
                                    .append_child_mut("fgColor", Some(&pattern_fill_id))
                                    .context("Pattern Fill Foreground Element Failed")?;
                                let mut fg_attributes: HashMap<String, String> = HashMap::new();
                                fg_attributes.insert(
                                    ColorSettingTypeValues::get_string(
                                        fg_setting.color_setting_type,
                                    ),
                                    fg_setting.value,
                                );
                                fg_element
                                    .set_attribute_mut(fg_attributes)
                                    .context("Set Foreground attribute Failed")?;
                            }
                            if let Some(bg_setting) = fill_data.background_color {
                                let bg_element = xml_doc_mut
                                    .append_child_mut("bgColor", Some(&pattern_fill_id))
                                    .context("Pattern Fill Background Element Failed")?;
                                let mut bg_attributes: HashMap<String, String> = HashMap::new();
                                bg_attributes.insert(
                                    ColorSettingTypeValues::get_string(
                                        bg_setting.color_setting_type,
                                    ),
                                    bg_setting.value,
                                );
                                bg_element
                                    .set_attribute_mut(bg_attributes)
                                    .context("Set Background attribute Failed")?;
                            }
                        }
                    }
                }
                // Create Border Elements
                {
                    let borders = xml_doc_mut
                        .insert_children_after_tag_mut("borders", "fills", None)
                        .context("Create borders parents Failed.")?;
                    let borders_id = borders.get_id();
                    if let Some(borders_query) = queries.get("select_border_style_table") {
                        fn row_mapper(row: &Row) -> AnyResult<BorderStyle, rusqlite::Error> {
                            let mut border_style = BorderStyle::default();
                            let left_border: String = row.get(0)?;
                            let top_border: String = row.get(1)?;
                            let right_border: String = row.get(2)?;
                            let bottom_border: String = row.get(3)?;
                            let diagonal_border: String = row.get(4)?;
                            border_style.left = from_str(&left_border).map_err(|_| {
                                rusqlite::Error::InvalidColumnName(
                                    "Left border parsing failed".to_string(),
                                )
                            })?;
                            border_style.top = from_str(&top_border).map_err(|_| {
                                rusqlite::Error::InvalidColumnName(
                                    "Top border parsing failed".to_string(),
                                )
                            })?;
                            border_style.right = from_str(&right_border).map_err(|_| {
                                rusqlite::Error::InvalidColumnName(
                                    "Right border parsing failed".to_string(),
                                )
                            })?;
                            border_style.bottom = from_str(&bottom_border).map_err(|_| {
                                rusqlite::Error::InvalidColumnName(
                                    "Bottom border parsing failed".to_string(),
                                )
                            })?;
                            border_style.diagonal = from_str(&diagonal_border).map_err(|_| {
                                rusqlite::Error::InvalidColumnName(
                                    "Diagonal border parsing failed".to_string(),
                                )
                            })?;
                            Ok(border_style)
                        }
                        let borders_data = office_doc
                            .get_database()
                            .find_many(borders_query, params![], row_mapper, None)
                            .context("Border Style Query Results Failed")?;
                        let mut attributes = HashMap::new();
                        attributes.insert("count".to_string(), borders_data.len().to_string());
                        borders
                            .set_attribute_mut(attributes)
                            .context("Updating Number Formats Element Attributes Failed")?;
                        for border_data in borders_data {
                            let border_id = xml_doc_mut
                                .append_child_mut("border", Some(&borders_id))
                                .context("Create Border Failed")?
                                .get_id();
                            // Left Border Setting
                            StylePart::add_border_element(
                                "left",
                                &mut xml_doc_mut,
                                &border_id,
                                border_data.left,
                            )?;
                            //Right Border Setting
                            StylePart::add_border_element(
                                "right",
                                &mut xml_doc_mut,
                                &border_id,
                                border_data.right,
                            )?;
                            // Top Border Setting
                            StylePart::add_border_element(
                                "top",
                                &mut xml_doc_mut,
                                &border_id,
                                border_data.top,
                            )?;
                            // Bottom Border Setting
                            StylePart::add_border_element(
                                "bottom",
                                &mut xml_doc_mut,
                                &border_id,
                                border_data.bottom,
                            )?;
                            // Diagonal Border Setting
                            StylePart::add_border_element(
                                "diagonal",
                                &mut xml_doc_mut,
                                &border_id,
                                border_data.diagonal,
                            )?;
                        }
                    }
                }
                // Create Cell Style Elements
                {
                    StylePart::add_cell_style(
                        "cellXfs",
                        "select_cell_xfs_table",
                        &mut xml_doc_mut,
                        &queries,
                        &office_doc,
                        true,
                    )?;
                }
                // Create Defined Cell Style Elements
                {
                    StylePart::add_cell_style(
                        "cellStyleXfs",
                        "select_cell_style_xfs_table",
                        &mut xml_doc_mut,
                        &queries,
                        &office_doc,
                        false,
                    )?;
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
        let insert_query_style_xfs = queries
            .get(query_target)
            .ok_or_else(|| anyhow!("Expected Query Not Found"))?;
        fn row_parser(cell_xf: CellXfs) -> Vec<Box<dyn ToSql>> {
            vec![
                Box::new(cell_xf.format_id),
                Box::new(cell_xf.number_format_id),
                Box::new(cell_xf.font_id),
                Box::new(cell_xf.fill_id),
                Box::new(cell_xf.border_id),
                Box::new(cell_xf.apply_font),
                Box::new(cell_xf.apply_alignment),
                Box::new(cell_xf.apply_fill),
                Box::new(cell_xf.apply_border),
                Box::new(cell_xf.apply_number_format),
                Box::new(cell_xf.apply_protection),
                Box::new(cell_xf.is_wrap_text),
                Box::new(HorizontalAlignmentValues::get_string(
                    cell_xf.horizontal_alignment,
                )),
                Box::new(VerticalAlignmentValues::get_string(
                    cell_xf.vertical_alignment,
                )),
            ]
        }
        office_doc
            .get_database()
            .insert_records(&insert_query_style_xfs, style_records, row_parser, None)
            .context("Insert border Style Failed")?;
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
        parent_tag: &str,
        query_id: &str,
        xml_doc_mut: &mut std::cell::RefMut<'_, XmlDocument>,
        queries: &HashMap<String, String>,
        office_doc: &std::cell::Ref<'_, OfficeDocument>,
        enable_format_id: bool,
    ) -> Result<(), AnyError> {
        let cell_style_xfs = xml_doc_mut
            .insert_children_after_tag_mut(parent_tag, "borders", None)
            .context("Create Cell Style parents Failed.")?;
        let cell_style_xfs_id = cell_style_xfs.get_id();
        Ok(if let Some(cell_style_xfs_query) = queries.get(query_id) {
            fn row_mapper(row: &Row) -> AnyResult<CellXfs, rusqlite::Error> {
                let horizontal_alignment: String = row.get(12)?;
                let vertical_alignment: String = row.get(13)?;
                let cell_style = CellXfs {
                    id: 0,
                    format_id: row.get(0)?,
                    number_format_id: row.get(1)?,
                    font_id: row.get(2)?,
                    fill_id: row.get(3)?,
                    border_id: row.get(4)?,
                    apply_font: row.get(5)?,
                    apply_alignment: row.get(6)?,
                    apply_fill: row.get(7)?,
                    apply_border: row.get(8)?,
                    apply_number_format: row.get(9)?,
                    apply_protection: row.get(10)?,
                    is_wrap_text: row.get(11)?,
                    horizontal_alignment: HorizontalAlignmentValues::get_enum(
                        &horizontal_alignment,
                    ),
                    vertical_alignment: VerticalAlignmentValues::get_enum(&vertical_alignment),
                };
                Ok(cell_style)
            }
            let cell_style_xfs_data = office_doc
                .get_database()
                .find_many(cell_style_xfs_query, params![], row_mapper, None)
                .context("Num Format Query Results Failed")?;
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
                            VerticalAlignmentValues::get_string(cell_style_xfs.vertical_alignment),
                        );
                    }
                    if cell_style_xfs.horizontal_alignment != HorizontalAlignmentValues::None {
                        alignment_attributes.insert(
                            "horizontal".to_string(),
                            HorizontalAlignmentValues::get_string(
                                cell_style_xfs.horizontal_alignment,
                            ),
                        );
                    }
                    alignment
                        .set_attribute_mut(alignment_attributes)
                        .context("Setting Alignment Attribute Failed")?;
                }
            }
        })
    }
}

// ################################## mut feature ########################
impl StylePart {
    fn generate_setting_hash<T: Hash>(&mut self, item: &T) -> u64 {
        let mut hasher = DefaultHasher::new();
        item.hash(&mut hasher);
        hasher.finish()
    }

    pub(crate) fn 
    get_style_id_mut(
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
                // Get Number Format ID
                {
                    
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
