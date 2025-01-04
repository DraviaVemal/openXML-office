use crate::global_2007::{models::HyperlinkProperties, traits::Enum};
use crate::spreadsheet_2007::models::StyleId;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CellDataType {
    /// Use if you want the package to auto detect best fit
    Auto,
    Number,
    Boolean,
    String,
    ShareString,
    InlineString,
    Error,
}

impl Enum<CellDataType> for CellDataType {
    fn get_string(input_enum: CellDataType) -> String {
        match input_enum {
            CellDataType::Boolean => "b".to_string(),
            CellDataType::String => "str".to_string(),
            CellDataType::ShareString => "s".to_string(),
            CellDataType::InlineString => "inlineStr".to_string(),
            CellDataType::Error => "e".to_string(),
            CellDataType::Number => "n".to_string(),
            _ => "a".to_string(),
        }
    }
    fn get_enum(input_string: &str) -> CellDataType {
        match input_string {
            "b" => CellDataType::Boolean,
            "str" => CellDataType::String,
            "s" => CellDataType::ShareString,
            "inlineStr" => CellDataType::InlineString,
            "n" => CellDataType::Number,
            "e" => CellDataType::Error,
            _ => CellDataType::Auto,
        }
    }
}

#[derive(Debug, Default, Clone)]
pub(crate) struct RowRecord {
    pub(crate) index: usize,
    pub(crate) hide: Option<bool>,
    pub(crate) span: Option<String>,
    pub(crate) height: Option<f32>,
    pub(crate) style_id: Option<StyleId>,
    pub(crate) thick_top: Option<bool>,
    pub(crate) thick_bottom: Option<bool>,
    pub(crate) group_level: Option<usize>,
    pub(crate) collapsed: Option<bool>,
    pub(crate) place_holder: Option<bool>,
}

#[derive(Debug, Default, Clone)]
pub(crate) struct CellRecord {
    pub(crate) row_index: usize,
    pub(crate) col_index: Option<usize>,
    pub(crate) style_id: Option<StyleId>,
    pub(crate) value: Option<String>,
    pub(crate) formula: Option<String>,
    pub(crate) data_type: Option<CellDataType>,
    pub(crate) metadata: Option<String>,
    pub(crate) place_holder: Option<bool>,
    pub(crate) comment_id: Option<usize>,
}

#[derive(Debug, Default)]
pub struct RowProperties {
    // Set Custom height for the row
    pub height: Option<f32>,
    // Hide The Specific Row
    pub style_id: Option<StyleId>,
    pub hidden: Option<bool>,
    pub thick_top: Option<bool>,
    pub thick_bottom: Option<bool>,
    // Column group to use with collapse expand
    pub(crate) group_level: Option<usize>,
    // Collapse the current column
    pub(crate) collapsed: Option<bool>,
    pub(crate) place_holder: Option<bool>,
    pub(crate) span: Option<String>,
}

#[derive(Debug)]
pub struct ColumnProperties {
    // Start Column index
    pub(crate) min: usize,
    // End Column Index
    pub(crate) max: usize,
    // width value
    pub width: Option<f32>,
    // hide the specific column
    pub hidden: Option<bool>,
    // Column level style setting
    pub style_id: Option<StyleId>,
    // Best fit/auto fit column
    pub best_fit: Option<bool>,
    // Column group to use with collapse expand
    pub(crate) group_level: usize,
    // Collapse the current column
    pub(crate) collapsed: Option<bool>,
}

impl Default for ColumnProperties {
    fn default() -> Self {
        Self {
            min: 1,
            max: 1,
            best_fit: None,
            collapsed: None,
            group_level: 0,
            hidden: None,
            style_id: None,
            width: None,
        }
    }
}

#[derive(Debug)]
pub struct ColumnCell {
    pub formula: Option<String>,
    pub value: Option<String>,
    pub data_type: CellDataType,
    pub hyperlink_properties: Option<HyperlinkProperties>,
    pub style_id: Option<StyleId>,
    // TODO: Future Items
    pub(crate) metadata: Option<String>,
    pub(crate) comment_id: Option<usize>,
    pub(crate) place_holder: Option<bool>,
}

impl Default for ColumnCell {
    fn default() -> Self {
        Self {
            formula: None,
            value: None,
            data_type: CellDataType::Auto,
            hyperlink_properties: None,
            style_id: None,
            metadata: None,
            comment_id: None,
            place_holder: None,
        }
    }
}
