use crate::global_2007::models::HyperlinkProperties;
use crate::global_2007::traits::Enum;
use crate::spreadsheet_2007::models::CellStyleSetting;

#[derive(Debug)]
pub enum CellDataType {
    DATE,
    NUMBER,
    BOOLEAN,
    STRING,
    SHARE_STRING,
    INLINE_STRING,
    ERROR,
    FORMULA,
}

impl Enum<CellDataType> for CellDataType {
    fn get_string(input_enum: CellDataType) -> String {
        match input_enum {
            CellDataType::DATE => "d".to_string(),
            CellDataType::NUMBER => "n".to_string(),
            CellDataType::BOOLEAN => "b".to_string(),
            CellDataType::STRING => "str".to_string(),
            CellDataType::SHARE_STRING => "s".to_string(),
            CellDataType::INLINE_STRING => "inlineStr".to_string(),
            CellDataType::ERROR => "e".to_string(),
            CellDataType::FORMULA => "".to_string(),
        }
    }
    fn get_enum(input_string: &str) -> CellDataType {
        match input_string {
            "d" => CellDataType::DATE,
            "n" => CellDataType::NUMBER,
            "b" => CellDataType::BOOLEAN,
            "str" => CellDataType::STRING,
            "s" => CellDataType::SHARE_STRING,
            "inlineStr" => CellDataType::INLINE_STRING,
            "e" => CellDataType::ERROR,
            _ => CellDataType::FORMULA,
        }
    }
}

#[derive(Debug, Default)]
pub struct ColumnProperties {
    best_fit: bool,
    hidden: bool,
    width: Option<usize>,
}

#[derive(Debug)]
pub struct ColumnCell {
    value: String,
    data_type: CellDataType,
    hyperlink_properties: HyperlinkProperties,
    style_setting: CellStyleSetting,
    style_id: Option<usize>,
}

#[derive(Debug)]
pub struct RowProperties {
    height: Option<usize>,
    hidden: bool,
}
