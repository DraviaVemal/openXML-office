use crate::global_2007::models::HyperlinkProperties;
use crate::spreadsheet_2007::models::CellStyleSetting;

pub enum CellDataType {
    DATE,
    NUMBER,
    STRING,
    FORMULA,
}
pub struct ColumnProperties {
    best_fit: bool,
    hidden: bool,
    width: Option<usize>,
}
pub struct ColumnCell {
    value: String,
    data_type: CellDataType,
    hyperlink_properties: HyperlinkProperties,
    style_setting: CellStyleSetting,
    style_id: Option<usize>,
}
pub struct RowProperties {
    height: Option<usize>,
    hidden: bool,
}
