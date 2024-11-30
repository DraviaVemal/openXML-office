pub enum HorizontalAlignmentValues {
    NONE,
    LEFT,
    CENTER,
    RIGHT,
    JUSTIFY,
}
pub enum VerticalAlignmentValues {
    NONE,
    TOP,
    MIDDLE,
    BOTTOM,
}
pub enum ColorSettingTypeValues {
    INDEXED,
    THEME,
    RGB,
}
pub enum StyleValues {
    NONE,
    THIN,
    THICK,
    DOTTED,
    DOUBLE,
    DASHED,
    DASH_DOT,
    DASH_DOT_DOT,
    MEDIUM,
    MEDIUM_DASHED,
    MEDIUM_DASH_DOT,
    MEDIUM_DASH_DOT_DOT,
    SLANT_DASH_DOT,
    HAIR,
}

pub struct ColorSetting {
    color_setting_type_values: ColorSettingTypeValues,
    value: String,
}

pub struct BorderSetting {
    border_color: ColorSetting,
    style: StyleValues,
}

pub struct CellStyleSetting {
    background_color: String,
    border_bottom: BorderSetting,
    border_left: BorderSetting,
    border_right: BorderSetting,
    border_top: BorderSetting,
    font_family: String,
    font_size: usize,
    foreground_color: String,
    is_bold: bool,
    is_double_underline: bool,
    is_italic: bool,
    is_underline: bool,
    is_wrap_text: bool,
    number_format: String,
    text_color: ColorSetting,
    horizontal_alignment: HorizontalAlignmentValues,
    vertical_alignment: VerticalAlignmentValues,
}
