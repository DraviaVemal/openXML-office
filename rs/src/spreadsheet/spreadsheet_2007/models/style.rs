use crate::global_2007::traits::Enum;
use serde::{Deserialize, Serialize};

#[derive(Debug, Deserialize, Serialize)]
pub enum SchemeValues {
    None,
    Minor,
    Major,
}

impl Enum<SchemeValues> for SchemeValues {
    fn get_string(input_enum: SchemeValues) -> String {
        match input_enum {
            SchemeValues::Major => "major".to_string(),
            SchemeValues::Minor => "minor".to_string(),
            SchemeValues::None => "none".to_string(),
        }
    }
    fn get_enum(input_string: &str) -> SchemeValues {
        match input_string {
            "major" => SchemeValues::Major,
            "minor" => SchemeValues::Minor,
            _ => SchemeValues::None,
        }
    }
}

#[derive(Debug, Deserialize, Serialize)]
pub enum PatternTypeValues {
    None,
    Gray125,
    Solid,
}

impl Enum<PatternTypeValues> for PatternTypeValues {
    fn get_string(input_enum: PatternTypeValues) -> String {
        match input_enum {
            PatternTypeValues::Solid => "solid".to_string(),
            PatternTypeValues::Gray125 => "gray125".to_string(),
            PatternTypeValues::None => "none".to_string(),
        }
    }

    fn get_enum(input_string: &str) -> PatternTypeValues {
        match input_string {
            "solid" => PatternTypeValues::Solid,
            "gray125" => PatternTypeValues::Gray125,
            _ => PatternTypeValues::None,
        }
    }
}

#[derive(Debug, Deserialize, Serialize, PartialEq)]
pub enum BorderStyleValues {
    None,
    Thin,
    Thick,
    Dotted,
    Double,
    Dashed,
    DashDot,
    DashDotDot,
    Medium,
    MediumDashed,
    MediumDashDot,
    MediumDashDotDot,
    SlantDashDot,
    Hair,
}
impl Enum<BorderStyleValues> for BorderStyleValues {
    fn get_string(input_enum: BorderStyleValues) -> String {
        match input_enum {
            BorderStyleValues::DashDot => "dashDot".to_string(),
            BorderStyleValues::DashDotDot => "dashDotDot".to_string(),
            BorderStyleValues::Dashed => "dashed".to_string(),
            BorderStyleValues::Dotted => "dotted".to_string(),
            BorderStyleValues::Double => "double".to_string(),
            BorderStyleValues::Hair => "hair".to_string(),
            BorderStyleValues::Medium => "medium".to_string(),
            BorderStyleValues::MediumDashDot => "mediumDashDot".to_string(),
            BorderStyleValues::MediumDashDotDot => "mediumDashDotDot".to_string(),
            BorderStyleValues::MediumDashed => "mediumDashed".to_string(),
            BorderStyleValues::SlantDashDot => "slantDashDot".to_string(),
            BorderStyleValues::Thick => "thick".to_string(),
            BorderStyleValues::Thin => "thin".to_string(),
            BorderStyleValues::None => "none".to_string(),
        }
    }

    fn get_enum(input_string: &str) -> BorderStyleValues {
        match input_string {
            "dashDot" => BorderStyleValues::DashDot,
            "dashDotDot" => BorderStyleValues::DashDotDot,
            "dashed" => BorderStyleValues::Dashed,
            "dotted" => BorderStyleValues::Dotted,
            "double" => BorderStyleValues::Double,
            "hair" => BorderStyleValues::Hair,
            "medium" => BorderStyleValues::Medium,
            "mediumDashDot" => BorderStyleValues::MediumDashDot,
            "mediumDashDotDot" => BorderStyleValues::MediumDashDotDot,
            "mediumDashed" => BorderStyleValues::MediumDashed,
            "slantDashDot" => BorderStyleValues::SlantDashDot,
            "thick" => BorderStyleValues::Thick,
            "thin" => BorderStyleValues::Thin,
            _ => BorderStyleValues::None,
        }
    }
}

#[derive(Debug, Deserialize, Serialize)]
pub enum ColorSettingTypeValues {
    Indexed,
    Theme,
    Rgb,
}

impl Enum<ColorSettingTypeValues> for ColorSettingTypeValues {
    fn get_string(input_enum: ColorSettingTypeValues) -> String {
        match input_enum {
            ColorSettingTypeValues::Theme => "theme".to_string(),
            ColorSettingTypeValues::Rgb => "rgb".to_string(),
            ColorSettingTypeValues::Indexed => "indexed".to_string(),
        }
    }
    fn get_enum(input_string: &str) -> ColorSettingTypeValues {
        match input_string {
            "theme" => ColorSettingTypeValues::Theme,
            "rgb" => ColorSettingTypeValues::Rgb,
            _ => ColorSettingTypeValues::Indexed,
        }
    }
}

#[derive(Debug, Deserialize, Serialize)]
pub struct ColorSetting {
    pub color_setting_type: ColorSettingTypeValues,
    pub value: String,
}

impl Default for ColorSetting {
    fn default() -> Self {
        Self {
            color_setting_type: ColorSettingTypeValues::Theme,
            value: "1".to_string(),
        }
    }
}

#[derive(Debug, Deserialize, Serialize)]
pub struct BorderSetting {
    pub border_color: Option<ColorSetting>,
    pub style: BorderStyleValues,
}

impl Default for BorderSetting {
    fn default() -> Self {
        Self {
            border_color: None,
            style: BorderStyleValues::None,
        }
    }
}

#[derive(Debug, Deserialize, Serialize)]
pub struct BorderStyle {
    pub id: u32,
    pub bottom: BorderSetting,
    pub left: BorderSetting,
    pub right: BorderSetting,
    pub top: BorderSetting,
}

impl Default for BorderStyle {
    fn default() -> Self {
        Self {
            id: 0,
            bottom: BorderSetting::default(),
            left: BorderSetting::default(),
            right: BorderSetting::default(),
            top: BorderSetting::default(),
        }
    }
}

#[derive(Debug, Deserialize, Serialize)]
pub enum HorizontalAlignmentValues {
    None,
    LEFT,
    CENTER,
    RIGHT,
    JUSTIFY,
}

impl Enum<HorizontalAlignmentValues> for HorizontalAlignmentValues {
    fn get_string(input_enum: HorizontalAlignmentValues) -> String {
        match input_enum {
            HorizontalAlignmentValues::LEFT => "left".to_string(),
            HorizontalAlignmentValues::CENTER => "center".to_string(),
            HorizontalAlignmentValues::RIGHT => "right".to_string(),
            HorizontalAlignmentValues::JUSTIFY => "justify".to_string(),
            HorizontalAlignmentValues::None => "none".to_string(),
        }
    }
    fn get_enum(input_string: &str) -> HorizontalAlignmentValues {
        match input_string {
            "left" => HorizontalAlignmentValues::LEFT,
            "center" => HorizontalAlignmentValues::CENTER,
            "right" => HorizontalAlignmentValues::RIGHT,
            "justify" => HorizontalAlignmentValues::JUSTIFY,
            _ => HorizontalAlignmentValues::None,
        }
    }
}

#[derive(Debug, Deserialize, Serialize)]
pub enum VerticalAlignmentValues {
    None,
    TOP,
    MIDDLE,
    BOTTOM,
}

impl Enum<VerticalAlignmentValues> for VerticalAlignmentValues {
    fn get_string(input_enum: VerticalAlignmentValues) -> String {
        match input_enum {
            VerticalAlignmentValues::TOP => "top".to_string(),
            VerticalAlignmentValues::MIDDLE => "center".to_string(),
            VerticalAlignmentValues::BOTTOM => "bottom".to_string(),
            VerticalAlignmentValues::None => "none".to_string(),
        }
    }
    fn get_enum(input_string: &str) -> VerticalAlignmentValues {
        match input_string {
            "top" => VerticalAlignmentValues::TOP,
            "center" => VerticalAlignmentValues::MIDDLE,
            "bottom" => VerticalAlignmentValues::BOTTOM,
            _ => VerticalAlignmentValues::None,
        }
    }
}

#[derive(Debug, Deserialize, Serialize)]
pub struct CellStyleSetting {
    pub background_color: Option<String>,
    pub border_bottom: BorderSetting,
    pub border_left: BorderSetting,
    pub border_right: BorderSetting,
    pub border_top: BorderSetting,
    pub font_family: String,
    pub font_size: u32,
    pub foreground_color: Option<String>,
    pub horizontal_alignment: HorizontalAlignmentValues,
    pub is_bold: bool,
    pub is_double_underline: bool,
    pub is_italic: bool,
    pub is_underline: bool,
    pub is_wrap_text: bool,
    pub number_format: String,
    pub text_color: ColorSetting,
    pub vertical_alignment: VerticalAlignmentValues,
}

impl Default for CellStyleSetting {
    fn default() -> Self {
        Self {
            background_color: None,
            border_bottom: BorderSetting::default(),
            border_left: BorderSetting::default(),
            border_right: BorderSetting::default(),
            border_top: BorderSetting::default(),
            font_family: "Calibri".to_string(),
            font_size: 11,
            foreground_color: None,
            horizontal_alignment: HorizontalAlignmentValues::None,
            is_bold: false,
            is_double_underline: false,
            is_italic: false,
            is_underline: false,
            is_wrap_text: false,
            number_format: "General".to_string(),
            text_color: ColorSetting {
                color_setting_type: ColorSettingTypeValues::Rgb,
                value: "000000".to_string(),
            },
            vertical_alignment: VerticalAlignmentValues::None,
        }
    }
}

#[derive(Debug, Deserialize, Serialize)]
pub struct CellXfs {
    pub id: u32,
    pub format_id: u8,
    pub number_format_id: u8,
    pub font_id: u8,
    pub fill_id: u8,
    pub border_id: u8,
    pub apply_font: u8,
    pub apply_alignment: u8,
    pub apply_fill: u8,
    pub apply_border: u8,
    pub apply_number_format: u8,
    pub apply_protection: u8,
    pub is_wrap_text: u8,
    pub horizontal_alignment: HorizontalAlignmentValues,
    pub vertical_alignment: VerticalAlignmentValues,
}

impl Default for CellXfs {
    fn default() -> Self {
        Self {
            id: 0,
            format_id: 0,
            number_format_id: 0,
            font_id: 0,
            fill_id: 0,
            border_id: 0,
            apply_font: 0,
            apply_alignment: 0,
            apply_fill: 0,
            apply_border: 0,
            apply_number_format: 0,
            apply_protection: 0,
            is_wrap_text: 0,
            horizontal_alignment: HorizontalAlignmentValues::None,
            vertical_alignment: VerticalAlignmentValues::None,
        }
    }
}

#[derive(Debug, Deserialize, Serialize)]
pub struct FillStyle {
    pub id: u32,
    pub pattern_type: PatternTypeValues,
    pub background_color: Option<ColorSetting>,
    pub foreground_color: Option<ColorSetting>,
}

impl Default for FillStyle {
    fn default() -> Self {
        Self {
            id: 0,
            background_color: None,
            foreground_color: None,
            pattern_type: PatternTypeValues::None,
        }
    }
}

#[derive(Debug, Deserialize, Serialize)]
pub struct FontStyle {
    pub id: u32,
    pub name: String,
    pub size: u32,
    pub color: ColorSetting,
    pub family: i32,
    pub font_scheme: SchemeValues,
    pub is_bold: bool,
    pub is_double_underline: bool,
    pub is_italic: bool,
    pub is_underline: bool,
}

impl Default for FontStyle {
    fn default() -> Self {
        Self {
            id: 0,
            color: ColorSetting {
                value: "1".to_string(),
                ..Default::default()
            },
            family: 2,
            font_scheme: SchemeValues::None,
            is_bold: false,
            is_double_underline: false,
            is_italic: false,
            is_underline: false,
            name: "Calibri".to_string(),
            size: 11,
        }
    }
}

#[derive(Debug, Deserialize, Serialize)]
pub struct NumberFormats {
    pub id: u32,
    pub format_code: String,
}

impl Default for NumberFormats {
    fn default() -> Self {
        Self {
            id: 0,
            format_code: "General".to_string(),
        }
    }
}
