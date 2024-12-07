use crate::global_2007::traits::Enum;

#[derive(Debug)]
pub enum HyperlinkPropertyTypeValues {
    EXISTING_FILE,
    WEB_URL,
    TARGET_SHEET,
    TARGET_SLIDE,
    NEXT_SLIDE,
    PREVIOUS_SLIDE,
    FIRST_SLIDE,
    LAST_SLIDE,
}

// impl Enum<HyperlinkPropertyTypeValues> for HyperlinkPropertyTypeValues {
//     fn get_string(input_enum: HyperlinkPropertyTypeValues) -> String {
//         match input_enum {
//             HyperlinkPropertyTypeValues::Theme => "theme".to_string(),
//             HyperlinkPropertyTypeValues::Rgb => "rgb".to_string(),
//             HyperlinkPropertyTypeValues::Indexed => "indexed".to_string(),
//         }
//     }
//     fn get_enum(input_string: &str) -> HyperlinkPropertyTypeValues {
//         match input_string {
//             "theme" => HyperlinkPropertyTypeValues::Theme,
//             "rgb" => HyperlinkPropertyTypeValues::Rgb,
//             _ => HyperlinkPropertyTypeValues::Indexed,
//         }
//     }
// }

#[derive(Debug)]
pub struct HyperlinkProperties {}
