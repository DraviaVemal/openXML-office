use anyhow::{anyhow, Context, Error as AnyError, Ok, Result as AnyResult};

pub struct ConverterUtil;

impl ConverterUtil {
    /// Return int Id of the column
    pub fn get_column_index(cell_ref: &str) -> AnyResult<u16, AnyError> {
        let column_part: String = cell_ref.chars().take_while(|c| c.is_alphabetic()).collect();
        if column_part.is_empty() {
            return Err(anyhow!("Failed to Convert to Column Key Id"));
        }
        let mut index = 0;
        for (i, c) in column_part.chars().rev().enumerate() {
            let char_value = c.to_ascii_uppercase() as u16 - 'A' as u16 + 1;
            index += char_value * 26_usize.pow(i as u32) as u16;
        }
        Ok(index)
    }
    /// Return String ref of the column
    pub fn get_column_ref(cell_id: u16) -> AnyResult<String, AnyError> {
        if cell_id == 0 {
            return Err(anyhow!("Index must be greater than 0"));
        }
        let mut index = cell_id;
        let mut column_name = String::new();

        while index > 0 {
            index -= 1;
            let char_value = (index % 26) as u8 + b'A';
            column_name.insert(0, char_value as char);
            index /= 26;
        }

        Ok(column_name)
    }

    /// Return
    pub fn get_cell_index(cell_ref: &str) -> AnyResult<(u32, u16), AnyError> {
        Ok((
            Self::extract_digits(cell_ref).context("Failed to extract int key")?,
            Self::get_column_index(cell_ref).context("Failed to Convert to int key")?,
        ))
    }

    /// convert open-xml bool flag property
    pub(crate) fn normalize_bool_property_u8(value: &str) -> u8 {
        match value.trim() {
            "true" | "1" => 1,
            _ => 0,
        }
    }
    /// convert bool to open-xml flag property
    pub(crate) fn bool_xml_flag(value: &bool) -> String {
        if *value {
            "1".to_string()
        } else {
            "0".to_string()
        }
    }
    /// convert open-xml bool flag property
    pub(crate) fn normalize_bool_property_bool(value: &str) -> bool {
        match value.trim() {
            "true" | "1" => true,
            _ => false,
        }
    }

    fn extract_digits(input: &str) -> AnyResult<u32> {
        input
            .chars()
            .filter(|c| c.is_digit(10))
            .collect::<String>()
            .parse()
            .context("Failed to Extract Digits")
    }
}
