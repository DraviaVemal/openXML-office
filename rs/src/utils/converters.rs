use anyhow::{anyhow, Error as AnyError, Result as AnyResult};

pub struct ConverterUtil;

impl ConverterUtil {
    pub fn get_column_int(cell_id: &str) -> AnyResult<usize, AnyError> {
        let column_part: String = cell_id.chars().take_while(|c| c.is_alphabetic()).collect();
        if column_part.is_empty() {
            return Err(anyhow!("Failed to Convert to Column Key Id"));
        }
        let mut index = 0;
        for (i, c) in column_part.chars().rev().enumerate() {
            let char_value = c.to_ascii_uppercase() as usize - 'A' as usize + 1;
            index += char_value * 26_usize.pow(i as u32);
        }
        Ok(index)
    }
}
