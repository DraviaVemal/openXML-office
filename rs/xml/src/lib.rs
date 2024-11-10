pub mod enums;
pub mod implements;
pub mod macros;
pub mod structs;
pub mod utils;

use anyhow::{Error as AnyError, Result as AnyResult};
pub use enums::*;
pub use macros::*;
pub use structs::*;
pub use utils::*;

/// Create new file to work with
pub fn create_file(is_in_memory: bool) -> AnyResult<OpenXmlFile, AnyError> {
    OpenXmlFile::create(is_in_memory)
}

/// Edit existing file content
pub fn open_file(
    file_path: String,
    is_editable: bool,
    is_in_memory: bool,
) -> AnyResult<OpenXmlFile, AnyError> {
    OpenXmlFile::open(&file_path, is_editable, is_in_memory)
}
