mod enums;
mod implements;
mod structs;
mod tests;
mod utils;

pub use crate::structs::common::*;

/// Create new file to work with
pub fn create_file() -> OpenXmlFile {
    return OpenXmlFile::create();
}

/// Edit existing file content
pub fn open_file(file_path: String, is_editable: bool) -> OpenXmlFile {
    return OpenXmlFile::open(&file_path, is_editable);
}
