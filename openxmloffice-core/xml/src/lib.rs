mod enums;
mod implements;
mod structs;
mod tests;
mod utils;

use crate::structs::common::*;

// Edit existing file content
pub fn open_file(file_path: String, is_editable: bool) -> OpenXmlFile {
    return OpenXmlFile::open(file_path, is_editable);
}
// Create new file to work with
pub fn create_file() -> OpenXmlFile {
    return OpenXmlFile::create();
}
