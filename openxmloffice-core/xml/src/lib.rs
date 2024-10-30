pub mod enums;
pub mod implements;
pub mod macros;
pub mod structs;
mod tests;
pub mod utils;

pub use enums::*;
pub use macros::*;
pub use structs::*;
pub use utils::*;

/// Create new file to work with
pub fn create_file() -> OpenXmlFile {
    return OpenXmlFile::create();
}

/// Edit existing file content
pub fn open_file(file_path: String, is_editable: bool) -> OpenXmlFile {
    return OpenXmlFile::open(&file_path, is_editable);
}
