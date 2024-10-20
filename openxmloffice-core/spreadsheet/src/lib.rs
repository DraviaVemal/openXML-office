pub mod enums;
pub mod implements;
pub mod structs;
pub mod tests;
pub mod utils;

use structs::excel::Excel;

pub fn create_excel() -> Excel {
    return Excel::new(None);
}

pub fn open_excel(file_path: String) -> Excel {
    return Excel::new(Some(file_path));
}

pub fn save_as(file_name: String) {}
