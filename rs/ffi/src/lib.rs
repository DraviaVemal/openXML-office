use openxmloffice_spreadsheet::Excel;

#[no_mangle]
pub extern "C" fn create_excel(file_name: Option<String>) {
    if let Some(file_name) = file_name {
        // return Excel::new(Some(file_name));
    } else {
        // return Excel::new(None);
    }
}
