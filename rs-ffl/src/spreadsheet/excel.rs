use flatbuffers::{root, Verifier, VerifierOptions};
use openxmloffice_fbs::spreadsheet_2007;
use openxmloffice_spreadsheet::Excel;
use std::{
    ffi::{c_char, CStr},
    slice::from_raw_parts,
};

#[no_mangle]
/// Creates a new Excel object.
///
/// Returns a pointer to the newly created Excel object.
/// If an error occurs, returns a null pointer.
pub extern "C" fn create_excel(
    optional_string: *const c_char,
    buffer: *const u8,
    length: usize,
) -> *mut Excel {
    let file_name = if optional_string.is_null() {
        None
    } else {
        // Safety: `optional_string` is a valid C string pointer
        Some(
            unsafe { CStr::from_ptr(optional_string) }
                .to_string_lossy()
                .into_owned(),
        )
    };
    let buffer_slice = unsafe { from_raw_parts(buffer, length) };
    let verifier = Verifier::new(&VerifierOptions::default(), buffer_slice)
        .in_buffer::<spreadsheet_2007::ExcelPropertiesModel>(0);
    if let Ok(_verifier) = verifier {
        let fbs_excel_properties: spreadsheet_2007::ExcelPropertiesModel =
            root::<spreadsheet_2007::ExcelPropertiesModel>(buffer_slice)
                .expect("Decoding Flat Buffer Failed");
        let excel_properties = openxmloffice_spreadsheet::ExcelPropertiesModel {
            is_in_memory: fbs_excel_properties.is_in_memory(),
        };
        let excel = if let Some(file_name) = file_name {
            Box::new(Excel::new(Some(file_name), excel_properties))
        } else {
            Box::new(Excel::new(None, excel_properties))
        };
        return Box::into_raw(excel);
    } else {
        return std::ptr::null_mut();
    }
}
