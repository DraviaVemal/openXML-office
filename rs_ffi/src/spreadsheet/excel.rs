use crate::{openxml_office, StatusCode};
use draviavema_openxml_office::{spreadsheet::Excel, ExcelPropertiesModel};
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
    file_name: *const c_char,
    buffer: *const u8,
    buffer_size: usize,
    out_excel: *mut *mut Excel,
) -> i8 {
    let file_name = if file_name.is_null() {
        None
    } else {
        Some(
            unsafe { CStr::from_ptr(file_name) }
                .to_string_lossy()
                .into_owned(),
        )
    };
    if buffer.is_null() || buffer_size == 0 {
        return StatusCode::InvalidArgument as i8;
    }
    let buffer_slice = unsafe { from_raw_parts(buffer, buffer_size) };
    match flatbuffers::root::<openxml_office::spreadsheet_2007::ExcelPropertiesModel>(buffer_slice)
    {
        Ok(fbs_excel_properties) => {
            let excel_properties = ExcelPropertiesModel {
                is_in_memory: fbs_excel_properties.is_in_memory(),
            };
            let excel = if let Some(file_name) = file_name {
                Excel::new(Some(file_name), excel_properties)
            } else {
                Excel::new(None, excel_properties)
            };
            match excel {
                Ok(excel) => {
                    unsafe {
                        *out_excel = Box::into_raw(Box::new(excel));
                    }
                    StatusCode::Success as i8
                }
                Err(e) => StatusCode::UnknownError as i8,
            }
        }
        Err(e) => {
            return StatusCode::FlatBufferError as i8;
        }
    }
}

#[no_mangle]
///Save the Excel File in provided file path
pub extern "C" fn save_as(excel_ptr: *const u8, file_name: *const c_char) -> i8 {
    if excel_ptr.is_null() || file_name.is_null() {
        eprintln!("Received null pointer");
        return StatusCode::InvalidArgument as i8;
    }
    let file_name = unsafe { CStr::from_ptr(file_name) }
        .to_string_lossy()
        .into_owned();
    let excel_ptr = excel_ptr as *mut Excel;
    let excel = unsafe { Box::from_raw(excel_ptr) };
    excel.save_as(&file_name);
    return StatusCode::Success as i8;
}
