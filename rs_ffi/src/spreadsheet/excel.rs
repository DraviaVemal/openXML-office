use crate::{chain_error, openxml_office_ffi, StatusCode};
use draviavemal_openxml_office::spreadsheet_2007::{Excel, ExcelPropertiesModel};
use std::{
    ffi::{c_char, c_void, CStr, CString},
    slice::from_raw_parts,
};

#[no_mangle]
/// Creates a new Excel object.
///
/// Returns a pointer to the newly created Excel object.
/// If an error occurs, returns a null pointer.
pub extern "C" fn excel_create(
    file_name: *const c_char,
    buffer: *const u8,
    buffer_size: usize,
    out_excel: *mut *mut c_void,
    out_error: *mut *const c_char,
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
    match flatbuffers::root::<openxml_office_ffi::spreadsheet_2007::ExcelPropertiesModel>(
        buffer_slice,
    ) {
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
                        *out_excel = Box::into_raw(Box::new(excel)) as *mut c_void;
                    }
                    StatusCode::Success as i8
                }
                Err(e) => {
                    unsafe { *out_error = chain_error(&e) };
                    StatusCode::UnknownError as i8
                }
            }
        }
        Err(e) => {
            unsafe { *out_error = chain_error(&e.into()) };
            StatusCode::FlatBufferError as i8
        }
    }
}

#[no_mangle]
/// Add New Sheet to the Excel
pub extern "C" fn excel_add_sheet(
    excel_ptr: *const c_void,
    sheet_name: *const c_char,
    out_worksheet: *mut *mut c_void,
    out_error: *mut *const c_char,
) -> i8 {
    if excel_ptr.is_null() {
        eprintln!("Received null pointer");
        return StatusCode::InvalidArgument as i8;
    }
    let excel_ptr = excel_ptr as *mut Excel;
    let mut excel = unsafe { Box::from_raw(excel_ptr) };
    if sheet_name.is_null() {
        match excel.add_sheet(None) {
            Result::Ok(worksheet) => {
                unsafe {
                    *out_worksheet = Box::into_raw(Box::new(worksheet)) as *mut c_void;
                }
                StatusCode::Success as i8
            }
            Err(e) => match CString::new(format!("Flat Buffer Parse Error. {}", e)) {
                Result::Ok(str) => {
                    unsafe { *out_error = str.into_raw() };
                    StatusCode::Success as i8
                }
                Err(e) => {
                    eprintln!("Error String send Error. {}", e);
                    StatusCode::IoError as i8
                }
            },
        }
    } else {
        let sheet_name = unsafe { CStr::from_ptr(sheet_name) }
            .to_string_lossy()
            .into_owned();
        match excel.add_sheet(Some(sheet_name)) {
            Result::Ok(worksheet) => {
                unsafe {
                    *out_worksheet = Box::into_raw(Box::new(worksheet)) as *mut c_void;
                }
                StatusCode::Success as i8
            }
            Err(e) => match CString::new(format!("Flat Buffer Parse Error. {}", e)) {
                Result::Ok(str) => {
                    unsafe { *out_error = str.into_raw() };
                    StatusCode::Success as i8
                }
                Err(e) => {
                    eprintln!("Error String send Error. {}", e);
                    StatusCode::IoError as i8
                }
            },
        }
    }
}

#[no_mangle]
/// Get Existing Sheet from Excel
pub extern "C" fn excel_rename_sheet(
    excel_ptr: *const c_void,
    old_sheet_name: *const c_char,
    new_sheet_name: *const c_char,
    out_error: *mut *const c_char,
) -> i8 {
    if excel_ptr.is_null() || old_sheet_name.is_null() || new_sheet_name.is_null() {
        eprintln!("Received null pointer");
        return StatusCode::InvalidArgument as i8;
    }
    let excel_ptr = excel_ptr as *mut Excel;
    let mut excel = unsafe { Box::from_raw(excel_ptr) };
    let old_sheet_name = unsafe { CStr::from_ptr(old_sheet_name) }
        .to_string_lossy()
        .into_owned();
    let new_sheet_name = unsafe { CStr::from_ptr(new_sheet_name) }
        .to_string_lossy()
        .into_owned();
    match excel.rename_sheet_name(&old_sheet_name, &new_sheet_name) {
        Result::Ok(()) => StatusCode::Success as i8,
        Err(e) => match CString::new(format!("Flat Buffer Parse Error. {}", e)) {
            Result::Ok(str) => {
                unsafe { *out_error = str.into_raw() };
                StatusCode::Success as i8
            }
            Err(e) => {
                eprintln!("Error String send Error. {}", e);
                StatusCode::IoError as i8
            }
        },
    }
}

#[no_mangle]
/// Get Existing Sheet from Excel
pub extern "C" fn excel_get_sheet(
    excel_ptr: *const c_void,
    sheet_name: *const c_char,
    out_worksheet: *mut *mut c_void,
    out_error: *mut *const c_char,
) -> i8 {
    if excel_ptr.is_null() || sheet_name.is_null() {
        eprintln!("Received null pointer");
        return StatusCode::InvalidArgument as i8;
    }
    let excel_ptr = excel_ptr as *mut Excel;
    let mut excel = unsafe { Box::from_raw(excel_ptr) };
    let sheet_name = unsafe { CStr::from_ptr(sheet_name) }
        .to_string_lossy()
        .into_owned();
    match excel.get_worksheet(&sheet_name) {
        Result::Ok(worksheet) => {
            unsafe {
                *out_worksheet = Box::into_raw(Box::new(worksheet)) as *mut c_void;
            }
            StatusCode::Success as i8
        }
        Err(e) => match CString::new(format!("Flat Buffer Parse Error. {}", e)) {
            Result::Ok(str) => {
                unsafe { *out_error = str.into_raw() };
                StatusCode::Success as i8
            }
            Err(e) => {
                eprintln!("Error String send Error. {}", e);
                StatusCode::IoError as i8
            }
        },
    }
}

#[no_mangle]
/// List Sheet Name from Excel
pub extern "C" fn excel_list_sheet_name(
    excel_ptr: *const c_void,
    out_error: *mut *const c_char,
) -> i8 {
    if excel_ptr.is_null() {
        eprintln!("Received null pointer");
        return StatusCode::InvalidArgument as i8;
    }
    let excel_ptr = excel_ptr as *mut Excel;
    let mut excel = unsafe { Box::from_raw(excel_ptr) };
    let sheet_names = excel.list_sheet_names();
    return StatusCode::InvalidArgument as i8;
}

#[no_mangle]
///Save the Excel File in provided file path
pub extern "C" fn excel_save_as(
    excel_ptr: *const c_void,
    file_name: *const c_char,
    out_error: *mut *const c_char,
) -> i8 {
    if excel_ptr.is_null() || file_name.is_null() {
        eprintln!("Received null pointer");
        return StatusCode::InvalidArgument as i8;
    }
    let file_name = unsafe { CStr::from_ptr(file_name) }
        .to_string_lossy()
        .into_owned();
    let excel_ptr = excel_ptr as *mut Excel;
    let excel = unsafe { Box::from_raw(excel_ptr) };
    match excel.save_as(&file_name) {
        Result::Ok(()) => StatusCode::Success as i8,
        Err(e) => match CString::new(format!("Flat Buffer Parse Error. {}", e)) {
            Result::Ok(str) => {
                unsafe { *out_error = str.into_raw() };
                StatusCode::Success as i8
            }
            Err(e) => {
                eprintln!("Error String send Error. {}", e);
                StatusCode::IoError as i8
            }
        },
    }
}
