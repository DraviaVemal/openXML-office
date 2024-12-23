use crate::{chain_error, openxml_office, StatusCode};
use draviavema_openxml_office::presentation_2007::{PowerPoint, PowerPointPropertiesModel};
use std::{
    ffi::{c_char, CStr, CString},
    slice::from_raw_parts,
};

#[no_mangle]
/// Creates a new Power Point object.
///
/// Returns a pointer to the newly created Power Point object.
/// If an error occurs, returns a null pointer.
pub extern "C" fn create_kl(
    file_name: *const c_char,
    buffer: *const u8,
    buffer_size: usize,
    out_power_point: *mut *mut PowerPoint,
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
    match flatbuffers::root::<openxml_office::presentation_2007::PresentationPropertiesModel>(
        buffer_slice,
    ) {
        Ok(fbs_power_point_properties) => {
            let power_point_properties = PowerPointPropertiesModel {
                is_in_memory: fbs_power_point_properties.is_in_memory(),
            };
            let power_point = if let Some(file_name) = file_name {
                PowerPoint::new(Some(file_name), power_point_properties)
            } else {
                PowerPoint::new(None, power_point_properties)
            };
            match power_point {
                Ok(power_point) => {
                    unsafe {
                        *out_power_point = Box::into_raw(Box::new(power_point));
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
///Save the Power Point File in provided file path
pub extern "C" fn save_as(
    power_point_ptr: *const u8,
    file_name: *const c_char,
    out_error: *mut *const c_char,
) -> i8 {
    if power_point_ptr.is_null() || file_name.is_null() {
        eprintln!("Received null pointer");
        return StatusCode::InvalidArgument as i8;
    }
    let file_name = unsafe { CStr::from_ptr(file_name) }
        .to_string_lossy()
        .into_owned();
    let power_point_ptr = power_point_ptr as *mut PowerPoint;
    let power_point = unsafe { Box::from_raw(power_point_ptr) };
    match power_point.save_as(&file_name) {
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
