mod utils;
mod enums;
mod structs;

use std::fs::copy;
use std::fs::File;
use zip::ZipArchive;
use tempfile::{tempdir, NamedTempFile};
use crate::structs::common::*;

impl OpenXmlFile {
    // Create Current file helper object from exiting source
    fn open(file_path: String, is_editable: bool) -> Self {
        // Create a temp directory to work with
        let temp_dir = tempdir().expect("Failed to create temporary directory");
        let temp_file = NamedTempFile::new_in(&temp_dir).expect("Failed to create temporary file");
        let temp_file_path = temp_file.path().to_str().unwrap().to_string();
        copy(&file_path, &temp_file_path).expect("Failed to copy file");
        // Create a clone copy of master file to work with code
        Self::read_initial_meta_data(&temp_file_path);
        Self {
            file_path: Some(file_path),
            temp_file_path,
            is_readonly: is_editable,
        }
    }
    // Create Current file helper object a new file to work with
    fn create() -> Self {
        // Create a temp directory to work with
        let temp_dir = tempdir().expect("Failed to create temporary directory");
        let temp_file = NamedTempFile::new_in(&temp_dir).expect("Failed to create temporary file");
        let temp_file_path = temp_file.path().to_str().unwrap().to_string();
        Self {
            file_path: None,
            temp_file_path,
            is_readonly: true,
        }
    }
    fn read_initial_meta_data(working_file: &str) {
        let file_buffer = File::open(working_file).expect("File Buffer for archive read failed");
        let archive = ZipArchive::new(file_buffer).expect("Actual archive file read failed");
        for file_name in archive.file_names() {

        };
    }
    // Read target file from archive
    fn read_zip_archive() {}
    // Write the content to archive file
    fn write_zip_archive() {}
    // Read file content and parse it to XML object
    fn read_xml(&self, file_path: String) {}
    // Use the XML object to write it to file format
    fn write_xml(&self, file_path: String) {}
}
// Edit existing file content
fn open_file(file_path: String, is_editable: bool) -> OpenXmlFile {
    return OpenXmlFile::open(file_path, is_editable);
}
// Create new file to work with
fn create_file() -> OpenXmlFile {
    return OpenXmlFile::create();
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn it_works() {
        let file = open_file(
            "/home/draviavemal/repo/OpenXML-Office/openxmloffice-core/xml/Book1.xlsx".to_string(),
            true,
        );
        println!("{}", file.temp_file_path);
        assert_eq!(true, true);
    }
}
