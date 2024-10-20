use crate::{CurrentNode, OpenXmlFile};
use std::fs::{copy, File};
use tempfile::{tempdir, NamedTempFile};
use zip::ZipArchive;

impl OpenXmlFile {
    /// Create Current file helper object from exiting source
    pub fn open(file_path: String, is_editable: bool) -> Self {
        // Create a temp directory to work with
        let temp_dir = tempdir().expect("Failed to create temporary directory");
        let temp_file = NamedTempFile::new_in(&temp_dir).expect("Failed to create temporary file");
        let temp_file_path = temp_file.path().to_str().unwrap().to_string();
        copy(&file_path, &temp_file_path).expect("Failed to copy file");
        // Create a clone copy of master file to work with code
        Self {
            file_path: Some(file_path),
            is_readonly: is_editable,
            archive_files: Self::read_initial_meta_data(&temp_file_path),
            temp_file_path,
        }
    }
    /// Create Current file helper object a new file to work with
    pub fn create() -> Self {
        // Create a temp directory to work with
        let temp_dir = tempdir().expect("Failed to create temporary directory");
        let temp_file = NamedTempFile::new_in(&temp_dir).expect("Failed to create temporary file");
        let temp_file_path = temp_file.path().to_str().unwrap().to_string();
        Self::create_initial_archive(&temp_file_path);
        Self {
            file_path: None,
            is_readonly: true,
            archive_files: Self::read_initial_meta_data(&temp_file_path),
            temp_file_path,
        }
    }
    fn read_initial_meta_data(working_file: &str) -> Vec<CurrentNode> {
        let file_buffer = File::open(working_file).expect("File Buffer for archive read failed");
        let archive = ZipArchive::new(file_buffer).expect("Actual archive file read failed");
        for file_name in archive.file_names() {
            println!("File Name : {}", file_name)
        }
        return vec![];
    }
    /// Read target file from archive
    fn read_zip_archive() {}
    /// This creates initial archive for openXML file
    fn create_initial_archive(temp_file_path: &str) {
        let physical_file =
            File::create(temp_file_path).expect("Creating Archive Physical File Failed");
        let mut archive = ZipArchive::new(physical_file).expect("Archive file creation Failed");
    }
    /// Write the content to archive file
    fn write_zip_archive() {}
    /// Read file content and parse it to XML object
    fn read_xml(&self, file_path: String) {}
    /// Use the XML object to write it to file format
    fn write_xml(&self, file_path: String) {}
}
