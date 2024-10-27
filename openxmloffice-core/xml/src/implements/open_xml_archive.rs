use crate::structs::{
    common::CurrentNode, open_xml_archive::OpenXmlFile, open_xml_archive_write::OpenXmlWrite,
};
use std::fs::{copy, metadata, remove_file, File};
use tempfile::NamedTempFile;
use zip::{ZipArchive, ZipWriter};

impl OpenXmlFile {
    /// Create Current file helper object from exiting source
    pub fn open(file_path: &str, is_editable: bool) -> Self {
        // Create a temp file to work with
        let temp_file = NamedTempFile::new().expect("Failed to create temporary file");
        let temp_file_path = temp_file
            .path()
            .to_str()
            .expect("str to String conversion fail");
        copy(&file_path, &temp_file_path).expect("Failed to copy file");
        // Create a clone copy of master file to work with code
        Self {
            file_path: Some(file_path.to_string()),
            is_readonly: is_editable,
            archive_files: Self::read_initial_meta_data(&temp_file_path),
            temp_file,
        }
    }
    /// Create Current file helper object a new file to work with
    pub fn create() -> Self {
        // Create a temp file to work with
        let temp_file = NamedTempFile::new().expect("Failed to create temporary file");
        let temp_file_path = temp_file
            .path()
            .to_str()
            .expect("str to String conversion fail");
        Self::create_initial_archive(temp_file_path);
        // Default List of files common for all types
        OpenXmlWrite::add_file("[Content_Types].xml");
        OpenXmlWrite::add_file("_rels/.rels");
        OpenXmlWrite::add_file("docProps/app.xml");
        OpenXmlWrite::add_file("docProps/core.xml");
        Self {
            file_path: None,
            is_readonly: true,
            archive_files: Self::read_initial_meta_data(&temp_file_path),
            temp_file,
        }
    }
    /// Save the current temp directory state file into final result
    pub fn save(&self, save_file: &str) {
        if metadata(save_file).is_ok() {
            remove_file(save_file).expect("Failed to Remove existing file");
        }
        copy(
            &self
                .temp_file
                .path()
                .to_str()
                .expect("str to String conversion fail"),
            &save_file,
        )
        .expect("Failed to place the result file");
    }

    fn read_initial_meta_data(working_file: &str) -> Vec<CurrentNode> {
        let file_buffer = File::open(working_file).expect("File Buffer for archive read failed");
        let archive = ZipArchive::new(file_buffer).expect("Actual archive file read failed");
        for file_name in archive.file_names() {
            println!("File Name : {}", file_name)
        }
        return vec![];
    }
    /// This creates initial archive for openXML file
    pub fn create_initial_archive(temp_file_path: &str) {
        let physical_file =
            File::create(temp_file_path).expect("Creating Archive Physical File Failed");
        let archive = ZipWriter::new(physical_file);
        archive.finish().expect("Archive File Failed");
    }
}
