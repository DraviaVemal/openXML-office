use crate::structs::{common::CurrentNode, open_xml_archive::OpenXmlFile};
use std::{
    fs::{copy, metadata, remove_file, File},
    io::Write,
};
use tempfile::NamedTempFile;
use zip::{write::SimpleFileOptions, ZipArchive, ZipWriter};

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
        Self::create_initial_archive(&temp_file_path);
        // Default List of files common for all types
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
    fn create_initial_archive(temp_file_path: &str) {
        let physical_file =
            File::create(temp_file_path).expect("Creating Archive Physical File Failed");
        let mut archive = ZipWriter::new(physical_file);
        let zip_option = SimpleFileOptions::default();
        archive
            .add_directory("_rels", zip_option)
            .expect("_rels Directory creation failed");
        archive
            .add_directory("docProps", zip_option)
            .expect("docProps Directory creation failed");
        archive
            .start_file("_rels/.rels", zip_option)
            .expect("File write fail");
        archive.write_all(b"<xml />").expect("Write content failed");
        archive.finish().expect("Archive File Failed");
    }
}
