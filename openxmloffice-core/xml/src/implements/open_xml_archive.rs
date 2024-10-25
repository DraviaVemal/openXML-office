use crate::structs::{
    common::CurrentNode, open_xml_archive::OpenXmlFile, open_xml_archive_write::OpenXmlEditable,
};
use std::{
    cell::RefCell,
    fs::{metadata, remove_file, File},
    io::{Cursor, Read, Write},
    rc::Rc,
};
use zip::ZipArchive;

impl OpenXmlFile {
    /// Create Current file helper object from exiting source
    pub fn open(file_path: &str, is_editable: bool) -> Self {
        let mut file = File::open(file_path).expect("Failed to read open file"); // Replace with your ZIP file path
        let mut buffer = Vec::new();
        file.read_to_end(&mut buffer);
        let working_buffer = Rc::new(RefCell::new(buffer));
        // Create a clone copy of master file to work with code
        Self {
            file_path: Some(file_path.to_string()),
            is_readonly: is_editable,
            archive_files: Self::read_initial_meta_data(&working_buffer),
            working_buffer,
        }
    }
    /// Create Current file helper object a new file to work with
    pub fn create() -> Self {
        let working_buffer = Rc::new(RefCell::new(Vec::new()));
        // Default List of files common for all types
        Self {
            file_path: None,
            is_readonly: true,
            archive_files: Self::read_initial_meta_data(&working_buffer),
            working_buffer,
        }
    }
    /// Save the current temp directory state file into final result
    pub fn save(&self, save_file: &str) {
        if metadata(save_file).is_ok() {
            remove_file(save_file).expect("Failed to Remove existing file");
        }
        let mut file = File::create(save_file).expect("Save file target failed");
        file.write_all(&self.working_buffer.borrow())
            .expect("File write failed");
    }

    fn read_initial_meta_data(working_file: &Rc<RefCell<Vec<u8>>>) -> Vec<CurrentNode> {
        let non_editable = OpenXmlEditable::new(working_file);
        // non_editable.read_archive_files();
        let content_buffer = working_file.borrow().clone();
        let archive =
            ZipArchive::new(Cursor::new(content_buffer)).expect("Actual archive file read failed");
        for file_name in archive.file_names() {
            println!("File Name : {}", file_name)
        }
        return vec![];
    }
    /// This creates initial archive for openXML file
    fn create_initial_archive(&self) {
        let editable = OpenXmlEditable::new(&self.working_buffer);
        editable
            .add_file("[Content_Types].xml", "<xml />")
            .expect("Adding File Failed");
        editable
            .add_file("_rels/.rels", "<xml />")
            .expect("Adding File Failed");
        editable
            .add_file("docProps/app.xml", "<xml />")
            .expect("Adding File Failed");
        editable
            .add_file("docProps/core.xml", "<xml />")
            .expect("Adding File Failed");
    }
}
