use super::common::CurrentNode;
use std::{cell::RefCell, rc::Rc};

/**
 * This contains the root document to work with
 */
pub struct OpenXmlFile {
    pub(crate) file_path: Option<String>,
    pub(crate) is_readonly: bool,
    pub(crate) archive_files: Vec<CurrentNode>,
    pub(crate) working_buffer: Rc<RefCell<Vec<u8>>>,
}
