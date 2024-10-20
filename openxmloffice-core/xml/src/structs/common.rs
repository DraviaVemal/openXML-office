use crate::enums::common::NodeContentType;

pub(crate) struct CurrentNode {
    name: String,
    content_type: NodeContentType,
    childs: Option<Vec<CurrentNode>>,
}

/**
 * This contains the root document to work with
 */
pub struct OpenXmlFile {
    pub(crate) file_path: Option<String>,
    pub(crate) temp_file_path: String,
    pub(crate) is_readonly: bool,
    pub(crate) archive_files: Vec<CurrentNode>,
}