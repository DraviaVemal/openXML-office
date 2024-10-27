use tempfile::NamedTempFile;

use super::common::CurrentNode;

/**
 * This contains the root document to work with
 */
pub struct OpenXmlFile {
    pub(crate) file_path: Option<String>,
    pub(crate) temp_file: NamedTempFile,
    pub(crate) is_readonly: bool,
    pub(crate) archive_files: Vec<CurrentNode>,
}
