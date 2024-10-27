use super::common::CurrentNode;
use tempfile::NamedTempFile;

/**
 * This contains the root document to work with
 */
pub struct OpenXmlFile {
    pub(crate) file_path: Option<String>,
    pub(crate) is_readonly: bool,
    pub(crate) archive_files: Vec<CurrentNode>,
    pub(crate) temp_file: NamedTempFile,
}
