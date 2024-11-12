use crate::OpenXmlFile;
use anyhow::{Error as AnyError, Ok, Result as AnyResult};
use std::{cell::RefCell, rc::Rc};

pub trait XmlElement {
    /// Create new object with file connector handle
    fn new(xml_fs: &Rc<RefCell<OpenXmlFile>>, file_name: Option<&str>) -> AnyResult<Self, AnyError>
    where
        Self: Sized;
    /// Save the current file state
    fn flush(self);
    /// Get content of the current xml
    fn get_content_xml(
        xml_fs: &Rc<RefCell<OpenXmlFile>>,
        file_name: &str,
    ) -> AnyResult<Vec<u8>, AnyError> {
        let results = xml_fs.borrow().get_xml_content(file_name)?;
        if let Some(results) = results {
            Ok(results)
        } else {
            Ok(Self::initialize_content_xml())
        }
    }
    /// Initialize the content if not already exist
    fn initialize_content_xml() -> Vec<u8>;
}
