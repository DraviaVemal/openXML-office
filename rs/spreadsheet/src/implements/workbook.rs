use std::{cell::RefCell, rc::Rc};

use crate::structs::workbook::Workbook;
use openxmloffice_global::xml_file::XmlElement;
use openxmloffice_xml::OpenXmlFile;

impl Drop for Workbook {
    fn drop(&mut self) {
        self.xml_fs
            .borrow()
            .add_update_xml_content(&self.file_name, &self.file_content)
            .expect("Workbook Save Failed");
    }
}

impl XmlElement for Workbook {
    /// Create workbook
    fn new(xml_fs: &Rc<RefCell<OpenXmlFile>>, _: Option<&str>) -> Self {
        let file_name = "xl/workbook.xml".to_string();
        return Self {
            xml_fs: Rc::clone(&xml_fs),
            file_content: Self::get_content_xml(&xml_fs, &file_name),
            file_name,
        };
    }

    fn flush(self) {}

    /// Read and load workbook xml to work with
    fn get_content_xml(xml_fs: &Rc<RefCell<OpenXmlFile>>, file_name: &str) -> Vec<u8> {
        let results = xml_fs.borrow().get_xml_content(file_name);
        if let Some(results) = results {
            return results;
        } else {
            return Self::initialize_content_xml(&xml_fs);
        }
    }

    /// Initialize workbook for new excel
    fn initialize_content_xml(xml_fs: &Rc<RefCell<OpenXmlFile>>) -> Vec<u8> {
        let content = Vec::new();
        return content;
    }
}
