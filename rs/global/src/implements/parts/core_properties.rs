use std::{cell::RefCell, rc::Rc};

use crate::{xml_file::XmlElement, CorePropertiesPart};
use openxmloffice_xml::OpenXmlFile;
use quick_xml::Reader;

impl Drop for CorePropertiesPart {
    fn drop(&mut self) {
        self.update_last_modified();
        self.xml_fs
            .borrow()
            .add_update_xml_content(&self.file_name, &self.file_content)
            .expect("Core Property Part Save Failed");
    }
}

impl XmlElement for CorePropertiesPart {
    fn new(xml_fs: &Rc<RefCell<OpenXmlFile>>, _: Option<&str>) -> CorePropertiesPart {
        let file_name = "docProps/core.xml".to_string();
        Self {
            xml_fs: Rc::clone(xml_fs),
            file_content: Self::get_content_xml(&xml_fs, &file_name),
            file_name,
        }
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
        let template_core_properties = include_str!("core_properties.xml");
        let mut xml_parsed = Reader::from_str(template_core_properties);
        return Vec::new();
    }
}

impl CorePropertiesPart {

    fn update_last_modified(&self){

    }
}
