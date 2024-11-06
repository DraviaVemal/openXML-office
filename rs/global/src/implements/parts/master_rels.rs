use std::{cell::RefCell, rc::Rc};

use crate::{xml_file::XmlElement, ContentTypesPart, MasterRelsPart};
use openxmloffice_xml::OpenXmlFile;
use quick_xml::Reader;

impl Drop for MasterRelsPart {
    fn drop(&mut self) {
        self.xml_fs
            .borrow()
            .add_update_xml_content(&self.file_name, &self.file_content)
            .expect("Master Rels Part Save Failed");
    }
}

impl XmlElement for MasterRelsPart {
    fn new(xml_fs: &Rc<RefCell<OpenXmlFile>>, _: Option<&str>) -> MasterRelsPart {
        let file_name = ".rels.xml".to_string();
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
        let template_core_properties = include_str!("master_rels.xml");
        let mut xml_parsed = Reader::from_str(template_core_properties);
        return Vec::new();
    }
}

impl ContentTypesPart {}
