use crate::{xml_file::XmlElement, ThemePart};
use openxmloffice_xml::OpenXmlFile;
use std::{cell::RefCell, rc::Rc};

impl Drop for ThemePart {
    fn drop(&mut self) {
        self.xml_fs
            .borrow()
            .add_update_xml_content(&self.file_name, &self.file_content)
    }
}

impl XmlElement for ThemePart {
    fn new(xml_fs: &Rc<RefCell<OpenXmlFile>>, file_name: Option<&str>) -> ThemePart {
        let mut local_file_name = "".to_string();
        if let Some(file_name) = file_name {
            local_file_name = file_name.to_string();
        }
        Self {
            xml_fs: Rc::clone(xml_fs),
            file_content: Self::get_content_xml(&xml_fs, &local_file_name),
            file_name: local_file_name.to_string(),
        }
    }

    fn flush(self) {}

    /// Initialize xml content for this part from base template
    fn initialize_content_xml() -> Vec<u8> {
        let template_core_properties = include_str!("theme.xml");
        return template_core_properties.as_bytes().to_vec();
    }
}

impl ThemePart {}
