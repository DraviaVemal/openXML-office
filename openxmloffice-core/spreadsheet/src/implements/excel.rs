use openxmloffice_core_xml::OpenXmlFile;

use crate::structs::{excel::Excel, worksheet::Worksheet};

impl Excel {
    pub fn new(file_name: Option<String>) -> Self {
        if let Some(file_name) = file_name {
            let xml_fs = OpenXmlFile::open(&file_name, true);
            return Self {
                xml_fs,
                // Todo: read file for name and add
                worksheet_names: Vec::new(),
            };
        } else {
            let xml_fs = OpenXmlFile::create();
            return Self {
                xml_fs,
                worksheet_names: Vec::new(),
            };
        }
    }
    pub fn add_sheet(&self, sheet_name: &str) -> Worksheet {
        return Worksheet::new(&self, Some(sheet_name));
    }

    pub fn save_as(&self, file_name: &str) {
        self.xml_fs.save(file_name);
    }
}
