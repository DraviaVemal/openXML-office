use crate::CoreProperties;
use openxmloffice_core_xml::OpenXmlFile;
use quick_xml::Reader;

impl CoreProperties {
    /// Initialize Core property for new file
    pub fn initialize_core_properties(xml_fs: &OpenXmlFile) {
        let template_core_properties = include_str!("core_properties.xml");
        let mut xml_parsed = Reader::from_str(&template_core_properties);
        xml_parsed.config_mut().trim_text(true);
        
    }
    /// Update File Core Property
    pub fn update_core_properties(xml_fs: &OpenXmlFile) {}
}
