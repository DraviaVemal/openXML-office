use crate::OpenXmlFile;

pub struct OpenXmlNonEditable<'file_handle> {
    pub open_xml_file: &'file_handle OpenXmlFile,
}
