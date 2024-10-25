use crate::OpenXmlFile;

pub struct OpenXmlEditable<'file_handle> {
    pub open_xml_file: &'file_handle OpenXmlFile,
}
