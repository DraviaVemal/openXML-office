use crate::files::XmlElement;
use anyhow::{anyhow, Context, Error as AnyError, Ok, Result as AnyResult};
use quick_xml::{
    events::{
        attributes::{AttrError, Attribute},
        Event,
    },
    NsReader,
};
use std::{collections::HashMap, io::Cursor};

enum MethodResult {
    XmlElement(XmlElement),
    String(String),
    EndTag(String),
    EndFile(()),
}

pub struct XmlSerializer {}

impl XmlSerializer {
    pub fn xml_str_to_xml_tree(xml_str: Vec<u8>) -> AnyResult<XmlElement, AnyError> {
        let mut reader: NsReader<Cursor<Vec<u8>>> = NsReader::from_reader(Cursor::new(xml_str));
        reader.config_mut().trim_text(true);
        match Self::recursive_xml_element_parser(&mut reader) {
            Result::Ok(method_result) => match method_result {
                MethodResult::XmlElement(element) => Ok(element),
                _ => Err(anyhow!("This is not valid XML to parse.")),
            },
            Err(e) => Err(e),
        }
    }

    fn recursive_xml_element_parser(
        reader: &mut NsReader<Cursor<Vec<u8>>>,
    ) -> AnyResult<MethodResult, AnyError> {
        let mut temp_buffer: Vec<u8> = Vec::new();
        loop {
            match reader.read_event_into(&mut temp_buffer) {
                // Error when processing the excel file
                Err(e) => break Err(e.into()),
                // Read XML declaration
                Result::Ok(Event::Decl(_)) => {
                    temp_buffer.clear();
                    continue;
                }
                // Read element tag and attributes, with text content
                Result::Ok(Event::Start(element)) => {
                    let attributes: HashMap<String, String> = element
                        .attributes()
                        .map(|attribute_result: Result<Attribute<'_>, AttrError>| {
                            let attribute: Attribute<'_> =
                                attribute_result.context("Failed to parse attribute")?;
                            let key: String =
                                String::from_utf8_lossy(attribute.key.into_inner()).to_string();
                            let value: String =
                                String::from_utf8_lossy(&attribute.value).to_string();
                            Ok((key, value))
                        })
                        .collect::<AnyResult<HashMap<String, String>>>()?;
                    // Create new element
                    let tag: String =
                        String::from_utf8_lossy(element.name().into_inner()).to_string();
                    let mut xml_element: XmlElement = XmlElement::new(tag.clone(), None);
                    if !attributes.is_empty() {
                        xml_element.set_attribute(attributes);
                    }
                    let mut continue_element_loop = true;
                    while continue_element_loop {
                        continue_element_loop = match Self::recursive_xml_element_parser(reader) {
                            Err(err) => Err(err),
                            Result::Ok(method_result) => match method_result {
                                MethodResult::XmlElement(child) => {
                                    xml_element
                                        .push_children(child)
                                        .context("Add Child Node Failed")?;
                                    Ok(true)
                                }
                                MethodResult::String(text) => {
                                    xml_element.set_value(text);
                                    Ok(true)
                                }
                                MethodResult::EndTag(tag) => {
                                    if tag == xml_element.get_tag() {
                                        Ok(false)
                                    } else {
                                        Err(anyhow!("Invalid XML. Parsing Failed"))
                                    }
                                }
                                MethodResult::EndFile(()) => Ok(false),
                            },
                        }
                        .context("XML Recursive Parsing Failed.")?;
                    }
                    break Ok(MethodResult::XmlElement(xml_element)); // Return the full element
                }
                // Read element tag and attributes, without text content
                Result::Ok(Event::Empty(element)) => {
                    let attributes: HashMap<String, String> = element
                        .attributes()
                        .map(|a| {
                            let attribute = a.context("Failed to parse attribute")?;
                            let key =
                                String::from_utf8_lossy(attribute.key.into_inner()).to_string();
                            let value = String::from_utf8_lossy(&attribute.value).to_string();
                            Ok((key, value))
                        })
                        .collect::<AnyResult<HashMap<String, String>>>()?;
                    // Create new element
                    let tag: String =
                        String::from_utf8_lossy(element.name().into_inner()).to_string();
                    let mut xml_element: XmlElement = XmlElement::new(tag.clone(), None);
                    if !attributes.is_empty() {
                        xml_element.set_attribute(attributes);
                    }
                    break Ok(MethodResult::XmlElement(xml_element)); // Return the full element
                }
                // Read text content within element tag
                Result::Ok(Event::Text(byte_text)) => {
                    let text = byte_text.unescape().context("XML Text parsing error")?;
                    break Ok(MethodResult::String(text.to_string()));
                }
                // End of element tag
                Result::Ok(Event::End(element)) => {
                    let tag = String::from_utf8_lossy(element.name().into_inner()).to_string();
                    break Ok(MethodResult::EndTag(tag));
                } // Return the full element
                // End of file
                Result::Ok(Event::Eof) => break Ok(MethodResult::EndFile(())),
                // Just Extend all other data
                _ => {
                    temp_buffer.clear();
                    continue;
                }
            }
        }
    }
}
