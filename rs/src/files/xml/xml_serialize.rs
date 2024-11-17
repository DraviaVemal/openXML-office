use crate::files::XmlElement;
use anyhow::{Error as AnyError, Result as AnyResult};
use quick_xml::{events::Event, NsReader};
use std::io::Cursor;

pub struct XmlSerializer {}

impl XmlSerializer {
    pub fn xml_str_to_xml_tree(xml_str: Vec<u8>) -> AnyResult<XmlElement, AnyError> {
        let mut reader = NsReader::from_reader(Cursor::new(xml_str));
        reader.config_mut().trim_text(true);
        let mut temp_buffer: Vec<u8> = Vec::new();
        loop {
            match reader.read_event_into(&mut temp_buffer) {
                // Error when processing the excel file
                Err(e) => panic!("Error at position {}: {:?}", reader.buffer_position(), e),
                // Read XML declaration
                Ok(Event::Decl(e)) => {
                    println!("XML Dec: {:?}", String::from_utf8(e.to_vec()).unwrap());
                }
                // Read element tag and attributes, with or without text content
                Ok(Event::Start(e) | Event::Empty(e)) => {
                    println!("Start tag: {:?}", e.name());
                }
                // Read text content within element tag
                Ok(Event::Text(e)) => {
                    println!("Text: {}", String::from_utf8(e.to_vec()).unwrap());
                }
                // End of element tag
                Ok(Event::End(e)) => {
                    println!("End tag: {:?}", e.name());
                }
                // End of file
                Ok(Event::Eof) => break,
                // Just Extend all other data
                _ => {
                    println!("Some Other Event");
                }
            }
            temp_buffer.clear();
        }
        Ok(XmlElement::new("test".to_string(), None))
    }
}
