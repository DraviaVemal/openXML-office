use quick_xml::{events::Event, NsReader};
use std::io::{Cursor, Write};

pub struct XmlHelper {
    root_element: Vec<u8>,
}

impl XmlHelper {
    /// Create new xml element object
    pub fn new(root_element: Vec<u8>) -> Self {
        Self { root_element }
    }

    /// Update/Replace target element text content
    pub fn update_element_text(&mut self, element_tag: &str, text: &str) {
        let mut reader = NsReader::from_reader(Cursor::new(&self.root_element));
        reader.config_mut().trim_text(true);
        let mut writer: Cursor<Vec<u8>> = Cursor::new(Vec::new());
        let mut temp_buffer: Vec<u8> = Vec::new();
        fn write_element(writer: &mut Cursor<Vec<u8>>, content: &[u8], wrap_closer: bool) {
            if wrap_closer {
                writer.write_all(b"<").expect("Wrap XML Failed");
                writer.write_all(content).expect("Clone Source Data");
                writer.write_all(b">").expect("Wrap XML Failed");
            } else {
                writer.write_all(content).expect("Clone Source Data");
            }
        }
        loop {
            match reader.read_event_into(&mut temp_buffer) {
                // Error when processing the excel file
                Err(e) => panic!("Error at position {}: {:?}", reader.buffer_position(), e),
                // Read XML declaration
                Ok(Event::Decl(e)) => {
                    println!("XML Dec: {:?}", String::from_utf8(e.to_vec()).unwrap());
                    write_element(&mut writer, &temp_buffer, true);
                }
                // Read element tag and attributes, with or without text content
                Ok(Event::Start(e) | Event::Empty(e)) => {
                    println!("Start tag: {:?}", e.name());
                    write_element(&mut writer, &temp_buffer, true);
                }
                // Read text content within element tag
                Ok(Event::Text(e)) => {
                    println!("Text: {}", String::from_utf8(e.to_vec()).unwrap());
                    write_element(&mut writer, &temp_buffer, false);
                }
                // End of element tag
                Ok(Event::End(e)) => {
                    println!("End tag: {:?}", e.name());
                    write_element(&mut writer, &temp_buffer, true);
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
        self.root_element = writer.into_inner();
    }

    pub fn get_first_element() {}

    pub fn get_last_element() {}

    pub fn update_element_value(element: &str, value: &str) {}

    pub fn update_element_attribute() {}
}
