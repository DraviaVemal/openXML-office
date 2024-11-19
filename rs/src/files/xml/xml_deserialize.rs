use crate::files::XmlElement;
use anyhow::{anyhow, Context, Error as AnyError, Result as AnyResult};
use bincode::deserialize;
pub struct XmlDeSerializer {}

impl XmlDeSerializer {
    pub fn xml_tree_to_xml_vec(xml_element: &XmlElement) -> AnyResult<Vec<u8>, AnyError> {
        let mut xml_content = String::new();
        xml_content.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        Self::recursive_xml_element_parser(xml_element, &mut xml_content)
            .context("Create XML Contact String Failed")?;
        Ok(xml_content.as_bytes().to_vec())
    }

    fn recursive_xml_element_parser(
        xml_element: &XmlElement,
        master_string: &mut String,
    ) -> AnyResult<(), AnyError> {
        let attributes_string = match xml_element.get_attribute() {
            Some(attributes) => attributes
                .iter()
                .map(|(key, value)| format!("{}=\"{}\"", key, value))
                .collect::<Vec<String>>()
                .join(" "),
            None => "".to_string(),
        };
        if xml_element.is_empty_tag() {
            master_string.push_str(&format!(
                "<{} {}/>",
                xml_element.get_tag(),
                attributes_string
            ));
        } else if let Some(value) = xml_element.get_value() {
            master_string.push_str(&format!(
                "<{} {}>{}</{}>",
                xml_element.get_tag(),
                attributes_string,
                value,
                xml_element.get_tag()
            ));
        } else if let Some(childrens) = xml_element.get_children() {
            master_string.push_str(&format!(
                "<{} {}>",
                xml_element.get_tag(),
                attributes_string
            ));
            for child in childrens {
                let child_element =
                    deserialize::<XmlElement>(&child).context("Deserialize XML Element Failed")?;
                Self::recursive_xml_element_parser(&child_element, master_string)?;
            }
            master_string.push_str(&format!("</{}>", xml_element.get_tag()));
        } else {
            Err(anyhow!("Parsing XML Tree To File Content Failed"))?
        }
        Ok(())
    }
}
