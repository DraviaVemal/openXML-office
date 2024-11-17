use anyhow::Context;
use anyhow::{Error as AnyError, Result as AnyResult};
use bincode::{deserialize, serialize};
use serde::{Deserialize, Serialize};
use std::collections::HashMap;

/// Normalized XML representation
#[derive(Serialize, Deserialize, Debug)]
pub struct XmlElement {
    /// Element Tag name
    tag: String,
    /// Namespace of the element if applicable
    namespace: Option<String>,
    /// Attributes of the element if applicable
    attributes: Option<HashMap<String, String>>,
    /// Bincode byte data of XmlElement Children's of the node if applicable
    children: Option<Vec<Vec<u8>>>,
    /// Child node tag name list if applicable
    children_meta: Option<Vec<String>>,
}

impl XmlElement {
    pub fn new(tag: String, namespace: Option<String>) -> Self {
        Self {
            tag,
            namespace,
            attributes: None,
            children: None,
            children_meta: None,
        }
    }

    pub fn get_first_children(&self) -> Option<XmlElement> {
        if let Some(children) = &self.children {
            if let Some(element_bytes) = children.get(0) {
                deserialize::<XmlElement>(element_bytes).ok()
            } else {
                None
            }
        } else {
            None
        }
    }

    pub fn get_last_children(&self) -> Option<XmlElement> {
        if let Some(children) = &self.children {
            if let Some(element_bytes) = children.last() {
                deserialize::<XmlElement>(element_bytes).ok()
            } else {
                None
            }
        } else {
            None
        }
    }

    pub fn push_children(&mut self, xml_element: XmlElement) -> AnyResult<(), AnyError> {
        if self.children_meta.is_none() {
            self.children = Some(Vec::new());
            self.children_meta = Some(Vec::new());
        } else {
            self.children_meta
                .as_mut()
                .unwrap()
                .push(xml_element.tag.to_string());
        }
        if let Some(children) = &mut self.children {
            let serialized = serialize(&xml_element).context("Serializing XML Node failed")?;
            children.push(serialized);
        }
        Ok(())
    }
}
