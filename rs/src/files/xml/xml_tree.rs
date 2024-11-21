use anyhow::{anyhow, Context};
use anyhow::{Error as AnyError, Result as AnyResult};
use bincode::{deserialize, serialize};
use serde::{Deserialize, Serialize};
use std::collections::HashMap;

#[derive(Debug)]
struct _XmlDocument {
    namespace_collection: HashMap<String, String>,
    xml_tree: XmlElement,
}
/// Master XML document
/// This helps in keeping the XML validation clean across document for all relations
/// TODO: Clear Unused namespace
/// TODO: Give clear handle for attribute and Namespace handling of root element
impl _XmlDocument {
    /// Create new XML Document Tree.
    /// The Root element will be created by default.
    pub fn _new(
        root_name: String,
        namespace_url: Option<String>,
        namespace_alias: Option<String>,
    ) -> Self {
        let mut namespace_collection: HashMap<String, String> = HashMap::new();
        if let Some(ns_url) = namespace_url {
            namespace_collection.insert(namespace_alias.clone().unwrap_or_default(), ns_url);
        }
        let xml_tree = XmlElement::new(root_name, namespace_alias);
        Self {
            namespace_collection,
            xml_tree,
        }
    }

    /// Get Root XML Element
    pub fn _get_root_element(self) -> XmlElement {
        self.xml_tree
    }

    /// Create Note Element in the document.
    /// Note: Its not inserted into tree. use push children to push it at specific leaf
    pub fn _create_element(
        &self,
        tag: String,
        namespace: Option<String>,
    ) -> AnyResult<XmlElement, AnyError> {
        if let Some(ns) = namespace {
            if self.namespace_collection.contains_key(&ns) {
                Ok(XmlElement::new(tag, Some(ns)))
            } else {
                Err(anyhow!("Provided Namespace : {} is not found in collection. Add the Namespace URL then proceed adding element",&ns))
            }
        } else {
            Ok(XmlElement::new(tag, None))
        }
    }
}

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
    value: Option<String>,
}

impl XmlElement {
    pub fn new(tag: String, namespace: Option<String>) -> Self {
        Self {
            tag,
            namespace,
            attributes: None,
            children: None,
            children_meta: None,
            value: None,
        }
    }
    // ===================== Data Read Only Methods ===============================
    pub fn is_empty_tag(&self) -> bool {
        self.value.is_none() && self.children.is_none()
    }

    pub fn get_tag(&self) -> &str {
        &self.tag
    }

    pub fn get_attribute(&self) -> &Option<HashMap<String, String>> {
        &self.attributes
    }

    pub fn get_value(&self) -> &Option<String> {
        &self.value
    }

    pub fn get_children(&self) -> &Option<Vec<Vec<u8>>> {
        &self.children
    }

    /// Get the element to read the content
    pub fn find_child_element_by_attribute(
        &self,
        attribute: &str,
        value: &str,
    ) -> AnyResult<Option<XmlElement>, AnyError> {
        if let Some(children) = &self.children {
            for item in children {
                let xml_element =
                    deserialize::<XmlElement>(item).context("Deserializing bincode failed")?;
                if let Some(attributes) = &xml_element.attributes {
                    if let Some((_, attr_value)) = attributes.get_key_value(attribute) {
                        if attr_value == value {
                            return Ok(Some(xml_element));
                        }
                    }
                }
            }
        }
        Ok(None)
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

    // ===================== Data Update Methods ===============================
    pub fn update_text_value(&mut self, xml_path: &str, new_value: &str) {
        let tree_path: Vec<&str> = xml_path.split("->").collect();
        
    }

    pub fn set_attribute(&mut self, attributes: HashMap<String, String>) -> () {
        self.attributes = Some(attributes)
    }

    pub fn set_value(&mut self, text: String) -> () {
        self.value = Some(text)
    }

    pub fn push_children(&mut self, xml_element: XmlElement) -> AnyResult<(), AnyError> {
        if self.children_meta.is_none() {
            self.children = Some(Vec::new());
            self.children_meta = Some(Vec::new());
        }
        if let Some(children) = &mut self.children {
            let serialized = serialize(&xml_element).context("Serializing XML Node failed")?;
            self.children_meta
                .as_mut()
                .unwrap()
                .push(xml_element.tag.to_string());
            children.push(serialized);
        }
        Ok(())
    }
}
