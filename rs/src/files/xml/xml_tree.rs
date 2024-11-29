use anyhow::{anyhow, Context, Error as AnyError, Result as AnyResult};
use std::{
    cell::RefCell,
    collections::HashMap,
    rc::{Rc, Weak},
};

#[derive(Debug)]
pub struct XmlDocument {
    running_id: usize,
    /// XML Namespace collection
    namespace_collection: Rc<RefCell<HashMap<String, String>>>,
    /// XML Element Collection
    xml_element_collection: HashMap<usize, XmlElement>,
}

/// ####################### Im-Mut Access Functions ####################
impl XmlDocument {
    pub fn new() -> Self {
        Self {
            running_id: 0,
            namespace_collection: Rc::new(RefCell::new(HashMap::new())),
            xml_element_collection: HashMap::new(),
        }
    }

    pub fn get_element_count(&self) -> usize {
        self.xml_element_collection.len()
    }

    pub fn get_root(&self) -> Option<&XmlElement> {
        self.xml_element_collection.get(&0)
    }

    pub fn get_element_by_attribute(
        &self,
        parent_id: &usize,
        attribute_key: &str,
        attribute_value: &str,
    ) -> Option<&XmlElement> {
        if let Some(parent_element) = self.xml_element_collection.get(&parent_id) {
            if let Some(found_child) = parent_element.children.borrow().iter().find(|item| {
                if let Some(current) = self.xml_element_collection.get(&item.id) {
                    if let Some(attribute) = current.get_attribute() {
                        if let Some(value) = attribute.get(attribute_key) {
                            return value == attribute_value;
                        }
                    }
                }
                false
            }) {
                return self.xml_element_collection.get(&found_child.id);
            }
        }
        None
    }

    pub fn get_element_ids_by_tag(
        &self,
        parent_id: &usize,
        filter_tag: &str,
    ) -> Option<Vec<usize>> {
        if let Some(parent_element) = self.xml_element_collection.get(parent_id) {
            let element_id_list = parent_element
                .children
                .borrow()
                .iter()
                .filter(|item| item.tag == filter_tag)
                .map(|item| item.id)
                .collect::<Vec<usize>>();
            if element_id_list.len() > 0 {
                return Some(element_id_list);
            }
        }
        None
    }

    pub fn get_element(&self, element_id: &usize) -> Option<&XmlElement> {
        self.xml_element_collection.get(element_id)
    }
}

/// ####################### Mut Access Functions ####################
impl XmlDocument {
    pub fn get_root_mut(&mut self) -> Option<&mut XmlElement> {
        self.xml_element_collection.get_mut(&0)
    }

    pub fn create_root_mut(&mut self, tag: &str) -> AnyResult<&mut XmlElement, AnyError> {
        let element = XmlElement::new(Rc::downgrade(&self.namespace_collection), tag)
            .context("Create XML Element Failed")?;
        self.xml_element_collection.insert(0, element);
        Ok(self.xml_element_collection.get_mut(&0).unwrap())
    }
    /// Removes the child reference from parent and remove the master element itself from collection
    pub fn pop_elements_by_tag_mut(
        &mut self,
        parent_id: &usize,
        filter_tag: &str,
    ) -> Option<Vec<XmlElement>> {
        if let Some(parent_element) = self.xml_element_collection.get(parent_id) {
            let element_id_list = parent_element
                .children
                .borrow()
                .iter()
                .filter(|item| item.tag == filter_tag)
                .map(|item| item.id)
                .collect::<Vec<usize>>();
            if element_id_list.len() > 0 {
                parent_element
                    .children
                    .borrow_mut()
                    .retain(|item| item.tag != filter_tag);
                let mut elements: Vec<XmlElement> = Vec::new();
                for element_id in element_id_list {
                    if let Some(item) = self.xml_element_collection.remove(&element_id) {
                        elements.push(item);
                    }
                }
                return Some(elements);
            }
        }
        None
    }

    pub fn pop_element_mut(&mut self, element_id: &usize) -> Option<XmlElement> {
        self.xml_element_collection.remove(element_id)
    }

    pub fn insert_children_before_tag_mut(
        &mut self,
        parent_id: &usize,
        find_tag: &str,
        tag: &str,
    ) -> AnyResult<&mut XmlElement, AnyError> {
        if let Some(parent_element) = self.xml_element_collection.get_mut(&parent_id) {
            let mut element = XmlElement::new(Rc::downgrade(&self.namespace_collection), tag)
                .context("Create XML Element Failed")?;
            element.set_parent_id_mut(*parent_id);
            self.running_id += 1;
            element.set_id_mut(self.running_id);
            parent_element.insert_children_before_tag_mut(self.running_id, find_tag, tag);
            self.xml_element_collection.insert(self.running_id, element);
            Ok(self
                .xml_element_collection
                .get_mut(&self.running_id)
                .unwrap())
        } else {
            return Err(anyhow!("Parent Element Not Found"));
        }
    }

    pub fn insert_child_at_mut(
        &mut self,
        parent_id: &usize,
        position: &usize,
        tag: &str,
    ) -> AnyResult<&mut XmlElement, AnyError> {
        if let Some(parent_element) = self.xml_element_collection.get_mut(&parent_id) {
            let mut element = XmlElement::new(Rc::downgrade(&self.namespace_collection), tag)
                .context("Create XML Element Failed")?;
            element.set_parent_id_mut(*parent_id);
            self.running_id += 1;
            element.set_id_mut(self.running_id);
            parent_element.insert_children_at_mut(self.running_id, *position, tag);
            self.xml_element_collection.insert(self.running_id, element);
            Ok(self
                .xml_element_collection
                .get_mut(&self.running_id)
                .unwrap())
        } else {
            return Err(anyhow!("Parent Element Not Found"));
        }
    }

    pub fn append_child_mut(
        &mut self,
        parent_id: &usize,
        tag: &str,
    ) -> AnyResult<&mut XmlElement, AnyError> {
        if let Some(parent_element) = self.xml_element_collection.get_mut(&parent_id) {
            let mut element = XmlElement::new(Rc::downgrade(&self.namespace_collection), tag)
                .context("Create XML Element Failed")?;
            element.set_parent_id_mut(*parent_id);
            self.running_id += 1;
            element.set_id_mut(self.running_id);
            parent_element.append_children_mut(self.running_id, tag);
            self.xml_element_collection.insert(self.running_id, element);
            Ok(self
                .xml_element_collection
                .get_mut(&self.running_id)
                .unwrap())
        } else {
            return Err(anyhow!("Parent Element Not Found"));
        }
    }

    pub fn get_first_element_mut(
        &mut self,
        start_element: &usize,
        mut element_tree: Vec<String>,
    ) -> AnyResult<Option<&mut XmlElement>, AnyError> {
        element_tree.reverse();
        let mut current_id = *start_element;
        let mut is_found = false;
        loop {
            if let Some(find_tag) = element_tree.pop() {
                let element = self.xml_element_collection.get(&current_id).unwrap();
                let element_child = element
                    .children
                    .try_borrow_mut()
                    .context("Children Borrow Failed")?;
                if let Some(found_child) = element_child.iter().find(|item| item.tag == find_tag) {
                    current_id = found_child.id;
                    if element_tree.len() == 0 {
                        is_found = true;
                        break;
                    }
                } else {
                    // Not Able to find any one Child
                    break;
                }
            } else {
                // If Vec has No Value to start
                break;
            }
        }
        if is_found {
            return Ok(self.xml_element_collection.get_mut(&current_id));
        }
        Ok(None)
    }

    pub fn get_element_by_attribute_mut(
        &mut self,
        parent_id: &usize,
        attribute_key: &str,
        attribute_value: &str,
    ) -> Option<&mut XmlElement> {
        let mut id: Option<usize> = None;
        if let Some(parent_element) = self.xml_element_collection.get(&parent_id) {
            if let Some(found_child) = parent_element.children.borrow().iter().find(|item| {
                if let Some(current) = self.xml_element_collection.get(&item.id) {
                    if let Some(attribute) = current.get_attribute() {
                        if let Some(value) = attribute.get(attribute_key) {
                            return value == attribute_value;
                        }
                    }
                }
                false
            }) {
                id = Some(found_child.id);
            }
        }
        if let Some(find_id) = id {
            return self.xml_element_collection.get_mut(&find_id);
        }
        None
    }

    pub fn get_element_mut(&mut self, element_id: &usize) -> Option<&mut XmlElement> {
        self.xml_element_collection.get_mut(element_id)
    }
}

#[derive(Debug)]
pub struct XmlElementChild {
    id: usize,
    tag: String,
}
/// Normalized XML representation
#[derive(Debug)]
pub struct XmlElement {
    /// Current node Id
    id: usize,
    /// Parent id if applicable else -1
    parent_id: usize,
    /// Element Tag name with Namespace
    tag: String,
    /// Attributes of the element if applicable with namespace
    attributes: Option<HashMap<String, String>>,
    /// Internal Value of the
    value: Option<String>,
    /// Child Element Names to pull up nodes quickly
    children: Rc<RefCell<Vec<XmlElementChild>>>,
    // ######################## Document Parts ##########################
    /// XML Namespace collection
    namespace_collection_ref: Weak<RefCell<HashMap<String, String>>>,
}

impl XmlElement {
    /// Create element with tree document reference
    fn new(
        namespace_collection: Weak<RefCell<HashMap<String, String>>>,
        tag: &str,
    ) -> AnyResult<Self, AnyError> {
        if Self::is_valid_xml_name(tag) {
            return Ok(Self {
                id: 0,
                parent_id: 0,
                tag: tag.to_string(),
                attributes: None,
                value: None,
                children: Rc::new(RefCell::new(Vec::new())),
                namespace_collection_ref: namespace_collection,
            });
        }
        Err(anyhow!("Provided Tag Name \"{}\" is not valid.", tag))
    }

    fn is_valid_xml_name(name: &str) -> bool {
        fn is_name_start_char(c: char) -> bool {
            c.is_ascii_alphabetic() || c == '_' || c.is_alphabetic()
        }

        fn is_name_char(c: char) -> bool {
            is_name_start_char(c) || c.is_ascii_digit() || c == '-' || c == '.' || c == ':'
        }

        if name.is_empty() {
            return false;
        }

        let mut chars = name.chars();
        if let Some(first_char) = chars.next() {
            if !is_name_start_char(first_char) {
                return false;
            }
        }
        let mut colon_count = 0;
        for c in chars {
            if c == ':' {
                colon_count += 1;
                if colon_count > 1 {
                    return false;
                }
            } else if !is_name_char(c) {
                return false;
            }
        }
        true
    }

    // ######################## Data Read Only Methods ###########################

    pub fn get_tag(&self) -> &str {
        &self.tag
    }

    pub fn has_value(&self) -> bool {
        self.get_value().is_some()
    }

    pub fn is_empty_tag(&self) -> bool {
        self.children.borrow().len() == 0 && self.value.is_none()
    }

    pub fn get_attribute(&self) -> Option<&HashMap<String, String>> {
        self.attributes.as_ref()
    }

    pub fn get_namespace(&self) -> Option<HashMap<String, String>> {
        if let Some(namespace_collection) = self.namespace_collection_ref.upgrade() {
            return Some(namespace_collection.borrow().clone());
        }
        None
    }

    pub fn get_value(&self) -> &Option<String> {
        &self.value
    }

    pub fn get_id(&self) -> usize {
        self.id
    }

    pub fn get_first_child_id(&self) -> Option<usize> {
        if let Some(child_element) = self.children.borrow().get(0) {
            return Some(child_element.id);
        }
        None
    }

    pub fn get_parent_id(&self) -> usize {
        self.parent_id
    }
}

// ########################## Data Write Methods ###########################
impl XmlElement {
    /// Remove the child reference irreversible
    pub fn pop_child_id_mut(&self) -> Option<usize> {
        if self.children.borrow_mut().len() > 0 {
            return Some(self.children.borrow_mut().remove(0).id);
        }
        None
    }

    fn append_children_mut(&mut self, child_id: usize, tag: &str) {
        self.children.borrow_mut().push(XmlElementChild {
            id: child_id,
            tag: tag.to_string(),
        });
    }

    fn insert_children_before_tag_mut(&mut self, child_id: usize, find_tag: &str, tag: &str) {
        let mut children = self.children.borrow_mut();
        if let Some(index) = children.iter().position(|item| item.tag == find_tag) {
            children.insert(
                index,
                XmlElementChild {
                    id: child_id,
                    tag: tag.to_string(),
                },
            );
        } else {
            children.push(XmlElementChild {
                id: child_id,
                tag: tag.to_string(),
            });
        }
    }

    fn insert_children_at_mut(&mut self, child_id: usize, position: usize, tag: &str) {
        self.children.borrow_mut().insert(
            position,
            XmlElementChild {
                id: child_id,
                tag: tag.to_string(),
            },
        );
    }

    pub fn get_attribute_mut(&mut self) -> Option<&mut HashMap<String, String>> {
        self.attributes.as_mut()
    }

    fn set_id_mut(&mut self, id: usize) {
        self.id = id;
    }

    fn set_parent_id_mut(&mut self, parent_id: usize) {
        self.parent_id = parent_id;
    }

    fn validate_namespace(
        &mut self,
        attributes: &HashMap<String, String>,
    ) -> AnyResult<(), AnyError> {
        Ok(())
    }

    pub fn set_attribute_mut(
        &mut self,
        mut attributes: HashMap<String, String>,
    ) -> AnyResult<&mut Self, AnyError> {
        let keys = attributes
            .keys()
            .clone()
            .map(|item| item.to_string())
            .collect::<Vec<String>>();
        if !keys
            .iter()
            .map(|item| item.to_string())
            .all(|item| Self::is_valid_xml_name(&item))
        {
            return Err(anyhow!("Not All the attribute satisfy naming standards"));
        }
        let ns_keys = keys
            .iter()
            .map(|item| item.to_string())
            .filter(|item| item.starts_with("xmlns"))
            .clone()
            .collect::<Vec<String>>();
        if !ns_keys.is_empty() {
            if let Some(ns_collection) = self.namespace_collection_ref.upgrade() {
                let mut namespace_collection = ns_collection
                    .try_borrow_mut()
                    .context("Namespace Collection Borrow Failed")?;
                for ns_key in ns_keys.clone() {
                    let ns_alias = ns_key.split(":").collect::<Vec<&str>>();
                    if ns_alias.len() > 1 {
                        if !namespace_collection.contains_key(&ns_alias[1].to_string()) {
                            namespace_collection.insert(
                                ns_alias[1].to_string(),
                                attributes.get(&ns_key).unwrap().to_string(),
                            );
                        }
                    } else {
                        namespace_collection.insert(
                            "<Default>".to_string(),
                            attributes.get(&ns_key).unwrap().to_string(),
                        );
                    }
                }
            }
        }
        attributes.retain(|key, _| !ns_keys.contains(key));
        self.validate_namespace(&attributes)
            .context("Attribute NameSpace Validation Failed")?;
        self.attributes = Some(attributes);
        Ok(self)
    }

    pub fn set_value(&mut self, text: String) -> &mut Self {
        self.value = Some(text);
        self
    }
}
