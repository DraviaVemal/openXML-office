use anyhow::{anyhow, Context, Error as AnyError, Result as AnyResult};
use std::{
    cell::RefCell,
    collections::HashMap,
    rc::{Rc, Weak},
};

#[derive(Debug)]
pub struct XmlDocument {
    document_tree: Rc<RefCell<XmlDocumentTree>>,
    running_id: usize,
    /// XML Element Collection
    xml_element_collection: HashMap<usize, XmlElement>,
}

/// XML Document In Execution
#[derive(Debug)]
pub struct XmlDocumentTree {
    namespace_collection: HashMap<String, String>,
    /// XML Tree Element
    xml_tree: HashMap<usize, XmlElementChildren>,
}

/// Normalized XML representation
#[derive(Debug, Clone)]
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
    /// XML Document Tree Reference
    document_tree_ref: Weak<RefCell<XmlDocumentTree>>,
}
///
#[derive(Debug)]
pub struct XmlElementChildren {
    /// Child Element Names to pull up nodes quickly
    children_tags: Vec<String>,
    /// Child Element Id that has full element
    children_id: Vec<usize>,
}

impl XmlDocument {
    pub fn new() -> Self {
        let document_tree: Rc<RefCell<XmlDocumentTree>> =
            Rc::new(RefCell::new(XmlDocumentTree::new()));
        Self {
            document_tree,
            running_id: 0,
            xml_element_collection: HashMap::new(),
        }
    }
    // #################### Non Editable Section ###################

    pub fn get_root(&self) -> Option<&XmlElement> {
        self.get_element(0)
    }

    pub fn find_child_by_attribute(
        &self,
        parent: &XmlElement,
        attribute_key: &str,
        attribute_value: &str,
    ) -> Option<&XmlElement> {
        if let Some(children) = self.document_tree.borrow().xml_tree.get(&parent.id) {
            for child_id in children.children_id.iter() {
                let element = self.get_element(*child_id).unwrap();
                if let Some(attributes) = element.get_attribute() {
                    let attr = attributes.get(attribute_key);
                    if let Some(val) = attr {
                        if val == attribute_value {
                            return Some(element);
                        }
                    }
                }
            }
        }
        None
    }

    //############### Editable Section ##############
    pub fn create_root(&mut self, element_tag: &str) -> &mut XmlElement {
        let root_element: XmlElement =
            XmlElement::new(Rc::downgrade(&self.document_tree), element_tag);
        self.insert_element_collection(0, root_element);
        self.get_element_mut(0).unwrap()
    }

    pub fn insert_element(&mut self, parent_id: &usize, element_tag: &str) -> &mut XmlElement {
        let mut child_element: XmlElement =
            XmlElement::new(Rc::downgrade(&self.document_tree), element_tag);
        let element_id = self.get_next_id();
        child_element.set_id(element_id);
        child_element.set_parent_id(*parent_id);
        self.document_tree
            .borrow_mut()
            .insert_tree_child(*parent_id, element_id, element_tag);
        self.insert_element_collection(element_id, child_element);
        self.get_element_mut(element_id).unwrap()
    }

    pub fn find_first_element_mut(&mut self, tag_path: &str) -> Option<&mut XmlElement> {
        if let Some(element_id) = self.find_first_element_id(tag_path) {
            return self.get_element_mut(element_id);
        }
        None
    }

    fn find_first_element_id(&self, tag_path: &str) -> Option<usize> {
        let mut found_id: Option<usize> = None;
        let mut path_parts = tag_path
            .split("->")
            .map(String::from)
            .collect::<Vec<String>>();
        path_parts.reverse();
        path_parts.pop();
        let xml_doc_tree = self.document_tree.borrow_mut();
        if let Some(tree) = xml_doc_tree.xml_tree.get(&0) {
            for tag in path_parts.iter() {
                let found_item = tree
                    .children_tags
                    .iter()
                    .position(|element_tag| element_tag == tag);
                if let Some(item) = found_item {
                    found_id = Some(tree.children_id[item]);
                }
            }
        }
        found_id
    }

    // ############### Actions Performed on xml_element_collection ############
    fn get_next_id(&mut self) -> usize {
        self.running_id += 1;
        self.running_id
    }

    fn insert_element_collection(&mut self, id: usize, xml_element: XmlElement) {
        self.xml_element_collection.insert(id, xml_element);
    }

    pub fn get_element_mut(&mut self, element_id: usize) -> Option<&mut XmlElement> {
        self.xml_element_collection.get_mut(&element_id)
    }

    pub fn get_element(&self, element_id: usize) -> Option<&XmlElement> {
        self.xml_element_collection.get(&element_id)
    }
}

/// Master XML document
/// This helps in keeping the XML validation clean across document for all relations
/// TODO: Clear Unused namespace
/// TODO: Give clear handle for attribute and Namespace handling of root element
impl XmlDocumentTree {
    /// Create new XML Document Tree.
    /// The Root element will be created by default.
    fn new() -> Self {
        Self {
            namespace_collection: HashMap::new(),
            xml_tree: HashMap::new(),
        }
    }

    fn insert_tree_child(&mut self, parent_id: usize, child_id: usize, element_tag: &str) {
        let element_child_option = self.xml_tree.get_mut(&parent_id);
        if let Some(element_child) = element_child_option {
            element_child.children_id.push(child_id);
            element_child.children_tags.push(element_tag.to_string());
        } else {
            let mut element_child = XmlElementChildren {
                children_tags: Vec::new(),
                children_id: Vec::new(),
            };
            element_child.children_id.push(child_id);
            element_child.children_tags.push(element_tag.to_string());
            self.xml_tree.insert(parent_id, element_child);
        }
    }

    fn pop_children(&mut self, id: usize) -> Option<usize> {
        if let Some(child) = self.xml_tree.get_mut(&id) {
            if child.children_tags.len() > 0 {
                child.children_tags.remove(0);
                return Some(child.children_id.remove(0));
            }
        }
        None
    }
}

impl XmlElement {
    /// Create element with tree document reference
    fn new(document_tree_ref: Weak<RefCell<XmlDocumentTree>>, tag: &str) -> Self {
        Self {
            id: 0,
            parent_id: 0,
            tag: tag.to_string(),
            attributes: None,
            value: None,
            document_tree_ref,
        }
    }
    // ######################## Data Read Only Methods ###########################

    pub fn get_tag(&self) -> &str {
        &self.tag
    }

    pub fn has_value(&self) -> bool {
        self.get_value().is_some()
    }

    pub fn is_empty_tag(&self) -> bool {
        self.attributes.is_none() && self.value.is_none()
    }

    pub fn get_attribute(&self) -> &Option<HashMap<String, String>> {
        &self.attributes
    }

    pub fn get_value(&self) -> &Option<String> {
        &self.value
    }

    pub fn get_id(&self) -> usize {
        self.id
    }

    pub fn pop_children(&self, id: usize) -> AnyResult<Option<usize>, AnyError> {
        if let Some(master_tree) = self.document_tree_ref.upgrade() {
            return Ok(master_tree
                .try_borrow_mut()
                .context("Document Tree is not valid")?
                .pop_children(id));
        }
        Err(anyhow!("Document Tree is no longer valid"))
    }

    pub fn get_parent_id(&self) -> usize {
        self.parent_id
    }
    // ########################## Data Write Methods ###########################
    fn set_id(&mut self, id: usize) {
        self.id = id;
    }

    fn set_parent_id(&mut self, parent_id: usize) {
        self.parent_id = parent_id;
    }

    pub fn set_attribute(&mut self, attributes: HashMap<String, String>) -> &mut Self {
        self.attributes = Some(attributes);
        self
    }

    pub fn set_value(&mut self, text: String) -> &mut Self {
        self.value = Some(text);
        self
    }
}
