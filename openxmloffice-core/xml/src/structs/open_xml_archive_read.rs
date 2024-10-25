use std::{cell::RefCell, rc::Rc};

pub struct OpenXmlNonEditable<'buffer> {
    pub(crate) working_buffer: &'buffer Rc<RefCell<Vec<u8>>>,
}
