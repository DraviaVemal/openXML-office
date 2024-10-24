use crate::enums::common::NodeContentType;

pub(crate) struct CurrentNode {
    name: String,
    content_type: NodeContentType,
    childs: Option<Vec<CurrentNode>>,
}
