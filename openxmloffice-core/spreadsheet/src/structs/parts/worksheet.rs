use crate::Excel;

pub struct Worksheet<'excel> {
    pub(crate) excel: &'excel Excel,
}
