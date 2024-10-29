use crate::{structs::worksheet::Worksheet, Excel};

impl<'excel> Worksheet<'excel> {
    /// Create New object for the group
    pub fn new(excel: &'excel Excel, sheet_name: Option<&str>) -> Self {
        if let Some(sheet_name) = sheet_name {
            
        } else { // Auto generate Sheet name
        }
        return Self { excel };
    }
    /// Set active status for the current worksheet
    pub fn set_active_sheet(&self, is_active: bool) {}
}
