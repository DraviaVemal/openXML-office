use crate::structs::{excel::Excel, worksheet::Worksheet};

impl Excel {
    pub fn new(file_name: Option<String>) -> Self {
        return Self {};
    }
    pub fn add_sheet(&self) -> Worksheet {
        return Worksheet::new();
    }

    pub fn save_as(file_name: String) {
        
    }
}
