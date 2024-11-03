use crate::structs::workbook::Workbook;
use openxmloffice_xml::{get_specific_queries, OpenXmlFile};
use rusqlite::{params, Row};

impl Workbook {
    /// Create workbook
    pub fn new(xml_fs: &OpenXmlFile) -> Self {
        let result = Self::get_workbook_xml(&xml_fs);
        return Self {};
    }

    /// Read and load workbook xml to work with
    fn get_workbook_xml(xml_fs: &OpenXmlFile) -> Vec<u8> {
        let query = get_specific_queries!("workbook.sql", "select_workbook")
            .expect("Workbook data query failed");
        fn row_mapper(row: &Row) -> Result<Vec<u8>, rusqlite::Error> {
            Ok(row.get(0)?)
        }
        let results = xml_fs
            .find_one(&query, params!["xl/workbook.xml"], row_mapper)
            .expect("Get workbook content failed");
        if let Some(results) = results {
            return results;
        } else {
            return Self::initialize_workbook_xml(&xml_fs);
        }
    }

    /// Initialize workbook for new excel
    fn initialize_workbook_xml(xml_fs: &OpenXmlFile) -> Vec<u8> {
        let query = get_specific_queries!("workbook.sql", "insert_workbook")
            .expect("Workbook data insert query failed");
        let content = Vec::new();
        xml_fs
            .execute_query(&query, params!["xl/workbook.xml", 0, 0, 0, "gzip", content])
            .expect("Workbook New Entry failed");
        return content;
    }
}
