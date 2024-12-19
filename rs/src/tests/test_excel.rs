use crate::global_2007::traits::XmlDocumentPartCommon;
use chrono::Utc;
use std::fs::{create_dir, exists};

fn get_save_file() -> String {
    let result_path = "test_results";
    if !exists(result_path).expect("Dir Check Failed") {
        create_dir(result_path).expect("Failed to Create")
    }
    format!(
        "{}/test-{}.xlsx",
        result_path,
        Utc::now().format("%Y-%m-%d-%H-%M-%S").to_string()
    )
}

#[test]
fn blank_excel() {
    let file = crate::spreadsheet_2007::Excel::new(
        None,
        crate::spreadsheet_2007::ExcelPropertiesModel {
            is_in_memory: false,
            is_editable: true,
        },
    )
    .expect("Create New File Failed");
    file.save_as(&get_save_file()).expect("File Save Failed");
    assert_eq!(true, true);
}

#[test]
fn sheet_handling() {
    let mut file = crate::spreadsheet_2007::Excel::new(
        None,
        crate::spreadsheet_2007::ExcelPropertiesModel::default(),
    )
    .expect("Create New File Failed");
    file.add_sheet(Some("Test".to_string()))
        .expect("Failed to add static Sheet");
    file.add_sheet(Some("bust".to_string()))
        .expect("Failed to add static Sheet");
    file.add_sheet(Some("RenameThisSheet".to_string()))
        .expect("Failed to add static Sheet");
    let close_sheet = file
        .add_sheet(Some("deletethis".to_string()))
        .expect("Failed to add static Sheet");
    close_sheet.flush().expect("Failed to Close Work Sheet");
    let delete_sheet = file
        .get_worksheet("deletethis".to_string())
        .expect("Failed to Get the Worksheet");
    delete_sheet
        .delete_sheet_mut()
        .expect("Failed to Delete Sheet");
    file.add_sheet(None).expect("Failed to add dynamic Sheet");
    file.rename_sheet_name("RenameThisSheet".to_string(), "RenamedSheet".to_string())
        .expect("Failed to rename the sheet");
    file.save_as(&get_save_file()).expect("File Save Failed");
    assert_eq!(true, true);
}

#[test]
fn edit_excel() {
    let file = crate::spreadsheet_2007::Excel::new(
        Some("src/tests/TestFiles/basic_test.xlsx".to_string()),
        crate::spreadsheet_2007::ExcelPropertiesModel {
            is_in_memory: false,
            is_editable: true,
        },
    )
    .expect("Open Existing File Failed");
    // file.add_sheet(&"Test".to_string());
    file.save_as(&get_save_file()).expect("Save File Failed");
    assert_eq!(true, true);
}
