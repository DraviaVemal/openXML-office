use crate::spreadsheet_2007::{Excel, ExcelPropertiesModel};

#[test]
fn blank_excel() {
    let file = Excel::new(
        None,
        ExcelPropertiesModel {
            is_in_memory: false,
        },
    )
    .expect("Create New File Failed");
    // file.add_sheet(&"Test".to_string());
    file.save_as(&"test.xlsx".to_string())
        .expect("File Save Failed");
    assert_eq!(true, true);
}

#[test]
fn edit_excel() {
    let file = Excel::new(
        Some("src/tests/TestFiles/basic_test.xlsx".to_string()),
        ExcelPropertiesModel {
            is_in_memory: false,
        },
    )
    .expect("Open Existing File Failed");
    // file.add_sheet(&"Test".to_string());
    file.save_as(&"test.xlsx".to_string())
        .expect("Save File Failed");
    assert_eq!(true, true);
}
