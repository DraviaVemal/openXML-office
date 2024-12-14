#[test]
fn blank_excel() {
    let file = crate::spreadsheet_2007::Excel::new(
        None,
        crate::spreadsheet_2007::ExcelPropertiesModel {
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
    let file = crate::spreadsheet_2007::Excel::new(
        Some("src/tests/TestFiles/basic_test.xlsx".to_string()),
        crate::spreadsheet_2007::ExcelPropertiesModel {
            is_in_memory: false,
        },
    )
    .expect("Open Existing File Failed");
    // file.add_sheet(&"Test".to_string());
    file.save_as(&"test.xlsx".to_string())
        .expect("Save File Failed");
    assert_eq!(true, true);
}
