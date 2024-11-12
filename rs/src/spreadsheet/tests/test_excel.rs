use crate::Excel;

#[test]
fn blank_excel() {
    let file = Excel::new(None, Excel::default()).expect("Create New File Failed");
    // file.add_sheet(&"Test".to_string());
    file.save_as(&"test.xlsx".to_string());
    assert_eq!(true, true);
}

#[test]
fn edit_excel() {
    let file = Excel::new(
        Some("/home/draviavemal/repo/OpenXML-Office/cs/Test/TestFiles/basic_test.xlsx".to_string()),
        Excel::default(),
    )
    .expect("Open Existing File Failed");
    // file.add_sheet(&"Test".to_string());
    file.save_as(&"test.xlsx".to_string());
    assert_eq!(true, true);
}
