#[test]
fn blank_excel() {
    let file = crate::Excel::new(None, crate::ExcelPropertiesModel { is_in_memory: true });
    file.add_sheet(&"Test".to_string());
    file.save_as(&"this.xlsx".to_string());
    assert_eq!(true, true);
}
#[test]
fn edit_excel() {
    let file = crate::Excel::new(
        Some(
            "/home/draviavemal/repo/OpenXML-Office/rs/spreadsheet/src/tests/test_file.xlsx"
                .to_string(),
        ),
        crate::ExcelPropertiesModel { is_in_memory: true },
    );
    file.add_sheet(&"Test".to_string());
    file.save_as(&"this.xlsx".to_string());
    assert_eq!(true, true);
}
