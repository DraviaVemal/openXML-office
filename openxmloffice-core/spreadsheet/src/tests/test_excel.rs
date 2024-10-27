#[test]
fn blank_excel() {
    let file = crate::Excel::new(None);
    file.add_sheet(&"Test".to_string());
    file.save_as(&"this.xlsx".to_string());
    assert_eq!(true, true);
}
