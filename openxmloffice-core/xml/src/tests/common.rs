#[test]
fn it_works() {
    let file = crate::open_file(
        "/home/draviavemal/repo/OpenXML-Office/openxmloffice-core/xml/Book1.xlsx".to_string(),
        true,
    );
    println!(
        "{}",
        file.temp_file
            .path()
            .to_str()
            .expect("Str Reference Failed")
    );
    assert_eq!(true, true);
}
