#[test]
fn blank_document() {
    let file = crate::document_2007::Word::new(None, crate::document_2007::Word::default())
        .expect("Blank File Open Failed");
    file.save_as(&"test.docx".to_string())
        .expect("Save Result Failed");
    assert_eq!(true, true);
}

#[test]
fn edit_document() {
    let file = crate::document_2007::Word::new(
        Some(
            "/home/draviavemal/repo/OpenXML-Office/rs/document/src/tests/test_file.docx"
                .to_string(),
        ),
        crate::document_2007::Word::default(),
    )
    .expect("Edit existing file failed");
    file.save_as(&"test.docx".to_string())
        .expect("Save Result Failed");
    assert_eq!(true, true);
}
