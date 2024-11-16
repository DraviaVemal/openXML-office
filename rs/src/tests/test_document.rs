use crate::document_2007::Word;

#[test]
fn blank_document() {
    let file = Word::new(None, Word::default()).expect("Blank File Open Failed");
    file.save_as(&"test.docx".to_string())
        .expect("Save Result Failed");
    assert_eq!(true, true);
}

#[test]
fn edit_document() {
    let file = Word::new(
        Some(
            "/home/draviavemal/repo/OpenXML-Office/rs/document/src/tests/test_file.docx"
                .to_string(),
        ),
        Word::default(),
    )
    .expect("Edit existing file failed");
    file.save_as(&"test.docx".to_string())
        .expect("Save Result Failed");
    assert_eq!(true, true);
}
