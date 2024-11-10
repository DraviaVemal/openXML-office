use crate::Word;

#[test]
fn blank_document() {
    let file = Word::new(None, Word::default());
    file.save_as(&"test.docx".to_string());
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
    );
    file.save_as(&"test.docx".to_string());
    assert_eq!(true, true);
}
