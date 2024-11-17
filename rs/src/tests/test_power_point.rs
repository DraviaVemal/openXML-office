use crate::presentation_2007::{PowerPoint, PowerPointPropertiesModel};

#[test]
fn blank_power_point() {
    let file = PowerPoint::new(None, PowerPoint::default()).expect("Create New File Failed");
    file.save_as(&"test.pptx".to_string())
        .expect("Failed to save Empty Power Point");
    assert_eq!(true, true);
}

#[test]
fn edit_power_point() {
    let file = PowerPoint::new(
        Some(
            "/home/draviavemal/repo/OpenXML-Office/rs/presentation/src/tests/test_file.pptx"
                .to_string(),
        ),
        PowerPoint::default(),
    )
    .expect("Open Existing file failed");
    file.save_as(&"test.pptx".to_string())
        .expect("Failed to save Edit Power Point");
    assert_eq!(true, true);
}
