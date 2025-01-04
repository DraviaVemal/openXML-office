use crate::global_2007::traits::XmlDocumentPartCommon;
use chrono::Utc;
use std::fs::{create_dir, exists};

fn get_save_file(dynamic_path: Option<&str>) -> String {
    let result_path = "test_results";
    if !exists(result_path).expect("Dir Check Failed") {
        create_dir(result_path).expect("Failed to Create")
    }
    format!(
        "{}/test-{}{}.xlsx",
        result_path,
        dynamic_path.unwrap_or(""),
        Utc::now().format("%Y-%m-%d-%H-%M-%S").to_string()
    )
}

#[test]
fn blank_excel() {
    let file = crate::spreadsheet_2007::Excel::new(
        None,
        crate::spreadsheet_2007::ExcelPropertiesModel {
            is_in_memory: false,
            is_editable: true,
        },
    )
    .expect("Create New File Failed");
    file.save_as(&get_save_file(None))
        .expect("File Save Failed");
    assert_eq!(true, true);
}

#[test]
fn excel_handling() {
    let mut file = crate::spreadsheet_2007::Excel::new(
        None,
        crate::spreadsheet_2007::ExcelPropertiesModel::default(),
    )
    .expect("Create New File Failed");
    file.add_sheet_mut(Some("Test".to_string()))
        .expect("Failed to add static Sheet");
    file.add_sheet_mut(Some("bust".to_string()))
        .expect("Failed to add static Sheet");
    file.add_sheet_mut(None)
        .expect("Failed to add static Sheet");
    file.add_sheet_mut(Some("Active".to_string()))
        .expect("Failed to add static Sheet");
    file.set_active_sheet_mut("Active".to_string())
        .expect("Failed To Set Active Sheet");
    file.add_sheet_mut(Some("hideSheet".to_string()))
        .expect("Failed to add static Sheet");
    file.hide_sheet_mut("hideSheet".to_string())
        .expect("Failed to hide the sheet");
    file.save_as(&get_save_file(None))
        .expect("File Save Failed");
    assert_eq!(true, true);
}

#[test]
fn sheet_handling() {
    let mut file = crate::spreadsheet_2007::Excel::new(
        None,
        crate::spreadsheet_2007::ExcelPropertiesModel {
            is_in_memory: false,
            ..crate::spreadsheet_2007::ExcelPropertiesModel::default()
        },
    )
    .expect("Create New File Failed");
    file.add_sheet_mut(Some("Test".to_string()))
        .expect("Failed to add static Sheet");
    file.add_sheet_mut(Some("bust".to_string()))
        .expect("Failed to add static Sheet");
    file.add_sheet_mut(Some("RenameThisSheet".to_string()))
        .expect("Failed to add static Sheet");
    let close_sheet = file
        .add_sheet_mut(Some("deletethis".to_string()))
        .expect("Failed to add static Sheet");
    close_sheet.flush().expect("Failed to Close Work Sheet");
    file.add_sheet_mut(None)
        .expect("Failed to add dynamic Sheet");
    let delete_sheet = file
        .get_worksheet_mut("deletethis".to_string())
        .expect("Failed to Get the Worksheet");
    delete_sheet
        .delete_sheet_mut()
        .expect("Failed to Delete Sheet");
    file.rename_sheet_name_mut("RenameThisSheet".to_string(), "RenamedSheet".to_string())
        .expect("Failed to rename the sheet");
    file.save_as(&get_save_file(None))
        .expect("File Save Failed");
    assert_eq!(true, true);
}

// #[test]
// fn excel_view() {
//     // let mut file = crate::spreadsheet_2007::Excel::new(
//     //     None,
//     //     crate::spreadsheet_2007::ExcelPropertiesModel::default(),
//     // )
//     // .expect("Create New File Failed");
//     // file.minimize_workbook_mut(true)
//     //     .expect("Failed to minimize workbook");
//     // file.save_as(&format!("{}", &get_save_file(Some("min"))))
//     //     .expect("File Save Failed");
//     // file = crate::spreadsheet_2007::Excel::new(
//     //     None,
//     //     crate::spreadsheet_2007::ExcelPropertiesModel::default(),
//     // )
//     // .expect("Create New File Failed");
//     // file.set_visibility_mut(false)
//     //     .expect("Failed to add static Sheet");
//     // file.save_as(&format!("{}", &get_save_file(Some("hide"))))
//     //     .expect("File Save Failed");

//     let mut file = crate::spreadsheet_2007::Excel::new(
//         None,
//         crate::spreadsheet_2007::ExcelPropertiesModel::default(),
//     )
//     .expect("Create New File Failed");
//     file.hide_ruler_mut(true)
//         .expect("Failed to add static Sheet");
//     file.save_as(&format!("{}", &get_save_file(Some("show_ruler"))))
//         .expect("File Save Failed");
//     file = crate::spreadsheet_2007::Excel::new(
//         None,
//         crate::spreadsheet_2007::ExcelPropertiesModel::default(),
//     )
//     .expect("Create New File Failed");
//     file.hide_grid_lines_mut(true)
//         .expect("Failed to add static Sheet");
//     file.save_as(&format!("{}", &get_save_file(Some("show_grid"))))
//         .expect("File Save Failed");
//     file = crate::spreadsheet_2007::Excel::new(
//         None,
//         crate::spreadsheet_2007::ExcelPropertiesModel::default(),
//     )
//     .expect("Create New File Failed");
//     file.hide_horizontal_scroll_mut(true)
//         .expect("Failed to add static Sheet");
//     file.save_as(&format!("{}", &get_save_file(Some("show_hor_scroll"))))
//         .expect("File Save Failed");
//     file = crate::spreadsheet_2007::Excel::new(
//         None,
//         crate::spreadsheet_2007::ExcelPropertiesModel::default(),
//     )
//     .expect("Create New File Failed");
//     file.hide_vertical_scroll_mut(true)
//         .expect("Failed to add static Sheet");
//     file.save_as(&format!("{}", &get_save_file(Some("show_ver_scroll"))))
//         .expect("File Save Failed");
//     file = crate::spreadsheet_2007::Excel::new(
//         None,
//         crate::spreadsheet_2007::ExcelPropertiesModel::default(),
//     )
//     .expect("Create New File Failed");
//     file.hide_sheet_tabs_mut(true)
//         .expect("Failed to add static Sheet");
//     file.save_as(&format!("{}", &get_save_file(Some("show_tab"))))
//         .expect("File Save Failed");
//     assert_eq!(true, true);
// }

#[test]
fn set_row_property() {
    let mut file = crate::spreadsheet_2007::Excel::new(
        Some("src/tests/TestFiles/basic_test.xlsx".to_string()),
        crate::spreadsheet_2007::ExcelPropertiesModel {
            is_in_memory: false,
            is_editable: true,
        },
    )
    .expect("Open Existing File Failed");
    let mut row_prop = file
        .add_sheet_mut(Some("row_property".to_string()))
        .expect("failed to add Sheet");
    row_prop
        .set_row_index_properties_mut(
            &1,
            Some(crate::spreadsheet_2007::models::RowProperties {
                height: Some(100 as f32),
                ..crate::spreadsheet_2007::models::RowProperties::default()
            }),
        )
        .expect("Failed to set row height");
    row_prop
        .set_row_index_properties_mut(
            &3,
            Some(crate::spreadsheet_2007::models::RowProperties {
                hidden: Some(true),
                ..crate::spreadsheet_2007::models::RowProperties::default()
            }),
        )
        .expect("Failed to set row height");
    row_prop
        .set_row_index_properties_mut(
            &5,
            Some(crate::spreadsheet_2007::models::RowProperties {
                thick_bottom: Some(true),
                ..crate::spreadsheet_2007::models::RowProperties::default()
            }),
        )
        .expect("Failed to set row height");
    row_prop
        .set_row_index_properties_mut(
            &7,
            Some(crate::spreadsheet_2007::models::RowProperties {
                thick_top: Some(true),
                ..crate::spreadsheet_2007::models::RowProperties::default()
            }),
        )
        .expect("Failed to set row height");
    row_prop.flush().expect("Failed to write Data");
    file.save_as(&get_save_file(None))
        .expect("Save File Failed");
    assert_eq!(true, true);
}

#[test]
fn set_column_property() {
    let mut file = crate::spreadsheet_2007::Excel::new(
        Some("src/tests/TestFiles/basic_test.xlsx".to_string()),
        crate::spreadsheet_2007::ExcelPropertiesModel {
            is_in_memory: false,
            is_editable: true,
        },
    )
    .expect("Open Existing File Failed");
    let mut col_prop = file
        .add_sheet_mut(Some("col_property".to_string()))
        .expect("failed to add Sheet");
    col_prop
        .set_column_index_properties_mut(
            &1,
            Some(crate::spreadsheet_2007::models::ColumnProperties {
                width: Some(200 as f32),
                ..crate::spreadsheet_2007::models::ColumnProperties::default()
            }),
        )
        .expect("Failed to Set Column prop");
    col_prop
        .set_column_index_properties_mut(
            &3,
            Some(crate::spreadsheet_2007::models::ColumnProperties {
                hidden: Some(true),
                ..crate::spreadsheet_2007::models::ColumnProperties::default()
            }),
        )
        .expect("Failed to Set Column prop");
    col_prop
        .set_column_index_properties_mut(
            &5,
            Some(crate::spreadsheet_2007::models::ColumnProperties {
                best_fit: Some(true),
                ..crate::spreadsheet_2007::models::ColumnProperties::default()
            }),
        )
        .expect("Failed to Set Column prop");
    col_prop.flush().expect("Failed to write Data");
    file.save_as(&get_save_file(None))
        .expect("Save File Failed");
    assert_eq!(true, true);
}

#[test]
fn edit_excel() {
    let mut file = crate::spreadsheet_2007::Excel::new(
        Some("src/tests/TestFiles/basic_test.xlsx".to_string()),
        crate::spreadsheet_2007::ExcelPropertiesModel {
            is_in_memory: false,
            is_editable: true,
        },
    )
    .expect("Open Existing File Failed");
    {
        let mut formula = file
            .get_worksheet_mut("formula".to_string())
            .expect("Failed to find the worksheet");
        formula
            .set_row_value_ref_mut(
                "V3",
                vec![
                    crate::spreadsheet_2007::models::ColumnCell {
                        value: Some("Dravia".to_string()),
                        data_type: crate::spreadsheet_2007::models::CellDataType::Auto,
                        ..crate::spreadsheet_2007::models::ColumnCell::default()
                    },
                    crate::spreadsheet_2007::models::ColumnCell {
                        value: Some("Vemal".to_string()),
                        data_type: crate::spreadsheet_2007::models::CellDataType::Auto,
                        ..crate::spreadsheet_2007::models::ColumnCell::default()
                    },
                ],
            )
            .expect("Failed To Set Row Value");
    }
    file.get_worksheet_mut("Style".to_string())
        .expect("Failed to find the worksheet");
    file.save_as(&get_save_file(None))
        .expect("Save File Failed");
    assert_eq!(true, true);
}
