use crate::{
    files::{OfficeDocument, XmlDocument, XmlSerializer},
    global_2007::{
        parts::{RelationsPart, ThemePart},
        traits::{XmlDocumentPart, XmlDocumentPartCommon},
    },
    reference_dictionary::EXCEL_TYPE_COLLECTION,
    spreadsheet_2007::{
        parts::WorkSheet,
        services::{CalculationChainPart, CommonServices, ShareStringPart, StylePart},
    },
};
use anyhow::{anyhow, Context, Error as AnyError, Result as AnyResult};
use std::{
    cell::RefCell,
    collections::HashMap,
    rc::{Rc, Weak},
};

#[derive(Debug)]
pub struct WorkbookPart {
    office_document: Weak<RefCell<OfficeDocument>>,
    xml_document: Weak<RefCell<XmlDocument>>,
    file_path: String,
    common_service: Rc<RefCell<CommonServices>>,
    workbook_relationship_part: Rc<RefCell<RelationsPart>>,
    theme_part: ThemePart,
    /// This contain the sheet name, relationId, active sheet, hide sheet
    sheet_collection: Rc<RefCell<Vec<(String, String, bool, bool)>>>,
    workbook_view: Option<WorkbookView>,
}

#[derive(Debug)]
pub(crate) struct WorkbookView {
    first_sheet: String,
    active_tab: String,
    visibility: String,
    minimize: bool,
    hide_horizontal_scroll: bool,
    hide_vertical_scroll: bool,
    hide_sheet_tab: bool,
    hide_ruler: bool,
    hide_grid_line: bool,
}

impl WorkbookView {
    pub(crate) fn default() -> Self {
        WorkbookView {
            active_tab: "0".to_string(),
            first_sheet: "1".to_string(),
            minimize: false,
            hide_grid_line: false,
            hide_ruler: false,
            hide_sheet_tab: false,
            hide_vertical_scroll: false,
            hide_horizontal_scroll: false,
            visibility: "visible".to_string(),
        }
    }
}

impl Drop for WorkbookPart {
    fn drop(&mut self) {
        let _ = self.close_document();
    }
}

impl XmlDocumentPartCommon for WorkbookPart {
    /// Initialize xml content for this part from base template
    fn initialize_content_xml(
    ) -> AnyResult<(XmlDocument, Option<String>, Option<String>, Option<String>), AnyError> {
        let content = EXCEL_TYPE_COLLECTION.get("workbook").unwrap();
        let template_core_properties = include_str!("workbook.xml");
        Ok((
            XmlSerializer::vec_to_xml_doc_tree(template_core_properties.as_bytes().to_vec())
                .context("Initializing Workbook Failed")?,
            Some(content.content_type.to_string()),
            Some(content.extension.to_string()),
            Some(content.extension_type.to_string()),
        ))
    }

    fn close_document(&mut self) -> AnyResult<(), AnyError>
    where
        Self: Sized,
    {
        self.theme_part.close_document()?;
        self.common_service
            .try_borrow_mut()
            .context("Failed to pull common Service Handle")?
            .close_service()
            .context("Failed to Close Common Service From Workbook")?;
        self.workbook_relationship_part
            .try_borrow_mut()
            .context("Failed to pull relationship handle")?
            .close_document()
            .context("Failed to Close work")?;
        // Write Sheet Records to Workbook
        if let Some(xml_document_mut) = self.xml_document.upgrade() {
            let mut xml_doc_mut = xml_document_mut
                .try_borrow_mut()
                .context("Borrow XML Document Failed")?;
            let mut sheet_count = 1;
            if let Some(workbook_view) = &self.workbook_view {
                // Create and Set BookViews
                let book_views_id = xml_doc_mut
                    .insert_children_after_tag_mut("bookViews", "fileVersion", None)
                    .context("Create book viewsD Node Failed")?
                    .get_id();
                let workbook_view_element = xml_doc_mut
                    .append_child_mut("workbookView", Some(&book_views_id))
                    .context("Failed to create workbook view")?;
                let mut attributes = HashMap::new();
                if workbook_view.active_tab != "0" {
                    attributes.insert(
                        "activeTab".to_string(),
                        workbook_view.active_tab.to_string(),
                    );
                }
                if workbook_view.first_sheet != "1" {
                    attributes.insert(
                        "firstSheet".to_string(),
                        workbook_view.first_sheet.to_string(),
                    );
                }
                // attributes.insert(
                //     "minimized".to_string(),
                //     if workbook_view.minimize {
                //         "true".to_string()
                //     } else {
                //         "false".to_string()
                //     },
                // );
                // if workbook_view.visibility != "visible" {
                //     attributes.insert(
                //         "visibility".to_string(),
                //         workbook_view.visibility.to_string(),
                //     );
                // }
                // attributes.insert(
                //     "showGridlines".to_string(),
                //     if workbook_view.hide_grid_line {
                //         "false".to_string()
                //     } else {
                //         "true".to_string()
                //     },
                // );
                // attributes.insert(
                //     "showRuler".to_string(),
                //     if workbook_view.hide_ruler {
                //         "0".to_string()
                //     } else {
                //         "1".to_string()
                //     },
                // );
                attributes.insert(
                    "showSheetTabs".to_string(),
                    if workbook_view.hide_sheet_tab {
                        "0".to_string()
                    } else {
                        "1".to_string()
                    },
                );
                attributes.insert(
                    "showVerticalScroll".to_string(),
                    if workbook_view.hide_vertical_scroll {
                        "0".to_string()
                    } else {
                        "1".to_string()
                    },
                );
                attributes.insert(
                    "showHorizontalScroll".to_string(),
                    if workbook_view.hide_horizontal_scroll {
                        "0".to_string()
                    } else {
                        "1".to_string()
                    },
                );
                workbook_view_element
                    .set_attribute_mut(attributes)
                    .context("Failed to set workbook view attributes")?;
            }
            // Create and set Sheets
            let sheets_id = xml_doc_mut
                .insert_children_after_tag_mut("sheets", "bookViews", None)
                .context("Create Sheets Node Failed")?
                .get_id();
            for (sheet_display_name, relationship_id, _, hide) in &self
                .sheet_collection
                .try_borrow_mut()
                .context("Failed to pull Sheet Name Collection")?
                .clone()
            {
                let sheet = xml_doc_mut
                    .append_child_mut("sheet", Some(&sheets_id))
                    .context("Create Sheet Node Failed")?;
                let mut attributes = HashMap::new();
                attributes.insert("name".to_string(), sheet_display_name.to_string());
                attributes.insert("sheetId".to_string(), sheet_count.to_string());
                attributes.insert("r:id".to_string(), relationship_id.to_string());
                if *hide {
                    attributes.insert("state".to_string(), "hidden".to_string());
                }
                sheet
                    .set_attribute_mut(attributes)
                    .context("Sheet Attributes Failed")?;
                sheet_count += 1;
            }
        }
        if let Some(xml_tree) = self.office_document.upgrade() {
            xml_tree
                .try_borrow_mut()
                .context("Failed To Pull XML Handle")?
                .close_xml_document(&self.file_path)?;
        }
        Ok(())
    }
}

/// ######################### Train implementation of XML Part - Only accessible within crate ##############
impl XmlDocumentPart for WorkbookPart {
    /// Create workbook
    fn new(
        office_document: Weak<RefCell<OfficeDocument>>,
        parent_relationship_part: Weak<RefCell<RelationsPart>>,
        _: Option<&str>,
    ) -> AnyResult<Self, AnyError> {
        let file_name = Self::get_workbook_file_name(&parent_relationship_part)
            .context("Failed to pull workbook file name")?
            .to_string();
        let mut file_tree = Self::get_xml_document(&office_document, &file_name)?;
        let workbook_relationship_part = Rc::new(RefCell::new(
            RelationsPart::new(
                office_document.clone(),
                &format!(
                    "{}/_rels/workbook.xml.rels",
                    &file_name[..file_name.rfind("/").unwrap()]
                ),
            )
            .context("Creating Relation ship part for workbook failed.")?,
        ));
        // Theme
        let theme_part = ThemePart::new(
            office_document.clone(),
            Rc::downgrade(&workbook_relationship_part),
            None,
        )
        .context("Loading Theme Part Failed")?;
        // Share String
        let share_string = ShareStringPart::new(
            office_document.clone(),
            Rc::downgrade(&workbook_relationship_part),
            None,
        )
        .context("Loading Share String Failed")?;
        // Calculation chain
        let calculation_chain = CalculationChainPart::new(
            office_document.clone(),
            Rc::downgrade(&workbook_relationship_part),
            None,
        )
        .context("Loading Calculation Chain Failed")?;
        // Style
        let style = StylePart::new(
            office_document.clone(),
            Rc::downgrade(&workbook_relationship_part),
            None,
        )
        .context("Loading Style Part Failed")?;
        let common_service = Rc::new(RefCell::new(CommonServices::new(
            calculation_chain,
            share_string,
            style,
        )));
        let (sheet_collection, workbook_view) =
            Self::load_sheet_names(&mut file_tree).context("Loading Sheet Names Failed")?;
        Ok(Self {
            office_document,
            xml_document: file_tree,
            file_path: file_name,
            common_service,
            workbook_relationship_part,
            theme_part,
            sheet_collection: Rc::new(RefCell::new(sheet_collection)),
            workbook_view,
        })
    }
}

// ############################# Internal Function ######################################
// ############################# mut Function ######################################
impl WorkbookPart {
    fn load_sheet_names(
        xml_document: &mut Weak<RefCell<XmlDocument>>,
    ) -> AnyResult<(Vec<(String, String, bool, bool)>, Option<WorkbookView>), AnyError> {
        let mut sheet_collection = Vec::new();
        let mut workbook_view = None;
        if let Some(xml_doc) = xml_document.upgrade() {
            let mut xml_doc_mut = xml_doc.try_borrow_mut().context("xml doc borrow failed")?;
            // Deconstruct Book View for sheet collection data
            if let Some(mut book_views_vec) = xml_doc_mut.pop_elements_by_tag_mut("bookViews", None)
            {
                if let Some(book_views) = book_views_vec.pop() {
                    loop {
                        if let Some(workbook_view_id) = book_views.pop_child_id_mut() {
                            if let Some(workbook_view_element) =
                                xml_doc_mut.pop_element_mut(&workbook_view_id)
                            {
                                if let Some(attributes) = workbook_view_element.get_attribute() {
                                    workbook_view = Some(WorkbookView {
                                        active_tab: if let Some(active_tab) =
                                            attributes.get("activeTab")
                                        {
                                            active_tab.to_string()
                                        } else {
                                            "0".to_string()
                                        },
                                        first_sheet: if let Some(first_sheet) =
                                            attributes.get("firstSheet")
                                        {
                                            first_sheet.to_string()
                                        } else {
                                            "1".to_string()
                                        },
                                        hide_sheet_tab: if let Some(hide_sheet_tab) =
                                            attributes.get("showSheetTabs")
                                        {
                                            hide_sheet_tab == "1"
                                        } else {
                                            false
                                        },
                                        visibility: if let Some(visibility) =
                                            attributes.get("visibility")
                                        {
                                            visibility.to_string()
                                        } else {
                                            "visible".to_string()
                                        },
                                        minimize: if let Some(minimize) =
                                            attributes.get("minimized")
                                        {
                                            minimize == "1"
                                        } else {
                                            false
                                        },
                                        hide_grid_line: if let Some(hide_grid_line) =
                                            attributes.get("showGridlines")
                                        {
                                            hide_grid_line == "0"
                                        } else {
                                            false
                                        },
                                        hide_ruler: if let Some(hide_ruler) =
                                            attributes.get("showRuler")
                                        {
                                            hide_ruler == "0"
                                        } else {
                                            false
                                        },
                                        hide_horizontal_scroll: if let Some(
                                            hide_horizontal_scroll,
                                        ) =
                                            attributes.get("showHorizontalScroll")
                                        {
                                            hide_horizontal_scroll == "0"
                                        } else {
                                            false
                                        },
                                        hide_vertical_scroll: if let Some(hide_vertical_scroll) =
                                            attributes.get("showVerticalScroll")
                                        {
                                            hide_vertical_scroll == "0"
                                        } else {
                                            false
                                        },
                                    })
                                }
                            }
                        } else {
                            break;
                        }
                    }
                }
            }
            // Deconstruct Sheets into collection
            if let Some(mut sheets_vec) = xml_doc_mut.pop_elements_by_tag_mut("sheets", None) {
                if let Some(sheets) = sheets_vec.pop() {
                    // Load Sheet from File if exist
                    loop {
                        if let Some(sheet_id) = sheets.pop_child_id_mut() {
                            if let Some(sheet) = xml_doc_mut.pop_element_mut(&sheet_id) {
                                if let Some(attributes) = sheet.get_attribute() {
                                    let name = attributes.get("name").ok_or(anyhow!(
                                        "Error When Trying to read Sheet Details."
                                    ))?;
                                    let r_id = attributes.get("r:id").ok_or(anyhow!(
                                        "Error When Trying to read Sheet Details."
                                    ))?;
                                    let state = attributes.get("state");
                                    sheet_collection.push((
                                        name.to_string(),
                                        r_id.to_string(),
                                        false,
                                        if let Some(state) = state {
                                            state == "hidden"
                                        } else {
                                            false
                                        },
                                    ));
                                }
                            }
                        } else {
                            break;
                        }
                    }
                }
            }
        }
        Ok((sheet_collection, workbook_view))
    }
}

// ############################# im-mut Function ######################################
impl WorkbookPart {
    fn get_workbook_file_name(
        relations_part: &Weak<RefCell<RelationsPart>>,
    ) -> AnyResult<String, AnyError> {
        let relationship_content = EXCEL_TYPE_COLLECTION.get("workbook").unwrap();
        if let Some(relations_part) = relations_part.upgrade() {
            Ok(relations_part
                .try_borrow_mut()
                .context("Failed to pull relationship connection")?
                .get_relationship_target_by_type_mut(
                    &relationship_content.schemas_type,
                    relationship_content,
                    None,
                    None,
                )
                .context("Pull Path From Existing File Failed")?)
        } else {
            Err(anyhow!("Failed to upgrade relation part"))
        }
    }
}

// ############################# Feature Function ######################################
// ############################# mut Function ######################################
impl WorkbookPart {
    pub(crate) fn add_sheet_mut(
        &mut self,
        sheet_name: Option<String>,
    ) -> AnyResult<WorkSheet, AnyError> {
        Ok(WorkSheet::new(
            self.office_document.clone(),
            Rc::downgrade(&self.sheet_collection),
            Rc::downgrade(&self.workbook_relationship_part),
            Rc::downgrade(&self.common_service),
            sheet_name,
        )
        .context("Worksheet Creation Failed")?)
    }

    pub(crate) fn get_worksheet_mut(&mut self, sheet_name: &str) -> AnyResult<WorkSheet, AnyError> {
        WorkSheet::new(
            self.office_document.clone(),
            Rc::downgrade(&self.sheet_collection),
            Rc::downgrade(&self.workbook_relationship_part),
            Rc::downgrade(&self.common_service),
            Some(sheet_name.to_string()),
        )
        .context("Worksheet Creation Failed")
    }

    /// Set Active sheet on opening the excel
    pub(crate) fn set_active_sheet_mut(&mut self, sheet_name: &str) -> AnyResult<(), AnyError> {
        for (current_sheet_name, _, active_sheet, _) in self
            .sheet_collection
            .try_borrow_mut()
            .context("Failed to pull Sheet Collection Handle")?
            .iter_mut()
        {
            if current_sheet_name == sheet_name {
                *active_sheet = true
            } else {
                *active_sheet = false
            }
        }
        Ok(())
    }

    // /// Set workbook visibility
    // pub(crate) fn set_visibility_mut(&mut self, is_visible: bool) -> AnyResult<(), AnyError> {
    //     if let Some(workbook_view) = &mut self.workbook_view {
    //         workbook_view.visibility = if is_visible {
    //             "visible".to_string()
    //         } else {
    //             "hidden".to_string()
    //         }
    //     } else {
    //         self.workbook_view = Some(WorkbookView {
    //             visibility: if is_visible {
    //                 "visible".to_string()
    //             } else {
    //                 "hidden".to_string()
    //             },
    //             ..WorkbookView::default()
    //         })
    //     }
    //     Ok(())
    // }

    // /// Set workbook minimized
    // pub(crate) fn minimize_workbook_mut(&mut self, is_minimized: bool) -> AnyResult<(), AnyError> {
    //     if let Some(workbook_view) = &mut self.workbook_view {
    //         workbook_view.minimize = is_minimized;
    //     } else {
    //         self.workbook_view = Some(WorkbookView {
    //             minimize: is_minimized,
    //             ..WorkbookView::default()
    //         })
    //     }
    //     Ok(())
    // }

    /// Set visibility of sheet tabs in workbook
    pub(crate) fn hide_sheet_tabs_mut(&mut self, hide_sheet_tab: bool) -> AnyResult<(), AnyError> {
        if let Some(workbook_view) = &mut self.workbook_view {
            workbook_view.hide_sheet_tab = hide_sheet_tab;
        } else {
            self.workbook_view = Some(WorkbookView {
                hide_sheet_tab,
                ..WorkbookView::default()
            })
        }
        Ok(())
    }

    /// Set workbook ruler Visibility
    pub(crate) fn hide_ruler_mut(&mut self, hide_ruler: bool) -> AnyResult<(), AnyError> {
        if let Some(workbook_view) = &mut self.workbook_view {
            workbook_view.hide_ruler = hide_ruler;
        } else {
            self.workbook_view = Some(WorkbookView {
                hide_ruler,
                ..WorkbookView::default()
            })
        }
        Ok(())
    }

    /// Set workbook grid line Visibility
    pub(crate) fn hide_grid_lines_mut(&mut self, hide_grid_line: bool) -> AnyResult<(), AnyError> {
        if let Some(workbook_view) = &mut self.workbook_view {
            workbook_view.hide_grid_line = hide_grid_line;
        } else {
            self.workbook_view = Some(WorkbookView {
                hide_grid_line,
                ..WorkbookView::default()
            })
        }
        Ok(())
    }

    /// Set workbook Vertical Scroll Visibility
    pub(crate) fn hide_vertical_scroll_mut(
        &mut self,
        hide_vertical_scroll: bool,
    ) -> AnyResult<(), AnyError> {
        if let Some(workbook_view) = &mut self.workbook_view {
            workbook_view.hide_vertical_scroll = hide_vertical_scroll;
        } else {
            self.workbook_view = Some(WorkbookView {
                hide_vertical_scroll,
                ..WorkbookView::default()
            })
        }
        Ok(())
    }

    /// Set workbook Horizontal Scroll Visibility
    pub(crate) fn hide_horizontal_scroll_mut(
        &mut self,
        hide_horizontal_scroll: bool,
    ) -> AnyResult<(), AnyError> {
        if let Some(workbook_view) = &mut self.workbook_view {
            workbook_view.hide_horizontal_scroll = hide_horizontal_scroll;
        } else {
            self.workbook_view = Some(WorkbookView {
                hide_horizontal_scroll,
                ..WorkbookView::default()
            })
        }
        Ok(())
    }

    /// Set Active sheet on opening the excel
    pub(crate) fn hide_sheet_mut(&mut self, sheet_name: &str) -> AnyResult<(), AnyError> {
        for (current_sheet_name, _, _, hide_sheet) in self
            .sheet_collection
            .try_borrow_mut()
            .context("Failed to pull Sheet Collection Handle")?
            .iter_mut()
        {
            if current_sheet_name == sheet_name {
                *hide_sheet = true
            } else {
                *hide_sheet = false
            }
        }
        Ok(())
    }

    pub(crate) fn rename_sheet_name_mut(
        &mut self,
        old_sheet_name: &str,
        new_sheet_name: &str,
    ) -> AnyResult<(), AnyError> {
        // Check if sheet with same name exist
        if self
            .sheet_collection
            .try_borrow()
            .context("Failed to pull Sheet Name Collection")?
            .iter()
            .any(|item| new_sheet_name == item.0)
        {
            Err(anyhow!("New Sheet Name Already exist in the stack"))
        } else {
            if let Some(record) = self
                .sheet_collection
                .try_borrow_mut()
                .context("Failed to pull Sheet Name Collection")?
                .iter_mut()
                .find(|item| item.0 == old_sheet_name)
            {
                record.0 = new_sheet_name.to_string();
                Ok(())
            } else {
                Err(anyhow!("Old Sheet Name not found in the stack"))
            }
        }
    }
}

// ############################# im-mut Function ######################################
impl WorkbookPart {
    pub(crate) fn list_sheet_names(&self) -> AnyResult<Vec<String>, AnyError> {
        Ok(self
            .sheet_collection
            .try_borrow()
            .context("Failed to pull Sheet Name Collection")?
            .iter()
            .map(|(sheet_name, _, _, _)| sheet_name.to_string())
            .collect::<Vec<String>>())
    }
}
