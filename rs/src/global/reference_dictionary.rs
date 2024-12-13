use phf::{phf_map, Map};

pub struct Content {
    pub schemas_namespace: &'static str,
    pub schemas_type: &'static str,
    pub alias: &'static str,
    pub content_type: &'static str,
    pub extension: &'static str,
    pub default_path: &'static str,
    pub default_name: &'static str,
}
/// Common Content
pub static COMMON_TYPE_COLLECTION: Map<&'static str, &'static Content> = phf_map! {
    "content_type"=>&Content{
        schemas_namespace:"http://schemas.openxmlformats.org/package/2006/content-types",
        schemas_type:"",
        alias:"",
        content_type:"",
        extension:"xml",
        default_path:".",
        default_name:""
    },
    "xml"=>&Content{
        schemas_namespace:"",
        schemas_type:"",
        alias:"",
        content_type:"application/xml",
        extension:"xml",
        default_path:".",
        default_name:""
    },
    "rels"=>&Content{
        schemas_namespace:"http://schemas.openxmlformats.org/package/2006/relationships",
        schemas_type:"",
        alias:"r",
        content_type:"application/vnd.openxmlformats-package.relationships+xml",
        extension:"rels",
        default_path:".",
        default_name:""
    },
    "docProps_core"=>&Content{
        schemas_namespace:"http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
        schemas_type:"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
        alias:"cp",
        content_type:"application/vnd.openxmlformats-package.core-properties+xml",
        extension:"xml",
        default_path:"docProps",
        default_name:"core"
    },
    "theme"=>&Content{
        schemas_namespace:"",
        schemas_type:"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
        alias:"",
        content_type:"application/vnd.openxmlformats-officedocument.theme+xml",
        extension:"xml",
        default_path:"theme",
        default_name:"theme1"
    },
};
/// Excel Related Content
pub static EXCEL_TYPE_COLLECTION: Map<&'static str, &'static Content> = phf_map! {
    "style"=>&Content{
        schemas_namespace:"http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        schemas_type:"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
        alias:"x",
        content_type:"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
        extension:"xml",
        default_path:"xl",
        default_name:"styles"
    },
    "share_string"=>&Content{
        schemas_namespace:"http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        schemas_type:"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
        alias:"x",
        content_type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
        extension:"xml",
        default_path:"xl",
        default_name:"sharedStrings"
    },
    "calc_chain"=>&Content{
        schemas_namespace:"http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        schemas_type:"http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain",
        alias:"x",
        content_type:"application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml",
        extension:"xml",
        default_path:"xl",
        default_name:"calcChain"
    },
    "workbook"=>&Content{
        schemas_namespace:"http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        schemas_type:"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
        alias:"x",
        content_type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
        extension:"xml",
        default_path:"xl",
        default_name:"workbook"
    },
    "table"=>&Content{
        schemas_namespace:"http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        schemas_type:"http://schemas.openxmlformats.org/officeDocument/2006/relationships/table",
        alias:"x",
        content_type:"application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml",
        extension:"xml",
        default_path:"xl/tables",
        default_name:"table1"
    },
    "worksheet"=>&Content{
        schemas_namespace:"http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        schemas_type:"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
        alias:"x",
        content_type:"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
        extension:"xml",
        default_path:"xl/worksheets",
        default_name:"sheet1"
    },
};
/// Power Point Related Content
pub static POWER_POINT_TYPE_COLLECTION: Map<&'static str, &'static Content> = phf_map! {};
/// Word Related Content
pub static DOCUMENT_TYPE_COLLECTION: Map<&'static str, &'static Content> = phf_map! {};
