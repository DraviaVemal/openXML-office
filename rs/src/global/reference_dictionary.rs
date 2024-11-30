use phf::{phf_map, Map};

pub struct NameSpace {
    pub url: &'static str,
    pub alias: &'static str,
}

pub static NAME_SPACE_COLLECTION: Map<&'static str, &'static NameSpace> = phf_map! {
    "https://draviavemal.com/openxml-office" => &NameSpace{
        url:"https://draviavemal.com/openxml-office",
        alias:"mdoo"
    },
    "http://purl.org/dc/elements/1.1/" => &NameSpace{
        url:"http://purl.org/dc/elements/1.1/",
        alias:"dc"
    },
    "http://purl.org/dc/terms/"  => &NameSpace{
        url:"http://purl.org/dc/terms/",
        alias:"dcterms"
    },
    "http://purl.org/dc/dcmitype/" => &NameSpace{
        url:"http://purl.org/dc/dcmitype/",
        alias:"dcmitype"
    },
    "http://www.w3.org/2001/XMLSchema-instance" => &NameSpace{
        url:"http://www.w3.org/2001/XMLSchema-instance",
        alias:"xsi"
    },
    "http://schemas.openxmlformats.org/drawingml/2006/main" => &NameSpace{
        url:"http://schemas.openxmlformats.org/drawingml/2006/main",
        alias:"a"
    },
    "http://schemas.openxmlformats.org/spreadsheetml/2006/main" => &NameSpace{
        url:"http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        alias:"x"
    },
    "http://schemas.openxmlformats.org/markup-compatibility/2006" => &NameSpace{
        url:"http://schemas.openxmlformats.org/markup-compatibility/2006",
        alias:"mc"
    },
    "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes" => &NameSpace{
        url:"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
        alias:"vt"
    },
};
