use std::env;
use std::path::Path;

use cbindgen::Config;

fn main() {
    let crate_dir = env::var("CARGO_MANIFEST_DIR").unwrap();
    let methods = cbindgen::Builder::new()
        .with_crate(crate_dir.clone())
        .with_config(Config {
            language:cbindgen::Language::C,
            autogen_warning:Some("/* Warning, this file is autogenerated by cbindgen. Don't modify this manually. */".to_string()),
            header:Some("/* Dravia Vemal. MIT License */".to_string()),
            include_version:true,
            package_version:true,
            no_includes:true,
            ..Default::default()
        })
        .generate()
        .unwrap();
    methods.write_to_file(Path::new("../target").join("methods.h"));

    let headers = cbindgen::Builder::new()
        .with_crate(crate_dir)
        .with_config(Config {
            language:cbindgen::Language::C,
            autogen_warning:Some("/* Warning, this file is autogenerated by cbindgen. Don't modify this manually. */".to_string()),
            header:Some("/* Dravia Vemal. MIT License */".to_string()),
            include_version:true,
            package_version:true,
            ..Default::default()
        })
        .generate()
        .unwrap();
    headers.write_to_file(Path::new("../target").join("headers.h"));
}
