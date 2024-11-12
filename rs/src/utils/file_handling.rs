use anyhow::{Error as AnyError, Ok, Result as AnyResult};
use core::str;
use flate2::bufread::GzDecoder;
use flate2::write::GzEncoder;
use flate2::Compression;
use std::io::{Read, Write};

pub fn vector_to_string(vector_data: &Vec<u8>) -> AnyResult<String, AnyError> {
    return match str::from_utf8(vector_data) {
        Result::Ok(v) => Ok(v.to_string()),
        Err(e) => Err(e.into()),
    };
}

pub fn compress_content(uncompressed_data: &[u8]) -> AnyResult<Vec<u8>, AnyError> {
    let mut encoder = GzEncoder::new(Vec::new(), Compression::default());
    encoder.write_all(uncompressed_data)?;
    let compressed_data = encoder.finish()?;
    Ok(compressed_data)
}

pub fn decompress_content(compressed_data: &[u8]) -> AnyResult<Vec<u8>, AnyError> {
    let mut decoder = GzDecoder::new(compressed_data);
    let mut decompressed_data = Vec::new();
    decoder.read_to_end(&mut decompressed_data)?;
    Ok(decompressed_data)
}
