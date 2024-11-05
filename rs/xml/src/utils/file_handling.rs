use anyhow::{Error, Result};
use flate2::bufread::GzDecoder;
use flate2::write::GzEncoder;
use flate2::Compression;
use std::io::{Read, Write};

pub fn compress_content(uncompressed_data: &[u8]) -> Result<Vec<u8>, Error> {
    let mut encoder = GzEncoder::new(Vec::new(), Compression::default());
    encoder.write_all(uncompressed_data)?;
    let compressed_data = encoder.finish()?;
    Ok(compressed_data)
}

pub fn decompress_content(compressed_data: &[u8]) -> Result<Vec<u8>, Error> {
    let mut decoder = GzDecoder::new(compressed_data);
    let mut decompressed_data = Vec::new();
    decoder.read_to_end(&mut decompressed_data)?;
    Ok(decompressed_data)
}
