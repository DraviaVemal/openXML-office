use std::io::Write;

use anyhow::Result;
use flate2::write::GzEncoder;
use flate2::Compression;

pub fn compress_content(data: &[u8]) -> Result<Vec<u8>> {
    let mut encoder = GzEncoder::new(Vec::new(), Compression::default());
    encoder.write_all(data)?;
    let compressed_data = encoder.finish()?;
    Ok(compressed_data)
}
