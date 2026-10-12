pub fn decode_resident_font(bytes: &[u8]) -> Result<Vec<u8>, String> {
    const MAX_FONT_BYTES: usize = 64 * 1024 * 1024;
    if bytes.len() > MAX_FONT_BYTES || bytes.len() < 12 {
        return Err("Invalid resident font size".into());
    }
    let signature = &bytes[..4];
    if signature == b"wOFF" || signature == b"wOF2" {
        let size = bytes
            .get(16..20)
            .and_then(|value| value.try_into().ok())
            .map(u32::from_be_bytes)
            .ok_or("Invalid web font header")?;
        if size as usize > MAX_FONT_BYTES {
            return Err("Invalid resident font size".into());
        }
    }
    let decoded = match signature {
        b"wOFF" => wuff::decompress_woff1(bytes).map_err(|error| error.to_string())?,
        b"wOF2" => wuff::decompress_woff2(bytes).map_err(|error| error.to_string())?,
        _ => bytes.to_vec(),
    };
    if decoded.len() > MAX_FONT_BYTES {
        return Err("Invalid resident font size".into());
    }
    ooxml_text::FontStore::new()
        .register(decoded.clone())
        .map_err(|error| error.to_string())?;
    Ok(decoded)
}
