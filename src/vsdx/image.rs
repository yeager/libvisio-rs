//! Embedded image handling for .vsdx files.

use base64::Engine;
use std::collections::HashMap;

/// MIME types for embedded images.
pub fn mime_for_ext(ext: &str) -> &'static str {
    match ext.to_lowercase().as_str() {
        ".png" | "png" => "image/png",
        ".jpg" | ".jpeg" | "jpg" | "jpeg" => "image/jpeg",
        ".gif" | "gif" => "image/gif",
        ".bmp" | "bmp" => "image/bmp",
        ".emf" | "emf" => "image/x-emf",
        ".wmf" | "wmf" => "image/x-wmf",
        ".tiff" | ".tif" | "tiff" | "tif" => "image/tiff",
        ".svg" | "svg" => "image/svg+xml",
        _ => "image/png",
    }
}

/// Convert image bytes to a base64 data URI.
pub fn image_to_data_uri(data: &[u8], filename: &str) -> String {
    let ext = std::path::Path::new(filename)
        .extension()
        .and_then(|e| e.to_str())
        .unwrap_or("png");
    let mime = mime_for_ext(ext);
    let b64 = base64::engine::general_purpose::STANDARD.encode(data);
    format!("data:{};base64,{}", mime, b64)
}

/// Extract all files from visio/media/ in the ZIP.
pub fn extract_media(zip: &mut zip::ZipArchive<std::io::Cursor<&[u8]>>) -> HashMap<String, Vec<u8>> {
    let mut media = HashMap::new();
    let names: Vec<String> = (0..zip.len())
        .filter_map(|i| zip.by_index(i).ok().map(|f| f.name().to_string()))
        .collect();

    for name in names {
        if name.starts_with("visio/media/") {
            if let Some(fname) = name.rsplit('/').next() {
                if !fname.is_empty() {
                    if let Ok(mut f) = zip.by_name(&name) {
                        let mut buf = Vec::new();
                        if std::io::Read::read_to_end(&mut f, &mut buf).is_ok() {
                            media.insert(fname.to_string(), buf);
                        }
                    }
                }
            }
        }
    }
    media
}

/// Parse relationship file for a page to map rId -> target path.
pub fn parse_rels(
    zip: &mut zip::ZipArchive<std::io::Cursor<&[u8]>>,
    page_file: &str,
) -> HashMap<String, String> {
    let mut rels = HashMap::new();
    let dir = std::path::Path::new(page_file)
        .parent()
        .and_then(|p| p.to_str())
        .unwrap_or("");
    let basename = std::path::Path::new(page_file)
        .file_name()
        .and_then(|n| n.to_str())
        .unwrap_or("");
    let rels_path = format!("{}/_rels/{}.rels", dir, basename);

    if let Ok(mut f) = zip.by_name(&rels_path) {
        let mut buf = Vec::new();
        if std::io::Read::read_to_end(&mut f, &mut buf).is_ok() {
            if let Ok(xml_str) = std::str::from_utf8(&buf) {
                if let Ok(doc) = roxmltree::Document::parse(xml_str) {
                    for node in doc.descendants() {
                        if node.tag_name().name() == "Relationship" {
                            let rid = node.attribute("Id").unwrap_or("");
                            let target = node.attribute("Target").unwrap_or("");
                            if !rid.is_empty() && !target.is_empty() {
                                rels.insert(rid.to_string(), target.to_string());
                            }
                        }
                    }
                }
            }
        }
    }
    rels
}
