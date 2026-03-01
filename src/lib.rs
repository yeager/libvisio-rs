//! libvisio-rs — A Rust library for parsing Microsoft Visio files and converting to SVG.
//!
//! Supports both .vsdx (ZIP+XML, Open Packaging) and .vsd (OLE2 binary) formats.
//! Produces high-quality SVG output with themes, gradients, shadows, rich text,
//! embedded images, connectors, and more.
//!
//! # Examples
//!
//! ```no_run
//! use libvisio_rs::{convert, get_page_info, extract_text};
//!
//! // Convert all pages to SVG
//! let svg_files = convert("diagram.vsdx", Some("/tmp/output"), None).unwrap();
//!
//! // Get page info
//! let pages = get_page_info("diagram.vsdx").unwrap();
//!
//! // Extract text
//! let text = extract_text("diagram.vsdx").unwrap();
//! ```

pub mod error;
pub mod model;
pub mod svg;
pub mod vsd;
pub mod vsdx;

use crate::error::{Result, VisioError};
use crate::model::*;
use crate::svg::render;
use std::path::Path;

/// Supported file extensions for XML-based formats.
pub const XML_EXTENSIONS: &[&str] = &[".vsdx", ".vstx", ".vssx", ".vsdm", ".vstm", ".vssm"];

/// Supported file extensions for binary formats.
pub const BINARY_EXTENSIONS: &[&str] = &[".vsd", ".vss", ".vst"];

/// All supported file extensions.
pub const ALL_EXTENSIONS: &[&str] = &[
    ".vsdx", ".vstx", ".vssx", ".vsdm", ".vstm", ".vssm", ".vsd", ".vss", ".vst",
];

/// Check if a file extension is supported.
pub fn is_supported(path: &str) -> bool {
    let ext = Path::new(path)
        .extension()
        .and_then(|e| e.to_str())
        .unwrap_or("");
    let dotted = format!(".{}", ext.to_lowercase());
    ALL_EXTENSIONS.contains(&dotted.as_str())
}

/// Parse a Visio file and return the parsed document.
pub fn parse(path: &str) -> Result<Document> {
    let data = std::fs::read(path)?;
    let ext = Path::new(path)
        .extension()
        .and_then(|e| e.to_str())
        .map(|e| e.to_lowercase())
        .unwrap_or_default();
    let dotted = format!(".{}", ext);

    if XML_EXTENSIONS.contains(&dotted.as_str()) {
        vsdx::parser::parse_vsdx(&data)
    } else if BINARY_EXTENSIONS.contains(&dotted.as_str()) {
        vsd::parser::parse_vsd(&data)
    } else {
        Err(VisioError::UnsupportedFormat(format!(
            "Unsupported format: .{}",
            ext
        )))
    }
}

/// Convert a Visio file to SVG pages.
///
/// Returns a list of SVG file paths (one per page).
pub fn convert(
    input_path: &str,
    output_dir: Option<&str>,
    page: Option<usize>,
) -> Result<Vec<String>> {
    let out_dir = output_dir.unwrap_or("/tmp/libvisio_rs_output");
    std::fs::create_dir_all(out_dir)?;

    let doc = parse(input_path)?;
    let basename = Path::new(input_path)
        .file_stem()
        .and_then(|s| s.to_str())
        .unwrap_or("visio");

    let mut svg_files = Vec::new();

    for p in &doc.pages {
        if let Some(page_num) = page {
            if p.index != page_num {
                continue;
            }
        }

        // Get background shapes if any
        let bg_shapes: Option<Vec<Shape>> = doc
            .background_map
            .get(&p.index)
            .and_then(|bg_idx| doc.pages.iter().find(|pp| pp.index == *bg_idx))
            .map(|bg_page| bg_page.shapes.clone());

        let svg_content = render::shapes_to_svg(
            &p.shapes,
            p.width,
            p.height,
            &doc.masters,
            &p.connects,
            &doc.media,
            &std::collections::HashMap::new(),
            bg_shapes.as_deref(),
            &doc.theme_colors,
            &p.layers,
        );

        let svg_path = format!("{}/{}_page{}.svg", out_dir, basename, p.index + 1);
        std::fs::write(&svg_path, &svg_content)?;
        svg_files.push(svg_path);
    }

    Ok(svg_files)
}

/// Convert a single page to SVG string (no file output).
pub fn convert_page_to_svg(input_path: &str, page_index: usize) -> Result<String> {
    let doc = parse(input_path)?;
    let page = doc
        .pages
        .iter()
        .find(|p| p.index == page_index)
        .ok_or(VisioError::PageNotFound(page_index))?;

    let bg_shapes: Option<Vec<Shape>> = doc
        .background_map
        .get(&page.index)
        .and_then(|bg_idx| doc.pages.iter().find(|pp| pp.index == *bg_idx))
        .map(|bg_page| bg_page.shapes.clone());

    Ok(render::shapes_to_svg(
        &page.shapes,
        page.width,
        page.height,
        &doc.masters,
        &page.connects,
        &doc.media,
        &std::collections::HashMap::new(),
        bg_shapes.as_deref(),
        &doc.theme_colors,
        &page.layers,
    ))
}

/// Get page information from a Visio file.
pub fn get_page_info(path: &str) -> Result<Vec<PageInfo>> {
    let doc = parse(path)?;
    Ok(doc
        .pages
        .iter()
        .map(|p| PageInfo {
            name: p.name.clone(),
            index: p.index,
            width: p.width,
            height: p.height,
        })
        .collect())
}

/// Extract all text from a Visio file.
pub fn extract_text(path: &str) -> Result<String> {
    let doc = parse(path)?;
    let mut text_lines = Vec::new();
    for page in &doc.pages {
        text_lines.push(format!(
            "--- {} ---",
            if page.name.is_empty() {
                format!("Page {}", page.index + 1)
            } else {
                page.name.clone()
            }
        ));
        for shape in &page.shapes {
            extract_shape_text(&mut text_lines, shape);
        }
        text_lines.push(String::new());
    }
    Ok(text_lines.join("\n"))
}

fn extract_shape_text(lines: &mut Vec<String>, shape: &Shape) {
    if !shape.text.is_empty() {
        lines.push(shape.text.clone());
    }
    for sub in &shape.sub_shapes {
        extract_shape_text(lines, sub);
    }
}

// ============================================================
// C ABI — extern "C" functions for FFI
// ============================================================

use std::ffi::{CStr, CString};
use std::os::raw::c_char;

/// Opaque handle to a parsed Visio document.
#[repr(C)]
pub struct VisioDocument {
    inner: Document,
}

/// Open and parse a Visio file. Returns null on error.
#[no_mangle]
pub extern "C" fn visio_open(path: *const c_char) -> *mut VisioDocument {
    if path.is_null() {
        return std::ptr::null_mut();
    }
    let path_str = match unsafe { CStr::from_ptr(path) }.to_str() {
        Ok(s) => s,
        Err(_) => return std::ptr::null_mut(),
    };
    match parse(path_str) {
        Ok(doc) => Box::into_raw(Box::new(VisioDocument { inner: doc })),
        Err(_) => std::ptr::null_mut(),
    }
}

/// Get the number of pages.
#[no_mangle]
pub extern "C" fn visio_get_page_count(doc: *const VisioDocument) -> usize {
    if doc.is_null() {
        return 0;
    }
    unsafe { &*doc }.inner.pages.len()
}

/// Convert a page to SVG. Returns null on error.
/// Caller must free with visio_free_string.
#[no_mangle]
pub extern "C" fn visio_convert_page_to_svg(doc: *const VisioDocument, page: usize) -> *mut c_char {
    if doc.is_null() {
        return std::ptr::null_mut();
    }
    let document = &unsafe { &*doc }.inner;
    let p = match document.pages.iter().find(|p| p.index == page) {
        Some(p) => p,
        None => return std::ptr::null_mut(),
    };

    let bg_shapes: Option<Vec<Shape>> = document
        .background_map
        .get(&p.index)
        .and_then(|bg_idx| document.pages.iter().find(|pp| pp.index == *bg_idx))
        .map(|bg_page| bg_page.shapes.clone());

    let svg = render::shapes_to_svg(
        &p.shapes,
        p.width,
        p.height,
        &document.masters,
        &p.connects,
        &document.media,
        &std::collections::HashMap::new(),
        bg_shapes.as_deref(),
        &document.theme_colors,
        &p.layers,
    );

    match CString::new(svg) {
        Ok(c) => c.into_raw(),
        Err(_) => std::ptr::null_mut(),
    }
}

/// Extract all text from the document.
/// Caller must free with visio_free_string.
#[no_mangle]
pub extern "C" fn visio_extract_text(doc: *const VisioDocument) -> *mut c_char {
    if doc.is_null() {
        return std::ptr::null_mut();
    }
    let document = &unsafe { &*doc }.inner;
    let mut lines = Vec::new();
    for page in &document.pages {
        lines.push(format!(
            "--- {} ---",
            if page.name.is_empty() {
                format!("Page {}", page.index + 1)
            } else {
                page.name.clone()
            }
        ));
        for shape in &page.shapes {
            if !shape.text.is_empty() {
                lines.push(shape.text.clone());
            }
        }
    }
    match CString::new(lines.join("\n")) {
        Ok(c) => c.into_raw(),
        Err(_) => std::ptr::null_mut(),
    }
}

/// Free a VisioDocument handle.
#[no_mangle]
pub extern "C" fn visio_free(doc: *mut VisioDocument) {
    if !doc.is_null() {
        unsafe {
            let _ = Box::from_raw(doc);
        }
    }
}

/// Free a string returned by visio_convert_page_to_svg or visio_extract_text.
#[no_mangle]
pub extern "C" fn visio_free_string(s: *mut c_char) {
    if !s.is_null() {
        unsafe {
            let _ = CString::from_raw(s);
        }
    }
}
