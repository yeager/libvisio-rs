//! Integration tests for libvisio-rs.

// =============================================================================
// File format support tests
// =============================================================================

#[test]
fn test_is_supported_vsdx() {
    assert!(libvisio_rs::is_supported("test.vsdx"));
}

#[test]
fn test_is_supported_vsd() {
    assert!(libvisio_rs::is_supported("test.vsd"));
}

#[test]
fn test_is_supported_vstx() {
    assert!(libvisio_rs::is_supported("test.vstx"));
}

#[test]
fn test_is_supported_vssx() {
    assert!(libvisio_rs::is_supported("test.vssx"));
}

#[test]
fn test_is_supported_vsdm() {
    assert!(libvisio_rs::is_supported("test.vsdm"));
}

#[test]
fn test_is_supported_vstm() {
    assert!(libvisio_rs::is_supported("test.vstm"));
}

#[test]
fn test_is_supported_vssm() {
    assert!(libvisio_rs::is_supported("test.vssm"));
}

#[test]
fn test_is_supported_vss() {
    assert!(libvisio_rs::is_supported("test.vss"));
}

#[test]
fn test_is_supported_vst() {
    assert!(libvisio_rs::is_supported("test.vst"));
}

#[test]
fn test_not_supported_pdf() {
    assert!(!libvisio_rs::is_supported("test.pdf"));
}

#[test]
fn test_not_supported_docx() {
    assert!(!libvisio_rs::is_supported("test.docx"));
}

#[test]
fn test_not_supported_xlsx() {
    assert!(!libvisio_rs::is_supported("test.xlsx"));
}

#[test]
fn test_not_supported_svg() {
    assert!(!libvisio_rs::is_supported("test.svg"));
}

#[test]
fn test_not_supported_txt() {
    assert!(!libvisio_rs::is_supported("test.txt"));
}

#[test]
fn test_is_supported_case_insensitive() {
    assert!(libvisio_rs::is_supported("test.VSDX"));
    assert!(libvisio_rs::is_supported("test.VSD"));
    assert!(libvisio_rs::is_supported("test.Vsdx"));
}

#[test]
fn test_is_supported_with_path() {
    assert!(libvisio_rs::is_supported("/home/user/docs/diagram.vsdx"));
    assert!(libvisio_rs::is_supported("C:\\Users\\docs\\test.vsd"));
}

#[test]
fn test_is_supported_empty() {
    assert!(!libvisio_rs::is_supported(""));
    assert!(!libvisio_rs::is_supported("noext"));
}

// =============================================================================
// Error handling tests
// =============================================================================

#[test]
fn test_unsupported_format() {
    let result = libvisio_rs::parse("test.pdf");
    assert!(result.is_err());
}

#[test]
fn test_file_not_found() {
    let result = libvisio_rs::parse("nonexistent.vsdx");
    assert!(result.is_err());
}

#[test]
fn test_file_not_found_vsd() {
    let result = libvisio_rs::parse("nonexistent.vsd");
    assert!(result.is_err());
}

#[test]
fn test_convert_file_not_found() {
    let result = libvisio_rs::convert("nonexistent.vsdx", None, None);
    assert!(result.is_err());
}

#[test]
fn test_get_page_info_file_not_found() {
    let result = libvisio_rs::get_page_info("nonexistent.vsdx");
    assert!(result.is_err());
}

#[test]
fn test_extract_text_file_not_found() {
    let result = libvisio_rs::extract_text("nonexistent.vsdx");
    assert!(result.is_err());
}

#[test]
fn test_convert_page_to_svg_file_not_found() {
    let result = libvisio_rs::convert_page_to_svg("nonexistent.vsdx", 0);
    assert!(result.is_err());
}

// =============================================================================
// Extension constants tests
// =============================================================================

#[test]
fn test_xml_extensions_count() {
    assert_eq!(libvisio_rs::XML_EXTENSIONS.len(), 6);
}

#[test]
fn test_binary_extensions_count() {
    assert_eq!(libvisio_rs::BINARY_EXTENSIONS.len(), 3);
}

#[test]
fn test_all_extensions_count() {
    assert_eq!(
        libvisio_rs::ALL_EXTENSIONS.len(),
        libvisio_rs::XML_EXTENSIONS.len() + libvisio_rs::BINARY_EXTENSIONS.len()
    );
}

#[test]
fn test_xml_extensions_contain_vsdx() {
    assert!(libvisio_rs::XML_EXTENSIONS.contains(&".vsdx"));
}

#[test]
fn test_binary_extensions_contain_vsd() {
    assert!(libvisio_rs::BINARY_EXTENSIONS.contains(&".vsd"));
}

#[test]
fn test_all_extensions_superset() {
    for ext in libvisio_rs::XML_EXTENSIONS {
        assert!(libvisio_rs::ALL_EXTENSIONS.contains(ext));
    }
    for ext in libvisio_rs::BINARY_EXTENSIONS {
        assert!(libvisio_rs::ALL_EXTENSIONS.contains(ext));
    }
}
