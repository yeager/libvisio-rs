//! Integration tests for libvisio-rs.

use libvisio_rs;

#[test]
fn test_is_supported() {
    assert!(libvisio_rs::is_supported("test.vsdx"));
    assert!(libvisio_rs::is_supported("test.vsd"));
    assert!(libvisio_rs::is_supported("test.vstx"));
    assert!(libvisio_rs::is_supported("test.vssx"));
    assert!(!libvisio_rs::is_supported("test.pdf"));
    assert!(!libvisio_rs::is_supported("test.docx"));
}

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
