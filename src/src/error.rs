//! Error types for libvisio-rs.

use thiserror::Error;

/// Main error type for libvisio-rs operations.
#[derive(Error, Debug)]
pub enum VisioError {
    #[error("I/O error: {0}")]
    Io(#[from] std::io::Error),

    #[error("ZIP error: {0}")]
    Zip(#[from] zip::result::ZipError),

    #[error("XML parse error: {0}")]
    Xml(String),

    #[error("Invalid Visio file: {0}")]
    InvalidFile(String),

    #[error("Unsupported format: {0}")]
    UnsupportedFormat(String),

    #[error("CFB/OLE2 error: {0}")]
    Cfb(String),

    #[error("Decompression error: {0}")]
    Decompression(String),

    #[error("Page not found: {0}")]
    PageNotFound(usize),
}

pub type Result<T> = std::result::Result<T, VisioError>;
