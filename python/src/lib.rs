//! Python bindings for libvisio-rs via PyO3.
//! Module name: libvisio_ng (drop-in replacement for the Python version).

use pyo3::prelude::*;
use pyo3::exceptions::PyRuntimeError;
use std::collections::HashMap;

#[pyfunction]
#[pyo3(signature = (path, output_dir=None, page=None))]
fn convert(path: &str, output_dir: Option<&str>, page: Option<usize>) -> PyResult<Vec<String>> {
    libvisio_rs::convert(path, output_dir, page)
        .map_err(|e| PyRuntimeError::new_err(e.to_string()))
}

#[pyfunction]
fn get_page_info(path: &str) -> PyResult<Vec<HashMap<String, PyObject>>> {
    Python::with_gil(|py| {
        let pages = libvisio_rs::get_page_info(path)
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
        Ok(pages.into_iter().map(|p| {
            let mut m = HashMap::new();
            m.insert("name".to_string(), p.name.into_pyobject(py).unwrap().into_any().unbind());
            m.insert("index".to_string(), p.index.into_pyobject(py).unwrap().into_any().unbind());
            m.insert("page_w".to_string(), p.width.into_pyobject(py).unwrap().into_any().unbind());
            m.insert("page_h".to_string(), p.height.into_pyobject(py).unwrap().into_any().unbind());
            m
        }).collect())
    })
}

#[pyfunction]
fn extract_text(path: &str) -> PyResult<String> {
    libvisio_rs::extract_text(path)
        .map_err(|e| PyRuntimeError::new_err(e.to_string()))
}

#[pymodule]
fn libvisio_ng(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_function(wrap_pyfunction!(convert, m)?)?;
    m.add_function(wrap_pyfunction!(get_page_info, m)?)?;
    m.add_function(wrap_pyfunction!(extract_text, m)?)?;
    Ok(())
}
