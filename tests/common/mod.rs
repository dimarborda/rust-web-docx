//! Helpers for tests that use real documents from `examples/`.
//!
//! The folder is git-ignored so everyone can test with their own Word files: put any .docx
//! there to run these tests. When a named document is missing the test is skipped.
#![allow(dead_code)]

use std::fs;
use std::path::Path;

/// Bytes of `examples/<name>`, or `None` (with a notice) when the file is not there
pub fn example(name: &str) -> Option<Vec<u8>> {
    match fs::read(Path::new("examples").join(name)) {
        Ok(bytes) => Some(bytes),
        Err(_) => {
            eprintln!("⏭  examples/{} no existe: test omitido", name);
            None
        }
    }
}

/// Every .docx in `examples/` as (file name, bytes), sorted by name
pub fn all_examples() -> Vec<(String, Vec<u8>)> {
    let Ok(dir) = fs::read_dir("examples") else { return Vec::new() };
    let mut out: Vec<(String, Vec<u8>)> = dir
        .flatten()
        .map(|e| e.path())
        .filter(|p| p.extension().and_then(|s| s.to_str()) == Some("docx"))
        .filter_map(|p| Some((p.file_name()?.to_string_lossy().to_string(), fs::read(&p).ok()?)))
        .collect();
    out.sort();
    out
}
