//! The package: a ZIP of parts, read with limits and written back losslessly, with undo/redo

use std::collections::HashMap;
use std::io::{Cursor, Read, Write};

use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

/// What a package (a ZIP archive) may expand to when it is opened. The defaults are far above
/// any real document and keep a malicious file from exhausting the browser's memory.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct ZipLimits {
    /// Entries in the archive
    pub max_entries: usize,
    /// Uncompressed size of one part (e.g. one picture or video), in bytes
    pub max_part_bytes: u64,
    /// Uncompressed size of all parts together, in bytes
    pub max_total_bytes: u64,
}

impl Default for ZipLimits {
    fn default() -> Self {
        ZipLimits { max_entries: 10_000, max_part_bytes: 256 * 1024 * 1024, max_total_bytes: 512 * 1024 * 1024 }
    }
}

/// One undoable step: previous contents of every part it changed, plus the editor selection
/// before and after (opaque JSON owned by the frontend)
#[derive(Default)]
struct UndoStep {
    parts: Vec<(String, Option<Vec<u8>>)>,
    selection_before: Option<String>,
    selection_after: Option<String>,
}

#[derive(Default)]
struct History {
    undo: Vec<UndoStep>,
    redo: Vec<UndoStep>,
    /// Step being recorded since the last checkpoint
    recording: Option<UndoStep>,
}

const MAX_UNDO_STEPS: usize = 200;

/// Result of an undo or redo: the selection to restore (`None` when unknown) and the parts
/// that changed, so the caller can refresh what it derived from them
pub struct Undone {
    pub selection: Option<String>,
    pub parts: Vec<String>,
}

/// The parts of an OOXML package by name. Parts are kept as bytes and written back in their
/// original order, so whatever the editor does not touch survives byte for byte.
pub struct Package {
    /// Format name used in messages ("DOCX", "XLSX")
    format: &'static str,
    parts: HashMap<String, Vec<u8>>,
    original_order: Vec<String>,
    original_size: usize,
    history: History,
}

impl Package {
    /// Reads a package from the bytes of its ZIP, refusing archives that expand past `limits`
    pub fn from_bytes(bytes: &[u8], limits: &ZipLimits, format: &'static str) -> Result<Self, String> {
        if bytes.len() < 4 {
            return Err("El archivo proporcionado está vacío o es demasiado pequeño.".to_string());
        }

        let original_size = bytes.len();
        let reader = Cursor::new(bytes);
        let mut archive = ZipArchive::new(reader)
            .map_err(|e| format!("No se pudo leer el archivo como formato {} (ZIP): {}", format, e))?;

        let mut parts = HashMap::new();
        let mut original_order = Vec::new();

        // A tiny malicious ZIP could expand to gigabytes (a "zip bomb"). Sizes declared in
        // the archive are not trusted: reading stops once a limit is passed.
        if archive.len() > limits.max_entries {
            return Err(format!(
                "El archivo tiene demasiadas partes ({}; el máximo es {}).",
                archive.len(),
                limits.max_entries
            ));
        }
        let mut total: u64 = 0;
        for i in 0..archive.len() {
            let mut file = archive
                .by_index(i)
                .map_err(|e| format!("Error al leer entrada ZIP {}: {}", i, e))?;
            let name = file.name().to_string();
            let cap = limits.max_part_bytes.min(limits.max_total_bytes.saturating_sub(total));
            let mut content = Vec::with_capacity(file.size().min(cap).min(16 * 1024 * 1024) as usize);
            (&mut file)
                .take(cap + 1)
                .read_to_end(&mut content)
                .map_err(|e| format!("Error al extraer '{}': {}", name, e))?;
            if content.len() as u64 > cap {
                return Err(if cap < limits.max_part_bytes {
                    format!(
                        "El documento ocupa más de {} MB al descomprimirse; no se abre por seguridad.",
                        limits.max_total_bytes / (1024 * 1024)
                    )
                } else {
                    format!(
                        "La parte '{}' ocupa más de {} MB al descomprimirse; no se abre por seguridad.",
                        name,
                        limits.max_part_bytes / (1024 * 1024)
                    )
                });
            }
            total += content.len() as u64;

            parts.insert(name.clone(), content);
            original_order.push(name);
        }

        Ok(Package { format, parts, original_order, original_size, history: History::default() })
    }

    /// All parts by name
    pub fn parts(&self) -> &HashMap<String, Vec<u8>> {
        &self.parts
    }

    pub fn get(&self, name: &str) -> Option<&Vec<u8>> {
        self.parts.get(name)
    }

    pub fn contains(&self, name: &str) -> bool {
        self.parts.contains_key(name)
    }

    /// A part as text
    pub fn get_string(&self, name: &str) -> Result<String, String> {
        match self.parts.get(name) {
            Some(bytes) => String::from_utf8(bytes.clone())
                .map_err(|e| format!("El archivo '{}' no contiene UTF-8 válido: {}", name, e)),
            None => Err(format!("No se encontró '{}' en el archivo {}.", name, self.format)),
        }
    }

    /// Writes a part, remembering its previous content for undo when a step is open
    pub fn put(&mut self, name: String, bytes: Vec<u8>) {
        if let Some(step) = self.history.recording.as_mut() {
            if !step.parts.iter().any(|(n, _)| *n == name) {
                step.parts.push((name.clone(), self.parts.get(&name).cloned()));
                self.history.redo.clear();
            }
        }
        self.parts.insert(name, bytes);
    }

    /// Deletes a part, remembering its content for undo when a step is open
    pub fn remove(&mut self, name: &str) -> Option<Vec<u8>> {
        if !self.parts.contains_key(name) {
            return None;
        }
        if let Some(step) = self.history.recording.as_mut() {
            if !step.parts.iter().any(|(n, _)| n == name) {
                step.parts.push((name.to_string(), self.parts.get(name).cloned()));
                self.history.redo.clear();
            }
        }
        self.parts.remove(name)
    }

    /// Names of the parts in the order they are written, starting with the original archive's
    pub fn order(&self) -> &[String] {
        &self.original_order
    }

    /// Gives a new part a fixed place after the existing ones (parts not listed are written
    /// last, in no particular order)
    pub fn keep_in_order(&mut self, name: &str) {
        if !self.original_order.iter().any(|n| n == name) {
            self.original_order.push(name.to_string());
        }
    }

    /// Size of the archive the package was read from, in bytes
    pub fn original_size(&self) -> usize {
        self.original_size
    }

    /// Writes the package as a ZIP: original parts in their original order, then new ones
    pub fn to_bytes(&self) -> Result<Vec<u8>, String> {
        let mut buffer = Cursor::new(Vec::new());
        let mut zip = ZipWriter::new(&mut buffer);

        let options = SimpleFileOptions::default()
            .compression_method(CompressionMethod::Deflated);

        for name in &self.original_order {
            if let Some(content) = self.parts.get(name) {
                zip.start_file(name, options)
                    .map_err(|e| format!("Error al escribir entrada ZIP '{}': {}", name, e))?;
                zip.write_all(content)
                    .map_err(|e| format!("Error al escribir contenido de '{}': {}", name, e))?;
            }
        }

        // Include any new parts added
        for (name, content) in &self.parts {
            if !self.original_order.contains(name) {
                zip.start_file(name, options)
                    .map_err(|e| format!("Error al escribir nueva entrada ZIP '{}': {}", name, e))?;
                zip.write_all(content)
                    .map_err(|e| format!("Error al escribir contenido de '{}': {}", name, e))?;
            }
        }

        zip.finish()
            .map_err(|e| format!("Error al generar archivo final {}: {}", self.format, e))?;

        Ok(buffer.into_inner())
    }

    /// Starts a new undo step; changes made until the next checkpoint undo together
    pub fn checkpoint(&mut self, selection_before: Option<String>) {
        self.finish_step();
        self.history.recording = Some(UndoStep { selection_before, ..Default::default() });
    }

    /// Records where the selection ended up after the current step's changes
    pub fn set_selection_after(&mut self, selection: Option<String>) {
        if let Some(step) = self.history.recording.as_mut() {
            step.selection_after = selection;
        }
    }

    /// True while a step is being recorded (edits can be merged into it, e.g. while typing)
    pub fn is_recording(&self) -> bool {
        self.history.recording.is_some()
    }

    fn finish_step(&mut self) {
        if let Some(step) = self.history.recording.take() {
            if !step.parts.is_empty() {
                self.history.undo.push(step);
                if self.history.undo.len() > MAX_UNDO_STEPS {
                    self.history.undo.remove(0);
                }
            }
        }
    }

    /// Swaps the stored part contents with the current ones and returns the inverse step
    fn apply_step(&mut self, step: UndoStep) -> UndoStep {
        let parts = step
            .parts
            .into_iter()
            .map(|(name, content)| {
                let current = match content {
                    Some(bytes) => self.parts.insert(name.clone(), bytes),
                    None => self.parts.remove(&name),
                };
                (name, current)
            })
            .collect::<Vec<_>>();
        UndoStep { parts, selection_before: step.selection_before, selection_after: step.selection_after }
    }

    /// Undoes the last step
    pub fn undo(&mut self) -> Option<Undone> {
        self.finish_step();
        let step = self.history.undo.pop()?;
        let inverse = self.apply_step(step);
        let done = Undone {
            selection: inverse.selection_before.clone(),
            parts: inverse.parts.iter().map(|(name, _)| name.clone()).collect(),
        };
        self.history.redo.push(inverse);
        Some(done)
    }

    /// Redoes the last undone step
    pub fn redo(&mut self) -> Option<Undone> {
        self.finish_step();
        let step = self.history.redo.pop()?;
        let inverse = self.apply_step(step);
        let done = Undone {
            selection: inverse.selection_after.clone(),
            parts: inverse.parts.iter().map(|(name, _)| name.clone()).collect(),
        };
        self.history.undo.push(inverse);
        Some(done)
    }

    pub fn can_undo(&self) -> bool {
        !self.history.undo.is_empty() || self.history.recording.as_ref().is_some_and(|s| !s.parts.is_empty())
    }

    pub fn can_redo(&self) -> bool {
        !self.history.redo.is_empty()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn zip_of(parts: &[(&str, &str)]) -> Vec<u8> {
        let mut buffer = Cursor::new(Vec::new());
        let mut zip = ZipWriter::new(&mut buffer);
        for (name, content) in parts {
            zip.start_file(*name, SimpleFileOptions::default()).unwrap();
            zip.write_all(content.as_bytes()).unwrap();
        }
        zip.finish().unwrap();
        buffer.into_inner()
    }

    fn names(bytes: &[u8]) -> Vec<String> {
        let mut archive = ZipArchive::new(Cursor::new(bytes)).unwrap();
        (0..archive.len()).map(|i| archive.by_index(i).unwrap().name().to_string()).collect()
    }

    #[test]
    fn round_trips_parts_in_original_order() {
        let bytes = zip_of(&[("b.xml", "<b/>"), ("a.xml", "<a/>"), ("c/d.bin", "\u{1}")]);
        let mut pkg = Package::from_bytes(&bytes, &ZipLimits::default(), "TEST").unwrap();
        pkg.put("new.xml".to_string(), b"<n/>".to_vec());
        let out = pkg.to_bytes().unwrap();
        assert_eq!(names(&out), ["b.xml", "a.xml", "c/d.bin", "new.xml"]);
        let again = Package::from_bytes(&out, &ZipLimits::default(), "TEST").unwrap();
        assert_eq!(again.get_string("a.xml").unwrap(), "<a/>");
        assert!(again.get_string("missing").unwrap_err().contains("archivo TEST"));
    }

    #[test]
    fn refuses_archives_past_the_limits() {
        let bytes = zip_of(&[("a", "12345"), ("b", "12345")]);
        let few = ZipLimits { max_entries: 1, ..ZipLimits::default() };
        assert!(Package::from_bytes(&bytes, &few, "TEST").is_err());
        let small = ZipLimits { max_part_bytes: 4, ..ZipLimits::default() };
        assert!(Package::from_bytes(&bytes, &small, "TEST").err().unwrap().contains("'a'"));
        let total = ZipLimits { max_total_bytes: 8, ..ZipLimits::default() };
        assert!(Package::from_bytes(&bytes, &total, "TEST").is_err());
        assert!(Package::from_bytes(b"xx", &ZipLimits::default(), "TEST").is_err());
    }

    #[test]
    fn undoes_and_redoes_steps() {
        let bytes = zip_of(&[("a", "1")]);
        let mut pkg = Package::from_bytes(&bytes, &ZipLimits::default(), "TEST").unwrap();
        pkg.put("a".to_string(), b"untracked".to_vec());
        assert!(!pkg.can_undo());

        pkg.checkpoint(Some("before".to_string()));
        pkg.put("a".to_string(), b"2".to_vec());
        pkg.put("b".to_string(), b"new".to_vec());
        pkg.set_selection_after(Some("after".to_string()));
        assert!(pkg.can_undo());

        let undone = pkg.undo().unwrap();
        assert_eq!(undone.selection.as_deref(), Some("before"));
        assert_eq!(undone.parts, ["a", "b"]);
        assert_eq!(pkg.get("a").unwrap(), b"untracked");
        assert!(!pkg.contains("b"));

        let redone = pkg.redo().unwrap();
        assert_eq!(redone.selection.as_deref(), Some("after"));
        assert_eq!(pkg.get("b").unwrap(), b"new");
        assert!(pkg.redo().is_none());
    }

    #[test]
    fn removes_parts_undoably() {
        let bytes = zip_of(&[("a", "1"), ("b", "2")]);
        let mut pkg = Package::from_bytes(&bytes, &ZipLimits::default(), "TEST").unwrap();
        pkg.checkpoint(None);
        assert_eq!(pkg.remove("a").unwrap(), b"1");
        assert!(pkg.remove("a").is_none());
        assert_eq!(names(&pkg.to_bytes().unwrap()), ["b"]);
        pkg.undo().unwrap();
        assert_eq!(pkg.get("a").unwrap(), b"1");
        assert_eq!(names(&pkg.to_bytes().unwrap()), ["a", "b"]);
    }
}
