//! A workbook: the package of an .xlsx read with `ooxml-core`, written back losslessly

use quick_xml::events::Event;
use quick_xml::reader::Reader;

use ooxml_core::xml::{get_attr_value, tag_is};
use ooxml_core::{Package, ZipLimits};

/// Main part of the workbook (sheet list, defined names)
pub const WORKBOOK_PART: &str = "xl/workbook.xml";

pub struct Workbook {
    package: Package,
}

impl Workbook {
    /// Opens an .xlsx from the bytes of its ZIP
    pub fn from_bytes(bytes: &[u8]) -> Result<Self, String> {
        Self::from_bytes_with_limits(bytes, &ZipLimits::default())
    }

    pub fn from_bytes_with_limits(bytes: &[u8], limits: &ZipLimits) -> Result<Self, String> {
        let package = Package::from_bytes(bytes, limits, "XLSX")?;
        if !package.contains(WORKBOOK_PART) {
            return Err(format!("El archivo no es un libro de Excel: falta '{}'.", WORKBOOK_PART));
        }
        Ok(Workbook { package })
    }

    /// Names of the sheets, in the order Excel shows their tabs
    pub fn sheet_names(&self) -> Result<Vec<String>, String> {
        let xml = self.package.get_string(WORKBOOK_PART)?;
        let mut reader = Reader::from_str(&xml);
        let mut names = Vec::new();
        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) if tag_is(e.name().as_ref(), "sheet") => {
                    if let Some(name) = get_attr_value(&e, "name") {
                        names.push(name);
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(format!("'{}' no es XML válido: {}", WORKBOOK_PART, e)),
                _ => {}
            }
        }
        Ok(names)
    }

    /// The workbook as .xlsx bytes: untouched parts are written back as they were read
    pub fn to_bytes(&self) -> Result<Vec<u8>, String> {
        self.package.to_bytes()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::blank::blank_xlsx;

    #[test]
    fn reads_the_sheets_of_a_blank_workbook() {
        let wb = Workbook::from_bytes(&blank_xlsx().unwrap()).unwrap();
        assert_eq!(wb.sheet_names().unwrap(), vec!["Hoja1".to_string()]);
    }

    #[test]
    fn round_trips_every_part_byte_for_byte() {
        let bytes = blank_xlsx().unwrap();
        let first = Workbook::from_bytes(&bytes).unwrap();
        let second = Workbook::from_bytes(&first.to_bytes().unwrap()).unwrap();
        assert_eq!(first.package.order(), second.package.order());
        for name in first.package.order() {
            assert_eq!(first.package.get(name), second.package.get(name), "part {}", name);
        }
    }

    #[test]
    fn refuses_a_zip_that_is_not_a_workbook() {
        let docx_like = crate::blank::zip_parts(&[("word/document.xml", "<w:document/>")]).unwrap();
        let err = Workbook::from_bytes(&docx_like).err().unwrap();
        assert!(err.contains(WORKBOOK_PART), "{}", err);
    }
}
