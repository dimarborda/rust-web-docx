use rust_web_docx::docx_parser::{DocxModifier, ZipLimits};
use std::io::{Cursor, Write};
use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipWriter};

const DOCUMENT: &str = r#"<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Hola</w:t></w:r></w:p></w:body></w:document>"#;

/// A .docx with `extra` parts of `size` zero bytes each (they compress to almost nothing)
fn docx_with(extra: usize, size: usize) -> Vec<u8> {
    let mut buf = Cursor::new(Vec::new());
    {
        let mut zip = ZipWriter::new(&mut buf);
        let opts = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
        zip.start_file("[Content_Types].xml", opts).unwrap();
        zip.write_all(br#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>"#).unwrap();
        zip.start_file("word/document.xml", opts).unwrap();
        zip.write_all(DOCUMENT.as_bytes()).unwrap();
        let zeros = vec![0u8; size];
        for i in 0..extra {
            zip.start_file(format!("word/media/part{}.bin", i), opts).unwrap();
            zip.write_all(&zeros).unwrap();
        }
        zip.finish().unwrap();
    }
    buf.into_inner()
}

const SMALL: ZipLimits = ZipLimits { max_entries: 20, max_part_bytes: 1024 * 1024, max_total_bytes: 3 * 1024 * 1024 };

#[test]
fn test_a_part_that_expands_too_much_is_refused() {
    // 4 MB of zeros compress to a few KB: a small file that expands far beyond its size
    let bomb = docx_with(1, 4 * 1024 * 1024);
    assert!(bomb.len() < 64 * 1024, "the archive itself is tiny: {} bytes", bomb.len());
    let err = DocxModifier::from_bytes_with_limits(&bomb, &SMALL).err().expect("refused");
    assert!(err.contains("part0.bin") && err.contains("1 MB"), "{}", err);
}

#[test]
fn test_parts_that_together_expand_too_much_are_refused() {
    // Each part fits (900 KB), but four of them pass the 3 MB total
    let many = docx_with(4, 900 * 1024);
    let err = DocxModifier::from_bytes_with_limits(&many, &SMALL).err().expect("refused");
    assert!(err.contains("3 MB"), "{}", err);
    // Three of them fit
    assert!(DocxModifier::from_bytes_with_limits(&docx_with(3, 900 * 1024), &SMALL).is_ok());
}

#[test]
fn test_too_many_entries_are_refused() {
    let err = DocxModifier::from_bytes_with_limits(&docx_with(30, 10), &SMALL).err().expect("refused");
    assert!(err.contains("demasiadas partes"), "{}", err);
}

#[test]
fn test_normal_documents_open_with_the_default_limits() {
    let doc = docx_with(2, 2 * 1024 * 1024);
    let m = DocxModifier::from_bytes(&doc).unwrap();
    assert_eq!(m.extract_raw_text().unwrap().trim(), "Hola");
    let limits = ZipLimits::default();
    assert!(limits.max_part_bytes >= 256 * 1024 * 1024 && limits.max_total_bytes >= limits.max_part_bytes);
}

// ---------- XML entities (quick-xml reports them apart from the text) ----------

fn docx_xml(body: &str) -> DocxModifier {
    let document = format!(
        r#"<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{}</w:body></w:document>"#,
        body
    );
    let mut buf = Cursor::new(Vec::new());
    {
        let mut zip = ZipWriter::new(&mut buf);
        let opts = SimpleFileOptions::default();
        zip.start_file("[Content_Types].xml", opts).unwrap();
        zip.write_all(br#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>"#).unwrap();
        zip.start_file("word/document.xml", opts).unwrap();
        zip.write_all(document.as_bytes()).unwrap();
        zip.finish().unwrap();
    }
    DocxModifier::from_bytes(&buf.into_inner()).unwrap()
}

#[test]
fn test_entities_are_read_edited_and_replaced() {
    use rust_web_docx::docx_parser::KeyValuePair;
    let mut m = docx_xml(
        r#"<w:p><w:r><w:t xml:space="preserve">Soporte &amp; mantenimiento &lt;24h&gt; caf&#233; &#x2014; </w:t></w:r><w:r><w:t>{{CLIENTE}}</w:t></w:r></w:p><w:p><w:r><w:t>&amp;</w:t></w:r></w:p>"#,
    );
    let texts = |m: &DocxModifier| m.extract_paragraphs().unwrap().into_iter().map(|p| p.text).collect::<Vec<_>>();
    assert_eq!(texts(&m), vec!["Soporte & mantenimiento <24h> café — {{CLIENTE}}", "&"]);

    // Typing keeps the entities around the edit
    m.replace_paragraph_range(0, 0, 0, "¡").unwrap();
    assert_eq!(texts(&m)[0], "¡Soporte & mantenimiento <24h> café — {{CLIENTE}}");
    // A paragraph that is only an entity is still text
    m.replace_paragraph_range(1, 1, 1, " más").unwrap();
    assert_eq!(texts(&m)[1], "& más");

    // Template variables, with an entity in the value and in the surrounding text
    m.batch_replace(&[KeyValuePair { key: "{{CLIENTE}}".into(), value: "Acme & Co".into() }]).unwrap();
    assert_eq!(texts(&m)[0], "¡Soporte & mantenimiento <24h> café — Acme & Co");

    // The saved XML stays well formed and escaped
    let reopened = DocxModifier::from_bytes(&m.to_bytes().unwrap()).unwrap();
    assert_eq!(texts(&reopened), texts(&m));
    let xml = reopened.get_file_string("word/document.xml").unwrap();
    assert!(xml.contains("Acme &amp; Co") && !xml.contains("Acme & Co"));
}
