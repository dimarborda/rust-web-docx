use regex::Regex;
use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use std::io::{Cursor, Read, Write};
use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

#[derive(Serialize, Deserialize, Debug, Clone)]
pub struct ParagraphInfo {
    pub index: usize,
    pub text: String,
    pub style: String,
    pub is_heading: bool,
    pub run_count: usize,
}

#[derive(Serialize, Deserialize, Debug, Clone)]
pub struct ParagraphUpdate {
    pub index: usize,
    pub text: String,
}

#[derive(Serialize, Deserialize, Debug, Clone)]
pub struct KeyValuePair {
    pub key: String,
    pub value: String,
}

#[derive(Serialize, Deserialize, Debug, Clone)]
pub struct DocxStats {
    pub paragraph_count: usize,
    pub word_count: usize,
    pub char_count: usize,
    pub files_in_zip: Vec<String>,
    pub original_size_bytes: usize,
}

#[derive(Serialize, Deserialize, Debug, Clone)]
pub struct ReplaceResult {
    pub occurrences_replaced: usize,
    pub affected_files: Vec<String>,
    pub message: String,
}

pub struct DocxModifier {
    files: HashMap<String, Vec<u8>>,
    original_order: Vec<String>,
    original_size: usize,
}

impl DocxModifier {
    /// Loads a docx from byte buffer
    pub fn from_bytes(bytes: &[u8]) -> Result<Self, String> {
        if bytes.len() < 4 {
            return Err("El archivo proporcionado está vacío o es demasiado pequeño.".to_string());
        }

        let original_size = bytes.len();
        let reader = Cursor::new(bytes);
        let mut archive = ZipArchive::new(reader)
            .map_err(|e| format!("No se pudo leer el archivo como formato DOCX (ZIP): {}", e))?;

        let mut files = HashMap::new();
        let mut original_order = Vec::new();

        for i in 0..archive.len() {
            let mut file = archive
                .by_index(i)
                .map_err(|e| format!("Error al leer entrada ZIP {}: {}", i, e))?;
            let name = file.name().to_string();
            let mut content = Vec::new();
            file.read_to_end(&mut content)
                .map_err(|e| format!("Error al extraer '{}': {}", name, e))?;

            files.insert(name.clone(), content);
            original_order.push(name);
        }

        // Verify that this is indeed a docx
        if !files.contains_key("word/document.xml") {
            return Err("El archivo no es un documento DOCX válido (falta 'word/document.xml').".to_string());
        }

        Ok(DocxModifier {
            files,
            original_order,
            original_size,
        })
    }

    /// Extracts list of paragraphs from word/document.xml
    pub fn extract_paragraphs(&self) -> Result<Vec<ParagraphInfo>, String> {
        let doc_xml = self.get_file_string("word/document.xml")?;
        let paragraphs = parse_paragraphs_from_xml(&doc_xml);
        Ok(paragraphs)
    }

    /// Extracts full consolidated plain text from the document
    pub fn extract_raw_text(&self) -> Result<String, String> {
        let paragraphs = self.extract_paragraphs()?;
        let full_text = paragraphs
            .into_iter()
            .map(|p| p.text)
            .filter(|t| !t.trim().is_empty())
            .collect::<Vec<_>>()
            .join("\n\n");
        Ok(full_text)
    }

    /// Finds and replaces text across word/document.xml, headers, and footers
    pub fn find_and_replace(
        &mut self,
        search: &str,
        replacement: &str,
        match_case: bool,
        use_regex: bool,
    ) -> Result<ReplaceResult, String> {
        if search.is_empty() {
            return Err("El término de búsqueda no puede estar vacío.".to_string());
        }

        let target_files: Vec<String> = self
            .files
            .keys()
            .filter(|name| {
                name.starts_with("word/")
                    && (name.ends_with(".xml") || name.ends_with(".rels"))
            })
            .cloned()
            .collect();

        let mut total_replacements = 0;
        let mut affected_files = Vec::new();

        for filename in target_files {
            if let Some(content_bytes) = self.files.get(&filename) {
                if let Ok(xml_str) = String::from_utf8(content_bytes.clone()) {
                    let (new_xml, count) = replace_in_docx_xml(
                        &xml_str,
                        search,
                        replacement,
                        match_case,
                        use_regex,
                    )?;

                    if count > 0 {
                        total_replacements += count;
                        affected_files.push(filename.clone());
                        self.files.insert(filename, new_xml.into_bytes());
                    }
                }
            }
        }

        Ok(ReplaceResult {
            occurrences_replaced: total_replacements,
            affected_files,
            message: format!(
                "Se reemplazaron exitosamente {} coincidencias.",
                total_replacements
            ),
        })
    }

    /// Replaces multiple key-value pairs (template variables)
    pub fn batch_replace(&mut self, pairs: &[KeyValuePair]) -> Result<ReplaceResult, String> {
        let mut total_replacements = 0;
        let mut affected = Vec::new();

        for pair in pairs {
            if !pair.key.trim().is_empty() {
                let res = self.find_and_replace(&pair.key, &pair.value, true, false)?;
                total_replacements += res.occurrences_replaced;
                for f in res.affected_files {
                    if !affected.contains(&f) {
                        affected.push(f);
                    }
                }
            }
        }

        Ok(ReplaceResult {
            occurrences_replaced: total_replacements,
            affected_files: affected,
            message: format!(
                "Reemplazo por lotes completado: {} variables reemplazadas.",
                total_replacements
            ),
        })
    }

    /// Updates individual paragraphs in word/document.xml by index
    pub fn update_paragraphs(&mut self, updates: &[ParagraphUpdate]) -> Result<usize, String> {
        let doc_xml = self.get_file_string("word/document.xml")?;
        let (new_xml, updated_count) = update_paragraphs_in_xml(&doc_xml, updates)?;
        self.files.insert("word/document.xml".to_string(), new_xml.into_bytes());
        Ok(updated_count)
    }

    /// Retrieves document statistics
    pub fn get_statistics(&self) -> Result<DocxStats, String> {
        let paragraphs = self.extract_paragraphs()?;
        let paragraph_count = paragraphs.len();

        let mut word_count = 0;
        let mut char_count = 0;

        for p in &paragraphs {
            char_count += p.text.chars().count();
            word_count += p.text.split_whitespace().count();
        }

        let files_in_zip = self.original_order.clone();

        Ok(DocxStats {
            paragraph_count,
            word_count,
            char_count,
            files_in_zip,
            original_size_bytes: self.original_size,
        })
    }

    /// Exports modified DOCX as bytes
    pub fn to_bytes(&self) -> Result<Vec<u8>, String> {
        let mut buffer = Cursor::new(Vec::new());
        let mut zip = ZipWriter::new(&mut buffer);

        let options = SimpleFileOptions::default()
            .compression_method(CompressionMethod::Deflated);

        for name in &self.original_order {
            if let Some(content) = self.files.get(name) {
                zip.start_file(name, options)
                    .map_err(|e| format!("Error al escribir entrada ZIP '{}': {}", name, e))?;
                zip.write_all(content)
                    .map_err(|e| format!("Error al escribir contenido de '{}': {}", name, e))?;
            }
        }

        // Include any new files added
        for (name, content) in &self.files {
            if !self.original_order.contains(name) {
                zip.start_file(name, options)
                    .map_err(|e| format!("Error al escribir nueva entrada ZIP '{}': {}", name, e))?;
                zip.write_all(content)
                    .map_err(|e| format!("Error al escribir contenido de '{}': {}", name, e))?;
            }
        }

        zip.finish()
            .map_err(|e| format!("Error al generar archivo final DOCX: {}", e))?;

        Ok(buffer.into_inner())
    }

    fn get_file_string(&self, path: &str) -> Result<String, String> {
        match self.files.get(path) {
            Some(bytes) => String::from_utf8(bytes.clone())
                .map_err(|e| format!("El archivo '{}' no contiene UTF-8 válido: {}", path, e)),
            None => Err(format!("No se encontró '{}' en el archivo DOCX.", path)),
        }
    }
}

// ---------------- Helper XML Functions ----------------

/// Parses `<w:p>...</w:p>` blocks from document XML
fn parse_paragraphs_from_xml(xml: &str) -> Vec<ParagraphInfo> {
    let mut paragraphs = Vec::new();
    let mut search_idx = 0;
    let mut p_index = 0;

    while let Some(start_p) = xml[search_idx..].find("<w:p") {
        let p_offset = search_idx + start_p;
        // Find closing tag </w:p>
        if let Some(end_p_rel) = xml[p_offset..].find("</w:p>") {
            let p_end = p_offset + end_p_rel + 6;
            let p_content = &xml[p_offset..p_end];

            // Extract style
            let style = extract_paragraph_style(p_content);
            let is_heading = style.to_lowercase().contains("heading")
                || style.to_lowercase().contains("title")
                || style.to_lowercase().contains("encabezado")
                || style.to_lowercase().contains("título");

            // Extract consolidated text and count runs
            let (text, run_count) = extract_text_and_runs_from_p(p_content);

            paragraphs.push(ParagraphInfo {
                index: p_index,
                text,
                style,
                is_heading,
                run_count,
            });

            p_index += 1;
            search_idx = p_end;
        } else {
            break;
        }
    }

    paragraphs
}

fn extract_paragraph_style(p_content: &str) -> String {
    if let Some(style_tag_pos) = p_content.find("<w:pStyle ") {
        let tag_slice = &p_content[style_tag_pos..];
        if let Some(val_pos) = tag_slice.find("w:val=\"") {
            let val_start = val_pos + 7;
            if let Some(val_end) = tag_slice[val_start..].find('"') {
                return tag_slice[val_start..val_start + val_end].to_string();
            }
        }
    }
    "Normal".to_string()
}

fn extract_text_and_runs_from_p(p_content: &str) -> (String, usize) {
    let mut combined_text = String::new();
    let mut run_count = 0;
    let mut search_idx = 0;

    // We search for <w:t ...>text</w:t> or <w:t>text</w:t>
    while let Some(t_open_rel) = p_content[search_idx..].find("<w:t") {
        let t_open_abs = search_idx + t_open_rel;
        run_count += 1;

        // Find '>' of opening tag
        if let Some(close_bracket) = p_content[t_open_abs..].find('>') {
            let content_start = t_open_abs + close_bracket + 1;
            if let Some(t_close_rel) = p_content[content_start..].find("</w:t>") {
                let content_end = content_start + t_close_rel;
                let raw_text = &p_content[content_start..content_end];
                combined_text.push_str(&unescape_xml(raw_text));
                search_idx = content_end + 6;
            } else {
                search_idx = content_start;
            }
        } else {
            search_idx = t_open_abs + 4;
        }
    }

    (combined_text, run_count)
}

/// Replaces text inside XML `<w:t>` elements and handles split-run matches in paragraphs
fn replace_in_docx_xml(
    xml: &str,
    search: &str,
    replacement: &str,
    match_case: bool,
    use_regex: bool,
) -> Result<(String, usize), String> {
    let mut total_replacements = 0;

    // First, try simple direct text-node replacement
    let mut result_xml = String::with_capacity(xml.len());
    let mut last_idx = 0;

    let regex_matcher = if use_regex {
        let pattern = if match_case {
            search.to_string()
        } else {
            format!("(?i){}", search)
        };
        Some(Regex::new(&pattern).map_err(|e| format!("Expresión regular inválida: {}", e))?)
    } else {
        None
    };

    while let Some(t_open_rel) = xml[last_idx..].find("<w:t") {
        let t_open_abs = last_idx + t_open_rel;
        result_xml.push_str(&xml[last_idx..t_open_abs]);

        if let Some(close_bracket) = xml[t_open_abs..].find('>') {
            let tag_header_end = t_open_abs + close_bracket + 1;
            let tag_header = &xml[t_open_abs..tag_header_end];
            result_xml.push_str(tag_header);

            if let Some(t_close_rel) = xml[tag_header_end..].find("</w:t>") {
                let content_end = tag_header_end + t_close_rel;
                let raw_node_text = &xml[tag_header_end..content_end];
                let unescaped = unescape_xml(raw_node_text);

                let (modified_unescaped, count) = if let Some(ref re) = regex_matcher {
                    let mut c = 0;
                    for _ in re.find_iter(&unescaped) {
                        c += 1;
                    }
                    let replaced = re.replace_all(&unescaped, replacement).to_string();
                    (replaced, c)
                } else if match_case {
                    let c = unescaped.matches(search).count();
                    let replaced = unescaped.replace(search, replacement);
                    (replaced, c)
                } else {
                    let (replaced, c) = replace_case_insensitive(&unescaped, search, replacement);
                    (replaced, c)
                };

                total_replacements += count;
                result_xml.push_str(&escape_xml(&modified_unescaped));
                last_idx = content_end;
            } else {
                last_idx = tag_header_end;
            }
        } else {
            result_xml.push_str("<w:t");
            last_idx = t_open_abs + 4;
        }
    }

    result_xml.push_str(&xml[last_idx..]);

    // If no replacement was found via single-node, check for cross-run split text in paragraphs!
    if total_replacements == 0 && !use_regex {
        let (cross_xml, cross_count) = replace_cross_run_text(&result_xml, search, replacement, match_case);
        if cross_count > 0 {
            return Ok((cross_xml, cross_count));
        }
    }

    Ok((result_xml, total_replacements))
}

/// Case insensitive string replacement
fn replace_case_insensitive(text: &str, search: &str, replacement: &str) -> (String, usize) {
    if search.is_empty() {
        return (text.to_string(), 0);
    }
    let lower_text = text.to_lowercase();
    let lower_search = search.to_lowercase();

    let mut count = 0;
    let mut result = String::new();
    let mut last_idx = 0;

    while let Some(found_idx) = lower_text[last_idx..].find(&lower_search) {
        let abs_idx = last_idx + found_idx;
        result.push_str(&text[last_idx..abs_idx]);
        result.push_str(replacement);
        count += 1;
        last_idx = abs_idx + search.len();
    }
    result.push_str(&text[last_idx..]);

    (result, count)
}

/// Replaces text when word split the phrase across adjacent `<w:r><w:t>` runs in `<w:p>`
fn replace_cross_run_text(
    xml: &str,
    search: &str,
    replacement: &str,
    match_case: bool,
) -> (String, usize) {
    let mut result = String::with_capacity(xml.len());
    let mut search_idx = 0;
    let mut count = 0;

    while let Some(p_start_rel) = xml[search_idx..].find("<w:p") {
        let p_start_abs = search_idx + p_start_rel;
        result.push_str(&xml[search_idx..p_start_abs]);

        if let Some(p_end_rel) = xml[p_start_abs..].find("</w:p>") {
            let p_end_abs = p_start_abs + p_end_rel + 6;
            let p_content = &xml[p_start_abs..p_end_abs];

            let (full_text, _) = extract_text_and_runs_from_p(p_content);
            let contains = if match_case {
                full_text.contains(search)
            } else {
                full_text.to_lowercase().contains(&search.to_lowercase())
            };

            if contains {
                // Modify this paragraph's runs
                let (new_full_text, c) = if match_case {
                    let cnt = full_text.matches(search).count();
                    (full_text.replace(search, replacement), cnt)
                } else {
                    replace_case_insensitive(&full_text, search, replacement)
                };

                let updated_p = set_paragraph_consolidated_text(p_content, &new_full_text);
                result.push_str(&updated_p);
                count += c;
            } else {
                result.push_str(p_content);
            }

            search_idx = p_end_abs;
        } else {
            result.push_str(&xml[p_start_abs..]);
            search_idx = xml.len();
            break;
        }
    }

    result.push_str(&xml[search_idx..]);
    (result, count)
}

/// Sets the consolidated text of a paragraph by writing to the first <w:t> and clearing remaining <w:t>
fn set_paragraph_consolidated_text(p_xml: &str, new_text: &str) -> String {
    let mut result = String::with_capacity(p_xml.len() + new_text.len());
    let mut search_idx = 0;
    let mut first_t_written = false;

    while let Some(t_open_rel) = p_xml[search_idx..].find("<w:t") {
        let t_open_abs = search_idx + t_open_rel;
        result.push_str(&p_xml[search_idx..t_open_abs]);

        if let Some(close_bracket) = p_xml[t_open_abs..].find('>') {
            let tag_header_end = t_open_abs + close_bracket + 1;
            // Ensure xml:space="preserve" is in header
            result.push_str("<w:t xml:space=\"preserve\">");

            if let Some(t_close_rel) = p_xml[tag_header_end..].find("</w:t>") {
                let content_end = tag_header_end + t_close_rel;
                if !first_t_written {
                    result.push_str(&escape_xml(new_text));
                    first_t_written = true;
                }
                result.push_str("</w:t>");
                search_idx = content_end + 6;
            } else {
                search_idx = tag_header_end;
            }
        } else {
            result.push_str("<w:t");
            search_idx = t_open_abs + 4;
        }
    }

    result.push_str(&p_xml[search_idx..]);

    // If paragraph had no <w:t> tags at all, insert a run inside the paragraph
    if !first_t_written {
        if let Some(closing_pos) = result.rfind("</w:p>") {
            let (before, after) = result.split_at(closing_pos);
            return format!(
                "{}<w:r><w:t xml:space=\"preserve\">{}</w:t></w:r>{}",
                before,
                escape_xml(new_text),
                after
            );
        }
    }

    result
}

/// Updates specific paragraphs in word/document.xml by index
fn update_paragraphs_in_xml(
    xml: &str,
    updates: &[ParagraphUpdate],
) -> Result<(String, usize), String> {
    let update_map: HashMap<usize, &str> = updates
        .iter()
        .map(|u| (u.index, u.text.as_str()))
        .collect();

    let mut result = String::with_capacity(xml.len());
    let mut search_idx = 0;
    let mut p_index = 0;
    let mut updated_count = 0;

    while let Some(p_start_rel) = xml[search_idx..].find("<w:p") {
        let p_start_abs = search_idx + p_start_rel;
        result.push_str(&xml[search_idx..p_start_abs]);

        if let Some(p_end_rel) = xml[p_start_abs..].find("</w:p>") {
            let p_end_abs = p_start_abs + p_end_rel + 6;
            let p_content = &xml[p_start_abs..p_end_abs];

            if let Some(new_text) = update_map.get(&p_index) {
                let updated_p = set_paragraph_consolidated_text(p_content, new_text);
                result.push_str(&updated_p);
                updated_count += 1;
            } else {
                result.push_str(p_content);
            }

            p_index += 1;
            search_idx = p_end_abs;
        } else {
            result.push_str(&xml[p_start_abs..]);
            search_idx = xml.len();
            break;
        }
    }

    result.push_str(&xml[search_idx..]);
    Ok((result, updated_count))
}

fn escape_xml(input: &str) -> String {
    input
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

fn unescape_xml(input: &str) -> String {
    input
        .replace("&apos;", "'")
        .replace("&quot;", "\"")
        .replace("&gt;", ">")
        .replace("&lt;", "<")
        .replace("&amp;", "&")
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::sample_generator::generate_sample_docx;

    #[test]
    fn test_sample_docx_parsing_and_modification() {
        let sample_bytes = generate_sample_docx().expect("Should generate sample docx");
        let mut modifier = DocxModifier::from_bytes(&sample_bytes).expect("Should parse docx");

        let paragraphs = modifier.extract_paragraphs().expect("Should extract paragraphs");
        assert!(paragraphs.len() >= 5);
        assert!(paragraphs[0].text.contains("Acuerdo"));

        // Test Find and Replace
        let rep_res = modifier
            .find_and_replace("{{NOMBRE_CLIENTE}}", "Empresa Ejemplo S.A.S.", true, false)
            .expect("Should replace");
        assert_eq!(rep_res.occurrences_replaced, 1);

        // Test Batch Replace
        let pairs = vec![
            KeyValuePair {
                key: "{{FECHA_CONTRATO}}".to_string(),
                value: "30 de Septiembre de 2026".to_string(),
            },
            KeyValuePair {
                key: "{{VALOR_PROYECTO}}".to_string(),
                value: "$25,000,000 COP".to_string(),
            },
        ];
        let batch_res = modifier.batch_replace(&pairs).expect("Should batch replace");
        assert_eq!(batch_res.occurrences_replaced, 2);

        // Export bytes and re-parse
        let exported_bytes = modifier.to_bytes().expect("Should export bytes");
        let reloaded = DocxModifier::from_bytes(&exported_bytes).expect("Should reload docx");
        let raw_text = reloaded.extract_raw_text().expect("Should get raw text");
        assert!(raw_text.contains("Empresa Ejemplo S.A.S."));
        assert!(raw_text.contains("30 de Septiembre de 2026"));
        assert!(raw_text.contains("$25,000,000 COP"));
    }
}
