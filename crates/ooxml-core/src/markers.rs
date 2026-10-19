//! The marker lexicon shared by every format: how a template or a draft says, in its own
//! text, what is still missing and what an automation may fill.
//!
//! - `{{KEY}}` field (canonical). Spaces around the key are allowed (`{{ KEY }}`) and a hint
//!   may follow a colon: `{{FECHA: dd/mm/aaaa}}`.
//! - `{key}` short field: no spaces inside (`{cliente}`, `{fecha:larga}`).
//! - `[ … ]` pending: something a person or an agent still has to write. The text inside is
//!   the hint (`[Completar: alcance del proyecto]`); dots alone (`[...]`, `[…]`) or blanks
//!   (`[____]`) mean "to complete" with no hint. Citations (`[1]`, `[2, 3]`, `[4-6]`),
//!   checkboxes (`[ ]`, `[x]`) and `[sic]` are ordinary text, not markers.
//!
//! A marker never spans a line break. Offsets are in Unicode code points, the unit the
//! editors use for caret positions.

/// What a marker asks for
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MarkerKind {
    /// `{{KEY}}` or `{key}`: a value an automation fills
    Field,
    /// `[ … ]`: content still to be written
    Pending,
}

impl MarkerKind {
    pub fn as_str(self) -> &'static str {
        match self {
            MarkerKind::Field => "field",
            MarkerKind::Pending => "pending",
        }
    }
}

/// How the marker was written
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MarkerSyntax {
    /// `{{ … }}`
    Double,
    /// `{ … }`
    Single,
    /// `[ … ]`
    Bracket,
}

impl MarkerSyntax {
    pub fn as_str(self) -> &'static str {
        match self {
            MarkerSyntax::Double => "double",
            MarkerSyntax::Single => "single",
            MarkerSyntax::Bracket => "bracket",
        }
    }
}

/// One marker found in a text
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Marker {
    pub kind: MarkerKind,
    pub syntax: MarkerSyntax,
    /// The marker as written, braces or brackets included
    pub raw: String,
    /// Field name (`CLIENTE` for `{{ CLIENTE: nombre }}`); `None` for a pending marker
    pub key: Option<String>,
    /// What the marker says about its content: the text after the colon of a field, the text
    /// inside a pending marker; `None` when there is nothing to say
    pub hint: Option<String>,
    /// First code point of the marker
    pub start: usize,
    /// Length in code points
    pub len: usize,
}

/// Every marker of `text`, in order of position
pub fn scan(text: &str) -> Vec<Marker> {
    if !text.contains(['{', '[']) {
        return Vec::new();
    }
    let chars: Vec<char> = text.chars().collect();
    let mut out = Vec::new();
    let mut i = 0;
    while i < chars.len() {
        let found = match chars[i] {
            '{' => field_at(&chars, i),
            '[' => pending_at(&chars, i),
            _ => None,
        };
        match found {
            Some(marker) => {
                i += marker.len;
                out.push(marker);
            }
            None => i += 1,
        }
    }
    out
}

/// Field names of `text` (each once, in order of first appearance)
pub fn field_keys(text: &str) -> Vec<String> {
    let mut keys: Vec<String> = Vec::new();
    for key in scan(text).into_iter().filter_map(|m| m.key) {
        if !keys.contains(&key) {
            keys.push(key);
        }
    }
    keys
}

/// Position of `close` after `from`, stopping at a line break or at any of `stop`
fn find_close(chars: &[char], from: usize, close: &[char], stop: &[char]) -> Option<usize> {
    let mut j = from;
    while j < chars.len() {
        if chars[j..].starts_with(close) {
            return Some(j);
        }
        if is_break(chars[j]) || stop.contains(&chars[j]) {
            return None;
        }
        j += 1;
    }
    None
}

fn is_break(c: char) -> bool {
    matches!(c, '\n' | '\r' | '\u{000B}' | '\u{2028}' | '\u{2029}')
}

fn field_at(chars: &[char], at: usize) -> Option<Marker> {
    let double = chars.get(at + 1) == Some(&'{');
    let (open, syntax) = if double { (2, MarkerSyntax::Double) } else { (1, MarkerSyntax::Single) };
    let close: &[char] = if double { &['}', '}'] } else { &['}'] };
    let end = find_close(chars, at + open, close, &['{', '}'])?;
    let inner: String = chars[at + open..end].iter().collect();
    if !double && (inner.is_empty() || inner.chars().any(char::is_whitespace)) {
        return None;
    }
    let (key, hint) = match inner.split_once(':') {
        Some((key, hint)) => (key.trim(), Some(hint.trim()).filter(|h| !h.is_empty())),
        None => (inner.trim(), None),
    };
    if key.is_empty() {
        return None;
    }
    let len = end + close.len() - at;
    Some(Marker {
        kind: MarkerKind::Field,
        syntax,
        raw: chars[at..at + len].iter().collect(),
        key: Some(key.to_string()),
        hint: hint.map(str::to_string),
        start: at,
        len,
    })
}

fn pending_at(chars: &[char], at: usize) -> Option<Marker> {
    let end = find_close(chars, at + 1, &[']'], &['['])?;
    let inner: String = chars[at + 1..end].iter().collect();
    let trimmed = inner.trim();
    if !is_pending(trimmed) {
        return None;
    }
    let filler = trimmed.chars().all(|c| matches!(c, '.' | '…' | '_' | '-' | ' ' | '\u{a0}'));
    let len = end + 1 - at;
    Some(Marker {
        kind: MarkerKind::Pending,
        syntax: MarkerSyntax::Bracket,
        raw: chars[at..at + len].iter().collect(),
        key: None,
        hint: (!filler).then(|| trimmed.to_string()),
        start: at,
        len,
    })
}

/// Bracketed text that is not a citation, a checkbox or `[sic]`
fn is_pending(inner: &str) -> bool {
    if inner.is_empty() || inner.eq_ignore_ascii_case("x") || inner.eq_ignore_ascii_case("sic") {
        return false;
    }
    let citation = inner.chars().any(|c| c.is_ascii_digit())
        && inner.chars().all(|c| c.is_ascii_digit() || matches!(c, ',' | ';' | '-' | '–' | ' '));
    !citation
}

#[cfg(test)]
mod tests {
    use super::*;

    fn kinds(text: &str) -> Vec<(&'static str, String, Option<String>, Option<String>, usize, usize)> {
        scan(text)
            .into_iter()
            .map(|m| (m.kind.as_str(), m.raw, m.key, m.hint, m.start, m.len))
            .collect()
    }

    fn s(v: &str) -> Option<String> {
        Some(v.to_string())
    }

    #[test]
    fn fields() {
        assert_eq!(
            kinds("Hola {{CLIENTE}}, {{ FECHA : dd/mm }} y {ciudad}."),
            vec![
                ("field", "{{CLIENTE}}".into(), s("CLIENTE"), None, 5, 11),
                ("field", "{{ FECHA : dd/mm }}".into(), s("FECHA"), s("dd/mm"), 18, 19),
                ("field", "{ciudad}".into(), s("ciudad"), None, 40, 8),
            ]
        );
        assert_eq!(field_keys("{{A}} {A} {{B}} {{ A }}"), vec!["A", "B"]);
    }

    #[test]
    fn not_fields() {
        for text in ["{}", "{{}}", "{{ : x }}", "{con espacio}", "{{abierto", "{{a\nb}}", "{a{b}"] {
            assert!(scan(text).iter().all(|m| m.kind != MarkerKind::Field || m.raw == "{b}"), "{}", text);
        }
    }

    #[test]
    fn pending() {
        assert_eq!(
            kinds("Alcance: [Completar: objetivos] […] [....] [____]"),
            vec![
                ("pending", "[Completar: objetivos]".into(), None, s("Completar: objetivos"), 9, 22),
                ("pending", "[…]".into(), None, None, 32, 3),
                ("pending", "[....]".into(), None, None, 36, 6),
                ("pending", "[____]".into(), None, None, 43, 6),
            ]
        );
    }

    #[test]
    fn citations_and_checkboxes_are_text() {
        assert!(scan("Según [1], [2, 3] y [4-6] [sic] [ ] [x] [X] []").is_empty());
        assert_eq!(scan("Ver [anexo 2]").len(), 1);
    }

    #[test]
    fn offsets_are_code_points() {
        let m = &scan("ñandú 🦙 {{X}}")[0];
        assert_eq!((m.start, m.len), (8, 5));
    }

    #[test]
    fn field_inside_pending_is_found_as_pending() {
        // The outer pending marker wins; its text keeps the field as the hint
        let found = scan("[{{NOMBRE}} o razón social]");
        assert_eq!(found.len(), 1);
        assert_eq!(found[0].kind, MarkerKind::Pending);
    }
}
