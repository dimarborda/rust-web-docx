//! Number formats: turns a cell value into the text Excel shows for it with a format code such
//! as `#,##0.00 "€"`, `0%`, `dd/mm/yyyy` or `[Red]-0.0;0.0`. Covers sections (positive;
//! negative;zero;text), conditions, colors, literals, grouping, scaling, percent, scientific
//! notation, simple fractions, dates, times and elapsed time. Separators and month and day
//! names follow the viewer's language.

use crate::dates::serial_to_unix_ms;

/// Separators and names of a display language
#[derive(Debug, Clone)]
pub struct Locale {
    pub decimal: char,
    pub group: char,
    months: [&'static str; 12],
    months_short: [&'static str; 12],
    /// From Sunday
    days: [&'static str; 7],
    days_short: [&'static str; 7],
    true_text: &'static str,
    false_text: &'static str,
    /// Built-in format 14 (short date) and 22 (date and time) in this language
    short_date: &'static str,
    date_time: &'static str,
}

impl Locale {
    /// The language of a BCP 47 tag (`es`, `es-CO`, `en-US`…); Spanish or English
    pub fn from_tag(tag: &str) -> Locale {
        if tag.trim().to_ascii_lowercase().starts_with("es") {
            Locale {
                decimal: ',',
                group: '.',
                months: ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"],
                months_short: ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"],
                days: ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"],
                days_short: ["dom", "lun", "mar", "mié", "jue", "vie", "sáb"],
                true_text: "VERDADERO",
                false_text: "FALSO",
                short_date: "dd/mm/yyyy",
                date_time: "dd/mm/yyyy h:mm",
            }
        } else {
            Locale {
                decimal: '.',
                group: ',',
                months: ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"],
                months_short: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
                days: ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"],
                days_short: ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"],
                true_text: "TRUE",
                false_text: "FALSE",
                short_date: "m/d/yyyy",
                date_time: "m/d/yyyy h:mm",
            }
        }
    }

    pub fn bool_text(&self, b: bool) -> &'static str {
        if b {
            self.true_text
        } else {
            self.false_text
        }
    }
}

/// The code of a built-in number format (`numFmtId` below 164 with no `<numFmt>`)
pub fn builtin_code(id: u32, locale: &Locale) -> &'static str {
    match id {
        1 => "0",
        2 => "0.00",
        3 => "#,##0",
        4 => "#,##0.00",
        5 => "$#,##0_);($#,##0)",
        6 => "$#,##0_);[Red]($#,##0)",
        7 => "$#,##0.00_);($#,##0.00)",
        8 => "$#,##0.00_);[Red]($#,##0.00)",
        9 => "0%",
        10 => "0.00%",
        11 => "0.00E+00",
        12 => "# ?/?",
        13 => "# ??/??",
        14 => locale.short_date,
        15 => "d-mmm-yy",
        16 => "d-mmm",
        17 => "mmm-yy",
        18 => "h:mm AM/PM",
        19 => "h:mm:ss AM/PM",
        20 => "h:mm",
        21 => "h:mm:ss",
        22 => locale.date_time,
        37 => "#,##0 ;(#,##0)",
        38 => "#,##0 ;[Red](#,##0)",
        39 => "#,##0.00;(#,##0.00)",
        40 => "#,##0.00;[Red](#,##0.00)",
        41 => r#"_(* #,##0_);_(* \(#,##0\);_(* "-"_);_(@_)"#,
        42 => r#"_("$"* #,##0_);_("$"* \(#,##0\);_("$"* "-"_);_(@_)"#,
        43 => r#"_(* #,##0.00_);_(* \(#,##0.00\);_(* "-"??_);_(@_)"#,
        44 => r#"_("$"* #,##0.00_);_("$"* \(#,##0.00\);_("$"* "-"??_);_(@_)"#,
        45 => "mm:ss",
        46 => "[h]:mm:ss",
        47 => "mm:ss.0",
        48 => "##0.0E+0",
        49 => "@",
        // Other ids are locale-specific dates in East Asian builds; General elsewhere
        27..=36 | 50..=58 => locale.short_date,
        _ => "General",
    }
}

/// A value to show
#[derive(Debug, Clone, Copy)]
pub enum Value<'a> {
    Number(f64),
    Text(&'a str),
    Bool(bool),
    Error(&'a str),
}

/// What a cell shows: its text and, when the format says so (`[Red]`), a color `#RRGGBB`
#[derive(Debug, Clone, PartialEq)]
pub struct Formatted {
    pub text: String,
    pub color: Option<String>,
}

impl Formatted {
    fn plain(text: String) -> Formatted {
        Formatted { text, color: None }
    }
}

#[derive(Debug, Clone, PartialEq)]
enum Tok {
    Lit(String),
    /// `0`, `#` or `?`
    Digit(char),
    Point,
    Comma,
    Percent,
    /// `E+` (true) or `E-` (false)
    Exp(bool),
    Slash,
    At,
    General,
    /// Date or time letter (`y`, `m`, `d`, `h`, `s`, lowercased; `M` = minutes once resolved)
    Date(char, usize),
    /// `[h]`, `[mm]`, `[ss]`: elapsed time
    Elapsed(char, usize),
    /// `AM/PM` (true) or `A/P` (false)
    AmPm(bool),
}

#[derive(Debug, Clone, Default)]
struct Section {
    toks: Vec<Tok>,
    color: Option<String>,
    condition: Option<(String, f64)>,
}

impl Section {
    fn is_date(&self) -> bool {
        self.toks.iter().any(|t| matches!(t, Tok::Date(..) | Tok::Elapsed(..) | Tok::AmPm(_)))
    }
    fn has_at(&self) -> bool {
        self.toks.contains(&Tok::At)
    }
    fn matches(&self, v: f64) -> bool {
        match &self.condition {
            None => true,
            Some((op, n)) => match op.as_str() {
                "<" => v < *n,
                "<=" => v <= *n,
                ">" => v > *n,
                ">=" => v >= *n,
                "=" => v == *n,
                "<>" => v != *n,
                _ => true,
            },
        }
    }
}

/// Splits a format code into its `;`-separated sections, ignoring `;` inside quotes,
/// brackets and escapes
fn split_sections(code: &str) -> Vec<String> {
    let mut sections = vec![String::new()];
    let mut chars = code.chars();
    while let Some(c) = chars.next() {
        let current = sections.last_mut().unwrap();
        match c {
            '"' => {
                current.push(c);
                for q in chars.by_ref() {
                    current.push(q);
                    if q == '"' {
                        break;
                    }
                }
            }
            '\\' | '_' | '*' => {
                current.push(c);
                if let Some(n) = chars.next() {
                    current.push(n);
                }
            }
            '[' => {
                current.push(c);
                for q in chars.by_ref() {
                    current.push(q);
                    if q == ']' {
                        break;
                    }
                }
            }
            ';' => sections.push(String::new()),
            _ => current.push(c),
        }
    }
    sections
}

fn color_name(name: &str) -> Option<&'static str> {
    Some(match name.to_ascii_lowercase().as_str() {
        "black" | "negro" => "#000000",
        "blue" | "azul" => "#0000FF",
        "cyan" | "aguamarina" => "#00FFFF",
        "green" | "verde" => "#00FF00",
        "magenta" => "#FF00FF",
        "red" | "rojo" => "#FF0000",
        "white" | "blanco" => "#FFFFFF",
        "yellow" | "amarillo" => "#FFFF00",
        _ => return None,
    })
}

fn parse_section(code: &str) -> Section {
    let mut s = Section::default();
    let chars: Vec<char> = code.chars().collect();
    let mut i = 0;
    let lit = |toks: &mut Vec<Tok>, text: String| {
        if let Some(Tok::Lit(prev)) = toks.last_mut() {
            prev.push_str(&text);
        } else {
            toks.push(Tok::Lit(text));
        }
    };
    while i < chars.len() {
        let c = chars[i];
        let rest: String = chars[i..].iter().take(5).collect::<String>().to_ascii_lowercase();
        match c {
            '"' => {
                let mut text = String::new();
                i += 1;
                while i < chars.len() && chars[i] != '"' {
                    text.push(chars[i]);
                    i += 1;
                }
                lit(&mut s.toks, text);
            }
            '\\' => {
                if let Some(n) = chars.get(i + 1) {
                    lit(&mut s.toks, n.to_string());
                }
                i += 1;
            }
            // Padding as wide as the next character, and fill: a space keeps the alignment
            '_' => {
                lit(&mut s.toks, " ".to_string());
                i += 1;
            }
            '*' => i += 1,
            '[' => {
                let mut inner = String::new();
                i += 1;
                while i < chars.len() && chars[i] != ']' {
                    inner.push(chars[i]);
                    i += 1;
                }
                let lower = inner.to_ascii_lowercase();
                if let Some(color) = color_name(&inner) {
                    s.color = Some(color.to_string());
                } else if lower.starts_with("color") {
                    let n: usize = lower[5..].parse().unwrap_or(0);
                    s.color = crate::theme::indexed_color(n.saturating_sub(1)).map(|c| format!("#{}", c));
                } else if let Some(currency) = inner.strip_prefix('$') {
                    let symbol = currency.split('-').next().unwrap_or("");
                    if !symbol.is_empty() {
                        lit(&mut s.toks, symbol.to_string());
                    }
                } else if inner.starts_with(['<', '>', '=']) {
                    let op_len = inner.chars().take_while(|c| matches!(c, '<' | '>' | '=')).count();
                    if let Ok(n) = inner[op_len..].trim().parse::<f64>() {
                        s.condition = Some((inner[..op_len].to_string(), n));
                    }
                } else if !lower.is_empty() && lower.chars().all(|ch| ch == lower.chars().next().unwrap()) && matches!(lower.chars().next(), Some('h' | 'm' | 's')) {
                    s.toks.push(Tok::Elapsed(lower.chars().next().unwrap(), lower.len()));
                }
            }
            '0' | '#' | '?' => s.toks.push(Tok::Digit(c)),
            '.' => s.toks.push(Tok::Point),
            ',' => s.toks.push(Tok::Comma),
            '%' => s.toks.push(Tok::Percent),
            '/' => s.toks.push(Tok::Slash),
            '@' => s.toks.push(Tok::At),
            'E' | 'e' if matches!(chars.get(i + 1), Some('+') | Some('-')) => {
                s.toks.push(Tok::Exp(chars[i + 1] == '+'));
                i += 1;
            }
            _ if rest.starts_with("am/pm") => {
                s.toks.push(Tok::AmPm(true));
                i += 4;
            }
            _ if rest.starts_with("a/p") => {
                s.toks.push(Tok::AmPm(false));
                i += 2;
            }
            _ if rest.starts_with("gener") && chars[i..].iter().take(7).collect::<String>().eq_ignore_ascii_case("general") => {
                s.toks.push(Tok::General);
                i += 6;
            }
            _ if matches!(c.to_ascii_lowercase(), 'y' | 'm' | 'd' | 'h' | 's') => {
                let letter = c.to_ascii_lowercase();
                let mut n = 1;
                while chars.get(i + n).map(|x| x.to_ascii_lowercase()) == Some(letter) {
                    n += 1;
                }
                s.toks.push(Tok::Date(letter, n));
                i += n - 1;
            }
            _ => lit(&mut s.toks, c.to_string()),
        }
        i += 1;
    }
    if s.is_date() {
        resolve_minutes(&mut s.toks);
    }
    s
}

/// `m` and `mm` right after hours or right before seconds are minutes, not months
fn resolve_minutes(toks: &mut [Tok]) {
    let kinds: Vec<Option<char>> = toks
        .iter()
        .map(|t| match t {
            Tok::Date(c, _) => Some(*c),
            Tok::Elapsed(c, _) => Some(*c),
            _ => None,
        })
        .collect();
    for i in 0..toks.len() {
        if let Tok::Date('m', n) = toks[i] {
            if n > 2 {
                continue;
            }
            let before = kinds[..i].iter().rev().flatten().next();
            let after = kinds[i + 1..].iter().flatten().next();
            if before == Some(&'h') || after == Some(&'s') {
                toks[i] = Tok::Date('M', n);
            }
        }
    }
}

/// The text Excel shows for `value` with the format `code`
pub fn format(value: Value, code: &str, locale: &Locale, date1904: bool) -> Formatted {
    let sections: Vec<Section> = split_sections(code).iter().map(|s| parse_section(s)).collect();
    match value {
        Value::Bool(b) => Formatted::plain(locale.bool_text(b).to_string()),
        Value::Error(e) => Formatted::plain(e.to_string()),
        Value::Text(t) => {
            let section = if sections.len() >= 4 {
                Some(&sections[3])
            } else {
                sections.iter().find(|s| s.has_at())
            };
            match section {
                Some(s) => Formatted { text: render_text(s, t), color: s.color.clone() },
                None => Formatted::plain(t.to_string()),
            }
        }
        Value::Number(v) => {
            let numeric = &sections[..sections.len().min(3)];
            let has_conditions = numeric.iter().any(|s| s.condition.is_some());
            let (section, v, signed) = if has_conditions {
                let chosen = numeric
                    .iter()
                    .position(|s| s.condition.is_some() && s.matches(v))
                    .or_else(|| numeric.iter().position(|s| s.condition.is_none()))
                    .unwrap_or(0);
                // A negative value shown by a later section loses its sign, like Excel's
                (&numeric[chosen], if chosen > 0 { v.abs() } else { v }, chosen == 0)
            } else if v < 0.0 && numeric.len() >= 2 {
                (&numeric[1], -v, false)
            } else if v == 0.0 && numeric.len() >= 3 {
                (&numeric[2], v, false)
            } else {
                (&numeric[0], v, true)
            };
            let text = if section.toks.is_empty() {
                String::new()
            } else if section.is_date() {
                render_date(section, v, locale, date1904)
            } else if section.toks.contains(&Tok::General) && !section.toks.iter().any(|t| matches!(t, Tok::Digit(_))) {
                let general = general(v.abs(), locale);
                let mut out = String::new();
                if v < 0.0 && signed {
                    out.push('-');
                }
                for t in &section.toks {
                    match t {
                        Tok::General => out.push_str(&general),
                        Tok::Lit(l) => out.push_str(l),
                        _ => {}
                    }
                }
                out
            } else {
                render_number(section, v, signed, locale)
            };
            Formatted { text, color: section.color.clone() }
        }
    }
}

fn render_text(s: &Section, text: &str) -> String {
    let mut out = String::new();
    for t in &s.toks {
        match t {
            Tok::At => out.push_str(text),
            Tok::Lit(l) => out.push_str(l),
            _ => {}
        }
    }
    out
}

/// Excel's General format: up to 11 characters, switching to scientific notation for very
/// large or very small numbers
pub fn general(v: f64, locale: &Locale) -> String {
    if v == 0.0 || !v.is_finite() {
        return "0".to_string();
    }
    let sign = if v < 0.0 { "-" } else { "" };
    let a = v.abs();
    if !(1e-9..1e11).contains(&a) {
        return format!("{}{}", sign, scientific_general(a, locale));
    }
    let int_digits = if a >= 1.0 { (a.log10().floor() as i32 + 1).max(1) } else { 1 };
    let decimals = (10 - int_digits).max(0) as usize;
    let (ip, mut fp) = decimal_parts(a, decimals);
    while fp.ends_with('0') {
        fp.pop();
    }
    if ip.len() > 11 {
        return format!("{}{}", sign, scientific_general(a, locale));
    }
    if fp.is_empty() {
        format!("{}{}", sign, ip)
    } else {
        format!("{}{}{}{}", sign, ip, locale.decimal, fp)
    }
}

fn scientific_general(a: f64, locale: &Locale) -> String {
    let mut exp = a.log10().floor() as i32;
    let (mut ip, mut fp) = decimal_parts(a / 10f64.powi(exp), 5);
    if ip.len() > 1 {
        exp += 1;
        (ip, fp) = decimal_parts(a / 10f64.powi(exp), 5);
    }
    while fp.ends_with('0') {
        fp.pop();
    }
    let mantissa = if fp.is_empty() { ip } else { format!("{}{}{}", ip, locale.decimal, fp) };
    format!("{}E{}{:02}", mantissa, if exp < 0 { '-' } else { '+' }, exp.abs())
}

/// Integer and decimal digits of `x` (≥ 0) rounded half up to `decimals` places, working on its
/// 15 significant digits as Excel does (so 2.675 rounds to 2.68)
fn decimal_parts(x: f64, decimals: usize) -> (String, String) {
    if x == 0.0 || !x.is_finite() {
        return ("0".to_string(), "0".repeat(decimals));
    }
    let sci = format!("{:.14e}", x);
    let (mantissa, exp) = sci.split_once('e').unwrap();
    let exp: i32 = exp.parse().unwrap_or(0);
    let mut digits: String = mantissa.chars().filter(|c| c.is_ascii_digit()).collect();
    let mut point = exp + 1;
    if point <= 0 {
        digits = format!("{}{}", "0".repeat((-point) as usize), digits);
        point = 0;
    }
    let point = point as usize;
    if digits.len() < point + decimals + 1 {
        digits.push_str(&"0".repeat(point + decimals + 1 - digits.len()));
    }
    let round_up = digits.as_bytes()[point + decimals] >= b'5';
    let mut kept: Vec<u8> = digits.as_bytes()[..point + decimals].to_vec();
    if round_up {
        let mut i = kept.len();
        loop {
            if i == 0 {
                kept.insert(0, b'1');
                break;
            }
            i -= 1;
            if kept[i] == b'9' {
                kept[i] = b'0';
            } else {
                kept[i] += 1;
                break;
            }
        }
    }
    let int_len = kept.len() - decimals;
    let ip = String::from_utf8(kept[..int_len].to_vec()).unwrap();
    let ip = ip.trim_start_matches('0');
    let ip = if ip.is_empty() { "0" } else { ip };
    (ip.to_string(), String::from_utf8(kept[int_len..].to_vec()).unwrap())
}

fn group_digits(digits: &str, sep: char) -> String {
    let mut out = String::new();
    let len = digits.len();
    for (i, c) in digits.chars().enumerate() {
        if i > 0 && (len - i) % 3 == 0 {
            out.push(sep);
        }
        out.push(c);
    }
    out
}

fn render_number(s: &Section, v: f64, signed: bool, locale: &Locale) -> String {
    let toks = &s.toks;
    let mut v = v;
    let negative = v < 0.0 && signed;
    v = v.abs();
    for t in toks {
        if *t == Tok::Percent {
            v *= 100.0;
        }
    }
    if let Some(slash) = toks.iter().position(|t| *t == Tok::Slash) {
        if toks[..slash].iter().any(|t| matches!(t, Tok::Digit(_))) {
            return render_fraction(toks, slash, v, negative);
        }
    }
    let exp_at = toks.iter().position(|t| matches!(t, Tok::Exp(_)));
    let mantissa_end = exp_at.unwrap_or(toks.len());
    let point_at = toks[..mantissa_end].iter().position(|t| *t == Tok::Point);
    let int_end = point_at.unwrap_or(mantissa_end);

    // Commas right after the last digit placeholder scale by 1000 each; commas between
    // integer placeholders group thousands
    let last_digit = toks[..mantissa_end].iter().rposition(|t| matches!(t, Tok::Digit(_)));
    let scaling = match last_digit {
        Some(last) => toks[last + 1..mantissa_end].iter().take_while(|t| **t == Tok::Comma).count() as i32,
        None => 0,
    };
    let last_int_digit = toks[..int_end].iter().rposition(|t| matches!(t, Tok::Digit(_)));
    let grouping = toks[..int_end].iter().enumerate().any(|(i, t)| *t == Tok::Comma && last_int_digit.is_some_and(|l| i < l));
    v /= 1000f64.powi(scaling);

    let int_slots: Vec<(usize, char)> = toks[..int_end].iter().enumerate().filter_map(|(i, t)| if let Tok::Digit(c) = t { Some((i, *c)) } else { None }).collect();
    let dec_slots: Vec<(usize, char)> = match point_at {
        Some(p) => toks[p + 1..mantissa_end].iter().enumerate().filter_map(|(i, t)| if let Tok::Digit(c) = t { Some((p + 1 + i, *c)) } else { None }).collect(),
        None => Vec::new(),
    };
    let exp_slots: Vec<char> = match exp_at {
        Some(e) => toks[e + 1..].iter().filter_map(|t| if let Tok::Digit(c) = t { Some(*c) } else { None }).collect(),
        None => Vec::new(),
    };

    let mut exponent = 0i32;
    let (ip, fp) = if exp_at.is_some() {
        let step = if int_slots.len() > 1 && int_slots.iter().any(|(_, c)| *c == '#') { int_slots.len() as i32 } else { 1 };
        if v != 0.0 {
            exponent = (v.log10().floor() as i32).div_euclid(step) * step;
        }
        let mut parts = decimal_parts(v / 10f64.powi(exponent), dec_slots.len());
        if v != 0.0 && parts.0.len() as i32 > step {
            exponent += step;
            parts = decimal_parts(v / 10f64.powi(exponent), dec_slots.len());
        }
        parts
    } else {
        decimal_parts(v, dec_slots.len())
    };
    let int_digits = if ip == "0" { String::new() } else { ip };
    let nonzero = !int_digits.is_empty() || fp.chars().any(|c| c != '0');

    // Integer digits per placeholder, right to left; extra digits go to the first one
    let mut int_out: Vec<String> = vec![String::new(); int_slots.len()];
    if grouping {
        let zeros_from = int_slots.iter().position(|(_, c)| *c == '0').map(|p| int_slots.len() - p).unwrap_or(0);
        let padded = format!("{}{}", "0".repeat(zeros_from.saturating_sub(int_digits.len())), int_digits);
        if let Some(first) = int_out.first_mut() {
            *first = group_digits(&padded, locale.group);
        }
    } else {
        let digits: Vec<char> = int_digits.chars().collect();
        let mut next = digits.len();
        for slot in (0..int_slots.len()).rev() {
            if next > 0 {
                next -= 1;
                int_out[slot] = digits[next].to_string();
            } else {
                int_out[slot] = match int_slots[slot].1 {
                    '0' => "0".to_string(),
                    '?' => " ".to_string(),
                    _ => String::new(),
                };
            }
        }
        if next > 0 && !int_out.is_empty() {
            let extra: String = digits[..next].iter().collect();
            int_out[0] = format!("{}{}", extra, int_out[0]);
        }
    }
    // Decimal digits left to right; trailing zeros vanish under `#` and become spaces under `?`
    let fp_chars: Vec<char> = fp.chars().collect();
    let last_nonzero = fp_chars.iter().rposition(|c| *c != '0');
    let dec_out: Vec<String> = dec_slots
        .iter()
        .enumerate()
        .map(|(i, (_, c))| {
            if last_nonzero.is_some_and(|l| i <= l) || *c == '0' {
                fp_chars[i].to_string()
            } else if *c == '?' {
                " ".to_string()
            } else {
                String::new()
            }
        })
        .collect();

    let mut out = String::new();
    if negative && nonzero {
        out.push('-');
    }
    let mut int_index = 0;
    let mut dec_index = 0;
    let mut exp_written = false;
    for (i, t) in toks.iter().enumerate() {
        match t {
            Tok::Lit(l) => out.push_str(l),
            Tok::Digit(_) if i < int_end => {
                out.push_str(&int_out[int_index]);
                int_index += 1;
            }
            Tok::Digit(_) if i < mantissa_end => {
                out.push_str(&dec_out[dec_index]);
                dec_index += 1;
            }
            Tok::Digit(_) => {
                if !exp_written {
                    let digits = exponent.abs().to_string();
                    let pad = exp_slots.iter().filter(|c| **c == '0').count();
                    out.push_str(&format!("{}{}", "0".repeat(pad.saturating_sub(digits.len())), digits));
                    exp_written = true;
                }
            }
            Tok::Point if Some(i) == point_at => {
                if int_slots.is_empty() && !int_digits.is_empty() {
                    out.push_str(&int_digits);
                }
                out.push(locale.decimal);
            }
            Tok::Point => out.push('.'),
            Tok::Percent => out.push('%'),
            Tok::Exp(plus) => {
                out.push('E');
                if exponent < 0 {
                    out.push('-');
                } else if *plus {
                    out.push('+');
                }
            }
            Tok::Slash => out.push('/'),
            Tok::General => out.push_str(&general(v, locale)),
            Tok::Date(c, n) => out.push_str(&c.to_string().repeat(*n)),
            Tok::Comma | Tok::At | Tok::Elapsed(..) | Tok::AmPm(_) => {}
        }
    }
    out
}

/// `# ?/?`, `# ??/??`, `?/8`…: whole part, then the closest fraction with the denominators
/// the placeholders allow (or the fixed one written)
fn render_fraction(toks: &[Tok], slash: usize, v: f64, negative: bool) -> String {
    let den_digits: String = toks[slash + 1..]
        .iter()
        .take_while(|t| matches!(t, Tok::Digit(_)) || matches!(t, Tok::Lit(l) if l.chars().all(|c| c.is_ascii_digit())))
        .map(|t| match t {
            Tok::Digit(c) => c.to_string(),
            Tok::Lit(l) => l.clone(),
            _ => String::new(),
        })
        .collect();
    let fixed: Option<u64> = if den_digits.chars().all(|c| c.is_ascii_digit()) && !den_digits.is_empty() { den_digits.parse().ok() } else { None };
    let max_den = fixed.unwrap_or_else(|| 10u64.pow(den_digits.len().clamp(1, 4) as u32) - 1);
    // Numerator placeholders are the run right before the slash; digits before them, if any,
    // are the whole part
    let num_start = toks[..slash].iter().rposition(|t| !matches!(t, Tok::Digit(_))).map(|p| p + 1).unwrap_or(0);
    let has_whole = toks[..num_start].iter().any(|t| matches!(t, Tok::Digit(_)));
    let mut whole = if has_whole { v.trunc() } else { 0.0 };
    let frac = v - whole;
    let (mut num, den) = match fixed {
        Some(d) => ((frac * d as f64).round() as u64, d),
        None => {
            let mut best = (0u64, 1u64, f64::MAX);
            for d in 1..=max_den {
                let n = (frac * d as f64).round();
                let err = (frac - n / d as f64).abs();
                if err < best.2 - 1e-12 {
                    best = (n as u64, d, err);
                }
            }
            (best.0, best.1)
        }
    };
    if has_whole && num == den {
        whole += 1.0;
        num = 0;
    }
    let mut out = String::new();
    if negative && (whole != 0.0 || num != 0) {
        out.push('-');
    }
    if has_whole && (whole != 0.0 || num == 0) {
        out.push_str(&format!("{}", whole as u64));
    }
    if num != 0 {
        if has_whole && whole != 0.0 {
            out.push(' ');
        }
        out.push_str(&format!("{}/{}", num, den));
    }
    out
}

/// Year, month and day of a day count since 1970-01-01
fn civil_from_days(z: i64) -> (i64, u32, u32) {
    let z = z + 719_468;
    let era = z.div_euclid(146_097);
    let doe = z - era * 146_097;
    let yoe = (doe - doe / 1460 + doe / 36_524 - doe / 146_096) / 365;
    let y = yoe + era * 400;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let d = (doy - (153 * mp + 2) / 5 + 1) as u32;
    let m = if mp < 10 { mp + 3 } else { mp - 9 } as u32;
    (if m <= 2 { y + 1 } else { y }, m, d)
}

fn render_date(s: &Section, serial: f64, locale: &Locale, date1904: bool) -> String {
    if !(0.0..2_958_466.0).contains(&serial) {
        return "#".repeat(8);
    }
    // Fractions of a second shown after the seconds (`ss.00`)
    let mut sub_digits = 0;
    for (i, t) in s.toks.iter().enumerate() {
        if *t == Tok::Point && i > 0 && matches!(s.toks[..i].iter().rev().find(|t| !matches!(t, Tok::Lit(_))), Some(Tok::Date('s', _)) | Some(Tok::Elapsed('s', _))) {
            sub_digits = s.toks[i + 1..].iter().take_while(|t| **t == Tok::Digit('0')).count().min(3);
        }
    }
    let unit = 10f64.powi(3 - sub_digits as i32);
    let total_ms = (serial * 86_400_000.0 / unit).round() * unit;
    let day_serial = (total_ms / 86_400_000.0).floor();
    let time_ms = total_ms - day_serial * 86_400_000.0;

    let (year, month, day) = if !date1904 && day_serial == 0.0 {
        (1900, 1, 0)
    } else if !date1904 && day_serial == 60.0 {
        (1900, 2, 29)
    } else {
        let days = (serial_to_unix_ms(day_serial, date1904) / 86_400_000.0).round() as i64;
        civil_from_days(days)
    };
    let weekday = if !date1904 && day_serial < 61.0 {
        // Excel's 1900 calendar is one day off before March 1900
        ((day_serial as i64 + 6).rem_euclid(7)) as usize
    } else {
        let days = (serial_to_unix_ms(day_serial, date1904) / 86_400_000.0).round() as i64;
        ((days + 4).rem_euclid(7)) as usize
    };
    let secs = (time_ms / 1000.0).floor() as u64;
    let hour = secs / 3600;
    let minute = (secs / 60) % 60;
    let second = secs % 60;
    let has_ampm = s.toks.iter().any(|t| matches!(t, Tok::AmPm(_)));

    let mut out = String::new();
    let mut i = 0;
    while i < s.toks.len() {
        match &s.toks[i] {
            Tok::Lit(l) => out.push_str(l),
            Tok::Date('y', n) => {
                if *n <= 2 {
                    out.push_str(&format!("{:02}", year.rem_euclid(100)));
                } else {
                    out.push_str(&format!("{:04}", year));
                }
            }
            Tok::Date('m', n) => match n {
                1 => out.push_str(&month.to_string()),
                2 => out.push_str(&format!("{:02}", month)),
                3 => out.push_str(locale.months_short[month as usize - 1]),
                4 => out.push_str(locale.months[month as usize - 1]),
                _ => out.push_str(&locale.months[month as usize - 1].chars().next().unwrap().to_uppercase().to_string()),
            },
            Tok::Date('M', n) => {
                if *n == 1 {
                    out.push_str(&minute.to_string());
                } else {
                    out.push_str(&format!("{:02}", minute));
                }
            }
            Tok::Date('d', n) => match n {
                1 => out.push_str(&day.to_string()),
                2 => out.push_str(&format!("{:02}", day)),
                3 => out.push_str(locale.days_short[weekday]),
                _ => out.push_str(locale.days[weekday]),
            },
            Tok::Date('h', n) => {
                let h = if has_ampm {
                    match hour % 12 {
                        0 => 12,
                        h => h,
                    }
                } else {
                    hour % 24
                };
                if *n == 1 {
                    out.push_str(&h.to_string());
                } else {
                    out.push_str(&format!("{:02}", h));
                }
            }
            Tok::Date('s', n) => {
                if *n == 1 {
                    out.push_str(&second.to_string());
                } else {
                    out.push_str(&format!("{:02}", second));
                }
            }
            Tok::Date(c, n) => out.push_str(&c.to_string().repeat(*n)),
            Tok::Elapsed(c, n) => {
                let total_secs = (total_ms / 1000.0).floor() as u64;
                let value = match c {
                    'h' => total_secs / 3600,
                    'm' => total_secs / 60,
                    _ => total_secs,
                };
                out.push_str(&format!("{:0width$}", value, width = *n));
            }
            Tok::AmPm(full) => {
                let pm = hour % 24 >= 12;
                out.push_str(match (full, pm) {
                    (true, false) => "AM",
                    (true, true) => "PM",
                    (false, false) => "A",
                    (false, true) => "P",
                });
            }
            Tok::Point if sub_digits > 0 => {
                out.push(locale.decimal);
                let frac = ((time_ms % 1000.0) / unit).round() as u64;
                out.push_str(&format!("{:0width$}", frac, width = sub_digits));
                i += sub_digits;
            }
            Tok::Point => out.push('.'),
            Tok::Digit(c) => out.push(*c),
            Tok::Comma => out.push(','),
            Tok::Percent => out.push('%'),
            Tok::Slash => out.push('/'),
            Tok::At | Tok::General | Tok::Exp(_) => {}
        }
        i += 1;
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    fn en(v: f64, code: &str) -> String {
        format(Value::Number(v), code, &Locale::from_tag("en"), false).text
    }
    fn es(v: f64, code: &str) -> String {
        format(Value::Number(v), code, &Locale::from_tag("es"), false).text
    }

    #[test]
    fn general_numbers() {
        assert_eq!(en(0.0, "General"), "0");
        assert_eq!(en(1234.0, "General"), "1234");
        assert_eq!(en(-12.5, "General"), "-12.5");
        assert_eq!(es(12.5, "General"), "12,5");
        assert_eq!(en(0.1 + 0.2, "General"), "0.3");
        assert_eq!(en(1.0 / 3.0, "General"), "0.333333333");
        assert_eq!(en(123456789012.0, "General"), "1.23457E+11");
        assert_eq!(en(0.0000000001234, "General"), "1.234E-10");
    }

    #[test]
    fn fixed_grouped_and_percent() {
        assert_eq!(en(1234.5, "0.00"), "1234.50");
        assert_eq!(en(2.675, "0.00"), "2.68");
        assert_eq!(en(1234567.891, "#,##0.00"), "1,234,567.89");
        assert_eq!(es(1234567.891, "#,##0.00"), "1.234.567,89");
        assert_eq!(en(0.0, "#,##0"), "0");
        assert_eq!(en(0.5, "#.##"), ".5");
        assert_eq!(en(0.256, "0%"), "26%");
        assert_eq!(en(0.25678, "0.00%"), "25.68%");
        assert_eq!(en(-5.0, "0.0"), "-5.0");
        assert_eq!(en(1234567.0, "#,##0,"), "1,235");
        assert_eq!(en(1234567.0, "0.0,,\" M\""), "1.2 M");
        assert_eq!(en(5551234567.0, "(###) ###-####"), "(555) 123-4567");
        assert_eq!(en(7.0, "000"), "007");
        assert_eq!(en(1.5, "0.0??"), "1.5  ");
    }

    #[test]
    fn literals_currency_and_sections() {
        assert_eq!(es(60.0, "#,##0.00 \"€\""), "60,00 €");
        assert_eq!(es(1234.5, "[$€-C0A] #,##0.00"), "€ 1.234,50");
        assert_eq!(en(-1234.0, "$#,##0_);($#,##0)"), "($1,234)");
        assert_eq!(en(1234.0, "$#,##0_);($#,##0)"), "$1,234 ");
        let red = format(Value::Number(-3.0), "0.0;[Red]-0.0", &Locale::from_tag("en"), false);
        assert_eq!(red, Formatted { text: "-3.0".to_string(), color: Some("#FF0000".to_string()) });
        assert_eq!(en(0.0, "0;-0;\"cero\""), "cero");
        assert_eq!(en(0.0, "0;-0;;@"), "");
        assert_eq!(en(150.0, "[>100]\"alto\";[<=100]\"bajo\""), "alto");
        assert_eq!(en(50.0, "[>100]\"alto\";[<=100]\"bajo\""), "bajo");
        assert_eq!(en(0.0, r#"_(* #,##0_);_(* \(#,##0\);_(* "-"_);_(@_)"#), " - ");
    }

    #[test]
    fn scientific_and_fractions() {
        assert_eq!(en(12345.0, "0.00E+00"), "1.23E+04");
        assert_eq!(en(0.00012, "0.00E+00"), "1.20E-04");
        assert_eq!(en(12345.0, "##0.0E+0"), "12.3E+3");
        assert_eq!(en(1.5, "# ?/?"), "1 1/2");
        assert_eq!(en(0.75, "# ?/?"), "3/4");
        assert_eq!(en(3.0, "# ?/?"), "3");
        assert_eq!(en(0.333, "# ??/??"), "1/3");
        assert_eq!(en(0.3, "# ?/8"), "2/8");
    }

    #[test]
    fn text_bool_and_errors() {
        let l = Locale::from_tag("es");
        assert_eq!(format(Value::Text("hola"), "0.00", &l, false).text, "hola");
        assert_eq!(format(Value::Text("hola"), "\"Sr. \"@", &l, false).text, "Sr. hola");
        assert_eq!(format(Value::Text("x"), "0;-0;0;\"[\"@\"]\"", &l, false).text, "[x]");
        assert_eq!(format(Value::Bool(true), "General", &l, false).text, "VERDADERO");
        assert_eq!(format(Value::Error("#DIV/0!"), "0.00", &l, false).text, "#DIV/0!");
    }

    #[test]
    fn dates_and_times() {
        // 2026-01-15 is serial 46037; 2026-10-17 is 46312 (a Saturday)
        assert_eq!(es(46037.0, "dd/mm/yyyy"), "15/01/2026");
        assert_eq!(en(46037.0, "m/d/yy"), "1/15/26");
        assert_eq!(es(46312.0, "dddd d \"de\" mmmm \"de\" yyyy"), "sábado 17 de octubre de 2026");
        assert_eq!(en(46312.0, "ddd, mmm d"), "Sat, Oct 17");
        assert_eq!(en(46023.520833333336, "m/d/yy h:mm"), "1/1/26 12:30");
        assert_eq!(en(0.75, "h:mm AM/PM"), "6:00 PM");
        assert_eq!(en(0.5 + 1.0 / 86400.0 * 5.25, "hh:mm:ss.00"), "12:00:05.25");
        assert_eq!(en(1.5, "[h]:mm"), "36:00");
        assert_eq!(en(0.0104166666, "mm:ss"), "15:00");
        assert_eq!(en(60.0, "yyyy-mm-dd"), "1900-02-29");
        assert_eq!(en(-1.0, "dd/mm/yyyy"), "########");
        assert_eq!(es(46037.0, builtin_code(14, &Locale::from_tag("es"))), "15/01/2026");
        assert_eq!(en(46037.0, builtin_code(14, &Locale::from_tag("en"))), "1/15/2026");
        assert_eq!(en(46037.0 - 1462.0, "yyyy-mm-dd"), "2022-01-14", "1900 system serial read as 1900");
        assert_eq!(format(Value::Number(46037.0 - 1462.0), "yyyy-mm-dd", &Locale::from_tag("en"), true).text, "2026-01-15");
    }
}
