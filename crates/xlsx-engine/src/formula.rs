//! Formulas as Excel writes them in `<f>` (English function names, `,` between arguments, no
//! leading `=`): parsed into an expression tree that `calc` evaluates, shifted for the cells
//! of a shared formula, and printed back.
//!
//! What the parser does not handle is reported as `Unsupported` rather than invalid, so such
//! formulas keep the value Excel cached: references to other workbooks (`[1]Hoja!A1`), 3-D
//! references (`Hoja1:Hoja3!A1`), table references (`Tabla1[Columna]`), whole rows (`1:1`) and
//! the union and intersection operators.

use crate::cell_ref::{column_name, MAX_COL, MAX_ROW};

/// One corner of a reference; `$` makes a row or column absolute (kept when a shared formula
/// is shifted)
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Pos {
    pub row: u32,
    pub col: u32,
    pub row_abs: bool,
    pub col_abs: bool,
}

/// A cell or a range, optionally on another sheet
#[derive(Debug, Clone, PartialEq)]
pub struct Area {
    pub sheet: Option<String>,
    pub first: Pos,
    pub last: Pos,
    /// Whole columns (`A:C`): rows 1 to the end of the sheet
    pub whole_cols: bool,
    /// Written as a range (`A1:A1` rather than `A1`)
    pub is_range: bool,
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum Op {
    Add,
    Sub,
    Mul,
    Div,
    Pow,
    Concat,
    Eq,
    Ne,
    Lt,
    Le,
    Gt,
    Ge,
}

impl Op {
    fn text(self) -> &'static str {
        match self {
            Op::Add => "+",
            Op::Sub => "-",
            Op::Mul => "*",
            Op::Div => "/",
            Op::Pow => "^",
            Op::Concat => "&",
            Op::Eq => "=",
            Op::Ne => "<>",
            Op::Lt => "<",
            Op::Le => "<=",
            Op::Gt => ">",
            Op::Ge => ">=",
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum Expr {
    Number(f64),
    Text(String),
    Bool(bool),
    /// Error literal such as `#N/A`
    Error(String),
    Ref(Area),
    /// A defined name
    Name(String),
    /// `-x` or `+x`
    Neg(Box<Expr>),
    Plus(Box<Expr>),
    Percent(Box<Expr>),
    Binary(Op, Box<Expr>, Box<Expr>),
    /// Function call; the name is upper case without the `_xlfn.` / `_xlws.` prefixes
    Call(String, Vec<Expr>),
    /// An omitted argument (`IF(A1,,2)`)
    Missing,
    /// `{1,2;3,4}`: constant rows
    Array(Vec<Vec<Expr>>),
    Paren(Box<Expr>),
}

#[derive(Debug, Clone, PartialEq)]
pub enum ParseError {
    /// Not a formula
    Invalid(String),
    /// A formula this engine does not evaluate
    Unsupported(String),
}

impl ParseError {
    pub fn message(&self) -> &str {
        match self {
            ParseError::Invalid(m) | ParseError::Unsupported(m) => m,
        }
    }
}

type Parsed<T> = Result<T, ParseError>;

fn invalid<T>(message: impl Into<String>) -> Parsed<T> {
    Err(ParseError::Invalid(message.into()))
}

fn unsupported<T>(message: impl Into<String>) -> Parsed<T> {
    Err(ParseError::Unsupported(message.into()))
}

/// Functions Excel stores with the `_xlfn.` prefix (added after Excel 2007)
const FUTURE_FUNCTIONS: &[&str] = &[
    "CONCAT", "TEXTJOIN", "IFS", "SWITCH", "MAXIFS", "MINIFS", "XLOOKUP", "XMATCH", "IFNA", "DAYS", "CEILING.MATH",
    "FLOOR.MATH", "XOR", "FILTER", "SORT", "UNIQUE", "SEQUENCE", "LET", "STDEV.S", "STDEV.P",
];

const ERRORS: &[&str] = &["#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#GETTING_DATA"];

#[derive(Debug, Clone, PartialEq)]
enum Tok {
    Number(f64),
    Text(String),
    Error(String),
    /// A word: cell reference, column, function or defined name, with `$` kept
    Word(String),
    /// `Hoja!` or `'Mi hoja'!`
    Sheet(String),
    Op(Op),
    Percent,
    Colon,
    Comma,
    Semicolon,
    Open,
    Close,
    BraceOpen,
    BraceClose,
}

fn is_word_char(c: char) -> bool {
    c.is_alphanumeric() || matches!(c, '_' | '.' | '$' | '\\' | '?')
}

fn tokenize(text: &str) -> Parsed<Vec<Tok>> {
    let chars: Vec<char> = text.chars().collect();
    let mut toks = Vec::new();
    let mut i = 0;
    while i < chars.len() {
        let c = chars[i];
        match c {
            ' ' | '\n' | '\r' | '\t' => {
                i += 1;
            }
            '"' => {
                let mut s = String::new();
                i += 1;
                loop {
                    match chars.get(i) {
                        None => return invalid("Falta cerrar unas comillas."),
                        Some('"') if chars.get(i + 1) == Some(&'"') => {
                            s.push('"');
                            i += 2;
                        }
                        Some('"') => {
                            i += 1;
                            break;
                        }
                        Some(ch) => {
                            s.push(*ch);
                            i += 1;
                        }
                    }
                }
                toks.push(Tok::Text(s));
            }
            '\'' => {
                let mut s = String::new();
                i += 1;
                loop {
                    match chars.get(i) {
                        None => return invalid("Falta cerrar el nombre de una hoja."),
                        Some('\'') if chars.get(i + 1) == Some(&'\'') => {
                            s.push('\'');
                            i += 2;
                        }
                        Some('\'') => {
                            i += 1;
                            break;
                        }
                        Some(ch) => {
                            s.push(*ch);
                            i += 1;
                        }
                    }
                }
                if chars.get(i) != Some(&'!') {
                    return invalid("Falta '!' después del nombre de la hoja.");
                }
                i += 1;
                if s.contains(':') {
                    return unsupported("referencias 3D a varias hojas");
                }
                toks.push(Tok::Sheet(s));
            }
            '#' => {
                let rest: String = chars[i..].iter().collect();
                let upper = rest.to_uppercase();
                match ERRORS.iter().find(|e| upper.starts_with(*e)) {
                    Some(e) => {
                        toks.push(Tok::Error(e.to_string()));
                        i += e.chars().count();
                    }
                    None => return invalid(format!("Error desconocido en '{}'.", rest)),
                }
            }
            '[' => return unsupported("referencias a otros libros o a tablas"),
            '0'..='9' | '.' if c != '.' || chars.get(i + 1).is_some_and(|d| d.is_ascii_digit()) => {
                let start = i;
                while i < chars.len() && (chars[i].is_ascii_digit() || chars[i] == '.') {
                    i += 1;
                }
                // A word that starts with digits is not a number (e.g. a row range 1:1 is
                // handled by the parser, a name cannot start with a digit)
                if i < chars.len() && matches!(chars[i], 'e' | 'E') {
                    let mut j = i + 1;
                    if j < chars.len() && matches!(chars[j], '+' | '-') {
                        j += 1;
                    }
                    if j < chars.len() && chars[j].is_ascii_digit() {
                        i = j;
                        while i < chars.len() && chars[i].is_ascii_digit() {
                            i += 1;
                        }
                    }
                }
                let s: String = chars[start..i].iter().collect();
                let n = s.parse::<f64>().or_else(|_| invalid(format!("'{}' no es un número.", s)))?;
                toks.push(Tok::Number(n));
            }
            '+' => {
                toks.push(Tok::Op(Op::Add));
                i += 1;
            }
            '-' => {
                toks.push(Tok::Op(Op::Sub));
                i += 1;
            }
            '*' => {
                toks.push(Tok::Op(Op::Mul));
                i += 1;
            }
            '/' => {
                toks.push(Tok::Op(Op::Div));
                i += 1;
            }
            '^' => {
                toks.push(Tok::Op(Op::Pow));
                i += 1;
            }
            '&' => {
                toks.push(Tok::Op(Op::Concat));
                i += 1;
            }
            '=' => {
                toks.push(Tok::Op(Op::Eq));
                i += 1;
            }
            '<' => {
                match chars.get(i + 1) {
                    Some('>') => {
                        toks.push(Tok::Op(Op::Ne));
                        i += 2;
                    }
                    Some('=') => {
                        toks.push(Tok::Op(Op::Le));
                        i += 2;
                    }
                    _ => {
                        toks.push(Tok::Op(Op::Lt));
                        i += 1;
                    }
                }
            }
            '>' => {
                if chars.get(i + 1) == Some(&'=') {
                    toks.push(Tok::Op(Op::Ge));
                    i += 2;
                } else {
                    toks.push(Tok::Op(Op::Gt));
                    i += 1;
                }
            }
            '%' => {
                toks.push(Tok::Percent);
                i += 1;
            }
            ':' => {
                toks.push(Tok::Colon);
                i += 1;
            }
            ',' => {
                toks.push(Tok::Comma);
                i += 1;
            }
            ';' => {
                toks.push(Tok::Semicolon);
                i += 1;
            }
            '(' => {
                toks.push(Tok::Open);
                i += 1;
            }
            ')' => {
                toks.push(Tok::Close);
                i += 1;
            }
            '{' => {
                toks.push(Tok::BraceOpen);
                i += 1;
            }
            '}' => {
                toks.push(Tok::BraceClose);
                i += 1;
            }
            c if is_word_char(c) => {
                let start = i;
                while i < chars.len() && is_word_char(chars[i]) {
                    i += 1;
                }
                let word: String = chars[start..i].iter().collect();
                if chars.get(i) == Some(&'!') {
                    i += 1;
                    toks.push(Tok::Sheet(word));
                } else if chars.get(i) == Some(&'[') {
                    return unsupported("referencias a tablas");
                } else if chars.get(i) == Some(&':') && sheet_range_ahead(&chars[i + 1..]) {
                    return unsupported("referencias 3D a varias hojas");
                } else {
                    toks.push(Tok::Word(word));
                }
            }
            other => return invalid(format!("Carácter inesperado '{}'.", other)),
        }
    }
    Ok(toks)
}

/// After `Hoja1:`, whether a `Hoja3!` follows (a 3-D reference)
fn sheet_range_ahead(rest: &[char]) -> bool {
    let n = rest.iter().take_while(|c| is_word_char(**c)).count();
    n > 0 && rest.get(n) == Some(&'!')
}

/// `$B$3` → corner; `None` when the word is not a cell reference
fn parse_pos(word: &str) -> Option<Pos> {
    let bytes = word.as_bytes();
    let mut i = 0;
    let col_abs = bytes.first() == Some(&b'$');
    if col_abs {
        i += 1;
    }
    let letters_start = i;
    while i < bytes.len() && bytes[i].is_ascii_alphabetic() {
        i += 1;
    }
    let letters = &word[letters_start..i];
    if letters.is_empty() || letters.len() > 3 {
        return None;
    }
    let row_abs = bytes.get(i) == Some(&b'$');
    if row_abs {
        i += 1;
    }
    let digits = &word[i..];
    if digits.is_empty() || !digits.bytes().all(|b| b.is_ascii_digit()) || digits.starts_with('0') || digits.len() > 7 {
        return None;
    }
    let col = letters.bytes().fold(0u32, |acc, b| acc * 26 + (b.to_ascii_uppercase() - b'A' + 1) as u32);
    let row: u32 = digits.parse().ok()?;
    ((1..=MAX_ROW).contains(&row) && (1..=MAX_COL).contains(&col)).then_some(Pos { row, col, row_abs, col_abs })
}

/// `$C` → (column, absolute); `None` when the word is not a column
fn parse_col(word: &str) -> Option<(u32, bool)> {
    let (abs, letters) = match word.strip_prefix('$') {
        Some(rest) => (true, rest),
        None => (false, word),
    };
    if letters.is_empty() || letters.len() > 3 || !letters.bytes().all(|b| b.is_ascii_alphabetic()) {
        return None;
    }
    let col = letters.bytes().fold(0u32, |acc, b| acc * 26 + (b.to_ascii_uppercase() - b'A' + 1) as u32);
    (col <= MAX_COL).then_some((col, abs))
}

struct Parser {
    toks: Vec<Tok>,
    at: usize,
}

impl Parser {
    fn peek(&self) -> Option<&Tok> {
        self.toks.get(self.at)
    }

    fn next(&mut self) -> Option<Tok> {
        let tok = self.toks.get(self.at).cloned();
        self.at += 1;
        tok
    }

    fn eat(&mut self, tok: &Tok) -> bool {
        if self.peek() == Some(tok) {
            self.at += 1;
            true
        } else {
            false
        }
    }

    fn comparison(&mut self) -> Parsed<Expr> {
        let mut left = self.concat()?;
        while let Some(Tok::Op(op @ (Op::Eq | Op::Ne | Op::Lt | Op::Le | Op::Gt | Op::Ge))) = self.peek().cloned() {
            self.at += 1;
            let right = self.concat()?;
            left = Expr::Binary(op, Box::new(left), Box::new(right));
        }
        Ok(left)
    }

    fn concat(&mut self) -> Parsed<Expr> {
        let mut left = self.additive()?;
        while self.eat(&Tok::Op(Op::Concat)) {
            let right = self.additive()?;
            left = Expr::Binary(Op::Concat, Box::new(left), Box::new(right));
        }
        Ok(left)
    }

    fn additive(&mut self) -> Parsed<Expr> {
        let mut left = self.multiplicative()?;
        while let Some(Tok::Op(op @ (Op::Add | Op::Sub))) = self.peek().cloned() {
            self.at += 1;
            let right = self.multiplicative()?;
            left = Expr::Binary(op, Box::new(left), Box::new(right));
        }
        Ok(left)
    }

    fn multiplicative(&mut self) -> Parsed<Expr> {
        let mut left = self.power()?;
        while let Some(Tok::Op(op @ (Op::Mul | Op::Div))) = self.peek().cloned() {
            self.at += 1;
            let right = self.power()?;
            left = Expr::Binary(op, Box::new(left), Box::new(right));
        }
        Ok(left)
    }

    fn power(&mut self) -> Parsed<Expr> {
        let mut left = self.percent()?;
        while self.eat(&Tok::Op(Op::Pow)) {
            let right = self.percent()?;
            left = Expr::Binary(Op::Pow, Box::new(left), Box::new(right));
        }
        Ok(left)
    }

    fn percent(&mut self) -> Parsed<Expr> {
        let mut e = self.unary()?;
        while self.eat(&Tok::Percent) {
            e = Expr::Percent(Box::new(e));
        }
        Ok(e)
    }

    /// In Excel negation binds tighter than `^`: `-2^2` is 4
    fn unary(&mut self) -> Parsed<Expr> {
        if self.eat(&Tok::Op(Op::Sub)) {
            return Ok(Expr::Neg(Box::new(self.unary()?)));
        }
        if self.eat(&Tok::Op(Op::Add)) {
            return Ok(Expr::Plus(Box::new(self.unary()?)));
        }
        self.primary()
    }

    fn primary(&mut self) -> Parsed<Expr> {
        match self.next() {
            Some(Tok::Number(n)) => {
                if self.peek() == Some(&Tok::Colon) {
                    return unsupported("rangos de filas completas");
                }
                Ok(Expr::Number(n))
            }
            Some(Tok::Text(s)) => Ok(Expr::Text(s)),
            Some(Tok::Error(e)) => Ok(Expr::Error(e)),
            Some(Tok::Open) => {
                let inner = self.comparison()?;
                if matches!(self.peek(), Some(Tok::Comma)) {
                    return unsupported("el operador de unión");
                }
                if !self.eat(&Tok::Close) {
                    return invalid("Falta cerrar un paréntesis.");
                }
                Ok(Expr::Paren(Box::new(inner)))
            }
            Some(Tok::BraceOpen) => self.array(),
            Some(Tok::Sheet(sheet)) => match self.next() {
                Some(Tok::Word(word)) => self.reference(Some(sheet), &word),
                Some(Tok::Error(e)) if e == "#REF!" => Ok(Expr::Error(e)),
                _ => invalid(format!("Falta la referencia después de '{}!'.", sheet)),
            },
            Some(Tok::Word(word)) => {
                if self.peek() == Some(&Tok::Open) {
                    self.at += 1;
                    return self.call(&word);
                }
                let upper = word.to_uppercase();
                if (upper == "TRUE" || upper == "FALSE") && self.peek() != Some(&Tok::Colon) {
                    return Ok(Expr::Bool(upper == "TRUE"));
                }
                self.reference(None, &word)
            }
            Some(tok) => invalid(format!("No se esperaba {:?}.", tok)),
            None => invalid("La fórmula está incompleta."),
        }
    }

    fn reference(&mut self, sheet: Option<String>, word: &str) -> Parsed<Expr> {
        if let Some(first) = parse_pos(word) {
            if self.eat(&Tok::Colon) {
                let last = match self.next() {
                    Some(Tok::Word(w)) => parse_pos(&w),
                    _ => None,
                };
                let Some(last) = last else { return invalid(format!("Rango incompleto después de '{}:'.", word)) };
                return Ok(Expr::Ref(Area { sheet, first, last, whole_cols: false, is_range: true }));
            }
            return Ok(Expr::Ref(Area { sheet, first, last: first, whole_cols: false, is_range: false }));
        }
        if self.peek() == Some(&Tok::Colon) {
            if let Some((c1, a1)) = parse_col(word) {
                self.at += 1;
                let Some(Tok::Word(w)) = self.next() else { return invalid("Rango de columnas incompleto.") };
                let Some((c2, a2)) = parse_col(&w) else { return invalid("Rango de columnas incompleto.") };
                return Ok(Expr::Ref(Area {
                    sheet,
                    first: Pos { row: 1, col: c1, row_abs: true, col_abs: a1 },
                    last: Pos { row: MAX_ROW, col: c2, row_abs: true, col_abs: a2 },
                    whole_cols: true,
                    is_range: true,
                }));
            }
            return unsupported("rangos con nombres");
        }
        if sheet.is_some() {
            return unsupported("nombres definidos en otra hoja");
        }
        if word.contains('$') || word.chars().next().is_some_and(|c| c.is_ascii_digit()) {
            return invalid(format!("'{}' no es una referencia válida.", word));
        }
        Ok(Expr::Name(word.to_string()))
    }

    fn call(&mut self, word: &str) -> Parsed<Expr> {
        let upper = word.to_uppercase();
        let name = upper.trim_start_matches("_XLFN.").trim_start_matches("_XLWS.").to_string();
        let mut args = Vec::new();
        if self.eat(&Tok::Close) {
            return Ok(Expr::Call(name, args));
        }
        loop {
            let arg = if matches!(self.peek(), Some(Tok::Comma | Tok::Semicolon | Tok::Close)) {
                Expr::Missing
            } else {
                self.comparison()?
            };
            args.push(arg);
            match self.next() {
                Some(Tok::Comma | Tok::Semicolon) => continue,
                Some(Tok::Close) => break,
                _ => return invalid(format!("Faltan argumentos o un paréntesis en {}().", name)),
            }
        }
        Ok(Expr::Call(name, args))
    }

    fn array(&mut self) -> Parsed<Expr> {
        let mut rows = vec![Vec::new()];
        loop {
            let negative = self.eat(&Tok::Op(Op::Sub));
            let item = match self.next() {
                Some(Tok::Number(n)) => Expr::Number(if negative { -n } else { n }),
                Some(Tok::Text(s)) if !negative => Expr::Text(s),
                Some(Tok::Error(e)) if !negative => Expr::Error(e),
                Some(Tok::Word(w)) if !negative && matches!(w.to_uppercase().as_str(), "TRUE" | "FALSE") => {
                    Expr::Bool(w.eq_ignore_ascii_case("TRUE"))
                }
                _ => return invalid("Una matriz constante solo admite números, textos, booleanos y errores."),
            };
            rows.last_mut().unwrap().push(item);
            match self.next() {
                Some(Tok::Comma) => {}
                Some(Tok::Semicolon) => rows.push(Vec::new()),
                Some(Tok::BraceClose) => break,
                _ => return invalid("Falta cerrar una matriz constante."),
            }
        }
        if rows.iter().any(|r| r.len() != rows[0].len()) {
            return invalid("Las filas de la matriz constante tienen distinta longitud.");
        }
        Ok(Expr::Array(rows))
    }
}

/// Parses a formula as written in `<f>` (a leading `=` is accepted and ignored)
pub fn parse(text: &str) -> Parsed<Expr> {
    let text = text.trim();
    let text = text.strip_prefix('=').unwrap_or(text);
    if text.trim().is_empty() {
        return invalid("La fórmula está vacía.");
    }
    let mut parser = Parser { toks: tokenize(text)?, at: 0 };
    let expr = parser.comparison()?;
    match parser.peek() {
        None => Ok(expr),
        Some(Tok::Comma) => unsupported("el operador de unión"),
        Some(tok) => invalid(format!("Sobra {:?} al final de la fórmula.", tok)),
    }
}

impl Expr {
    /// The formula of a cell `rows` below and `cols` to the right of the one it was written
    /// for: relative references move, absolute ones stay; references pushed off the sheet
    /// become `#REF!`
    pub fn shifted(&self, rows: i64, cols: i64) -> Expr {
        let boxed = |e: &Expr| Box::new(e.shifted(rows, cols));
        match self {
            Expr::Ref(area) => {
                let shift = |p: Pos, whole_cols: bool| -> Option<Pos> {
                    let row = if p.row_abs || whole_cols { p.row as i64 } else { p.row as i64 + rows };
                    let col = if p.col_abs { p.col as i64 } else { p.col as i64 + cols };
                    (row >= 1 && row <= MAX_ROW as i64 && col >= 1 && col <= MAX_COL as i64)
                        .then_some(Pos { row: row as u32, col: col as u32, ..p })
                };
                match (shift(area.first, area.whole_cols), shift(area.last, area.whole_cols)) {
                    (Some(first), Some(last)) => Expr::Ref(Area { first, last, ..area.clone() }),
                    _ => Expr::Error("#REF!".to_string()),
                }
            }
            Expr::Neg(e) => Expr::Neg(boxed(e)),
            Expr::Plus(e) => Expr::Plus(boxed(e)),
            Expr::Percent(e) => Expr::Percent(boxed(e)),
            Expr::Paren(e) => Expr::Paren(boxed(e)),
            Expr::Binary(op, a, b) => Expr::Binary(*op, boxed(a), boxed(b)),
            Expr::Call(name, args) => Expr::Call(name.clone(), args.iter().map(|a| a.shifted(rows, cols)).collect()),
            other => other.clone(),
        }
    }

    /// The formula as Excel stores it: `_xlfn.` on newer functions, `,` between arguments
    pub fn to_formula(&self) -> String {
        let mut out = String::new();
        self.write(&mut out, true);
        out
    }

    /// The formula as people read it (no `_xlfn.` prefixes)
    pub fn to_display(&self) -> String {
        let mut out = String::new();
        self.write(&mut out, false);
        out
    }

    fn write(&self, out: &mut String, stored: bool) {
        match self {
            Expr::Number(n) => out.push_str(&number_literal(*n)),
            Expr::Text(s) => {
                out.push('"');
                out.push_str(&s.replace('"', "\"\""));
                out.push('"');
            }
            Expr::Bool(b) => out.push_str(if *b { "TRUE" } else { "FALSE" }),
            Expr::Error(e) => out.push_str(e),
            Expr::Ref(area) => write_area(area, out),
            Expr::Name(n) => out.push_str(n),
            Expr::Neg(e) => {
                out.push('-');
                e.write(out, stored);
            }
            Expr::Plus(e) => {
                out.push('+');
                e.write(out, stored);
            }
            Expr::Percent(e) => {
                e.write(out, stored);
                out.push('%');
            }
            Expr::Paren(e) => {
                out.push('(');
                e.write(out, stored);
                out.push(')');
            }
            Expr::Binary(op, a, b) => {
                a.write(out, stored);
                out.push_str(op.text());
                b.write(out, stored);
            }
            Expr::Call(name, args) => {
                if stored && FUTURE_FUNCTIONS.contains(&name.as_str()) {
                    out.push_str("_xlfn.");
                }
                out.push_str(name);
                out.push('(');
                for (i, arg) in args.iter().enumerate() {
                    if i > 0 {
                        out.push(',');
                    }
                    arg.write(out, stored);
                }
                out.push(')');
            }
            Expr::Missing => {}
            Expr::Array(rows) => {
                out.push('{');
                for (r, row) in rows.iter().enumerate() {
                    if r > 0 {
                        out.push(';');
                    }
                    for (c, item) in row.iter().enumerate() {
                        if c > 0 {
                            out.push(',');
                        }
                        item.write(out, stored);
                    }
                }
                out.push('}');
            }
        }
    }

    /// Calls `visit` with every reference and name in the formula
    pub fn visit_refs(&self, visit: &mut impl FnMut(&Expr)) {
        match self {
            Expr::Ref(_) | Expr::Name(_) => visit(self),
            Expr::Neg(e) | Expr::Plus(e) | Expr::Percent(e) | Expr::Paren(e) => e.visit_refs(visit),
            Expr::Binary(_, a, b) => {
                a.visit_refs(visit);
                b.visit_refs(visit);
            }
            Expr::Call(_, args) => args.iter().for_each(|a| a.visit_refs(visit)),
            _ => {}
        }
    }
}

fn number_literal(n: f64) -> String {
    if n.fract() == 0.0 && n.abs() < 1e15 {
        format!("{}", n as i64)
    } else {
        format!("{}", n)
    }
}

fn write_sheet(sheet: &str, out: &mut String) {
    let plain = !sheet.is_empty()
        && sheet.chars().all(|c| c.is_alphanumeric() || c == '_' || c == '.')
        && !sheet.chars().next().unwrap().is_ascii_digit()
        && parse_pos(sheet).is_none();
    if plain {
        out.push_str(sheet);
    } else {
        out.push('\'');
        out.push_str(&sheet.replace('\'', "''"));
        out.push('\'');
    }
    out.push('!');
}

fn write_area(area: &Area, out: &mut String) {
    if let Some(sheet) = &area.sheet {
        write_sheet(sheet, out);
    }
    let col = |p: Pos, out: &mut String| {
        if p.col_abs {
            out.push('$');
        }
        out.push_str(&column_name(p.col));
    };
    let pos = |p: Pos, out: &mut String| {
        col(p, out);
        if p.row_abs {
            out.push('$');
        }
        out.push_str(&p.row.to_string());
    };
    if area.whole_cols {
        col(area.first, out);
        out.push(':');
        col(area.last, out);
    } else {
        pos(area.first, out);
        if area.is_range {
            out.push(':');
            pos(area.last, out);
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn round_trip(text: &str) -> String {
        parse(text).unwrap().to_formula()
    }

    #[test]
    fn parses_and_prints_formulas() {
        for f in [
            "SUM(A1:B3)*2",
            "IF(A1>=10,\"alto\",\"bajo\")",
            "'Mis datos'!$B$2+Hoja2!C3",
            "-2^2",
            "A1&\" \"&B1",
            "SUMIF(A:A,\">5\",B:B)",
            "_xlfn.XLOOKUP(A2,D:D,E:E,\"\")",
            "{1,2;3,4}",
            "IF(A1,,2)",
            "(1+2)*3%",
            "Total*IVA",
            "#N/A",
        ] {
            assert_eq!(round_trip(f), f, "{}", f);
        }
        assert_eq!(round_trip("=sum( a1 ; b2 )"), "SUM(a1,b2)".replace("a1", "A1").replace("b2", "B2"));
        assert_eq!(round_trip("XLOOKUP(1,A:A,B:B)"), "_xlfn.XLOOKUP(1,A:A,B:B)");
        assert_eq!(parse("_xlfn.CONCAT(A1)").unwrap().to_display(), "CONCAT(A1)");
    }

    #[test]
    fn precedence_follows_excel() {
        let Expr::Binary(Op::Pow, left, _) = parse("-2^2").unwrap() else { panic!() };
        assert!(matches!(*left, Expr::Neg(_)));
        let Expr::Binary(Op::Add, _, right) = parse("1+2*3").unwrap() else { panic!() };
        assert!(matches!(*right, Expr::Binary(Op::Mul, _, _)));
        let Expr::Binary(Op::Eq, _, _) = parse("A1&\"x\"=\"1x\"").unwrap() else { panic!() };
    }

    #[test]
    fn reports_unsupported_and_invalid_formulas() {
        for f in ["[1]Hoja1!A1", "Hoja1:Hoja3!A1", "Tabla1[Importe]", "SUM((A1,B1))", "1:1"] {
            assert!(matches!(parse(f), Err(ParseError::Unsupported(_))), "{}", f);
        }
        for f in ["", "SUM(A1", "1+", "\"abc", "A1 B1", "$$A1"] {
            assert!(matches!(parse(f), Err(ParseError::Invalid(_))), "{}", f);
        }
    }

    #[test]
    fn shifts_relative_references_only() {
        let f = parse("B2*$C$1+SUM(D$1:D2)+A:A").unwrap();
        assert_eq!(f.shifted(3, 1).to_formula(), "C5*$C$1+SUM(E$1:E5)+B:B");
        assert_eq!(parse("A1").unwrap().shifted(-1, 0).to_formula(), "#REF!");
    }
}
