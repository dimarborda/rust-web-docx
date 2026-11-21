//! Evaluates formulas the way Excel does for the common cases: arithmetic, text and comparison
//! operators with Excel's coercions, references and ranges across sheets, defined names, and
//! about ninety frequently used functions (math, logic, conditional sums and counts, lookups,
//! text and dates).
//!
//! Anything else (an unknown function, a result that would spill over several cells) is
//! `Unsupported`: the caller then keeps the value Excel cached for that formula.

use std::cmp::Ordering;

use crate::cell_ref::CellRef;
use crate::dates::{serial_to_unix_ms, unix_ms_to_serial};
use crate::formula::{Area, Expr, Op};
use crate::numfmt::{self, Locale};

/// A value as formulas see it; dates are serial numbers
#[derive(Debug, Clone, PartialEq)]
pub enum Val {
    Empty,
    Number(f64),
    Text(String),
    Bool(bool),
    /// `#DIV/0!`, `#N/A`…
    Error(String),
}

/// Why evaluation stopped: an Excel error value, or something this engine does not evaluate
#[derive(Debug, Clone, PartialEq)]
pub enum Flow {
    Error(String),
    Unsupported(String),
}

type R<T> = Result<T, Flow>;

fn fail<T>(code: &str) -> R<T> {
    Err(Flow::Error(code.to_string()))
}

fn unsupported<T>(what: impl Into<String>) -> R<T> {
    Err(Flow::Unsupported(what.into()))
}

const DIV0: &str = "#DIV/0!";
const VALUE: &str = "#VALUE!";
const NA: &str = "#N/A";
const NUM: &str = "#NUM!";
const REF: &str = "#REF!";
const NAME: &str = "#NAME?";

/// Characters 0x80-0x9F of Windows-1252, the code page CHAR and CODE use (the rest of
/// 1-255 matches Latin-1); the five codes Windows leaves undefined map to themselves
const CP1252_HIGH: [char; 32] = [
    '€', '\u{81}', '‚', 'ƒ', '„', '…', '†', '‡', 'ˆ', '‰', 'Š', '‹', 'Œ', '\u{8D}', 'Ž', '\u{8F}',
    '\u{90}', '‘', '’', '“', '”', '•', '–', '—', '˜', '™', 'š', '›', 'œ', '\u{9D}', 'ž', 'Ÿ',
];

/// Most cells an array operation may materialize
const MAX_ARRAY_CELLS: u64 = 1_000_000;

/// What formulas read: cell values (already calculated ones for formula cells), sheets and names
pub trait Context {
    fn value(&self, sheet: usize, at: CellRef) -> Val;
    fn sheet_index(&self, name: &str) -> Option<usize>;
    /// Definition of the defined name `name` as seen from `sheet`, and the sheet its
    /// references without a sheet name point to
    fn name(&self, name: &str, sheet: usize) -> Option<(&Expr, usize)>;
    /// Last used row and column of a sheet (0 when empty)
    fn bounds(&self, sheet: usize) -> (u32, u32);
    fn date1904(&self) -> bool;
    /// The current local date and time as a serial number
    fn now(&self) -> f64;
}

/// A rectangle of cells on one sheet
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct RangeRef {
    pub sheet: usize,
    pub first: CellRef,
    pub last: CellRef,
}

impl RangeRef {
    fn rows(&self) -> usize {
        (self.last.row - self.first.row + 1) as usize
    }

    fn cols(&self) -> usize {
        (self.last.col - self.first.col + 1) as usize
    }
}

/// The result of evaluating part of a formula
#[derive(Debug, Clone, PartialEq)]
pub enum Res {
    Val(Val),
    Range(RangeRef),
    Array(Vec<Vec<Val>>),
}

/// A range or array read cell by cell
struct Grid<'c> {
    ctx: &'c dyn Context,
    kind: GridKind,
}

enum GridKind {
    Range(RangeRef, u32, u32),
    Array(Vec<Vec<Val>>),
}

impl<'c> Grid<'c> {
    fn rows(&self) -> usize {
        match &self.kind {
            GridKind::Range(r, _, _) => r.rows(),
            GridKind::Array(a) => a.len(),
        }
    }

    fn cols(&self) -> usize {
        match &self.kind {
            GridKind::Range(r, _, _) => r.cols(),
            GridKind::Array(a) => a.first().map(Vec::len).unwrap_or(0),
        }
    }

    /// Rows and columns that may hold values (a range is cut at the sheet's used area)
    fn used(&self) -> (usize, usize) {
        match &self.kind {
            GridKind::Range(r, max_row, max_col) => (
                (max_row.saturating_sub(r.first.row - 1) as usize).min(r.rows()),
                (max_col.saturating_sub(r.first.col - 1) as usize).min(r.cols()),
            ),
            GridKind::Array(_) => (self.rows(), self.cols()),
        }
    }

    fn get(&self, row: usize, col: usize) -> Val {
        match &self.kind {
            GridKind::Range(r, max_row, max_col) => {
                let at = CellRef {
                    row: r.first.row + row as u32,
                    col: r.first.col + col as u32,
                };
                if at.row > *max_row || at.col > *max_col || row >= r.rows() || col >= r.cols() {
                    Val::Empty
                } else {
                    self.ctx.value(r.sheet, at)
                }
            }
            GridKind::Array(a) => a
                .get(row)
                .and_then(|r| r.get(col))
                .cloned()
                .unwrap_or(Val::Empty),
        }
    }

    /// Values of the used area, by row
    fn values(&self) -> impl Iterator<Item = Val> + '_ {
        let (rows, cols) = self.used();
        (0..rows).flat_map(move |r| (0..cols).map(move |c| self.get(r, c)))
    }

    /// Range or array as a vector (one row or one column), else `None`
    fn vector_len(&self) -> Option<usize> {
        match (self.rows(), self.cols()) {
            (1, n) | (n, 1) => Some(n),
            _ => None,
        }
    }

    fn vector_get(&self, i: usize) -> Val {
        if self.rows() == 1 {
            self.get(0, i)
        } else {
            self.get(i, 0)
        }
    }
}

/// Evaluation of one formula, written for the cell `at` of `sheet`
pub struct Eval<'c> {
    ctx: &'c dyn Context,
    sheet: usize,
    at: CellRef,
    depth: usize,
}

/// Evaluates the formula of a cell. `Err` means it uses something not supported here.
pub fn evaluate(ctx: &dyn Context, sheet: usize, at: CellRef, expr: &Expr) -> Result<Val, String> {
    let mut eval = Eval {
        ctx,
        sheet,
        at,
        depth: 0,
    };
    let result = eval.eval(expr).and_then(|res| match res {
        Res::Array(rows) if rows.len() > 1 || rows.first().is_some_and(|r| r.len() > 1) => {
            unsupported("una matriz que ocupa varias celdas")
        }
        other => eval.scalar(other),
    });
    match result {
        Ok(Val::Empty) => Ok(Val::Number(0.0)),
        Ok(v) => Ok(v),
        Err(Flow::Error(e)) => Ok(Val::Error(e)),
        Err(Flow::Unsupported(what)) => Err(what),
    }
}

// ---------- Coercions ----------

/// Text to number as Excel does when text meets arithmetic: `" 12 "`, `"1e3"`, `"50%"`
fn parse_number(text: &str) -> Option<f64> {
    let t = text.trim();
    if t.is_empty() {
        return None;
    }
    if let Some(p) = t.strip_suffix('%') {
        return p
            .trim()
            .parse::<f64>()
            .ok()
            .filter(|n| n.is_finite())
            .map(|n| n / 100.0);
    }
    t.parse::<f64>().ok().filter(|n| n.is_finite())
}

fn to_num(v: &Val) -> R<f64> {
    match v {
        Val::Number(n) => Ok(*n),
        Val::Bool(b) => Ok(if *b { 1.0 } else { 0.0 }),
        Val::Empty => Ok(0.0),
        Val::Text(t) => parse_number(t).map(Ok).unwrap_or_else(|| fail(VALUE)),
        Val::Error(e) => fail(e),
    }
}

/// A number as text when joined with `&` (Excel's General, up to 15 significant digits)
pub fn number_text(n: f64) -> String {
    if n == 0.0 {
        return "0".to_string();
    }
    let a = n.abs();
    if !(1e-9..1e15).contains(&a) {
        let s = format!("{:.14E}", n);
        let (mantissa, exp) = s.split_once('E').unwrap();
        let mantissa = mantissa.trim_end_matches('0').trim_end_matches('.');
        let exp: i32 = exp.parse().unwrap_or(0);
        return format!(
            "{}E{}{:02}",
            mantissa,
            if exp < 0 { '-' } else { '+' },
            exp.abs()
        );
    }
    let rounded = round_sig(n, 15);
    let s = format!("{}", rounded);
    s
}

fn to_text(v: &Val) -> R<String> {
    match v {
        Val::Number(n) => Ok(number_text(*n)),
        Val::Text(t) => Ok(t.clone()),
        Val::Bool(b) => Ok(if *b { "TRUE" } else { "FALSE" }.to_string()),
        Val::Empty => Ok(String::new()),
        Val::Error(e) => fail(e),
    }
}

fn to_bool(v: &Val) -> R<bool> {
    match v {
        Val::Bool(b) => Ok(*b),
        Val::Number(n) => Ok(*n != 0.0),
        Val::Empty => Ok(false),
        Val::Text(t) if t.eq_ignore_ascii_case("TRUE") => Ok(true),
        Val::Text(t) if t.eq_ignore_ascii_case("FALSE") => Ok(false),
        Val::Text(_) => fail(VALUE),
        Val::Error(e) => fail(e),
    }
}

/// Rounds to `digits` significant digits (Excel keeps 15)
fn round_sig(n: f64, digits: usize) -> f64 {
    if n == 0.0 || !n.is_finite() {
        return n;
    }
    format!("{:.*e}", digits - 1, n).parse().unwrap_or(n)
}

fn finite(n: f64) -> R<Val> {
    if n.is_finite() {
        Ok(Val::Number(n))
    } else {
        fail(NUM)
    }
}

/// Excel's order: numbers < text < booleans; text without case; empty is 0, "" or FALSE
fn compare(a: &Val, b: &Val) -> Ordering {
    fn rank(v: &Val) -> u8 {
        match v {
            Val::Number(_) | Val::Empty => 0,
            Val::Text(_) => 1,
            Val::Bool(_) => 2,
            Val::Error(_) => 3,
        }
    }
    match (a, b) {
        (Val::Empty, Val::Empty) => Ordering::Equal,
        (Val::Empty, Val::Text(t)) => "".cmp(t.to_lowercase().as_str()),
        (Val::Text(t), Val::Empty) => t.to_lowercase().as_str().cmp(""),
        (Val::Empty, Val::Bool(b)) => false.cmp(b),
        (Val::Bool(b), Val::Empty) => b.cmp(&false),
        (Val::Number(x), Val::Empty) => x.partial_cmp(&0.0).unwrap_or(Ordering::Equal),
        (Val::Empty, Val::Number(y)) => 0.0.partial_cmp(y).unwrap_or(Ordering::Equal),
        (Val::Number(x), Val::Number(y)) => x.partial_cmp(y).unwrap_or(Ordering::Equal),
        (Val::Text(x), Val::Text(y)) => x.to_lowercase().cmp(&y.to_lowercase()),
        (Val::Bool(x), Val::Bool(y)) => x.cmp(y),
        _ => rank(a).cmp(&rank(b)),
    }
}

fn binary(op: Op, a: &Val, b: &Val) -> Val {
    let result = (|| -> R<Val> {
        match op {
            Op::Concat => Ok(Val::Text(to_text(a)? + &to_text(b)?)),
            Op::Eq | Op::Ne | Op::Lt | Op::Le | Op::Gt | Op::Ge => {
                if let Val::Error(e) = a {
                    return fail(e);
                }
                if let Val::Error(e) = b {
                    return fail(e);
                }
                let ord = compare(a, b);
                Ok(Val::Bool(match op {
                    Op::Eq => ord == Ordering::Equal,
                    Op::Ne => ord != Ordering::Equal,
                    Op::Lt => ord == Ordering::Less,
                    Op::Le => ord != Ordering::Greater,
                    Op::Gt => ord == Ordering::Greater,
                    _ => ord != Ordering::Less,
                }))
            }
            _ => {
                let (x, y) = (to_num(a)?, to_num(b)?);
                match op {
                    Op::Add => finite(x + y),
                    Op::Sub => finite(x - y),
                    Op::Mul => finite(x * y),
                    Op::Div if y == 0.0 => fail(DIV0),
                    Op::Div => finite(x / y),
                    _ if x == 0.0 && y < 0.0 => fail(DIV0),
                    _ => finite(x.powf(y)),
                }
            }
        }
    })();
    result.unwrap_or_else(|f| match f {
        Flow::Error(e) => Val::Error(e),
        Flow::Unsupported(_) => Val::Error(VALUE.to_string()),
    })
}

// ---------- Criteria (SUMIF, COUNTIF…) ----------

enum Criterion {
    Number(Op, f64),
    Text(Op, String),
    Bool(Op, bool),
    /// `""` / `"="`: blank cells; `"<>"`: non-blank ones
    Blank(bool),
    Error(Op, String),
}

impl Criterion {
    fn parse(v: &Val) -> Criterion {
        match v {
            Val::Number(n) => Criterion::Number(Op::Eq, *n),
            Val::Bool(b) => Criterion::Bool(Op::Eq, *b),
            Val::Empty => Criterion::Number(Op::Eq, 0.0),
            Val::Error(e) => Criterion::Error(Op::Eq, e.clone()),
            Val::Text(t) => {
                let (op, rest) = [
                    (">=", Op::Ge),
                    ("<=", Op::Le),
                    ("<>", Op::Ne),
                    (">", Op::Gt),
                    ("<", Op::Lt),
                    ("=", Op::Eq),
                ]
                .iter()
                .find_map(|(p, op)| t.strip_prefix(p).map(|rest| (*op, rest)))
                .unwrap_or((Op::Eq, t.as_str()));
                if rest.is_empty() {
                    return Criterion::Blank(op != Op::Ne);
                }
                if let Some(n) = parse_number(rest) {
                    return Criterion::Number(op, n);
                }
                if rest.eq_ignore_ascii_case("TRUE") || rest.eq_ignore_ascii_case("FALSE") {
                    return Criterion::Bool(op, rest.eq_ignore_ascii_case("TRUE"));
                }
                if rest.starts_with('#') {
                    return Criterion::Error(op, rest.to_uppercase());
                }
                Criterion::Text(op, rest.to_string())
            }
        }
    }

    fn holds(op: Op, ord: Ordering) -> bool {
        match op {
            Op::Eq => ord == Ordering::Equal,
            Op::Ne => ord != Ordering::Equal,
            Op::Lt => ord == Ordering::Less,
            Op::Le => ord != Ordering::Greater,
            Op::Gt => ord == Ordering::Greater,
            _ => ord != Ordering::Less,
        }
    }

    fn matches(&self, v: &Val) -> bool {
        match self {
            Criterion::Blank(true) => {
                matches!(v, Val::Empty) || matches!(v, Val::Text(t) if t.is_empty())
            }
            Criterion::Blank(false) => {
                !(matches!(v, Val::Empty) || matches!(v, Val::Text(t) if t.is_empty()))
            }
            Criterion::Number(op, n) => match v {
                Val::Number(x) => Self::holds(*op, x.partial_cmp(n).unwrap_or(Ordering::Equal)),
                Val::Text(t) if *op == Op::Eq => {
                    parse_number(t) == Some(*n) && !t.trim().is_empty()
                }
                _ => *op == Op::Ne,
            },
            Criterion::Bool(op, b) => match v {
                Val::Bool(x) => Self::holds(*op, x.cmp(b)),
                _ => *op == Op::Ne,
            },
            Criterion::Error(op, e) => match v {
                Val::Error(x) => (x == e) == (*op != Op::Ne),
                _ => *op == Op::Ne,
            },
            Criterion::Text(op, pattern) => match (op, v) {
                (Op::Eq, Val::Text(t)) => wildcard_match(pattern, t),
                (Op::Eq, _) => false,
                (Op::Ne, Val::Text(t)) => !wildcard_match(pattern, t),
                (Op::Ne, _) => true,
                (op, Val::Text(t)) => {
                    Self::holds(*op, t.to_lowercase().cmp(&pattern.to_lowercase()))
                }
                _ => false,
            },
        }
    }
}

/// `*` any text, `?` one character, `~` escapes them; without case
fn wildcard_match(pattern: &str, text: &str) -> bool {
    let p: Vec<char> = pattern.to_lowercase().chars().collect();
    let t: Vec<char> = text.to_lowercase().chars().collect();
    // Tokens: Some(c) a literal, None for `*`, Some('\0') never used; `?` kept apart
    enum P {
        Lit(char),
        One,
        Any,
    }
    let mut tokens = Vec::new();
    let mut i = 0;
    while i < p.len() {
        match p[i] {
            '~' if i + 1 < p.len() => {
                tokens.push(P::Lit(p[i + 1]));
                i += 2;
                continue;
            }
            '*' => tokens.push(P::Any),
            '?' => tokens.push(P::One),
            c => tokens.push(P::Lit(c)),
        }
        i += 1;
    }
    let (mut ti, mut pi) = (0usize, 0usize);
    let mut star: Option<(usize, usize)> = None;
    while ti < t.len() {
        match tokens.get(pi) {
            Some(P::Lit(c)) if *c == t[ti] => {
                ti += 1;
                pi += 1;
            }
            Some(P::One) => {
                ti += 1;
                pi += 1;
            }
            Some(P::Any) => {
                star = Some((pi, ti));
                pi += 1;
            }
            _ => match star {
                Some((sp, st)) => {
                    pi = sp + 1;
                    ti = st + 1;
                    star = Some((sp, st + 1));
                }
                None => return false,
            },
        }
    }
    tokens[pi..].iter().all(|t| matches!(t, P::Any))
}

fn has_wildcards(text: &str) -> bool {
    text.contains('*') || text.contains('?')
}

/// Equality used by lookups: text without case (with wildcards when the lookup has them)
fn lookup_equal(lookup: &Val, v: &Val, wildcards: bool) -> bool {
    match (lookup, v) {
        (Val::Text(p), Val::Text(t)) if wildcards && has_wildcards(p) => wildcard_match(p, t),
        (Val::Text(a), Val::Text(b)) => a.to_lowercase() == b.to_lowercase(),
        (Val::Number(a), Val::Number(b)) => a == b,
        (Val::Bool(a), Val::Bool(b)) => a == b,
        _ => false,
    }
}

fn same_kind(a: &Val, b: &Val) -> bool {
    matches!(
        (a, b),
        (Val::Number(_), Val::Number(_))
            | (Val::Text(_), Val::Text(_))
            | (Val::Bool(_), Val::Bool(_))
    )
}

// ---------- Dates ----------

const MS_PER_DAY: f64 = 86_400_000.0;

/// Days since 1970-01-01 of a proleptic Gregorian date
fn days_from_civil(y: i64, m: i64, d: i64) -> i64 {
    let y = if m <= 2 { y - 1 } else { y };
    let era = y.div_euclid(400);
    let yoe = y - era * 400;
    let mp = (m + 9) % 12;
    let doy = (153 * mp + 2) / 5 + d - 1;
    let doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
    era * 146_097 + doe - 719_468
}

fn civil_from_days(z: i64) -> (i64, i64, i64) {
    let z = z + 719_468;
    let era = z.div_euclid(146_097);
    let doe = z - era * 146_097;
    let yoe = (doe - doe / 1460 + doe / 36_524 - doe / 146_096) / 365;
    let y = yoe + era * 400;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let d = doy - (153 * mp + 2) / 5 + 1;
    let m = if mp < 10 { mp + 3 } else { mp - 9 };
    (if m <= 2 { y + 1 } else { y }, m, d)
}

impl<'c> Eval<'c> {
    fn ymd(&self, serial: f64) -> R<(i64, i64, i64)> {
        if serial < 0.0 {
            return fail(NUM);
        }
        if !self.ctx.date1904() && serial.floor() == 60.0 {
            return Ok((1900, 2, 29));
        }
        let days =
            (serial_to_unix_ms(serial.floor(), self.ctx.date1904()) / MS_PER_DAY).floor() as i64;
        Ok(civil_from_days(days))
    }

    /// Day of the week of a serial, 0 = Sunday
    fn weekday(&self, serial: f64) -> R<i64> {
        let (y, m, d) = self.ymd(serial)?;
        Ok((days_from_civil(y, m, d) + 4).rem_euclid(7))
    }

    /// Monday to Friday, the working days of NETWORKDAYS and WORKDAY
    fn is_weekday(&self, serial: f64) -> R<bool> {
        Ok(!matches!(self.weekday(serial)?, 0 | 6))
    }

    /// Serial of a date; months and days out of range roll over as in Excel's DATE
    fn serial(&self, y: i64, m: i64, d: i64) -> R<f64> {
        let y = if (0..1900).contains(&y) { y + 1900 } else { y };
        if !(0..10_000).contains(&y) {
            return fail(NUM);
        }
        let months = y * 12 + (m - 1);
        let (y, m) = (months.div_euclid(12), months.rem_euclid(12) + 1);
        let days = days_from_civil(y, m, 1) + d - 1;
        unix_ms_to_serial(days as f64 * MS_PER_DAY, self.ctx.date1904()).or_else(|_| fail(NUM))
    }
}

// ---------- Evaluation ----------

impl<'c> Eval<'c> {
    fn grid(&self, res: Res) -> R<Grid<'c>> {
        Ok(Grid {
            ctx: self.ctx,
            kind: match res {
                Res::Range(r) => {
                    let (max_row, max_col) = self.ctx.bounds(r.sheet);
                    GridKind::Range(r, max_row, max_col)
                }
                Res::Array(a) => GridKind::Array(a),
                Res::Val(Val::Error(e)) => return fail(&e),
                Res::Val(v) => GridKind::Array(vec![vec![v]]),
            },
        })
    }

    /// A single value: ranges by implicit intersection with the formula's row or column
    fn scalar(&self, res: Res) -> R<Val> {
        let v = match res {
            Res::Val(v) => v,
            Res::Array(a) => a
                .first()
                .and_then(|r| r.first())
                .cloned()
                .unwrap_or(Val::Empty),
            Res::Range(r) => {
                let at = if r.rows() == 1 && r.cols() == 1 {
                    r.first
                } else if r.cols() == 1
                    && r.sheet == self.sheet
                    && (r.first.row..=r.last.row).contains(&self.at.row)
                {
                    CellRef {
                        row: self.at.row,
                        col: r.first.col,
                    }
                } else if r.rows() == 1
                    && r.sheet == self.sheet
                    && (r.first.col..=r.last.col).contains(&self.at.col)
                {
                    CellRef {
                        row: r.first.row,
                        col: self.at.col,
                    }
                } else {
                    return fail(VALUE);
                };
                self.ctx.value(r.sheet, at)
            }
        };
        match v {
            Val::Error(e) => fail(&e),
            v => Ok(v),
        }
    }

    fn area(&self, area: &Area) -> R<RangeRef> {
        let sheet = match &area.sheet {
            Some(name) => self
                .ctx
                .sheet_index(name)
                .map(Ok)
                .unwrap_or_else(|| fail(REF))?,
            None => self.sheet,
        };
        let first = CellRef {
            row: area.first.row.min(area.last.row),
            col: area.first.col.min(area.last.col),
        };
        let last = CellRef {
            row: area.first.row.max(area.last.row),
            col: area.first.col.max(area.last.col),
        };
        Ok(RangeRef { sheet, first, last })
    }

    fn eval(&mut self, expr: &Expr) -> R<Res> {
        Ok(match expr {
            Expr::Number(n) => Res::Val(Val::Number(*n)),
            Expr::Text(s) => Res::Val(Val::Text(s.clone())),
            Expr::Bool(b) => Res::Val(Val::Bool(*b)),
            Expr::Error(e) => return fail(e),
            Expr::Missing => Res::Val(Val::Empty),
            Expr::Ref(area) => Res::Range(self.area(area)?),
            Expr::Name(name) => {
                if self.depth > 16 {
                    return unsupported("nombres definidos anidados");
                }
                let Some((definition, sheet)) = self.ctx.name(name, self.sheet) else {
                    return fail(NAME);
                };
                let mut inner = Eval {
                    ctx: self.ctx,
                    sheet,
                    at: self.at,
                    depth: self.depth + 1,
                };
                inner.eval(definition)?
            }
            Expr::Paren(e) | Expr::Plus(e) => self.eval(e)?,
            Expr::Neg(e) => {
                let res = self.eval(e)?;
                self.map(res, |v| binary(Op::Sub, &Val::Number(0.0), v))?
            }
            Expr::Percent(e) => {
                let res = self.eval(e)?;
                self.map(res, |v| binary(Op::Div, v, &Val::Number(100.0)))?
            }
            Expr::Binary(op, a, b) => {
                let (a, b) = (self.eval(a)?, self.eval(b)?);
                self.combine(*op, a, b)?
            }
            Expr::Array(rows) => {
                let mut out = Vec::with_capacity(rows.len());
                for row in rows {
                    let mut values = Vec::with_capacity(row.len());
                    for item in row {
                        values.push(match self.eval(item) {
                            Ok(res) => self
                                .scalar(res)
                                .unwrap_or_else(|_| Val::Error(VALUE.into())),
                            Err(Flow::Error(e)) => Val::Error(e),
                            Err(other) => return Err(other),
                        });
                    }
                    out.push(values);
                }
                Res::Array(out)
            }
            Expr::Call(name, args) => self.call(name, args)?,
        })
    }

    fn matrix(&self, res: Res) -> R<Vec<Vec<Val>>> {
        let grid = self.grid(res)?;
        if grid.rows() as u64 * grid.cols() as u64 > MAX_ARRAY_CELLS {
            return unsupported("operaciones con rangos demasiado grandes");
        }
        Ok((0..grid.rows())
            .map(|r| (0..grid.cols()).map(|c| grid.get(r, c)).collect())
            .collect())
    }

    fn map(&self, res: Res, f: impl Fn(&Val) -> Val) -> R<Res> {
        match res {
            Res::Val(v) => match f(&v) {
                Val::Error(e) => fail(&e),
                v => Ok(Res::Val(v)),
            },
            other => {
                let m = self.matrix(other)?;
                Ok(Res::Array(
                    m.iter().map(|r| r.iter().map(&f).collect()).collect(),
                ))
            }
        }
    }

    /// An operator: on two values, or element by element when either side is a range or array
    fn combine(&self, op: Op, a: Res, b: Res) -> R<Res> {
        if let (Res::Val(x), Res::Val(y)) = (&a, &b) {
            return match binary(op, x, y) {
                Val::Error(e) => fail(&e),
                v => Ok(Res::Val(v)),
            };
        }
        let single = |r: &Res| matches!(r, Res::Range(rr) if rr.rows() == 1 && rr.cols() == 1);
        if single(&a) && single(&b)
            || single(&a) && matches!(b, Res::Val(_))
            || matches!(a, Res::Val(_)) && single(&b)
        {
            let (x, y) = (self.scalar(a)?, self.scalar(b)?);
            return match binary(op, &x, &y) {
                Val::Error(e) => fail(&e),
                v => Ok(Res::Val(v)),
            };
        }
        let (ma, mb) = (self.matrix(a)?, self.matrix(b)?);
        let rows = ma.len().max(mb.len());
        let cols = ma
            .first()
            .map(Vec::len)
            .unwrap_or(0)
            .max(mb.first().map(Vec::len).unwrap_or(0));
        let pick = |m: &Vec<Vec<Val>>, r: usize, c: usize| -> Val {
            let r = if m.len() == 1 { 0 } else { r };
            let c = if m.first().map(Vec::len) == Some(1) {
                0
            } else {
                c
            };
            m.get(r)
                .and_then(|row| row.get(c))
                .cloned()
                .unwrap_or(Val::Error(NA.into()))
        };
        Ok(Res::Array(
            (0..rows)
                .map(|r| {
                    (0..cols)
                        .map(|c| binary(op, &pick(&ma, r, c), &pick(&mb, r, c)))
                        .collect()
                })
                .collect(),
        ))
    }

    // ----- Argument helpers -----

    fn val(&mut self, e: &Expr) -> R<Val> {
        let res = self.eval(e)?;
        self.scalar(res)
    }

    fn num(&mut self, e: &Expr) -> R<f64> {
        to_num(&self.val(e)?)
    }

    fn text(&mut self, e: &Expr) -> R<String> {
        to_text(&self.val(e)?)
    }

    fn boolean(&mut self, e: &Expr) -> R<bool> {
        to_bool(&self.val(e)?)
    }

    fn opt_num(&mut self, args: &[Expr], i: usize, default: f64) -> R<f64> {
        match args.get(i) {
            None | Some(Expr::Missing) => Ok(default),
            Some(e) => self.num(e),
        }
    }

    /// Day serials of an optional holidays argument (a range, an array or one date)
    fn holidays(&mut self, args: &[Expr], i: usize) -> R<std::collections::HashSet<i64>> {
        match args.get(i) {
            None | Some(Expr::Missing) => Ok(Default::default()),
            Some(_) => Ok(self
                .numbers(&args[i..=i])?
                .into_iter()
                .map(|x| x.floor() as i64)
                .collect()),
        }
    }

    fn arg_grid(&mut self, e: &Expr) -> R<Grid<'c>> {
        let res = self.eval(e)?;
        self.grid(res)
    }

    /// Numbers of the arguments as SUM reads them: in ranges only numbers count (errors
    /// propagate); values typed as arguments are converted
    fn numbers(&mut self, args: &[Expr]) -> R<Vec<f64>> {
        let mut out = Vec::new();
        for arg in args {
            match self.eval(arg)? {
                Res::Val(v) => match v {
                    Val::Empty => {}
                    v => out.push(to_num(&v)?),
                },
                other => {
                    for v in self.grid(other)?.values() {
                        match v {
                            Val::Number(n) => out.push(n),
                            Val::Error(e) => return fail(&e),
                            _ => {}
                        }
                    }
                }
            }
        }
        Ok(out)
    }

    /// Every value of the arguments, ranges flattened by row
    fn all_values(&mut self, args: &[Expr]) -> R<Vec<(Val, bool)>> {
        let mut out = Vec::new();
        for arg in args {
            match self.eval(arg)? {
                Res::Val(v) => out.push((v, false)),
                other => out.extend(self.grid(other)?.values().map(|v| (v, true))),
            }
        }
        Ok(out)
    }

    fn arity(name: &str, args: &[Expr], min: usize, max: usize) -> R<()> {
        if args.len() < min || args.len() > max {
            return unsupported(format!("{} con {} argumentos", name, args.len()));
        }
        Ok(())
    }

    // ----- Functions -----

    fn call(&mut self, name: &str, args: &[Expr]) -> R<Res> {
        let v = |v: Val| Ok(Res::Val(v));
        let n = |x: f64| finite(x).map(Res::Val);
        match name {
            // Math
            "SUM" => n(self.numbers(args)?.iter().sum()),
            "PRODUCT" => n(self.numbers(args)?.iter().product()),
            "AVERAGE" => {
                let nums = self.numbers(args)?;
                if nums.is_empty() {
                    return fail(DIV0);
                }
                n(nums.iter().sum::<f64>() / nums.len() as f64)
            }
            "MIN" | "MAX" => {
                let nums = self.numbers(args)?;
                let pick = if name == "MIN" { f64::min } else { f64::max };
                n(nums.iter().copied().reduce(pick).unwrap_or(0.0))
            }
            "MEDIAN" => {
                let mut nums = self.numbers(args)?;
                if nums.is_empty() {
                    return fail(NUM);
                }
                nums.sort_by(|a, b| a.partial_cmp(b).unwrap_or(Ordering::Equal));
                let mid = nums.len() / 2;
                n(if nums.len() % 2 == 0 {
                    (nums[mid - 1] + nums[mid]) / 2.0
                } else {
                    nums[mid]
                })
            }
            "COUNT" => {
                let values = self.all_values(args)?;
                let count = values.iter().filter(|(v, from_range)| match v {
                    Val::Number(_) => true,
                    Val::Bool(_) => !from_range,
                    Val::Text(t) => !from_range && parse_number(t).is_some(),
                    _ => false,
                });
                v(Val::Number(count.count() as f64))
            }
            "COUNTA" => {
                let count = self
                    .all_values(args)?
                    .iter()
                    .filter(|(v, from_range)| !(matches!(v, Val::Empty) && *from_range))
                    .count();
                v(Val::Number(count as f64))
            }
            "COUNTBLANK" => {
                Self::arity(name, args, 1, 1)?;
                let grid = self.arg_grid(&args[0])?;
                let total = grid.rows() as f64 * grid.cols() as f64;
                let filled = grid
                    .values()
                    .filter(|v| {
                        !(matches!(v, Val::Empty) || matches!(v, Val::Text(t) if t.is_empty()))
                    })
                    .count();
                v(Val::Number(total - filled as f64))
            }
            "ROUND" | "ROUNDUP" | "ROUNDDOWN" => {
                Self::arity(name, args, 2, 2)?;
                let (x, digits) = (self.num(&args[0])?, self.num(&args[1])?.trunc());
                let f = 10f64.powi(digits as i32);
                let scaled = round_sig(round_sig(x, 15) * f, 15);
                let r = match name {
                    "ROUND" => scaled.round(),
                    "ROUNDUP" => scaled.abs().ceil().copysign(scaled),
                    _ => scaled.trunc(),
                };
                n(round_sig(r / f, 15))
            }
            "INT" => n(self.num(args.first().unwrap_or(&Expr::Missing))?.floor()),
            "TRUNC" => {
                let x = self.num(args.first().unwrap_or(&Expr::Missing))?;
                let f = 10f64.powi(self.opt_num(args, 1, 0.0)?.trunc() as i32);
                n((x * f).trunc() / f)
            }
            "ABS" => n(self.num(args.first().unwrap_or(&Expr::Missing))?.abs()),
            "SIGN" => {
                let x = self.num(args.first().unwrap_or(&Expr::Missing))?;
                n(if x > 0.0 {
                    1.0
                } else if x < 0.0 {
                    -1.0
                } else {
                    0.0
                })
            }
            "SQRT" => {
                let x = self.num(args.first().unwrap_or(&Expr::Missing))?;
                if x < 0.0 {
                    return fail(NUM);
                }
                n(x.sqrt())
            }
            "POWER" => {
                Self::arity(name, args, 2, 2)?;
                let (a, b) = (self.val(&args[0])?, self.val(&args[1])?);
                match binary(Op::Pow, &a, &b) {
                    Val::Error(e) => fail(&e),
                    r => v(r),
                }
            }
            "MOD" => {
                Self::arity(name, args, 2, 2)?;
                let (x, d) = (self.num(&args[0])?, self.num(&args[1])?);
                if d == 0.0 {
                    return fail(DIV0);
                }
                n(x - d * (x / d).floor())
            }
            "PI" => n(std::f64::consts::PI),
            "CEILING" | "FLOOR" => {
                Self::arity(name, args, 2, 2)?;
                let (x, sig) = (self.num(&args[0])?, self.num(&args[1])?);
                if sig == 0.0 {
                    return n(0.0);
                }
                if x > 0.0 && sig < 0.0 {
                    return fail(NUM);
                }
                let q = round_sig(x / sig, 15);
                n(if name == "CEILING" {
                    q.ceil()
                } else {
                    q.floor()
                } * sig)
            }
            "CEILING.MATH" | "FLOOR.MATH" => {
                let x = self.num(args.first().unwrap_or(&Expr::Missing))?;
                let sig = self.opt_num(args, 1, 1.0)?.abs();
                let away = self.opt_num(args, 2, 0.0)? != 0.0 && x < 0.0;
                if sig == 0.0 {
                    return n(0.0);
                }
                let q = round_sig(x / sig, 15);
                let up = name == "CEILING.MATH";
                let r = if away {
                    if up {
                        q.floor()
                    } else {
                        q.ceil()
                    }
                } else if up {
                    q.ceil()
                } else {
                    q.floor()
                };
                n(r * sig)
            }
            "SUMPRODUCT" => {
                let mut matrices = Vec::new();
                for arg in args {
                    let res = self.eval(arg)?;
                    matrices.push(self.matrix(res)?);
                }
                let Some(first) = matrices.first() else {
                    return fail(VALUE);
                };
                let (rows, cols) = (first.len(), first.first().map(Vec::len).unwrap_or(0));
                if matrices
                    .iter()
                    .any(|m| m.len() != rows || m.first().map(Vec::len).unwrap_or(0) != cols)
                {
                    return fail(VALUE);
                }
                let mut total = 0.0;
                for r in 0..rows {
                    for c in 0..cols {
                        let mut product = 1.0;
                        for m in &matrices {
                            match &m[r][c] {
                                Val::Number(x) => product *= x,
                                Val::Error(e) => return fail(e),
                                _ => product = 0.0,
                            }
                        }
                        total += product;
                    }
                }
                n(total)
            }

            "LN" | "LOG10" | "EXP" => {
                Self::arity(name, args, 1, 1)?;
                let x = self.num(&args[0])?;
                match name {
                    "EXP" => n(x.exp()),
                    _ if x <= 0.0 => fail(NUM),
                    "LN" => n(x.ln()),
                    _ => n(x.log10()),
                }
            }
            "LOG" => {
                Self::arity(name, args, 1, 2)?;
                let (x, base) = (self.num(&args[0])?, self.opt_num(args, 1, 10.0)?);
                if x <= 0.0 || base <= 0.0 {
                    return fail(NUM);
                }
                if base == 1.0 {
                    return fail(DIV0);
                }
                n(round_sig(x.ln() / base.ln(), 15))
            }
            "EVEN" | "ODD" => {
                Self::arity(name, args, 1, 1)?;
                let x = self.num(&args[0])?;
                let up = round_sig(x.abs(), 15).ceil();
                let r = match (name, up % 2.0 == 0.0) {
                    ("EVEN", true) | ("ODD", false) => up,
                    _ => up + 1.0,
                };
                n(if x < 0.0 { -r } else { r })
            }
            "MROUND" => {
                Self::arity(name, args, 2, 2)?;
                let (x, m) = (self.num(&args[0])?, self.num(&args[1])?);
                if m == 0.0 || x == 0.0 {
                    return n(0.0);
                }
                if (x > 0.0) != (m > 0.0) {
                    return fail(NUM);
                }
                n(round_sig(round_sig(x / m, 15).round() * m, 15))
            }
            "QUOTIENT" => {
                Self::arity(name, args, 2, 2)?;
                let (a, b) = (self.num(&args[0])?, self.num(&args[1])?);
                if b == 0.0 {
                    return fail(DIV0);
                }
                n((a / b).trunc())
            }
            "SUMSQ" => n(self.numbers(args)?.iter().map(|x| x * x).sum()),
            "GCD" | "LCM" => {
                let mut whole = Vec::new();
                for x in self.numbers(args)? {
                    if x < 0.0 || x >= 2f64.powi(53) {
                        return fail(NUM);
                    }
                    whole.push(x.trunc() as u128);
                }
                let gcd = |mut a: u128, mut b: u128| {
                    while b != 0 {
                        (a, b) = (b, a % b);
                    }
                    a
                };
                let r = if name == "GCD" {
                    whole.into_iter().fold(0, gcd)
                } else {
                    let mut acc: u128 = 1;
                    for x in whole {
                        if x == 0 {
                            return n(0.0);
                        }
                        acc = acc / gcd(acc, x) * x;
                        if acc >= 1 << 53 {
                            return fail(NUM);
                        }
                    }
                    acc
                };
                n(r as f64)
            }
            "ISEVEN" | "ISODD" => {
                Self::arity(name, args, 1, 1)?;
                let even = self.num(&args[0])?.trunc() % 2.0 == 0.0;
                v(Val::Bool(even == (name == "ISEVEN")))
            }

            // Statistics
            "LARGE" | "SMALL" => {
                Self::arity(name, args, 2, 2)?;
                let mut nums = self.numbers(&args[..1])?;
                let k = self.num(&args[1])?.ceil();
                if k < 1.0 || k > nums.len() as f64 {
                    return fail(NUM);
                }
                nums.sort_by(|a, b| a.partial_cmp(b).unwrap_or(Ordering::Equal));
                if name == "LARGE" {
                    nums.reverse();
                }
                n(nums[k as usize - 1])
            }
            "RANK" | "RANK.EQ" => {
                Self::arity(name, args, 2, 3)?;
                let x = self.num(&args[0])?;
                let nums = self.numbers(&args[1..2])?;
                let ascending = self.opt_num(args, 2, 0.0)? != 0.0;
                if !nums.contains(&x) {
                    return fail(NA);
                }
                let before = nums
                    .iter()
                    .filter(|&&y| if ascending { y < x } else { y > x })
                    .count();
                n(before as f64 + 1.0)
            }
            "STDEV" | "STDEV.S" | "STDEV.P" | "VAR" | "VAR.S" | "VAR.P" => {
                let nums = self.numbers(args)?;
                let sample = !name.ends_with(".P");
                let min = if sample { 2 } else { 1 };
                if nums.len() < min {
                    return fail(DIV0);
                }
                let mean = nums.iter().sum::<f64>() / nums.len() as f64;
                let squares: f64 = nums.iter().map(|x| (x - mean).powi(2)).sum();
                let var = squares / (nums.len() - usize::from(sample)) as f64;
                n(if name.starts_with("STDEV") {
                    var.sqrt()
                } else {
                    var
                })
            }
            "MODE" | "MODE.SNGL" => {
                let nums = self.numbers(args)?;
                // count and first position of each value: ties go to the one seen first
                let mut seen: std::collections::HashMap<u64, (usize, usize)> =
                    std::collections::HashMap::new();
                for (i, x) in nums.iter().enumerate() {
                    seen.entry(x.to_bits()).or_insert((0, i)).0 += 1;
                }
                let best = seen
                    .into_iter()
                    .filter(|(_, (count, _))| *count > 1)
                    .max_by(|a, b| a.1 .0.cmp(&b.1 .0).then(b.1 .1.cmp(&a.1 .1)));
                match best {
                    Some((bits, _)) => n(f64::from_bits(bits)),
                    None => fail(NA),
                }
            }

            // Logic
            "TRUE" => v(Val::Bool(true)),
            "FALSE" => v(Val::Bool(false)),
            "IF" => {
                Self::arity(name, args, 1, 3)?;
                let cond = self.boolean(&args[0])?;
                let branch = if cond { args.get(1) } else { args.get(2) };
                match branch {
                    None if !cond => v(Val::Bool(false)),
                    None | Some(Expr::Missing) => v(Val::Number(0.0)),
                    Some(e) => self.eval(e),
                }
            }
            "IFS" => {
                for pair in args.chunks(2) {
                    let [cond, value] = pair else {
                        return unsupported("IFS sin valor");
                    };
                    if self.boolean(cond)? {
                        return self.eval(value);
                    }
                }
                fail(NA)
            }
            "SWITCH" => {
                if args.len() < 3 {
                    return unsupported("SWITCH incompleto");
                }
                let subject = self.val(&args[0])?;
                let rest = &args[1..];
                for pair in rest.chunks(2) {
                    match pair {
                        [case, value] => {
                            if lookup_equal(&subject, &self.val(case)?, false) {
                                return self.eval(value);
                            }
                        }
                        [default] => return self.eval(default),
                        _ => {}
                    }
                }
                fail(NA)
            }
            "IFERROR" | "IFNA" => {
                Self::arity(name, args, 2, 2)?;
                let first = self.eval(&args[0]).and_then(|res| match res {
                    Res::Val(Val::Error(e)) => Err(Flow::Error(e)),
                    Res::Range(r) if r.rows() == 1 && r.cols() == 1 => {
                        self.scalar(Res::Range(r)).map(Res::Val)
                    }
                    other => Ok(other),
                });
                match first {
                    Err(Flow::Error(e)) if name == "IFERROR" || e == NA => self.eval(&args[1]),
                    other => other,
                }
            }
            "AND" | "OR" | "XOR" => {
                let mut seen = Vec::new();
                for (value, from_range) in self.all_values(args)? {
                    match value {
                        Val::Error(e) => return fail(&e),
                        Val::Bool(b) => seen.push(b),
                        Val::Number(x) => seen.push(x != 0.0),
                        Val::Text(_) | Val::Empty if from_range => {}
                        Val::Empty => {}
                        other => seen.push(to_bool(&other)?),
                    }
                }
                if seen.is_empty() {
                    return fail(VALUE);
                }
                v(Val::Bool(match name {
                    "AND" => seen.iter().all(|b| *b),
                    "OR" => seen.iter().any(|b| *b),
                    _ => seen.iter().filter(|b| **b).count() % 2 == 1,
                }))
            }
            "NOT" => {
                Self::arity(name, args, 1, 1)?;
                v(Val::Bool(!self.boolean(&args[0])?))
            }

            // Information
            "ISBLANK" | "ISNUMBER" | "ISTEXT" | "ISLOGICAL" | "ISNONTEXT" | "ISERROR" | "ISERR"
            | "ISNA" => {
                Self::arity(name, args, 1, 1)?;
                let value = match self.val(&args[0]) {
                    Ok(v) => v,
                    Err(Flow::Error(e)) => Val::Error(e),
                    Err(other) => return Err(other),
                };
                v(Val::Bool(match (name, &value) {
                    ("ISBLANK", Val::Empty) => true,
                    ("ISNUMBER", Val::Number(_)) => true,
                    ("ISTEXT", Val::Text(_)) => true,
                    ("ISNONTEXT", v) => !matches!(v, Val::Text(_)),
                    ("ISLOGICAL", Val::Bool(_)) => true,
                    ("ISERROR", Val::Error(_)) => true,
                    ("ISERR", Val::Error(e)) => e != NA,
                    ("ISNA", Val::Error(e)) => e == NA,
                    _ => false,
                }))
            }
            "NA" => fail(NA),

            // Conditional aggregation
            "SUMIF" | "AVERAGEIF" | "COUNTIF" => {
                let (range, criterion, target) = match (name, args) {
                    ("COUNTIF", [r, c]) => (r, c, None),
                    (_, [r, c]) => (r, c, None),
                    (n2, [r, c, t]) if n2 != "COUNTIF" => (r, c, Some(t)),
                    _ => return unsupported(format!("{} con {} argumentos", name, args.len())),
                };
                let range = self.arg_grid(range)?;
                let criterion = Criterion::parse(&self.val(criterion)?);
                let target = match target {
                    Some(t) => Some(self.arg_grid(t)?),
                    None => None,
                };
                self.conditional(name, &[(range, criterion)], target.as_ref())
            }
            "SUMIFS" | "AVERAGEIFS" | "MAXIFS" | "MINIFS" | "COUNTIFS" => {
                let (target, pairs) = if name == "COUNTIFS" {
                    (None, args)
                } else {
                    (args.first(), args.get(1..).unwrap_or(&[]))
                };
                if pairs.is_empty() || pairs.len() % 2 != 0 {
                    return unsupported(format!("{} con {} argumentos", name, args.len()));
                }
                let target = match target {
                    Some(t) => Some(self.arg_grid(t)?),
                    None => None,
                };
                let mut conditions = Vec::new();
                for pair in pairs.chunks(2) {
                    let grid = self.arg_grid(&pair[0])?;
                    let criterion = Criterion::parse(&self.val(&pair[1])?);
                    conditions.push((grid, criterion));
                }
                self.conditional(name, &conditions, target.as_ref())
            }

            // Lookup
            "VLOOKUP" | "HLOOKUP" => {
                Self::arity(name, args, 3, 4)?;
                let lookup = self.val(&args[0])?;
                let table = self.arg_grid(&args[1])?;
                let index = self.num(&args[2])?.trunc() as i64;
                let approximate = match args.get(3) {
                    None | Some(Expr::Missing) => true,
                    Some(e) => self.boolean(e)?,
                };
                let vertical = name == "VLOOKUP";
                let (count, width) = if vertical {
                    (table.rows(), table.cols())
                } else {
                    (table.cols(), table.rows())
                };
                if index < 1 {
                    return fail(VALUE);
                }
                if index as usize > width {
                    return fail(REF);
                }
                let key = |i: usize| {
                    if vertical {
                        table.get(i, 0)
                    } else {
                        table.get(0, i)
                    }
                };
                let searched = if vertical {
                    table.used().0
                } else {
                    table.used().1
                };
                let found = if approximate {
                    approximate_position(&lookup, (0..searched.min(count)).map(key), 1)
                } else {
                    (0..searched.min(count)).find(|i| lookup_equal(&lookup, &key(*i), true))
                };
                let Some(i) = found else { return fail(NA) };
                let col = index as usize - 1;
                v(if vertical {
                    table.get(i, col)
                } else {
                    table.get(col, i)
                })
            }
            "MATCH" => {
                Self::arity(name, args, 2, 3)?;
                let lookup = self.val(&args[0])?;
                let grid = self.arg_grid(&args[1])?;
                let mode = self.opt_num(args, 2, 1.0)?;
                let Some(len) = grid.vector_len() else {
                    return fail(NA);
                };
                let used = if grid.rows() == 1 {
                    grid.used().1
                } else {
                    grid.used().0
                };
                let values = (0..used.min(len)).map(|i| grid.vector_get(i));
                let found = if mode == 0.0 {
                    values
                        .enumerate()
                        .find(|(_, v)| lookup_equal(&lookup, v, true))
                        .map(|(i, _)| i)
                } else {
                    approximate_position(&lookup, values, if mode > 0.0 { 1 } else { -1 })
                };
                found
                    .map(|i| Res::Val(Val::Number(i as f64 + 1.0)))
                    .map(Ok)
                    .unwrap_or_else(|| fail(NA))
            }
            "XLOOKUP" => {
                Self::arity(name, args, 3, 6)?;
                let lookup = self.val(&args[0])?;
                let keys = self.arg_grid(&args[1])?;
                let results = self.eval(&args[2])?;
                let mode = self.opt_num(args, 4, 0.0)?;
                let search = self.opt_num(args, 5, 1.0)?;
                let Some(len) = keys.vector_len() else {
                    return fail(VALUE);
                };
                let used = if keys.rows() == 1 {
                    keys.used().1
                } else {
                    keys.used().0
                };
                let mut order: Vec<usize> = (0..used.min(len)).collect();
                if search < 0.0 {
                    order.reverse();
                } else if search.abs() == 2.0 {
                    return unsupported("XLOOKUP con búsqueda binaria");
                }
                let mut found = order
                    .iter()
                    .copied()
                    .find(|i| lookup_equal(&lookup, &keys.vector_get(*i), mode == 2.0));
                if found.is_none() && (mode == -1.0 || mode == 1.0) {
                    let mut best: Option<(usize, Val)> = None;
                    for i in order {
                        let candidate = keys.vector_get(i);
                        if !same_kind(&candidate, &lookup) {
                            continue;
                        }
                        let ord = compare(&candidate, &lookup);
                        let fits = if mode < 0.0 {
                            ord == Ordering::Less
                        } else {
                            ord == Ordering::Greater
                        };
                        let better = best.as_ref().is_none_or(|(_, b)| {
                            let o = compare(&candidate, b);
                            if mode < 0.0 {
                                o == Ordering::Greater
                            } else {
                                o == Ordering::Less
                            }
                        });
                        if fits && better {
                            best = Some((i, candidate));
                        }
                    }
                    found = best.map(|(i, _)| i);
                }
                let Some(i) = found else {
                    return match args.get(3) {
                        Some(e) if *e != Expr::Missing => self.eval(e),
                        _ => fail(NA),
                    };
                };
                let horizontal = keys.rows() == 1 && len > 1;
                match results {
                    Res::Range(r) => {
                        let picked = if horizontal {
                            RangeRef {
                                first: CellRef {
                                    row: r.first.row,
                                    col: r.first.col + i as u32,
                                },
                                last: CellRef {
                                    row: r.last.row,
                                    col: r.first.col + i as u32,
                                },
                                ..r
                            }
                        } else {
                            RangeRef {
                                first: CellRef {
                                    row: r.first.row + i as u32,
                                    col: r.first.col,
                                },
                                last: CellRef {
                                    row: r.first.row + i as u32,
                                    col: r.last.col,
                                },
                                ..r
                            }
                        };
                        if picked.last.row > r.last.row && !horizontal
                            || picked.last.col > r.last.col && horizontal
                        {
                            return fail(VALUE);
                        }
                        Ok(Res::Range(picked))
                    }
                    other => {
                        let grid = self.grid(other)?;
                        v(if horizontal {
                            grid.get(0, i)
                        } else {
                            grid.get(i, 0)
                        })
                    }
                }
            }
            "INDEX" => {
                Self::arity(name, args, 2, 3)?;
                let res = self.eval(&args[0])?;
                let row = self.num(&args[1])?.trunc() as i64;
                let col = self.opt_num(args, 2, 0.0)?.trunc() as i64;
                if row < 0 || col < 0 {
                    return fail(VALUE);
                }
                match res {
                    Res::Range(r) => {
                        let (rows, cols) = (r.rows() as i64, r.cols() as i64);
                        // A vector takes a single index along its length
                        let (row, col) = if args.len() == 2 && rows == 1 {
                            (1, row)
                        } else if args.len() == 2 && cols == 1 {
                            (row, 1)
                        } else {
                            (row, col)
                        };
                        if row > rows || col > cols {
                            return fail(REF);
                        }
                        let (r1, r2) = if row == 0 {
                            (r.first.row, r.last.row)
                        } else {
                            (r.first.row + row as u32 - 1, r.first.row + row as u32 - 1)
                        };
                        let (c1, c2) = if col == 0 {
                            (r.first.col, r.last.col)
                        } else {
                            (r.first.col + col as u32 - 1, r.first.col + col as u32 - 1)
                        };
                        Ok(Res::Range(RangeRef {
                            sheet: r.sheet,
                            first: CellRef { row: r1, col: c1 },
                            last: CellRef { row: r2, col: c2 },
                        }))
                    }
                    other => {
                        let grid = self.grid(other)?;
                        let (row, col) = if args.len() == 2 && grid.rows() == 1 {
                            (1, row)
                        } else if args.len() == 2 && grid.cols() == 1 {
                            (row, 1)
                        } else {
                            (row, col)
                        };
                        if row == 0 || col == 0 {
                            return unsupported("INDEX de una fila o columna entera de una matriz");
                        }
                        if row as usize > grid.rows() || col as usize > grid.cols() {
                            return fail(REF);
                        }
                        v(grid.get(row as usize - 1, col as usize - 1))
                    }
                }
            }
            "CHOOSE" => {
                if args.len() < 2 {
                    return unsupported("CHOOSE incompleto");
                }
                let i = self.num(&args[0])?.trunc() as usize;
                match args.get(i).filter(|_| i >= 1) {
                    Some(e) => self.eval(e),
                    None => fail(VALUE),
                }
            }
            "ROW" | "COLUMN" => match args.first() {
                None => v(Val::Number(if name == "ROW" {
                    self.at.row
                } else {
                    self.at.col
                } as f64)),
                Some(Expr::Ref(area)) => {
                    let r = self.area(area)?;
                    v(Val::Number(if name == "ROW" {
                        r.first.row
                    } else {
                        r.first.col
                    } as f64))
                }
                _ => unsupported(format!("{} de algo que no es una referencia", name)),
            },
            "ROWS" | "COLUMNS" => {
                Self::arity(name, args, 1, 1)?;
                let grid = self.arg_grid(&args[0])?;
                v(Val::Number(if name == "ROWS" {
                    grid.rows()
                } else {
                    grid.cols()
                } as f64))
            }

            // Text
            "CONCATENATE" => {
                let mut out = String::new();
                for arg in args {
                    out.push_str(&self.text(arg)?);
                }
                v(Val::Text(out))
            }
            "CONCAT" => {
                let mut out = String::new();
                for (value, _) in self.all_values(args)? {
                    out.push_str(&to_text(&value)?);
                }
                v(Val::Text(out))
            }
            "TEXTJOIN" => {
                if args.len() < 3 {
                    return unsupported("TEXTJOIN incompleto");
                }
                let delimiter = self.text(&args[0])?;
                let skip_empty = self.boolean(&args[1])?;
                let mut parts = Vec::new();
                for (value, _) in self.all_values(&args[2..])? {
                    let t = to_text(&value)?;
                    if !(skip_empty && t.is_empty()) {
                        parts.push(t);
                    }
                }
                v(Val::Text(parts.join(&delimiter)))
            }
            "LEFT" | "RIGHT" => {
                Self::arity(name, args, 1, 2)?;
                let t: Vec<char> = self.text(&args[0])?.chars().collect();
                let count = self.opt_num(args, 1, 1.0)?;
                if count < 0.0 {
                    return fail(VALUE);
                }
                let count = (count as usize).min(t.len());
                let slice = if name == "LEFT" {
                    &t[..count]
                } else {
                    &t[t.len() - count..]
                };
                v(Val::Text(slice.iter().collect()))
            }
            "MID" => {
                Self::arity(name, args, 3, 3)?;
                let t: Vec<char> = self.text(&args[0])?.chars().collect();
                let (start, count) = (self.num(&args[1])?.trunc(), self.num(&args[2])?.trunc());
                if start < 1.0 || count < 0.0 {
                    return fail(VALUE);
                }
                let from = (start as usize - 1).min(t.len());
                let to = (from + count as usize).min(t.len());
                v(Val::Text(t[from..to].iter().collect()))
            }
            "LEN" => {
                Self::arity(name, args, 1, 1)?;
                v(Val::Number(self.text(&args[0])?.chars().count() as f64))
            }
            "UPPER" | "LOWER" | "PROPER" | "TRIM" => {
                Self::arity(name, args, 1, 1)?;
                let t = self.text(&args[0])?;
                v(Val::Text(match name {
                    "UPPER" => t.to_uppercase(),
                    "LOWER" => t.to_lowercase(),
                    "TRIM" => t
                        .split(' ')
                        .filter(|w| !w.is_empty())
                        .collect::<Vec<_>>()
                        .join(" "),
                    _ => {
                        let mut out = String::new();
                        let mut after_letter = false;
                        for c in t.chars() {
                            if after_letter {
                                out.extend(c.to_lowercase());
                            } else {
                                out.extend(c.to_uppercase());
                            }
                            after_letter = c.is_alphabetic();
                        }
                        out
                    }
                }))
            }
            "SUBSTITUTE" => {
                Self::arity(name, args, 3, 4)?;
                let (t, old, new) = (
                    self.text(&args[0])?,
                    self.text(&args[1])?,
                    self.text(&args[2])?,
                );
                if old.is_empty() {
                    return v(Val::Text(t));
                }
                match args.get(3) {
                    None => v(Val::Text(t.replace(&old, &new))),
                    Some(e) => {
                        let nth = self.num(e)?.trunc();
                        if nth < 1.0 {
                            return fail(VALUE);
                        }
                        match t.match_indices(&old).nth(nth as usize - 1) {
                            Some((at, _)) => v(Val::Text(format!(
                                "{}{}{}",
                                &t[..at],
                                new,
                                &t[at + old.len()..]
                            ))),
                            None => v(Val::Text(t)),
                        }
                    }
                }
            }
            "REPT" => {
                Self::arity(name, args, 2, 2)?;
                let (t, times) = (self.text(&args[0])?, self.num(&args[1])?.trunc());
                if times < 0.0 || t.len() as f64 * times > 32_767.0 {
                    return fail(VALUE);
                }
                v(Val::Text(t.repeat(times as usize)))
            }
            "EXACT" => {
                Self::arity(name, args, 2, 2)?;
                v(Val::Bool(self.text(&args[0])? == self.text(&args[1])?))
            }
            "FIND" | "SEARCH" => {
                Self::arity(name, args, 2, 3)?;
                let (needle, haystack) = (self.text(&args[0])?, self.text(&args[1])?);
                let start = self.opt_num(args, 2, 1.0)?.trunc();
                let chars: Vec<char> = haystack.chars().collect();
                if start < 1.0 || start as usize > chars.len() + 1 {
                    return fail(VALUE);
                }
                let from = start as usize - 1;
                let rest: String = chars[from..].iter().collect();
                let found = if name == "FIND" {
                    rest.find(&needle).map(|b| rest[..b].chars().count())
                } else if has_wildcards(&needle) {
                    let rest_chars: Vec<char> = rest.chars().collect();
                    (0..=rest_chars.len()).find(|i| {
                        (*i..=rest_chars.len()).any(|j| {
                            wildcard_match(&needle, &rest_chars[*i..j].iter().collect::<String>())
                        })
                    })
                } else {
                    let lower = rest.to_lowercase();
                    lower
                        .find(&needle.to_lowercase())
                        .map(|b| lower[..b].chars().count())
                };
                found
                    .map(|i| Res::Val(Val::Number((from + i + 1) as f64)))
                    .map(Ok)
                    .unwrap_or_else(|| fail(VALUE))
            }
            "CHAR" => {
                Self::arity(name, args, 1, 1)?;
                let code = self.num(&args[0])?.trunc();
                if !(1.0..=255.0).contains(&code) {
                    return fail(VALUE);
                }
                let code = code as u8;
                v(Val::Text(
                    match code {
                        0x80..=0x9F => CP1252_HIGH[(code - 0x80) as usize],
                        _ => char::from(code),
                    }
                    .to_string(),
                ))
            }
            "CODE" => {
                Self::arity(name, args, 1, 1)?;
                let Some(c) = self.text(&args[0])?.chars().next() else {
                    return fail(VALUE);
                };
                let code = match CP1252_HIGH.iter().position(|&x| x == c) {
                    Some(i) => 0x80 + i as u32,
                    None if (c as u32) < 256 => c as u32,
                    // Windows shows characters outside its code page as "?"
                    None => 63,
                };
                n(code as f64)
            }
            "CLEAN" => {
                Self::arity(name, args, 1, 1)?;
                let t = self.text(&args[0])?;
                v(Val::Text(t.chars().filter(|c| *c as u32 >= 32).collect()))
            }
            "VALUE" => {
                Self::arity(name, args, 1, 1)?;
                match self.val(&args[0])? {
                    Val::Number(x) => n(x),
                    Val::Empty => n(0.0),
                    Val::Text(t) => parse_number(&t).map(n).unwrap_or_else(|| fail(VALUE)),
                    _ => fail(VALUE),
                }
            }
            "TEXT" => {
                Self::arity(name, args, 2, 2)?;
                let value = self.val(&args[0])?;
                let code = self.text(&args[1])?;
                let locale = Locale::from_tag("en");
                let shown = match &value {
                    Val::Text(t) => match parse_number(t) {
                        Some(x) => numfmt::format(
                            numfmt::Value::Number(x),
                            &code,
                            &locale,
                            self.ctx.date1904(),
                        ),
                        None => numfmt::format(
                            numfmt::Value::Text(t),
                            &code,
                            &locale,
                            self.ctx.date1904(),
                        ),
                    },
                    Val::Bool(b) => {
                        numfmt::format(numfmt::Value::Bool(*b), &code, &locale, self.ctx.date1904())
                    }
                    other => numfmt::format(
                        numfmt::Value::Number(to_num(other)?),
                        &code,
                        &locale,
                        self.ctx.date1904(),
                    ),
                };
                v(Val::Text(shown.text))
            }

            // Dates
            "TODAY" => n(self.ctx.now().floor()),
            "NOW" => n(self.ctx.now()),
            "DATE" => {
                Self::arity(name, args, 3, 3)?;
                let (y, m, d) = (
                    self.num(&args[0])?.trunc(),
                    self.num(&args[1])?.trunc(),
                    self.num(&args[2])?.trunc(),
                );
                n(self.serial(y as i64, m as i64, d as i64)?)
            }
            "YEAR" | "MONTH" | "DAY" => {
                Self::arity(name, args, 1, 1)?;
                let serial = self.num(&args[0])?;
                let (y, m, d) = self.ymd(serial)?;
                n(match name {
                    "YEAR" => y,
                    "MONTH" => m,
                    _ => d,
                } as f64)
            }
            "WEEKDAY" => {
                Self::arity(name, args, 1, 2)?;
                let serial = self.num(&args[0])?.floor();
                let dow = self.weekday(serial)?;
                let kind = self.opt_num(args, 1, 1.0)?;
                n(match kind as i64 {
                    1 | 17 => dow + 1,
                    2 | 11 => (dow + 6) % 7 + 1,
                    3 => (dow + 6) % 7,
                    _ => return fail(NUM),
                } as f64)
            }
            "EDATE" | "EOMONTH" => {
                Self::arity(name, args, 2, 2)?;
                let serial = self.num(&args[0])?;
                let (y, m, d) = self.ymd(serial)?;
                let months = self.num(&args[1])?.trunc() as i64;
                let total = y * 12 + (m - 1) + months;
                let (ny, nm) = (total.div_euclid(12), total.rem_euclid(12) + 1);
                let last_day = days_from_civil(
                    if nm == 12 { ny + 1 } else { ny },
                    if nm == 12 { 1 } else { nm + 1 },
                    1,
                ) - days_from_civil(ny, nm, 1);
                let day = if name == "EOMONTH" {
                    last_day
                } else {
                    d.min(last_day)
                };
                n(self.serial(ny, nm, day)?)
            }
            "DAYS" => {
                Self::arity(name, args, 2, 2)?;
                n(self.num(&args[0])?.floor() - self.num(&args[1])?.floor())
            }

            "TIME" => {
                Self::arity(name, args, 3, 3)?;
                let (h, m, sec) = (
                    self.num(&args[0])?.trunc(),
                    self.num(&args[1])?.trunc(),
                    self.num(&args[2])?.trunc(),
                );
                let total = h * 3600.0 + m * 60.0 + sec;
                if total < 0.0 {
                    return fail(NUM);
                }
                n(total.rem_euclid(86_400.0) / 86_400.0)
            }
            "HOUR" | "MINUTE" | "SECOND" => {
                Self::arity(name, args, 1, 1)?;
                let x = self.num(&args[0])?;
                if x < 0.0 {
                    return fail(NUM);
                }
                let secs = (round_sig(x - x.floor(), 15) * 86_400.0).round() as i64 % 86_400;
                n(match name {
                    "HOUR" => secs / 3600,
                    "MINUTE" => secs / 60 % 60,
                    _ => secs % 60,
                } as f64)
            }
            "DATEDIF" => {
                Self::arity(name, args, 3, 3)?;
                let (start, end) = (self.num(&args[0])?.floor(), self.num(&args[1])?.floor());
                let unit = self.text(&args[2])?.to_uppercase();
                if start > end {
                    return fail(NUM);
                }
                let ((sy, sm, sd), (ey, em, ed)) = (self.ymd(start)?, self.ymd(end)?);
                let months = (ey - sy) * 12 + (em - sm) - i64::from(ed < sd);
                n(match unit.as_str() {
                    "D" => end - start,
                    "M" => months as f64,
                    "Y" => (months / 12) as f64,
                    "YM" => (months % 12) as f64,
                    // days since the same day of the month before (Excel's own rule,
                    // which can go negative after a longer month)
                    "MD" if ed >= sd => (ed - sd) as f64,
                    "MD" => end - self.serial(ey, em - 1, sd)?,
                    "YD" => {
                        let mut anniversary = self.serial(ey, sm, sd)?;
                        if anniversary > end {
                            anniversary = self.serial(ey - 1, sm, sd)?;
                        }
                        end - anniversary
                    }
                    _ => return fail(NUM),
                })
            }
            "NETWORKDAYS" => {
                Self::arity(name, args, 2, 3)?;
                let (a, b) = (self.num(&args[0])?.floor(), self.num(&args[1])?.floor());
                let holidays = self.holidays(args, 2)?;
                let (lo, hi) = if a <= b { (a, b) } else { (b, a) };
                let total = (hi - lo) as i64 + 1;
                let full_weeks = total / 7;
                let mut count = full_weeks * 5;
                for k in 0..total % 7 {
                    if self.is_weekday(lo + (full_weeks * 7 + k) as f64)? {
                        count += 1;
                    }
                }
                for &h in &holidays {
                    let h = h as f64;
                    if (lo..=hi).contains(&h) && self.is_weekday(h)? {
                        count -= 1;
                    }
                }
                n(if a <= b { count } else { -count } as f64)
            }
            "WORKDAY" => {
                Self::arity(name, args, 2, 3)?;
                let mut day = self.num(&args[0])?.floor();
                let days = self.num(&args[1])?.trunc();
                let holidays = self.holidays(args, 2)?;
                if days.abs() > 1_000_000.0 {
                    return fail(NUM);
                }
                let (step, mut left) = (days.signum(), days.abs() as i64);
                while left > 0 {
                    day += step;
                    if self.is_weekday(day)? && !holidays.contains(&(day as i64)) {
                        left -= 1;
                    }
                }
                n(day)
            }
            "WEEKNUM" | "ISOWEEKNUM" => {
                Self::arity(name, args, 1, if name == "WEEKNUM" { 2 } else { 1 })?;
                let day = self.num(&args[0])?.floor();
                let kind = if name == "WEEKNUM" {
                    self.opt_num(args, 1, 1.0)? as i64
                } else {
                    21
                };
                // first day of the week, 0 = Sunday
                let first = match kind {
                    1 | 17 => 0,
                    2 | 11 => 1,
                    12..=16 => kind - 10,
                    21 => {
                        // ISO 8601: the week belongs to the year of its Thursday
                        let monday_based = (self.weekday(day)? + 6) % 7;
                        let thursday = day - monday_based as f64 + 3.0;
                        let (year, _, _) = self.ymd(thursday)?;
                        let jan1 = self.serial(year, 1, 1)?;
                        return n(((thursday - jan1) / 7.0).floor() + 1.0);
                    }
                    _ => return fail(NUM),
                };
                let (year, _, _) = self.ymd(day)?;
                let jan1 = self.serial(year, 1, 1)?;
                let offset = (self.weekday(jan1)? - first).rem_euclid(7);
                n(((day - jan1) as i64 + offset).div_euclid(7) as f64 + 1.0)
            }

            // Finance: Excel's sign convention, money paid out is negative
            "PMT" | "PV" | "FV" | "NPER" => {
                Self::arity(name, args, 3, 5)?;
                let (rate, a, b) = (
                    self.num(&args[0])?,
                    self.num(&args[1])?,
                    self.num(&args[2])?,
                );
                let rest = self.opt_num(args, 3, 0.0)?;
                let due = if self.opt_num(args, 4, 0.0)? != 0.0 {
                    1.0
                } else {
                    0.0
                };
                let growth = |periods: f64| (1.0 + rate).powf(periods);
                match name {
                    // PMT(rate, nper, pv, [fv], [type])
                    "PMT" => {
                        let (nper, pv, fv) = (a, b, rest);
                        if nper == 0.0 {
                            return fail(NUM);
                        }
                        if rate == 0.0 {
                            return n(-(pv + fv) / nper);
                        }
                        let g = growth(nper);
                        n(-(rate * (fv + pv * g)) / ((1.0 + rate * due) * (g - 1.0)))
                    }
                    // PV(rate, nper, pmt, [fv], [type])
                    "PV" => {
                        let (nper, pmt, fv) = (a, b, rest);
                        if rate == 0.0 {
                            return n(-(fv + pmt * nper));
                        }
                        let g = growth(nper);
                        n(-(fv + pmt * (1.0 + rate * due) * (g - 1.0) / rate) / g)
                    }
                    // FV(rate, nper, pmt, [pv], [type])
                    "FV" => {
                        let (nper, pmt, pv) = (a, b, rest);
                        if rate == 0.0 {
                            return n(-(pv + pmt * nper));
                        }
                        let g = growth(nper);
                        n(-(pv * g + pmt * (1.0 + rate * due) * (g - 1.0) / rate))
                    }
                    // NPER(rate, pmt, pv, [fv], [type])
                    _ => {
                        let (pmt, pv, fv) = (a, b, rest);
                        if rate == 0.0 {
                            if pmt == 0.0 {
                                return fail(NUM);
                            }
                            return n(-(pv + fv) / pmt);
                        }
                        let k = pmt * (1.0 + rate * due) / rate;
                        let ratio = (k - fv) / (k + pv);
                        if ratio <= 0.0 || rate <= -1.0 {
                            return fail(NUM);
                        }
                        n(ratio.ln() / (1.0 + rate).ln())
                    }
                }
            }

            _ => unsupported(format!("la función {}", name)),
        }
    }

    /// SUMIF(S), COUNTIF(S), AVERAGEIF(S), MAXIFS and MINIFS
    fn conditional(
        &self,
        name: &str,
        conditions: &[(Grid, Criterion)],
        target: Option<&Grid>,
    ) -> R<Res> {
        let (rows, cols) = (conditions[0].0.rows(), conditions[0].0.cols());
        if conditions
            .iter()
            .any(|(g, _)| g.rows() != rows || g.cols() != cols)
            && name.ends_with('S')
        {
            return fail(VALUE);
        }
        // Cells past every used area are empty: only count them, for criteria that match blanks
        let (mut used_rows, mut used_cols) = (0, 0);
        for (g, _) in conditions {
            let (r, c) = g.used();
            used_rows = used_rows.max(r);
            used_cols = used_cols.max(c);
        }
        if let Some(t) = target {
            let (r, c) = t.used();
            used_rows = used_rows.max(r);
            used_cols = used_cols.max(c);
        }
        let (used_rows, used_cols) = (used_rows.min(rows), used_cols.min(cols));
        let mut numbers = Vec::new();
        let mut count = 0f64;
        for r in 0..used_rows {
            for c in 0..used_cols {
                if conditions
                    .iter()
                    .all(|(g, crit)| crit.matches(&g.get(r, c)))
                {
                    count += 1.0;
                    if let Some(t) = target {
                        match t.get(r, c) {
                            Val::Number(x) => numbers.push(x),
                            Val::Error(e) if name != "COUNTIF" => return fail(&e),
                            _ => {}
                        }
                    }
                }
            }
        }
        if name.starts_with("COUNT") {
            if conditions.iter().all(|(_, crit)| crit.matches(&Val::Empty)) {
                count += (rows * cols - used_rows * used_cols) as f64;
            }
            return Ok(Res::Val(Val::Number(count)));
        }
        if target.is_none() {
            // SUMIF / AVERAGEIF without a sum range: the criteria range itself
            let (g, crit) = &conditions[0];
            for r in 0..used_rows {
                for c in 0..used_cols {
                    match g.get(r, c) {
                        Val::Number(x) if crit.matches(&Val::Number(x)) => numbers.push(x),
                        _ => {}
                    }
                }
            }
        }
        let value = match name {
            "SUMIF" | "SUMIFS" => numbers.iter().sum(),
            "AVERAGEIF" | "AVERAGEIFS" if numbers.is_empty() => return fail(DIV0),
            "AVERAGEIF" | "AVERAGEIFS" => numbers.iter().sum::<f64>() / numbers.len() as f64,
            "MAXIFS" => numbers.iter().copied().reduce(f64::max).unwrap_or(0.0),
            _ => numbers.iter().copied().reduce(f64::min).unwrap_or(0.0),
        };
        finite(value).map(Res::Val)
    }
}

/// Approximate match over values sorted ascending (`direction` 1: largest value ≤ lookup) or
/// descending (-1: smallest value ≥ lookup); values of another type are skipped
fn approximate_position(
    lookup: &Val,
    values: impl Iterator<Item = Val>,
    direction: i32,
) -> Option<usize> {
    let mut found = None;
    for (i, v) in values.enumerate() {
        if !same_kind(&v, lookup) {
            continue;
        }
        let ord = compare(&v, lookup);
        let fits = if direction > 0 {
            ord != Ordering::Greater
        } else {
            ord != Ordering::Less
        };
        if fits {
            found = Some(i);
            if ord == Ordering::Equal {
                break;
            }
        } else {
            break;
        }
    }
    found
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::formula::parse;
    use std::collections::HashMap;

    struct Sheet {
        cells: HashMap<CellRef, Val>,
        names: HashMap<String, Expr>,
    }

    impl Context for Sheet {
        fn value(&self, _sheet: usize, at: CellRef) -> Val {
            self.cells.get(&at).cloned().unwrap_or(Val::Empty)
        }
        fn sheet_index(&self, name: &str) -> Option<usize> {
            (name.eq_ignore_ascii_case("Hoja1")).then_some(0)
        }
        fn name(&self, name: &str, _sheet: usize) -> Option<(&Expr, usize)> {
            self.names.get(&name.to_uppercase()).map(|e| (e, 0))
        }
        fn bounds(&self, _sheet: usize) -> (u32, u32) {
            let rows = self.cells.keys().map(|c| c.row).max().unwrap_or(0);
            let cols = self.cells.keys().map(|c| c.col).max().unwrap_or(0);
            (rows, cols)
        }
        fn date1904(&self) -> bool {
            false
        }
        fn now(&self) -> f64 {
            45_000.75
        }
    }

    /// A1:A5 = 10, 20, "x", 40, TRUE; B1:B5 = norte, sur, norte, este, norte; C1 = 0.5
    fn sheet() -> Sheet {
        let mut cells = HashMap::new();
        let mut put = |r: &str, v: Val| {
            cells.insert(CellRef::parse(r).unwrap(), v);
        };
        put("A1", Val::Number(10.0));
        put("A2", Val::Number(20.0));
        put("A3", Val::Text("x".into()));
        put("A4", Val::Number(40.0));
        put("A5", Val::Bool(true));
        for (i, region) in ["norte", "sur", "norte", "este", "norte"]
            .iter()
            .enumerate()
        {
            put(&format!("B{}", i + 1), Val::Text(region.to_string()));
        }
        put("C1", Val::Number(0.5));
        put("D1", Val::Number(1.0));
        put("D2", Val::Number(2.0));
        put("D3", Val::Number(3.0));
        put("E1", Val::Text("uno".into()));
        put("E2", Val::Text("dos".into()));
        put("E3", Val::Text("tres".into()));
        let mut names = HashMap::new();
        names.insert("IVA".to_string(), parse("Hoja1!$C$1").unwrap());
        Sheet { cells, names }
    }

    fn calc(formula: &str) -> Val {
        let ctx = sheet();
        evaluate(
            &ctx,
            0,
            CellRef::parse("H2").unwrap(),
            &parse(formula).unwrap(),
        )
        .unwrap()
    }

    fn num(formula: &str) -> f64 {
        match calc(formula) {
            Val::Number(n) => n,
            other => panic!("{} = {:?}", formula, other),
        }
    }

    fn text(formula: &str) -> String {
        match calc(formula) {
            Val::Text(t) => t,
            other => panic!("{} = {:?}", formula, other),
        }
    }

    fn error(formula: &str) -> String {
        match calc(formula) {
            Val::Error(e) => e,
            other => panic!("{} = {:?}", formula, other),
        }
    }

    #[test]
    fn arithmetic_and_coercions() {
        assert_eq!(num("1+2*3"), 7.0);
        assert_eq!(num("-2^2"), 4.0);
        assert_eq!(num("2^3^2"), 64.0);
        assert_eq!(num("50%*A1"), 5.0);
        assert_eq!(num("\"3\"+A1"), 13.0);
        assert_eq!(num("TRUE+1"), 2.0);
        assert_eq!(num("Z99+1"), 1.0);
        assert_eq!(num("Z99"), 0.0);
        assert_eq!(error("A1/0"), "#DIV/0!");
        assert_eq!(error("A3*2"), "#VALUE!");
        assert_eq!(text("A1&\"-\"&C1&\"-\"&A5"), "10-0.5-TRUE");
        assert_eq!(text("0.1+0.2&\"\""), "0.3");
        assert_eq!(calc("A1>5"), Val::Bool(true));
        assert_eq!(calc("\"a\"=\"A\""), Val::Bool(true));
        assert_eq!(calc("\"b\">1"), Val::Bool(true));
        assert_eq!(num("A1*IVA"), 5.0);
        assert_eq!(error("NOEXISTE+1"), "#NAME?");
    }

    #[test]
    fn aggregates_skip_text_in_ranges() {
        assert_eq!(num("SUM(A1:A5)"), 70.0);
        assert_eq!(num("SUM(A:A)"), 70.0);
        assert_eq!(num("SUM(A3)"), 0.0);
        assert_eq!(num("SUM(1,\"2\",TRUE)"), 4.0);
        assert_eq!(num("AVERAGE(A1:A5)"), 70.0 / 3.0);
        assert_eq!(num("COUNT(A1:A5)"), 3.0);
        assert_eq!(num("COUNTA(A1:A5)"), 5.0);
        assert_eq!(num("COUNTBLANK(A1:A7)"), 2.0);
        assert_eq!(num("MAX(A1:A5)"), 40.0);
        assert_eq!(num("MIN(A1:A5,5)"), 5.0);
        assert_eq!(num("ROUND(2.675,2)"), 2.68);
        assert_eq!(num("ROUND(-2.5,0)"), -3.0);
        assert_eq!(num("ROUNDUP(1.21,1)"), 1.3);
        assert_eq!(num("ROUNDDOWN(-1.29,1)"), -1.2);
        assert_eq!(num("MOD(-3,2)"), 1.0);
        assert_eq!(num("SUMPRODUCT(D1:D3,D1:D3)"), 14.0);
        assert_eq!(num("SUMPRODUCT((B1:B3=\"norte\")*D1:D3)"), 4.0);
        assert_eq!(error("SUMPRODUCT((B1:B5=\"norte\")*A1:A5)"), "#VALUE!");
        assert_eq!(error("AVERAGE(B1:B5)"), "#DIV/0!");
    }

    #[test]
    fn conditional_sums_and_counts() {
        assert_eq!(num("SUMIF(B1:B5,\"norte\",A1:A5)"), 10.0);
        assert_eq!(num("SUMIF(A1:A5,\">15\")"), 60.0);
        assert_eq!(num("COUNTIF(B:B,\"n*\")"), 3.0);
        assert_eq!(num("COUNTIF(B1:B5,\"<>norte\")"), 2.0);
        assert_eq!(num("COUNTIF(A1:A7,\"\")"), 2.0);
        assert_eq!(num("SUMIFS(A1:A5,B1:B5,\"norte\",A1:A5,\">5\")"), 10.0);
        assert_eq!(num("COUNTIFS(B1:B5,\"norte\",A1:A5,\"<>x\")"), 2.0);
        assert_eq!(num("AVERAGEIF(B1:B5,\"sur\",A1:A5)"), 20.0);
        assert_eq!(num("MAXIFS(A1:A5,B1:B5,\"<>norte\")"), 40.0);
    }

    #[test]
    fn logic_and_errors() {
        assert_eq!(text("IF(A1>15,\"alto\",\"bajo\")"), "bajo");
        assert_eq!(calc("IF(A1>15,\"alto\")"), Val::Bool(false));
        assert_eq!(num("IFERROR(1/0,7)"), 7.0);
        assert_eq!(error("IFNA(1/0,7)"), "#DIV/0!");
        assert_eq!(num("IFNA(NA(),7)"), 7.0);
        assert_eq!(calc("AND(A1>5,A2>5)"), Val::Bool(true));
        assert_eq!(calc("OR(A1:A2>15)"), Val::Bool(true));
        assert_eq!(calc("ISERROR(A3*1)"), Val::Bool(true));
        assert_eq!(calc("ISBLANK(Z9)"), Val::Bool(true));
        assert_eq!(text("IFS(A1>50,\"a\",A1>5,\"b\")"), "b");
        assert_eq!(text("SWITCH(D2,1,\"uno\",2,\"dos\",\"otro\")"), "dos");
    }

    #[test]
    fn lookups() {
        assert_eq!(text("VLOOKUP(2,D1:E3,2,FALSE)"), "dos");
        assert_eq!(text("VLOOKUP(2.5,D1:E3,2)"), "dos");
        assert_eq!(error("VLOOKUP(9,D1:E3,2,FALSE)"), "#N/A");
        assert_eq!(error("VLOOKUP(2,D1:E3,3,FALSE)"), "#REF!");
        assert_eq!(text("HLOOKUP(\"dos\",E2:E2,1,FALSE)"), "dos");
        assert_eq!(num("MATCH(\"tres\",E1:E3,0)"), 3.0);
        assert_eq!(num("MATCH(\"T*\",E1:E3,0)"), 3.0);
        assert_eq!(num("MATCH(2.9,D1:D3)"), 2.0);
        assert_eq!(text("INDEX(E1:E3,2)"), "dos");
        assert_eq!(num("INDEX(D1:E3,3,1)"), 3.0);
        assert_eq!(num("SUM(INDEX(D1:E3,0,1))"), 6.0);
        assert_eq!(text("XLOOKUP(3,D1:D3,E1:E3)"), "tres");
        assert_eq!(text("XLOOKUP(9,D1:D3,E1:E3,\"no\")"), "no");
        assert_eq!(text("XLOOKUP(2.5,D1:D3,E1:E3,,1)"), "tres");
        assert_eq!(text("CHOOSE(2,\"a\",\"b\")"), "b");
    }

    #[test]
    fn text_functions() {
        assert_eq!(text("CONCATENATE(\"a\",A1,\"b\")"), "a10b");
        assert_eq!(text("CONCAT(D1:D3)"), "123");
        assert_eq!(text("TEXTJOIN(\", \",TRUE,E1:E3,Z1)"), "uno, dos, tres");
        assert_eq!(
            text("LEFT(\"España\",3)&RIGHT(\"España\",2)&MID(\"España\",2,3)"),
            "Espñaspa"
        );
        assert_eq!(num("LEN(\"ñandú\")"), 5.0);
        assert_eq!(text("PROPER(\"hola mUNDO\")"), "Hola Mundo");
        assert_eq!(text("TRIM(\"  a   b \")"), "a b");
        assert_eq!(text("SUBSTITUTE(\"a-b-c\",\"-\",\"+\",2)"), "a-b+c");
        assert_eq!(num("FIND(\"b\",\"abcb\",3)"), 4.0);
        assert_eq!(num("SEARCH(\"B\",\"abc\")"), 2.0);
        assert_eq!(num("VALUE(\"12.5\")"), 12.5);
        assert_eq!(text("TEXT(1234.5,\"#,##0.00\")"), "1,234.50");
        assert_eq!(text("TEXT(45000,\"yyyy-mm-dd\")"), "2023-03-15");
    }

    #[test]
    fn dates() {
        assert_eq!(num("DATE(2023,3,15)"), 45000.0);
        assert_eq!(num("DATE(2023,14,1)"), num("DATE(2024,2,1)"));
        assert_eq!(
            num("YEAR(45000)*10000+MONTH(45000)*100+DAY(45000)"),
            20230315.0
        );
        assert_eq!(num("WEEKDAY(45000)"), 4.0);
        assert_eq!(num("WEEKDAY(45000,2)"), 3.0);
        assert_eq!(num("EOMONTH(45000,0)"), num("DATE(2023,3,31)"));
        assert_eq!(num("EDATE(DATE(2024,1,31),1)"), num("DATE(2024,2,29)"));
        assert_eq!(num("TODAY()"), 45000.0);
        assert_eq!(num("NOW()"), 45000.75);
        assert_eq!(num("DAYS(DATE(2024,3,1),DATE(2024,2,1))"), 29.0);
    }

    /// `formula` evaluates to `expected` within `tolerance`
    fn near(formula: &str, expected: f64, tolerance: f64) {
        let got = num(formula);
        assert!(
            (got - expected).abs() <= tolerance,
            "{} = {}, expected {}",
            formula,
            got,
            expected
        );
    }

    #[test]
    fn more_math() {
        near("LN(EXP(1))", 1.0, 1e-12);
        assert_eq!(num("LOG(8,2)"), 3.0);
        assert_eq!(num("LOG(100)"), 2.0);
        assert_eq!(num("LOG10(1000)"), 3.0);
        assert_eq!(error("LN(0)"), NUM);
        assert_eq!(error("LOG(10,1)"), DIV0);
        assert_eq!(num("EVEN(1.5)"), 2.0);
        assert_eq!(num("EVEN(3)"), 4.0);
        assert_eq!(num("EVEN(-1)"), -2.0);
        assert_eq!(num("EVEN(0)"), 0.0);
        assert_eq!(num("ODD(1.5)"), 3.0);
        assert_eq!(num("ODD(2)"), 3.0);
        assert_eq!(num("ODD(-1)"), -1.0);
        assert_eq!(num("ODD(0)"), 1.0);
        assert_eq!(num("MROUND(10,3)"), 9.0);
        assert_eq!(num("MROUND(-10,-3)"), -9.0);
        assert_eq!(num("MROUND(1.3,0.2)"), 1.4);
        assert_eq!(error("MROUND(5,-2)"), NUM);
        assert_eq!(num("QUOTIENT(5,2)"), 2.0);
        assert_eq!(num("QUOTIENT(-10,3)"), -3.0);
        assert_eq!(error("QUOTIENT(1,0)"), DIV0);
        assert_eq!(num("SUMSQ(3,4)"), 25.0);
        assert_eq!(num("SUMSQ(D1:D3)"), 14.0);
        assert_eq!(num("GCD(24,36)"), 12.0);
        assert_eq!(num("GCD(5,2)"), 1.0);
        assert_eq!(num("LCM(24,36)"), 72.0);
        assert_eq!(num("LCM(D1:D3)"), 6.0);
        assert_eq!(num("LCM(0,5)"), 0.0);
        assert_eq!(error("GCD(-1,2)"), NUM);
        assert_eq!(calc("ISEVEN(2.5)"), Val::Bool(true));
        assert_eq!(calc("ISODD(-1)"), Val::Bool(true));
        assert_eq!(calc("ISODD(D2)"), Val::Bool(false));
    }

    #[test]
    fn statistics() {
        // A1:A5 = 10, 20, "x", 40, TRUE: only the numbers count
        assert_eq!(num("LARGE(A1:A5,1)"), 40.0);
        assert_eq!(num("LARGE(A1:A5,3)"), 10.0);
        assert_eq!(num("SMALL(A1:A5,1)"), 10.0);
        assert_eq!(num("SMALL({3,1,2},2)"), 2.0);
        assert_eq!(error("LARGE(A1:A5,4)"), NUM);
        assert_eq!(error("SMALL(A1:A5,0)"), NUM);
        assert_eq!(num("RANK(20,A1:A5)"), 2.0);
        assert_eq!(num("RANK(20,A1:A5,1)"), 2.0);
        assert_eq!(num("RANK.EQ(40,A1:A5)"), 1.0);
        assert_eq!(num("RANK(3.5,{7,3.5,3.5,1,2},1)"), 3.0);
        assert_eq!(error("RANK(99,A1:A5)"), NA);
        let sample = "{1345,1301,1368,1322,1310,1370,1318,1350,1303,1299}";
        near(&format!("STDEV.S({})", sample), 27.46392, 1e-5);
        near(&format!("STDEV({})", sample), 27.46392, 1e-5);
        near(&format!("STDEV.P({})", sample), 26.05456, 1e-5);
        near(&format!("VAR.S({})", sample), 754.26667, 1e-5);
        near(&format!("VAR.P({})", sample), 678.84, 1e-9);
        assert_eq!(error("STDEV.S(5)"), DIV0);
        assert_eq!(num("VAR.P(5)"), 0.0);
        assert_eq!(num("MODE({5.6,4,4,3,2,4})"), 4.0);
        assert_eq!(num("MODE.SNGL({1,2,2,3,3})"), 2.0);
        assert_eq!(error("MODE(D1:D3)"), NA);
    }

    #[test]
    fn times_and_working_days() {
        assert_eq!(num("TIME(12,0,0)"), 0.5);
        near("TIME(16,48,10)", 0.700115741, 1e-9);
        assert_eq!(num("TIME(24,0,0)"), 0.0);
        assert_eq!(error("TIME(-1,0,0)"), NUM);
        assert_eq!(num("HOUR(0.75)"), 18.0);
        assert_eq!(num("HOUR(45000.75)"), 18.0);
        assert_eq!(num("MINUTE(TIME(12,45,0))"), 45.0);
        assert_eq!(num("SECOND(TIME(4,48,18))"), 18.0);
        assert_eq!(num("DATEDIF(DATE(2001,1,1),DATE(2003,1,1),\"Y\")"), 2.0);
        assert_eq!(num("DATEDIF(DATE(2001,6,1),DATE(2002,8,15),\"D\")"), 440.0);
        assert_eq!(num("DATEDIF(DATE(2001,6,1),DATE(2002,8,15),\"YD\")"), 75.0);
        assert_eq!(num("DATEDIF(DATE(2001,6,1),DATE(2002,8,15),\"md\")"), 14.0);
        assert_eq!(num("DATEDIF(DATE(2001,6,1),DATE(2002,8,15),\"M\")"), 14.0);
        assert_eq!(num("DATEDIF(DATE(2001,6,1),DATE(2002,8,15),\"YM\")"), 2.0);
        assert_eq!(num("DATEDIF(DATE(2024,1,31),DATE(2024,2,29),\"M\")"), 0.0);
        assert_eq!(error("DATEDIF(DATE(2003,1,1),DATE(2001,1,1),\"Y\")"), NUM);
        assert_eq!(error("DATEDIF(1,2,\"W\")"), NUM);
        assert_eq!(num("NETWORKDAYS(DATE(2012,10,1),DATE(2013,3,1))"), 110.0);
        assert_eq!(
            num("NETWORKDAYS(DATE(2012,10,1),DATE(2013,3,1),DATE(2012,11,22))"),
            109.0
        );
        assert_eq!(num("NETWORKDAYS(DATE(2013,3,1),DATE(2012,10,1))"), -110.0);
        assert_eq!(num("WORKDAY(DATE(2008,10,1),151)"), num("DATE(2009,4,30)"));
        assert_eq!(
            num("WORKDAY(DATE(2008,10,1),151,{39778,39786,39834})"),
            num("DATE(2009,5,5)")
        );
        assert_eq!(num("WORKDAY(DATE(2024,1,8),-1)"), num("DATE(2024,1,5)"));
        assert_eq!(num("WEEKNUM(DATE(2012,3,9))"), 10.0);
        assert_eq!(num("WEEKNUM(DATE(2012,3,9),2)"), 11.0);
        assert_eq!(num("ISOWEEKNUM(DATE(2012,3,9))"), 10.0);
        assert_eq!(num("WEEKNUM(DATE(2012,3,9),21)"), 10.0);
        // 1 January 2021 (a Friday) is in ISO week 53 of 2020
        assert_eq!(num("ISOWEEKNUM(DATE(2021,1,1))"), 53.0);
        assert_eq!(error("WEEKNUM(45000,4)"), NUM);
    }

    #[test]
    fn character_codes() {
        assert_eq!(text("CHAR(65)"), "A");
        assert_eq!(text("CHAR(241)"), "ñ");
        assert_eq!(text("CHAR(128)"), "€");
        assert_eq!(error("CHAR(0)"), VALUE);
        assert_eq!(num("CODE(\"A\")"), 65.0);
        assert_eq!(num("CODE(\"€uro\")"), 128.0);
        assert_eq!(num("CODE(\"ñ\")"), 241.0);
        assert_eq!(num("CODE(\"漢\")"), 63.0);
        assert_eq!(error("CODE(\"\")"), VALUE);
        assert_eq!(text("CLEAN(CHAR(9)&\"Informe\"&CHAR(10))"), "Informe");
    }

    #[test]
    fn finance() {
        near("PMT(0.08/12,10,10000)", -1037.032089, 1e-6);
        near("PMT(0.06/12,18*12,0,50000)", -129.0811609, 1e-6);
        assert_eq!(num("PMT(0,10,1000)"), -100.0);
        assert_eq!(error("PMT(0.1,0,1000)"), NUM);
        near("FV(0.06/12,10,-200,-500,1)", 2581.403374, 1e-6);
        assert_eq!(num("FV(0,10,-100)"), 1000.0);
        near("PV(0.08/12,12*20,500)", -59777.14585, 1e-5);
        near("NPER(0.12/12,-100,-1000,10000,1)", 59.67386567, 1e-7);
        // ln(10000 / 11000) / ln(1.01)
        near("NPER(0.01,-100,-1000)", -9.578594040, 1e-7);
        assert_eq!(num("NPER(0,-100,1000)"), 10.0);
        assert_eq!(error("NPER(0,0,1000)"), NUM);
    }

    #[test]
    fn unsupported_formulas_are_reported() {
        let ctx = sheet();
        let at = CellRef::parse("H2").unwrap();
        assert!(evaluate(&ctx, 0, at, &parse("OFFSET(A1,1,1)").unwrap()).is_err());
        assert!(evaluate(&ctx, 0, at, &parse("D1:D3*2").unwrap()).is_err());
        assert!(evaluate(&ctx, 0, at, &parse("ROUND(1)").unwrap()).is_err());
    }
}
