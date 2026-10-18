//! A1-style cell references: `B3` ↔ (row 3, column 2), both counted from 1

/// Last row and column of a worksheet (row 1 048 576, column XFD)
pub const MAX_ROW: u32 = 1_048_576;
pub const MAX_COL: u32 = 16_384;

/// A cell position, row and column counted from 1
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct CellRef {
    pub row: u32,
    pub col: u32,
}

impl CellRef {
    pub fn new(row: u32, col: u32) -> Result<CellRef, String> {
        if row == 0 || row > MAX_ROW || col == 0 || col > MAX_COL {
            return Err(format!("La celda (fila {}, columna {}) está fuera de la hoja.", row, col));
        }
        Ok(CellRef { row, col })
    }

    /// Parses `B3` (also `b3` and `$B$3`)
    pub fn parse(text: &str) -> Result<CellRef, String> {
        let invalid = || format!("'{}' no es una referencia de celda válida (por ejemplo, B3).", text);
        let clean: String = text.trim().chars().filter(|c| *c != '$').collect();
        let split = clean.find(|c: char| c.is_ascii_digit()).ok_or_else(invalid)?;
        let (letters, digits) = clean.split_at(split);
        if letters.is_empty() || letters.len() > 3 || !letters.chars().all(|c| c.is_ascii_alphabetic()) {
            return Err(invalid());
        }
        if !digits.chars().all(|c| c.is_ascii_digit()) || digits.starts_with('0') || digits.len() > 7 {
            return Err(invalid());
        }
        let col = letters.chars().fold(0u32, |acc, c| acc * 26 + (c.to_ascii_uppercase() as u32 - 'A' as u32 + 1));
        let row = digits.parse::<u32>().map_err(|_| invalid())?;
        CellRef::new(row, col).map_err(|_| invalid())
    }

    /// The cell `rows` below and `cols` to the right
    pub fn offset(self, rows: usize, cols: usize) -> Result<CellRef, String> {
        let row = self.row as usize + rows;
        let col = self.col as usize + cols;
        if row > MAX_ROW as usize || col > MAX_COL as usize {
            return Err(format!("Los valores no caben en la hoja a partir de {}.", self));
        }
        CellRef::new(row as u32, col as u32)
    }
}

impl std::fmt::Display for CellRef {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "{}{}", column_name(self.col), self.row)
    }
}

/// Column letters: 1 → A, 27 → AA
pub fn column_name(mut col: u32) -> String {
    let mut letters = Vec::new();
    while col > 0 {
        let rem = (col - 1) % 26;
        letters.push((b'A' + rem as u8) as char);
        col = (col - 1) / 26;
    }
    letters.iter().rev().collect()
}

/// A rectangle of cells, corners included
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CellRange {
    pub first: CellRef,
    pub last: CellRef,
}

impl CellRange {
    /// Parses `A1:C3` (corners in any order) or a single cell `B2`
    pub fn parse(text: &str) -> Result<CellRange, String> {
        let (a, b) = match text.split_once(':') {
            Some((a, b)) => (CellRef::parse(a)?, CellRef::parse(b)?),
            None => {
                let cell = CellRef::parse(text)?;
                (cell, cell)
            }
        };
        Ok(CellRange::around(a, b))
    }

    /// Smallest range holding both cells
    pub fn around(a: CellRef, b: CellRef) -> CellRange {
        CellRange {
            first: CellRef { row: a.row.min(b.row), col: a.col.min(b.col) },
            last: CellRef { row: a.row.max(b.row), col: a.col.max(b.col) },
        }
    }

    pub fn union(self, other: CellRange) -> CellRange {
        CellRange::around(
            CellRef { row: self.first.row.min(other.first.row), col: self.first.col.min(other.first.col) },
            CellRef { row: self.last.row.max(other.last.row), col: self.last.col.max(other.last.col) },
        )
    }

    pub fn cell_count(&self) -> u64 {
        (self.last.row - self.first.row + 1) as u64 * (self.last.col - self.first.col + 1) as u64
    }
}

impl std::fmt::Display for CellRange {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        if self.first == self.last {
            write!(f, "{}", self.first)
        } else {
            write!(f, "{}:{}", self.first, self.last)
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_and_prints_references() {
        assert_eq!(CellRef::parse("B3").unwrap(), CellRef { row: 3, col: 2 });
        assert_eq!(CellRef::parse("$aa$10").unwrap(), CellRef { row: 10, col: 27 });
        assert_eq!(CellRef::parse("XFD1048576").unwrap().to_string(), "XFD1048576");
        for bad in ["", "3", "B", "B0", "B03", "XFE1", "A1048577", "A-1", "1A"] {
            assert!(CellRef::parse(bad).is_err(), "{}", bad);
        }
        assert_eq!(column_name(1), "A");
        assert_eq!(column_name(26), "Z");
        assert_eq!(column_name(28), "AB");
        assert_eq!(column_name(703), "AAA");
    }

    #[test]
    fn parses_ranges_in_any_order() {
        let range = CellRange::parse("C3:A1").unwrap();
        assert_eq!(range.to_string(), "A1:C3");
        assert_eq!(range.cell_count(), 9);
        assert_eq!(CellRange::parse("B2").unwrap().to_string(), "B2");
        assert_eq!(range.union(CellRange::parse("D1").unwrap()).to_string(), "A1:D3");
    }
}
