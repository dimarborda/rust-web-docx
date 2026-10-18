//! Dates in cells: Excel stores them as serial numbers (days since 1900 or 1904) shown through a
//! date number format. Times are the fraction of the day. Values cross to JavaScript as
//! milliseconds since 1970-01-01 in UTC, the wall-clock time Excel shows, with no time zone.

const MS_PER_DAY: f64 = 86_400_000.0;
/// Serial of 1970-01-01 in the 1900 system (valid from serial 61, 1900-03-01)
const UNIX_EPOCH_1900: f64 = 25_569.0;
/// Serial of 1970-01-01 in the 1904 system
const UNIX_EPOCH_1904: f64 = 24_107.0;
/// Days from 1970-01-01 back to 1900-03-01
const MARCH_1900: f64 = -25_508.0;

/// A serial date to milliseconds since 1970-01-01 UTC. In the 1900 system Excel counts a
/// 29 February 1900 that never existed (serial 60): earlier serials are shifted one day.
pub fn serial_to_unix_ms(serial: f64, date1904: bool) -> f64 {
    let days = if date1904 {
        serial - UNIX_EPOCH_1904
    } else if serial < 60.0 {
        serial - (UNIX_EPOCH_1900 - 1.0)
    } else {
        serial - UNIX_EPOCH_1900
    };
    (days * MS_PER_DAY).round()
}

/// Milliseconds since 1970-01-01 UTC to a serial date, or an error before the epoch of the system
pub fn unix_ms_to_serial(ms: f64, date1904: bool) -> Result<f64, String> {
    if !ms.is_finite() {
        return Err("La fecha no es válida.".to_string());
    }
    let days = ms / MS_PER_DAY;
    let serial = if date1904 {
        days + UNIX_EPOCH_1904
    } else if days < MARCH_1900 {
        days + UNIX_EPOCH_1900 - 1.0
    } else {
        days + UNIX_EPOCH_1900
    };
    if serial < if date1904 { 0.0 } else { 1.0 } {
        let first = if date1904 { "1904" } else { "1900" };
        return Err(format!("Excel no guarda fechas anteriores al 1 de enero de {} en este libro.", first));
    }
    // Rounded to the millisecond, as Excel shows it
    Ok((serial * MS_PER_DAY).round() / MS_PER_DAY)
}

/// Whether a time of day is part of the instant (so the cell needs a date-time format)
pub fn has_time(ms: f64) -> bool {
    ms.rem_euclid(MS_PER_DAY) != 0.0
}

/// An ISO 8601 date as written in `t="d"` cells (`2024-01-15`, `2024-01-15T10:30:00Z`) to
/// milliseconds since 1970-01-01 UTC; a time zone offset other than Z is ignored
pub fn iso_to_unix_ms(text: &str) -> Option<f64> {
    let text = text.trim();
    let (date, time) = match text.split_once('T') {
        Some((d, t)) => (d, Some(t)),
        None => (text, None),
    };
    let mut parts = date.split('-');
    let year: i64 = parts.next()?.parse().ok()?;
    let month: u32 = parts.next()?.parse().ok()?;
    let day: u32 = parts.next()?.parse().ok()?;
    if parts.next().is_some() || !(1..=12).contains(&month) || !(1..=31).contains(&day) {
        return None;
    }
    let mut ms = days_from_civil(year, month, day) as f64 * MS_PER_DAY;
    if let Some(time) = time {
        let time = time.trim_end_matches('Z');
        let time = time.split(['+', '-']).next()?;
        let mut fields = time.split(':');
        let h: f64 = fields.next()?.parse().ok()?;
        let m: f64 = fields.next().unwrap_or("0").parse().ok()?;
        let s: f64 = fields.next().unwrap_or("0").parse().ok()?;
        ms += ((h * 60.0 + m) * 60.0 + s) * 1000.0;
    }
    Some(ms.round())
}

/// Days since 1970-01-01 of a date in the proleptic Gregorian calendar (H. Hinnant's algorithm)
fn days_from_civil(year: i64, month: u32, day: u32) -> i64 {
    let y = if month <= 2 { year - 1 } else { year };
    let era = y.div_euclid(400);
    let yoe = y - era * 400;
    let mp = (month as i64 + 9) % 12;
    let doy = (153 * mp + 2) / 5 + day as i64 - 1;
    let doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
    era * 146_097 + doe - 719_468
}

/// Whether one of Excel's built-in number formats (by id) shows a date or a time
pub fn is_builtin_date_format(id: u32) -> bool {
    matches!(id, 14..=22 | 27..=36 | 45..=47 | 50..=58)
}

/// Built-in format ids used when writing dates
pub const DATE_FORMAT_ID: u32 = 14;
pub const DATE_TIME_FORMAT_ID: u32 = 22;

/// Whether a number format code shows a date or a time: it has d, m, y, h or s outside quoted
/// text, escaped characters and bracketed colors or conditions (`[h]` elapsed time counts)
pub fn is_date_format_code(code: &str) -> bool {
    if code.eq_ignore_ascii_case("general") {
        return false;
    }
    let mut chars = code.chars();
    while let Some(c) = chars.next() {
        match c {
            '"' => {
                for q in chars.by_ref() {
                    if q == '"' {
                        break;
                    }
                }
            }
            '\\' | '_' | '*' => {
                chars.next();
            }
            '[' => {
                let inner: String = chars.by_ref().take_while(|&b| b != ']').collect();
                let inner = inner.to_ascii_lowercase();
                if !inner.is_empty() && inner.chars().all(|b| matches!(b, 'h' | 'm' | 's')) {
                    return true;
                }
            }
            'd' | 'D' | 'm' | 'M' | 'y' | 'Y' | 'h' | 'H' | 's' | 'S' => return true,
            _ => {}
        }
    }
    false
}

#[cfg(test)]
mod tests {
    use super::*;

    const JAN_15_2024: f64 = 1_705_276_800_000.0;

    #[test]
    fn converts_serials_in_both_systems() {
        assert_eq!(serial_to_unix_ms(45306.0, false), JAN_15_2024);
        assert_eq!(serial_to_unix_ms(45306.0 - 1462.0, true), JAN_15_2024);
        assert_eq!(serial_to_unix_ms(25569.0, false), 0.0);
        assert_eq!(serial_to_unix_ms(45306.5, false), JAN_15_2024 + 43_200_000.0);
        // 1900-01-01 is serial 1; 1900-03-01 is 61 (60 is the phantom 29 February)
        assert_eq!(serial_to_unix_ms(1.0, false), days_from_civil(1900, 1, 1) as f64 * MS_PER_DAY);
        assert_eq!(serial_to_unix_ms(61.0, false), days_from_civil(1900, 3, 1) as f64 * MS_PER_DAY);

        assert_eq!(unix_ms_to_serial(JAN_15_2024, false).unwrap(), 45306.0);
        assert_eq!(unix_ms_to_serial(JAN_15_2024, true).unwrap(), 45306.0 - 1462.0);
        assert_eq!(unix_ms_to_serial(days_from_civil(1900, 1, 1) as f64 * MS_PER_DAY, false).unwrap(), 1.0);
        assert_eq!(unix_ms_to_serial(days_from_civil(1900, 3, 1) as f64 * MS_PER_DAY, false).unwrap(), 61.0);
        assert!(unix_ms_to_serial(days_from_civil(1899, 6, 1) as f64 * MS_PER_DAY, false).is_err());
        assert!(unix_ms_to_serial(days_from_civil(1903, 6, 1) as f64 * MS_PER_DAY, true).is_err());
        assert!(has_time(JAN_15_2024 + 1000.0) && !has_time(JAN_15_2024));
    }

    #[test]
    fn reads_iso_dates() {
        assert_eq!(iso_to_unix_ms("2024-01-15"), Some(JAN_15_2024));
        assert_eq!(iso_to_unix_ms("2024-01-15T12:00:00Z"), Some(JAN_15_2024 + 43_200_000.0));
        assert_eq!(iso_to_unix_ms("2024-13-01"), None);
        assert_eq!(iso_to_unix_ms("hola"), None);
    }

    #[test]
    fn recognises_date_formats() {
        for code in ["dd/mm/yyyy", "m/d/yy h:mm", "[h]:mm:ss", "yyyy\\-mm", "[$-es-ES]d \"de\" mmmm", "hh:mm AM/PM"] {
            assert!(is_date_format_code(code), "{}", code);
        }
        for code in ["General", "0.00", "#,##0 \"días\"", "[Red]0.0", "_-* #,##0_-", "0\\d", "@"] {
            assert!(!is_date_format_code(code), "{}", code);
        }
        assert!(is_builtin_date_format(14) && is_builtin_date_format(22) && !is_builtin_date_format(4));
    }
}
