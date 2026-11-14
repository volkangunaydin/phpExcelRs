use std::{fs::File, io::BufReader};

use calamine::{
    Data, DataRef, ExcelDateTime, Range, Reader as _, Sheets, Xlsx, open_workbook_auto,
};
use ext_php_rs::{
    boxed::ZBox,
    convert::IntoZval,
    prelude::*,
    types::{ZendCallable, ZendHashTable, Zval},
};

use crate::err;

/// `ExcelRs\Reader` — reads xlsx, xlsm, xlsb, xls and ods files.
#[php_class]
#[php(name = "ExcelRs\\Reader")]
pub struct Reader {
    workbook: Sheets<BufReader<File>>,
}

#[php_impl]
impl Reader {
    /// Open a workbook. The format is detected from the file extension.
    pub fn __construct(path: &str) -> PhpResult<Self> {
        let workbook = open_workbook_auto(path).map_err(|e| err(format!("{path}: {e}")))?;
        Ok(Self { workbook })
    }

    /// Static alias of the constructor: `Reader::open($path)`.
    pub fn open(path: &str) -> PhpResult<Self> {
        Self::__construct(path)
    }

    /// Names of all sheets, in workbook order.
    pub fn sheet_names(&self) -> Vec<String> {
        self.workbook.sheet_names()
    }

    /// All rows of a sheet as a list of lists.
    ///
    /// `$sheet` is a sheet name, a zero-based index, or null for the first
    /// sheet. Rows and columns start at A1, so array positions match cell
    /// positions; empty cells are null.
    #[php(optional = sheet)]
    pub fn rows(&mut self, sheet: Option<&Zval>) -> PhpResult<ZBox<ZendHashTable>> {
        let name = self.resolve_sheet(sheet)?;
        let range = self.workbook.worksheet_range(&name).map_err(err)?;
        let mut rows = ZendHashTable::new();
        range_rows(&range, |row| {
            rows.push(row).map_err(err)?;
            Ok(true)
        })?;
        Ok(rows)
    }

    /// Call `$callback(array $row, int $index)` for every row of a sheet and
    /// return the number of rows delivered. Returning `false` from the
    /// callback stops the iteration.
    ///
    /// For xlsx/xlsm files the sheet is streamed, so memory use does not grow
    /// with the number of rows. Other formats are loaded first.
    #[php(optional = sheet)]
    pub fn each(&mut self, callback: ZendCallable, sheet: Option<&Zval>) -> PhpResult<i64> {
        let name = self.resolve_sheet(sheet)?;
        let mut count = 0i64;
        let mut emit = |row: ZBox<ZendHashTable>| -> PhpResult<bool> {
            let row = row.into_zval(false).map_err(err)?;
            // On failure the PHP exception is already pending and propagates.
            let result = callback.try_call(vec![&row, &count]).map_err(err)?;
            count += 1;
            Ok(!result.is_false())
        };
        match &mut self.workbook {
            Sheets::Xlsx(xlsx) => stream_xlsx(xlsx, &name, &mut emit)?,
            other => {
                let range = other.worksheet_range(&name).map_err(err)?;
                range_rows(&range, &mut emit)?;
            }
        }
        Ok(count)
    }
}

impl Reader {
    fn resolve_sheet(&self, sheet: Option<&Zval>) -> PhpResult<String> {
        let names = self.workbook.sheet_names();
        let by_index = |index: i64| {
            usize::try_from(index)
                .ok()
                .and_then(|i| names.get(i).cloned())
                .ok_or_else(|| err(format!("Sheet index {index} is out of range")))
        };
        match sheet {
            None => by_index(0),
            Some(z) if z.is_null() => by_index(0),
            Some(z) if z.is_long() => by_index(z.long().unwrap_or(-1)),
            Some(z) if z.is_string() => {
                let name = z
                    .str()
                    .ok_or_else(|| err("Sheet name must be valid UTF-8"))?;
                if names.iter().any(|n| n == name) {
                    Ok(name.to_owned())
                } else {
                    Err(err(format!("Sheet \"{name}\" not found")))
                }
            }
            Some(_) => Err(err("Sheet must be a name (string), an index (int) or null")),
        }
    }
}

/// Deliver the rows of a loaded range, padded so that index 0 is A1.
/// `emit` returns false to stop.
fn range_rows(
    range: &Range<Data>,
    mut emit: impl FnMut(ZBox<ZendHashTable>) -> PhpResult<bool>,
) -> PhpResult<()> {
    let Some((_, end_col)) = range.end() else {
        return Ok(());
    };
    let (start_row, start_col) = range.start().unwrap_or((0, 0));
    let width = end_col + 1;

    for _ in 0..start_row {
        if !emit(null_row(width)?)? {
            return Ok(());
        }
    }
    for cells in range.rows() {
        let mut row = ZendHashTable::with_capacity(width);
        pad(&mut row, start_col)?;
        for cell in cells {
            row.push(data_to_zval(cell)?).map_err(err)?;
        }
        if !emit(row)? {
            return Ok(());
        }
    }
    Ok(())
}

/// Stream an xlsx sheet cell by cell, assembling one row at a time.
///
/// Rows are padded to the width the sheet declares; a row with cells beyond
/// that width is simply longer.
fn stream_xlsx(
    xlsx: &mut Xlsx<BufReader<File>>,
    name: &str,
    mut emit: impl FnMut(ZBox<ZendHashTable>) -> PhpResult<bool>,
) -> PhpResult<()> {
    let mut cells = xlsx.worksheet_cells_reader(name).map_err(err)?;
    let width = cells.dimensions().end.1 + 1;

    // Row being assembled: (row index, cells so far, next free column).
    let mut current: Option<(u32, ZBox<ZendHashTable>, u32)> = None;
    let mut next_row = 0u32;

    while let Some(cell) = cells.next_cell().map_err(err)? {
        let (r, c) = cell.get_position();
        if let Some((row_index, mut row, next_col)) = current.take_if(|(i, ..)| *i != r) {
            pad(&mut row, width.saturating_sub(next_col))?;
            if !emit(row)? {
                return Ok(());
            }
            next_row = row_index + 1;
        }
        if current.is_none() {
            if r < next_row {
                return Err(err("Sheet rows are not in ascending order"));
            }
            for _ in next_row..r {
                if !emit(null_row(width)?)? {
                    return Ok(());
                }
            }
            current = Some((r, ZendHashTable::with_capacity(width), 0));
        }
        if let Some((_, row, next_col)) = current.as_mut() {
            if c < *next_col {
                return Err(err("Sheet cells are not in ascending order"));
            }
            pad(row, c - *next_col)?;
            row.push(dataref_to_zval(cell.get_value())?).map_err(err)?;
            *next_col = c + 1;
        }
    }
    if let Some((_, mut row, next_col)) = current {
        pad(&mut row, width.saturating_sub(next_col))?;
        emit(row)?;
    }
    Ok(())
}

fn pad(row: &mut ZendHashTable, count: u32) -> PhpResult<()> {
    for _ in 0..count {
        row.push(()).map_err(err)?;
    }
    Ok(())
}

fn null_row(width: u32) -> PhpResult<ZBox<ZendHashTable>> {
    let mut row = ZendHashTable::with_capacity(width);
    pad(&mut row, width)?;
    Ok(row)
}

fn data_to_zval(cell: &Data) -> PhpResult<Zval> {
    match cell {
        Data::Empty => ().into_zval(false),
        Data::Int(v) => v.into_zval(false),
        Data::Float(v) => v.into_zval(false),
        Data::Bool(v) => v.into_zval(false),
        Data::String(v) | Data::DateTimeIso(v) | Data::DurationIso(v) => {
            v.as_str().into_zval(false)
        }
        Data::DateTime(v) => return datetime_to_zval(v),
        // Cell errors are returned as Excel shows them, e.g. "#DIV/0!".
        Data::Error(e) => e.to_string().into_zval(false),
    }
    .map_err(err)
}

fn dataref_to_zval(cell: &DataRef) -> PhpResult<Zval> {
    match cell {
        DataRef::Empty => ().into_zval(false),
        DataRef::Int(v) => v.into_zval(false),
        DataRef::Float(v) => v.into_zval(false),
        DataRef::Bool(v) => v.into_zval(false),
        DataRef::SharedString(v) => v.into_zval(false),
        DataRef::String(v) | DataRef::DateTimeIso(v) | DataRef::DurationIso(v) => {
            v.as_str().into_zval(false)
        }
        DataRef::DateTime(v) => return datetime_to_zval(v),
        DataRef::Error(e) => e.to_string().into_zval(false),
    }
    .map_err(err)
}

/// Dates become "Y-m-d" / "Y-m-d H:i:s" strings; durations stay numeric.
fn datetime_to_zval(value: &ExcelDateTime) -> PhpResult<Zval> {
    match value.as_datetime() {
        Some(dt) if value.is_datetime() => {
            let format = if dt.time() == chrono::NaiveTime::MIN {
                "%Y-%m-%d"
            } else {
                "%Y-%m-%d %H:%M:%S"
            };
            dt.format(format).to_string().into_zval(false)
        }
        _ => value.as_f64().into_zval(false),
    }
    .map_err(err)
}
