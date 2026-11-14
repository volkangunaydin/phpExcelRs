use ext_php_rs::{
    convert::FromZval,
    prelude::*,
    types::{ZendHashTable, Zval},
    zend::ClassEntry,
};
use rust_xlsxwriter::{ExcelDateTime, Format, Workbook, Worksheet};

use crate::err;

/// `ExcelRs\Formula` — marks a cell value as a formula: `new Formula('=SUM(A1:A9)')`.
///
/// Plain strings are never treated as formulas, so exported user data that
/// happens to start with `=` stays text.
#[php_class]
#[php(name = "ExcelRs\\Formula")]
pub struct Formula {
    expression: String,
}

#[php_impl]
impl Formula {
    pub fn __construct(expression: String) -> Self {
        Self { expression }
    }

    pub fn __to_string(&self) -> String {
        self.expression.clone()
    }
}

/// `ExcelRs\Writer` — builds an xlsx workbook in memory and saves it.
#[php_class]
#[php(name = "ExcelRs\\Writer")]
pub struct Writer {
    workbook: Workbook,
    /// Next free row per sheet, used by the append-style methods.
    next_row: Vec<u32>,
}

#[php_impl]
impl Writer {
    pub fn __construct() -> Self {
        Self {
            workbook: Workbook::new(),
            next_row: Vec::new(),
        }
    }

    /// Add a sheet and return its zero-based index.
    ///
    /// With `$lowMemory` the sheet's rows are flushed to a temporary file as
    /// they are written instead of being kept in memory until `save()`.
    #[php(defaults(low_memory = false))]
    pub fn add_sheet(&mut self, name: &str, low_memory: bool) -> PhpResult<i64> {
        let sheet = if low_memory {
            self.workbook.add_worksheet_with_low_memory().map_err(err)?
        } else {
            self.workbook.add_worksheet()
        };
        sheet.set_name(name).map_err(err)?;
        self.next_row.push(0);
        Ok(self.next_row.len() as i64 - 1)
    }

    /// Append one row to a sheet.
    ///
    /// Cells may be null, bool, int, float, string, `DateTimeInterface` or
    /// `ExcelRs\Formula`. `$style` applies to every cell of the row; keys:
    /// `bold`, `italic` (bool), `numberFormat` (e.g. `"#,##0.00"`; dates
    /// keep their date format), `color` and `background` (`"#RRGGBB"`).
    #[php(optional = style)]
    pub fn write_row(
        &mut self,
        sheet: i64,
        row: &ZendHashTable,
        style: Option<&ZendHashTable>,
    ) -> PhpResult<()> {
        let style = RowStyle::parse(style)?;
        let (worksheet, next_row) = self.sheet_mut(sheet)?;
        write_cells(worksheet, *next_row, row, &style)?;
        *next_row += 1;
        Ok(())
    }

    /// Append many rows (a list of lists) to a sheet. See `writeRow()`.
    #[php(optional = style)]
    pub fn write_rows(
        &mut self,
        sheet: i64,
        rows: &ZendHashTable,
        style: Option<&ZendHashTable>,
    ) -> PhpResult<()> {
        let style = RowStyle::parse(style)?;
        let (worksheet, next_row) = self.sheet_mut(sheet)?;
        for row in rows.values() {
            let cells = row
                .array()
                .ok_or_else(|| err(format!("Row {} is not an array", *next_row + 1)))?;
            write_cells(worksheet, *next_row, cells, &style)?;
            *next_row += 1;
        }
        Ok(())
    }

    /// Set the width of a column (zero-based) in character units.
    pub fn set_column_width(&mut self, sheet: i64, column: i64, width: f64) -> PhpResult<()> {
        let column = u16::try_from(column).map_err(|_| err("Column index is out of range"))?;
        let (worksheet, _) = self.sheet_mut(sheet)?;
        worksheet.set_column_width(column, width).map_err(err)?;
        Ok(())
    }

    /// Size every column to fit the data written so far.
    pub fn autofit(&mut self, sheet: i64) -> PhpResult<()> {
        let (worksheet, _) = self.sheet_mut(sheet)?;
        worksheet.autofit();
        Ok(())
    }

    /// Keep the first `$rows` rows and `$columns` columns visible while scrolling.
    #[php(defaults(columns = 0))]
    pub fn freeze_panes(&mut self, sheet: i64, rows: i64, columns: i64) -> PhpResult<()> {
        let rows = u32::try_from(rows).map_err(|_| err("Row count is out of range"))?;
        let columns = u16::try_from(columns).map_err(|_| err("Column count is out of range"))?;
        let (worksheet, _) = self.sheet_mut(sheet)?;
        worksheet.set_freeze_panes(rows, columns).map_err(err)?;
        Ok(())
    }

    /// Write the workbook to disk as xlsx.
    pub fn save(&mut self, path: &str) -> PhpResult<()> {
        if self.next_row.is_empty() {
            return Err(err("Workbook has no sheets; call addSheet() first"));
        }
        self.workbook
            .save(path)
            .map_err(|e| err(format!("{path}: {e}")))
    }
}

impl Writer {
    fn sheet_mut(&mut self, sheet: i64) -> PhpResult<(&mut Worksheet, &mut u32)> {
        let index = usize::try_from(sheet)
            .ok()
            .filter(|i| *i < self.next_row.len())
            .ok_or_else(|| err(format!("Sheet index {sheet} is out of range")))?;
        let worksheet = self.workbook.worksheet_from_index(index).map_err(err)?;
        Ok((worksheet, &mut self.next_row[index]))
    }
}

/// The formats used for one `writeRow()` / `writeRows()` call.
struct RowStyle {
    /// Whether the caller asked for any styling; empty cells are only
    /// written (as blanks) when they need to carry a style.
    styled: bool,
    base: Format,
    date: Format,
    datetime: Format,
}

impl RowStyle {
    fn parse(style: Option<&ZendHashTable>) -> PhpResult<Self> {
        let mut base = Format::new();
        let mut number_format = None;
        for (key, value) in style.into_iter().flat_map(|s| s.iter()) {
            let key = key.to_string();
            let text = || {
                value
                    .str()
                    .filter(|_| value.is_string())
                    .ok_or_else(|| err(format!("Style \"{key}\" must be a string")))
            };
            let flag = || {
                value
                    .bool()
                    .filter(|_| value.is_bool())
                    .ok_or_else(|| err(format!("Style \"{key}\" must be a bool")))
            };
            base = match key.as_str() {
                "bold" if flag()? => base.set_bold(),
                "italic" if flag()? => base.set_italic(),
                "bold" | "italic" => base,
                "numberFormat" => {
                    number_format = Some(text()?.to_owned());
                    base
                }
                "color" => base.set_font_color(parse_color(text()?)?),
                "background" => base.set_background_color(parse_color(text()?)?),
                _ => return Err(err(format!("Unknown style key \"{key}\""))),
            };
        }
        // The number format is for numbers; dates in the same row keep a
        // date format, otherwise they would show up as serial numbers.
        Ok(Self {
            styled: style.is_some_and(|s| !s.is_empty()),
            date: base.clone().set_num_format("yyyy-mm-dd"),
            datetime: base.clone().set_num_format("yyyy-mm-dd hh:mm:ss"),
            base: match number_format {
                Some(f) => base.set_num_format(f),
                None => base,
            },
        })
    }
}

fn parse_color(value: &str) -> PhpResult<u32> {
    let hex = value.strip_prefix('#').unwrap_or(value);
    u32::from_str_radix(hex, 16)
        .ok()
        .filter(|_| hex.len() == 6)
        .ok_or_else(|| err(format!("Invalid color \"{value}\"; expected \"#RRGGBB\"")))
}

fn write_cells(
    worksheet: &mut Worksheet,
    row: u32,
    cells: &ZendHashTable,
    style: &RowStyle,
) -> PhpResult<()> {
    for (i, cell) in cells.values().enumerate() {
        let col = u16::try_from(i).map_err(|_| err("Too many columns in row"))?;
        write_cell(worksheet, row, col, cell, style)?;
    }
    Ok(())
}

fn write_cell(
    worksheet: &mut Worksheet,
    row: u32,
    col: u16,
    cell: &Zval,
    style: &RowStyle,
) -> PhpResult<()> {
    let at = || format!("row {}, column {}", row + 1, col + 1);
    let format = &style.base;
    let result = if cell.is_null() {
        if !style.styled {
            return Ok(());
        }
        worksheet.write_blank(row, col, format)
    } else if let Some(v) = cell.long().filter(|_| cell.is_long()) {
        // Excel stores every number as a double.
        worksheet.write_number_with_format(row, col, v as f64, format)
    } else if let Some(v) = cell.double().filter(|_| cell.is_double()) {
        worksheet.write_number_with_format(row, col, v, format)
    } else if cell.is_bool() {
        worksheet.write_boolean_with_format(row, col, cell.is_true(), format)
    } else if cell.is_string() {
        let text = cell
            .str()
            .ok_or_else(|| err(format!("Cell at {} is not valid UTF-8", at())))?;
        worksheet.write_string_with_format(row, col, text, format)
    } else if let Some(formula) = <&Formula>::from_zval(cell) {
        worksheet.write_formula_with_format(row, col, formula.expression.as_str(), format)
    } else if let Some(text) = php_datetime(cell)? {
        let datetime = ExcelDateTime::parse_from_str(&text).map_err(err)?;
        let format = if text.ends_with(" 00:00:00") {
            &style.date
        } else {
            &style.datetime
        };
        worksheet.write_datetime_with_format(row, col, &datetime, format)
    } else {
        return Err(err(format!(
            "Unsupported value at {}: expected null, bool, int, float, string, \
             DateTimeInterface or ExcelRs\\Formula",
            at()
        )));
    };
    result.map(|_| ()).map_err(err)
}

/// The wall-clock time of a `DateTimeInterface` object as "Y-m-d H:i:s", or
/// None if the value is not one.
fn php_datetime(cell: &Zval) -> PhpResult<Option<String>> {
    let Some(object) = cell.object() else {
        return Ok(None);
    };
    let is_datetime = ClassEntry::try_find("DateTimeInterface")
        .is_some_and(|interface| object.instance_of(interface));
    if !is_datetime {
        return Ok(None);
    }
    let formatted = object
        .try_call_method("format", vec![&"Y-m-d H:i:s"])
        .map_err(err)?;
    Ok(formatted.string())
}
