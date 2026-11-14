# excel_rs

A PHP extension for reading and writing Excel files, written in Rust.

- **Read**: xlsx, xlsm, xlsb, xls, ods (via [calamine](https://crates.io/crates/calamine))
- **Write**: xlsx (via [rust_xlsxwriter](https://crates.io/crates/rust_xlsxwriter))
- Built with [ext-php-rs](https://crates.io/crates/ext-php-rs); PHP 8.0 – 8.5

## Usage

```php
use ExcelRs\Reader;
use ExcelRs\Writer;

$writer = new Writer();
$sheet = $writer->addSheet('Products');
$writer->writeRow($sheet, ['Code', 'Name', 'Price']);
$writer->writeRows($sheet, [
    ['A-1', 'Tea', 12.5],
    ['A-2', 'Sugar', 40],
]);
$writer->save('products.xlsx');

$reader = Reader::open('products.xlsx');
$reader->sheetNames();      // ['Products']
$reader->rows();            // first sheet, list of lists
$reader->rows('Products');  // by name
$reader->rows(0);           // by zero-based index
```

Cell values map to `null`, `bool`, `int`, `float` and `string`. Date cells are
read as `Y-m-d` or `Y-m-d H:i:s` strings. Every error is thrown as
`ExcelRs\Exception`.

### Dates, formulas and styles

```php
use ExcelRs\Formula;

$writer->writeRow($sheet, ['Date', 'Amount'], [
    'bold' => true,
    'background' => '#DDEEFF',
]);
$writer->writeRows($sheet, [
    [new DateTimeImmutable('2026-11-14'), 1250.5],
    [new DateTimeImmutable('2026-11-15'), 10],
], ['numberFormat' => '#,##0.00']);
$writer->writeRow($sheet, ['Total', new Formula('=SUM(B2:B3)')]);

$writer->setColumnWidth($sheet, 0, 14);  // column A
$writer->freezePanes($sheet, 1);         // keep the header row visible
$writer->autofit($sheet);
```

- `DateTimeInterface` values become real Excel dates (`Y-m-d`, or
  `Y-m-d H:i:s` when they carry a time).
- Only `Formula` objects are written as formulas. A plain string starting with
  `=` stays text, so exported user data cannot inject formulas. Formula results
  are calculated by the spreadsheet application when the file is opened;
  reading the file back before that yields `0`.
- The optional style array applies to every cell of the row(s): `bold`,
  `italic`, `numberFormat` (numbers only), `color`, `background`.

### Large files

`rows()` returns the whole sheet at once. For large sheets use `each()`, which
hands rows to a callback one at a time; xlsx/xlsm sheets are streamed from the
file. Return `false` from the callback to stop early.

```php
$count = $reader->each(function (array $row, int $index) {
    // ...
}, 'Products');
```

When writing, pass `true` as the second argument of `addSheet()` to flush rows
to a temporary file as they are written instead of holding them until `save()`:

```php
$sheet = $writer->addSheet('Export', true);
```

Measured with 300,000 rows × 7 columns (PHP 8.4, release build):

| Operation | Time | Peak memory |
|---|---|---|
| `rows()` | 2.0 s | 281 MB |
| `each()` | 1.9 s | 84 MB |
| write | 4.1 s | 441 MB |
| write, low memory | 3.5 s | 141 MB |

The remaining memory is mostly the workbook's shared string table, which grows
with the number of distinct strings.

## Building

Requires Rust, the PHP development headers (`php-dev`) and `clang`.

```sh
cargo build --release
php -d extension=$PWD/target/release/libexcel_rs.so tests/roundtrip.php
```

The binary is tied to the PHP version it was built against (and to NTS/ZTS).
To build for another installed version, point the build at its tools:

```sh
PHP=/usr/bin/php8.3 PHP_CONFIG=/usr/bin/php-config8.3 cargo build --release
```

To install, copy the `.so` into `php-config --extension-dir` and add
`extension=excel_rs` to `php.ini`.

## License

MIT
