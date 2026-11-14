<?php
// Smoke test: php -d extension=target/debug/libexcel_rs.so tests/roundtrip.php

use ExcelRs\Reader;
use ExcelRs\Writer;

function check(bool $ok, string $what): void
{
    echo ($ok ? 'ok   ' : 'FAIL ') . $what . PHP_EOL;
    if (!$ok) {
        exit(1);
    }
}

$path = sys_get_temp_dir() . '/excel_rs_' . getmypid() . '.xlsx';

$writer = new Writer();
$first = $writer->addSheet('Ürünler');
$writer->writeRow($first, ['Kod', 'Ad', 'Fiyat', 'Aktif']);
$writer->writeRows($first, [
    ['A-1', 'Çay', 12.5, true],
    ['A-2', 'Şeker', 40, false],
    ['A-3', null, 0.1, true],
]);
$second = $writer->addSheet('Boş');
$writer->save($path);

$reader = Reader::open($path);
check($reader->sheetNames() === ['Ürünler', 'Boş'], 'sheet names');

$rows = $reader->rows();
check(count($rows) === 4, 'row count');
check($rows[0] === ['Kod', 'Ad', 'Fiyat', 'Aktif'], 'header row');
check($rows[1] === ['A-1', 'Çay', 12.5, true], 'string/float/bool cells');
check($rows[2][2] == 40 && $rows[2][3] === false, 'int and false cells');
check($rows[3][1] === null, 'null cell');
check($reader->rows('Ürünler') === $rows, 'sheet by name');
check($reader->rows(0) === $rows, 'sheet by index');
check($reader->rows(1) === [], 'empty sheet');

// --- each(): streaming iteration ---------------------------------------
$seen = [];
$count = $reader->each(function (array $row, int $index) use (&$seen) {
    $seen[$index] = $row;
});
check($count === 4 && $seen === $rows, 'each() delivers the same rows as rows()');

$seen = 0;
$count = $reader->each(function () use (&$seen) {
    return ++$seen < 2;
}, 'Ürünler');
check($count === 2 && $seen === 2, 'each() stops when the callback returns false');
check($reader->each(fn () => null, 1) === 0, 'each() on an empty sheet');

try {
    $reader->each(fn () => throw new RuntimeException('boom'));
    check(false, 'callback exception should propagate');
} catch (RuntimeException $e) {
    check($e->getMessage() === 'boom', 'callback exception propagates unchanged');
}

// --- gaps: empty rows and leading empty cells keep their positions ------
$gapPath = $path . '.gap.xlsx';
$gap = new Writer();
$g = $gap->addSheet('Gap', true);
$gap->writeRow($g, []);
$gap->writeRow($g, [null, null, 'C2']);
$gap->writeRow($g, []);
$gap->writeRow($g, ['A4']);
$gap->save($gapPath);
$expected = [
    [null, null, null],
    [null, null, 'C2'],
    [null, null, null],
    ['A4', null, null],
];
$gapReader = new Reader($gapPath);
check($gapReader->rows() === $expected, 'rows() keeps gaps (low-memory sheet)');
$seen = [];
$gapReader->each(function (array $row) use (&$seen) {
    $seen[] = $row;
});
check($seen === $expected, 'each() keeps gaps');
unlink($gapPath);

// --- dates, formulas and styles ------------------------------------------
$stylePath = $path . '.style.xlsx';
$styled = new Writer();
$st = $styled->addSheet('Rapor');
$styled->writeRow($st, ['Tarih', 'Zaman', 'Tutar'], ['bold' => true, 'background' => '#DDEEFF', 'color' => '#003366']);
$styled->writeRows($st, [
    [new DateTimeImmutable('2026-11-14'), new DateTime('2026-11-14 09:30:15'), 1250.5],
    [new DateTimeImmutable('1999-12-31'), null, 10],
], ['numberFormat' => '#,##0.00']);
$styled->writeRow($st, ['=not a formula', null, new ExcelRs\Formula('=SUM(C2:C3)')], ['italic' => true]);
$styled->setColumnWidth($st, 0, 14);
$styled->freezePanes($st, 1);
$styled->autofit($st);
$styled->save($stylePath);

$styledRows = (new Reader($stylePath))->rows();
check($styledRows[1][0] === '2026-11-14', 'date cell reads back as Y-m-d');
check($styledRows[1][1] === '2026-11-14 09:30:15', 'datetime cell reads back as Y-m-d H:i:s');
check($styledRows[2][0] === '1999-12-31' && $styledRows[2][1] === null, 'date with custom number format');
check($styledRows[1][2] === 1250.5, 'number with number format keeps its value');
check($styledRows[3][0] === '=not a formula', 'string starting with = stays text');
check((string) new ExcelRs\Formula('=A1') === '=A1', 'Formula is stringable');
if (getenv('KEEP')) { copy($stylePath, getenv('KEEP')); }
unlink($stylePath);

foreach ([
    'missing sheet' => fn () => $reader->rows('yok'),
    'bad index' => fn () => $reader->rows(9),
    'missing sheet in each' => fn () => $reader->each(fn () => null, 'yok'),
    'missing file' => fn () => new Reader('/nonexistent/x.xlsx'),
    'unsupported cell' => fn () => $writer->writeRow($first, [new stdClass()]),
    'unknown style key' => fn () => $writer->writeRow($first, ['x'], ['font' => 'Arial']),
    'bad color' => fn () => $writer->writeRow($first, ['x'], ['color' => 'red']),
    'bad style type' => fn () => $writer->writeRow($first, ['x'], ['bold' => 'yes']),
    'bad sheet index on write' => fn () => $writer->writeRow(7, ['x']),
] as $what => $call) {
    try {
        $call();
        check(false, "$what should throw");
    } catch (ExcelRs\Exception $e) {
        check(true, "$what throws: " . $e->getMessage());
    }
}

unlink($path);
echo "all passed\n";
