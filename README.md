PHP_XLSXWriter
==============

This library is designed to be lightweight, and have minimal memory usage.

It is designed to output an Excel compatible spreadsheet in (Office 2007+) xlsx format, with just basic features supported:
* supports PHP 5.4+ (CI runs the test suite on PHP 7.3 – 8.5)
* takes UTF-8 encoded input
* multiple worksheets
* supports currency/date/numeric cell formatting, simple formulas
* supports basic cell styling
* supports writing huge 100K+ row spreadsheets

[Never run out of memory with PHPExcel again](https://github.com/mk-j/PHP_XLSXWriter).

Simple PHP CLI example:
```php
$data = array(
    array('year','month','amount'),
    array('2003','1','220'),
    array('2003','2','153.5'),
);

$writer = new XLSXWriter();
$writer->writeSheet($data);
$writer->writeToFile('output.xlsx');
```

Simple/Advanced Cell Formats:
```php
$header = array(
  'created'=>'date',
  'product_id'=>'integer',
  'quantity'=>'#,##0',
  'amount'=>'price',
  'description'=>'string',
  'tax'=>'[$$-1009]#,##0.00;[RED]-[$$-1009]#,##0.00',
);
$data = array(
    array('2015-01-01',873,1,'44.00','misc','=D2*0.05'),
    array('2015-01-12',324,2,'88.00','none','=D3*0.05'),
);

$writer = new XLSXWriter();
$writer->writeSheetHeader('Sheet1', $header );
foreach($data as $row)
	$writer->writeSheetRow('Sheet1', $row );
$writer->writeToFile('example.xlsx');
```

50000 rows: (1.4s, 0MB memory usage)
```php
include_once("xlsxwriter.class.php");
$writer = new XLSXWriter();
$writer->writeSheetHeader('Sheet1', array('c1'=>'integer','c2'=>'integer','c3'=>'integer','c4'=>'integer') );
for($i=0; $i<50000; $i++)
{
    $writer->writeSheetRow('Sheet1', array($i, $i+1, $i+2, $i+3) );
}
$writer->writeToFile('huge.xlsx');
echo '#'.floor((memory_get_peak_usage())/1024/1024)."MB"."\n";
```
| rows   | time | memory |
| ------ | ---- | ------ |
|  50000 | 1.4s | 0MB    |
| 100000 | 2.7s | 0MB    |
| 150000 | 4.1s | 0MB    |
| 200000 | 5.7s | 0MB    |
| 250000 | 7.0s | 0MB    |

Simple cell formats map to more advanced cell formats

| simple formats | format code |
| ---------- | ---- |
| string   | @ |
| integer  | 0 |
| date     | YYYY-MM-DD |
| datetime | YYYY-MM-DD HH:MM:SS |
| time     | HH:MM:SS |
| price    | #,##0.00 |
| dollar   | [$$-1009]#,##0.00;[RED]-[$$-1009]#,##0.00 |
| euro     | #,##0.00 [$€-407];[RED]-#,##0.00 [$€-407] |


Basic cell styles have been available since version 0.30

| style        | allowed values |
| ------------ | ---- |
| font         | Arial, Times New Roman, Courier New, Comic Sans MS |
| font-size    | 8,9,10,11,12 ... |
| font-style   | bold, italic, underline, strikethrough or multiple ie: 'bold,italic' |
| border       | left, right, top, bottom,   or multiple ie: 'top,left' |
| border-style | thin, medium, thick, dashDot, dashDotDot, dashed, dotted, double, hair, mediumDashDot, mediumDashDotDot, mediumDashed, slantDashDot |
| border-color | #RRGGBB, ie: #ff99cc or #f9c |
| color        | #RRGGBB, ie: #ff99cc or #f9c |
| fill         | #RRGGBB, ie: #eeffee or #efe |
| halign       | general, left, right, justify, center, centerContinuous, distributed, fill |
| valign       | bottom, center, distributed, top, justify |
| wrap_text    | true, false |

Row options (`writeSheetRow()` 3rd argument) and column options (`writeSheetHeader()` 3rd argument)
can be mixed with the styles above:

| option         | where             | example |
| -------------- | ----------------- | ---- |
| height         | writeSheetRow     | `['height'=>30]` |
| hidden         | writeSheetRow     | `['hidden'=>true]` |
| collapsed      | writeSheetRow     | `['collapsed'=>true]` |
| widths         | writeSheetHeader  | `['widths'=>[10,20,30]]` |
| suppress_row   | writeSheetHeader  | `['suppress_row'=>true]` |
| auto_filter    | writeSheetHeader  | `['auto_filter'=>true]` |
| freeze_rows    | writeSheetHeader  | `['freeze_rows'=>1]` |
| freeze_columns | writeSheetHeader  | `['freeze_columns'=>1]` |

How values are written

| PHP value | result |
| --------- | ------ |
| `int`, `float` | number (`NAN`/`INF` are written as text) |
| numeric string in a `GENERAL` column | number, except leading-zero strings (`'00123'`) and integers of 16+ digits (card numbers, barcodes), which stay text |
| non-numeric string in a numeric column | text, rather than a corrupt cell |
| `true` / `false` | boolean cell |
| string starting with `=` | formula |
| `DateTime` / `DateTimeImmutable` | date/time (in a date/datetime/time column) |
| `null`, `''`, arrays | empty cell |

Sheet names are sanitized (`\ / ? * : [ ]` removed, max 31 characters) and made unique, so `'a/b'` and
`'a?b'` become `a b` and `a b (2)`.

Error handling

`writeToFile()`, `writeToStdOut()` return `true`/`false`, `writeToString()` returns `false` on failure.
Details go to `XLSXWriter::log()`, which calls `error_log()`. Subclass and override `log()` to send them elsewhere:

```php
class MyWriter extends XLSXWriter {
    public static function log($message) { MyLogger::warning($message); }
}
```

Running the tests

```
composer install
composer test
```

The `examples/` directory has a runnable script for each feature, e.g. `php examples/ex13-types-and-errors.php`.


