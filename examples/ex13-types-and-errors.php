<?php
include_once __DIR__."/../xlsxwriter.class.php";

// Shows how different PHP values are written, and how to check for errors.
$header = array(
  'label'=>'string',
  'auto'=>'GENERAL',     //auto-detects numbers vs text
  'number'=>'0.00',
  'date'=>'date',
  'datetime'=>'datetime',
  'time'=>'time',
);
$rows = array(
  array('numbers',      '42',                 '3.14159',  '2024-02-29',             '2024-02-29 13:45:00',          '13:45:00'),
  array('leading zero', '00123',              '1e3',      43000,                    43000.5,                        0.25),//numbers in date columns are excel serials
  array('long id',      '4111111111111111',   'n/a',      new DateTime('2024-12-31'), new DateTime('2024-12-31 23:59:59'), '23:59:59'),//16+ digit ids stay text, non-numbers stay text
  array('booleans',     true,                 false,      'not a date',             '',                             null),
  array('formula',      '=B2*2',              '=SUM(C2:C3)', '=TODAY()',            '=NOW()',                       '=NOW()-TODAY()'),
  array('  padded  ',   "tab\tseparated",     '-0.5',     '1999-12-31',             '1999-12-31 00:00:00',          '00:00:01'),
);

$writer = new XLSXWriter();
$writer->writeSheetHeader('Types', $header, array('widths'=>array(14,20,12,12,20,10), 'font-style'=>'bold', 'border'=>'bottom'));
foreach($rows as $row)
	$writer->writeSheetRow('Types', $row);

// Sheet names are sanitized and made unique: these become "a b", "a b (2)" and "History_"
$writer->writeSheetRow('a/b', array('first'));
$writer->writeSheetRow('a?b', array('second'));
$writer->writeSheetRow('History', array('"History" is reserved by Excel'));

// writeToFile() returns false on failure; the reason is sent to XLSXWriter::log() (error_log by default)
$filename = __DIR__.'/xlsx-types-and-errors.xlsx';
if (!$writer->writeToFile($filename)) {
	fwrite(STDERR, "Could not write $filename\n");
	exit(1);
}
echo "Wrote $filename\n";
