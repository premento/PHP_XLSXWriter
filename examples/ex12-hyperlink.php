<?php
include_once __DIR__."/../xlsxwriter.class.php";

$rows = array(
  array("not clickable","also not clickable",'=HYPERLINK("http://github.com", "Click for Github.com")'),
);
$writer = new XLSXWriter();
foreach($rows as $row)
	$writer->writeSheetRow('Sheet1', $row);

$writer->writeToFile(__DIR__.'/xlsx-hyperlink.xlsx');
//$writer->writeToStdOut();
//echo $writer->writeToString();

