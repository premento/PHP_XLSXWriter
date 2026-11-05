<?php
include_once __DIR__."/../xlsxwriter.class.php";

$header = array(
  'c1-text'=>'string',//text
  'c2-text'=>'@',//text
);
$rows = array(
  array('abcdefg','hijklmnop'),
);
$writer = new XLSXWriter();
$writer->setRightToLeft(true);

$writer->writeSheetHeader('Sheet1', $header);
foreach($rows as $row)
	$writer->writeSheetRow('Sheet1', $row);
$writer->writeToFile(__DIR__.'/xlsx-right-to-left.xlsx');

