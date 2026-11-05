<?php
/*
 * @license MIT License
 * */

class XLSXWriter
{
	//http://www.ecma-international.org/publications/standards/Ecma-376.htm
	//http://officeopenxml.com/SSstyles.php
	//------------------------------------------------------------------
	//http://office.microsoft.com/en-us/excel-help/excel-specifications-and-limits-HP010073849.aspx
	const EXCEL_2007_MAX_ROW=1048576;
	const EXCEL_2007_MAX_COL=16384;
	const EXCEL_MAX_CELL_CHARS=32767;
	const EXCEL_MAX_SHEETNAME_CHARS=31;
	//------------------------------------------------------------------
	protected $title;
	protected $subject;
	protected $author;
	protected $isRightToLeft;
	protected $company;
	protected $description;
	protected $keywords = array();
	protected $tempdir;

	protected $current_sheet;
	protected $sheets = array();
	protected $temp_files = array();
	protected $cell_styles = array();
	protected $cell_styles_lookup = array();//cell style string => index, avoids a linear search per styled cell
	protected $number_formats = array();
	protected $number_formats_lookup = array();//number format => index

	//row/column options that are not cell styles, these are stripped before a style is registered
	protected static $non_style_options = array(
		'height'=>1, 'hidden'=>1, 'collapsed'=>1, 'widths'=>1, 'suppress_row'=>1,
		'auto_filter'=>1, 'freeze_rows'=>1, 'freeze_columns'=>1,
	);

	public function __construct()
	{
		defined('ENT_XML1') or define('ENT_XML1',16);//for php 5.3, avoid fatal error
		is_writeable($this->tempFilename()) or static::log("Warning: tempdir ".sys_get_temp_dir()." not writeable, use ->setTempDir()");
		class_exists('ZipArchive') or static::log("Error: ZipArchive class does not exist");
		$this->addCellStyle('GENERAL', null);
	}

	public function setTitle($title='') { $this->title=$title; }
	public function setSubject($subject='') { $this->subject=$subject; }
	public function setAuthor($author='') { $this->author=$author; }
	public function setCompany($company='') { $this->company=$company; }
	public function setKeywords($keywords='') { $this->keywords=$keywords; }
	public function setDescription($description='') { $this->description=$description; }
	public function setTempDir($tempdir='') { $this->tempdir=$tempdir; }
	public function setRightToLeft($isRightToLeft=false){ $this->isRightToLeft=$isRightToLeft; }

	public function __destruct()
	{
		//close any still-open sheet files first, an open handle cannot be unlinked on Windows
		foreach($this->sheets as $sheet) {
			if (isset($sheet->file_writer)) {
				$sheet->file_writer->close();
			}
		}
		foreach($this->temp_files as $temp_file) {
			@unlink($temp_file);
		}
	}

	protected function tempFilename()
	{
		$tempdir = !empty($this->tempdir) ? $this->tempdir : sys_get_temp_dir();
		$filename = tempnam($tempdir, "xlsx_writer_");
		if (!$filename) {
			// If you are seeing this error, it's possible you may have too many open
			// file handles. If you're creating a spreadsheet with many small inserts,
			// it is possible to exceed the default 1024 open file handles. Run 'ulimit -a'
			// and try increasing the 'open files' number with 'ulimit -n 8192'
			throw new \Exception("Unable to create tempfile - check file handle limits?");
		}
		$this->temp_files[] = $filename;
		return $filename;
	}

	public function writeToStdOut()
	{
		$temp_file = $this->tempFilename();
		if (!$this->writeToFile($temp_file)) {
			return false;
		}
		readfile($temp_file);
		return true;
	}

	public function writeToString()
	{
		$temp_file = $this->tempFilename();
		if (!$this->writeToFile($temp_file)) {
			return false;
		}
		return file_get_contents($temp_file);
	}

	/**
	 * @return bool true on success, false on failure (details are sent to the error log)
	 */
	public function writeToFile($filename)
	{
		if (empty($this->sheets)) {
			static::log("Error in ".__CLASS__."::".__FUNCTION__.", no worksheets defined.");
			return false;
		}
		if (!class_exists('ZipArchive')) {
			static::log("Error in ".__CLASS__."::".__FUNCTION__.", ZipArchive class does not exist.");
			return false;
		}

		foreach($this->sheets as $sheet_name => $sheet) {
			$this->finalizeSheet($sheet_name);//making sure all footers have been written
		}

		if (file_exists($filename)) {
			if (is_writable($filename)) {
				@unlink($filename); //if the zip already exists, remove it
			} else {
				static::log("Error in " . __CLASS__ . "::" . __FUNCTION__ . ", file is not writeable.");
				return false;
			}
		}
		$directory = dirname($filename);
		if (!is_dir($directory) || !is_writable($directory)) {
			//ZipArchive::open() succeeds here and only close() fails, check up front instead
			static::log("Error in " . __CLASS__ . "::" . __FUNCTION__ . ", unable to create zip, directory $directory is not writeable.");
			return false;
		}

		$zip = new ZipArchive();
		//ZipArchive::open() returns true or an int error code (which is truthy), so compare strictly
		if ($zip->open($filename, ZipArchive::CREATE | ZipArchive::OVERWRITE) !== true) {
			static::log("Error in ".__CLASS__."::".__FUNCTION__.", unable to create zip.");
			return false;
		}

		$zip->addEmptyDir("docProps/");
		$zip->addFromString("docProps/app.xml" , $this->buildAppXML() );
		$zip->addFromString("docProps/core.xml", $this->buildCoreXML());

		$zip->addEmptyDir("_rels/");
		$zip->addFromString("_rels/.rels", $this->buildRelationshipsXML());

		$zip->addEmptyDir("xl/worksheets/");
		foreach($this->sheets as $sheet) {
			$zip->addFile($sheet->filename, "xl/worksheets/".$sheet->xmlname );
		}
		$zip->addFromString("xl/workbook.xml"         , $this->buildWorkbookXML() );
		$zip->addFile($this->writeStylesXML(), "xl/styles.xml" );
		$zip->addFromString("[Content_Types].xml"     , $this->buildContentTypesXML() );

		$zip->addEmptyDir("xl/_rels/");
		$zip->addFromString("xl/_rels/workbook.xml.rels", $this->buildWorkbookRelsXML() );
		if (!@$zip->close()) {//failure is reported through log() below
			static::log("Error in ".__CLASS__."::".__FUNCTION__.", unable to write zip.");
			return false;
		}
		return true;
	}

	protected static function isValidSheetName($sheet_name)
	{
		//note: '0' is a valid sheet name, so empty() can not be used here
		return is_scalar($sheet_name) && !is_bool($sheet_name) && strlen((string)$sheet_name) > 0;
	}

	protected function initializeSheet($sheet_name, $col_widths=array(), $auto_filter=false, $freeze_rows=false, $freeze_columns=false )
	{
		//if already initialized
		if (isset($this->sheets[$sheet_name]))
			return;

		$sheet_filename = $this->tempFilename();
		$sheet_xmlname = 'sheet' . (count($this->sheets) + 1).".xml";
		$this->sheets[$sheet_name] = (object)array(
			'filename' => $sheet_filename,
			'sheetname' => $sheet_name,
			'safe_sheetname' => $this->uniqueSheetName($sheet_name),
			'xmlname' => $sheet_xmlname,
			'row_count' => 0,
			'file_writer' => new XLSXWriter_BuffererWriter($sheet_filename),
			'columns' => array(),
			'merge_cells' => array(),
			'max_cell_tag_start' => 0,
			'max_cell_tag_end' => 0,
			'auto_filter' => $auto_filter,
			'freeze_rows' => $freeze_rows,
			'freeze_columns' => $freeze_columns,
			'finalized' => false,
			'overflow_logged' => false,
		);
		$rightToLeftValue = $this->isRightToLeft ? 'true' : 'false';
		$sheet = $this->sheets[$sheet_name];
		$tabselected = count($this->sheets) == 1 ? 'true' : 'false';//only first sheet is selected
		$max_cell=XLSXWriter::xlsCell(self::EXCEL_2007_MAX_ROW, self::EXCEL_2007_MAX_COL);//XFE1048577
		$sheet->file_writer->write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\n");
		$sheet->file_writer->write('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">');
		$sheet->file_writer->write(  '<sheetPr filterMode="false">');
		$sheet->file_writer->write(    '<pageSetUpPr fitToPage="false"/>');
		$sheet->file_writer->write(  '</sheetPr>');
		$sheet->max_cell_tag_start = $sheet->file_writer->ftell();
		$sheet->file_writer->write('<dimension ref="A1:' . $max_cell . '"/>');
		$sheet->max_cell_tag_end = $sheet->file_writer->ftell();
		$sheet->file_writer->write(  '<sheetViews>');
		$sheet->file_writer->write(    '<sheetView colorId="64" defaultGridColor="true" rightToLeft="'.$rightToLeftValue.'" showFormulas="false" showGridLines="true" showOutlineSymbols="true" showRowColHeaders="true" showZeros="true" tabSelected="' . $tabselected . '" topLeftCell="A1" view="normal" windowProtection="false" workbookViewId="0" zoomScale="100" zoomScaleNormal="100" zoomScalePageLayoutView="100">');
		if ($sheet->freeze_rows && $sheet->freeze_columns) {
			$top_right   = self::xlsCell(0, $sheet->freeze_columns);
			$bottom_left = self::xlsCell($sheet->freeze_rows, 0);
			$bottom_right= self::xlsCell($sheet->freeze_rows, $sheet->freeze_columns);
			$sheet->file_writer->write(      '<pane ySplit="'.$sheet->freeze_rows.'" xSplit="'.$sheet->freeze_columns.'" topLeftCell="'.$bottom_right.'" activePane="bottomRight" state="frozen"/>');
			$sheet->file_writer->write(      '<selection activeCell="'.$top_right.'" activeCellId="0" pane="topRight" sqref="'.$top_right.'"/>');
			$sheet->file_writer->write(      '<selection activeCell="'.$bottom_left.'" activeCellId="0" pane="bottomLeft" sqref="'.$bottom_left.'"/>');
			$sheet->file_writer->write(      '<selection activeCell="'.$bottom_right.'" activeCellId="0" pane="bottomRight" sqref="'.$bottom_right.'"/>');
		}
		elseif ($sheet->freeze_rows) {
			$sheet->file_writer->write(      '<pane ySplit="'.$sheet->freeze_rows.'" topLeftCell="'.self::xlsCell($sheet->freeze_rows, 0).'" activePane="bottomLeft" state="frozen"/>');
			$sheet->file_writer->write(      '<selection activeCell="'.self::xlsCell($sheet->freeze_rows, 0).'" activeCellId="0" pane="bottomLeft" sqref="'.self::xlsCell($sheet->freeze_rows, 0).'"/>');
		}
		elseif ($sheet->freeze_columns) {
			$sheet->file_writer->write(      '<pane xSplit="'.$sheet->freeze_columns.'" topLeftCell="'.self::xlsCell(0, $sheet->freeze_columns).'" activePane="topRight" state="frozen"/>');
			$sheet->file_writer->write(      '<selection activeCell="'.self::xlsCell(0, $sheet->freeze_columns).'" activeCellId="0" pane="topRight" sqref="'.self::xlsCell(0, $sheet->freeze_columns).'"/>');
		}
		else { // not frozen
			$sheet->file_writer->write(      '<selection activeCell="A1" activeCellId="0" pane="topLeft" sqref="A1"/>');
		}
		$sheet->file_writer->write(    '</sheetView>');
		$sheet->file_writer->write(  '</sheetViews>');
		$sheet->file_writer->write(  '<cols>');
		$i=0;
		if (!empty($col_widths)) {
			foreach($col_widths as $column_width) {
				$sheet->file_writer->write(  '<col collapsed="false" hidden="false" max="'.($i+1).'" min="'.($i+1).'" style="0" customWidth="true" width="'.floatval($column_width).'"/>');
				$i++;
			}
		}
		if ($i < 1024) {
			$sheet->file_writer->write(  '<col collapsed="false" hidden="false" max="1024" min="'.($i+1).'" style="0" customWidth="false" width="11.5"/>');
		}
		$sheet->file_writer->write(  '</cols>');
		$sheet->file_writer->write(  '<sheetData>');
	}

	/**
	 * Excel requires sheet names to be unique (case-insensitively) and at most 31 chars.
	 * The safe name is computed once per sheet, so workbook.xml and definedNames always agree.
	 */
	protected function uniqueSheetName($sheet_name)
	{
		$base = self::sanitize_sheetname($sheet_name);
		$taken = array();
		foreach($this->sheets as $sheet) {
			if (isset($sheet->safe_sheetname)) {
				$taken[self::lowercase($sheet->safe_sheetname)] = true;
			}
		}
		$candidate = $base;
		for($n=2; isset($taken[self::lowercase($candidate)]); $n++) {
			$suffix = " ($n)";
			$candidate = self::substring($base, 0, self::EXCEL_MAX_SHEETNAME_CHARS - strlen($suffix)) . $suffix;
		}
		return $candidate;
	}

	private function addCellStyle($number_format, $cell_style_string)
	{
		if (isset($this->number_formats_lookup[$number_format])) {
			$number_format_idx = $this->number_formats_lookup[$number_format];
		} else {
			$number_format_idx = count($this->number_formats);
			$this->number_formats[] = $number_format;
			$this->number_formats_lookup[$number_format] = $number_format_idx;
		}
		$lookup_string = $number_format_idx.";".$cell_style_string;
		if (isset($this->cell_styles_lookup[$lookup_string])) {
			return $this->cell_styles_lookup[$lookup_string];
		}
		$cell_style_idx = count($this->cell_styles);
		$this->cell_styles[] = $lookup_string;
		$this->cell_styles_lookup[$lookup_string] = $cell_style_idx;
		return $cell_style_idx;
	}

	private function initializeColumnTypes($header_types)
	{
		$column_types = array();
		foreach($header_types as $v)
		{
			$number_format = self::numberFormatStandardized($v);
			$number_format_type = self::determineNumberFormatType($number_format);
			$cell_style_idx = $this->addCellStyle($number_format, null);
			$column_types[] = array('number_format' => $number_format,//contains excel format like 'YYYY-MM-DD HH:MM:SS'
									'number_format_type' => $number_format_type, //contains friendly format like 'datetime'
									'default_cell_style' => $cell_style_idx,
									);
		}
		return $column_types;
	}

	/**
	 * Strips row/column options (height, widths...) so they do not create needless cell styles.
	 * @return array|null array($style, $is_per_column_list) or null when there is no style
	 */
	protected static function extractStyle($options)
	{
		if (empty($options) || !is_array($options)) {
			return null;
		}
		if (array_key_exists(0, $options)) {
			return array($options, true);//list of styles, one per column
		}
		$style = array_diff_key($options, self::$non_style_options);
		return empty($style) ? null : array($style, false);
	}

	protected function resolveCellStyles(array $columns, $style, $num_cells, $number_format=null)
	{
		$styles = array();
		if ($style === null) {
			for($c=0; $c<$num_cells; $c++) {
				$styles[$c] = $columns[$c]['default_cell_style'];
			}
			return $styles;
		}
		list($style, $is_list) = $style;
		$shared_json = $is_list ? null : json_encode($style);//encode once per row, not per cell
		for($c=0; $c<$num_cells; $c++) {
			$format = $number_format !== null ? $number_format : $columns[$c]['number_format'];
			if ($is_list) {
				$cell_style = isset($style[$c]) ? $style[$c] : null;
				$styles[$c] = empty($cell_style) ? $columns[$c]['default_cell_style'] : $this->addCellStyle($format, json_encode($cell_style));
			} else {
				$styles[$c] = $this->addCellStyle($format, $shared_json);
			}
		}
		return $styles;
	}

	public function writeSheetHeader($sheet_name, array $header_types, $col_options = null)
	{
		if (!self::isValidSheetName($sheet_name) || empty($header_types))
			return;

		if (is_bool($col_options))
		{
			static::log( 'Warning! passing $suppress_row=false|true to writeSheetHeader() is deprecated, this will be removed in a future version.' );
			$col_options = array('suppress_row' => $col_options);
		}
		$col_options = is_array($col_options) ? $col_options : array();
		if (isset($this->sheets[$sheet_name]) && $this->sheets[$sheet_name]->finalized) {
			static::log("Error in ".__CLASS__."::".__FUNCTION__.", sheet '$sheet_name' has already been written, header ignored.");
			return;
		}

		$suppress_row = !empty($col_options['suppress_row']);
		$col_widths = isset($col_options['widths']) ? (array)$col_options['widths'] : array();
		$auto_filter = isset($col_options['auto_filter']) ? intval($col_options['auto_filter']) : false;
		$freeze_rows = isset($col_options['freeze_rows']) ? intval($col_options['freeze_rows']) : false;
		$freeze_columns = isset($col_options['freeze_columns']) ? intval($col_options['freeze_columns']) : false;
		$this->initializeSheet($sheet_name, $col_widths, $auto_filter, $freeze_rows, $freeze_columns);
		$sheet = $this->sheets[$sheet_name];
		$sheet->columns = $this->initializeColumnTypes($header_types);
		if (!$suppress_row)
		{
			$header_row = array_keys($header_types);
			$styles = $this->resolveCellStyles($sheet->columns, self::extractStyle($col_options), count($header_row), 'GENERAL');

			$sheet->file_writer->write('<row collapsed="false" customFormat="false" customHeight="false" hidden="false" ht="12.1" outlineLevel="0" r="' . ($sheet->row_count+1) . '">');
			foreach ($header_row as $c => $v) {
				$this->writeCell($sheet->file_writer, $sheet->row_count, $c, $v, 'n_string', $styles[$c]);
			}
			$sheet->file_writer->write('</row>');
			$sheet->row_count++;
		}
		$this->current_sheet = $sheet_name;
	}

	public function writeSheetRow($sheet_name, array $row, $row_options=null)
	{
		if (!self::isValidSheetName($sheet_name))
			return;

		$this->initializeSheet($sheet_name);
		$sheet = $this->sheets[$sheet_name];
		if ($sheet->finalized) {
			static::log("Error in ".__CLASS__."::".__FUNCTION__.", sheet '$sheet_name' has already been written, row ignored.");
			return;
		}
		if ($sheet->row_count >= self::EXCEL_2007_MAX_ROW) {
			if (!$sheet->overflow_logged) {
				static::log("Error in ".__CLASS__."::".__FUNCTION__.", sheet '$sheet_name' exceeds ".self::EXCEL_2007_MAX_ROW." rows, extra rows ignored.");
				$sheet->overflow_logged = true;
			}
			return;
		}
		$num_cells = count($row);
		$num_columns = count($sheet->columns);
		if ($num_columns < $num_cells) {
			//only add the missing columns, otherwise the sheet dimension/autofilter range grows too wide
			$default_column_types = $this->initializeColumnTypes( array_fill(0, $num_cells - $num_columns, 'GENERAL') );//will map to n_auto
			$sheet->columns = array_merge($sheet->columns, $default_column_types);
		}

		if (!empty($row_options) && is_array($row_options))
		{
			$customHt = isset($row_options['height']);
			$ht = $customHt ? floatval($row_options['height']) : 12.1;
			$hidden = !empty($row_options['hidden']);
			$collapsed = !empty($row_options['collapsed']);
			$sheet->file_writer->write('<row collapsed="'.($collapsed ? 'true' : 'false').'" customFormat="false" customHeight="'.($customHt ? 'true' : 'false').'" hidden="'.($hidden ? 'true' : 'false').'" ht="'.($ht).'" outlineLevel="0" r="' . ($sheet->row_count + 1) . '">');
		}
		else
		{
			$sheet->file_writer->write('<row collapsed="false" customFormat="false" customHeight="false" hidden="false" ht="12.1" outlineLevel="0" r="' . ($sheet->row_count + 1) . '">');
		}

		$styles = $this->resolveCellStyles($sheet->columns, self::extractStyle($row_options), $num_cells);
		$c=0;
		foreach ($row as $v) {
			$this->writeCell($sheet->file_writer, $sheet->row_count, $c, $v, $sheet->columns[$c]['number_format_type'], $styles[$c]);
			$c++;
		}
		$sheet->file_writer->write('</row>');
		$sheet->row_count++;
		$this->current_sheet = $sheet_name;
	}

	public function countSheetRows($sheet_name = '')
	{
		$sheet_name = self::isValidSheetName($sheet_name) ? $sheet_name : $this->current_sheet;
		return self::isValidSheetName($sheet_name) && isset($this->sheets[$sheet_name]) ? $this->sheets[$sheet_name]->row_count : 0;
	}

	protected function finalizeSheet($sheet_name)
	{
		if (!self::isValidSheetName($sheet_name) || !isset($this->sheets[$sheet_name]) || $this->sheets[$sheet_name]->finalized)
			return;

		$sheet = $this->sheets[$sheet_name];

		$sheet->file_writer->write(    '</sheetData>');

		if (!empty($sheet->merge_cells)) {
			$sheet->file_writer->write(    '<mergeCells>');
			foreach ($sheet->merge_cells as $range) {
				$sheet->file_writer->write(        '<mergeCell ref="' . $range . '"/>');
			}
			$sheet->file_writer->write(    '</mergeCells>');
		}

		$max_cell = self::xlsCell(max(0, $sheet->row_count - 1), max(0, count($sheet->columns) - 1));

		if ($sheet->auto_filter) {
			$sheet->file_writer->write(    '<autoFilter ref="A1:' . $max_cell . '"/>');
		}

		$sheet->file_writer->write(    '<printOptions headings="false" gridLines="false" gridLinesSet="true" horizontalCentered="false" verticalCentered="false"/>');
		$sheet->file_writer->write(    '<pageMargins left="0.5" right="0.5" top="1.0" bottom="1.0" header="0.5" footer="0.5"/>');
		$sheet->file_writer->write(    '<pageSetup blackAndWhite="false" cellComments="none" copies="1" draft="false" firstPageNumber="1" fitToHeight="1" fitToWidth="1" horizontalDpi="300" orientation="portrait" pageOrder="downThenOver" paperSize="1" scale="100" useFirstPageNumber="true" usePrinterDefaults="false" verticalDpi="300"/>');
		$sheet->file_writer->write(    '<headerFooter differentFirst="false" differentOddEven="false">');
		$sheet->file_writer->write(        '<oddHeader>&amp;C&amp;&quot;Times New Roman,Regular&quot;&amp;12&amp;A</oddHeader>');
		$sheet->file_writer->write(        '<oddFooter>&amp;C&amp;&quot;Times New Roman,Regular&quot;&amp;12Page &amp;P</oddFooter>');
		$sheet->file_writer->write(    '</headerFooter>');
		$sheet->file_writer->write('</worksheet>');

		$max_cell_tag = '<dimension ref="A1:' . $max_cell . '"/>';
		$padding_length = $sheet->max_cell_tag_end - $sheet->max_cell_tag_start - strlen($max_cell_tag);
		$sheet->file_writer->fseek($sheet->max_cell_tag_start);
		$sheet->file_writer->write($max_cell_tag.str_repeat(" ", $padding_length));
		$sheet->file_writer->close();
		$sheet->finalized=true;
	}

	public function markMergedCell($sheet_name, $start_cell_row, $start_cell_column, $end_cell_row, $end_cell_column)
	{
		if (!self::isValidSheetName($sheet_name))
			return;

		$this->initializeSheet($sheet_name);
		$sheet = $this->sheets[$sheet_name];
		if ($sheet->finalized)
			return;

		$startCell = self::xlsCell($start_cell_row, $start_cell_column);
		$endCell = self::xlsCell($end_cell_row, $end_cell_column);
		$sheet->merge_cells[] = $startCell . ":" . $endCell;
	}

	public function writeSheet(array $data, $sheet_name='', array $header_types=array())
	{
		$sheet_name = self::isValidSheetName($sheet_name) ? $sheet_name : 'Sheet1';
		$data = empty($data) ? array(array('')) : $data;
		if (!empty($header_types))
		{
			$this->writeSheetHeader($sheet_name, $header_types);
		}
		foreach($data as $row)
		{
			$this->writeSheetRow($sheet_name, (array)$row);
		}
		$this->finalizeSheet($sheet_name);
	}

	protected function writeCell(XLSXWriter_BuffererWriter &$file, $row_number, $column_number, $value, $num_format_type, $cell_style_idx)
	{
		$cell_name = self::xlsCell($row_number, $column_number);
		$cell_attrs = 'r="'.$cell_name.'" s="'.$cell_style_idx.'"';

		if ($value instanceof \DateTimeInterface || $value instanceof \DateTime) {
			$value = $value->format('Y-m-d H:i:s');
		}

		if (is_bool($value)) {
			$file->write('<c '.$cell_attrs.' t="b"><v>'.($value ? 1 : 0).'</v></c>');
		} elseif (!is_scalar($value) || $value==='') { //objects, array, null, empty
			$file->write('<c '.$cell_attrs.'/>');
		} elseif (is_string($value) && $value[0]=='=') {
			//no t="s" here: that type means "index into sharedStrings.xml", which this writer never creates
			$file->write('<c '.$cell_attrs.'><f>'.self::xmlspecialchars(substr($value, 1)).'</f></c>');
		} elseif ($num_format_type=='n_date' || $num_format_type=='n_datetime') {
			if (is_int($value) || is_float($value)) {
				$serial = $value;//already an excel serial date
			} elseif (preg_match('/\d{4}-\d{2}-\d{2}|\d+:\d{2}:\d{2}/', $value)) {
				$serial = self::convert_date_time($value);
			} else {
				$file->write(self::inlineStringCell($cell_attrs, $value));//unparseable, keep the data rather than writing 0
				return;
			}
			$file->write('<c '.$cell_attrs.' t="n"><v>'.($num_format_type=='n_date' ? intval($serial) : $serial).'</v></c>');
		} elseif ($num_format_type=='n_numeric') {
			$number = self::numericValue($value);
			$file->write($number !== null ? '<c '.$cell_attrs.' t="n"><v>'.$number.'</v></c>' : self::inlineStringCell($cell_attrs, $value));//int,float,currency
		} elseif ($num_format_type=='n_string') {
			$file->write(self::inlineStringCell($cell_attrs, $value));
		} else { //n_auto: auto-detect unknown column types
			if (!is_string($value)) {
				$number = self::numericValue($value);//int or float
			} elseif (($value[0]!='0' || $value==='0') && ctype_digit($value)) {
				$number = isset($value[15]) ? null : $value;//16+ digit ids (credit cards, barcodes...) would lose precision as numbers
			} elseif (preg_match("/^\-?(0|[1-9][0-9]*)(\.[0-9]+)?$/", $value)) {
				$number = $value;
			} else {
				$number = null;
			}
			$file->write($number !== null ? '<c '.$cell_attrs.' t="n"><v>'.$number.'</v></c>' : self::inlineStringCell($cell_attrs, $value));
		}
	}

	/**
	 * @return string|null the value formatted for a <v> numeric element, or null when it is not a finite number
	 */
	protected static function numericValue($value)
	{
		if (is_int($value)) {
			return (string)$value;
		}
		if (is_float($value)) {
			return is_finite($value) ? (string)$value : null;
		}
		if (is_string($value) && is_numeric($value)) {
			$value = trim($value);
			return is_finite((float)$value) ? $value : null;
		}
		return null;
	}

	protected static function inlineStringCell($cell_attrs, $value)
	{
		if (is_float($value) && !is_finite($value)) {
			$value = is_nan($value) ? 'NAN' : ($value > 0 ? 'INF' : '-INF');//php 8.5 warns when NAN is cast to string
		}
		$value = (string)$value;
		if (isset($value[self::EXCEL_MAX_CELL_CHARS])) {//byte length check first, it is cheap
			$value = self::substring($value, 0, self::EXCEL_MAX_CELL_CHARS);//excel refuses to open cells longer than this
		}
		$first = $value[0];
		$last = substr($value, -1);
		$preserve = ($first==' ' || $first=="\t" || $first=="\n" || $first=="\r" || $last==' ' || $last=="\t" || $last=="\n" || $last=="\r") ? ' xml:space="preserve"' : '';
		return '<c '.$cell_attrs.' t="inlineStr"><is><t'.$preserve.'>'.self::xmlspecialchars($value).'</t></is></c>';
	}

	/**
	 * @return string|null 'FFRRGGBB' for '#rgb' / '#rrggbb' input, null for anything else
	 */
	protected static function normalizeColor($color)
	{
		if (!is_string($color) || !isset($color[0]) || $color[0]!='#') {
			return null;
		}
		$v = substr($color,1,6);
		$v = strlen($v)==3 ? $v[0].$v[0].$v[1].$v[1].$v[2].$v[2] : $v;// expand cf0 => ccff00
		if (strlen($v)!=6 || !ctype_xdigit($v)) {
			return null;//also prevents injecting markup into the styles.xml attribute
		}
		return "FF".strtoupper($v);
	}

	protected function styleFontIndexes()
	{
		static $border_allowed = array('left','right','top','bottom');
		static $border_style_allowed = array('thin','medium','thick','dashDot','dashDotDot','dashed','dotted','double','hair','mediumDashDot','mediumDashDotDot','mediumDashed','slantDashDot');
		static $horizontal_allowed = array('general','left','right','justify','center','centerContinuous','distributed','fill');
		static $vertical_allowed = array('bottom','center','distributed','top','justify');
		$default_font = array('size'=>'10','name'=>'Arial','family'=>'2');
		$fills = array('','');//2 placeholders for static xml later
		$fonts = array('','','','');//4 placeholders for static xml later
		$borders = array('');//1 placeholder for static xml later
		$style_indexes = array();
		foreach($this->cell_styles as $i=>$cell_style_string)
		{
			$semi_colon_pos = strpos($cell_style_string,";");
			$number_format_idx = substr($cell_style_string, 0, $semi_colon_pos);
			$style_json_string = substr($cell_style_string, $semi_colon_pos+1);
			$style = $style_json_string!=='' ? json_decode($style_json_string, true) : null;
			$style = is_array($style) ? $style : array();

			$style_indexes[$i] = array('num_fmt_idx'=>$number_format_idx);//initialize entry
			if (isset($style['border']) && is_string($style['border']))//border is a comma delimited str
			{
				$border_value = array();//reset per style, otherwise style/color leak from the previous style
				$requested_sides = array_map('trim', explode(",", $style['border']));
				$border_value['side'] = array_values(array_intersect($border_allowed, $requested_sides));
				if (isset($style['border-style']) && in_array($style['border-style'],$border_style_allowed,true))
				{
					$border_value['style'] = $style['border-style'];
				}
				$border_color = isset($style['border-color']) ? self::normalizeColor($style['border-color']) : null;
				if ($border_color!==null)
				{
					$border_value['color'] = $border_color;
				}
				$style_indexes[$i]['border_idx'] = self::add_to_list_get_index($borders, json_encode($border_value));
			}
			$fill = isset($style['fill']) ? self::normalizeColor($style['fill']) : null;
			if ($fill!==null)
			{
				$style_indexes[$i]['fill_idx'] = self::add_to_list_get_index($fills, $fill);
			}
			if (isset($style['halign']) && in_array($style['halign'],$horizontal_allowed,true))
			{
				$style_indexes[$i]['alignment'] = true;
				$style_indexes[$i]['halign'] = $style['halign'];
			}
			if (isset($style['valign']) && in_array($style['valign'],$vertical_allowed,true))
			{
				$style_indexes[$i]['alignment'] = true;
				$style_indexes[$i]['valign'] = $style['valign'];
			}
			if (isset($style['wrap_text']))
			{
				$style_indexes[$i]['alignment'] = true;
				$style_indexes[$i]['wrap_text'] = (bool)$style['wrap_text'];
			}

			$font = $default_font;
			if (isset($style['font-size']) && is_numeric($style['font-size']) && $style['font-size'] > 0)
			{
				$font['size'] = floatval($style['font-size']);//floatval to allow "10.5" etc
			}
			if (isset($style['font']) && is_string($style['font']))
			{
				if ($style['font']=='Comic Sans MS') { $font['family']=4; }
				if ($style['font']=='Times New Roman') { $font['family']=1; }
				if ($style['font']=='Courier New') { $font['family']=3; }
				$font['name'] = strval($style['font']);
			}
			if (isset($style['font-style']) && is_string($style['font-style']))
			{
				if (strpos($style['font-style'], 'bold')!==false) { $font['bold'] = true; }
				if (strpos($style['font-style'], 'italic')!==false) { $font['italic'] = true; }
				if (strpos($style['font-style'], 'strike')!==false) { $font['strike'] = true; }
				if (strpos($style['font-style'], 'underline')!==false) { $font['underline'] = true; }
			}
			$color = isset($style['color']) ? self::normalizeColor($style['color']) : null;
			if ($color!==null)
			{
				$font['color'] = $color;
			}
			if ($font!=$default_font)
			{
				$style_indexes[$i]['font_idx'] = self::add_to_list_get_index($fonts, json_encode($font) );
			}
		}
		return array('fills'=>$fills,'fonts'=>$fonts,'borders'=>$borders,'styles'=>$style_indexes );
	}

	protected function writeStylesXML()
	{
		$r = $this->styleFontIndexes();
		$fills = $r['fills'];
		$fonts = $r['fonts'];
		$borders = $r['borders'];
		$style_indexes = $r['styles'];

		$temporary_filename = $this->tempFilename();
		$file = new XLSXWriter_BuffererWriter($temporary_filename);
		$file->write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'."\n");
		$file->write('<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">');
		$file->write('<numFmts count="'.count($this->number_formats).'">');
		foreach($this->number_formats as $i=>$v) {
			$file->write('<numFmt numFmtId="'.(164+$i).'" formatCode="'.self::xmlspecialchars($v).'" />');
		}
		$file->write('</numFmts>');

		$file->write('<fonts count="'.(count($fonts)).'">');
		$file->write(		'<font><name val="Arial"/><charset val="1"/><family val="2"/><sz val="10"/></font>');
		$file->write(		'<font><name val="Arial"/><family val="0"/><sz val="10"/></font>');
		$file->write(		'<font><name val="Arial"/><family val="0"/><sz val="10"/></font>');
		$file->write(		'<font><name val="Arial"/><family val="0"/><sz val="10"/></font>');

		foreach($fonts as $font) {
			if (!empty($font)) { //fonts have 4 empty placeholders in array to offset the 4 static xml entries above
				$f = json_decode($font,true);
				$file->write('<font>');
				$file->write(	'<name val="'.self::xmlspecialchars($f['name']).'"/><charset val="1"/><family val="'.intval($f['family']).'"/>');
				$file->write(	'<sz val="'.floatval($f['size']).'"/>');//sz is a double, keep "10.5"
				if (!empty($f['color'])) { $file->write('<color rgb="'.strval($f['color']).'"/>'); }
				if (!empty($f['bold'])) { $file->write('<b val="true"/>'); }
				if (!empty($f['italic'])) { $file->write('<i val="true"/>'); }
				if (!empty($f['underline'])) { $file->write('<u val="single"/>'); }
				if (!empty($f['strike'])) { $file->write('<strike val="true"/>'); }
				$file->write('</font>');
			}
		}
		$file->write('</fonts>');

		$file->write('<fills count="'.(count($fills)).'">');
		$file->write(	'<fill><patternFill patternType="none"/></fill>');
		$file->write(	'<fill><patternFill patternType="gray125"/></fill>');
		foreach($fills as $fill) {
			if (!empty($fill)) { //fills have 2 empty placeholders in array to offset the 2 static xml entries above
				$file->write('<fill><patternFill patternType="solid"><fgColor rgb="'.strval($fill).'"/><bgColor indexed="64"/></patternFill></fill>');
			}
		}
		$file->write('</fills>');

		$file->write('<borders count="'.(count($borders)).'">');
		$file->write(    '<border diagonalDown="false" diagonalUp="false"><left/><right/><top/><bottom/><diagonal/></border>');
		foreach($borders as $border) {
			if (!empty($border)) { //borders have an empty placeholder in the array to offset the static xml entry above
				$pieces = json_decode($border,true);
				$border_style = !empty($pieces['style']) ? $pieces['style'] : 'hair';
				$border_color = !empty($pieces['color']) ? '<color rgb="'.strval($pieces['color']).'"/>' : '';
				$file->write('<border diagonalDown="false" diagonalUp="false">');
				foreach (array('left', 'right', 'top', 'bottom') as $side)
				{
					$show_side = in_array($side,$pieces['side'],true);
					$file->write($show_side ? "<$side style=\"$border_style\">$border_color</$side>" : "<$side/>");
				}
				$file->write(  '<diagonal/>');
				$file->write('</border>');
			}
		}
		$file->write('</borders>');

		$file->write('<cellStyleXfs count="20">');
		$file->write(		'<xf applyAlignment="true" applyBorder="true" applyFont="true" applyProtection="true" borderId="0" fillId="0" fontId="0" numFmtId="164">');
		$file->write(		'<alignment horizontal="general" indent="0" shrinkToFit="false" textRotation="0" vertical="bottom" wrapText="false"/>');
		$file->write(		'<protection hidden="false" locked="true"/>');
		$file->write(		'</xf>');
		$file->write(		'<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="0"/>');
		$file->write(		'<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="0"/>');
		$file->write(		'<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="2" numFmtId="0"/>');
		$file->write(		'<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="2" numFmtId="0"/>');
		for($x=0; $x<10; $x++) {
			$file->write(	'<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
		}
		$file->write(		'<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="43"/>');
		$file->write(		'<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="41"/>');
		$file->write(		'<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="44"/>');
		$file->write(		'<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="42"/>');
		$file->write(		'<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="9"/>');
		$file->write('</cellStyleXfs>');

		$file->write('<cellXfs count="'.(count($style_indexes)).'">');
		foreach($style_indexes as $v)
		{
			$applyAlignment = isset($v['alignment']) ? 'true' : 'false';
			$wrapText = !empty($v['wrap_text']) ? 'true' : 'false';
			$horizAlignment = isset($v['halign']) ? $v['halign'] : 'general';
			$vertAlignment = isset($v['valign']) ? $v['valign'] : 'bottom';
			$applyBorder = isset($v['border_idx']) ? 'true' : 'false';
			$applyFont = 'true';
			$borderIdx = isset($v['border_idx']) ? intval($v['border_idx']) : 0;
			$fillIdx = isset($v['fill_idx']) ? intval($v['fill_idx']) : 0;
			$fontIdx = isset($v['font_idx']) ? intval($v['font_idx']) : 0;
			$file->write('<xf applyAlignment="'.$applyAlignment.'" applyBorder="'.$applyBorder.'" applyFont="'.$applyFont.'" applyProtection="false" borderId="'.($borderIdx).'" fillId="'.($fillIdx).'" fontId="'.($fontIdx).'" numFmtId="'.(164+$v['num_fmt_idx']).'" xfId="0">');
			$file->write('	<alignment horizontal="'.$horizAlignment.'" vertical="'.$vertAlignment.'" textRotation="0" wrapText="'.$wrapText.'" indent="0" shrinkToFit="false"/>');
			$file->write('	<protection locked="true" hidden="false"/>');
			$file->write('</xf>');
		}
		$file->write('</cellXfs>');
		$file->write(	'<cellStyles count="6">');
		$file->write(		'<cellStyle builtinId="0" customBuiltin="false" name="Normal" xfId="0"/>');
		$file->write(		'<cellStyle builtinId="3" customBuiltin="false" name="Comma" xfId="15"/>');
		$file->write(		'<cellStyle builtinId="6" customBuiltin="false" name="Comma [0]" xfId="16"/>');
		$file->write(		'<cellStyle builtinId="4" customBuiltin="false" name="Currency" xfId="17"/>');
		$file->write(		'<cellStyle builtinId="7" customBuiltin="false" name="Currency [0]" xfId="18"/>');
		$file->write(		'<cellStyle builtinId="5" customBuiltin="false" name="Percent" xfId="19"/>');
		$file->write(	'</cellStyles>');
		$file->write('</styleSheet>');
		$file->close();
		return $temporary_filename;
	}

	protected function buildAppXML()
	{
		$app_xml="";
		$app_xml.='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'."\n";
		$app_xml.='<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">';
		$app_xml.='<TotalTime>0</TotalTime>';
		$app_xml.='<Company>'.self::xmlspecialchars($this->company).'</Company>';
		$app_xml.='</Properties>';
		return $app_xml;
	}

	protected function buildCoreXML()
	{
		$core_xml="";
		$core_xml.='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'."\n";
		$core_xml.='<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">';
		$core_xml.='<dcterms:created xsi:type="dcterms:W3CDTF">'.gmdate("Y-m-d\TH:i:s.00\Z").'</dcterms:created>';//UTC, the trailing Z says so; $date_time = '2014-10-25T15:54:37.00Z';
		$core_xml.='<dc:title>'.self::xmlspecialchars($this->title).'</dc:title>';
		$core_xml.='<dc:subject>'.self::xmlspecialchars($this->subject).'</dc:subject>';
		$core_xml.='<dc:creator>'.self::xmlspecialchars($this->author).'</dc:creator>';
		if (!empty($this->keywords)) {
			$core_xml.='<cp:keywords>'.self::xmlspecialchars(implode (", ", (array)$this->keywords)).'</cp:keywords>';
		}
		$core_xml.='<dc:description>'.self::xmlspecialchars($this->description).'</dc:description>';
		$core_xml.='<cp:revision>0</cp:revision>';
		$core_xml.='</cp:coreProperties>';
		return $core_xml;
	}

	protected function buildRelationshipsXML()
	{
		$rels_xml="";
		$rels_xml.='<?xml version="1.0" encoding="UTF-8"?>'."\n";
		$rels_xml.='<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
		$rels_xml.='<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>';
		$rels_xml.='<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>';
		$rels_xml.='<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>';
		$rels_xml.="\n";
		$rels_xml.='</Relationships>';
		return $rels_xml;
	}

	protected function buildWorkbookXML()
	{
		$workbook_xml="";
		$workbook_xml.='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'."\n";
		$workbook_xml.='<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';
		$workbook_xml.='<fileVersion appName="Calc"/><workbookPr backupFile="false" showObjects="all" date1904="false"/><workbookProtection/>';
		$workbook_xml.='<bookViews><workbookView activeTab="0" firstSheet="0" showHorizontalScroll="true" showSheetTabs="true" showVerticalScroll="true" tabRatio="212" windowHeight="8192" windowWidth="16384" xWindow="0" yWindow="0"/></bookViews>';
		$workbook_xml.='<sheets>';
		$defined_names = '';
		$i=0;
		foreach($this->sheets as $sheet) {
			$workbook_xml.='<sheet name="'.self::xmlspecialchars($sheet->safe_sheetname).'" sheetId="'.($i+1).'" state="visible" r:id="rId'.($i+2).'"/>';
			if ($sheet->auto_filter) {
				//localSheetId is the 0-based position of the sheet, apostrophes in a quoted sheet name are doubled
				$quoted_name = "'".str_replace("'", "''", $sheet->safe_sheetname)."'";
				$last_cell = self::xlsCell(max(0, $sheet->row_count - 1), max(0, count($sheet->columns) - 1), true);
				$defined_names.='<definedName name="_xlnm._FilterDatabase" localSheetId="'.$i.'" hidden="1">'.self::xmlspecialchars($quoted_name).'!$A$1:'.$last_cell.'</definedName>';
			}
			$i++;
		}
		$workbook_xml.='</sheets>';
		if ($defined_names!=='') {
			$workbook_xml.='<definedNames>'.$defined_names.'</definedNames>';
		}
		$workbook_xml.='<calcPr iterateCount="100" refMode="A1" iterate="false" iterateDelta="0.001" fullCalcOnLoad="1"/></workbook>';
		return $workbook_xml;
	}

	protected function buildWorkbookRelsXML()
	{
		$i=0;
		$wkbkrels_xml="";
		$wkbkrels_xml.='<?xml version="1.0" encoding="UTF-8"?>'."\n";
		$wkbkrels_xml.='<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
		$wkbkrels_xml.='<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>';
		foreach($this->sheets as $sheet) {
			$wkbkrels_xml.='<Relationship Id="rId'.($i+2).'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/'.($sheet->xmlname).'"/>';
			$i++;
		}
		$wkbkrels_xml.="\n";
		$wkbkrels_xml.='</Relationships>';
		return $wkbkrels_xml;
	}

	protected function buildContentTypesXML()
	{
		$content_types_xml="";
		$content_types_xml.='<?xml version="1.0" encoding="UTF-8"?>'."\n";
		$content_types_xml.='<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">';
		$content_types_xml.='<Override PartName="/_rels/.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>';
		$content_types_xml.='<Override PartName="/xl/_rels/workbook.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>';
		foreach($this->sheets as $sheet) {
			$content_types_xml.='<Override PartName="/xl/worksheets/'.($sheet->xmlname).'" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
		}
		$content_types_xml.='<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>';
		$content_types_xml.='<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>';
		$content_types_xml.='<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>';
		$content_types_xml.='<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>';
		$content_types_xml.="\n";
		$content_types_xml.='</Types>';
		return $content_types_xml;
	}

	//------------------------------------------------------------------
	/*
	 * @param $row_number int, zero based
	 * @param $column_number int, zero based
	 * @param $absolute bool
	 * @return Cell label/coordinates, ex: A1, C3, AA42 (or if $absolute==true: $A$1, $C$3, $AA$42)
	 * */
	public static function xlsCell($row_number, $column_number, $absolute=false)
	{
		static $column_letters = array();//called for every cell, so cache the column letters
		$column_number = (int)$column_number;
		if (!isset($column_letters[$column_number])) {
			for($r = "", $n = $column_number; $n >= 0; $n = intval($n / 26) - 1) {
				$r = chr($n%26 + 0x41) . $r;
			}
			$column_letters[$column_number] = $r;
		}
		$r = $column_letters[$column_number];
		$row = (int)$row_number + 1;
		if ($absolute) {
			return '$' . $r . '$' . $row;
		}
		return $r . $row;
	}
	//------------------------------------------------------------------
	public static function log($string)
	{
		error_log(date("Y-m-d H:i:s:").rtrim(is_array($string) ? json_encode($string) : $string)."\n");
	}
	//------------------------------------------------------------------
	public static function sanitize_filename($filename) //http://msdn.microsoft.com/en-us/library/aa365247%28VS.85%29.aspx
	{
		static $all_invalids = null;
		if ($all_invalids === null) {
			$all_invalids = array_merge(array_map('chr', range(0,31)), array('<', '>', '?', '"', ':', '|', '\\', '/', '*', '&'));
		}
		return str_replace($all_invalids, "", (string)$filename);
	}
	//------------------------------------------------------------------
	public static function sanitize_sheetname($sheetname)
	{
		static $badchars  = '\\/?*:[]';
		static $goodchars = '        ';
		$sheetname = strtr((string)$sheetname, $badchars, $goodchars);
		$sheetname = self::substring($sheetname, 0, self::EXCEL_MAX_SHEETNAME_CHARS);
		$sheetname = trim(trim(trim($sheetname),"'"));//trim before and after trimming single quotes
		if (strcasecmp($sheetname, 'History')==0) {
			$sheetname .= '_';//"History" is reserved by excel
		}
		return $sheetname!=='' ? $sheetname : 'Sheet'.((rand()%900)+100);
	}
	//------------------------------------------------------------------
	public static function xmlspecialchars($val)
	{
		//note, badchars does not include \t\n\r (\x09\x0a\x0d)
		static $badchars = "\x00\x01\x02\x03\x04\x05\x06\x07\x08\x0b\x0c\x0e\x0f\x10\x11\x12\x13\x14\x15\x16\x17\x18\x19\x1a\x1b\x1c\x1d\x1e\x1f\x7f";
		static $goodchars = "                              ";
		return strtr(htmlspecialchars((string)$val, ENT_QUOTES | ENT_XML1 | ENT_SUBSTITUTE, 'UTF-8'), $badchars, $goodchars);//strtr appears to be faster than str_replace
	}
	//------------------------------------------------------------------
	public static function array_first_key(array $arr)
	{
		reset($arr);
		$first_key = key($arr);
		return $first_key;
	}
	//------------------------------------------------------------------
	protected static function substring($string, $start, $length)
	{
		return function_exists('mb_substr') ? mb_substr($string, $start, $length, 'UTF-8') : substr($string, $start, $length);
	}
	//------------------------------------------------------------------
	protected static function lowercase($string)
	{
		return function_exists('mb_strtolower') ? mb_strtolower($string, 'UTF-8') : strtolower($string);
	}
	//------------------------------------------------------------------
	private static function determineNumberFormatType($num_format)
	{
		$num_format = preg_replace("/\[(Black|Blue|Cyan|Green|Magenta|Red|White|Yellow)\]/i", "", $num_format);
		if ($num_format=='GENERAL') return 'n_auto';
		if ($num_format=='@') return 'n_string';
		if ($num_format=='0') return 'n_numeric';
		if (preg_match('/[H]{1,2}:[M]{1,2}(?![^"]*+")/i', $num_format)) return 'n_datetime';
		if (preg_match('/[M]{1,2}:[S]{1,2}(?![^"]*+")/i', $num_format)) return 'n_datetime';
		if (preg_match('/[Y]{2,4}(?![^"]*+")/i', $num_format)) return 'n_date';
		if (preg_match('/[D]{1,2}(?![^"]*+")/i', $num_format)) return 'n_date';
		if (preg_match('/[M]{1,2}(?![^"]*+")/i', $num_format)) return 'n_date';
		if (preg_match('/\$(?![^"]*+")/', $num_format)) return 'n_numeric';//literal $, an unescaped $ (end anchor) matched every format
		if (preg_match('/%(?![^"]*+")/', $num_format)) return 'n_numeric';
		if (preg_match('/0(?![^"]*+")/', $num_format)) return 'n_numeric';
		return 'n_auto';
	}
	//------------------------------------------------------------------
	private static function numberFormatStandardized($num_format)
	{
		$num_format = (string)$num_format;
		if ($num_format=='money') { $num_format='dollar'; }
		if ($num_format=='number') { $num_format='integer'; }

		if      ($num_format=='string')   $num_format='@';
		else if ($num_format=='integer')  $num_format='0';
		else if ($num_format=='date')     $num_format='YYYY-MM-DD';
		else if ($num_format=='datetime') $num_format='YYYY-MM-DD HH:MM:SS';
		else if ($num_format=='time')     $num_format='HH:MM:SS';
		else if ($num_format=='price')    $num_format='#,##0.00';
		else if ($num_format=='dollar')   $num_format='[$$-1009]#,##0.00;[RED]-[$$-1009]#,##0.00';
		else if ($num_format=='euro')     $num_format='#,##0.00 [$€-407];[RED]-#,##0.00 [$€-407]';
		$ignore_until='';
		$escaped = '';
		for($i=0,$ix=strlen($num_format); $i<$ix; $i++)
		{
			$c = $num_format[$i];
			if ($ignore_until=='' && $c=='[')
				$ignore_until=']';
			else if ($ignore_until=='' && $c=='"')
				$ignore_until='"';
			else if ($ignore_until==$c)
				$ignore_until='';
			if ($ignore_until=='' && ($c==' ' || $c=='-'  || $c=='('  || $c==')') && ($i==0 || $num_format[$i-1]!='_'))
				$escaped.= "\\".$c;
			else
				$escaped.= $c;
		}
		return $escaped;
	}
	//------------------------------------------------------------------
	public static function add_to_list_get_index(&$haystack, $needle)
	{
		$existing_idx = array_search($needle, $haystack, true);
		if ($existing_idx===false)
		{
			$existing_idx = count($haystack);
			$haystack[] = $needle;
		}
		return $existing_idx;
	}
	//------------------------------------------------------------------
	public static function convert_date_time($date_input) //thanks to Excel::Writer::XLSX::Worksheet.pm (perl)
	{
		$days    = 0;    # Number of days since epoch
		$seconds = 0;    # Time expressed as fraction of 24h hours in seconds
		$year=$month=$day=0;
		$hour=$min  =$sec=0;
		$has_date = false;

		$date_time = (string)$date_input;
		if (preg_match("/(\d{4})\-(\d{2})\-(\d{2})/", $date_time, $matches))
		{
			list(,$year,$month,$day) = $matches;
			$year = (int)$year; $month = (int)$month; $day = (int)$day;
			$has_date = true;
		}
		if (preg_match("/(\d+):(\d{2}):(\d{2})/", $date_time, $matches))
		{
			list(,$hour,$min,$sec) = $matches;
			$seconds = ( $hour * 60 * 60 + $min * 60 + $sec ) / ( 24 * 60 * 60 );
		}

		//time only (ie: "23:59:59"), the serial is just the fraction of a day
		if (!$has_date) return $seconds;

		//using 1900 as epoch, not 1904, ignoring 1904 special case

		# Special cases for Excel.
		if ($year==1899 && $month==12 && $day==31)  return $seconds      ;    # Excel 1900 epoch
		if ($year==1900 && $month==1  && $day==0 )  return $seconds      ;    # Excel 1900 epoch
		if ($year==1900 && $month==2  && $day==29)  return 60 + $seconds ;    # Excel false leapday

		# We calculate the date by calculating the number of days since the epoch
		# and adjust for the number of leap days. We calculate the number of leap
		# days by normalising the year in relation to the epoch. Thus the year 2000
		# becomes 100 for 4 and 100 year leapdays and 400 for 400 year leapdays.
		$epoch  = 1900;
		$offset = 0;
		$norm   = 300;
		$range  = $year - $epoch;

		# Set month days and check for leap year.
		$leap = (($year % 400 == 0) || (($year % 4 == 0) && ($year % 100)) ) ? 1 : 0;
		$mdays = array( 31, ($leap ? 29 : 28), 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 );

		# Some boundary checks
		if($year < $epoch || $year > 9999) return 0;
		if($month < 1     || $month > 12)  return 0;
		if($day < 1       || $day > $mdays[ $month - 1 ]) return 0;

		# Accumulate the number of days since the epoch.
		$days = $day;    # Add days for current month
		$days += array_sum( array_slice($mdays, 0, $month-1 ) );    # Add days for past months
		$days += $range * 365;                      # Add days for past years
		$days += intval( ( $range ) / 4 );             # Add leapdays
		$days -= intval( ( $range + $offset ) / 100 ); # Subtract 100 year leapdays
		$days += intval( ( $range + $offset + $norm ) / 400 );  # Add 400 year leapdays
		$days -= $leap;                                      # Already counted above

		# Adjust for Excel erroneously treating 1900 as a leap year.
		if ($days > 59) { $days++;}

		return $days + $seconds;
	}
	//------------------------------------------------------------------
}

class XLSXWriter_BuffererWriter
{
	const BUFFER_SIZE = 65536;//fewer, larger fwrite() calls

	protected $fd=null;
	protected $buffer='';
	protected $check_utf8=false;

	public function __construct($filename, $fd_fopen_flags='w', $check_utf8=false)
	{
		$this->check_utf8 = $check_utf8;
		$this->fd = fopen($filename, $fd_fopen_flags);
		if ($this->fd===false) {
			XLSXWriter::log("Unable to open $filename for writing.");
		}
	}

	public function write($string)
	{
		$this->buffer.=$string;
		if (isset($this->buffer[self::BUFFER_SIZE-1])) {
			$this->purge();
		}
	}

	protected function purge()
	{
		if ($this->fd) {
			if ($this->check_utf8 && !self::isValidUTF8($this->buffer)) {
				XLSXWriter::log("Error, invalid UTF8 encoding detected.");
				$this->check_utf8 = false;
			}
			fwrite($this->fd, $this->buffer);
		}
		$this->buffer='';//discard even without a handle, so memory can not grow without bound
	}

	public function close()
	{
		$this->purge();
		if ($this->fd) {
			fclose($this->fd);
			$this->fd=null;
		}
	}

	public function __destruct()
	{
		$this->close();
	}

	public function ftell()
	{
		if ($this->fd) {
			$this->purge();
			return ftell($this->fd);
		}
		return -1;
	}

	public function fseek($pos)
	{
		if ($this->fd) {
			$this->purge();
			return fseek($this->fd, $pos);
		}
		return -1;
	}

	protected static function isValidUTF8($string)
	{
		if (function_exists('mb_check_encoding'))
		{
			return mb_check_encoding($string, 'UTF-8') ? true : false;
		}
		return preg_match("//u", $string) ? true : false;
	}
}



// vim: set filetype=php expandtab tabstop=4 shiftwidth=4 autoindent smartindent:
