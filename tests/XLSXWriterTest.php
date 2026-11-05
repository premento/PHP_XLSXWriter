<?php

use PHPUnit\Framework\TestCase;

/**
 * Exposes protected internals for testing.
 */
class TestableXLSXWriter extends XLSXWriter
{
    public function writeCellToString($value, $num_format_type, $cell_style_idx = 0)
    {
        $filename = tempnam(sys_get_temp_dir(), 'xlsx_cell_');
        $file_writer = new XLSXWriter_BuffererWriter($filename);
        $this->writeCell($file_writer, 0, 0, $value, $num_format_type, $cell_style_idx);
        $file_writer->close();
        $xml = file_get_contents($filename);
        @unlink($filename);
        return $xml;
    }

    public function cellStyleCount()
    {
        return count($this->cell_styles);
    }

    /** @var string[] messages passed to log() */
    public static $logged = array();

    public static function log($string)
    {
        self::$logged[] = is_array($string) ? json_encode($string) : (string)$string;
    }
}

class XLSXWriterTest extends TestCase
{
    /** @var string[] */
    private $files = array();
    protected function setUp(): void
    {
        TestableXLSXWriter::$logged = array();
    }

    protected function tearDown(): void
    {
        foreach ($this->files as $file) {
            @unlink($file);
        }
    }

    private function tempXlsx()
    {
        $filename = tempnam(sys_get_temp_dir(), 'xlsx_test_') . '.xlsx';
        $this->files[] = $filename;
        return $filename;
    }

    /**
     * Writes the workbook and returns every part of the zip, keyed by name.
     * @return array<string,string>
     */
    private function parts(XLSXWriter $writer)
    {
        $filename = $this->tempXlsx();
        $this->assertTrue($writer->writeToFile($filename));
        $zip = new ZipArchive();
        $this->assertTrue($zip->open($filename));
        $parts = array();
        for ($i = 0; $i < $zip->numFiles; $i++) {
            $name = $zip->getNameIndex($i);
            if (substr($name, -1) !== '/') {
                $parts[$name] = $zip->getFromIndex($i);
            }
        }
        $zip->close();
        foreach ($parts as $name => $xml) {
            if (substr($name, -4) === '.xml' || substr($name, -5) === '.rels') {
                $this->assertWellFormedXml($xml, $name);
            }
        }
        return $parts;
    }

    private function assertWellFormedXml($xml, $name)
    {
        $previous = libxml_use_internal_errors(true);
        $doc = simplexml_load_string($xml);
        libxml_clear_errors();
        libxml_use_internal_errors($previous);
        $this->assertNotFalse($doc, "$name is not well formed XML");
    }

    private function sheetCells($sheet_xml)
    {
        $output = array();
        $xml = new SimpleXMLElement($sheet_xml);
        foreach ($xml->sheetData->row as $row) {
            $cells = array();
            foreach ($row->c as $c) {
                $cells[] = isset($c->is) ? (string)$c->is->t : (string)$c->v;
            }
            $output[] = $cells;
        }
        return $output;
    }

    private function logContents()
    {
        return implode("\n", TestableXLSXWriter::$logged);
    }

    // ---------------------------------------------------------------- cells

    public function testStringCellIsInlineAndKeepsLeadingZero()
    {
        $w = new TestableXLSXWriter();
        $this->assertSame('<c r="A1" s="0" t="inlineStr"><is><t>0123</t></is></c>', $w->writeCellToString('0123', 'n_string'));
        $this->assertSame('<c r="A1" s="0" t="inlineStr"><is><t>0123</t></is></c>', $w->writeCellToString('0123', 'n_auto'));
    }

    public function testAutoDetectNumbers()
    {
        $w = new TestableXLSXWriter();
        $this->assertSame('<c r="A1" s="0" t="n"><v>123</v></c>', $w->writeCellToString('123', 'n_auto'));
        $this->assertSame('<c r="A1" s="0" t="n"><v>0</v></c>', $w->writeCellToString('0', 'n_auto'));
        $this->assertSame('<c r="A1" s="0" t="n"><v>-50.5</v></c>', $w->writeCellToString('-50.5', 'n_auto'));
        $this->assertSame('<c r="A1" s="0" t="n"><v>42</v></c>', $w->writeCellToString(42, 'n_auto'));
        $this->assertSame('<c r="A1" s="0" t="n"><v>1.5</v></c>', $w->writeCellToString(1.5, 'n_auto'));
    }

    public function testLongDigitStringsStayText()
    {
        $w = new TestableXLSXWriter();
        $this->assertSame('<c r="A1" s="0" t="inlineStr"><is><t>4111111111111111</t></is></c>', $w->writeCellToString('4111111111111111', 'n_auto'));
        $this->assertSame('<c r="A1" s="0" t="n"><v>411111111111111</v></c>', $w->writeCellToString('411111111111111', 'n_auto'));
    }

    public function testFormulaHasNoSharedStringType()
    {
        $w = new TestableXLSXWriter();
        $this->assertSame('<c r="A1" s="0"><f>SUM(B1:B2)</f></c>', $w->writeCellToString('=SUM(B1:B2)', 'n_auto'));
        $this->assertSame('<c r="A1" s="0"><f>&quot;a&quot;&amp;&quot;b&quot;</f></c>', $w->writeCellToString('="a"&"b"', 'n_string'));
        $this->assertSame('<c r="A1" s="0"><f>=A2</f></c>', $w->writeCellToString('==A2', 'n_auto'), 'only the first = is the formula marker');
    }

    public function testEmptyAndNonScalarCells()
    {
        $w = new TestableXLSXWriter();
        $this->assertSame('<c r="A1" s="0"/>', $w->writeCellToString('', 'n_auto'));
        $this->assertSame('<c r="A1" s="0"/>', $w->writeCellToString(null, 'n_numeric'));
        $this->assertSame('<c r="A1" s="0"/>', $w->writeCellToString(array(1), 'n_string'));
    }

    public function testBooleanCells()
    {
        $w = new TestableXLSXWriter();
        $this->assertSame('<c r="A1" s="0" t="b"><v>1</v></c>', $w->writeCellToString(true, 'n_auto'));
        $this->assertSame('<c r="A1" s="0" t="b"><v>0</v></c>', $w->writeCellToString(false, 'n_numeric'));
    }

    public function testNumericColumnNeverWritesInvalidNumbers()
    {
        $w = new TestableXLSXWriter();
        $this->assertSame('<c r="A1" s="0" t="n"><v>12.5</v></c>', $w->writeCellToString(' 12.5', 'n_numeric'));
        $this->assertSame('<c r="A1" s="0" t="inlineStr"><is><t>n/a</t></is></c>', $w->writeCellToString('n/a', 'n_numeric'));
        $this->assertSame('<c r="A1" s="0" t="inlineStr"><is><t>INF</t></is></c>', $w->writeCellToString(INF, 'n_numeric'));
        $this->assertSame('<c r="A1" s="0" t="inlineStr"><is><t>NAN</t></is></c>', $w->writeCellToString(NAN, 'n_auto'));
    }

    public function testDateCells()
    {
        $w = new TestableXLSXWriter();
        $this->assertSame('<c r="A1" s="0" t="n"><v>43107</v></c>', $w->writeCellToString('2018-01-07', 'n_date'));
        $this->assertSame('<c r="A1" s="0" t="n"><v>43107</v></c>', $w->writeCellToString('2018-01-07 12:00:00', 'n_date'));
        $this->assertSame('<c r="A1" s="0" t="n"><v>43107.5</v></c>', $w->writeCellToString('2018-01-07 12:00:00', 'n_datetime'));
        $this->assertSame('<c r="A1" s="0" t="n"><v>43107.5</v></c>', $w->writeCellToString(new DateTime('2018-01-07 12:00:00'), 'n_datetime'));
        $this->assertSame('<c r="A1" s="0" t="n"><v>43107</v></c>', $w->writeCellToString(43107, 'n_date'), 'numbers are taken as excel serials');
        $this->assertSame('<c r="A1" s="0" t="inlineStr"><is><t>unknown</t></is></c>', $w->writeCellToString('unknown', 'n_date'));
    }

    public function testXmlEscapingAndControlCharacters()
    {
        $w = new TestableXLSXWriter();
        $this->assertSame('<c r="A1" s="0" t="inlineStr"><is><t>&lt;a href=&quot;x&quot;&gt;&amp;&apos;</t></is></c>', $w->writeCellToString('<a href="x">&\'', 'n_string'));
        $this->assertSame('<c r="A1" s="0" t="inlineStr"><is><t>a b</t></is></c>', $w->writeCellToString("a\x07b", 'n_string'));
        $this->assertStringContainsString("\xEF\xBF\xBD", $w->writeCellToString("bad \xC3\x28 utf8", 'n_string'), 'invalid UTF-8 is substituted');
    }

    public function testWhitespaceIsPreserved()
    {
        $w = new TestableXLSXWriter();
        $this->assertSame('<c r="A1" s="0" t="inlineStr"><is><t xml:space="preserve">  padded </t></is></c>', $w->writeCellToString('  padded ', 'n_string'));
        $this->assertSame('<c r="A1" s="0" t="inlineStr"><is><t>no padding</t></is></c>', $w->writeCellToString('no padding', 'n_string'));
    }

    public function testOverlongStringIsTruncatedToExcelLimit()
    {
        $w = new TestableXLSXWriter();
        $xml = $w->writeCellToString(str_repeat('é', 40000), 'n_string');
        $text = (string)simplexml_load_string($xml)->is->t;
        $this->assertSame(XLSXWriter::EXCEL_MAX_CELL_CHARS, mb_strlen($text, 'UTF-8'));
    }

    // ---------------------------------------------------------------- helpers

    public function testXlsCell()
    {
        $this->assertSame('A1', XLSXWriter::xlsCell(0, 0));
        $this->assertSame('Z1', XLSXWriter::xlsCell(0, 25));
        $this->assertSame('AA2', XLSXWriter::xlsCell(1, 26));
        $this->assertSame('AZ1', XLSXWriter::xlsCell(0, 51));
        $this->assertSame('XFD1048576', XLSXWriter::xlsCell(1048575, 16383));
        $this->assertSame('$C$3', XLSXWriter::xlsCell(2, 2, true));
        $this->assertSame('B1', XLSXWriter::xlsCell(0, 1.5), 'fractional input must not raise a deprecation');
    }

    public function testConvertDateTime()
    {
        $this->assertSame(1, XLSXWriter::convert_date_time('1900-01-01'));
        $this->assertSame(60, XLSXWriter::convert_date_time('1900-02-29'));
        $this->assertSame(61, XLSXWriter::convert_date_time('1900-03-01'));
        $this->assertSame(36525, XLSXWriter::convert_date_time('1999-12-31'));
        $this->assertSame(43107, XLSXWriter::convert_date_time('2018-01-07'));
        $this->assertEqualsWithDelta(0.5, XLSXWriter::convert_date_time('12:00:00'), 1e-9, 'time only');
        $this->assertEqualsWithDelta(86399 / 86400, XLSXWriter::convert_date_time('23:59:59'), 1e-9);
        $this->assertSame(0, XLSXWriter::convert_date_time('0000-00-00'));
        $this->assertSame(0, XLSXWriter::convert_date_time('2018-02-30'));
        $this->assertSame(0, XLSXWriter::convert_date_time('1800-01-01'));
    }

    public function testSanitizeSheetname()
    {
        $this->assertSame('a b c', XLSXWriter::sanitize_sheetname('a/b?c'));
        $this->assertSame(31, mb_strlen(XLSXWriter::sanitize_sheetname(str_repeat('ж', 40)), 'UTF-8'));
        $this->assertSame('quoted', XLSXWriter::sanitize_sheetname("'quoted'"));
        $this->assertSame('History_', XLSXWriter::sanitize_sheetname('History'));
        $this->assertMatchesRegularExpression('/^Sheet\d{3}$/', XLSXWriter::sanitize_sheetname('[]'));
    }

    public function testSanitizeFilename()
    {
        $this->assertSame('reportQ1.xlsx', XLSXWriter::sanitize_filename("report<Q1>\n.xlsx"));
    }

    // ---------------------------------------------------------------- workbook

    public function testWriteToFileRoundTrip()
    {
        $header = array('c1' => 'string', 'c2' => 'string', 'c3' => 'string', 'c4' => 'string');
        $sheet = array(
            array('55', '66', '77', '88'),
            array('10', '11', '12', '13'),
        );
        $w = new XLSXWriter();
        $w->writeSheet($sheet, 'mysheet', $header);
        $parts = $this->parts($w);

        foreach (array('[Content_Types].xml', '_rels/.rels', 'docProps/app.xml', 'docProps/core.xml',
                     'xl/workbook.xml', 'xl/styles.xml', 'xl/_rels/workbook.xml.rels', 'xl/worksheets/sheet1.xml') as $name) {
            $this->assertArrayHasKey($name, $parts);
        }
        $cells = $this->sheetCells($parts['xl/worksheets/sheet1.xml']);
        $this->assertSame(array_keys($header), array_shift($cells));
        $this->assertSame($sheet, $cells);
        $this->assertStringContainsString('<dimension ref="A1:D3"/>', $parts['xl/worksheets/sheet1.xml']);
    }

    public function testDocumentProperties()
    {
        $w = new XLSXWriter();
        $w->setTitle('T & T');
        $w->setAuthor('Author');
        $w->setCompany('ACME');
        $w->setKeywords(array('a', 'b'));
        $w->writeSheetRow('Sheet1', array(1));
        $parts = $this->parts($w);
        $this->assertStringContainsString('<dc:title>T &amp; T</dc:title>', $parts['docProps/core.xml']);
        $this->assertStringContainsString('<cp:keywords>a, b</cp:keywords>', $parts['docProps/core.xml']);
        $this->assertStringContainsString('<Company>ACME</Company>', $parts['docProps/app.xml']);
        $this->assertMatchesRegularExpression('/<dcterms:created xsi:type="dcterms:W3CDTF">' . gmdate('Y-m-d') . 'T\d\d:\d\d:\d\d\.00Z</', $parts['docProps/core.xml']);
    }

    public function testWriteToStringAndStdOut()
    {
        $w = new XLSXWriter();
        $w->writeSheetRow('Sheet1', array('a'));
        $this->assertSame('PK', substr($w->writeToString(), 0, 2));

        ob_start();
        $this->assertTrue($w->writeToStdOut());
        $this->assertSame('PK', substr(ob_get_clean(), 0, 2));
    }

    public function testWriteToFileFailures()
    {
        $w = new TestableXLSXWriter();
        $this->assertFalse($w->writeToFile($this->tempXlsx()), 'no sheets');
        $this->assertStringContainsString('no worksheets defined', $this->logContents());

        $w->writeSheetRow('Sheet1', array(1));
        $this->assertFalse($w->writeToFile(sys_get_temp_dir() . '/does-not-exist-' . uniqid() . '/out.xlsx'));
        $this->assertStringContainsString('unable to', $this->logContents());
    }

    public function testSheetNamedZeroAndNumericLookalikes()
    {
        $w = new XLSXWriter();
        $w->writeSheetRow('0', array('zero'));
        $w->writeSheetRow('1', array('one'));
        $w->writeSheetRow('01', array('zero-one'));
        $this->assertSame(1, $w->countSheetRows('0'));
        $this->assertSame(1, $w->countSheetRows('01'));
        $parts = $this->parts($w);
        $this->assertStringContainsString('<sheet name="0" sheetId="1"', $parts['xl/workbook.xml']);
        $this->assertStringContainsString('<sheet name="01" sheetId="3"', $parts['xl/workbook.xml']);
        $this->assertSame(array(array('zero-one')), $this->sheetCells($parts['xl/worksheets/sheet3.xml']));
    }

    public function testCountSheetRowsWithoutSheets()
    {
        $w = new XLSXWriter();
        $this->assertSame(0, $w->countSheetRows());
        $this->assertSame(0, $w->countSheetRows('missing'));
    }

    public function testDuplicateSanitizedSheetNamesAreMadeUnique()
    {
        $w = new XLSXWriter();
        $w->writeSheetRow('a/b', array(1));
        $w->writeSheetRow('a?b', array(2));
        $w->writeSheetRow('A B', array(3));
        $w->writeSheetRow(str_repeat('x', 40), array(4));
        $w->writeSheetRow(str_repeat('x', 35), array(5));
        $workbook = $this->parts($w)['xl/workbook.xml'];
        preg_match_all('/<sheet name="([^"]*)"/', $workbook, $m);
        $this->assertSame(array('a b', 'a b (2)', 'A B (3)', str_repeat('x', 31), str_repeat('x', 27) . ' (2)'), $m[1]);
    }

    public function testRowWiderThanHeaderOnlyAddsMissingColumns()
    {
        $w = new XLSXWriter();
        $w->writeSheetHeader('S', array('a' => 'integer', 'b' => 'integer'), array('auto_filter' => true));
        $w->writeSheetRow('S', array(1, 2, 3));
        $parts = $this->parts($w);
        $this->assertStringContainsString('<dimension ref="A1:C2"/>', $parts['xl/worksheets/sheet1.xml']);
        $this->assertStringContainsString('<autoFilter ref="A1:C2"/>', $parts['xl/worksheets/sheet1.xml']);
    }

    public function testEmptySheetHasValidDimension()
    {
        $w = new XLSXWriter();
        $w->writeSheetHeader('Empty', array('x' => 'string'), array('suppress_row' => true));
        $parts = $this->parts($w);
        $this->assertStringContainsString('<dimension ref="A1:A1"/>', $parts['xl/worksheets/sheet1.xml']);
    }

    public function testAutoFilterDefinedNameUsesSheetPosition()
    {
        $w = new XLSXWriter();
        $w->writeSheetRow('Plain', array(1));
        $w->writeSheetHeader("Bob's", array('x' => 'string', 'y' => 'string'), array('auto_filter' => true));
        $w->writeSheetRow("Bob's", array('a', 'b'));
        $workbook = $this->parts($w)['xl/workbook.xml'];
        $this->assertStringContainsString('localSheetId="1"', $workbook);
        $this->assertStringContainsString('&apos;Bob&apos;&apos;s&apos;!$A$1:$B$2', $workbook);
    }

    public function testNoEmptyDefinedNames()
    {
        $w = new XLSXWriter();
        $w->writeSheetRow('Sheet1', array(1));
        $this->assertStringNotContainsString('definedNames', $this->parts($w)['xl/workbook.xml']);
    }

    public function testMarkMergedCellOnNewSheet()
    {
        $w = new XLSXWriter();
        $w->markMergedCell('New', 0, 0, 0, 3);
        $w->writeSheetRow('New', array('merged'));
        $this->assertStringContainsString('<mergeCell ref="A1:D1"/>', $this->parts($w)['xl/worksheets/sheet1.xml']);
    }

    public function testRowsAfterFinalizeAreRejectedLoudly()
    {
        $w = new TestableXLSXWriter();
        $w->writeSheet(array(array(1)), 'Done');
        $w->writeSheetRow('Done', array(2));
        $this->assertSame(1, $w->countSheetRows('Done'));
        $this->assertStringContainsString("sheet 'Done' has already been written", $this->logContents());
    }

    public function testFreezePaneSelections()
    {
        $w = new XLSXWriter();
        $w->writeSheetHeader('S', array('a' => 'string', 'b' => 'string'), array('freeze_rows' => 1, 'freeze_columns' => 1));
        $xml = $this->parts($w)['xl/worksheets/sheet1.xml'];
        $this->assertStringContainsString('<pane ySplit="1" xSplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/>', $xml);
        $this->assertStringContainsString('<selection activeCell="B1" activeCellId="0" pane="topRight" sqref="B1"/>', $xml);
        $this->assertStringContainsString('<selection activeCell="A2" activeCellId="0" pane="bottomLeft" sqref="A2"/>', $xml);
    }

    public function testColumnWidthsAndRowOptions()
    {
        $w = new XLSXWriter();
        $w->writeSheetHeader('S', array('a' => 'string', 'b' => 'string'), array('widths' => array(10, 20.5)));
        $w->writeSheetRow('S', array('x', 'y'), array('height' => 30, 'hidden' => true));
        $xml = $this->parts($w)['xl/worksheets/sheet1.xml'];
        $this->assertStringContainsString('max="2" min="2" style="0" customWidth="true" width="20.5"', $xml);
        $this->assertStringContainsString('customHeight="true" hidden="true" ht="30"', $xml);
    }

    public function testRightToLeft()
    {
        $w = new XLSXWriter();
        $w->setRightToLeft(true);
        $w->writeSheetRow('S', array('x'));
        $this->assertStringContainsString('rightToLeft="true"', $this->parts($w)['xl/worksheets/sheet1.xml']);
    }

    // ---------------------------------------------------------------- styles

    public function testRowOptionsDoNotCreateExtraStyles()
    {
        $w = new TestableXLSXWriter();
        $before = $w->cellStyleCount();
        for ($i = 0; $i < 20; $i++) {
            $w->writeSheetRow('S', array($i), array('height' => 10 + $i));
        }
        $this->assertSame($before, $w->cellStyleCount(), 'height/hidden must not register cell styles');
    }

    public function testShortStyleListDoesNotWarn()
    {
        $w = new XLSXWriter();
        $w->writeSheetRow('S', array(1, 2, 3), array(array('fill' => '#f00')));
        $styles = $this->parts($w)['xl/styles.xml'];
        $this->assertStringContainsString('<fgColor rgb="FFFF0000"/>', $styles);
    }

    public function testBorderStyleDoesNotLeakBetweenStyles()
    {
        $w = new XLSXWriter();
        $w->writeSheetRow('S', array('x'), array('border' => 'left', 'border-style' => 'thick', 'border-color' => '#f00'));
        $w->writeSheetRow('S', array('y'), array('border' => 'left'));
        $styles = $this->parts($w)['xl/styles.xml'];
        $this->assertStringContainsString('<left style="thick"><color rgb="FFFF0000"/></left>', $styles);
        $this->assertStringContainsString('<left style="hair"></left>', $styles);
    }

    public function testInvalidColorsAreIgnoredSafely()
    {
        $w = new XLSXWriter();
        $w->writeSheetRow('S', array(1), array('fill' => '', 'color' => '#', 'border' => 'top', 'border-color' => '#zzz'));
        $w->writeSheetRow('S', array(2), array('fill' => '#"/><x a="'));
        $styles = $this->parts($w)['xl/styles.xml'];
        $this->assertStringNotContainsString('<x a=', $styles);
        $this->assertStringContainsString('<fills count="2">', $styles);
    }

    public function testFontStyles()
    {
        $w = new XLSXWriter();
        $w->writeSheetRow('S', array('x'), array('font' => 'Courier New', 'font-size' => 10.5, 'font-style' => 'bold,italic,underline,strikethrough', 'color' => '#00f'));
        $styles = $this->parts($w)['xl/styles.xml'];
        $this->assertStringContainsString('<name val="Courier New"/><charset val="1"/><family val="3"/><sz val="10.5"/><color rgb="FF0000FF"/><b val="true"/><i val="true"/><u val="single"/><strike val="true"/>', $styles);
    }

    public function testAlignmentAndWrap()
    {
        $w = new XLSXWriter();
        $w->writeSheetRow('S', array('x'), array('halign' => 'center', 'valign' => 'top', 'wrap_text' => true));
        $styles = $this->parts($w)['xl/styles.xml'];
        $this->assertStringContainsString('<alignment horizontal="center" vertical="top" textRotation="0" wrapText="true"', $styles);
    }

    public function testHeaderStyle()
    {
        $w = new XLSXWriter();
        $w->writeSheetHeader('S', array('a' => 'integer'), array('font-style' => 'bold', 'widths' => array(10)));
        $styles = $this->parts($w)['xl/styles.xml'];
        $this->assertStringContainsString('<b val="true"/>', $styles);
        $this->assertStringNotContainsString('widths', $styles);
    }

    public function testNumberFormats()
    {
        $w = new XLSXWriter();
        $w->writeSheetHeader('S', array('d' => 'date', 'm' => 'dollar', 't' => 'time', 'p' => '0.00%'));
        $w->writeSheetRow('S', array('2018-01-07', '12.5', '06:00:00', '0.25'));
        $parts = $this->parts($w);
        $styles = $parts['xl/styles.xml'];
        $this->assertStringContainsString('formatCode="YYYY\-MM\-DD"', $styles);
        $this->assertStringContainsString('formatCode="HH:MM:SS"', $styles);
        $this->assertSame(array(array('d', 'm', 't', 'p'), array('43107', '12.5', '0.25', '0.25')), $this->sheetCells($parts['xl/worksheets/sheet1.xml']));
    }

    public function testDeprecatedBooleanHeaderOptionStillWorks()
    {
        $w = new TestableXLSXWriter();
        $w->writeSheetHeader('S', array('a' => 'string'), true);
        $this->assertSame(0, $w->countSheetRows('S'));
        $this->assertStringContainsString('passing $suppress_row=false|true', $this->logContents());
    }
}
