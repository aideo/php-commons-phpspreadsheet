<?php

namespace Ideo\Utilities\PhpOffice;

use Ideo\TestCase;
use PhpOffice\PhpSpreadsheet\Exception as PhpSpreadsheetException;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Reader\Exception as PhpSpreadsheetReaderException;

class PhpSpreadsheetUtilsTest extends TestCase
{

    /**
     * @throws PhpSpreadsheetException
     * @throws PhpSpreadsheetReaderException
     */
    public function testGetCellValue()
    {
        $wb = IOFactory::load(__DIR__ . '/test.xlsx');
        $sh = $wb->getSheet(0);

        $this->assertEquals(PhpSpreadsheetUtils::getCellValue($sh, 2, 2), "A");
        $this->assertEquals(PhpSpreadsheetUtils::getCellValue($sh, 2, 3), "B");
        $this->assertEquals(PhpSpreadsheetUtils::getCellValue($sh, 2, 4), "C");
        $this->assertEquals(PhpSpreadsheetUtils::getCellValue($sh, 4, 2), "Hello !! World !!");
        $this->assertEquals(PhpSpreadsheetUtils::getCellValue($sh, 4, 4), "1/1/2017 0:01");
        $this->assertEquals(PhpSpreadsheetUtils::getCellValue($sh, 4, 5), 10000);

        $this->assertEquals(PhpSpreadsheetUtils::getCellValue($sh, 4, 3), date('n/j/Y G:i'));
    }

    /**
     * @throws PhpSpreadsheetException
     * @throws PhpSpreadsheetReaderException
     */
    public function testGetCellValueByName()
    {
        $wb = IOFactory::load(__DIR__ . '/test.xlsx');
        $sh = $wb->getSheet(0);

        $this->assertEquals(PhpSpreadsheetUtils::getCellValueByName($sh, 'Named_Cell'), "Named Cell");
    }

    /**
     * @throws PhpSpreadsheetException
     * @throws PhpSpreadsheetReaderException
     */
    public function testGetHighestRowAndColumnIndex()
    {
        $wb = IOFactory::load(__DIR__ . '/test.xlsx');
        $sh = $wb->getSheet(0);

        $this->assertEquals(PhpSpreadsheetUtils::getHighestRowAndColumnIndex($sh), [4, 12]);
        $this->assertEquals(PhpSpreadsheetUtils::getCellValue($sh, 4, 12), 'Last');
    }

    /**
     * @throws PhpSpreadsheetException
     * @throws PhpSpreadsheetReaderException
     */
    public function testInMergeRange()
    {
        $wb = IOFactory::load(__DIR__ . '/test.xlsx');
        $sh = $wb->getSheet(0);

        $this->assertTrue(PhpSpreadsheetUtils::inMergeRange($sh, 2, 8));
        $this->assertTrue(PhpSpreadsheetUtils::inMergeRange($sh, 3, 9));
        $this->assertFalse(PhpSpreadsheetUtils::inMergeRange($sh, 1, 8));
        $this->assertFalse(PhpSpreadsheetUtils::inMergeRange($sh, 1, 9));
    }

}
