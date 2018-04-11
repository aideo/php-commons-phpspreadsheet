<?php

namespace Ideo\Utilities\PhpOffice;

use Ideo\Utilities\StringUtils;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Exception as PhpSpreadsheetException;
use PhpOffice\PhpSpreadsheet\NamedRange;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

/**
 * PhpSpreadsheet 使用時にあると便利な処理を提供します。
 *
 * @package Ideo\Utilities\PHPOffice
 */
class PhpSpreadsheetUtils
{

    /**
     * ワークシートの列・行を指定してセルの値を取得します。
     *
     * @param Worksheet $sheet ワークシート。
     * @param int|string $columnIndex 列インデックス (数値の場合 1 開始)。
     * @param int|string $rowIndex 行番号 (数値の場合 1 開始)。
     * @param string|null $defaultValue セルが有効ではない場合のデフォルト値。
     *
     * @return null|string セルの値。
     */
    public static function getCellValue(Worksheet $sheet, $columnIndex = 1, $rowIndex = 1, $defaultValue = null)
    {
        if ($sheet->cellExistsByColumnAndRow($columnIndex, $rowIndex)) {
            $cell = $sheet->getCellByColumnAndRow($columnIndex, $rowIndex);

            return StringUtils::trim($cell->getFormattedValue());
        } else {
            return $defaultValue;
        }
    }

    /**
     * ワークシートから、名前付きセル名を指定して値を取得します。
     *
     * @param Worksheet $sheet ワークシート。
     * @param string $name 名前付きセルの名前。
     * @param string|null $defaultValue 名前付きセルが有効ではない場合のデフォルト値。
     *
     * @return null|string セルの値。
     * @throws PhpSpreadsheetException セルの取得時に例外が発生した場合。
     */
    public static function getCellValueByName(Worksheet $sheet, $name, $defaultValue = null)
    {
        $namedRange = NamedRange::resolveRange($name, $sheet);

        if ($namedRange !== null && $sheet->cellExists(($range = $namedRange->getRange()))) {
            $cell = $sheet->getCell($range);

            return StringUtils::trim($cell->getFormattedValue());
        } else {
            return $defaultValue;
        }
    }

    /**
     * ワークシートの最大列と最大行をインデックスで取得します。
     *
     * @param Worksheet $sheet 対象のワークシート。
     *
     * @return array 0 要素目に最大列のインデックス、1 要素目に最大行のインデックスが格納された配列。
     * @throws PhpSpreadsheetException 列インデックスを取得する際に例外が発生した場合。
     */
    public static function getHighestRowAndColumnIndex(Worksheet $sheet)
    {
        $highestRowAndColumn = $sheet->getHighestRowAndColumn();

        $columnIndex = Coordinate::columnIndexFromString($highestRowAndColumn['column']);
        $rowIndex = $highestRowAndColumn['row'];

        return [$columnIndex, $rowIndex];
    }

    /**
     * 指定したセルがマージセルかどうかを取得します。
     *
     * @param Worksheet $sheet ワークシート。
     * @param int|string $columnIndex 列インデックス (数値の場合 1 開始)。
     * @param int|string $rowIndex 行番号 (数値の場合 1 開始)。
     *
     * @return bool マージセルの場合 true, それ以外は false 。
     */
    public static function inMergeRange(Worksheet $sheet, $columnIndex, $rowIndex)
    {
        $cell = $sheet->getCellByColumnAndRow($columnIndex, $rowIndex);

        return $cell->isInMergeRange();
    }

    /**
     * 対象シートの名前付きセルの定義をクリアします。
     *
     * @param Spreadsheet $workbook 対象のワークブック。
     * @param Worksheet $sheet 対象のシート。
     */
    public static function removeNamedRangesBySheet(Spreadsheet $workbook, Worksheet $sheet)
    {
        $namedRanges = $workbook->getNamedRanges();

        foreach ($namedRanges as $key => $namedRange) {
            if ($namedRange->getWorksheet() !== $sheet) {
                continue;
            }

            $workbook->removeNamedRange($namedRange->getName(), $sheet);
        }
    }

    /**
     * ワークシートの列・行を指定してセルの値を設定します。
     *
     * @param Worksheet $sheet ワークシート。
     * @param int|string $columnIndex 列インデックス (数値の場合 1 開始)。
     * @param int|string $rowIndex 行番号 (数値の場合 1 開始)。
     * @param int|string $value 設定するセルの値。
     *
     * @throws PhpSpreadsheetException 指定されたセルが有効ではない場合。
     */
    public static function setCellValue(Worksheet $sheet, $columnIndex = 1, $rowIndex = 1, $value)
    {
        if (!$sheet->cellExistsByColumnAndRow($columnIndex, $rowIndex)) {
            throw new PhpSpreadsheetException('The specified cell is not found.');
        }

        $cell = $sheet->getCellByColumnAndRow($columnIndex, $rowIndex);
        $cell->setValue($value);
    }

    /**
     * 対象のセルがマージセルに含まれる場合、そのマージセル全体をキャンセルします。
     *
     * @param Worksheet $sheet ワークシート。
     * @param int|string $columnIndex 列インデックス (数値の場合 1 開始)。
     * @param int|string $rowIndex 行番号 (数値の場合 1 開始)。
     *
     * @throws PhpSpreadsheetException マージセルのキャンセルに失敗した場合。
     */
    public static function unmergeContainedCells(Worksheet $sheet, $columnIndex, $rowIndex)
    {
        $cell = $sheet->getCellByColumnAndRow($columnIndex, $rowIndex);

        if ($cell->isInMergeRange()) {
            $sheet->unmergeCells($cell->getMergeRange());
        }
    }

}
