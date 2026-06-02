<?php

namespace App\Services;

use App\Filters\MappingReadFilter;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class SpreadsheetTransformService
{
    /**
     * @param string $sourcePath
     * @param string $outputPath
     * @return void
     */
    public function transform(string $sourcePath, string $outputPath): void
    {
        $reader = IOFactory::createReader('Xlsx');
        $reader->setReadDataOnly(true);
        $reader->setReadFilter(
            new MappingReadFilter(array_values(config('mappings.spreadsheet_transform')))
        );
        $source = $reader->load($sourcePath);
        $worksheet = $source->getSheet(0);

        $lastRow = $worksheet->getHighestRow();

        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();

        $this->setMeta($spreadsheet);
        $sheet->setTitle('Something6');

        $this->writeRow($worksheet, $sheet, 2, 2);

        $newRow = 3;

        for ($row = 3; $row <= $lastRow; $row++) {
            if ($this->shouldSkipRow($worksheet, $row)) {
                continue;
            }

            $this->writeRow($worksheet, $sheet, $row, $newRow);
            $newRow++;
        }

        $this->save($spreadsheet, $outputPath);
    }

    /**
     * @param Worksheet $sourceSheet
     * @param Worksheet $targetSheet
     * @param int $sourceRow
     * @param int $targetRow
     * @return void
     */
    private function writeRow(Worksheet $sourceSheet, Worksheet $targetSheet, int $sourceRow, int $targetRow): void
    {
        // if C column is empty, ignore whole row and enter empty values
        $isEmpty = $sourceSheet->getCell('C' . $sourceRow)->getValue() == '';

        foreach (config('mappings.spreadsheet_transform') as $targetCol => $sourceCol) {
            $targetSheet->setCellValue(
                "{$targetCol}{$targetRow}",
                !$isEmpty ? $sourceSheet->getCell("{$sourceCol}{$sourceRow}")->getValue() : ''
            );
        }
    }

    /**
     * Checks if row should be skipped, mainly for double empty rows
     * 
     * @param Worksheet $sheet
     * @param int $row
     * @return bool
     */
    private function shouldSkipRow(Worksheet $sheet, int $row): bool
    {
        return $sheet->getCell('C' . $row)->getValue() == ''
            && $sheet->getCell('C' . ($row - 1))->getValue() == '';
    }

    /**
     * @param Spreadsheet $spreadsheet
     * @return void
     */
    private function setMeta(Spreadsheet $spreadsheet): void
    {
        $spreadsheet->getProperties()
            ->setCreator("Something1")
            ->setLastModifiedBy("Something1")
            ->setTitle("Something1")
            ->setSubject("Something2")
            ->setDescription("Something3")
            ->setKeywords("Something4")
            ->setCategory("Something5");
    }

    /**
     * @param Spreadsheet $spreadsheet
     * @param string $path
     * @return void
     */
    private function save(Spreadsheet $spreadsheet, string $path): void
    {
        $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
        $writer->save($path);
    }
}