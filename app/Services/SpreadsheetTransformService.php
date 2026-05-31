<?php

namespace App\Services;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

class SpreadsheetTransformService
{
    public function transform(string $sourcePath, string $outputPath): void
    {
        $source = IOFactory::load($sourcePath);
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

    private function writeRow($sourceSheet, $targetSheet, int $sourceRow, int $targetRow): void
    {
        foreach (config('mappings.spreadsheet_transform') as $targetCol => $sourceCol) {
            $targetSheet->setCellValue(
                "{$targetCol}{$targetRow}",
                $sourceSheet->getCell("{$sourceCol}{$sourceRow}")->getValue()
            );
        }
    }

    private function shouldSkipRow($sheet, int $row): bool
    {
        return $sheet->getCell('C' . $row)->getValue() == ''
            && $sheet->getCell('C' . ($row - 1))->getValue() == '';
    }

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

    private function save(Spreadsheet $spreadsheet, string $path): void
    {
        $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
        $writer->save($path);
    }
}