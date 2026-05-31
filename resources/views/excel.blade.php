<?php

require '../vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

// Load source workbook
$sourceSpreadsheet = IOFactory::load("files/DB2.xlsx");
$worksheet = $sourceSpreadsheet->getSheet(0);
$lastRow = $worksheet->getHighestRow();

// Create new spreadsheet
$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();

// Set document properties
$spreadsheet->getProperties()
    ->setCreator("Something1")
    ->setLastModifiedBy("Something1")
    ->setTitle("Something1")
    ->setSubject("Something2")
    ->setDescription("Something3")
    ->setKeywords("Something4")
    ->setCategory("Something5");

$sheet->setTitle('Something6');


// ---- First row (row 2) ----
$row = 2;

$sheet->setCellValue('A2', $worksheet->getCell('CC' . $row)->getValue());
$sheet->setCellValue('B2', $worksheet->getCell('C' . $row)->getValue());
$sheet->setCellValue('C2', $worksheet->getCell('D' . $row)->getValue());
$sheet->setCellValue('D2', $worksheet->getCell('E' . $row)->getValue());
$sheet->setCellValue('E2', $worksheet->getCell('F' . $row)->getValue());
$sheet->setCellValue('F2', $worksheet->getCell('G' . $row)->getValue());
$sheet->setCellValue('G2', $worksheet->getCell('H' . $row)->getValue());
$sheet->setCellValue('H2', $worksheet->getCell('I' . $row)->getValue());
$sheet->setCellValue('I2', $worksheet->getCell('J' . $row)->getValue());
$sheet->setCellValue('J2', $worksheet->getCell('K' . $row)->getValue());
$sheet->setCellValue('K2', $worksheet->getCell('L' . $row)->getValue());
$sheet->setCellValue('L2', $worksheet->getCell('M' . $row)->getValue());
$sheet->setCellValue('M2', $worksheet->getCell('N' . $row)->getValue());
$sheet->setCellValue('N2', $worksheet->getCell('O' . $row)->getValue());
$sheet->setCellValue('O2', $worksheet->getCell('P' . $row)->getValue());
$sheet->setCellValue('P2', $worksheet->getCell('Q' . $row)->getValue());
$sheet->setCellValue('Q2', $worksheet->getCell('R' . $row)->getValue());
$sheet->setCellValue('R2', $worksheet->getCell('S' . $row)->getValue());
$sheet->setCellValue('S2', $worksheet->getCell('T' . $row)->getValue());
$sheet->setCellValue('T2', $worksheet->getCell('Y' . $row)->getValue());
$sheet->setCellValue('U2', $worksheet->getCell('Z' . $row)->getValue());
$sheet->setCellValue('V2', $worksheet->getCell('AA' . $row)->getValue());
$sheet->setCellValue('W2', $worksheet->getCell('AB' . $row)->getValue());
$sheet->setCellValue('X2', $worksheet->getCell('AC' . $row)->getValue());
$sheet->setCellValue('Y2', $worksheet->getCell('AD' . $row)->getValue());
$sheet->setCellValue('Z2', $worksheet->getCell('AE' . $row)->getValue());
$sheet->setCellValue('AA2', $worksheet->getCell('AF' . $row)->getValue());


// ---- Loop rows ----
$row = 3;
$newrow = 3;

for (; $row <= $lastRow; $row++) {

    $rowbefore = $row - 1;

    if (
        $worksheet->getCell('C' . $row)->getValue() == "" &&
        $worksheet->getCell('C' . $rowbefore)->getValue() == ""
    ) {
        continue;
    }

    $sheet->setCellValue('A' . $newrow, $worksheet->getCell('CC' . $row)->getValue());
    $sheet->setCellValue('B' . $newrow, $worksheet->getCell('C' . $row)->getValue());
    $sheet->setCellValue('C' . $newrow, $worksheet->getCell('D' . $row)->getValue());
    $sheet->setCellValue('D' . $newrow, $worksheet->getCell('E' . $row)->getValue());
    $sheet->setCellValue('E' . $newrow, $worksheet->getCell('F' . $row)->getValue());
    $sheet->setCellValue('F' . $newrow, $worksheet->getCell('G' . $row)->getValue());
    $sheet->setCellValue('G' . $newrow, $worksheet->getCell('H' . $row)->getValue());
    $sheet->setCellValue('H' . $newrow, $worksheet->getCell('I' . $row)->getValue());
    $sheet->setCellValue('I' . $newrow, $worksheet->getCell('J' . $row)->getValue());
    $sheet->setCellValue('J' . $newrow, $worksheet->getCell('K' . $row)->getValue());
    $sheet->setCellValue('K' . $newrow, $worksheet->getCell('L' . $row)->getValue());
    $sheet->setCellValue('L' . $newrow, $worksheet->getCell('M' . $row)->getValue());
    $sheet->setCellValue('M' . $newrow, $worksheet->getCell('N' . $row)->getValue());
    $sheet->setCellValue('N' . $newrow, $worksheet->getCell('O' . $row)->getValue());
    $sheet->setCellValue('O' . $newrow, $worksheet->getCell('P' . $row)->getValue());
    $sheet->setCellValue('P' . $newrow, $worksheet->getCell('Q' . $row)->getValue());
    $sheet->setCellValue('Q' . $newrow, $worksheet->getCell('R' . $row)->getValue());
    $sheet->setCellValue('R' . $newrow, $worksheet->getCell('S' . $row)->getValue());
    $sheet->setCellValue('S' . $newrow, $worksheet->getCell('T' . $row)->getValue());
    $sheet->setCellValue('T' . $newrow, $worksheet->getCell('Y' . $row)->getValue());
    $sheet->setCellValue('U' . $newrow, $worksheet->getCell('Z' . $row)->getValue());
    $sheet->setCellValue('V' . $newrow, $worksheet->getCell('AA' . $row)->getValue());
    $sheet->setCellValue('W' . $newrow, $worksheet->getCell('AB' . $row)->getValue());
    $sheet->setCellValue('X' . $newrow, $worksheet->getCell('AC' . $row)->getValue());
    $sheet->setCellValue('Y' . $newrow, $worksheet->getCell('AD' . $row)->getValue());
    $sheet->setCellValue('Z' . $newrow, $worksheet->getCell('AE' . $row)->getValue());
    $sheet->setCellValue('AA' . $newrow, $worksheet->getCell('AF' . $row)->getValue());

    $newrow++;
}

// ---- Save file ----
$writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
$path = storage_path('app/MyExcel.xlsx');
$writer->save($path);