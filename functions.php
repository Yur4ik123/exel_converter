<?php

use \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

function getRowsCount(Worksheet $originalItem): int {
    $emptyRows = 0;
    $countOfRows = 0;

    for ($i = 2; true; $i++) {
        $isNotNull = $originalItem->getCell([1, $i])->getValue();

        if ($isNotNull) {
            $countOfRows++;
            $emptyRows = 0;
        } else {
            $emptyRows++;
        }

        if ($emptyRows > 4) break;
    }

    return $countOfRows;
}

function processCompany (Worksheet $worksheet, Worksheet $newWorksheet, int $companyIndex, int $startRow, int $rowsCount): void {
    $hour = 0;
    $day = 1;
    $emptyRows = 0;
    $companyName = $worksheet->getCell([$companyIndex + 1, 1])->getValue();
    $newWorksheet->setCellValue([1, $startRow], $companyName);

    if (!$companyName) return;

    for ($rowIndex = 1; $rowIndex <= $rowsCount; $rowIndex++) {
        if ($emptyRows > 4) return;

        $companyHPosition = $companyIndex + 1;
        $colPosition = $rowIndex + 1;
        $date = $worksheet->getCell([1, $colPosition])->getValue();

        preg_match('/^([0-9][0-9]\.[0-9][0-9]\.[0-9][0-9][0-9][0-9]) (.*)$/', $date, $time);
        $rowDay = explode('.', $date)[0] * 1;

        if (!$date) {
            $emptyRows++;
            continue;
        } else {
            $emptyRows = 0;
        }

        if ($rowDay > $day) {
            $day++;
            $hour = 0;
        }

        $hour++;
        $newWorksheet->setCellValue([1, $startRow + $day], $time[1]);
        $newWorksheet->setCellValue([$hour + 1, $startRow], $time[2]);
        $value = explode(',', $worksheet->getCell([$companyHPosition, $colPosition])->getValue())[0];
        $newWorksheet->setCellValue([$hour + 1, $startRow + $day], $worksheet->getCell([$companyHPosition, $colPosition])->getValue());
    }

    processCompany($worksheet, $newWorksheet, $companyIndex + 1, $startRow + 33, $rowsCount);
}
