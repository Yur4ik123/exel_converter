<?php

require 'vendor/autoload.php';
require 'functions.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx as WriterXlsx;
use PhpOffice\PhpSpreadsheet\IOFactory as Reader;

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    $filename = basename($_FILES['file']['name']);
    move_uploaded_file($_FILES['file']['tmp_name'], $filename);

    $spreadsheet = new Spreadsheet();

    try {
        $originalSheet = Reader::load($filename);
    } catch (Exception $r) {
        echo "<pre>";
        var_dump($r);
    }

    $originalItem = $originalSheet->getActiveSheet();

    $rowsCount = getRowsCount($originalItem);

    $day = 1;
    $companyIndex = 1;
    $emptyRows = 0;

    try {
        processCompany($originalItem, $spreadsheet->getActiveSheet(), $companyIndex, 1, $rowsCount);
    } catch (Throwable $e) {
        var_dump($e);
    }

    $writer = new WriterXlsx($spreadsheet);

    try {
        $filename = 'file.xlsx';
        $writer->save($filename);

        if (file_exists($filename)) {
            $file = file_get_contents($filename);
            $size = strlen($file);
            header("Content-Disposition: attachment; filename = $filename");
            header("Content-Length: $size");
            header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            echo $file;
        }
    } catch (\PhpOffice\PhpSpreadsheet\Writer\Exception $e) {
        var_dump($e->getMessage());
    }
}
