<?php

require 'vendor/autoload.php';
require 'functions.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx as WriterXlsx;
use PhpOffice\PhpSpreadsheet\IOFactory as Reader;

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    $spreadsheet = new Spreadsheet();

    try {
        // Read directly from PHP's managed temp file — no extra copy on disk
        $originalSheet = Reader::load($_FILES['file']['tmp_name']);
    } catch (Exception $r) {
        echo "<pre>";
        var_dump($r);
        exit;
    }

    $originalItem = $originalSheet->getActiveSheet();
    $rowsCount = getRowsCount($originalItem);

    try {
        processCompany($originalItem, $spreadsheet->getActiveSheet(), 1, 1, $rowsCount);
    } catch (Throwable $e) {
        var_dump($e);
        exit;
    }

    // Free the source spreadsheet from memory before building the output
    $originalSheet->disconnectWorksheets();
    unset($originalSheet, $originalItem);

    $writer = new WriterXlsx($spreadsheet);

    try {
        // Write into an in-memory stream — no file ever touches disk
        $stream = fopen('php://temp', 'wb+');
        $writer->save($stream);
        $size = ftell($stream);
        rewind($stream);

        header('Content-Disposition: attachment; filename="file.xlsx"');
        header('Content-Length: ' . $size);
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        fpassthru($stream);
        fclose($stream);
    } catch (\PhpOffice\PhpSpreadsheet\Writer\Exception $e) {
        var_dump($e->getMessage());
    } finally {
        $spreadsheet->disconnectWorksheets();
        unset($spreadsheet, $writer);
    }
}
