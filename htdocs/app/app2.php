<?php
require '../vendor/autoload.php';
include ('index.html');
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$tableData = [];

if ($_SERVER['REQUEST_METHOD'] === 'POST' && !empty($_POST['html_input'])) {
    $html = $_POST['html_input'];

    $dom = new DOMDocument();
    @$dom->loadHTML($html); 

    $xpath = new DOMXPath($dom);

    $trNodes = $xpath->query('//tr');

    foreach ($trNodes as $tr) {
        $tdNodes = $xpath->query('.//td', $tr); 
        if ($tdNodes->length >= 4) {
            $index = $tdNodes->item(0)->nodeValue; 
            $port = $tdNodes->item(1)->nodeValue; 
            $number = $tdNodes->item(2)->nodeValue; 
            $mac = $tdNodes->item(3)->nodeValue; 
            $tableData[] = [$index, $port, $number, $mac];
        }
    }

    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();

    $sheet->setCellValue('A1', 'Index');
    $sheet->setCellValue('B1', 'Port');
    $sheet->setCellValue('C1', 'Number');
    $sheet->setCellValue('D1', 'MAC Address');

    $row = 2; 

    foreach ($tableData as $dataRow) {
        $sheet->setCellValue('A' . $row, $dataRow[0]);
        $sheet->setCellValue('B' . $row, $dataRow[1]);
        $sheet->setCellValue('C' . $row, $dataRow[2]);
        $sheet->setCellValue('D' . $row, $dataRow[3]);
        $row++;
    }

    $writer = new Xlsx($spreadsheet);
    $filename = 'output.xlsx';
    $writer->save($filename);

}


