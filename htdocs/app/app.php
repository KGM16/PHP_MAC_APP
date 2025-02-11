<?php
require '../vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

if ($_SERVER['REQUEST_METHOD'] === 'POST' && !empty($_POST['html_input'])) {
    $html = $_POST['html_input'];

    $dom = new DOMDocument();
    @$dom->loadHTML($html); 

    $xpath = new DOMXPath($dom);

    $trNodes = $xpath->query('//tr');

    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();

    $sheet->setCellValue('A1', 'Index');
    $sheet->setCellValue('B1', 'Port');
    $sheet->setCellValue('C1', 'Number');
    $sheet->setCellValue('D1', 'MAC Address');

    $row = 2; 
    foreach ($trNodes as $tr) {
        $tdNodes = $xpath->query('.//td', $tr); 

        if ($tdNodes->length >= 4) {
            $index = $tdNodes->item(0)->nodeValue; 
            $port = $tdNodes->item(1)->nodeValue; 
            $number = $tdNodes->item(2)->nodeValue; 
            $mac = $tdNodes->item(3)->nodeValue; 

            // Write data to the Excel sheet
            $sheet->setCellValue('A' . $row, $index);
            $sheet->setCellValue('B' . $row, $port);
            $sheet->setCellValue('C' . $row, $number);
            $sheet->setCellValue('D' . $row, $mac);

            $row++;
        }
    }

    $writer = new Xlsx($spreadsheet);
    $filename = 'output.xlsx';
    $writer->save($filename);

    echo "Excel file generated successfully: <a href='$filename'>Download</a>";
} else {
    echo 'Please provide HTML input.';
}
?>

<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>HTML to Excel</title>
</head>
<body>
    <h1>Paste HTML Here</h1>
    <form method="POST">
        <textarea name="html_input" rows="10" cols="50" placeholder="Paste your HTML here..."></textarea><br><br>
        <button type="submit">Generate Excel</button>
    </form>
</body>
</html>