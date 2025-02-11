<?php
require '../vendor/autoload.php';
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

} else {
}
?>

<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>HTML to Excel</title>
    <link rel="stylesheet" href="style.css">
</head>
<body>
    <div class="container">
        <h1>Paste HTML Here</h1>
        <form method="POST">
            <textarea name="html_input" rows="10" cols="50" placeholder="Paste your HTML here..."></textarea><br><br>
            <button type="submit">Generate Excel</button>
            <a href="<?php echo $filename; ?>">Download Excel File</a>

        </form>

        <?php if (!empty($tableData)): ?>
            <h2>Extracted Data Preview:</h2> 
            <table>
                <thead>
                    <tr>
                        <th>Index</th>
                        <th>Port</th>
                        <th>Number</th>
                        <th>MAC Address</th>
                    </tr>
                </thead>
                <tbody>
                    <?php foreach ($tableData as $row): ?>
                        <tr>
                            <td><?php echo htmlspecialchars($row[0]); ?></td>
                            <td><?php echo htmlspecialchars($row[1]); ?></td>
                            <td><?php echo htmlspecialchars($row[2]); ?></td>
                            <td><?php echo htmlspecialchars($row[3]); ?></td>
                        </tr>
                    <?php endforeach; ?>
                </tbody>
            </table>
            <a href="<?php echo $filename; ?>">Download Excel File</a>
        <?php endif; ?>
    </div>
</body>
</html>
