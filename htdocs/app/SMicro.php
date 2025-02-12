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

    $trNodes = $xpath->query('//tbody/tr');

    foreach ($trNodes as $tr) {
        $tdNodes = $xpath->query('.//td', $tr); 
        if ($tdNodes->length >= 7) { // Ensure there are enough columns
            $mac = trim($tdNodes->item(3)->nodeValue); 
            $vid = trim($tdNodes->item(4)->nodeValue); 
            $onInterface = trim($tdNodes->item(5)->nodeValue); 
            $bridge = trim($tdNodes->item(6)->nodeValue); 
            $tableData[] = [$mac, $vid, $onInterface, $bridge];
        }
    }

    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();

    // Set headers
    $sheet->setCellValue('A1', 'MAC Address');
    $sheet->setCellValue('B1', 'VID');
    $sheet->setCellValue('C1', 'On Interface');
    $sheet->setCellValue('D1', 'Bridge');

    // Insert data
    $row = 2; 
    foreach ($tableData as $dataRow) {
        $sheet->setCellValue('A' . $row, $dataRow[0]);
        $sheet->setCellValue('B' . $row, $dataRow[1]);
        $sheet->setCellValue('C' . $row, $dataRow[2]);
        $sheet->setCellValue('D' . $row, $dataRow[3]);
        $row++;
    }

    // Save the spreadsheet
    $writer = new Xlsx($spreadsheet);
    $filename = 'output.xlsx';
    $writer->save($filename);

} else {
    $filename = '';
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
            <?php if ($filename): ?>
                <a href="<?php echo $filename; ?>">Download Excel File</a>
            <?php endif; ?>
        </form>

        <?php if (!empty($tableData)): ?>
            <h2>Extracted Data Preview:</h2> 
            <table>
                <thead>
                    <tr>
                        <th>MAC Address</th>
                        <th>VID</th>
                        <th>On Interface</th>
                        <th>Bridge</th>
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