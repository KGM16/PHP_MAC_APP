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
        $vlanNode = $xpath->query('.//td[@id[starts-with(., "lblVLAN_")]]', $tr);
        $vlan = $vlanNode->length > 0 ? trim($vlanNode->item(0)->nodeValue) : '';

        $macNode = $xpath->query('.//td[contains(., ":")]', $tr);
        if ($macNode->length > 0) {
            $macTd = $macNode->item(0);
            foreach ($xpath->query('.//script', $macTd) as $script) {
                $script->parentNode->removeChild($script);
            }
            $mac = trim($macTd->nodeValue);
        } else {
            $mac = '';
        }

        $interfaceNode = $xpath->query('.//td[contains(., "GE")]', $tr);
        if ($interfaceNode->length > 0) {
            $interfaceTd = $interfaceNode->item(0);
            foreach ($xpath->query('.//script', $interfaceTd) as $script) {
                $script->parentNode->removeChild($script);
            }
            $interface = trim($interfaceTd->nodeValue);
        } else {
            $interface = '';
        }

        if (!empty($vlan) && !empty($mac) && !empty($interface)) {
            $tableData[] = [$vlan, $mac, $interface];
        }
    }

    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();

    $sheet->setCellValue('A1', 'VLAN');
    $sheet->setCellValue('B1', 'MAC Address');
    $sheet->setCellValue('C1', 'Interface');

    $row = 2; 

    foreach ($tableData as $dataRow) {
        $sheet->setCellValue('A' . $row, $dataRow[0]);
        $sheet->setCellValue('B' . $row, $dataRow[1]);
        $sheet->setCellValue('C' . $row, $dataRow[2]);
        $row++;
    }

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
            <a href="<?php echo $filename; ?>">Download Excel File</a>
        </form>

        <?php if (!empty($tableData)): ?>
            <h2>Extracted Data Preview:</h2> 
            <table>
                <thead>
                    <tr>
                        <th>VLAN</th>
                        <th>MAC Address</th>
                        <th>Interface</th>
                    </tr>
                </thead>
                <tbody>
                    <?php foreach ($tableData as $row): ?>
                        <tr>
                            <td><?php echo htmlspecialchars($row[0]); ?></td>
                            <td><?php echo htmlspecialchars($row[1]); ?></td>
                            <td><?php echo htmlspecialchars($row[2]); ?></td>
                        </tr>
                    <?php endforeach; ?>
                </tbody>
            </table>
            <a href="<?php echo $filename; ?>">Download Excel File</a>
        <?php endif; ?>
    </div>
</body>
</html>