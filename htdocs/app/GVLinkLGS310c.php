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
        $tdNodes = $xpath->query('.//td[contains(@class, "h7")]', $tr); 
        if ($tdNodes->length >= 4) {
            $vlan = $tdNodes->item(1)->nodeValue; 
            $tagged = $tdNodes->item(3)->nodeValue; 
            $untagged = $tdNodes->item(4)->nodeValue; 

            // Process tagged ports
            $taggedPorts = processPorts($tagged);

            // Process untagged ports
            $untaggedPorts = processPorts($untagged);

            // Combine tagged and untagged ports
            $allPorts = array_unique(array_merge($taggedPorts, $untaggedPorts));
            if ($vlan == 'VLAN') {
                continue;
            }
            foreach ($allPorts as $port) {
                if ($vlan == 'VLAN' || $port == 0 || $vlan == 0) {
                    continue;
                }
                $taggedUntagged = '';
                if (in_array($port, $taggedPorts)) {
                    $taggedUntagged .= 'T';
                }
                if (in_array($port, $untaggedPorts)) {
                    $taggedUntagged .= 'U';
                }
                $tableData[] = [$vlan, $port, $taggedUntagged];
            }
        }
    }

    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();

    $sheet->setCellValue('A1', 'VLAN');
    $sheet->setCellValue('B1', 'Port');
    $sheet->setCellValue('C1', 'Tagged/Untagged');

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
}

function processPorts($ports) {
    $portArray = [];
    $portParts = explode(',', $ports);
    foreach ($portParts as $part) {
        if (strpos($part, '-') !== false) {
            list($start, $end) = explode('-', $part);
            for ($i = (int)$start; $i <= (int)$end; $i++) {
                $portArray[] = $i;
            }
        } else {
            $portArray[] = (int)$part;
        }
    }
    return $portArray;
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
                        <th>Port</th>
                        <th>Tagged/Untagged</th>
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