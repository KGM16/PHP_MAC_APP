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
        $interfaceNode = $xpath->query('.//td[contains(@class, "interface")]', $tr);
        $pvidNode = $xpath->query('.//td[contains(@class, "pvid")]', $tr);
        $taggedNode = $xpath->query('.//td[contains(@class, "tagged")]', $tr);
        $untaggedNode = $xpath->query('.//td[contains(@class, "untagged")]', $tr);

        if ($interfaceNode->length > 0 && $pvidNode->length > 0 && $taggedNode->length > 0 && $untaggedNode->length > 0) {
            $interface = $interfaceNode->item(0)->nodeValue;
            $pvid = $pvidNode->item(0)->nodeValue;
            $tagged = $taggedNode->item(0)->nodeValue;
            $untagged = $untaggedNode->item(0)->nodeValue;

            // Process tagged ports
            $taggedPorts = [];
            if (strpos($tagged, ',') !== false) {
                list($start, $end) = explode(',', $tagged);
                for ($i = (int)$start; $i <= (int)$end; $i++) {
                    $taggedPorts[] = $i;
                }
            } elseif (strpos($tagged, '-') !== false) {
                list($start, $end) = explode('-', $tagged);
                for ($i = (int)$start; $i <= (int)$end; $i++) {
                    $taggedPorts[] = $i;
                }
            } else {
                $taggedPorts[] = (int)$tagged;
            }

            // Process untagged ports
            $untaggedPorts = [];
            if (strpos($untagged, '-') !== false) {
                list($start, $end) = explode('-', $untagged);
                for ($i = (int)$start; $i <= (int)$end; $i++) {
                    $untaggedPorts[] = $i;
                }
            } else {
                $untaggedPorts[] = (int)$untagged;
            }

            $allPorts = array_unique(array_merge($taggedPorts, $untaggedPorts));
            foreach ($allPorts as $port) {
                if ($port === 0) {
                    continue;
                }
                $taggedUntagged = '';
                if (in_array($port, $taggedPorts)) {
                    $taggedUntagged .= 'T';
                }
                if (in_array($port, $untaggedPorts)) {
                    $taggedUntagged .= 'U';
                }
                $tableData[] = [$interface, $pvid, $port, $taggedUntagged];
            }
        }
    }

    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();

    $sheet->setCellValue('A1', 'Interface');
    $sheet->setCellValue('B1', 'PVID');
    $sheet->setCellValue('C1', 'Port');
    $sheet->setCellValue('D1', 'Tagged/Untagged');

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
                        <th>Interface</th>
                        <th>PVID</th>
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