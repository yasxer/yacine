<?php
error_reporting(E_ALL);
ini_set('display_errors', 1);

require __DIR__ . '/vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Reader\Xls;
use PhpOffice\PhpSpreadsheet\Shared\Date as ExcelDate;
use Dompdf\Dompdf;
use Dompdf\Options;

// Simple helper to escape output for HTML
function h($s)
{
    return htmlspecialchars($s, ENT_QUOTES | ENT_SUBSTITUTE, 'UTF-8');
}

if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
    echo 'Use the upload form.';
    exit;
}

if (!isset($_FILES['xls_file']) || $_FILES['xls_file']['error'] !== UPLOAD_ERR_OK) {
    echo 'File upload error. Please provide a valid .xls file.';
    exit;
}

$uploadedName = $_FILES['xls_file']['name'];
$ext = strtolower(pathinfo($uploadedName, PATHINFO_EXTENSION));
if ($ext !== 'xls') {
    echo 'Unsupported file type. Please upload a .xls file.';
    exit;
}

$tmpFile = $_FILES['xls_file']['tmp_name'];

// Read spreadsheet
$reader = new Xls();
$reader->setReadDataOnly(true);
try {
    $spreadsheet = $reader->load($tmpFile);
} catch (Exception $e) {
    echo 'Could not read Excel file: ' . h($e->getMessage());
    exit;
}

$sheet = $spreadsheet->getActiveSheet();

$report_type = $_POST['report_type'] ?? 'overdue';
$today = new DateTimeImmutable('today');

// Determine min/max date based on report type (relative to TODAY)
switch ($report_type) {
    case '30_37':
        $minPaymentDate = $today->sub(new DateInterval('P37D'));
        $maxPaymentDate = $today->sub(new DateInterval('P30D'));
        $label = 'Early Warning (30-37 days)';
        break;
    case '37_44':
        $minPaymentDate = $today->sub(new DateInterval('P44D'));
        $maxPaymentDate = $today->sub(new DateInterval('P37D'));
        $label = 'Advanced Warning (37-44 days)';
        break;
    case '24_31':
        $minPaymentDate = $today->sub(new DateInterval('P31D'));
        $maxPaymentDate = $today->sub(new DateInterval('P24D'));
        $label = 'Early Warning (24-31 days)';
        break;
    case 'overdue':
    default:
        // Over 30 days overdue: any last payment date on or before (today - 30 days)
        $minPaymentDate = new DateTimeImmutable('1900-01-01');
        $maxPaymentDate = $today->sub(new DateInterval('P30D'));
        $label = 'Over 30 days overdue (overdue)';
        break;
}

$results = [];
$highestRow = $sheet->getHighestRow();

for ($row = 2; $row <= $highestRow; $row++) {
    $clientCode = trim((string)$sheet->getCell('A' . $row)->getValue());
    $clientName = trim((string)$sheet->getCell('B' . $row)->getValue());
    $contact = trim((string)$sheet->getCell('C' . $row)->getValue());
    $soldeRaw = $sheet->getCell('E' . $row)->getValue();
    $lastPaymentRaw = $sheet->getCell('F' . $row)->getValue();

    if ($contact === '') {
        $contact = 'inconnu';
    }

    // Normalize solde to float (handle comma separators and NBSP)
    $solde = 0.0;
    if ($soldeRaw !== null && $soldeRaw !== '') {
        $s = str_replace(["\xC2\xA0", ' '], '', (string)$soldeRaw);
        $s = str_replace(',', '.', $s);
        $solde = floatval($s);
    }

    // Secondary filter: solde must be > 0
    if ($solde <= 0) continue;

    // Parse last payment date (column F)
    $lastPaymentDate = null;
    if ($lastPaymentRaw === null || $lastPaymentRaw === '') {
        // skip rows without last payment date
        continue;
    }

    if (is_numeric($lastPaymentRaw)) {
        // Excel numeric date
        try {
            $dt = ExcelDate::excelToDateTimeObject($lastPaymentRaw);
            $lastPaymentDate = DateTimeImmutable::createFromMutable($dt);
        } catch (Exception $e) {
            $lastPaymentDate = null;
        }
    } else {
        // Try strtotime then common formats
        $parsed = strtotime($lastPaymentRaw);
        if ($parsed !== false) {
            $lastPaymentDate = (new DateTimeImmutable())->setTimestamp($parsed);
        } else {
            $d1 = DateTime::createFromFormat('d/m/Y', $lastPaymentRaw);
            if ($d1) $lastPaymentDate = DateTimeImmutable::createFromMutable($d1);
            else {
                $d2 = DateTime::createFromFormat('Y-m-d', $lastPaymentRaw);
                if ($d2) $lastPaymentDate = DateTimeImmutable::createFromMutable($d2);
            }
        }
    }

    if (!$lastPaymentDate) continue;

    // Primary filter: last payment date must fall within min/max (inclusive)
    if ($minPaymentDate && $lastPaymentDate < $minPaymentDate) continue;
    if ($maxPaymentDate && $lastPaymentDate > $maxPaymentDate) continue;

    $results[] = [
        'code' => $clientCode,
        'name' => $clientName,
        'contact' => $contact,
        // keep numeric raw for totals and formatted string for display
        'solde_raw' => $solde,
        'solde' => number_format($solde, 2, ',', ' '),
        'last_payment' => $lastPaymentDate->format('Y-m-d'),
    ];
}

// Build LTR HTML table
$rangeLabel = $minPaymentDate->format('Y-m-d') . ' → ' . $maxPaymentDate->format('Y-m-d');
$html = '<!doctype html><html lang="en" dir="ltr"><head><meta charset="utf-8"><style>';
$html .= 'body{font-family: "DejaVu Sans", Tahoma, Arial; direction:ltr; color:#222;}
header{display:flex;justify-content:space-between;align-items:center;margin-bottom:12px}
.brand{font-size:18px;font-weight:700;color:#2468f2}
.sub{color:#666;font-size:12px}
.card{background:#fff;padding:12px;border-radius:8px;box-shadow:0 6px 18px rgba(36,104,242,0.06)}
table{border-collapse:collapse; width:100%;}
th,td{border:1px solid #e6eefc; padding:10px; text-align:left; font-size:12px}
th{background:#f6fbff;color:#123;padding:12px 10px}
tbody tr:nth-child(odd){background:#fbfdff}
.totals{font-weight:700;background:#f1f6ff}
.meta{font-size:12px;color:#444;margin-bottom:8px}
';
$html .= '</style></head><body>';
$html .= '<header><div><div class="brand">' . h($label) . '</div><div class="sub">Generated: ' . h((new DateTime())->format('Y-m-d')) . ' &nbsp; | &nbsp; Date range: ' . h($rangeLabel) . '</div></div><div><img src="" alt="" style="width:84px;opacity:.15"></div></header>';

if (empty($results)) {
    $html .= '<p>No records matched the chosen criteria.</p>';
} else {
    $html .= '<table class="card"><thead><tr>';
    $html .= '<th>Client Code</th><th>Name</th><th>Contact</th><th>Solde</th><th>Date of Last Payment</th>';
    $html .= '</tr></thead><tbody>';
    foreach ($results as $r) {
        $html .= '<tr>';
        $html .= '<td>' . h($r['code']) . '</td>';
        $html .= '<td>' . h($r['name']) . '</td>';
        $html .= '<td>' . h($r['contact']) . '</td>';
        $html .= '<td>' . h($r['solde']) . '</td>';
        $html .= '<td>' . h($r['last_payment']) . '</td>';
        $html .= '</tr>';
    }
    // totals
    $total = 0.0;
    foreach ($results as $r) {
        $total += floatval($r['solde_raw'] ?? 0);
    }
    $html .= '<tr class="totals"><td colspan="3">Total outstanding</td><td>' . h(number_format($total, 2, ',', ' ')) . '</td><td></td></tr>';
    $html .= '</tbody></table>';
}

$html .= '</body></html>';

// Generate PDF via Dompdf
$options = new Options();
$options->set('isRemoteEnabled', true);
$options->set('defaultFont', 'DejaVu Sans');

$dompdf = new Dompdf($options);
$dompdf->loadHtml($html);
$dompdf->setPaper('A4', 'portrait');
$dompdf->render();

// Add simple page numbers in footer using canvas
try {
    $canvas = $dompdf->get_canvas();
    $font = $dompdf->getFontMetrics()->get_font('DejaVu Sans', 'normal');
    $w = $canvas->get_width();
    $h = $canvas->get_height();
    $text = "Page {PAGE_NUM} / {PAGE_COUNT}";
    // place near bottom-right (adjust offsets if needed)
    $canvas->page_text($w - 80, $h - 30, $text, $font, 10, array(0, 0, 0));
} catch (Exception $e) {
    // ignore footer failures
}

// Stream to browser as download
$dompdf->stream('overdue_report.pdf', ["Attachment" => 1]);
exit;
