<?php
require __DIR__ . '/includes/auth.php';
require_login();
require __DIR__ . '/config/db.php';
require __DIR__ . '/includes/tables_config.php';
require __DIR__ . '/includes/functions.php';
require __DIR__ . '/includes/xlsx_reader.php';

[$key, $cfg] = resolve_table($TABLES, $_GET['table'] ?? null);

$result = null; // ['inserted' => int, 'skipped' => int, 'notes' => [...], 'mode' => 'flat'|'summary']
$error  = null;

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    if (empty($_FILES['file']) || $_FILES['file']['error'] !== UPLOAD_ERR_OK) {
        $error = 'Please choose a file to upload.';
    } else {
        $tmpPath = $_FILES['file']['tmp_name'];
        $origName = $_FILES['file']['name'];
        $ext = strtolower(pathinfo($origName, PATHINFO_EXTENSION));

        try {
            if ($ext === 'csv') {
                $table = csv_read_as_table($tmpPath);
            } elseif ($ext === 'xlsx') {
                $table = xlsx_read_first_sheet_as_table($tmpPath);
            } elseif ($ext === 'xls') {
                throw new Exception('The old .xls format is not supported — please save the file as .xlsx or .csv and try again.');
            } else {
                throw new Exception('Unsupported file type. Please upload a .xlsx or .csv file.');
            }

            $headerRow = $table[0] ?? [];
            $headerMap = map_headers_to_keys($cfg, $headerRow);

            if (count($headerMap) >= 1) {
                // ---------- flat header-matched import ----------
                $inserted = 0;
                $notes = [];
                for ($i = 1; $i < count($table); $i++) {
                    $row = $table[$i];
                    $isBlank = true;
                    foreach ($row as $v) { if (trim((string)$v) !== '') { $isBlank = false; break; } }
                    if ($isBlank) continue;

                    $fieldValues = [];
                    $outcomeValues = [];
                    foreach ($headerMap as $colIdx => $mappedKey) {
                        $val = trim((string)($row[$colIdx] ?? ''));
                        if (in_array($mappedKey, $cfg['outcomes'], true)) {
                            $outcomeValues[$mappedKey] = $val === '' ? null : $val;
                        } else {
                            $fieldValues[$mappedKey] = $val;
                        }
                    }

                    $missing = [];
                    foreach ($cfg['fields'] as $f) {
                        if (!empty($f['required']) && trim((string)($fieldValues[$f['key']] ?? '')) === '') {
                            $missing[] = $f['label'];
                        }
                    }
                    if (!empty($missing)) {
                        $notes[] = 'Row ' . ($i + 1) . ' skipped — missing ' . implode(', ', $missing) . '.';
                        continue;
                    }

                    $badNumbers = [];
                    foreach ($outcomeValues as $ok => $ov) {
                        if ($ov !== null && !is_numeric($ov)) $badNumbers[] = strtoupper($ok);
                    }
                    if (!empty($badNumbers)) {
                        $notes[] = 'Row ' . ($i + 1) . ' skipped — non-numeric value for ' . implode(', ', $badNumbers) . '.';
                        continue;
                    }

                    insert_record($pdo, $key, $cfg, $fieldValues, $outcomeValues);
                    $inserted++;
                }
                $result = ['inserted' => $inserted, 'skipped' => count($notes), 'notes' => $notes, 'mode' => 'flat'];

            } elseif ($key === 'course_exit_survey' && $ext === 'xlsx') {
                // ---------- fallback: department "Course Evaluation Plan" workbook ----------
                $allSheets = xlsx_read_all_sheets($tmpPath);
                $summary = extract_course_final_summary($allSheets);
                if ($summary === null) {
                    throw new Exception('Could not recognize the column headers in this file, and it doesn\'t match the standard Course Evaluation Plan workbook layout either. Download the template below and match its headers, or check the CO ATTAINMENT block is present.');
                }
                $fieldValues = [
                    'acad_year'     => $summary['acad_year'],
                    'sem'           => $summary['sem'],
                    'c_name'        => $summary['c_name'],
                    'c_code'        => $summary['c_code'],
                    'c_coordinator' => $summary['c_coordinator'],
                ];
                $outcomeValues = [
                    'co1' => $summary['co1'], 'co2' => $summary['co2'], 'co3' => $summary['co3'],
                    'co4' => $summary['co4'], 'co5' => $summary['co5'], 'co6' => $summary['co6'],
                ];
                if (trim((string)$fieldValues['acad_year']) === '') {
                    throw new Exception('Found a CO ATTAINMENT block (sheet "' . e($summary['_source_sheet']) . '") but could not read the Academic Year / Batch / Term — please check the file or use the manual "Add record" form instead.');
                }
                insert_record($pdo, $key, $cfg, $fieldValues, $outcomeValues);
                $result = [
                    'inserted' => 1, 'skipped' => 0,
                    'notes' => ['Read from sheet "' . $summary['_source_sheet'] . '": ' . $summary['c_code'] . ' — ' . $summary['c_name']],
                    'mode' => 'summary',
                ];
            } else {
                throw new Exception('None of the columns in the first row matched this activity\'s fields. Download the template below, keep its header row, and fill data underneath it.');
            }

        } catch (Exception $e) {
            $error = $e->getMessage();
        }
    }
}

$active = 'history';
require __DIR__ . '/includes/header.php';
?>

<div class="wrap narrow">
  <div class="eyebrow">Bulk Import</div>
  <h1 class="page-title">Import <?= e($cfg['label']) ?> from Excel</h1>
  <p class="page-sub" style="margin-bottom:22px;">
    Upload a .xlsx or .csv file whose first row has column headers matching this activity's fields
    <?= $key === 'course_exit_survey' ? '&mdash; or upload the standard department Course Evaluation Plan workbook and the CO ATTAINMENT summary will be pulled out automatically.' : '.' ?>
  </p>

  <?php if ($error): ?>
    <div class="flash error"><?= e($error) ?></div>
  <?php endif; ?>

  <?php if ($result): ?>
    <div class="flash <?= $result['inserted'] > 0 ? 'success' : 'error' ?>">
      <?= (int)$result['inserted'] ?> record<?= $result['inserted'] === 1 ? '' : 's' ?> imported<?= $result['skipped'] ? ', ' . (int)$result['skipped'] . ' row' . ($result['skipped'] === 1 ? '' : 's') . ' skipped' : '' ?>.
    </div>
    <?php if (!empty($result['notes'])): ?>
      <div class="form-card" style="margin-bottom:22px;">
        <h3 style="font-size:14px;margin:0 0 10px;">Details</h3>
        <ul style="margin:0;padding-left:20px;font-size:13.5px;color:var(--ink-soft);line-height:1.7;">
          <?php foreach (array_slice($result['notes'], 0, 25) as $n): ?>
            <li><?= e($n) ?></li>
          <?php endforeach; ?>
        </ul>
      </div>
    <?php endif; ?>
    <div style="margin-bottom:26px;">
      <a href="history.php?table=<?= e($key) ?>" class="btn btn-primary">View history</a>
      <a href="import.php?table=<?= e($key) ?>" class="btn btn-outline">Import another file</a>
    </div>
  <?php endif; ?>

  <div class="form-card">
    <div class="form-section" style="margin-bottom:8px;">
      <h3>Upload file</h3>
      <form method="post" enctype="multipart/form-data">
        <div class="field">
          <label>Excel (.xlsx) or CSV file</label>
          <input type="file" name="file" accept=".xlsx,.csv" required>
        </div>
        <div class="form-actions" style="border-top:none;padding-top:0;">
          <a href="template.php?table=<?= e($key) ?>" class="btn btn-outline">Download template</a>
          <button type="submit" class="btn btn-primary">Upload &amp; import</button>
        </div>
      </form>
    </div>
  </div>

  <p class="hint" style="margin-top:16px;">
    Expected columns: <?= e(implode(', ', array_map(fn($f) => $f['label'], $cfg['fields']))) ?>,
    <?= e(implode(', ', array_map('strtoupper', $cfg['outcomes']))) ?>.
    Outcome columns may be left blank per row.
  </p>
</div>

<?php require __DIR__ . '/includes/footer.php'; ?>
