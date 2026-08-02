<?php
/**
 * Minimal XLSX reader — no Composer / PhpSpreadsheet needed, just the
 * zip and xml extensions that ship enabled with a stock XAMPP install.
 * Good enough to read simple cell values (text + numbers) and shared strings.
 */

function xlsx_col_to_index(string $letters): int {
    $letters = strtoupper($letters);
    $col = 0;
    for ($i = 0; $i < strlen($letters); $i++) {
        $col = $col * 26 + (ord($letters[$i]) - ord('A') + 1);
    }
    return $col - 1; // zero-based
}

/** Open the workbook and return [$zip, $sharedStrings, $sheets] or throw Exception. */
function xlsx_open(string $path): array {
    if (!class_exists('ZipArchive')) {
        throw new Exception('The PHP zip extension is not enabled, so .xlsx files cannot be read. Please enable it in php.ini, or upload a .csv file instead.');
    }
    $zip = new ZipArchive();
    if ($zip->open($path) !== true) {
        throw new Exception('Could not open the uploaded file as an Excel workbook.');
    }

    // shared strings
    $sharedStrings = [];
    $ssXml = $zip->getFromName('xl/sharedStrings.xml');
    if ($ssXml !== false) {
        $ssDom = @simplexml_load_string($ssXml);
        if ($ssDom !== false) {
            foreach ($ssDom->si as $si) {
                if (isset($si->t)) {
                    $sharedStrings[] = (string)$si->t;
                } else {
                    $text = '';
                    foreach ($si->r as $run) {
                        $text .= (string)$run->t;
                    }
                    $sharedStrings[] = $text;
                }
            }
        }
    }

    // sheet name -> xml path, via workbook.xml + its rels
    $sheets = [];
    $wbXml = $zip->getFromName('xl/workbook.xml');
    $relsXml = $zip->getFromName('xl/_rels/workbook.xml.rels');
    if ($wbXml !== false && $relsXml !== false) {
        $wbDom = @simplexml_load_string($wbXml);
        $relsDom = @simplexml_load_string($relsXml);
        $relMap = [];
        if ($relsDom !== false) {
            foreach ($relsDom->Relationship as $rel) {
                $relMap[(string)$rel['Id']] = (string)$rel['Target'];
            }
        }
        if ($wbDom !== false) {
            $ns = $wbDom->getNamespaces(true);
            $rNs = $ns['r'] ?? 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
            foreach ($wbDom->sheets->sheet as $sheet) {
                $attrs = $sheet->attributes($rNs);
                $rid = (string)$attrs['id'];
                $target = $relMap[$rid] ?? null;
                if ($target) {
                    $target = ltrim($target, '/');
                    if (strpos($target, 'worksheets/') === 0) $target = 'xl/' . $target;
                    $sheets[(string)$sheet['name']] = $target;
                }
            }
        }
    }
    if (empty($sheets)) {
        // fallback: just grab any worksheet files in order
        for ($i = 0; $i < $zip->numFiles; $i++) {
            $name = $zip->getNameIndex($i);
            if (preg_match('#^xl/worksheets/sheet(\d+)\.xml$#', $name, $m)) {
                $sheets['Sheet' . $m[1]] = $name;
            }
        }
    }

    return [$zip, $sharedStrings, $sheets];
}

/** Read one worksheet into a zero-indexed 2D array of string values. */
function xlsx_read_sheet($zip, array $sharedStrings, string $path): array {
    $xml = $zip->getFromName($path);
    if ($xml === false) return [];
    $dom = @simplexml_load_string($xml);
    if ($dom === false || !isset($dom->sheetData)) return [];

    $rows = [];
    foreach ($dom->sheetData->row as $row) {
        $rowIdx = (int)$row['r'] - 1;
        $rowData = [];
        foreach ($row->c as $c) {
            $ref = (string)$c['r'];
            if (!preg_match('/^([A-Z]+)(\d+)$/', $ref, $m)) continue;
            $colIdx = xlsx_col_to_index($m[1]);
            $type = (string)$c['t'];
            $value = '';
            if ($type === 'inlineStr') {
                $value = isset($c->is->t) ? (string)$c->is->t : '';
            } elseif (isset($c->v)) {
                $raw = (string)$c->v;
                $value = ($type === 's') ? ($sharedStrings[(int)$raw] ?? '') : $raw;
            }
            $rowData[$colIdx] = trim($value);
        }
        if (!empty($rowData)) {
            ksort($rowData);
            $maxCol = max(array_keys($rowData));
            $normalized = [];
            for ($i = 0; $i <= $maxCol; $i++) {
                $normalized[$i] = $rowData[$i] ?? '';
            }
            $rows[$rowIdx] = $normalized;
        }
    }
    ksort($rows);
    return array_values($rows);
}

/** Read the first worksheet of an .xlsx file as a plain 2D array. */
function xlsx_read_first_sheet_as_table(string $path): array {
    [$zip, $sharedStrings, $sheets] = xlsx_open($path);
    try {
        if (empty($sheets)) return [];
        $firstPath = array_values($sheets)[0];
        return xlsx_read_sheet($zip, $sharedStrings, $firstPath);
    } finally {
        $zip->close();
    }
}

/** Read every worksheet of an .xlsx file. Returns ['SheetName' => rows, ...]. */
function xlsx_read_all_sheets(string $path): array {
    [$zip, $sharedStrings, $sheets] = xlsx_open($path);
    $out = [];
    try {
        foreach ($sheets as $name => $sheetPath) {
            $out[$name] = xlsx_read_sheet($zip, $sharedStrings, $sheetPath);
        }
    } finally {
        $zip->close();
    }
    return $out;
}

/** Parse a .csv file into a plain 2D array. */
function csv_read_as_table(string $path): array {
    $rows = [];
    if (($h = fopen($path, 'r')) !== false) {
        while (($data = fgetcsv($h)) !== false) {
            $rows[] = array_map('trim', $data);
        }
        fclose($h);
    }
    return $rows;
}

/** In a raw row/col grid, find a labeled cell (e.g. "Course Code:") and return the next non-blank cell after it on the same row. */
function grid_find_label_value(array $rows, string $label): ?string {
    $target = strtolower(trim(rtrim(trim($label), ':')));
    foreach ($rows as $row) {
        foreach ($row as $idx => $val) {
            $norm = strtolower(trim(rtrim(trim((string)$val), ':')));
            if ($norm === $target) {
                for ($j = $idx + 1; $j < count($row); $j++) {
                    if (trim((string)$row[$j]) !== '') {
                        return trim((string)$row[$j]);
                    }
                }
            }
        }
    }
    return null;
}

/**
 * Auto-detect the department's standard "Course Evaluation Plan" workbook
 * (the FINAL sheet's CO ATTAINMENT block) and extract one Course Exit
 * Survey record from it: acad_year, sem, c_name, c_code, co1..co6.
 * Returns null if the workbook doesn't match this layout.
 */
function extract_course_final_summary(array $allSheets): ?array {
    foreach ($allSheets as $sheetName => $rows) {
        // Locate the header row containing both a "CO" label column and a "CO ATTAIN" column.
        $headerRowIdx = null;
        $coLabelCol = null;
        $coAttainCol = null;
        foreach ($rows as $rIdx => $row) {
            foreach ($row as $cIdx => $val) {
                $norm = strtolower(trim((string)$val));
                if ($norm === 'co attain') $coAttainCol = $cIdx;
                if ($norm === 'co') $coLabelCol = $cIdx;
            }
            if ($coAttainCol !== null && $coLabelCol !== null) {
                $headerRowIdx = $rIdx;
                break;
            }
            $coLabelCol = null; // reset per row - both must be on the same row
            $coAttainCol = null;
        }

        if ($headerRowIdx === null) continue;

        // Walk rows below the header, pulling CO1..CO6 attainment values.
        // Stop as soon as the block ends (CO6 reached, or a non-CO row after we've started)
        // so we don't wander into an unrelated table further down that reuses "CO1"/"CO2" labels.
        $coValues = [];
        for ($r = $headerRowIdx + 1; $r < count($rows) && $r <= $headerRowIdx + 20; $r++) {
            $row = $rows[$r] ?? [];
            $label = isset($row[$coLabelCol]) ? trim((string)$row[$coLabelCol]) : '';
            if (preg_match('/^CO([1-6])$/i', $label, $m)) {
                $raw = $row[$coAttainCol] ?? '';
                $coValues[(int)$m[1]] = ($raw === '' ? null : (float)$raw);
                if ((int)$m[1] === 6) break; // CO6 is always last
            } elseif (!empty($coValues)) {
                break; // block ended
            }
        }

        if (empty($coValues)) continue;

        $courseCode  = grid_find_label_value($rows, 'Course Code');
        $courseName  = grid_find_label_value($rows, 'Course Name');
        $semester    = grid_find_label_value($rows, 'Semester');
        $batch       = grid_find_label_value($rows, 'Batch');
        $term        = grid_find_label_value($rows, 'Term');

        return [
            'acad_year'     => $batch ?: $term ?: '',
            'sem'           => $semester ?: '',
            'c_name'        => $courseName ?: '',
            'c_code'        => $courseCode ?: '',
            'c_coordinator' => '',
            'co1' => $coValues[1] ?? null,
            'co2' => $coValues[2] ?? null,
            'co3' => $coValues[3] ?? null,
            'co4' => $coValues[4] ?? null,
            'co5' => $coValues[5] ?? null,
            'co6' => $coValues[6] ?? null,
            '_source_sheet' => $sheetName,
        ];
    }
    return null;
}
