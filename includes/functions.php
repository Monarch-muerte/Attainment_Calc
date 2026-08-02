<?php

function e($v): string {
    return htmlspecialchars((string)($v ?? ''), ENT_QUOTES, 'UTF-8');
}

/** Validate a ?table= key against the whitelist config. Dies with 404 if invalid. */
function resolve_table(array $TABLES, ?string $key): array {
    if ($key === null || !array_key_exists($key, $TABLES)) {
        http_response_code(404);
        die('<div style="font-family:sans-serif;padding:60px;text-align:center;color:#7a1f1f;">Unknown activity type.</div>');
    }
    return [$key, $TABLES[$key]];
}

/** Next serial number for a table (whitelisted table name only). */
function next_sr(PDO $pdo, string $table): int {
    $stmt = $pdo->query("SELECT MAX(sr) AS m FROM `$table`");
    $row = $stmt->fetch();
    return ((int)($row['m'] ?? 0)) + 1;
}

/** Format a nullable numeric outcome value for display. */
function fmt_outcome($v): string {
    if ($v === null || $v === '') return '—';
    return rtrim(rtrim(number_format((float)$v, 2, '.', ''), '0'), '.') ?: '0';
}

/** Format an average for display, or an em-dash if there is no data. */
function fmt_avg(?float $v): string {
    return $v === null ? '—' : number_format($v, 2);
}

/**
 * Combined PO/PSO averages across the PO_TABLES, ignoring NULLs.
 * Returns ['po1' => float|null, ...]
 */
function combined_po_averages(PDO $pdo): array {
    $out = [];
    foreach (PO_PSO as $col) {
        $unions = [];
        foreach (PO_TABLES as $t) {
            $unions[] = "SELECT `$col` AS v FROM `$t`";
        }
        $sql = 'SELECT AVG(v) AS avg_v FROM (' . implode(' UNION ALL ', $unions) . ') t';
        $row = $pdo->query($sql)->fetch();
        $out[$col] = $row['avg_v'] !== null ? (float)$row['avg_v'] : null;
    }
    return $out;
}

/** CO1-CO6 averages across course_exit_survey, ignoring NULLs. */
function course_exit_averages(PDO $pdo): array {
    $cols = implode(', ', array_map(fn($c) => "AVG(`$c`) AS `$c`", CO_LIST));
    $row = $pdo->query("SELECT $cols FROM course_exit_survey")->fetch();
    $out = [];
    foreach (CO_LIST as $c) {
        $out[$c] = ($row[$c] ?? null) !== null ? (float)$row[$c] : null;
    }
    return $out;
}

function table_count(PDO $pdo, string $table): int {
    return (int)$pdo->query("SELECT COUNT(*) AS c FROM `$table`")->fetch()['c'];
}

/** Insert one record into a whitelisted table, auto-assigning sr. Returns the new row id. */
function insert_record(PDO $pdo, string $table, array $cfg, array $fieldValues, array $outcomeValues): int {
    $sr = next_sr($pdo, $table);
    $fieldCols = array_column($cfg['fields'], 'key');
    $allCols   = array_merge(['sr'], $fieldCols, $cfg['outcomes']);
    $colList   = implode(', ', array_map(fn($c) => "`$c`", $allCols));
    $paramList = implode(', ', array_map(fn($c) => ":$c", $allCols));

    $params = ['sr' => $sr];
    foreach ($fieldCols as $fc) {
        $v = $fieldValues[$fc] ?? '';
        $params[$fc] = ($v === '' || $v === null) ? null : $v;
    }
    foreach ($cfg['outcomes'] as $o) {
        $v = $outcomeValues[$o] ?? null;
        $params[$o] = ($v === '' || $v === null) ? null : (float)$v;
    }

    $stmt = $pdo->prepare("INSERT INTO `$table` ($colList) VALUES ($paramList)");
    $stmt->execute($params);
    return (int)$pdo->lastInsertId();
}

/** Normalize a spreadsheet header cell for loose matching against field/outcome keys. */
function normalize_header(string $h): string {
    $h = strtolower(trim($h));
    $h = preg_replace('/[^a-z0-9]+/', '_', $h);
    return trim($h, '_');
}

/**
 * Given a table config and a header row, build [colIndex => fieldOrOutcomeKey].
 * Matches by key, label, or (for outcomes) the bare code like "po1"/"co1".
 */
function map_headers_to_keys(array $cfg, array $headerRow): array {
    $candidates = [];
    foreach ($cfg['fields'] as $f) {
        $candidates[normalize_header($f['key'])] = $f['key'];
        $candidates[normalize_header($f['label'])] = $f['key'];
    }
    foreach ($cfg['outcomes'] as $o) {
        $candidates[normalize_header($o)] = $o;
    }

    $map = [];
    foreach ($headerRow as $idx => $h) {
        $norm = normalize_header((string)$h);
        if ($norm !== '' && isset($candidates[$norm])) {
            $map[$idx] = $candidates[$norm];
        }
    }
    return $map;
}

function flash_message(): ?array {
    if (!empty($_SESSION['flash'])) {
        $f = $_SESSION['flash'];
        unset($_SESSION['flash']);
        return $f;
    }
    return null;
}

function set_flash(string $type, string $msg): void {
    $_SESSION['flash'] = ['type' => $type, 'msg' => $msg];
}
