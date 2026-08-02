<?php
require __DIR__ . '/includes/auth.php';
require_login();
require __DIR__ . '/config/db.php';
require __DIR__ . '/includes/tables_config.php';
require __DIR__ . '/includes/functions.php';

[$key, $cfg] = resolve_table($TABLES, $_GET['table'] ?? null);

$headers = array_merge(
    array_map(fn($f) => $f['label'], $cfg['fields']),
    array_map('strtoupper', $cfg['outcomes'])
);

header('Content-Type: text/csv; charset=utf-8');
header('Content-Disposition: attachment; filename="' . $key . '_template.csv"');

$out = fopen('php://output', 'w');
fputcsv($out, $headers);
// one blank example row so column order is obvious in Excel
fputcsv($out, array_fill(0, count($headers), ''));
fclose($out);
exit;
