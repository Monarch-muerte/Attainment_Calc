<?php
require __DIR__ . '/includes/auth.php';
require_login();
require __DIR__ . '/config/db.php';
require __DIR__ . '/includes/tables_config.php';
require __DIR__ . '/includes/functions.php';

[$key, $cfg] = resolve_table($TABLES, $_GET['table'] ?? null);

$isEdit = false;
$record = [];
$errors = [];

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    foreach ($cfg['fields'] as $f) {
        $val = trim($_POST[$f['key']] ?? '');
        if (!empty($f['required']) && $val === '') {
            $errors[] = $f['label'] . ' is required.';
        }
        $record[$f['key']] = $val;
    }

    $outcomeVals = [];
    foreach ($cfg['outcomes'] as $o) {
        $v = trim($_POST[$o] ?? '');
        if ($v !== '' && !is_numeric($v)) {
            $errors[] = strtoupper($o) . ' must be a number.';
        }
        $outcomeVals[$o] = $v === '' ? null : (float)$v;
        $record[$o] = $v;
    }

    if (empty($errors)) {
        insert_record($pdo, $key, $cfg, $record, $outcomeVals);
        set_flash('success', $cfg['label'] . ' record saved.');
        header('Location: history.php?table=' . urlencode($key));
        exit;
    }
}

$active = 'history';
require __DIR__ . '/includes/header.php';
require __DIR__ . '/includes/form_view.php';
require __DIR__ . '/includes/footer.php';
