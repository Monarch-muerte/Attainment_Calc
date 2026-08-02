<?php
require __DIR__ . '/includes/auth.php';
require_login();
require __DIR__ . '/config/db.php';
require __DIR__ . '/includes/tables_config.php';
require __DIR__ . '/includes/functions.php';

[$key, $cfg] = resolve_table($TABLES, $_GET['table'] ?? null);
$id = (int)($_GET['id'] ?? 0);

$stmt = $pdo->prepare("SELECT * FROM `$key` WHERE id = ? LIMIT 1");
$stmt->execute([$id]);
$existing = $stmt->fetch();

if (!$existing) {
    http_response_code(404);
    die('<div style="font-family:sans-serif;padding:60px;text-align:center;color:#7a1f1f;">Record not found.</div>');
}

$isEdit = true;
$record = $existing;
$errors = [];

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    $fieldCols = array_column($cfg['fields'], 'key');
    $setParts  = [];
    $params    = ['id' => $id];

    foreach ($cfg['fields'] as $f) {
        $val = trim($_POST[$f['key']] ?? '');
        if (!empty($f['required']) && $val === '') {
            $errors[] = $f['label'] . ' is required.';
        }
        $record[$f['key']] = $val;
        $setParts[] = "`{$f['key']}` = :{$f['key']}";
        $params[$f['key']] = $val === '' ? null : $val;
    }

    foreach ($cfg['outcomes'] as $o) {
        $v = trim($_POST[$o] ?? '');
        if ($v !== '' && !is_numeric($v)) {
            $errors[] = strtoupper($o) . ' must be a number.';
        }
        $record[$o] = $v;
        $setParts[] = "`$o` = :$o";
        $params[$o] = $v === '' ? null : (float)$v;
    }

    if (empty($errors)) {
        $sql = "UPDATE `$key` SET " . implode(', ', $setParts) . " WHERE id = :id";
        $stmt = $pdo->prepare($sql);
        $stmt->execute($params);

        set_flash('success', $cfg['label'] . ' record updated.');
        header('Location: history.php?table=' . urlencode($key));
        exit;
    }
}

$active = 'history';
require __DIR__ . '/includes/header.php';
require __DIR__ . '/includes/form_view.php';
require __DIR__ . '/includes/footer.php';
