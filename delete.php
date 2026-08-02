<?php
require __DIR__ . '/includes/auth.php';
require_login();
require __DIR__ . '/config/db.php';
require __DIR__ . '/includes/tables_config.php';
require __DIR__ . '/includes/functions.php';

if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
    header('Location: history.php');
    exit;
}

[$key, $cfg] = resolve_table($TABLES, $_POST['table'] ?? null);
$id = (int)($_POST['id'] ?? 0);

$stmt = $pdo->prepare("DELETE FROM `$key` WHERE id = ?");
$stmt->execute([$id]);

set_flash('success', 'Record deleted.');
header('Location: history.php?table=' . urlencode($key));
exit;
