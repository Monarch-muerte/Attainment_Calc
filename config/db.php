<?php
/**
 * Database connection.
 * Default values match a stock XAMPP install (MySQL on localhost, user root, no password).
 * Change these if your setup differs.
 */
$DB_HOST = 'localhost';
$DB_NAME = 'college_attainment';
$DB_USER = 'root';
$DB_PASS = '';

try {
    $pdo = new PDO(
        "mysql:host={$DB_HOST};dbname={$DB_NAME};charset=utf8mb4",
        $DB_USER,
        $DB_PASS,
        [
            PDO::ATTR_ERRMODE            => PDO::ERRMODE_EXCEPTION,
            PDO::ATTR_DEFAULT_FETCH_MODE => PDO::FETCH_ASSOC,
            PDO::ATTR_EMULATE_PREPARES   => false,
        ]
    );
} catch (PDOException $e) {
    http_response_code(500);
    die(
        '<div style="font-family:sans-serif;max-width:640px;margin:60px auto;padding:24px;'
        . 'border:1px solid #f1c0c0;background:#fff5f5;border-radius:10px;color:#7a1f1f;">'
        . '<h2 style="margin-top:0;">Database connection failed</h2>'
        . '<p>Could not connect to MySQL database <b>' . htmlspecialchars($DB_NAME) . '</b>.</p>'
        . '<p>Check that MySQL is running in XAMPP and that you have imported <code>sql/schema.sql</code>, '
        . 'then verify the credentials in <code>config/db.php</code>.</p>'
        . '<p style="color:#a55;font-size:13px;">Detail: ' . htmlspecialchars($e->getMessage()) . '</p>'
        . '</div>'
    );
}
