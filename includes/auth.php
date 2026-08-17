<?php
if (session_status() === PHP_SESSION_NONE) {
    session_start();
}

function is_logged_in(): bool {
    return !empty($_SESSION['user_id']);
}

function require_login(): void {
    if (!is_logged_in()) {
        header('Location: index.php');
        exit;
    }
}

function current_user_name(): string {
    return $_SESSION['user_name'] ?? 'Administrator';
}
