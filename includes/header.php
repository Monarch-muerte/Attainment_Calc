<?php
// Expects $active (string: 'dashboard' | 'history') to be set by the including page.
$active = $active ?? '';
?>
<!doctype html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>College Attainment Portal</title>
<link rel="stylesheet" href="assets/css/style.css">
</head>
<body>
<div id="app">
  <div class="topbar">
    <a href="dashboard.php" class="brand">
      <div class="brand-badge">CA</div>
      <div class="brand-name">College Attainment</div>
    </a>
    <div class="nav-center">
      <div class="nav-links">
        <a href="dashboard.php" class="<?= $active === 'dashboard' ? 'active' : '' ?>">Dashboard</a>
        <a href="history.php" class="<?= $active === 'history' ? 'active' : '' ?>">History</a>
      </div>
    </div>
    <div class="nav-right">
      <span>Hello, <?= e(current_user_name()) ?></span>
      <a href="logout.php">Sign out</a>
    </div>
  </div>
  <main>
