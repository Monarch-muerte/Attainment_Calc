<?php
require __DIR__ . '/includes/auth.php';
require __DIR__ . '/config/db.php';
require __DIR__ . '/includes/functions.php';

if (is_logged_in()) {
    header('Location: dashboard.php');
    exit;
}

$error = '';

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    $username = trim($_POST['username'] ?? '');
    $password = $_POST['password'] ?? '';

    $stmt = $pdo->prepare('SELECT id, username, password_hash, full_name FROM users WHERE username = ? LIMIT 1');
    $stmt->execute([$username]);
    $user = $stmt->fetch();

    if ($user && password_verify($password, $user['password_hash'])) {
        $_SESSION['user_id']   = $user['id'];
        $_SESSION['user_name'] = $user['full_name'];
        header('Location: dashboard.php');
        exit;
    }
    $error = 'Incorrect username or password.';
}
?>
<!doctype html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Sign in &middot; College Attainment Portal</title>
<link rel="stylesheet" href="assets/css/style.css">
</head>
<body>
<div class="login-shell">
  <div class="login-card">
    <div class="brand-badge">CA</div>
    <div class="eyebrow">K.D. POLYTECHNIC &middot; COMPUTER ENGINEERING</div>
    <h1>Attainment Portal</h1>
    <p class="sub">Record and report CO / PO / PSO attainment evidence.</p>

    <?php if ($error): ?>
      <div class="login-err"><?= e($error) ?></div>
    <?php endif; ?>

    <form method="post">
      <div class="field">
        <label>Username</label>
        <input type="text" name="username" autocomplete="username" autofocus value="<?= e($_POST['username'] ?? '') ?>">
      </div>
      <div class="field">
        <label>Password</label>
        <input type="password" name="password" autocomplete="current-password">
      </div>
      <button type="submit" class="btn btn-primary" style="width:100%;padding:12px;">Sign in</button>
    </form>

    <div class="login-note">Default administrator: <b>Admin9</b> / <b>654321</b></div>
  </div>
</div>
</body>
</html>
