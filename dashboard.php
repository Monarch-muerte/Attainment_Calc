<?php
require __DIR__ . '/includes/auth.php';
require_login();
require __DIR__ . '/config/db.php';
require __DIR__ . '/includes/tables_config.php';
require __DIR__ . '/includes/functions.php';

$counts = [];
foreach ($TABLES as $key => $cfg) {
    $counts[$key] = table_count($pdo, $key);
}

$avgs   = combined_po_averages($pdo);
$coAvgs = course_exit_averages($pdo);

$active = 'dashboard';
require __DIR__ . '/includes/header.php';
?>

<div class="wrap">
  <div class="page-head">
    <div>
      <div class="eyebrow">Academic Quality Assurance</div>
      <h1 class="page-title">CO / PO / PSO attainment at a glance</h1>
      <p class="page-sub">Enter evidence by activity, retain complete history, and use the combined average for attainment reporting.</p>
    </div>
    <div>
      <a href="add.php?table=program_exit_survey" class="btn btn-primary">+ Add evidence</a>
    </div>
  </div>

  <div class="stat-grid">
    <?php foreach ($TABLES as $key => $cfg): ?>
      <div class="stat-card">
        <div class="lbl"><?= e($cfg['label']) ?></div>
        <div class="num"><?= $counts[$key] ?></div>
        <div class="link"><a href="history.php?table=<?= e($key) ?>">View history &rarr;</a></div>
      </div>
    <?php endforeach; ?>
  </div>

  <div class="avg-panel">
    <h2>Combined PO / PSO attainment average</h2>
    <p>Average across all saved Program Exit Survey, Expert Lecture, Industry Visit, Alumni Survey and Industry Survey records.</p>
    <div class="avg-grid">
      <?php foreach (PO_PSO as $k): $v = $avgs[$k]; ?>
        <div class="avg-cell">
          <div class="k"><?= strtoupper($k) ?></div>
          <div class="v <?= $v === null ? 'empty' : '' ?>"><?= fmt_avg($v) ?></div>
        </div>
      <?php endforeach; ?>
    </div>
  </div>

  <div class="co-panel">
    <h2>Course Exit Survey &mdash; CO summary</h2>
    <p>Average across all saved Course Exit Survey records (CO1&ndash;CO6). Reported per course; map to POs separately via your CO&ndash;PO matrix.</p>
    <div class="co-grid">
      <?php foreach (CO_LIST as $k): $v = $coAvgs[$k]; ?>
        <div class="co-cell">
          <div class="k"><?= strtoupper($k) ?></div>
          <div class="v"><?= fmt_avg($v) ?></div>
        </div>
      <?php endforeach; ?>
    </div>
  </div>

  <div class="card-grid">
    <?php foreach ($TABLES as $key => $cfg): ?>
      <div class="activity-card">
        <h3><?= e($cfg['label']) ?></h3>
        <p><?= e($cfg['desc']) ?></p>
        <div class="row">
          <a href="add.php?table=<?= e($key) ?>" class="btn btn-outline btn-sm">Add record</a>
          <a href="import.php?table=<?= e($key) ?>" class="btn-text">Import</a>
          <a href="history.php?table=<?= e($key) ?>" class="btn-text">History</a>
        </div>
      </div>
    <?php endforeach; ?>
  </div>
</div>

<?php require __DIR__ . '/includes/footer.php'; ?>
