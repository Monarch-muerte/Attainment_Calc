<?php
require __DIR__ . '/includes/auth.php';
require_login();
require __DIR__ . '/config/db.php';
require __DIR__ . '/includes/tables_config.php';
require __DIR__ . '/includes/functions.php';

$active = 'history';
$tableParam = $_GET['table'] ?? null;
$flash = flash_message();

require __DIR__ . '/includes/header.php';

if ($flash): ?>
  <div class="wrap" style="padding-bottom:0;">
    <div class="flash <?= e($flash['type']) ?>"><?= e($flash['msg']) ?></div>
  </div>
<?php endif;

if ($tableParam === null):
    // ---------- picker view ----------
    ?>
    <div class="wrap">
      <div class="page-head">
        <div>
          <div class="eyebrow">Records</div>
          <h1 class="page-title">Browse history</h1>
          <p class="page-sub">Choose an activity type to view, edit, or remove saved evidence.</p>
        </div>
      </div>
      <div class="chooser-grid">
        <?php foreach ($TABLES as $key => $cfg):
          $count = table_count($pdo, $key);
        ?>
          <a class="chooser-card" href="history.php?table=<?= e($key) ?>">
            <div class="n"><?= $count ?></div>
            <div class="l"><?= e($cfg['label']) ?></div>
          </a>
        <?php endforeach; ?>
      </div>
    </div>
    <?php
else:
    // ---------- single table view ----------
    [$key, $cfg] = resolve_table($TABLES, $tableParam);

    $stmt = $pdo->query("SELECT * FROM `$key` ORDER BY sr DESC");
    $rows = $stmt->fetchAll();
    ?>
    <div class="wrap">
      <div class="page-head">
        <div>
          <div class="eyebrow">Records</div>
          <h1 class="page-title"><?= e($cfg['label']) ?> &mdash; history</h1>
          <p class="page-sub"><?= count($rows) ?> saved record<?= count($rows) === 1 ? '' : 's' ?>.</p>
        </div>
        <div>
          <a href="import.php?table=<?= e($key) ?>" class="btn btn-outline">Import from Excel</a>
          <a href="add.php?table=<?= e($key) ?>" class="btn btn-primary">+ Add record</a>
        </div>
      </div>

      <div class="table-wrap scroll-x">
        <table class="hist">
          <thead>
            <tr>
              <th>Sr</th>
              <?php foreach ($cfg['fields'] as $f): ?><th><?= e($f['label']) ?></th><?php endforeach; ?>
              <?php foreach ($cfg['outcomes'] as $o): ?><th><?= strtoupper($o) ?></th><?php endforeach; ?>
              <th></th>
            </tr>
          </thead>
          <tbody>
            <?php if (empty($rows)): ?>
              <tr>
                <td colspan="<?= 2 + count($cfg['fields']) + count($cfg['outcomes']) ?>">
                  <div class="empty-state">
                    <div class="big">No records yet</div>
                    Add your first <?= e(strtolower($cfg['label'])) ?> record to start tracking attainment.
                  </div>
                </td>
              </tr>
            <?php else: ?>
              <?php foreach ($rows as $r): ?>
                <tr>
                  <td class="num"><?= (int)$r['sr'] ?></td>
                  <?php foreach ($cfg['fields'] as $f): ?>
                    <td><?= e($r[$f['key']] !== null && $r[$f['key']] !== '' ? $r[$f['key']] : '—') ?></td>
                  <?php endforeach; ?>
                  <?php foreach ($cfg['outcomes'] as $o): ?>
                    <td class="num"><?= fmt_outcome($r[$o]) ?></td>
                  <?php endforeach; ?>
                  <td>
                    <a class="btn-text" href="edit.php?table=<?= e($key) ?>&id=<?= (int)$r['id'] ?>">Edit</a>
                    <form method="post" action="delete.php" class="inline-form" onsubmit="return confirm('Delete this record? This cannot be undone.');">
                      <input type="hidden" name="table" value="<?= e($key) ?>">
                      <input type="hidden" name="id" value="<?= (int)$r['id'] ?>">
                      <button type="submit" class="btn-danger-text">Delete</button>
                    </form>
                  </td>
                </tr>
              <?php endforeach; ?>
            <?php endif; ?>
          </tbody>
        </table>
      </div>
    </div>
    <?php
endif;

require __DIR__ . '/includes/footer.php';
