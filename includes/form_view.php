<?php
/**
 * Expects the including page to define:
 *   $key      - whitelisted table key
 *   $cfg      - table config array
 *   $isEdit   - bool
 *   $record   - assoc array of current values (empty array for a new record)
 *   $errors   - array of error strings
 */
?>
<div class="wrap narrow">
  <div class="eyebrow">Evidence Entry</div>
  <h1 class="page-title"><?= $isEdit ? 'Edit' : 'Add' ?> <?= e($cfg['label']) ?></h1>
  <p class="page-sub" style="margin-bottom:22px;">
    <?= $cfg['outcome_label'] === 'CO' ? 'CO inputs are stored as DOUBLE values; leave blank if not assessed.' : 'PO/PSO inputs are stored as DOUBLE values; leave blank if not assessed.' ?>
  </p>

  <?php if (!empty($errors)): ?>
    <div class="flash error">
      <?= e(implode(' ', $errors)) ?>
    </div>
  <?php endif; ?>

  <form method="post" class="form-card">
    <div class="form-section">
      <h3>Activity details</h3>
      <div class="grid-fields">
        <?php foreach ($cfg['fields'] as $f): ?>
          <div class="field">
            <label><?= e($f['label']) ?><?= !empty($f['required']) ? ' *' : '' ?></label>
            <input
              type="<?= e($f['type']) ?>"
              name="<?= e($f['key']) ?>"
              value="<?= e($record[$f['key']] ?? '') ?>"
              placeholder="<?= e($f['ph'] ?? '') ?>">
          </div>
        <?php endforeach; ?>
      </div>
    </div>

    <div class="form-section" style="margin-bottom:6px;">
      <h3><?= $cfg['outcome_label'] === 'CO' ? 'Course outcome (CO) attainment values' : 'Outcome attainment values' ?></h3>
      <div class="grid-outcomes">
        <?php foreach ($cfg['outcomes'] as $o): ?>
          <div class="field">
            <label><?= strtoupper($o) ?></label>
            <input type="number" step="any" name="<?= e($o) ?>" value="<?= e($record[$o] ?? '') ?>" placeholder="&mdash;">
          </div>
        <?php endforeach; ?>
      </div>
    </div>

    <div class="form-actions">
      <a href="history.php?table=<?= e($key) ?>" class="btn btn-outline">Cancel</a>
      <button type="submit" class="btn btn-primary">Save record</button>
    </div>
  </form>
</div>
