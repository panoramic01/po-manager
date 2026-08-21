/**
 * Month-end close: a full physical count reconciled against the Stock
 * Ledger, and a per-period report of the result.
 *
 * Compares physical count against ledger only, for now -- nothing posts to
 * QuickBooks yet, so there is no QBO leg to reconcile against. The Period
 * Close row leaves room for one (qboValue stays blank) so wiring that in
 * later is a column, not a redesign.
 *
 * Every count adjustment stays a deliberate, human decision (the locked
 * design): this module never posts a COUNT_ADJUST itself. It flags what's
 * worth reviewing; the admin opens the existing Count form
 * (Stock_Ledger_Api.gs / adjustStockCount) per material to actually fix it.
 */

// --- Sheets ----------------------------------------------------------------

var PHYSICAL_COUNT_SHEET = 'Physical Count';
var PHYSICAL_COUNT_HEADERS = [
  'Period', 'Material Id', 'Material Name', 'Counted Qty', 'Ledger Qty At Count',
  'Variance Qty', 'Variance Value', 'Counted By', 'Counted At'
];
var PHYSICAL_COUNT_COL = {};
PHYSICAL_COUNT_HEADERS.forEach(function(h, i) { PHYSICAL_COUNT_COL[h] = i; });

var PERIOD_CLOSE_SHEET = 'Period Close';
var PERIOD_CLOSE_HEADERS = [
  'Period', 'Status', 'Ledger Value', 'Materials Counted', 'Materials Flagged',
  'Variance Total Value', 'QBO Value', 'Closed By', 'Closed At', 'Note'
];
var PERIOD_CLOSE_COL = {};
PERIOD_CLOSE_HEADERS.forEach(function(h, i) { PERIOD_CLOSE_COL[h] = i; });

function physicalCountSheet_() {
  return ensureSheetWithHeaders_(stockSheetName_(PHYSICAL_COUNT_SHEET), PHYSICAL_COUNT_HEADERS);
}
function periodCloseSheet_() {
  return ensureSheetWithHeaders_(stockSheetName_(PERIOD_CLOSE_SHEET), PERIOD_CLOSE_HEADERS);
}

// --- Tolerance ---------------------------------------------------------------

/**
 * $25 or 2% of the material's own ledger value, whichever is larger --
 * the locked tolerance. Script Properties so it is tunable without a
 * redeploy once real months of data show whether it is too loose or
 * too tight. STOCK_CLOSE_TOLERANCE_FLOOR / STOCK_CLOSE_TOLERANCE_PCT.
 */
function getStockCloseTolerance_() {
  var props = PropertiesService.getScriptProperties();
  var floor = parseFloat(props.getProperty('STOCK_CLOSE_TOLERANCE_FLOOR'));
  var pct = parseFloat(props.getProperty('STOCK_CLOSE_TOLERANCE_PCT'));
  return {
    floor: isNaN(floor) ? 25 : floor,
    pct: isNaN(pct) ? 0.02 : pct
  };
}

/**
 * Whether a quantity variance is worth a human's attention. A material with
 * no cost yet (never received, or a data problem) can't be priced into a
 * dollar tolerance at all -- rather than silently pass it, any nonzero
 * variance on an unpriced material is flagged, since there is no dollar
 * basis to say it's immaterial.
 */
function isStockVarianceFlagged_(varianceQty, avgCost, materialLedgerValue) {
  if (!varianceQty) return false;
  if (!(avgCost > 0)) return true;
  var tol = getStockCloseTolerance_();
  var varianceValue = Math.abs(varianceQty) * avgCost;
  var threshold = Math.max(tol.floor, Math.abs(materialLedgerValue) * tol.pct);
  return varianceValue > threshold;
}

// --- Period targeting ----------------------------------------------------

/**
 * The period the close screen operates on: the oldest period with ledger
 * activity that hasn't been closed yet. Ordered so months close in
 * sequence, the way a real books close does. Included even when that
 * happens to be the current, still-in-progress month -- useful for a
 * system this new, where there may be no completed prior month yet -- but
 * the caller gets isCurrentPeriod so the screen can say so.
 */
function determineCloseablePeriod_() {
  var currentPeriod = periodOf_(new Date());
  var moves = readStockLedgerMoves_(0);
  var seen = {};
  moves.forEach(function(m) { if (m.period) seen[m.period] = true; });
  var periods = Object.keys(seen).sort();

  var closedSet = {};
  var sheet = periodCloseSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    var data = sheet.getRange(2, 1, lastRow - 1, PERIOD_CLOSE_HEADERS.length).getValues();
    data.forEach(function(r) {
      var per = (r[PERIOD_CLOSE_COL['Period']] || '').toString().trim();
      var status = (r[PERIOD_CLOSE_COL['Status']] || '').toString().trim();
      if (per && status === 'closed') closedSet[per] = true;
    });
  }

  var target = null;
  for (var i = 0; i < periods.length; i++) {
    if (closedSet[periods[i]]) continue;
    target = periods[i];
    break;
  }
  return { period: target, currentPeriod: currentPeriod, isCurrentPeriod: target === currentPeriod };
}

// --- Physical count --------------------------------------------------------

/** Every count row for a period, keyed by Material Id. */
function readPhysicalCounts_(period) {
  var sheet = physicalCountSheet_();
  var lastRow = sheet.getLastRow();
  var out = {};
  if (lastRow < 2) return out;
  var data = sheet.getRange(2, 1, lastRow - 1, PHYSICAL_COUNT_HEADERS.length).getValues();
  data.forEach(function(r, i) {
    if ((r[PHYSICAL_COUNT_COL['Period']] || '').toString().trim() !== period) return;
    var materialId = (r[PHYSICAL_COUNT_COL['Material Id']] || '').toString().trim();
    if (!materialId) return;
    out[materialId] = {
      rowIndex: i + 2,
      materialId: materialId,
      materialName: (r[PHYSICAL_COUNT_COL['Material Name']] || '').toString(),
      countedQty: parseFloat(r[PHYSICAL_COUNT_COL['Counted Qty']]) || 0,
      ledgerQtyAtCount: parseFloat(r[PHYSICAL_COUNT_COL['Ledger Qty At Count']]) || 0,
      varianceQty: parseFloat(r[PHYSICAL_COUNT_COL['Variance Qty']]) || 0,
      varianceValue: parseFloat(r[PHYSICAL_COUNT_COL['Variance Value']]) || 0,
      countedBy: (r[PHYSICAL_COUNT_COL['Counted By']] || '').toString(),
      countedAt: r[PHYSICAL_COUNT_COL['Counted At']]
    };
  });
  return out;
}

function periodIsClosed_(period) {
  var sheet = periodCloseSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  var data = sheet.getRange(2, 1, lastRow - 1, PERIOD_CLOSE_HEADERS.length).getValues();
  for (var i = 0; i < data.length; i++) {
    if ((data[i][PERIOD_CLOSE_COL['Period']] || '').toString().trim() === period &&
        (data[i][PERIOD_CLOSE_COL['Status']] || '').toString().trim() === 'closed') {
      return true;
    }
  }
  return false;
}

/**
 * Everything the count-entry screen needs: every active material with its
 * current ledger on-hand, and whatever's already been counted this period
 * (so a partial count can be resumed/corrected before close, not just
 * started over).
 */
function computeStockCountSheet_() {
  var target = determineCloseablePeriod_();
  if (!target.period) return { materials: [], period: null, message: 'No ledger activity yet -- nothing to count.' };
  if (periodIsClosed_(target.period)) return { materials: [], period: target.period, alreadyClosed: true };

  var catalog = getMaterialCatalog_();
  var position = getStockPosition_();
  var counted = readPhysicalCounts_(target.period);

  var materials = Object.keys(catalog)
    .filter(function(id) { return catalog[id].active; })
    .map(function(id) {
      var pos = position.positions[id];
      var onHand = pos ? pos.qty : 0;
      var c = counted[id];
      return {
        materialId: id,
        materialName: catalog[id].materialName,
        unit: catalog[id].unit,
        ledgerQty: onHand,
        countedQty: c ? c.countedQty : null,
        alreadyCounted: !!c
      };
    })
    .sort(function(a, b) { return a.materialName.localeCompare(b.materialName); });

  return {
    materials: materials,
    period: target.period,
    isCurrentPeriod: target.isCurrentPeriod,
    countedSoFar: Object.keys(counted).length,
    totalActive: materials.length
  };
}

function getStockCountSheet(payload) {
  try {
    var auth = authorizeCaller(payload, STOCK_READ_ROLES);
    if (!auth.ok) return { error: auth.error, code: auth.code, materials: [] };
    return withStockSheets_(payload, computeStockCountSheet_);
  } catch (e) {
    return { error: e.toString(), materials: [] };
  }
}

/**
 * Writes/updates this period's count rows. Upsert on (Period, Material Id)
 * so recounting a material before close corrects it rather than
 * duplicating a row. Refuses outright once the period is closed --
 * physical count history is part of the closed record and does not
 * change after the fact.
 * payload: {lines: [{materialId, countedQty}]}.
 */
function writePhysicalCount_(lines, byEmail) {
  var target = determineCloseablePeriod_();
  if (!target.period) return { success: false, error: 'No ledger activity yet -- nothing to count.' };
  if (periodIsClosed_(target.period)) return { success: false, error: 'This period is already closed -- counts are part of the closed record.' };

  var catalog = getMaterialCatalog_();
  var position = getStockPosition_();
  var existing = readPhysicalCounts_(target.period);
  var now = new Date();
  var sheet = physicalCountSheet_();

  for (var i = 0; i < lines.length; i++) {
    var materialId = (lines[i].materialId || '').toString().trim();
    var countedQty = parseFloat(lines[i].countedQty);
    if (!materialId || isNaN(countedQty) || countedQty < 0) {
      return { success: false, error: 'Every counted line needs a material and a quantity of 0 or more.' };
    }
    var cat = catalog[materialId];
    if (!cat) return { success: false, error: 'Material "' + materialId + '" is not in the catalog.' };

    var pos = position.positions[materialId];
    var ledgerQty = pos ? pos.qty : 0;
    var avgCost = pos && pos.qty ? pos.value / pos.qty : 0;
    var varianceQty = countedQty - ledgerQty;
    var varianceValue = Math.round(varianceQty * avgCost * 100) / 100;

    var row = [];
    for (var c = 0; c < PHYSICAL_COUNT_HEADERS.length; c++) row.push('');
    row[PHYSICAL_COUNT_COL['Period']]              = target.period;
    row[PHYSICAL_COUNT_COL['Material Id']]         = materialId;
    row[PHYSICAL_COUNT_COL['Material Name']]       = cat.materialName;
    row[PHYSICAL_COUNT_COL['Counted Qty']]         = countedQty;
    row[PHYSICAL_COUNT_COL['Ledger Qty At Count']] = ledgerQty;
    row[PHYSICAL_COUNT_COL['Variance Qty']]        = varianceQty;
    row[PHYSICAL_COUNT_COL['Variance Value']]      = varianceValue;
    row[PHYSICAL_COUNT_COL['Counted By']]          = byEmail || '';
    row[PHYSICAL_COUNT_COL['Counted At']]          = now;

    var prior = existing[materialId];
    if (prior) {
      sheet.getRange(prior.rowIndex, 1, 1, PHYSICAL_COUNT_HEADERS.length).setValues([row]);
    } else {
      sheet.appendRow(row);
      existing[materialId] = { rowIndex: sheet.getLastRow() };
    }
  }
  return { success: true, period: target.period, counted: lines.length };
}

/**
 * Writes/updates this period's count rows. Upsert on (Period, Material Id)
 * so recounting a material before close corrects it rather than
 * duplicating a row. Refuses outright once the period is closed --
 * physical count history is part of the closed record and does not
 * change after the fact.
 * payload: {lines: [{materialId, countedQty}]}.
 */
function submitPhysicalCount(payload) {
  try {
    var auth = authorizeCaller(payload, STOCK_WRITE_ROLES);
    if (!auth.ok) return { success: false, error: auth.error, code: auth.code };

    var lines = Array.isArray(payload.lines) ? payload.lines : [];
    if (!lines.length) return { success: false, error: 'Count at least one material.' };

    var lock = LockService.getScriptLock();
    if (!lock.tryLock(15000)) return { success: false, error: 'Server is busy - try again in a moment.' };
    try {
      return withStockSheets_(payload, function() { return writePhysicalCount_(lines, auth.email); });
    } finally {
      lock.releaseLock();
    }
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// --- Close review ----------------------------------------------------------

/**
 * The close screen's data: completeness (has every active material been
 * counted this period), the variance list with tolerance flags applied,
 * and the totals a confirmed close will record. Read-only -- nothing is
 * written here.
 */
function computeStockCloseReview_() {
  var target = determineCloseablePeriod_();
  if (!target.period) return { error: 'No ledger activity yet -- nothing to close.' };
  if (periodIsClosed_(target.period)) return { error: 'This period is already closed.', period: target.period, alreadyClosed: true };

  var catalog = getMaterialCatalog_();
  var position = getStockPosition_();
  var counted = readPhysicalCounts_(target.period);
  var activeIds = Object.keys(catalog).filter(function(id) { return catalog[id].active; });
  var missing = activeIds.filter(function(id) { return !counted[id]; });

  var lines = [];
  var flaggedCount = 0;
  var varianceTotal = 0;
  var ledgerValue = 0;
  Object.keys(position.positions).forEach(function(id) { ledgerValue += position.positions[id].value; });

  activeIds.forEach(function(id) {
    var c = counted[id];
    if (!c) return;
    var pos = position.positions[id];
    var avgCost = pos && pos.qty ? pos.value / pos.qty : 0;
    var flagged = isStockVarianceFlagged_(c.varianceQty, avgCost, pos ? pos.value : 0);
    if (flagged) { flaggedCount++; varianceTotal += Math.abs(c.varianceValue); }
    lines.push({
      materialId: id,
      materialName: catalog[id].materialName,
      countedQty: c.countedQty,
      ledgerQty: c.ledgerQtyAtCount,
      varianceQty: c.varianceQty,
      varianceValue: c.varianceValue,
      flagged: flagged
    });
  });
  lines.sort(function(a, b) {
    if (a.flagged !== b.flagged) return a.flagged ? -1 : 1;
    return a.materialName.localeCompare(b.materialName);
  });

  return {
    period: target.period,
    isCurrentPeriod: target.isCurrentPeriod,
    readyToClose: missing.length === 0,
    missingMaterials: missing.map(function(id) { return catalog[id].materialName; }),
    materialsCounted: activeIds.length - missing.length,
    materialsTotal: activeIds.length,
    materialsFlagged: flaggedCount,
    varianceTotalValue: Math.round(varianceTotal * 100) / 100,
    ledgerValue: Math.round(ledgerValue * 100) / 100,
    lines: lines,
    tolerance: getStockCloseTolerance_()
  };
}

/**
 * The close screen's data: completeness (has every active material been
 * counted this period), the variance list with tolerance flags applied,
 * and the totals a confirmed close will record. Read-only -- nothing is
 * written here.
 */
function getStockCloseReview(payload) {
  try {
    var auth = authorizeStockAdmin_(payload);
    if (!auth.ok) return { error: auth.error, code: auth.code };
    return withStockSheets_(payload, computeStockCloseReview_);
  } catch (e) {
    return { error: e.toString() };
  }
}

/**
 * Freezes the period: writes the Period Close summary row and the
 * per-material Stock Snapshot (writeStockSnapshot_, Stock_Ledger.gs) in
 * one step, so a close is never left half-written. Requires the full
 * count to be complete -- per the locked "full count every month"
 * decision, there is no partial close. Does NOT require every variance
 * to be resolved: a flagged variance can be closed as-is (with a note),
 * since deciding whether to adjust for it is a human call this module
 * never makes on its own -- it just refuses to let a flagged variance
 * disappear unrecorded.
 * payload: {note?}.
 */
function writeStockClose_(note, byEmail) {
  var review = computeStockCloseReview_();
  if (review.error) return { success: false, error: review.error };
  if (!review.readyToClose) {
    return { success: false, error: 'Every active material needs a count before closing -- missing: ' + review.missingMaterials.join(', ') };
  }
  if (review.materialsFlagged > 0 && !(note || '').toString().trim()) {
    return { success: false, error: review.materialsFlagged + ' variance' + (review.materialsFlagged === 1 ? '' : 's') + ' still flagged -- add a note explaining them (or adjust the count) before closing.' };
  }

  var position = getStockPosition_();
  var seq = currentStockSeq_();
  var now = new Date();
  writeStockSnapshot_(review.period, position.positions, seq, byEmail);

  var row = [];
  for (var c = 0; c < PERIOD_CLOSE_HEADERS.length; c++) row.push('');
  row[PERIOD_CLOSE_COL['Period']]                = review.period;
  row[PERIOD_CLOSE_COL['Status']]                = 'closed';
  row[PERIOD_CLOSE_COL['Ledger Value']]          = review.ledgerValue;
  row[PERIOD_CLOSE_COL['Materials Counted']]     = review.materialsCounted;
  row[PERIOD_CLOSE_COL['Materials Flagged']]     = review.materialsFlagged;
  row[PERIOD_CLOSE_COL['Variance Total Value']]  = review.varianceTotalValue;
  row[PERIOD_CLOSE_COL['QBO Value']]             = ''; // populated once QuickBooks posting exists
  row[PERIOD_CLOSE_COL['Closed By']]             = byEmail || '';
  row[PERIOD_CLOSE_COL['Closed At']]             = now;
  row[PERIOD_CLOSE_COL['Note']]                  = (note || '').toString().trim();
  periodCloseSheet_().appendRow(row);

  return { success: true, period: review.period };
}

/**
 * Freezes the period: writes the Period Close summary row and the
 * per-material Stock Snapshot (writeStockSnapshot_, Stock_Ledger.gs) in
 * one step, so a close is never left half-written. Requires the full
 * count to be complete -- per the locked "full count every month"
 * decision, there is no partial close. Does NOT require every variance
 * to be resolved: a flagged variance can be closed as-is (with a note),
 * since deciding whether to adjust for it is a human call this module
 * never makes on its own -- it just refuses to let a flagged variance
 * disappear unrecorded.
 * payload: {note?}.
 */
function confirmStockClose(payload) {
  try {
    var auth = authorizeStockAdmin_(payload);
    if (!auth.ok) return { success: false, error: auth.error, code: auth.code };

    var lock = LockService.getScriptLock();
    if (!lock.tryLock(15000)) return { success: false, error: 'Server is busy - try again in a moment.' };
    try {
      return withStockSheets_(payload, function() { return writeStockClose_(payload.note, auth.email); });
    } finally {
      lock.releaseLock();
    }
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// --- Report ------------------------------------------------------------------

function readStockCloseHistory_() {
  var sheet = periodCloseSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { closes: [] };
  var data = sheet.getRange(2, 1, lastRow - 1, PERIOD_CLOSE_HEADERS.length).getValues();
  var closes = data.map(function(r) {
    return {
      period:             (r[PERIOD_CLOSE_COL['Period']] || '').toString(),
      status:             (r[PERIOD_CLOSE_COL['Status']] || '').toString(),
      ledgerValue:        parseFloat(r[PERIOD_CLOSE_COL['Ledger Value']]) || 0,
      materialsCounted:   parseFloat(r[PERIOD_CLOSE_COL['Materials Counted']]) || 0,
      materialsFlagged:   parseFloat(r[PERIOD_CLOSE_COL['Materials Flagged']]) || 0,
      varianceTotalValue: parseFloat(r[PERIOD_CLOSE_COL['Variance Total Value']]) || 0,
      qboValue:           r[PERIOD_CLOSE_COL['QBO Value']] === '' ? null : parseFloat(r[PERIOD_CLOSE_COL['QBO Value']]),
      closedBy:           (r[PERIOD_CLOSE_COL['Closed By']] || '').toString(),
      closedAt:           r[PERIOD_CLOSE_COL['Closed At']] instanceof Date ? r[PERIOD_CLOSE_COL['Closed At']].toISOString() : '',
      note:               (r[PERIOD_CLOSE_COL['Note']] || '').toString()
    };
  });
  closes.sort(function(a, b) { return b.period.localeCompare(a.period); });
  return { closes: closes };
}

/** Every closed period, most recent first. */
function getStockCloseHistory(payload) {
  try {
    var auth = authorizeCaller(payload, STOCK_READ_ROLES);
    if (!auth.ok) return { error: auth.error, code: auth.code, closes: [] };
    return withStockSheets_(payload, readStockCloseHistory_);
  } catch (e) {
    return { error: e.toString(), closes: [] };
  }
}

function readStockPeriodDetail_(period) {
  var sheet = ensureSheetWithHeaders_(stockSheetName_(STOCK_SNAPSHOT_SHEET), STOCK_SNAPSHOT_HEADERS);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { materials: [], period: period };
  var data = sheet.getRange(2, 1, lastRow - 1, STOCK_SNAPSHOT_HEADERS.length).getValues();
  var materials = data
    .filter(function(r) { return (r[STOCK_SNAPSHOT_COL['Period']] || '').toString().trim() === period; })
    .map(function(r) {
      return {
        materialId:   r[STOCK_SNAPSHOT_COL['Material Id']],
        materialName: r[STOCK_SNAPSHOT_COL['Material Name']],
        qty:          parseFloat(r[STOCK_SNAPSHOT_COL['Qty']]) || 0,
        value:        parseFloat(r[STOCK_SNAPSHOT_COL['Value']]) || 0
      };
    })
    .sort(function(a, b) { return a.materialName.localeCompare(b.materialName); });
  return { materials: materials, period: period };
}

/** Per-material detail for one closed period, from its Stock Snapshot rows. */
function getStockPeriodDetail(payload) {
  try {
    var auth = authorizeCaller(payload, STOCK_READ_ROLES);
    if (!auth.ok) return { error: auth.error, code: auth.code, materials: [] };

    var period = (payload.period || '').toString().trim();
    if (!period) return { error: 'Missing period.', materials: [] };

    return withStockSheets_(payload, function() { return readStockPeriodDetail_(period); });
  } catch (e) {
    return { error: e.toString(), materials: [] };
  }
}
