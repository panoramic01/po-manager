/**
 * Stock Ledger -- the append-only record of every physical stock movement,
 * and the FIFO engine that derives cost from it.
 *
 * This module is the authority on BOTH quantity and cost. QuickBooks holds
 * dollars only: stocked materials post against a small set of NON-INVENTORY
 * QBO items (one per category / COGS account) with the specific material
 * name carried in the line description. QBO therefore never builds cost
 * layers of its own, and there is exactly one FIFO calculation in the
 * business -- this one.
 *
 * Because of that, the ledger keys on our own Material Id, not a QBO Item
 * Id. The QBO posting item is resolved from the Material Catalog at post
 * time, so re-pointing a category at a different account later is a catalog
 * edit rather than a ledger rewrite.
 *
 * Nothing in here talks to QuickBooks. Posting lives in a separate module
 * so the engine below stays pure and testable -- run testStockLedger_()
 * from the Apps Script editor to exercise it without touching anything.
 */

// --- Sheets --------------------------------------------------------------

var STOCK_LEDGER_SHEET = 'Stock Ledger';
var STOCK_LEDGER_HEADERS = [
  'Seq', 'Ledger Id', 'Effective Date', 'Recorded At', 'Material Id', 'Material Name',
  'Qty Delta', 'Unit Cost', 'Move Type', 'Job Ref', 'Reverses', 'Source Doc',
  'QBO Txn Type', 'QBO Txn Id', 'Post Status', 'Period', 'By', 'Note'
];
var STOCK_LEDGER_COL = {};
STOCK_LEDGER_HEADERS.forEach(function(h, i) { STOCK_LEDGER_COL[h] = i; });

var STOCK_SNAPSHOT_SHEET = 'Stock Snapshot';
var STOCK_SNAPSHOT_HEADERS = [
  'Period', 'Material Id', 'Material Name', 'Qty', 'Value', 'Layers JSON',
  'Through Seq', 'Status', 'Closed At', 'Closed By'
];
var STOCK_SNAPSHOT_COL = {};
STOCK_SNAPSHOT_HEADERS.forEach(function(h, i) { STOCK_SNAPSHOT_COL[h] = i; });

/**
 * The material master. One row per stocked material -- this is the
 * vocabulary the ledger, the receiving screen and the pull screen all speak.
 * "QBO Post Item" is the non-inventory QBO item its dollars land on, shared
 * by every material in that category; the material's own name goes in the
 * posted line's description so the detail survives in QuickBooks without
 * QuickBooks having to track it as stock.
 */
var MATERIAL_CATALOG_SHEET = 'Material Catalog';
var MATERIAL_CATALOG_HEADERS = [
  'Material Id', 'Material Name', 'Unit', 'Category',
  'QBO Post Item Id', 'QBO Post Item Name', 'Active', 'Notes'
];
var MATERIAL_CATALOG_COL = {};
MATERIAL_CATALOG_HEADERS.forEach(function(h, i) { MATERIAL_CATALOG_COL[h] = i; });

/**
 * Every sheet this module touches resolves its name through here, so the
 * integration test can run the REAL read/write paths against disposable
 * "(TEST)" sheets in the same spreadsheet and delete them afterwards.
 * Production leaves the suffix empty. This exists because the pure engine
 * test proves the arithmetic but proves nothing about the sheet I/O, which
 * is where a column-order or type-coercion mistake would actually bite.
 */
var STOCK_SHEET_SUFFIX_ = '';

function stockSheetName_(base) {
  return base + STOCK_SHEET_SUFFIX_;
}

// --- Enums ---------------------------------------------------------------

/**
 * Move types. The sign is fixed per type and enforced at entry
 * (validateStockMove_) so a typo cannot quietly invert a movement.
 * COUNT_ADJUST is the only type allowed to go either way.
 *
 * RECEIPT_STOCK  vendor -> warehouse. Unit Cost from the vendor Bill line.
 * ISSUE_JOB      warehouse -> job. Unit Cost DERIVED, never supplied.
 * RETURN_PULLED  job -> warehouse, of stock we previously issued. Unit Cost
 *                is resolved at ENTRY time from the original issue and
 *                written onto the row, so the replay treats it as an
 *                ordinary receipt and never has to reach back across a
 *                snapshot boundary to price it.
 * RETURN_DIRECT  job -> warehouse, of material billed direct to the job and
 *                never in stock. Unit Cost from the original purchase (the
 *                Purchase Line Item Log has it).
 * COUNT_ADJUST   physical count correction. Never automatic -- a person
 *                decides every one, per the locked decision.
 * OPENING        one-time go-live balance.
 */
var STOCK_MOVE = {
  OPENING:       { sign:  1, needsCost: true,  needsJob: false },
  RECEIPT_STOCK: { sign:  1, needsCost: true,  needsJob: false },
  ISSUE_JOB:     { sign: -1, needsCost: false, needsJob: true  },
  RETURN_PULLED: { sign:  1, needsCost: true,  needsJob: true  },
  RETURN_DIRECT: { sign:  1, needsCost: true,  needsJob: true  },
  COUNT_ADJUST:  { sign:  0, needsCost: false, needsJob: false }
};

var STOCK_POST_PENDING = 'pending';
var STOCK_POST_POSTED  = 'posted';
var STOCK_POST_FAILED  = 'failed';
var STOCK_POST_VOID    = 'void';

// --- Money ---------------------------------------------------------------

/**
 * All money is carried as integer cents inside the engine. Unit costs stay
 * as 4dp decimals (that is how vendors quote them) but every extended amount
 * is computed with a single round to cents -- never accumulated as floats,
 * which is the classic way an inventory value drifts a few dollars a month
 * for no visible reason.
 */
var UNIT_COST_DP = 4;

function centsToDollars_(cents) {
  return Math.round(cents) / 100;
}

function extendCents_(qty, unitCost) {
  return Math.round((parseFloat(qty) || 0) * (parseFloat(unitCost) || 0) * 100);
}

function roundUnitCost_(v) {
  var f = Math.pow(10, UNIT_COST_DP);
  return Math.round((parseFloat(v) || 0) * f) / f;
}

// --- The FIFO engine (pure) ---------------------------------------------

/**
 * Replays a set of moves into per-material FIFO positions. PURE -- no sheet
 * reads, no clock, no QuickBooks. Everything it needs arrives as arguments,
 * which is what makes testStockLedger_() possible.
 *
 * moves: [{seq, ledgerId, effectiveDate: Date, materialId, materialName,
 *          qtyDelta, unitCost, moveType}]
 * openingLayers: { materialId: [{qty, unitCost}] } -- from the last
 *          snapshot, or {} to replay from the beginning of time.
 *
 * Returns {
 *   positions:  { materialId: {materialId, materialName, qty, valueCents, value, layers} },
 *   issueCosts: { ledgerId: {cogsCents, unitCost} },
 *   warnings:   [{ledgerId, materialId, message}]
 * }
 *
 * Layers are held as {qty, unitCost} and their value is always DERIVED
 * (qty x unitCost), never decremented in place -- a decremented running
 * value accumulates rounding error across a partial consumption, a derived
 * one cannot.
 */
function replayStockFifo_(moves, openingLayers) {
  var positions = {};
  var issueCosts = {};
  var warnings = [];
  var seedLayers = openingLayers || {};

  var ensure = function(materialId, materialName) {
    if (!positions[materialId]) {
      var seed = seedLayers[materialId] || [];
      positions[materialId] = {
        materialId: materialId,
        materialName: materialName || '',
        layers: seed.map(function(l) {
          return { qty: parseFloat(l.qty) || 0, unitCost: roundUnitCost_(l.unitCost) };
        })
      };
    }
    if (materialName && !positions[materialId].materialName) {
      positions[materialId].materialName = materialName;
    }
    return positions[materialId];
  };

  // Seed every material that has an opening layer, even if no move touches
  // it this period -- otherwise a slow-moving material silently vanishes
  // from the position report between snapshots.
  Object.keys(seedLayers).forEach(function(materialId) { ensure(materialId, ''); });

  sortStockMoves_(moves).forEach(function(m) {
    var pos = ensure(m.materialId, m.materialName);
    var qty = parseFloat(m.qtyDelta) || 0;
    if (!qty) return;

    if (qty > 0) {
      // Receipt of any kind. Whether a $0 receipt was ALLOWED to be
      // written is decided once, at append time (validateStockMove_) --
      // replay never re-litigates it. A row that made it into the ledger
      // replays faithfully forever, $0 included, so closing the zero-cost
      // window later never breaks a receipt legitimately written while it
      // was open.
      pos.layers.push({ qty: qty, unitCost: roundUnitCost_(m.unitCost) });
      return;
    }

    // Issue. Consume from the head of the queue until satisfied.
    var remaining = -qty;
    var cogsCents = 0;
    // Remembered as we go: a fully consumed layer is shifted off the queue,
    // so by the time a shortfall needs pricing the layer that would have
    // priced it is already gone. Without this the fallback finds no layers
    // and prices the shortfall at zero, silently understating COGS.
    var lastConsumedCost = 0;
    while (remaining > 0 && pos.layers.length && pos.layers[0].qty > 0) {
      var layer = pos.layers[0];
      var taken = Math.min(layer.qty, remaining);
      cogsCents += extendCents_(taken, layer.unitCost);
      if (layer.unitCost > 0) lastConsumedCost = layer.unitCost;
      layer.qty = layer.qty - taken;
      remaining = remaining - taken;
      if (layer.qty <= 0) pos.layers.shift();
    }

    if (remaining > 0) {
      // Going negative is hard-blocked at entry (see validateStockMove_),
      // so reaching here means a movement arrived from outside the app or a
      // count is wrong. Price it at the last known cost, record the
      // shortfall as a negative layer so the next receipt back-fills it,
      // and shout about it.
      var lastCost = lastConsumedCost > 0 ? lastConsumedCost : lastKnownUnitCost_(pos, seedLayers, m.materialId);
      cogsCents += extendCents_(remaining, lastCost);
      pos.layers.push({ qty: -remaining, unitCost: lastCost });
      warnings.push({
        ledgerId: m.ledgerId,
        materialId: m.materialId,
        message: 'Issued ' + remaining + ' more than was on hand. Priced at last known cost ' + lastCost + '. Investigate before closing the period.'
      });
    }

    issueCosts[m.ledgerId] = {
      cogsCents: cogsCents,
      unitCost: roundUnitCost_(cogsCents / 100 / (-qty))
    };
  });

  // Collapse to totals. Value is derived from the surviving layers, so it
  // always agrees with them by construction.
  Object.keys(positions).forEach(function(materialId) {
    var pos = positions[materialId];
    pos.layers = pos.layers.filter(function(l) { return l.qty !== 0; });
    pos.qty = pos.layers.reduce(function(s, l) { return s + l.qty; }, 0);
    pos.valueCents = pos.layers.reduce(function(s, l) { return s + extendCents_(l.qty, l.unitCost); }, 0);
    pos.value = centsToDollars_(pos.valueCents);
  });

  return { positions: positions, issueCosts: issueCosts, warnings: warnings };
}

/**
 * Deterministic ordering: effective date first (that is what FIFO means),
 * then Seq as the tiebreaker. Without the Seq tiebreaker two moves recorded
 * in the same second could replay in either order and produce different
 * costs on different runs -- which would make the monthly variance
 * unreproducible and impossible to argue with.
 */
function sortStockMoves_(moves) {
  return (moves || []).slice().sort(function(a, b) {
    var at = a.effectiveDate instanceof Date ? a.effectiveDate.getTime() : 0;
    var bt = b.effectiveDate instanceof Date ? b.effectiveDate.getTime() : 0;
    if (at !== bt) return at - bt;
    return (parseFloat(a.seq) || 0) - (parseFloat(b.seq) || 0);
  });
}

/** Best available cost for pricing a short issue: the layer just exhausted, else the seed, else 0. */
function lastKnownUnitCost_(pos, seedLayers, materialId) {
  for (var i = pos.layers.length - 1; i >= 0; i--) {
    if (pos.layers[i].unitCost > 0) return pos.layers[i].unitCost;
  }
  var seed = (seedLayers && seedLayers[materialId]) || [];
  for (var j = seed.length - 1; j >= 0; j--) {
    if (seed[j].unitCost > 0) return roundUnitCost_(seed[j].unitCost);
  }
  return 0;
}

// --- Validation ----------------------------------------------------------

/**
 * Gate every move before it reaches the sheet. Returns null when valid, an
 * error string otherwise. Deliberately strict: the ledger is append-only,
 * so a bad row cannot be edited away later -- it can only be corrected by a
 * further row, which is exactly the mess worth refusing up front.
 */
function validateStockMove_(move, onHandByMaterial) {
  if (!move || !move.materialId) return 'Missing material.';
  var spec = STOCK_MOVE[move.moveType];
  if (!spec) return 'Unknown move type "' + move.moveType + '".';

  var qty = parseFloat(move.qtyDelta);
  if (!qty || isNaN(qty)) return 'Quantity must be a non-zero number.';
  if (spec.sign > 0 && qty < 0) return move.moveType + ' must be a positive quantity.';
  if (spec.sign < 0 && qty > 0) return move.moveType + ' must be a negative quantity.';

  // A COUNT_ADJUST that ADDS stock is a receipt and must be priced -- there
  // is no cost layer to inherit from. One that REMOVES stock consumes
  // existing layers like any issue, so supplying a cost would let a caller
  // override FIFO. Hence the rule depends on direction, not just type.
  var needsCost = spec.needsCost || (move.moveType === 'COUNT_ADJUST' && qty > 0);
  if (needsCost) {
    var cost = roundUnitCost_(move.unitCost);
    // Ordinary receipt-shaped types (spec.needsCost) allow $0 -- material
    // already in the warehouse whose cost was already expensed on a prior
    // job can be brought in without double-counting it. A COUNT_ADJUST
    // that adds stock is a correction against a physical count, not a
    // bulk load, and always needs a real, defensible number.
    var costOk = spec.needsCost ? cost >= 0 : cost > 0;
    if (!costOk) {
      return move.moveType + (move.moveType === 'COUNT_ADJUST' ? ' that adds stock' : '') +
        (spec.needsCost ? ' needs a unit cost of 0 or more.' : ' needs a unit cost greater than zero.');
    }
  }
  if (!needsCost && move.unitCost) {
    return move.moveType + ' derives its cost from the ledger -- do not supply a unit cost.';
  }
  if (spec.needsJob && !(move.jobRef || '').toString().trim()) {
    return move.moveType + ' needs a job reference.';
  }
  if (!(move.effectiveDate instanceof Date) || isNaN(move.effectiveDate.getTime())) {
    return 'Missing or invalid effective date.';
  }

  // Hard stop on negative stock. You cannot pull what is not on the shelf;
  // if the app says 0 and the shelf has 5 then the COUNT is wrong, and that
  // is a COUNT_ADJUST with someone's name on it, not a negative issue.
  if (qty < 0 && onHandByMaterial) {
    var onHand = parseFloat(onHandByMaterial[move.materialId]) || 0;
    if (-qty > onHand) {
      return 'Only ' + onHand + ' of "' + (move.materialName || move.materialId) + '" on hand -- cannot move ' + (-qty) + '. Count it and post an adjustment if the shelf disagrees.';
    }
  }
  return null;
}

// --- Sheet I/O -----------------------------------------------------------

function stockLedgerSheet_() {
  return ensureSheetWithHeaders_(stockSheetName_(STOCK_LEDGER_SHEET), STOCK_LEDGER_HEADERS);
}

function periodOf_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM');
}

function ledgerRowToMove_(row) {
  var eff = row[STOCK_LEDGER_COL['Effective Date']];
  var rawCost = row[STOCK_LEDGER_COL['Unit Cost']];
  return {
    seq:           parseFloat(row[STOCK_LEDGER_COL['Seq']]) || 0,
    ledgerId:      (row[STOCK_LEDGER_COL['Ledger Id']] || '').toString(),
    effectiveDate: eff instanceof Date ? eff : new Date(eff),
    materialId:    (row[STOCK_LEDGER_COL['Material Id']] || '').toString().trim(),
    materialName:  (row[STOCK_LEDGER_COL['Material Name']] || '').toString(),
    qtyDelta:      parseFloat(row[STOCK_LEDGER_COL['Qty Delta']]) || 0,
    unitCost:      rawCost === '' || rawCost == null ? null : parseFloat(rawCost),
    moveType:      (row[STOCK_LEDGER_COL['Move Type']] || '').toString().trim(),
    jobRef:        (row[STOCK_LEDGER_COL['Job Ref']] || '').toString(),
    reverses:      (row[STOCK_LEDGER_COL['Reverses']] || '').toString(),
    postStatus:    (row[STOCK_LEDGER_COL['Post Status']] || '').toString().trim(),
    period:        (row[STOCK_LEDGER_COL['Period']] || '').toString().trim()
  };
}

/** Every ledger row with Seq greater than afterSeq. Voided rows never replay. */
function readStockLedgerMoves_(afterSeq) {
  var sheet = stockLedgerSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var data = sheet.getRange(2, 1, lastRow - 1, STOCK_LEDGER_HEADERS.length).getValues();
  var out = [];
  for (var i = 0; i < data.length; i++) {
    var move = ledgerRowToMove_(data[i]);
    if (!move.materialId || !move.moveType) continue;
    if (move.postStatus === STOCK_POST_VOID) continue;
    if (afterSeq && move.seq <= afterSeq) continue;
    out.push(move);
  }
  return out;
}

/**
 * Appends validated moves and hands back the rows as written. Takes the
 * script lock and re-derives on-hand INSIDE it, so two runners pulling the
 * same material at the same moment cannot both pass the on-hand check --
 * the race the previewMaterialPull / pushMaterialPull pair allows today.
 *
 * Rows land as 'pending'. Posting to QuickBooks happens afterwards and
 * stamps the txn id, so an Intuit outage can never lose the fact that
 * material physically moved.
 */
function appendStockMoves_(moves, byEmail) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { success: false, error: 'Server is busy - try again in a moment.' };
  try {
    var position = getStockPosition_();
    var onHand = {};
    Object.keys(position.positions).forEach(function(id) {
      onHand[id] = position.positions[id].qty;
    });

    // Validate the batch against a RUNNING on-hand, so two lines pulling the
    // same material in one submission cannot jointly exceed stock -- a
    // per-line check against a fixed starting number would let them.
    var prepared = [];
    for (var i = 0; i < moves.length; i++) {
      var m = moves[i];
      var err = validateStockMove_(m, onHand);
      if (err) return { success: false, error: err, failedIndex: i };
      onHand[m.materialId] = (parseFloat(onHand[m.materialId]) || 0) + (parseFloat(m.qtyDelta) || 0);
      prepared.push(m);
    }

    var sheet = stockLedgerSheet_();
    var seq = nextStockSeq_(sheet);
    var now = new Date();
    var stamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyMMddHHmmss');
    var rows = [];
    var written = [];

    prepared.forEach(function(m, idx) {
      var ledgerId = 'SL-' + stamp + '-' + (idx + 1);
      var seqNo = seq + idx;
      var row = [];
      for (var c = 0; c < STOCK_LEDGER_HEADERS.length; c++) row.push('');
      row[STOCK_LEDGER_COL['Seq']]            = seqNo;
      row[STOCK_LEDGER_COL['Ledger Id']]      = ledgerId;
      row[STOCK_LEDGER_COL['Effective Date']] = m.effectiveDate;
      row[STOCK_LEDGER_COL['Recorded At']]    = now;
      row[STOCK_LEDGER_COL['Material Id']]    = m.materialId;
      row[STOCK_LEDGER_COL['Material Name']]  = m.materialName || '';
      row[STOCK_LEDGER_COL['Qty Delta']]      = m.qtyDelta;
      row[STOCK_LEDGER_COL['Unit Cost']]      = m.unitCost != null ? roundUnitCost_(m.unitCost) : '';
      row[STOCK_LEDGER_COL['Move Type']]      = m.moveType;
      row[STOCK_LEDGER_COL['Job Ref']]        = m.jobRef || '';
      row[STOCK_LEDGER_COL['Reverses']]       = m.reverses || '';
      row[STOCK_LEDGER_COL['Source Doc']]     = m.sourceDoc || '';
      row[STOCK_LEDGER_COL['Post Status']]    = STOCK_POST_PENDING;
      row[STOCK_LEDGER_COL['Period']]         = periodOf_(m.effectiveDate);
      row[STOCK_LEDGER_COL['By']]             = byEmail || '';
      row[STOCK_LEDGER_COL['Note']]           = m.note || '';
      rows.push(row);
      written.push({ ledgerId: ledgerId, seq: seqNo, materialId: m.materialId, qtyDelta: m.qtyDelta });
    });

    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, STOCK_LEDGER_HEADERS.length).setValues(rows);
    return { success: true, rows: written };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

/** Next sequence number. Called under the script lock only. */
function nextStockSeq_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 1;
  var seqs = sheet.getRange(2, STOCK_LEDGER_COL['Seq'] + 1, lastRow - 1, 1).getValues();
  var max = 0;
  for (var i = 0; i < seqs.length; i++) {
    var v = parseFloat(seqs[i][0]) || 0;
    if (v > max) max = v;
  }
  return max + 1;
}

/**
 * Current position for every material: quantity, value, and surviving cost
 * layers. Replays from the most recent snapshot forward rather than from
 * the beginning of time -- that bound is the whole reason the monthly close
 * writes a snapshot, and it is what keeps this fast enough to call on every
 * screen load instead of round-tripping to QuickBooks the way
 * getWarehouseItemsOnHand_ does today.
 */
function getStockPosition_() {
  var snap = readLatestStockSnapshot_();
  var moves = readStockLedgerMoves_(snap.throughSeq);
  var result = replayStockFifo_(moves, snap.layersByItem);
  result.fromSnapshot = snap.period || null;
  result.throughSeq = snap.throughSeq;
  return result;
}

/** The newest confirmed snapshot, as replay seed. Empty seed when none exists yet. */
function readLatestStockSnapshot_() {
  var sheet = ensureSheetWithHeaders_(stockSheetName_(STOCK_SNAPSHOT_SHEET), STOCK_SNAPSHOT_HEADERS);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { period: null, throughSeq: 0, layersByItem: {} };

  var data = sheet.getRange(2, 1, lastRow - 1, STOCK_SNAPSHOT_HEADERS.length).getValues();
  var bestPeriod = null;
  var bestSeq = 0;
  data.forEach(function(r) {
    var period = (r[STOCK_SNAPSHOT_COL['Period']] || '').toString().trim();
    var status = (r[STOCK_SNAPSHOT_COL['Status']] || '').toString().trim();
    // A stale snapshot is one a backdated entry invalidated -- ignoring it
    // here means the replay falls back to the previous good one rather than
    // seeding from numbers known to be wrong.
    if (!period || status === 'stale') return;
    if (!bestPeriod || period > bestPeriod) {
      bestPeriod = period;
      bestSeq = parseFloat(r[STOCK_SNAPSHOT_COL['Through Seq']]) || 0;
    }
  });
  if (!bestPeriod) return { period: null, throughSeq: 0, layersByItem: {} };

  var layersByItem = {};
  data.forEach(function(r) {
    if ((r[STOCK_SNAPSHOT_COL['Period']] || '').toString().trim() !== bestPeriod) return;
    var materialId = (r[STOCK_SNAPSHOT_COL['Material Id']] || '').toString().trim();
    if (!materialId) return;
    var layers = [];
    try { layers = JSON.parse(r[STOCK_SNAPSHOT_COL['Layers JSON']] || '[]'); } catch (e) { layers = []; }
    layersByItem[materialId] = layers;
  });

  return { period: bestPeriod, throughSeq: bestSeq, layersByItem: layersByItem };
}

/**
 * Material Id -> catalog entry, including the non-inventory QBO item its
 * dollars post against. Read by the posting module; kept here because the
 * catalog defines the ledger's vocabulary.
 */
function getMaterialCatalog_() {
  var out = {};
  var sheet = ensureSheetWithHeaders_(stockSheetName_(MATERIAL_CATALOG_SHEET), MATERIAL_CATALOG_HEADERS);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return out;
  var data = sheet.getRange(2, 1, lastRow - 1, MATERIAL_CATALOG_HEADERS.length).getValues();
  data.forEach(function(r) {
    var id = (r[MATERIAL_CATALOG_COL['Material Id']] || '').toString().trim();
    if (!id) return;
    var active = r[MATERIAL_CATALOG_COL['Active']];
    out[id] = {
      materialId:      id,
      materialName:    (r[MATERIAL_CATALOG_COL['Material Name']] || '').toString().trim(),
      unit:            (r[MATERIAL_CATALOG_COL['Unit']] || '').toString().trim(),
      category:        (r[MATERIAL_CATALOG_COL['Category']] || '').toString().trim(),
      // .toString() matters: Sheets stores a numeric-looking QBO id as a
      // Number, and QBO's API rejects a bare JSON number where its schema
      // wants a string -- the same trap stagingRowToObject_ documents.
      qboPostItemId:   (r[MATERIAL_CATALOG_COL['QBO Post Item Id']] || '').toString().trim(),
      qboPostItemName: (r[MATERIAL_CATALOG_COL['QBO Post Item Name']] || '').toString().trim(),
      active:          active === '' || active == null ? true : (active === true || /^(true|yes|y|1)$/i.test(active.toString().trim()))
    };
  });
  return out;
}

// --- Self-test -----------------------------------------------------------

/**
 * Exercises the engine against hand-computed expectations. Run from the
 * Apps Script editor; touches nothing, reads nothing, posts nothing.
 * Logs a PASS/FAIL line per case and returns the failure count.
 */
function testStockLedger_() {
  var failures = 0;
  var d = function(s) { return new Date(s + 'T12:00:00'); };
  var check = function(name, actual, expected) {
    var a = JSON.stringify(actual), e = JSON.stringify(expected);
    var ok = a === e;
    if (!ok) failures++;
    Logger.log((ok ? 'PASS  ' : 'FAIL  ') + name + (ok ? '' : '\n   expected ' + e + '\n   actual   ' + a));
  };

  // Two receipts at different costs, then an issue spanning both layers.
  // 10 @ 4.00 then 10 @ 6.00; issue 15 => 10x4.00 + 5x6.00 = 70.00 COGS,
  // leaving 5 @ 6.00 = 30.00 on hand.
  var r1 = replayStockFifo_([
    { seq: 1, ledgerId: 'A', effectiveDate: d('2026-08-01'), materialId: 'M1', qtyDelta:  10, unitCost: 4, moveType: 'RECEIPT_STOCK' },
    { seq: 2, ledgerId: 'B', effectiveDate: d('2026-08-02'), materialId: 'M1', qtyDelta:  10, unitCost: 6, moveType: 'RECEIPT_STOCK' },
    { seq: 3, ledgerId: 'C', effectiveDate: d('2026-08-03'), materialId: 'M1', qtyDelta: -15, unitCost: null, moveType: 'ISSUE_JOB' }
  ], {});
  check('spans two layers: COGS', r1.issueCosts['C'].cogsCents, 7000);
  check('spans two layers: qty left', r1.positions['M1'].qty, 5);
  check('spans two layers: value left', r1.positions['M1'].valueCents, 3000);

  // Out-of-order entry: a receipt recorded later but dated EARLIER must be
  // consumed first. This is what makes date-then-seq ordering load-bearing
  // rather than cosmetic.
  var r2 = replayStockFifo_([
    { seq: 1, ledgerId: 'A', effectiveDate: d('2026-08-10'), materialId: 'M1', qtyDelta:  5, unitCost: 10, moveType: 'RECEIPT_STOCK' },
    { seq: 2, ledgerId: 'B', effectiveDate: d('2026-08-05'), materialId: 'M1', qtyDelta:  5, unitCost:  2, moveType: 'RECEIPT_STOCK' },
    { seq: 3, ledgerId: 'C', effectiveDate: d('2026-08-11'), materialId: 'M1', qtyDelta: -5, unitCost: null, moveType: 'ISSUE_JOB' }
  ], {});
  check('backdated receipt consumed first', r2.issueCosts['C'].cogsCents, 1000);

  // A return re-enters at the cost it was issued at, so the pair is a
  // perfect round trip: value ends exactly where it started. This is the
  // drift bug the Bill + VendorCredit pair produces today by returning at
  // Pricing-sheet cost instead of issued cost.
  var r3 = replayStockFifo_([
    { seq: 1, ledgerId: 'A', effectiveDate: d('2026-08-01'), materialId: 'M1', qtyDelta: 10, unitCost: 4, moveType: 'RECEIPT_STOCK' },
    { seq: 2, ledgerId: 'B', effectiveDate: d('2026-08-02'), materialId: 'M1', qtyDelta: -5, unitCost: null, moveType: 'ISSUE_JOB' },
    { seq: 3, ledgerId: 'C', effectiveDate: d('2026-08-03'), materialId: 'M1', qtyDelta:  5, unitCost: 4, moveType: 'RETURN_PULLED', reverses: 'B' }
  ], {});
  check('return round-trips exactly: qty', r3.positions['M1'].qty, 10);
  check('return round-trips exactly: value', r3.positions['M1'].valueCents, 4000);

  // Fractional quantity against a 4dp unit cost: 3 x 12.3456 = 37.0368,
  // rounded once at the line to 3704 cents rather than accumulated.
  var r4 = replayStockFifo_([
    { seq: 1, ledgerId: 'A', effectiveDate: d('2026-08-01'), materialId: 'M1', qtyDelta: 3, unitCost: 12.3456, moveType: 'RECEIPT_STOCK' }
  ], {});
  check('4dp cost rounds once', r4.positions['M1'].valueCents, 3704);

  // Snapshot seed replaces replaying from the beginning of time.
  var r5 = replayStockFifo_([
    { seq: 9, ledgerId: 'Z', effectiveDate: d('2026-09-02'), materialId: 'M1', qtyDelta: -2, unitCost: null, moveType: 'ISSUE_JOB' }
  ], { M1: [{ qty: 5, unitCost: 8 }] });
  check('snapshot seed: COGS', r5.issueCosts['Z'].cogsCents, 1600);
  check('snapshot seed: qty left', r5.positions['M1'].qty, 3);

  // A material with an opening layer and no movement must still be reported.
  var r6 = replayStockFifo_([], { M9: [{ qty: 7, unitCost: 3 }] });
  check('idle material still reported', r6.positions['M9'].qty, 7);

  // Replay is purely faithful: whether a $0 receipt was ALLOWED to be
  // written is a validateStockMove_/append-time decision, never replay's --
  // a $0 row that made it into the ledger replays as a real, qty-bearing
  // layer at $0 cost.
  var r7 = replayStockFifo_([
    { seq: 1, ledgerId: 'A', effectiveDate: d('2026-08-01'), materialId: 'M1', qtyDelta: 5, unitCost: 0, moveType: 'RECEIPT_STOCK' }
  ], {});
  check('replay admits a $0 receipt as a real layer', r7.positions['M1'].qty, 5);
  check('replay prices that layer at $0', r7.positions['M1'].valueCents, 0);

  // A short issue is priced at last known cost and flagged loudly.
  var r8 = replayStockFifo_([
    { seq: 1, ledgerId: 'A', effectiveDate: d('2026-08-01'), materialId: 'M1', qtyDelta:  2, unitCost: 5, moveType: 'RECEIPT_STOCK' },
    { seq: 2, ledgerId: 'B', effectiveDate: d('2026-08-02'), materialId: 'M1', qtyDelta: -5, unitCost: null, moveType: 'ISSUE_JOB' }
  ], {});
  check('short issue: COGS at last cost', r8.issueCosts['B'].cogsCents, 2500);
  check('short issue: goes negative', r8.positions['M1'].qty, -3);
  check('short issue: warned', r8.warnings.length, 1);

  // A later receipt back-fills the negative layer rather than sitting beside it.
  var r9 = replayStockFifo_([
    { seq: 1, ledgerId: 'A', effectiveDate: d('2026-08-01'), materialId: 'M1', qtyDelta:  2, unitCost: 5, moveType: 'RECEIPT_STOCK' },
    { seq: 2, ledgerId: 'B', effectiveDate: d('2026-08-02'), materialId: 'M1', qtyDelta: -5, unitCost: null, moveType: 'ISSUE_JOB' },
    { seq: 3, ledgerId: 'C', effectiveDate: d('2026-08-03'), materialId: 'M1', qtyDelta: 10, unitCost: 5, moveType: 'RECEIPT_STOCK' }
  ], {});
  check('negative layer nets against later receipt', r9.positions['M1'].qty, 7);

  // Validation rejects the shapes that would corrupt the ledger.
  check('rejects issue past on-hand',
    !!validateStockMove_({ materialId: 'M1', moveType: 'ISSUE_JOB', qtyDelta: -5, jobRef: 'J1', effectiveDate: d('2026-08-01') }, { M1: 2 }), true);
  check('rejects supplied cost on an issue',
    !!validateStockMove_({ materialId: 'M1', moveType: 'ISSUE_JOB', qtyDelta: -1, unitCost: 5, jobRef: 'J1', effectiveDate: d('2026-08-01') }, { M1: 9 }), true);
  // A receipt with unitCost omitted defaults to $0, same as an explicit 0
  // -- receipts now allow $0 by default. The client's own numeric input
  // validation is what actually guards against an accidentally-blank field
  // reaching here in practice.
  check('receipt with no cost defaults to $0, not refused',
    validateStockMove_({ materialId: 'M1', moveType: 'RECEIPT_STOCK', qtyDelta: 5, effectiveDate: d('2026-08-01') }, {}), null);
  check('rejects wrong sign',
    !!validateStockMove_({ materialId: 'M1', moveType: 'RECEIPT_STOCK', qtyDelta: -5, unitCost: 2, effectiveDate: d('2026-08-01') }, {}), true);
  check('rejects issue with no job',
    !!validateStockMove_({ materialId: 'M1', moveType: 'ISSUE_JOB', qtyDelta: -1, effectiveDate: d('2026-08-01') }, { M1: 9 }), true);
  check('accepts a good issue',
    validateStockMove_({ materialId: 'M1', moveType: 'ISSUE_JOB', qtyDelta: -2, jobRef: 'J1', effectiveDate: d('2026-08-01') }, { M1: 9 }), null);

  // Receipt-shaped types allow $0 by default -- material already in the
  // warehouse whose cost was already expensed on a prior job. A negative
  // cost is still always rejected, and a COUNT_ADJUST that adds stock still
  // always needs a real, defensible number -- a count correction is not a
  // bulk opening load.
  check('accepts $0 on a receipt',
    validateStockMove_({ materialId: 'M1', moveType: 'RECEIPT_STOCK', qtyDelta: 5, unitCost: 0, effectiveDate: d('2026-08-01') }, {}), null);
  check('still rejects a NEGATIVE cost on a receipt',
    !!validateStockMove_({ materialId: 'M1', moveType: 'RECEIPT_STOCK', qtyDelta: 5, unitCost: -1, effectiveDate: d('2026-08-01') }, {}), true);
  check('a COUNT_ADJUST increase still always needs a real cost',
    !!validateStockMove_({ materialId: 'M1', moveType: 'COUNT_ADJUST', qtyDelta: 5, effectiveDate: d('2026-08-01') }, { M1: 0 }), true);

  Logger.log(failures === 0 ? 'ALL PASS' : (failures + ' FAILED'));
  return failures;
}

// --- Snapshot writing ----------------------------------------------------

/** Highest Seq currently in the ledger. A snapshot records it so the next replay knows where to resume. */
function currentStockSeq_() {
  var sheet = stockLedgerSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  var seqs = sheet.getRange(2, STOCK_LEDGER_COL['Seq'] + 1, lastRow - 1, 1).getValues();
  var max = 0;
  for (var i = 0; i < seqs.length; i++) {
    var v = parseFloat(seqs[i][0]) || 0;
    if (v > max) max = v;
  }
  return max;
}

/**
 * Freezes a period's position as the next replay origin. Writes the full
 * layer state, not just qty and value -- a snapshot recording only totals
 * would lose the cost layers and force exactly the full-history replay it
 * exists to avoid.
 */
function writeStockSnapshot_(period, positions, throughSeq, byEmail) {
  var sheet = ensureSheetWithHeaders_(stockSheetName_(STOCK_SNAPSHOT_SHEET), STOCK_SNAPSHOT_HEADERS);
  var now = new Date();
  var rows = [];
  Object.keys(positions).forEach(function(materialId) {
    var pos = positions[materialId];
    var row = [];
    for (var c = 0; c < STOCK_SNAPSHOT_HEADERS.length; c++) row.push('');
    row[STOCK_SNAPSHOT_COL['Period']]        = period;
    row[STOCK_SNAPSHOT_COL['Material Id']]   = materialId;
    row[STOCK_SNAPSHOT_COL['Material Name']] = pos.materialName || '';
    row[STOCK_SNAPSHOT_COL['Qty']]           = pos.qty;
    row[STOCK_SNAPSHOT_COL['Value']]         = pos.value;
    row[STOCK_SNAPSHOT_COL['Layers JSON']]   = JSON.stringify(pos.layers);
    row[STOCK_SNAPSHOT_COL['Through Seq']]   = throughSeq;
    row[STOCK_SNAPSHOT_COL['Status']]        = 'soft-closed';
    row[STOCK_SNAPSHOT_COL['Closed At']]     = now;
    row[STOCK_SNAPSHOT_COL['Closed By']]     = byEmail || '';
    rows.push(row);
  });
  if (!rows.length) return { success: true, rows: 0 };
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, STOCK_SNAPSHOT_HEADERS.length).setValues(rows);
  return { success: true, rows: rows.length };
}
