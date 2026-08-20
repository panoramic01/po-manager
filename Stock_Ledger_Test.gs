/**
 * Stock Ledger tests. Nothing in here runs on its own and nothing here is
 * reachable from the client -- run testStockAll_() from the Apps Script
 * editor and read the Execution log.
 *
 * Two suites, deliberately separate:
 *
 *   testStockLedger_()             lives in Stock_Ledger.gs. Pure arithmetic,
 *                                  no sheets, no clock. Proves FIFO.
 *   testStockLedgerIntegration_()  below. Drives the REAL append / read /
 *                                  snapshot code against disposable "(TEST)"
 *                                  sheets, then deletes them.
 *
 * The integration suite exists because the pure one proves the arithmetic
 * and proves nothing about the sheet layer -- which is where a column
 * ordering slip, or one of the Sheets type coercions that turn an id into a
 * Number and a blank into an empty string, would actually bite.
 *
 * Production sheets are never opened: the name suffix is set before anything
 * runs and restored in a finally block even if an assertion throws.
 */

function stockTestSheetNames_() {
  return [STOCK_LEDGER_SHEET, STOCK_SNAPSHOT_SHEET, MATERIAL_CATALOG_SHEET].map(function(base) {
    return base + ' (TEST)';
  });
}

function dropStockTestSheets_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  stockTestSheetNames_().forEach(function(name) {
    var sh = ss.getSheetByName(name);
    if (sh) ss.deleteSheet(sh);
  });
}

/** Safe to run repeatedly. Returns the failure count. */
function testStockLedgerIntegration_() {
  var failures = 0;
  var check = function(name, actual, expected) {
    var a = JSON.stringify(actual), e = JSON.stringify(expected);
    var ok = a === e;
    if (!ok) failures++;
    Logger.log((ok ? 'PASS  ' : 'FAIL  ') + name + (ok ? '' : '  expected ' + e + ' / actual ' + a));
  };
  var d = function(str) { return new Date(str + 'T12:00:00'); };
  var qtyOf = function(res, id) { return res.positions[id] ? res.positions[id].qty : null; };
  var valOf = function(res, id) { return res.positions[id] ? res.positions[id].valueCents : null; };
  var who = 'test@example.com';

  STOCK_SHEET_SUFFIX_ = ' (TEST)';
  try {
    dropStockTestSheets_();

    // --- catalog -------------------------------------------------------
    var cat = ensureSheetWithHeaders_(stockSheetName_(MATERIAL_CATALOG_SHEET), MATERIAL_CATALOG_HEADERS);
    cat.getRange(2, 1, 2, MATERIAL_CATALOG_HEADERS.length).setValues([
      ['M-SHINGLE',  'Architectural shingle',  'SQ', 'Roofing', 101, 'Roofing Materials', true, ''],
      ['M-UNDERLAY', 'Synthetic underlayment', 'RL', 'Roofing', 101, 'Roofing Materials', true, '']
    ]);
    var catalog = getMaterialCatalog_();
    check('catalog loads both materials', Object.keys(catalog).length, 2);
    // Sheets hands a numeric-looking id back as a Number, and QBO rejects a
    // bare JSON number where its schema wants a string.
    check('catalog coerces QBO id to string', typeof catalog['M-SHINGLE'].qboPostItemId, 'string');
    check('catalog reads Active as boolean', catalog['M-UNDERLAY'].active, true);

    // --- receiving -----------------------------------------------------
    var rec = appendStockMoves_([
      { materialId: 'M-SHINGLE',  materialName: 'Architectural shingle',  moveType: 'RECEIPT_STOCK', qtyDelta: 20, unitCost: 95.50, effectiveDate: d('2026-08-01'), sourceDoc: 'PO-1' },
      { materialId: 'M-UNDERLAY', materialName: 'Synthetic underlayment', moveType: 'RECEIPT_STOCK', qtyDelta: 10, unitCost: 42.00, effectiveDate: d('2026-08-01'), sourceDoc: 'PO-1' }
    ], who);
    check('receipt accepted', rec.success, true);
    check('receipt assigns sequential Seq', [rec.rows[0].seq, rec.rows[1].seq], [1, 2]);

    var p1 = getStockPosition_();
    check('shingle qty after receipt', qtyOf(p1, 'M-SHINGLE'), 20);
    check('shingle value after receipt', valOf(p1, 'M-SHINGLE'), 191000);
    check('underlay value after receipt', valOf(p1, 'M-UNDERLAY'), 42000);

    // --- second cost layer, then an issue spanning both ----------------
    appendStockMoves_([
      { materialId: 'M-SHINGLE', materialName: 'Architectural shingle', moveType: 'RECEIPT_STOCK', qtyDelta: 10, unitCost: 102.00, effectiveDate: d('2026-08-05'), sourceDoc: 'PO-2' }
    ], who);
    appendStockMoves_([
      { materialId: 'M-SHINGLE', materialName: 'Architectural shingle', moveType: 'ISSUE_JOB', qtyDelta: -25, jobRef: 'JOB-100', effectiveDate: d('2026-08-06') }
    ], who);

    // 20 x 95.50 + 5 x 102.00 consumed; 5 x 102.00 = 510.00 left.
    var p2 = getStockPosition_();
    check('shingle qty after issue', qtyOf(p2, 'M-SHINGLE'), 5);
    check('shingle value after issue', valOf(p2, 'M-SHINGLE'), 51000);

    // --- the on-hand hard stop -----------------------------------------
    var over = appendStockMoves_([
      { materialId: 'M-SHINGLE', materialName: 'Architectural shingle', moveType: 'ISSUE_JOB', qtyDelta: -99, jobRef: 'JOB-101', effectiveDate: d('2026-08-07') }
    ], who);
    check('over-issue blocked', over.success, false);
    check('blocked issue wrote nothing', qtyOf(getStockPosition_(), 'M-SHINGLE'), 5);

    // Two lines that each fit but jointly exceed stock must also be refused
    // -- that is the running-total check inside appendStockMoves_, and a
    // per-line check against a fixed starting number would let them through.
    var joint = appendStockMoves_([
      { materialId: 'M-SHINGLE', materialName: 'Architectural shingle', moveType: 'ISSUE_JOB', qtyDelta: -3, jobRef: 'JOB-102', effectiveDate: d('2026-08-07') },
      { materialId: 'M-SHINGLE', materialName: 'Architectural shingle', moveType: 'ISSUE_JOB', qtyDelta: -3, jobRef: 'JOB-102', effectiveDate: d('2026-08-07') }
    ], who);
    check('jointly-over batch blocked', joint.success, false);
    check('jointly-over names the failing line', joint.failedIndex, 1);
    check('blocked batch wrote nothing', qtyOf(getStockPosition_(), 'M-SHINGLE'), 5);

    // --- return round trip ---------------------------------------------
    appendStockMoves_([
      { materialId: 'M-SHINGLE', materialName: 'Architectural shingle', moveType: 'RETURN_PULLED', qtyDelta: 2, unitCost: 102.00, jobRef: 'JOB-100', effectiveDate: d('2026-08-08') }
    ], who);
    var p3 = getStockPosition_();
    check('shingle qty after return', qtyOf(p3, 'M-SHINGLE'), 7);
    check('shingle value after return', valOf(p3, 'M-SHINGLE'), 71400);

    // --- count adjustments ---------------------------------------------
    var shrink = appendStockMoves_([
      { materialId: 'M-SHINGLE', materialName: 'Architectural shingle', moveType: 'COUNT_ADJUST', qtyDelta: -1, effectiveDate: d('2026-08-09'), note: 'monthly count' }
    ], who);
    check('count shrink accepted', shrink.success, true);
    check('shingle qty after shrink', qtyOf(getStockPosition_(), 'M-SHINGLE'), 6);

    var foundNoCost = appendStockMoves_([
      { materialId: 'M-SHINGLE', materialName: 'Architectural shingle', moveType: 'COUNT_ADJUST', qtyDelta: 1, effectiveDate: d('2026-08-09') }
    ], who);
    check('count increase without a cost is refused', foundNoCost.success, false);

    // --- zero-cost opening window --------------------------------------
    // Default (no opts): a $0 receipt is refused, matching production
    // behavior when the Script Property window is closed.
    var zeroClosed = appendStockMoves_([
      { materialId: 'M-UNDERLAY', materialName: 'Synthetic underlayment', moveType: 'RECEIPT_STOCK', qtyDelta: 8, unitCost: 0, effectiveDate: d('2026-08-10') }
    ], who);
    check('$0 receipt refused with the window closed', zeroClosed.success, false);

    // With the window explicitly open: the receipt is accepted, the
    // quantity is real (not silently dropped), and it costs exactly $0 --
    // this is the case that a naive validation-only fix would get wrong,
    // since the OLD replay engine used to skip a $0 receipt's quantity too.
    var zeroOpen = appendStockMoves_([
      { materialId: 'M-UNDERLAY', materialName: 'Synthetic underlayment', moveType: 'RECEIPT_STOCK', qtyDelta: 8, unitCost: 0, effectiveDate: d('2026-08-10'), note: 'opening balance -- already expensed on a prior job' }
    ], who, { allowZeroCost: true });
    check('$0 receipt accepted with the window open', zeroOpen.success, true);
    var afterZero = getStockPosition_();
    // 10 @ 42.00 (42000 cents) already on hand; +8 @ $0 adds quantity but
    // not value -- 18 on hand, value unchanged at 42000.
    check('$0 receipt quantity actually lands', qtyOf(afterZero, 'M-UNDERLAY'), 18);
    check('$0 receipt adds no value', valOf(afterZero, 'M-UNDERLAY'), 42000);

    // --- snapshot seeds the next replay --------------------------------
    var before = getStockPosition_();
    writeStockSnapshot_('2026-08', before.positions, currentStockSeq_(), who);

    var after = getStockPosition_();
    check('replay now starts from the snapshot', after.fromSnapshot, '2026-08');
    check('snapshot preserves qty', qtyOf(after, 'M-SHINGLE'), qtyOf(before, 'M-SHINGLE'));
    check('snapshot preserves value', valOf(after, 'M-SHINGLE'), valOf(before, 'M-SHINGLE'));
    // Includes the earlier $0 receipt (10 @ 42.00 + 8 @ 0.00 = 18 units,
    // still 42000 cents of value) -- the snapshot must round-trip that
    // mixed-cost layer set exactly, not just a single-layer material.
    check('snapshot preserves a $0-layered material', qtyOf(after, 'M-UNDERLAY'), 18);
    check('snapshot preserves its value', valOf(after, 'M-UNDERLAY'), 42000);

    // A move after the snapshot must layer on top of it, not replace it.
    appendStockMoves_([
      { materialId: 'M-UNDERLAY', materialName: 'Synthetic underlayment', moveType: 'ISSUE_JOB', qtyDelta: -4, jobRef: 'JOB-103', effectiveDate: d('2026-09-02') }
    ], who);
    var p4 = getStockPosition_();
    // Issuing 4 consumes from the $42.00 layer first (FIFO): 18 - 4 = 14 on
    // hand; value drops by 4 x 42.00 = 168.00 (16800 cents) from the paid
    // layer, the $0 layer untouched -- 42000 - 16800 = 25200.
    check('post-snapshot issue applies', qtyOf(p4, 'M-UNDERLAY'), 14);
    check('post-snapshot value correct', valOf(p4, 'M-UNDERLAY'), 25200);

  } catch (e) {
    failures++;
    Logger.log('FAIL  threw: ' + e.toString());
  } finally {
    dropStockTestSheets_();
    STOCK_SHEET_SUFFIX_ = '';
  }

  Logger.log(failures === 0 ? 'INTEGRATION: ALL PASS' : ('INTEGRATION: ' + failures + ' FAILED'));
  return failures;
}

/**
 * Leaves a realistic scenario sitting in the "(TEST)" sheets for hands-on
 * poking -- unlike the suite above, this does NOT clean up after itself, so
 * the sheets stay there to be read, sorted and argued with. Run
 * clearStockDemoData_() when finished. Production sheets are untouched.
 */
function seedStockDemoData_() {
  STOCK_SHEET_SUFFIX_ = ' (TEST)';
  try {
    dropStockTestSheets_();
    var d = function(str) { return new Date(str + 'T12:00:00'); };
    var who = 'demo@example.com';

    var cat = ensureSheetWithHeaders_(stockSheetName_(MATERIAL_CATALOG_SHEET), MATERIAL_CATALOG_HEADERS);
    cat.getRange(2, 1, 3, MATERIAL_CATALOG_HEADERS.length).setValues([
      ['M-SHINGLE',  'Architectural shingle',  'SQ', 'Roofing', 101, 'Roofing Materials', true, ''],
      ['M-UNDERLAY', 'Synthetic underlayment', 'RL', 'Roofing', 101, 'Roofing Materials', true, ''],
      ['M-RIDGE',    'Ridge vent 4ft',         'EA', 'Roofing', 101, 'Roofing Materials', true, '']
    ]);

    appendStockMoves_([
      { materialId: 'M-SHINGLE',  materialName: 'Architectural shingle',  moveType: 'RECEIPT_STOCK', qtyDelta: 40, unitCost: 95.50, effectiveDate: d('2026-08-03'), sourceDoc: 'PO 26-03-114' },
      { materialId: 'M-UNDERLAY', materialName: 'Synthetic underlayment', moveType: 'RECEIPT_STOCK', qtyDelta: 25, unitCost: 42.00, effectiveDate: d('2026-08-03'), sourceDoc: 'PO 26-03-114' },
      { materialId: 'M-RIDGE',    materialName: 'Ridge vent 4ft',         moveType: 'RECEIPT_STOCK', qtyDelta: 120, unitCost: 8.75, effectiveDate: d('2026-08-03'), sourceDoc: 'PO 26-03-114' }
    ], who);

    appendStockMoves_([
      { materialId: 'M-SHINGLE', materialName: 'Architectural shingle', moveType: 'ISSUE_JOB', qtyDelta: -18, jobRef: 'Hollis / Lot 12', effectiveDate: d('2026-08-06') },
      { materialId: 'M-RIDGE',   materialName: 'Ridge vent 4ft',        moveType: 'ISSUE_JOB', qtyDelta: -20, jobRef: 'Hollis / Lot 12', effectiveDate: d('2026-08-06') }
    ], who);

    // A second, pricier cost layer -- this is what makes the next issue
    // interesting, because it has to span both layers.
    appendStockMoves_([
      { materialId: 'M-SHINGLE', materialName: 'Architectural shingle', moveType: 'RECEIPT_STOCK', qtyDelta: 30, unitCost: 102.25, effectiveDate: d('2026-08-11'), sourceDoc: 'PO 26-03-121' }
    ], who);

    appendStockMoves_([
      { materialId: 'M-SHINGLE',  materialName: 'Architectural shingle',  moveType: 'ISSUE_JOB', qtyDelta: -30, jobRef: 'Weatherly / Lot 4', effectiveDate: d('2026-08-14') },
      { materialId: 'M-UNDERLAY', materialName: 'Synthetic underlayment', moveType: 'ISSUE_JOB', qtyDelta: -9,  jobRef: 'Weatherly / Lot 4', effectiveDate: d('2026-08-14') }
    ], who);

    appendStockMoves_([
      { materialId: 'M-SHINGLE', materialName: 'Architectural shingle', moveType: 'RETURN_PULLED', qtyDelta: 3, unitCost: 102.25, jobRef: 'Weatherly / Lot 4', effectiveDate: d('2026-08-19'), note: 'over-ordered' }
    ], who);

    var pos = getStockPosition_();
    Logger.log('Demo data written to the "(TEST)" sheets. Position:');
    Object.keys(pos.positions).sort().forEach(function(id) {
      var p = pos.positions[id];
      Logger.log('  ' + id + '  qty ' + p.qty + '  value $' + p.value.toFixed(2) + '  layers ' + JSON.stringify(p.layers));
    });
    if (pos.warnings.length) Logger.log('  warnings: ' + JSON.stringify(pos.warnings));
    return pos;
  } finally {
    STOCK_SHEET_SUFFIX_ = '';
  }
}

/** Removes the demo sheets seedStockDemoData_ left behind. */
function clearStockDemoData_() {
  dropStockTestSheets_();
  Logger.log('Demo sheets removed.');
}

/** Runs both suites. This is the one to run from the editor. */
function testStockAll_() {
  var pure = testStockLedger_();
  var integration = testStockLedgerIntegration_();
  Logger.log('TOTAL FAILURES: ' + (pure + integration));
  return pure + integration;
}
