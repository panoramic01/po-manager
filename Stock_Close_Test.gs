/**
 * Month-end close tests. Run testStockCloseAll_() from the Apps Script
 * editor. Same disposable "(TEST)" sheet pattern as Stock_Ledger_Test.gs --
 * production sheets are never opened.
 */

function stockCloseTestSheetNames_() {
  return stockTestSheetNames_().concat(
    [PHYSICAL_COUNT_SHEET, PERIOD_CLOSE_SHEET].map(function(base) { return base + ' (TEST)'; })
  );
}

function dropStockCloseTestSheets_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  stockCloseTestSheetNames_().forEach(function(name) {
    var sh = ss.getSheetByName(name);
    if (sh) ss.deleteSheet(sh);
  });
}

/** Pure: the tolerance rule itself, no sheets involved. */
function testStockCloseTolerance_() {
  var failures = 0;
  var check = function(name, actual, expected) {
    var ok = actual === expected;
    if (!ok) failures++;
    Logger.log((ok ? 'PASS  ' : 'FAIL  ') + name + (ok ? '' : '  expected ' + expected + ' / actual ' + actual));
  };

  // $25 floor: a small material where 2% would be tiny stays flagged past $25.
  check('flags a variance just over the $25 floor',
    isStockVarianceFlagged_(1, 30, 100), true); // 1 x $30 = $30 > max(25, 100*.02=2)
  check('does not flag a variance under the $25 floor',
    isStockVarianceFlagged_(1, 20, 100), false); // $20 <= max(25, 2)

  // 2% ceiling: on a large-value material, 2% exceeds the $25 floor.
  check('flags a variance over 2% on a high-value material',
    isStockVarianceFlagged_(2, 100, 5000), true); // $200 > max(25, 5000*.02=100)
  check('does not flag a variance under 2% on a high-value material',
    isStockVarianceFlagged_(1, 50, 5000), false); // $50 <= max(25, 100)

  check('zero variance is never flagged', isStockVarianceFlagged_(0, 50, 5000), false);
  check('unpriced material flags ANY nonzero variance -- no dollar basis to call it immaterial',
    isStockVarianceFlagged_(3, 0, 0), true);

  Logger.log(failures === 0 ? 'TOLERANCE: ALL PASS' : ('TOLERANCE: ' + failures + ' FAILED'));
  return failures;
}

/** Real sheets, disposable, drives the full count -> review -> close -> report flow. */
function testStockCloseIntegration_() {
  var failures = 0;
  var check = function(name, actual, expected) {
    var a = JSON.stringify(actual), e = JSON.stringify(expected);
    var ok = a === e;
    if (!ok) failures++;
    Logger.log((ok ? 'PASS  ' : 'FAIL  ') + name + (ok ? '' : '  expected ' + e + ' / actual ' + a));
  };
  var d = function(str) { return new Date(str + 'T12:00:00'); };
  var who = 'closer@example.com';

  STOCK_SHEET_SUFFIX_ = ' (TEST)';
  try {
    dropStockCloseTestSheets_();

    var cat = ensureSheetWithHeaders_(stockSheetName_(MATERIAL_CATALOG_SHEET), MATERIAL_CATALOG_HEADERS);
    cat.getRange(2, 1, 3, MATERIAL_CATALOG_HEADERS.length).setValues([
      ['M-SHINGLE',  'Architectural shingle',  'SQ', 'Roofing', '', '', true, ''],
      ['M-UNDERLAY', 'Synthetic underlayment', 'RL', 'Roofing', '', '', true, ''],
      ['M-RIDGE',    'Ridge vent 4ft',         'EA', 'Roofing', '', '', true, '']
    ]);

    appendStockMoves_([
      { materialId: 'M-SHINGLE',  materialName: 'Architectural shingle',  moveType: 'RECEIPT_STOCK', qtyDelta: 40, unitCost: 95.50, effectiveDate: d('2026-08-01'), sourceDoc: 'PO-1' },
      { materialId: 'M-UNDERLAY', materialName: 'Synthetic underlayment', moveType: 'RECEIPT_STOCK', qtyDelta: 25, unitCost: 42.00, effectiveDate: d('2026-08-01'), sourceDoc: 'PO-1' },
      { materialId: 'M-RIDGE',    materialName: 'Ridge vent 4ft',         moveType: 'RECEIPT_STOCK', qtyDelta: 100, unitCost: 8.75, effectiveDate: d('2026-08-01'), sourceDoc: 'PO-1' }
    ], who);

    // --- count sheet reflects catalog + current on-hand, none counted yet ---
    var sheetView = computeStockCountSheet_();
    check('count sheet lists all 3 active materials', sheetView.materials.length, 3);
    check('nothing counted yet', sheetView.countedSoFar, 0);
    var period = sheetView.period;

    // --- review before any count: not ready ---
    var reviewBefore = computeStockCloseReview_();
    check('not ready before any count', reviewBefore.readyToClose, false);
    check('all 3 missing before any count', reviewBefore.missingMaterials.length, 3);

    // --- partial count: still not ready ---
    var partial = writePhysicalCount_([
      { materialId: 'M-SHINGLE', countedQty: 40 }
    ], who);
    check('partial count accepted', partial.success, true);
    var reviewPartial = computeStockCloseReview_();
    check('still not ready with a partial count', reviewPartial.readyToClose, false);
    check('2 remain missing', reviewPartial.missingMaterials.length, 2);

    // --- confirm refuses on an incomplete count ---
    var closeIncomplete = writeStockClose_(null, who);
    check('close refuses an incomplete count', closeIncomplete.success, false);

    // --- complete the count: shingle exact, underlay short by 2 (flagged,
    // 2 x $42 = $84 > $25), ridge exact ---
    var full = writePhysicalCount_([
      { materialId: 'M-SHINGLE',  countedQty: 40 },
      { materialId: 'M-UNDERLAY', countedQty: 23 },
      { materialId: 'M-RIDGE',    countedQty: 100 }
    ], who);
    check('full count accepted', full.success, true);

    var reviewFull = computeStockCloseReview_();
    check('ready once every material is counted', reviewFull.readyToClose, true);
    check('one material flagged', reviewFull.materialsFlagged, 1);
    check('flagged one is the underlay', reviewFull.lines.filter(function(l){return l.flagged;})[0].materialId, 'M-UNDERLAY');
    check('variance total is 84', reviewFull.varianceTotalValue, 84);
    // 40*95.50 + 25*42.00 + 100*8.75 = 3820 + 1050 + 875 = 5745
    check('ledger value totals correctly', reviewFull.ledgerValue, 5745);

    // --- confirm refuses a flagged variance with no note ---
    var closeNoNote = writeStockClose_(null, who);
    check('close refuses a flagged variance with no note', closeNoNote.success, false);

    // --- confirm succeeds with a note ---
    var closeOk = writeStockClose_('underlay short 2 -- shrinkage, not adjusting yet', who);
    check('close succeeds once noted', closeOk.success, true);
    check('closes the right period', closeOk.period, period);

    // --- period is now closed: with no next-period activity yet, review
    // correctly has nothing to show -- determineCloseablePeriod_ skips
    // already-closed periods when picking a target, so there is no
    // "already closed" period left to review, only "nothing yet."
    var reviewClosed = computeStockCloseReview_();
    check('nothing left to review once the only period is closed', !!reviewClosed.error, true);
    var recount = writePhysicalCount_([{ materialId: 'M-SHINGLE', countedQty: 41 }], who);
    check('cannot recount once nothing is open to count against', recount.success, false);

    // --- report shows the close ---
    var history = readStockCloseHistory_();
    check('history has one closed period', history.closes.length, 1);
    check('history period matches', history.closes[0].period, period);
    check('history flagged count matches', history.closes[0].materialsFlagged, 1);

    var detail = readStockPeriodDetail_(period);
    check('detail has all 3 materials', detail.materials.length, 3);
    var underlayDetail = detail.materials.filter(function(m){ return m.materialId === 'M-UNDERLAY'; })[0];
    // Snapshot reflects the LEDGER's 25 on hand, not the counted 23 -- the
    // close records what happened, it never silently rewrites the ledger.
    check('snapshot preserves the ledger quantity, not the physical count', underlayDetail.qty, 25);

  } catch (e) {
    failures++;
    Logger.log('FAIL  threw: ' + e.toString());
  } finally {
    dropStockCloseTestSheets_();
    STOCK_SHEET_SUFFIX_ = '';
  }

  Logger.log(failures === 0 ? 'CLOSE INTEGRATION: ALL PASS' : ('CLOSE INTEGRATION: ' + failures + ' FAILED'));
  return failures;
}

function testStockCloseAll_() {
  var a = testStockCloseTolerance_();
  var b = testStockCloseIntegration_();
  Logger.log('CLOSE TOTAL FAILURES: ' + (a + b));
  return a + b;
}
