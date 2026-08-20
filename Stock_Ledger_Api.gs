/**
 * Client-facing entry points for the Stock Ledger. Everything here is
 * reachable from the PWA via gasCall; the engine itself (Stock_Ledger.gs)
 * stays unreachable from the client so the only way in is through these
 * authorization and validation gates.
 *
 * Nothing here touches QuickBooks. Moves land in the ledger as 'pending'
 * and stay that way -- posting is a later, separate module, and the whole
 * point of that split is that the warehouse can be run and proven correct
 * before any accounting integration exists.
 *
 * TEST MODE: every function accepts payload.testMode. When set (admin or
 * owner only) the call runs against disposable "(TEST)" sheets instead of
 * the real ones, so the screens can be exercised with throwaway data in the
 * live app without touching production stock. The suffix is always restored
 * in a finally block.
 */

// --- Authorization -------------------------------------------------------

/** Read access: anyone who can see the warehouse. */
var STOCK_READ_ROLES  = ['admin', 'office', 'runner', 'site_manager'];
/** Move access: whoever physically handles material. */
var STOCK_WRITE_ROLES = ['admin', 'office', 'runner'];

/**
 * Admin-or-owner gate, for the operations that are financial corrections
 * rather than warehouse work: count adjustments, catalog edits, and test
 * mode. Owner emails are checked separately because an owner does not
 * necessarily carry the 'admin' role in the HR sheet.
 */
function authorizeStockAdmin_(payload) {
  var auth = authorizeCaller(payload, ['admin']);
  if (auth.ok) return auth;
  var email = verifySessionEmail_(payload && payload.sessionToken);
  if (email && isOwnerEmail(email)) return { ok: true, email: email, roles: ['owner'] };
  return auth;
}

/**
 * Runs fn against the "(TEST)" sheets when the caller asked for test mode
 * and is allowed it. Restores the suffix in a finally block so a throw can
 * never leave the whole script pointed at test sheets -- which would be a
 * quietly catastrophic way to lose a day of real warehouse activity.
 */
function withStockSheets_(payload, fn) {
  var wantsTest = !!(payload && payload.testMode);
  if (wantsTest && !authorizeStockAdmin_(payload).ok) {
    return { success: false, error: 'Test mode is limited to admin and owner.' };
  }
  var prev = STOCK_SHEET_SUFFIX_;
  STOCK_SHEET_SUFFIX_ = wantsTest ? ' (TEST)' : '';
  try {
    return fn();
  } finally {
    STOCK_SHEET_SUFFIX_ = prev;
  }
}

// --- Shared shaping ------------------------------------------------------

/** Client-safe date -> a Date the engine can order by. Defaults to today. */
function stockDateFrom_(raw) {
  if (!raw) return new Date();
  var d = raw instanceof Date ? raw : new Date(raw);
  return isNaN(d.getTime()) ? new Date() : d;
}

/**
 * Turns the engine's position map into the sorted array the screen renders,
 * merging in catalog metadata (unit, category) and flagging materials that
 * are in the catalog but have never moved -- those should still be listed,
 * at zero, so somebody can receive against them.
 */
function shapeStockMaterials_(positions, catalog) {
  var out = [];
  var seen = {};

  Object.keys(positions).forEach(function(materialId) {
    var pos = positions[materialId];
    var cat = catalog[materialId] || {};
    seen[materialId] = true;
    out.push({
      materialId:   materialId,
      materialName: pos.materialName || cat.materialName || materialId,
      unit:         cat.unit || '',
      category:     cat.category || '',
      onHand:       pos.qty,
      value:        pos.value,
      layers:       pos.layers,
      avgCost:      pos.qty ? Math.round((pos.value / pos.qty) * 10000) / 10000 : 0,
      inCatalog:    !!catalog[materialId]
    });
  });

  Object.keys(catalog).forEach(function(materialId) {
    if (seen[materialId] || !catalog[materialId].active) return;
    var cat = catalog[materialId];
    out.push({
      materialId: materialId, materialName: cat.materialName, unit: cat.unit,
      category: cat.category, onHand: 0, value: 0, layers: [], avgCost: 0, inCatalog: true
    });
  });

  out.sort(function(a, b) { return a.materialName.localeCompare(b.materialName); });
  return out;
}

// --- Read ----------------------------------------------------------------

/**
 * Everything the Stock Ledger screen needs in one call: current position
 * per material, the catalog, and the most recent moves. Deliberately a
 * single round trip -- the screen is opened constantly and this replaces a
 * QuickBooks round-trip per load with one local sheet read.
 */
function getStockLedgerView(payload) {
  try {
    var auth = authorizeCaller(payload, STOCK_READ_ROLES);
    if (!auth.ok) return { error: auth.error, code: auth.code, materials: [], recentMoves: [] };

    return withStockSheets_(payload, function() {
      var catalog = getMaterialCatalog_();
      var position = getStockPosition_();
      var moves = readStockLedgerMoves_(0);

      var recent = moves.slice().sort(function(a, b) { return b.seq - a.seq; }).slice(0, 60).map(function(m) {
        return {
          seq: m.seq, ledgerId: m.ledgerId,
          effectiveDate: m.effectiveDate instanceof Date ? m.effectiveDate.toISOString() : '',
          materialId: m.materialId, materialName: m.materialName,
          qtyDelta: m.qtyDelta, unitCost: m.unitCost, moveType: m.moveType,
          jobRef: m.jobRef, postStatus: m.postStatus
        };
      });

      var totalValue = 0;
      Object.keys(position.positions).forEach(function(id) { totalValue += position.positions[id].valueCents; });

      return {
        materials:    shapeStockMaterials_(position.positions, catalog),
        recentMoves:  recent,
        totalValue:   Math.round(totalValue) / 100,
        moveCount:    moves.length,
        fromSnapshot: position.fromSnapshot,
        warnings:     position.warnings,
        testMode:     !!(payload && payload.testMode)
      };
    });
  } catch (e) {
    return { error: e.toString(), materials: [], recentMoves: [] };
  }
}

/**
 * What an issue WOULD cost, without writing anything. Runs the same replay
 * the real issue will, then re-runs it with the proposed lines appended --
 * so the number shown is the number that will post, not an approximation
 * from a stored average.
 */
function previewStockIssue(payload) {
  try {
    var auth = authorizeCaller(payload, STOCK_WRITE_ROLES);
    if (!auth.ok) return { success: false, error: auth.error, code: auth.code };

    var lines = Array.isArray(payload.lines) ? payload.lines : [];
    if (!lines.length) return { success: false, error: 'Add at least one material.' };

    return withStockSheets_(payload, function() {
      var position = getStockPosition_();
      var onHand = {};
      Object.keys(position.positions).forEach(function(id) { onHand[id] = position.positions[id].qty; });

      // Layers as they stand now become the seed, so the hypothetical replay
      // costs exactly the layers the real one would consume.
      var seed = {};
      Object.keys(position.positions).forEach(function(id) { seed[id] = position.positions[id].layers; });

      var effectiveDate = stockDateFrom_(payload.effectiveDate);
      var hypothetical = [];
      var results = [];
      for (var i = 0; i < lines.length; i++) {
        var qty = Math.abs(parseFloat(lines[i].qty) || 0);
        var materialId = (lines[i].materialId || '').toString().trim();
        if (!materialId || !qty) return { success: false, error: 'Every line needs a material and a quantity.' };
        var available = parseFloat(onHand[materialId]) || 0;
        if (qty > available) {
          var nm = position.positions[materialId] ? position.positions[materialId].materialName : materialId;
          return { success: false, error: 'Only ' + available + ' of "' + nm + '" on hand -- cannot pull ' + qty + '.' };
        }
        onHand[materialId] = available - qty;
        hypothetical.push({
          seq: i + 1, ledgerId: 'PREVIEW-' + i, effectiveDate: effectiveDate,
          materialId: materialId, qtyDelta: -qty, unitCost: null, moveType: 'ISSUE_JOB'
        });
      }

      var replay = replayStockFifo_(hypothetical, seed);
      var grandCents = 0;
      hypothetical.forEach(function(h, i) {
        var cost = replay.issueCosts[h.ledgerId] || { cogsCents: 0, unitCost: 0 };
        grandCents += cost.cogsCents;
        var pos = position.positions[h.materialId];
        results.push({
          materialId: h.materialId,
          materialName: pos ? pos.materialName : h.materialId,
          qty: -h.qtyDelta,
          unitCost: cost.unitCost,
          lineTotal: Math.round(cost.cogsCents) / 100,
          onHandAfter: (pos ? pos.qty : 0) - (-h.qtyDelta)
        });
      });

      return { success: true, lines: results, grandTotal: Math.round(grandCents) / 100 };
    });
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// --- Write ---------------------------------------------------------------

/** Vendor -> warehouse. payload: {lines:[{materialId, qty, unitCost}], sourceDoc, effectiveDate, note}. */
function receiveStockMaterial(payload) {
  try {
    var auth = authorizeCaller(payload, STOCK_WRITE_ROLES);
    if (!auth.ok) return { success: false, error: auth.error, code: auth.code };

    var lines = Array.isArray(payload.lines) ? payload.lines : [];
    if (!lines.length) return { success: false, error: 'Add at least one material.' };

    return withStockSheets_(payload, function() {
      var catalog = getMaterialCatalog_();
      var effectiveDate = stockDateFrom_(payload.effectiveDate);
      var moves = [];
      for (var i = 0; i < lines.length; i++) {
        var materialId = (lines[i].materialId || '').toString().trim();
        var cat = catalog[materialId];
        if (!cat) return { success: false, error: 'Material "' + materialId + '" is not in the catalog.' };
        moves.push({
          materialId:    materialId,
          materialName:  cat.materialName,
          moveType:      'RECEIPT_STOCK',
          qtyDelta:      Math.abs(parseFloat(lines[i].qty) || 0),
          unitCost:      parseFloat(lines[i].unitCost),
          effectiveDate: effectiveDate,
          sourceDoc:     (payload.sourceDoc || '').toString().trim(),
          note:          (payload.note || '').toString().trim()
        });
      }
      return appendStockMoves_(moves, auth.email);
    });
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/** Warehouse -> job. payload: {lines:[{materialId, qty}], jobRef, effectiveDate, note}. */
function issueStockToJob(payload) {
  try {
    var auth = authorizeCaller(payload, STOCK_WRITE_ROLES);
    if (!auth.ok) return { success: false, error: auth.error, code: auth.code };

    var jobRef = (payload.jobRef || '').toString().trim();
    if (!jobRef) return { success: false, error: 'Job reference is required.' };
    var lines = Array.isArray(payload.lines) ? payload.lines : [];
    if (!lines.length) return { success: false, error: 'Add at least one material.' };

    return withStockSheets_(payload, function() {
      var catalog = getMaterialCatalog_();
      var effectiveDate = stockDateFrom_(payload.effectiveDate);
      var moves = [];
      for (var i = 0; i < lines.length; i++) {
        var materialId = (lines[i].materialId || '').toString().trim();
        var cat = catalog[materialId] || {};
        moves.push({
          materialId:    materialId,
          materialName:  cat.materialName || materialId,
          moveType:      'ISSUE_JOB',
          // Negated here, not by the client -- an issue is always a
          // reduction and the sign should not be something a caller can get
          // wrong or invert.
          qtyDelta:      -Math.abs(parseFloat(lines[i].qty) || 0),
          jobRef:        jobRef,
          effectiveDate: effectiveDate,
          note:          (payload.note || '').toString().trim()
        });
      }
      return appendStockMoves_(moves, auth.email);
    });
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Job -> warehouse. Two flavours, because they are genuinely different
 * events: RETURN_PULLED puts back stock we issued (priced at what it was
 * issued at, so the round trip nets to zero) and RETURN_DIRECT capitalizes
 * material that was billed straight to the job and never was stock.
 * Either way the caller supplies the cost -- the UI is responsible for
 * offering the right one.
 * payload: {lines:[{materialId, qty, unitCost, kind}], jobRef, effectiveDate, note}.
 */
function returnStockToWarehouse(payload) {
  try {
    var auth = authorizeCaller(payload, STOCK_WRITE_ROLES);
    if (!auth.ok) return { success: false, error: auth.error, code: auth.code };

    var jobRef = (payload.jobRef || '').toString().trim();
    if (!jobRef) return { success: false, error: 'Job reference is required.' };
    var lines = Array.isArray(payload.lines) ? payload.lines : [];
    if (!lines.length) return { success: false, error: 'Add at least one material.' };

    return withStockSheets_(payload, function() {
      var catalog = getMaterialCatalog_();
      var effectiveDate = stockDateFrom_(payload.effectiveDate);
      var moves = [];
      for (var i = 0; i < lines.length; i++) {
        var materialId = (lines[i].materialId || '').toString().trim();
        var cat = catalog[materialId];
        if (!cat) return { success: false, error: 'Material "' + materialId + '" is not in the catalog.' };
        moves.push({
          materialId:    materialId,
          materialName:  cat.materialName,
          moveType:      lines[i].kind === 'direct' ? 'RETURN_DIRECT' : 'RETURN_PULLED',
          qtyDelta:      Math.abs(parseFloat(lines[i].qty) || 0),
          unitCost:      parseFloat(lines[i].unitCost),
          jobRef:        jobRef,
          reverses:      (lines[i].reverses || '').toString().trim(),
          effectiveDate: effectiveDate,
          note:          (payload.note || '').toString().trim()
        });
      }
      return appendStockMoves_(moves, auth.email);
    });
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Count correction. Admin/owner only -- this is the one move that changes
 * stock without a physical event behind it, so it is a financial decision
 * rather than warehouse work. Never called automatically: every adjustment
 * has a person's name on it, per the locked decision.
 * payload: {materialId, countedQty, unitCost, effectiveDate, note}.
 */
function adjustStockCount(payload) {
  try {
    var auth = authorizeStockAdmin_(payload);
    if (!auth.ok) return { success: false, error: auth.error, code: auth.code };

    var materialId = (payload.materialId || '').toString().trim();
    if (!materialId) return { success: false, error: 'Pick a material.' };
    var counted = parseFloat(payload.countedQty);
    if (isNaN(counted)) return { success: false, error: 'Enter the counted quantity.' };
    var note = (payload.note || '').toString().trim();
    if (!note) return { success: false, error: 'A reason is required for a count adjustment.' };

    return withStockSheets_(payload, function() {
      var catalog = getMaterialCatalog_();
      var cat = catalog[materialId];
      if (!cat) return { success: false, error: 'Material "' + materialId + '" is not in the catalog.' };

      var position = getStockPosition_();
      var pos = position.positions[materialId];
      var current = pos ? pos.qty : 0;
      var delta = counted - current;
      if (!delta) return { success: false, error: 'Counted quantity already matches on hand (' + current + ') -- nothing to adjust.' };

      var move = {
        materialId:    materialId,
        materialName:  cat.materialName,
        moveType:      'COUNT_ADJUST',
        qtyDelta:      delta,
        effectiveDate: stockDateFrom_(payload.effectiveDate),
        note:          note
      };
      // Only an increase needs a price -- a decrease consumes existing
      // layers, and supplying a cost there would override FIFO.
      if (delta > 0) move.unitCost = parseFloat(payload.unitCost);

      var res = appendStockMoves_([move], auth.email);
      if (res.success) { res.from = current; res.to = counted; res.delta = delta; }
      return res;
    });
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// --- Catalog -------------------------------------------------------------

/**
 * Adds or updates one material. Admin/owner only -- the catalog is the
 * ledger's vocabulary, and a duplicate or renamed id would fragment a
 * material's history across two rows that no longer add up.
 * payload: {materialId, materialName, unit, category, qboPostItemId, qboPostItemName, active}.
 */
function saveMaterialCatalogItem(payload) {
  try {
    var auth = authorizeStockAdmin_(payload);
    if (!auth.ok) return { success: false, error: auth.error, code: auth.code };

    var materialId = (payload.materialId || '').toString().trim().toUpperCase();
    var materialName = (payload.materialName || '').toString().trim();
    if (!materialId) return { success: false, error: 'Material Id is required.' };
    if (!materialName) return { success: false, error: 'Material name is required.' };
    if (!/^[A-Z0-9][A-Z0-9_-]*$/.test(materialId)) {
      return { success: false, error: 'Material Id can only contain letters, numbers, hyphens and underscores.' };
    }

    var lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) return { success: false, error: 'Server is busy - try again in a moment.' };
    try {
      return withStockSheets_(payload, function() {
        var sheet = ensureSheetWithHeaders_(stockSheetName_(MATERIAL_CATALOG_SHEET), MATERIAL_CATALOG_HEADERS);
        var row = [];
        for (var c = 0; c < MATERIAL_CATALOG_HEADERS.length; c++) row.push('');
        row[MATERIAL_CATALOG_COL['Material Id']]        = materialId;
        row[MATERIAL_CATALOG_COL['Material Name']]      = materialName;
        row[MATERIAL_CATALOG_COL['Unit']]               = (payload.unit || '').toString().trim();
        row[MATERIAL_CATALOG_COL['Category']]           = (payload.category || '').toString().trim();
        row[MATERIAL_CATALOG_COL['QBO Post Item Id']]   = (payload.qboPostItemId || '').toString().trim();
        row[MATERIAL_CATALOG_COL['QBO Post Item Name']] = (payload.qboPostItemName || '').toString().trim();
        row[MATERIAL_CATALOG_COL['Active']]             = payload.active === false ? false : true;
        row[MATERIAL_CATALOG_COL['Notes']]              = (payload.notes || '').toString().trim();

        var lastRow = sheet.getLastRow();
        if (lastRow >= 2) {
          var ids = sheet.getRange(2, MATERIAL_CATALOG_COL['Material Id'] + 1, lastRow - 1, 1).getValues();
          for (var i = 0; i < ids.length; i++) {
            if ((ids[i][0] || '').toString().trim().toUpperCase() === materialId) {
              sheet.getRange(i + 2, 1, 1, MATERIAL_CATALOG_HEADERS.length).setValues([row]);
              return { success: true, materialId: materialId, updated: true };
            }
          }
        }
        sheet.getRange(sheet.getLastRow() + 1, 1, 1, MATERIAL_CATALOG_HEADERS.length).setValues([row]);
        return { success: true, materialId: materialId, updated: false };
      });
    } finally {
      lock.releaseLock();
    }
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Wipes the "(TEST)" sheets so a testing session can start clean. Refuses
 * outright unless testMode is set -- there is deliberately no code path
 * here that can touch a production sheet.
 */
function resetStockTestData(payload) {
  try {
    var auth = authorizeStockAdmin_(payload);
    if (!auth.ok) return { success: false, error: auth.error, code: auth.code };
    if (!(payload && payload.testMode)) {
      return { success: false, error: 'This only clears test data, and test mode is not on.' };
    }
    dropStockTestSheets_();
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}
