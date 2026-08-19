/**
 * PO Manager Web App - Panoramic Building
 * =========================================
 * Paste this into your Google Apps Script project (Extensions > Apps Script).
 * Also paste the contents of PO_Manager_index.html into a new HTML file named "index".
 * Then deploy: Deploy > New Deployment > Web App.
 */

var SHEET_NAME  = "PO Database";
var ROLES_SHEET = "HR";

var GOOGLE_CLIENT_ID = '740908602873-3k73e1sscs32ohhbtoc4ha8hdpvp05t9.apps.googleusercontent.com';
var GOOGLE_HD_DOMAIN = 'panoramicbuildingllc.com';

// Owner accounts always resolve to admin and can never be demoted or removed
// through the app, regardless of what the HR sheet says.
var OWNER_EMAILS = ['aidan@panoramicbuildingllc.com', 'aidansalisbury213@gmail.com'];

function isOwnerEmail(email) {
  email = (email || '').toString().toLowerCase().trim();
  for (var i = 0; i < OWNER_EMAILS.length; i++) {
    if (OWNER_EMAILS[i].toLowerCase() === email) return true;
  }
  return false;
}

// 'aidan' is the owner-only role label; it carries the exact same permissions
// as 'admin' everywhere in the app. 'purchaser' is the old token for what's
// now called 'office' (business-side rename) -- kept as an alias here so
// existing HR sheet rows that still literally say "purchaser" keep working
// without a manual data migration. Every role comparison should normalize
// through this first so allow-lists never need to spell out either legacy
// token.
function normalizeRole_(role) {
  if (role === 'aidan') return 'admin';
  if (role === 'purchaser') return 'office';
  return role;
}

// ── Multi-role support ───────────────────────────────────────────────────────
// An employee's HR sheet role cell (column D) can hold one role ("admin") or
// several comma-separated roles ("admin,human_resources"). Every place that
// used to compare a single role string now works with these list helpers so
// an account can carry multiple roles and gets the union of their permissions.

/** Parses a raw role cell into a deduped, lowercase, trimmed array of role tokens. */
function parseRoleList_(raw) {
  var seen = {};
  var out = [];
  (raw || '').toString().split(',').forEach(function(tok) {
    var r = tok.toLowerCase().trim();
    if (!r || seen[r]) return;
    seen[r] = true;
    out.push(r);
  });
  return out;
}

/** Normalizes each role in a list (via normalizeRole_) and dedupes the result. */
function normalizeRoleList_(rawList) {
  var seen = {};
  var out = [];
  rawList.forEach(function(r) {
    var n = normalizeRole_(r);
    if (!seen[n]) { seen[n] = true; out.push(n); }
  });
  return out;
}

/** True if any role in effRoles (already normalized) appears in the allowed list. */
function hasAnyRole_(effRoles, allowed) {
  for (var i = 0; i < allowed.length; i++) {
    if (effRoles.indexOf(allowed[i]) !== -1) return true;
  }
  return false;
}

// The only role tokens the app understands. Anything else in a client-supplied
// role list is dropped rather than written to the HR sheet - keeps garbage or
// script-like text out of a cell that gets echoed back into the UI verbatim.
var VALID_EMPLOYEE_ROLES = ['runner', 'site_manager', 'office', 'human_resources', 'admin', 'aidan'];
function filterValidRoles_(roleList) {
  return roleList.filter(function(r) { return VALID_EMPLOYEE_ROLES.indexOf(r) !== -1; });
}

var STATUS_OPTIONS = [
  "Pending Pickup",
  "Pending Delivery",
  "Pending Delivery to Supplier",
  "Ordered",
  "Being made",
  "Currently Picking Up",
  "Delivered",
  "Ready to Reconcile",
  "Invoice Missing",
  "Needs Review",
  "Complete",
  "Draft",
  "Canceled",
  "Other"
];

// Rebuilt from actual PO Database usage (debugListActualVendors_ tally, 2026-08)
// rather than a hand-maintained guess -- "Plaster" never appeared even once
// in real usage and has been dropped; everything below appeared 2+ times,
// with near-duplicate free-typed variants folded into one canonical name
// (e.g. "LKL"/"Lkl" -> "LKL Associates", "Cmg"/"Coated metals group" ->
// "Coated Metals Group"). One-off vendors (Amazon, Lowes, Home Depot,
// Grainger, etc., each used exactly once) are left to the "Other" free-text
// fallback rather than cluttering the dropdown.
var VENDOR_OPTIONS = [
  "ABC Interiors",
  "Timberline",
  "Castalite",
  "Lansing",
  "Harristone",
  "Interstate Brick",
  "Leak Tech",
  "LKL Associates",
  "Transcending",
  "Rustic",
  "Alside",
  "Coated Metals Group",
  "Swanson",
  "Burton Lumber",
  "Metal Super Markets",
  "Concrete Color and Design",
  "Builders First Source",
  "Architectural Depot",
  "Beehive",
  "Linc",
  "Stone Connections",
  "Artistic Stone",
  "Rocky Mountain Supply",
  "Tresselwood",
  "Other"
];

// ─── Web App Entry Point ─────────────────────────────────────────────────────

function doGet() {
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("Panoramic Ops")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * REST-style POST handler - replaces google.script.run for all client calls.
 * Expects JSON body: { action: string, payload: object }
 * Returns JSON via ContentService.
 */
function doPost(e) {
  try {
    var body    = JSON.parse(e.postData.contents);
    var action  = body.action;
    var payload = body.payload || {};
    var result;

    if      (action === 'getConfig')        result = getConfig(payload);
    else if (action === 'verifyLogin')       result = verifyLogin(payload.email, payload.password);
    else if (action === 'verifyGoogleLogin') result = verifyGoogleLogin(payload.credential);
    else if (action === 'getSheetData')      result = getSheetData(payload);
    else if (action === 'createPO')          result = createPO(payload);
    else if (action === 'createSubPO')       result = createSubPO(payload);
    else if (action === 'updatePO')          result = updatePO(payload);
    else if (action === 'findPOByNumber')    result = findPOByNumber(payload);
    else if (action === 'savePhotoToDrive')  result = savePhotoToDrive(payload);
    else if (action === 'createProject')       result = createProjectAndTask(payload);
    else if (action === 'saveFileToFolderById') result = saveFileToFolderById(payload);
    else if (action === 'setProjectDriveLink') result = setProjectDriveLink(payload);
    else if (action === 'getPricingData')    result = getPricingData(payload);
    else if (action === 'updatePricing')     result = updatePricing(payload);
    else if (action === 'getContacts')         result = getContacts(payload);
    else if (action === 'updateContact')       result = updateContact(payload);
    else if (action === 'addContact')          result = addContact(payload);
    else if (action === 'deleteContact')       result = deleteContact(payload);
    else if (action === 'reconcileStatement')  result = reconcileStatement(payload);
    else if (action === 'getJobList')          result = getJobList();
    else if (action === 'getJobCostSummary')   result = getJobCostSummary(payload);
    else if (action === 'getMissingInvoices')  result = getMissingInvoices(payload);
    else if (action === 'getJobDashboard')     result = getJobDashboard(payload);
    else if (action === 'getVendorSpend')      result = getVendorSpend(payload);
    else if (action === 'categorizeInvoices')  result = categorizeInvoices(payload);
    else if (action === 'suggestCategories')   result = suggestCategories(payload);
    else if (action === 'processEstimateWithMatching') result = processEstimateWithMatching(payload);
    else if (action === 'getSopData')                  result = getSopData();
    else if (action === 'saveMaterialHistory')          result = saveMaterialHistory(payload);
    else if (action === 'getAsanaJobs')                result = getAsanaJobs();
    else if (action === 'getJobsByPhase')               result = getJobsByPhase(payload);
    else if (action === 'getRecentQualityWalks')        result = getRecentQualityWalks(payload);
    else if (action === 'getMyAsanaTasks')              result = getMyAsanaTasks(payload);
    else if (action === 'submitQualityCheck')           result = submitQualityCheck(payload);
    else if (action === 'getQualityWalkPhotos')         result = getQualityWalkPhotos(payload);
    else if (action === 'submitOfficeNote')             result = submitOfficeNote(payload);
    else if (action === 'getAssignableEmployees')       result = getAssignableEmployees(payload);
    else if (action === 'saveOfficeNotePhoto')          result = saveOfficeNotePhoto(payload);
    else if (action === 'saveMileageCommissionPdf')     result = saveMileageCommissionPdf(payload);
    else if (action === 'getPTOData')                  result = getPTOData(payload);
    else if (action === 'submitPTORequest')             result = submitPTORequest(payload);
    else if (action === 'getPTOQueue')                  result = getPTOQueue(payload);
    else if (action === 'approvePTO')                   result = approvePTO(payload);
    else if (action === 'denyPTO')                      result = denyPTO(payload);
    else if (action === 'cancelPTORequest')             result = cancelPTORequest(payload);
    else if (action === 'clockIn')                      result = clockIn(payload);
    else if (action === 'clockOut')                     result = clockOut(payload);
    else if (action === 'getClockStatus')               result = getClockStatus(payload);
    else if (action === 'getTimesheet')                 result = getTimesheet(payload);
    else if (action === 'updateProfile')               result = updateProfile(payload);
    else if (action === 'getEmployees')                result = getEmployees(payload);
    else if (action === 'addEmployee')                 result = addEmployee(payload);
    else if (action === 'updateEmployee')              result = updateEmployee(payload);
    else if (action === 'removeEmployee')              result = removeEmployee(payload);
    else if (action === 'getPTOOverview')              result = getPTOOverview(payload);
    else if (action === 'getPayrollSummary')           result = getPayrollSummary(payload);
    else if (action === 'emailPayroll')                result = emailPayroll(payload);
    else if (action === 'getPayrollPdf')               result = getPayrollPdf(payload);
    else if (action === 'approveTimesheet')            result = approveTimesheet(payload);
    else if (action === 'unapproveTimesheet')          result = unapproveTimesheet(payload);
    else if (action === 'approveMyTimesheet')          result = approveMyTimesheet(payload);
    else if (action === 'getMyPeriodDetail')           result = getMyPeriodDetail(payload);
    else if (action === 'getShiftForDate')               result = getShiftForDate(payload);
    else if (action === 'submitTimeCorrection')         result = submitTimeCorrection(payload);
    else if (action === 'submitTimeCorrectionsBatch')   result = submitTimeCorrectionsBatch(payload);
    else if (action === 'getMyTimeCorrections')         result = getMyTimeCorrections(payload);
    else if (action === 'getTimeCorrectionQueue')       result = getTimeCorrectionQueue(payload);
    else if (action === 'approveTimeCorrection')        result = approveTimeCorrection(payload);
    else if (action === 'denyTimeCorrection')           result = denyTimeCorrection(payload);
    else if (action === 'adminSetTimeEntry')            result = adminSetTimeEntry(payload);
    else if (action === 'adminDeleteTimeEntry')         result = adminDeleteTimeEntry(payload);
    else if (action === 'getInventory')                result = getInventory(payload);
    else if (action === 'addAsset')                    result = addAsset(payload);
    else if (action === 'updateAsset')                 result = updateAsset(payload);
    else if (action === 'deleteAsset')                 result = deleteAsset(payload);
    else if (action === 'getAssetMaintenanceLog')      result = getAssetMaintenanceLog(payload);
    else if (action === 'addMaintenanceLog')           result = addMaintenanceLog(payload);
    else if (action === 'registerPushToken')           result = registerPushToken(payload);
    else if (action === 'unregisterPushToken')         result = unregisterPushToken(payload);
    else if (action === 'getMaterialCatalog')          result = getMaterialCatalog(payload);
    else if (action === 'getMaterialInventory')        result = getMaterialInventory(payload);
    else if (action === 'logMaterialTransaction')      result = logMaterialTransaction(payload);
    else if (action === 'deleteMaterialLogEntry')      result = deleteMaterialLogEntry(payload);
    else if (action === 'getInternalInvoices')         result = getInternalInvoices(payload);
    else if (action === 'getInternalInvoicePdf')       result = getInternalInvoicePdf(payload);
    else if (action === 'getPortalCredentials')        result = getPortalCredentials(payload);
    else if (action === 'savePortalCredential')        result = savePortalCredential(payload);
    else if (action === 'updatePortalCredential')      result = updatePortalCredential(payload);
    else if (action === 'deletePortalCredential')      result = deletePortalCredential(payload);
    else if (action === 'getQuickBooksAuthUrl')        result = getQuickBooksAuthorizationUrl(payload);
    else if (action === 'getQuickBooksStatus')         result = getQuickBooksStatus(payload);
    else if (action === 'disconnectQuickBooks')        result = disconnectQuickBooks(payload);
    else if (action === 'testQuickBooksConnection')    result = testQuickBooksConnection(payload);
    else if (action === 'testQuickBooksVendors')       result = testQuickBooksVendors(payload);
    else if (action === 'testQuickBooksCustomers')     result = testQuickBooksCustomers(payload);
    else if (action === 'testQuickBooksAllJobs')       result = testQuickBooksAllJobs(payload);
    else if (action === 'getQuickBooksVendorMap')      result = getQuickBooksVendorMap(payload);
    else if (action === 'saveQuickBooksVendorMapping') result = saveQuickBooksVendorMapping(payload);
    else if (action === 'extractInvoiceLineItems')     result = extractInvoiceLineItems(payload);
    else if (action === 'getInvoiceStaging')           result = getInvoiceStaging(payload);
    else if (action === 'saveInvoiceStagingReview')    result = saveInvoiceStagingReview(payload);
    else if (action === 'getQuickBooksItemCatalogForReview') result = getQuickBooksItemCatalogForReview(payload);
    else if (action === 'refreshQuickBooksItemCatalog') result = refreshQuickBooksItemCatalog(payload);
    else if (action === 'matchInvoiceLineItems')       result = matchInvoiceLineItems(payload);
    else if (action === 'createQuickBooksBill')        result = createQuickBooksBill(payload);
    else if (action === 'getQuickBooksAccountsForNewItem') result = getQuickBooksAccountsForNewItem(payload);
    else if (action === 'createQuickBooksItem')        result = createQuickBooksItem(payload);
    else if (action === 'getQuickBooksLinkStatusForInvoice') result = getQuickBooksLinkStatusForInvoice(payload);
    else if (action === 'saveProjectQuickBooksCustomerId') result = saveProjectQuickBooksCustomerId(payload);
    else                                        result = { error: 'Unknown action: ' + action };

    if (result && result.success === false) {
      logError_(action, result.error, payload);
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    try { logError_(action || 'doPost:parse', err.toString(), payload || null); } catch (e2) {}
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ─── Data Access ─────────────────────────────────────────────────────────────

/**
 * Returns all valid PO rows from the sheet as an array of objects.
 * Rows without a valid PO number (YY-QQ-###) are skipped automatically,
 * so the input/header rows at the top of the sheet are ignored.
 *
 * Invoice fields (invoiceTotal, invoiceFile, invoiceLink) are only
 * populated for admin/office callers -- this mirrors the client's
 * canViewInvoice gate, but enforced here too so the raw response can't be
 * used to read invoice data for a role the UI hides it from. An
 * unresolvable caller (missing/unknown email) is treated as the lowest
 * privilege and gets the fields stripped, same as any other role.
 */
function getSheetData(payload) {
  var callerEmail = verifySessionEmail_(payload && payload.sessionToken) || '';
  var callerRoles = getRoleByEmail(callerEmail).effRoles;
  var canViewInvoice = hasAnyRole_(callerRoles, ['admin', 'office']);

  var sheet = getSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var numRows = lastRow - 1;
  var data     = sheet.getRange(2, 1, numRows, 20).getValues();
  var tz       = Session.getScriptTimeZone();
  var pos      = [];

  // getRichTextValues lets us read hyperlinks that getValues() strips out.
  // Column A (index 1) holds a legacy manually-pasted invoice hyperlink on
  // the PO number cell -- only used as a fallback when column N (Invoice
  // File) is empty. Column J (index 10) holds the issued-PO link.
  var colARich = sheet.getRange(2, 1,  numRows, 1).getRichTextValues();
  var colJRich = sheet.getRange(2, 10, numRows, 1).getRichTextValues();

  data.forEach(function(row, i) {
    var poNum = row[0] ? row[0].toString().trim() : "";
    if (!isValidPONumber(poNum)) return; // skip header / input rows

    var dateIssued   = formatDateCell(row[1], tz);
    var deliveryDate = formatDateCell(row[8], tz);

    // Extract hyperlink URLs from rich-text cells
    var legacyInvoiceLink = "";
    var issuedPOLink      = "";
    try { legacyInvoiceLink = colARich[i][0].getLinkUrl() || ""; } catch(e) {}
    try { issuedPOLink      = colJRich[i][0].getLinkUrl() || ""; } catch(e) {}

    // Column J may also just contain a plain-text URL
    if (!issuedPOLink) issuedPOLink = str(row[9]);

    var invoiceFile = canViewInvoice ? str(row[14]) : "";

    var receivedBy = [];
    try { receivedBy = JSON.parse(row[15] || '[]'); } catch (e) { receivedBy = []; }
    receivedBy = receivedBy.map(function(p) {
      return { name: p.name || p.email || '', email: p.email || '', at: formatIsoTimestamp_(p.at, tz) };
    });

    pos.push({
      rowIndex:     i + 2,
      poNum:        poNum,
      dateIssued:   dateIssued,
      builder:      str(row[2]),
      jobRef:       str(row[3]),
      vendor:       str(row[4]),
      vendorInvoice:str(row[5]),
      status:       str(row[6]).trim(),
      invoiceTotal: canViewInvoice ? str(row[7]) : "",
      deliveryDate: deliveryDate,
      issuedPO:     str(row[9]),
      issuedPOLink: issuedPOLink,
      invoiceFile:  invoiceFile,
      invoiceLink:  canViewInvoice ? (invoiceFile || legacyInvoiceLink) : "",
      receivedNote: str(row[10]),
      notes:        str(row[11]),
      additionalNotes: str(row[12]),
      orderedBy:    str(row[13]),
      receivedBy:   receivedBy,
      invoiceApprovedBy: canViewInvoice ? str(row[16]) : "",
      invoiceApprovedAt: canViewInvoice ? formatIsoTimestamp_(row[17], tz) : "",
      invoiceUploadedBy: canViewInvoice ? str(row[18]) : "",
      invoiceUploadedAt: canViewInvoice ? formatIsoTimestamp_(row[19], tz) : ""
    });
  });

  return pos;
}

/**
 * Returns just the first whitespace-separated token of a full name.
 */
function getFirstName(fullName) {
  var trimmed = (fullName || "").toString().trim();
  if (!trimmed) return "";
  return trimmed.split(/\s+/)[0];
}

/**
 * Creates a new PO row and returns { success, poNumber } or { success: false, error }.
 */
function createPO(data) {
  try {
    var auth = authorizeCaller(data, ['admin', 'office', 'site_manager']);
    if (!auth.ok) return { success: false, error: auth.error, code: auth.code };

    if (!data.jobRef || !data.vendor) {
      return { success: false, error: "Job Reference and Vendor are required." };
    }

    var lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) {
      return { success: false, error: "Server is busy - try again in a moment." };
    }
    var nextRow, poNumber, result, cache, cacheKey;
    try {
      // Idempotency guard: the client keeps re-sending the same key on a
      // retry after a client-side timeout (the server keeps running after
      // the client gives up waiting, so a naive retry would create a real
      // duplicate PO). A cache hit here means this exact submission already
      // succeeded -- return that prior result instead of writing a new row.
      var idemKey = (data.idempotencyKey || '').toString().trim();
      cache = CacheService.getScriptCache();
      cacheKey = idemKey ? ('idem_createpo_' + idemKey) : null;
      if (cacheKey) {
        var cached = null;
        try { cached = cache.get(cacheKey); } catch (e) { /* ignore, fall through to a fresh write */ }
        if (cached) return JSON.parse(cached);
      }

      var sheet = getSheet();
      var now   = new Date();
      var tz    = Session.getScriptTimeZone();
      var year  = Utilities.formatDate(now, tz, "yy");
      var qtr   = Math.ceil((now.getMonth() + 1) / 3);
      var paddedQtr = ("0" + qtr).slice(-2);

      nextRow  = sheet.getLastRow() + 1;
      poNumber = year + "-" + paddedQtr + "-" + Utilities.formatString("%03d", nextRow);
      var today    = Utilities.formatDate(now, tz, "MM/dd/yyyy");
      var status   = data.status || "Pending Pickup";

      // Pending Pickup POs are picked up the same day they're created, so
      // default the pickup/delivery date to today rather than leaving it blank.
      var pickupDate = (status === "Pending Pickup") ? today : "";

      var row = [
        poNumber,                          // 1
        today,                             // 2
        data.builder       || "",          // 3
        data.jobRef        || "",          // 4
        data.vendor        || "",          // 5
        data.vendorInvoice || "",          // 6
        status,                            // 7
        data.invoiceTotal  || "",          // 8
        pickupDate,                        // 9
        "",                                // 10
        "",                                // 11
        data.notes           || "",        // 12
        data.additionalNotes || "",        // 13
        getFirstName(data.orderedBy)       // 14
      ];
      sheet.getRange(nextRow, 1, 1, row.length).setValues([row]);

      result = { success: true, poNumber: poNumber, rowIndex: nextRow };
      if (cacheKey) {
        try { cache.put(cacheKey, JSON.stringify(result), 300); } catch (e) { /* fine to skip -- worst case a genuine retry within 10s window creates a second row, same as before this fix */ }
      }
      invalidateConfigOptionsCache_(); // this builder/job pair may be new -- don't let the picker lag up to CONFIG_OPTIONS_CACHE_TTL_SEC behind
    } finally {
      lock.releaseLock();
    }

    // Sent after the lock is released -- it's a live network round-trip
    // (even parallelized across recipients/devices) and holding the lock
    // for it only makes concurrent PO submissions wait longer than needed.
    sendPushNotification(OWNER_EMAILS, 'New PO Created: ' + poNumber,
      'By ' + (data.orderedBy || auth.email) + ' - ' + (data.jobRef || '') + ' / ' + (data.vendor || ''), '/');

    return result;
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Creates a "sub-PO" that shares an existing PO's base number with a letter
 * suffix (e.g. 25-04-132 -> 25-04-132B -> 25-04-132C ...), for splitting one
 * job/vendor order into several. Builder/Job are inherited from the parent
 * row; everything else comes from `data` just like createPO.
 * Returns { success, poNumber, rowIndex } or { success: false, error }.
 */
function createSubPO(data) {
  try {
    var auth = authorizeCaller(data, ['admin', 'office', 'site_manager']);
    if (!auth.ok) return { success: false, error: auth.error, code: auth.code };

    if (!data.parentPoNumber || !data.vendor) {
      return { success: false, error: "Parent PO number and Vendor are required." };
    }

    var baseMatch = data.parentPoNumber.toString().trim().match(/^(\d{2}-\d{2}-\d{3,4})([A-Z])?$/);
    if (!baseMatch) return { success: false, error: "Invalid parent PO number." };
    var basePoNumber = baseMatch[1];

    var lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) {
      return { success: false, error: "Server is busy - try again in a moment." };
    }
    var nextRow, subPoNumber, jobRef, result, cache, cacheKey;
    try {
      // Idempotency guard -- see createPO() for the full rationale. A cache
      // hit here means this exact submission already succeeded.
      var idemKey = (data.idempotencyKey || '').toString().trim();
      cache = CacheService.getScriptCache();
      cacheKey = idemKey ? ('idem_createsubpo_' + idemKey) : null;
      if (cacheKey) {
        var cached = null;
        try { cached = cache.get(cacheKey); } catch (e) { /* ignore, fall through to a fresh write */ }
        if (cached) return JSON.parse(cached);
      }

      var sheet   = getSheet();
      var lastRow = sheet.getLastRow();
      var numRows = lastRow - 1;
      var poCol   = numRows > 0 ? sheet.getRange(2, 1, numRows, 1).getValues() : [];

      var parentRow    = null;
      var maxLetterCode = 64; // 'A' - 1, so first sub-PO is 'B'
      for (var i = 0; i < poCol.length; i++) {
        var v = (poCol[i][0] || "").toString().trim();
        var m = v.match(/^(\d{2}-\d{2}-\d{3,4})([A-Z])?$/);
        if (!m || m[1] !== basePoNumber) continue;
        if (!m[2]) {
          parentRow = i + 2;
        } else if (m[2].charCodeAt(0) > maxLetterCode) {
          maxLetterCode = m[2].charCodeAt(0);
        }
      }
      if (!parentRow) {
        return { success: false, error: "Original PO " + basePoNumber + " not found." };
      }

      var nextLetter  = String.fromCharCode(maxLetterCode + 1);
      subPoNumber = basePoNumber + nextLetter;

      var parentVals = sheet.getRange(parentRow, 1, 1, 4).getValues()[0];
      var builder    = parentVals[2];
      jobRef         = parentVals[3];

      nextRow = lastRow + 1;
      var now     = new Date();
      var tz      = Session.getScriptTimeZone();
      var today   = Utilities.formatDate(now, tz, "MM/dd/yyyy");
      var status  = data.status || "Pending Pickup";

      var pickupDate = (status === "Pending Pickup") ? today : "";

      var row = [
        subPoNumber,                       // 1
        today,                             // 2
        builder || "",                     // 3
        jobRef  || "",                     // 4
        data.vendor        || "",          // 5
        data.vendorInvoice || "",          // 6
        status,                            // 7
        data.invoiceTotal  || "",          // 8
        pickupDate,                        // 9
        data.issuedPO        || "",        // 10
        "",                                // 11
        data.notes           || "",        // 12
        data.additionalNotes || "",        // 13
        getFirstName(data.orderedBy),      // 14
        data.invoiceFile     || ""         // 15
      ];
      sheet.getRange(nextRow, 1, 1, row.length).setValues([row]);

      result = { success: true, poNumber: subPoNumber, rowIndex: nextRow };
      if (cacheKey) {
        try { cache.put(cacheKey, JSON.stringify(result), 300); } catch (e) { /* fine to skip */ }
      }
      invalidateConfigOptionsCache_();
    } finally {
      lock.releaseLock();
    }

    // Sent after the lock is released, same reasoning as createPO.
    sendPushNotification(OWNER_EMAILS, 'Sub-PO Created: ' + subPoNumber,
      'By ' + (data.orderedBy || auth.email) + ' - ' + (jobRef || '') + ' / ' + (data.vendor || ''), '/');

    return result;
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Updates specific fields on an existing PO row.
 * Only fields present in `updates` are written.
 */
function updatePO(payload) {
  try {
    // Every role except human_resources can reach this via the Receive flow
    // (canReceivePO in index.html), which only ever writes status/receivedNote/
    // notes — so runner is included here even though it can't open the full
    // Edit form (canEdit).
    var auth = authorizeCaller(payload, ['admin', 'office', 'site_manager', 'runner']);
    if (!auth.ok) return { success: false, error: auth.error, code: auth.code };

    var rowIndex = payload.rowIndex;
    var updates  = payload.updates || {};
    var sheet = getSheet();

    if (updates.builder       !== undefined) sheet.getRange(rowIndex, 3).setValue(updates.builder);
    if (updates.jobRef        !== undefined) sheet.getRange(rowIndex, 4).setValue(updates.jobRef);
    if (updates.vendor        !== undefined) sheet.getRange(rowIndex, 5).setValue(updates.vendor);
    if (updates.vendorInvoice !== undefined) sheet.getRange(rowIndex, 6).setValue(updates.vendorInvoice);
    if (updates.status        !== undefined) sheet.getRange(rowIndex, 7).setValue(updates.status);
    if (updates.invoiceTotal  !== undefined) sheet.getRange(rowIndex, 8).setValue(updates.invoiceTotal);
    if (updates.deliveryDate  !== undefined) sheet.getRange(rowIndex, 9).setValue(updates.deliveryDate);
    if (updates.issuedPO      !== undefined) sheet.getRange(rowIndex, 10).setValue(updates.issuedPO);
    if (updates.receivedNote     !== undefined) sheet.getRange(rowIndex, 11).setValue(updates.receivedNote);
    if (updates.notes            !== undefined) sheet.getRange(rowIndex, 12).setValue(updates.notes);
    if (updates.additionalNotes  !== undefined) sheet.getRange(rowIndex, 13).setValue(updates.additionalNotes);
    if (updates.orderedBy        !== undefined) sheet.getRange(rowIndex, 14).setValue(updates.orderedBy);
    if (updates.invoiceFile      !== undefined) sheet.getRange(rowIndex, 15).setValue(updates.invoiceFile);
    if (updates.invoiceApprovedBy !== undefined) sheet.getRange(rowIndex, 17).setValue(updates.invoiceApprovedBy);
    if (updates.invoiceApprovedAt !== undefined) sheet.getRange(rowIndex, 18).setValue(updates.invoiceApprovedAt);
    if (updates.invoiceUploadedBy !== undefined) sheet.getRange(rowIndex, 19).setValue(updates.invoiceUploadedBy);
    if (updates.invoiceUploadedAt !== undefined) sheet.getRange(rowIndex, 20).setValue(updates.invoiceUploadedAt);

    // Received By (col 16): an accumulating, deduped-by-email list of
    // {name, email, at} -- who received this PO is never client-typed, only
    // ever the verified session identity, same trust model as Reviewed By on
    // the QB Invoice Staging sheet.
    if (payload.appendReceivedBy) {
      var receivedByCell = sheet.getRange(rowIndex, 16);
      var receivedBy = [];
      try { receivedBy = JSON.parse(receivedByCell.getValue() || '[]'); } catch (e) { receivedBy = []; }
      var displayName = (getRoleByEmail(auth.email).name || auth.email || '').toString();
      var existing = receivedBy.filter(function(p) { return p.email === auth.email; })[0];
      if (existing) {
        existing.at = new Date().toISOString();
        existing.name = displayName;
      } else {
        receivedBy.push({ name: displayName, email: auth.email, at: new Date().toISOString() });
      }
      receivedByCell.setValue(JSON.stringify(receivedBy));
    }

    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Looks up a single PO by number. Returns the PO object or null.
 */
function findPOByNumber(payload) {
  try {
    var poNum = payload && payload.poNum;
    if (!poNum) return null;
    var callerEmail = verifySessionEmail_(payload && payload.sessionToken) || '';
    var canViewInvoice = hasAnyRole_(getRoleByEmail(callerEmail).effRoles, ['admin', 'office']);
    var sheet  = getSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return null;
    var numRows = lastRow - 1;
    // Search column A only - much faster than loading all columns
    var colA = sheet.getRange(2, 1, numRows, 1).getValues();
    for (var i = 0; i < colA.length; i++) {
      var cell = (colA[i][0] || '').toString().trim();
      if (cell !== poNum) continue;
      // Found - load just this single row
      var rowIndex = i + 2;
      var tz  = Session.getScriptTimeZone();
      var row = sheet.getRange(rowIndex, 1, 1, 20).getValues()[0];
      var legacyInvoiceLink = '', issuedPOLink = '';
      try { legacyInvoiceLink = sheet.getRange(rowIndex, 1,  1, 1).getRichTextValues()[0][0].getLinkUrl() || ''; } catch(e2) {}
      try { issuedPOLink      = sheet.getRange(rowIndex, 10, 1, 1).getRichTextValues()[0][0].getLinkUrl() || ''; } catch(e2) {}
      if (!issuedPOLink) issuedPOLink = str(row[9]);
      var invoiceFile = canViewInvoice ? str(row[14]) : "";
      var receivedBy = [];
      try { receivedBy = JSON.parse(row[15] || '[]'); } catch (e2) { receivedBy = []; }
      receivedBy = receivedBy.map(function(p) {
        return { name: p.name || p.email || '', email: p.email || '', at: formatIsoTimestamp_(p.at, tz) };
      });
      return {
        rowIndex:      rowIndex,
        poNum:         (row[0] || '').toString().trim(),
        dateIssued:    formatDateCell(row[1], tz),
        builder:       str(row[2]),
        jobRef:        str(row[3]),
        vendor:        str(row[4]),
        vendorInvoice: str(row[5]),
        status:        str(row[6]).trim(),
        invoiceTotal:  canViewInvoice ? str(row[7]) : "",
        deliveryDate:  formatDateCell(row[8], tz),
        issuedPO:      str(row[9]),
        issuedPOLink:  issuedPOLink,
        invoiceFile:   invoiceFile,
        invoiceLink:   canViewInvoice ? (invoiceFile || legacyInvoiceLink) : "",
        receivedNote:  str(row[10]),
        notes:         str(row[11]),
        additionalNotes: str(row[12]),
        orderedBy:     str(row[13]),
        receivedBy:    receivedBy,
        invoiceApprovedBy: canViewInvoice ? str(row[16]) : "",
        invoiceApprovedAt: canViewInvoice ? formatIsoTimestamp_(row[17], tz) : "",
        invoiceUploadedBy: canViewInvoice ? str(row[18]) : "",
        invoiceUploadedAt: canViewInvoice ? formatIsoTimestamp_(row[19], tz) : ""
      };
    }
    return null;
  } catch(e) {
    return { error: e.toString() };
  }
}

/**
 * Verifies an email + password against the Roles sheet.
 * Roles sheet columns: A = Email, B = Role, C = Password
 * Returns { success, role, email, error }
 */
function verifyLogin(email, password) {
  try {
    if (!email || !password) return { success: false, error: 'Enter your email and password.' };
    email = email.toLowerCase().trim();

    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(ROLES_SHEET);
    if (!sheet) return { success: false, error: 'System error. Contact admin.' };

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var rowEmail = (data[i][1] || '').toString().toLowerCase().trim(); // Column B
      var rowRole  = (data[i][3] || '').toString().toLowerCase().trim(); // Column D
      var rowPass  = (data[i][4] || '').toString().trim();               // Column E
      var rowName  = (data[i][0] || '').toString().trim();               // Column A
      var rowPhone = (data[i][2] || '').toString().trim();               // Column C
      if (rowEmail === email) {
        if (rowPass && rowPass === password) {
          var loginRoleList = parseRoleList_(rowRole);
          if (isOwnerEmail(email) && loginRoleList.indexOf('aidan') === -1) loginRoleList.push('aidan');
          if (!loginRoleList.length) loginRoleList = ['runner'];
          rowRole = loginRoleList.join(',');
          return {
            success: true, role: rowRole, email: email, sessionToken: issueSessionToken_(email),
            config: { statusOptions: STATUS_OPTIONS, vendorOptions: VENDOR_OPTIONS, userRole: rowRole, userEmail: email, userName: rowName, userPhone: rowPhone }
          };
        } else {
          return { success: false, error: 'Incorrect password.' };
        }
      }
    }
    return { success: false, error: 'Email not recognized. Contact your admin.' };
  } catch(e) {
    return { success: false, error: 'System error. Try again.' };
  }
}

/**
 * Verifies a Google Identity Services ID token (from the "Sign in with Google"
 * button) and looks up the resulting email in the Roles sheet, same as verifyLogin.
 * Restricted to GOOGLE_HD_DOMAIN — other Google accounts must use email+password.
 */
function verifyGoogleLogin(idToken) {
  try {
    if (!idToken) return { success: false, error: 'Missing Google credential.' };

    var resp = UrlFetchApp.fetch(
      'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken),
      { muteHttpExceptions: true }
    );
    if (resp.getResponseCode() !== 200) {
      return { success: false, error: 'Could not verify Google sign-in. Try again.' };
    }

    var token = JSON.parse(resp.getContentText());
    if (token.aud !== GOOGLE_CLIENT_ID) {
      return { success: false, error: 'Invalid Google sign-in.' };
    }
    if (token.email_verified !== 'true' && token.email_verified !== true) {
      return { success: false, error: 'Google email is not verified.' };
    }

    var email = (token.email || '').toLowerCase().trim();
    if (!email || email.split('@')[1] !== GOOGLE_HD_DOMAIN) {
      return { success: false, error: 'Google sign-in is limited to @' + GOOGLE_HD_DOMAIN + ' accounts. Use your email and password instead.' };
    }
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(ROLES_SHEET);
    if (!sheet) return { success: false, error: 'System error. Contact admin.' };

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var rowEmail = (data[i][1] || '').toString().toLowerCase().trim(); // Column B
      if (rowEmail === email) {
        var rowRole  = (data[i][3] || '').toString().toLowerCase().trim(); // Column D
        var rowName  = (data[i][0] || '').toString().trim();               // Column A
        var rowPhone = (data[i][2] || '').toString().trim();               // Column C
        var googleRoleList = parseRoleList_(rowRole);
        if (isOwnerEmail(email) && googleRoleList.indexOf('aidan') === -1) googleRoleList.push('aidan');
        if (!googleRoleList.length) googleRoleList = ['runner'];
        rowRole = googleRoleList.join(',');
        return {
          success: true, role: rowRole, email: email, sessionToken: issueSessionToken_(email),
          config: { statusOptions: STATUS_OPTIONS, vendorOptions: VENDOR_OPTIONS, userRole: rowRole, userEmail: email, userName: rowName, userPhone: rowPhone }
        };
      }
    }
    return { success: false, error: 'Email not recognized. Contact your admin.' };
  } catch(e) {
    return { success: false, error: 'System error. Try again.' };
  }
}

/**
 * Returns config (status/vendor lists) + role for a cached/returning user.
 * Also called pre-login (to pre-warm GAS, with an empty email) and can be
 * called with any payload.email -- personal fields (name/phone/role) are
 * only included once payload.sessionToken verifies to an actual session, so
 * a caller can't read another user's name/phone/role just by supplying
 * their email address.
 */
var CONFIG_OPTIONS_CACHE_KEY = 'config_options_v1';
var CONFIG_OPTIONS_CACHE_TTL_SEC = 120;

/**
 * Caches { builderOptions, jobOptions } -- getBuilderNames()/getRecentJobs()
 * each do a couple of full-sheet scans (Projects + PO Database) and
 * getConfig() calls both on essentially every app load, so this was the
 * single most repeated expensive pair of calls in the backend. Same
 * CacheService pattern as getRolesMap_()/getAsanaJobs(). Invalidated by
 * createPO/createSubPO/createProjectAndTask right after a successful write
 * that could introduce a new builder/job pair, so a fresh one shows up
 * promptly instead of waiting out the full TTL.
 */
function getConfigOptions_() {
  var cache = CacheService.getScriptCache();
  try {
    var cached = cache.get(CONFIG_OPTIONS_CACHE_KEY);
    if (cached) return JSON.parse(cached);
  } catch (e) { /* fall through and rebuild */ }

  var opts = {
    builderOptions: getBuilderNames(),
    jobOptions:     getRecentJobs()
  };
  try { cache.put(CONFIG_OPTIONS_CACHE_KEY, JSON.stringify(opts), CONFIG_OPTIONS_CACHE_TTL_SEC); } catch (e) { /* fine to skip */ }
  return opts;
}

function invalidateConfigOptionsCache_() {
  try { CacheService.getScriptCache().remove(CONFIG_OPTIONS_CACHE_KEY); } catch (e) {}
}

function getConfig(payload) {
  var configOptions = getConfigOptions_();
  var base = {
    statusOptions:  STATUS_OPTIONS,
    vendorOptions:  VENDOR_OPTIONS,
    builderOptions: configOptions.builderOptions,
    jobOptions:     configOptions.jobOptions
  };
  var verifiedEmail = verifySessionEmail_(payload && payload.sessionToken);
  if (!verifiedEmail) return base;

  var roleData = getRoleByEmail(verifiedEmail);
  base.userRole  = roleData.role;
  base.userEmail = roleData.email;
  base.userName  = roleData.name;
  base.userPhone = roleData.phone;
  return base;
}

/**
 * Collapses a Builder or Job string to a loose dedup key: lowercased, with
 * every run of non-alphanumeric characters (spaces, hyphens, underscores)
 * folded to a single space. "BRIO_HOMES" / "Brio-Homes" / "brio homes" all
 * map to the same key. Free-typed Builder+Job combos drift in separator
 * and casing between entry points (Projects sheet vs. a typed PO combo),
 * which otherwise produces near-duplicate entries in the New PO picker for
 * what's really the same job -- see the duplicate-job-names investigation.
 */
function normJobKey_(s) {
  return (s || '').toString().toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
}

/**
 * Builds the searchable Builder+Job list that powers the New Purchase
 * Order form's combined lookup field. Order matters -- it's the suggestion
 * order in the datalist:
 *   1. Every "Projects" sheet row (most-recently-added row first) -- these
 *      are jobs that may not have a PO yet, so they're the most likely
 *      thing someone is about to create a first PO for.
 *   2. Distinct Builder+Job pairs from "PO Database", most recent Date
 *      Issued first, capped at MAX_PO_ENTRIES so the payload stays small.
 * De-duplicated via normJobKey_ (case- and separator-insensitive). A pair
 * already covered by a Projects row is not repeated from PO Database --
 * this also means that when a job has both a Projects row and inconsistent
 * PO-side spellings, the Projects row's exact spelling always wins as the
 * one shown/returned, since getProjectFolderId (Drive folder resolution)
 * and the Asana Task GID column key off that exact Projects-sheet text.
 */
function getRecentJobs() {
  try {
    var MAX_PO_ENTRIES = 300;
    var seen  = {};
    var result = [];
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    var projectsSheet = ss.getSheetByName(PROJECTS_SHEET_NAME);
    if (projectsSheet) {
      var pLastRow = projectsSheet.getLastRow();
      if (pLastRow >= 2) {
        var pData = projectsSheet.getRange(2, 1, pLastRow - 1, 2).getValues(); // A:Contractor, B:Job Name
        for (var i = pData.length - 1; i >= 0; i--) {
          var b = (pData[i][0] || '').toString().trim();
          var j = (pData[i][1] || '').toString().trim();
          if (!b || !j) continue;
          var pKey = normJobKey_(b) + '|' + normJobKey_(j);
          if (seen[pKey]) continue;
          seen[pKey] = true;
          result.push({ builder: b, job: j });
        }
      }
    }

    var poSheet = ss.getSheetByName(SHEET_NAME);
    if (poSheet) {
      var poLastRow = poSheet.getLastRow();
      if (poLastRow >= 6) {
        var poData = poSheet.getRange(6, 2, poLastRow - 5, 3).getValues(); // B:Date, C:Builder, D:JobRef
        var latestByKey = {};
        for (var k = 0; k < poData.length; k++) {
          var builder = (poData[k][1] || '').toString().trim();
          var jobRef  = (poData[k][2] || '').toString().trim();
          if (!builder || !jobRef) continue;
          var dKey = normJobKey_(builder) + '|' + normJobKey_(jobRef);
          if (seen[dKey]) continue; // already covered by a Projects row
          var dateVal = poData[k][0];
          var ts = (dateVal instanceof Date) ? dateVal.getTime() : 0;
          if (!latestByKey[dKey] || ts > latestByKey[dKey].ts) {
            latestByKey[dKey] = { builder: builder, job: jobRef, ts: ts };
          }
        }
        var poEntries = [];
        for (var key in latestByKey) poEntries.push(latestByKey[key]);
        poEntries.sort(function(a, b2) { return b2.ts - a.ts; });

        for (var m = 0; m < poEntries.length && m < MAX_PO_ENTRIES; m++) {
          result.push({ builder: poEntries[m].builder, job: poEntries[m].job });
        }
      }
    }

    return result;
  } catch (e) {
    return [];
  }
}

/**
 * Builds a de-duplicated (case- and separator-insensitive, via normJobKey_),
 * alphabetically sorted list of builder/company names already in use,
 * pulled from the "Projects" sheet (Contractor, col A) and the "PO
 * Database" sheet (Builder, col C). Powers the New Project form's company
 * dropdown so names stay consistent instead of drifting across free-text
 * entries. Always ends with "Other" so a genuinely new company can still
 * be typed in. Projects-sheet spelling wins ties (same reasoning as
 * getRecentJobs -- it's read first, and it's what getProjectFolderId
 * matches against for Drive folder resolution).
 */
function getBuilderNames() {
  try {
    var seen  = {}; // normJobKey_(name) -> canonical display value
    var names = [];

    function addName(raw) {
      var s = (raw || '').toString().trim();
      if (!s) return;
      var key = normJobKey_(s);
      if (!seen[key]) {
        seen[key] = true;
        names.push(s);
      }
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();

    var projectsSheet = ss.getSheetByName(PROJECTS_SHEET_NAME);
    if (projectsSheet) {
      var pLastRow = projectsSheet.getLastRow();
      if (pLastRow >= 2) {
        projectsSheet.getRange(2, 1, pLastRow - 1, 1).getValues().forEach(function(row) {
          addName(row[0]);
        });
      }
    }

    var poSheet = ss.getSheetByName(SHEET_NAME);
    if (poSheet) {
      // Rows 1-5 hold header/label rows (not data) on this sheet -- real PO
      // rows start at row 6. Reading from row 2 previously picked up the
      // "Contractor" column-label text itself as if it were a builder name.
      var poLastRow = poSheet.getLastRow();
      if (poLastRow >= 6) {
        poSheet.getRange(6, 3, poLastRow - 5, 1).getValues().forEach(function(row) { // col C = Builder
          addName(row[0]);
        });
      }
    }

    names.sort(function(a, b) { return a.toLowerCase().localeCompare(b.toLowerCase()); });
    names.push('Other');
    return names;
  } catch (e) {
    return ['Other'];
  }
}

/**
 * Updates name and phone for an employee in the HR sheet.
 */
function updateProfile(payload) {
  try {
    var auth = requireVerifiedEmail_(payload);
    if (auth.error) return auth;
    var email = auth.email;
    var name  = (payload.name  || '').toString().trim();
    var phone = (payload.phone || '').toString().trim();
    if (!email || !name) return { error: 'Missing email or name' };
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(ROLES_SHEET);
    if (!sheet) return { error: 'HR sheet not found' };
    var data  = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var rowEmail = (data[i][1] || '').toString().toLowerCase().trim();
      if (rowEmail === email) {
        sheet.getRange(i + 1, 1).setValue(name);  // Column A: Name
        sheet.getRange(i + 1, 3).setValue(phone); // Column C: Phone
        invalidateRolesCache_();
        return { success: true };
      }
    }
    return { error: 'Employee not found' };
  } catch(e) {
    return { error: e.message };
  }
}

/**
 * Looks up an employee's role(s) by a caller-supplied email address.
 * Returns { role, roles, effRoles, email, name, phone }:
 *   - role:     raw comma-joined role string (e.g. "admin,human_resources")
 *   - roles:    the same, as an array
 *   - effRoles: roles with 'aidan' normalized to 'admin' - use this for permission checks
 * Falls back to 'runner' if not found.
 */
var HR_ROLES_CACHE_KEY = 'hr_roles_map_v1';
var HR_ROLES_CACHE_TTL_SEC = 120; // short TTL bounds staleness for role edits made directly in the Sheet UI (bypassing addEmployee/updateEmployee/removeEmployee/updateProfile, which invalidate this cache themselves)

/**
 * Builds { emailLower: {role, name, phone} } from the HR sheet, or reads it
 * from cache. getRoleByEmail() previously did a full getDataRange() + linear
 * scan on every single call, and it's invoked by authorizeCaller() on nearly
 * every privileged action app-wide -- caching this made every one of those
 * calls faster, not just PO/note creation.
 */
function getRolesMap_() {
  var cache = CacheService.getScriptCache();
  try {
    var cached = cache.get(HR_ROLES_CACHE_KEY);
    if (cached) return JSON.parse(cached);
  } catch (e) { /* fall through and rebuild */ }

  var map = {};
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(ROLES_SHEET);
  if (sheet) {
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var email = (data[i][1] || '').toString().toLowerCase().trim(); // Column B
      if (!email) continue;
      map[email] = {
        role:  (data[i][3] || '').toString().toLowerCase().trim(), // Column D
        name:  (data[i][0] || '').toString().trim(),               // Column A
        phone: (data[i][2] || '').toString().trim()                // Column C
      };
    }
  }
  try { cache.put(HR_ROLES_CACHE_KEY, JSON.stringify(map), HR_ROLES_CACHE_TTL_SEC); } catch (e) { /* e.g. over the 100KB cache limit -- fine, just skip caching */ }
  return map;
}

/** Clears the cached role map. Call after any write to the HR sheet's Name/Email/Phone/Role columns. */
function invalidateRolesCache_() {
  try { CacheService.getScriptCache().remove(HR_ROLES_CACHE_KEY); } catch (e) {}
}

function getRoleByEmail(email) {
  try {
    if (!email) return { role: 'runner', roles: ['runner'], effRoles: ['runner'], email: '' };
    email = email.toLowerCase().trim();

    var row = getRolesMap_()[email];
    if (!row) {
      var notFoundRoles = isOwnerEmail(email) ? ['aidan'] : ['runner'];
      return { role: notFoundRoles.join(','), roles: notFoundRoles, effRoles: normalizeRoleList_(notFoundRoles), email: email, name: '', phone: '' };
    }

    var roleList = parseRoleList_(row.role);
    if (isOwnerEmail(email) && roleList.indexOf('aidan') === -1) roleList.push('aidan');
    if (!roleList.length) roleList = ['runner'];
    return { role: roleList.join(','), roles: roleList, effRoles: normalizeRoleList_(roleList), email: email, name: row.name, phone: row.phone };
  } catch(e) {
    var errRoles = isOwnerEmail(email) ? ['aidan'] : ['runner'];
    return { role: errRoles.join(','), roles: errRoles, effRoles: normalizeRoleList_(errRoles), email: email, name: '', phone: '' };
  }
}

// ─── Session Tokens ──────────────────────────────────────────────────────────
// Signed at login so server-side code never has to trust a client-supplied
// email for identity -- closes buddy-punching / payroll-spoofing since a
// token can't be forged without SESSION_SECRET, which never leaves the
// server. Same PropertiesService-backed-secret pattern as CLAUDE_API_KEY.

var SESSION_TOKEN_TTL_MS = 30 * 24 * 60 * 60 * 1000; // 30 days

function getSessionSecret_() {
  var props  = PropertiesService.getScriptProperties();
  var secret = props.getProperty('SESSION_SECRET');
  if (!secret) {
    secret = Utilities.getUuid() + Utilities.getUuid();
    props.setProperty('SESSION_SECRET', secret);
  }
  return secret;
}

/** Issues a signed session token for email, to be stored client-side and sent back on every call. */
function issueSessionToken_(email) {
  email = (email || '').toString().toLowerCase().trim();
  var body = email + '|' + (Date.now() + SESSION_TOKEN_TTL_MS);
  var sig  = Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(body, getSessionSecret_()));
  return Utilities.base64EncodeWebSafe(body) + '.' + sig;
}

/**
 * Verifies a session token's signature and expiry and returns the email it
 * was issued for, or null if the token is missing, malformed, expired, or
 * doesn't match its signature (forged / tampered / signed with a stale secret).
 */
function verifySessionEmail_(token) {
  try {
    if (!token || token.indexOf('.') === -1) return null;
    var dot  = token.indexOf('.');
    var body = Utilities.newBlob(Utilities.base64DecodeWebSafe(token.substring(0, dot))).getDataAsString();
    var sig  = token.substring(dot + 1);
    var expectedSig = Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(body, getSessionSecret_()));
    if (sig !== expectedSig) return null;

    var pipe    = body.lastIndexOf('|');
    var email   = body.substring(0, pipe).toLowerCase().trim();
    var expires = parseInt(body.substring(pipe + 1), 10);
    if (!email || !expires || Date.now() > expires) return null;
    return email;
  } catch (e) {
    return null;
  }
}

/**
 * Requires a valid session token on payload.sessionToken and returns the
 * verified email, or an { error, code: 'AUTH_REQUIRED' } object if missing/
 * invalid. Use for actions any logged-in user may take on their own behalf
 * (no role restriction) where the identity itself must still be genuine --
 * e.g. clocking in/out as yourself.
 */
function requireVerifiedEmail_(payload) {
  var email = verifySessionEmail_(payload && payload.sessionToken);
  if (!email) return { error: 'Your session has expired. Please sign in again.', code: 'AUTH_REQUIRED' };
  return { email: email };
}

/**
 * Server-side authorization gate for privileged actions. Requires
 * payload.sessionToken to verify (see verifySessionEmail_ above) to an email
 * that resolves (via getRoleByEmail, which applies the owner override above)
 * to one of allowedRoles. Callers must check .ok before proceeding.
 */
function authorizeCaller(payload, allowedRoles) {
  var callerEmail = verifySessionEmail_(payload && payload.sessionToken);
  if (!callerEmail) return { ok: false, code: 'AUTH_REQUIRED', error: 'Your session has expired. Please sign in again.' };
  var effRoles = getRoleByEmail(callerEmail).effRoles;
  if (!hasAnyRole_(effRoles, allowedRoles)) {
    return { ok: false, code: 'FORBIDDEN', error: 'You do not have permission to do this.' };
  }
  return { ok: true, role: effRoles[0], roles: effRoles, email: callerEmail };
}

/** Counts rows whose role list (column D, index 3) includes 'admin' after normalizing (covers both 'admin' and 'aidan'). */
function countAdminRows(data) {
  var n = 0;
  for (var i = 0; i < data.length; i++) {
    var rowEffRoles = normalizeRoleList_(parseRoleList_(data[i][3]));
    if (rowEffRoles.indexOf('admin') !== -1) n++;
  }
  return n;
}

// ─── Pricing ─────────────────────────────────────────────────────────────────

var PRICING_SHEET = "Pricing";
var PRICING_SHEET_CACHE_KEY = 'pricing_sheet_raw_v1';
var PRICING_SHEET_CACHE_TTL_SEC = 60; // short -- pricing edits should show up quickly after updatePricing writes

/**
 * Caches the raw Pricing sheet read ({headers, data}) that getPricingData()
 * and getMaterialCatalog() were each independently doing on every call.
 * Invalidated by updatePricing() so a same-user edit is never seen as stale
 * by the very screen that just wrote it.
 */
function getPricingSheetRaw_() {
  var cache = CacheService.getScriptCache();
  try {
    var cached = cache.get(PRICING_SHEET_CACHE_KEY);
    if (cached) return JSON.parse(cached);
  } catch (e) { /* fall through and rebuild */ }

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(PRICING_SHEET);
  var lastRow = sheet ? sheet.getLastRow() : 0;
  var lastCol = sheet ? sheet.getLastColumn() : 0;
  var raw = { headers: [], data: [] };
  if (sheet && lastRow >= 2 && lastCol >= 2) {
    raw.headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    raw.data    = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  }
  try { cache.put(PRICING_SHEET_CACHE_KEY, JSON.stringify(raw), PRICING_SHEET_CACHE_TTL_SEC); } catch (e) { /* e.g. over the 100KB cache limit -- fine, just skip caching */ }
  return raw;
}

function invalidatePricingCache_() {
  try { CacheService.getScriptCache().remove(PRICING_SHEET_CACHE_KEY); } catch (e) {}
}

/**
 * Reads the Pricing sheet and returns { vendors, items }.
 * Vendor columns are read dynamically from the header row (E onwards),
 * so adding a new vendor column to the sheet requires no code changes.
 *
 * Layout: A=Description, B=U/M, C=Best Price, D=empty, E+=Vendors
 * Category header rows: description in A, everything else blank - no U/M and no prices.
 */
function getPricingData(payload) {
  try {
    var auth = authorizeCaller(payload, ['admin', 'office']);
    if (!auth.ok) return { error: auth.error, code: auth.code };

    var raw = getPricingSheetRaw_();
    var headers = raw.headers;
    var data    = raw.data;
    if (!data.length || headers.length < 5) return { vendors: [], items: [] };

    // Discover vendor columns (E onwards = index 4+)
    var vendorCols = []; // [{ name, colIndex }]
    for (var c = 4; c < headers.length; c++) {
      var h = (headers[c] || '').toString().trim();
      if (h) vendorCols.push({ name: h, colIndex: c });
    }

    var items = [];
    var currentCategory = '';

    data.forEach(function(row, rowIdx) {
      var desc = (row[0] || '').toString().trim();
      var um   = (row[1] || '').toString().trim();
      if (!desc) return;

      var bestPrice = parseFloat(row[2]) || 0;

      // Collect vendor prices from all discovered vendor columns
      var prices = {};
      vendorCols.forEach(function(vc) {
        var v = row[vc.colIndex];
        if (v !== '' && v !== null && v !== undefined && v !== 0) {
          prices[vc.name] = parseFloat(v) || 0;
        }
      });

      var hasPrices = bestPrice > 0 || Object.keys(prices).length > 0;

      // Category header: description in A, no U/M, no prices
      if (!um && !hasPrices) {
        currentCategory = desc;
        return;
      }

      items.push({
        description:  desc,
        um:           um,
        bestPrice:    bestPrice,
        prices:       prices,
        category:     currentCategory,
        rowIndex:     rowIdx + 2
      });
    });

    var lastUpdated = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId()).getLastUpdated();
    var tz = Session.getScriptTimeZone();
    var lastUpdatedStr = Utilities.formatDate(lastUpdated, tz, "MMM d, yyyy");

    return { vendors: vendorCols.map(function(vc){ return vc.name; }), items: items, lastUpdated: lastUpdatedStr };
  } catch(e) {
    return [];
  }
}

/**
 * Updates vendor prices for a single material row.
 * Auto-calculates best price as the minimum of all entered vendor prices.
 */
function updatePricing(payload) {
  try {
    var auth = authorizeCaller(payload, ['admin', 'office']);
    if (!auth.ok) return { success: false, error: auth.error, code: auth.code };

    var rowIndex     = payload.rowIndex;
    var vendorPrices = payload.vendorPrices || {};

    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(PRICING_SHEET);
    if (!sheet) return { success: false, error: 'Pricing sheet not found' };

    // Vendor columns are read dynamically from the header row (E onwards),
    // same as getPricingData() -- this used to reference an undefined
    // PRICING_VENDORS constant, so every pricing edit failed silently.
    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var vendorCols = []; // [{ name, colIndex(1-based) }], in sheet column order
    for (var c = 4; c < headers.length; c++) {
      var h = (headers[c] || '').toString().trim();
      if (h) vendorCols.push({ name: h, colIndex: c + 1 });
    }
    if (!vendorCols.length) return { success: false, error: 'No vendor columns found on Pricing sheet' };

    // Vendor columns are contiguous (E onward) -- batch them into one
    // setValues() call instead of one setValue() per vendor.
    var minCol = vendorCols[0].colIndex;
    var maxCol = vendorCols[vendorCols.length - 1].colIndex;
    var allPrices = [];
    var rowValues = [];
    for (var col = minCol; col <= maxCol; col++) {
      var vc = null;
      for (var i = 0; i < vendorCols.length; i++) {
        if (vendorCols[i].colIndex === col) { vc = vendorCols[i]; break; }
      }
      if (!vc) { rowValues.push(''); continue; }
      var price = vendorPrices[vc.name];
      if (price !== '' && price !== null && price !== undefined) {
        var val = parseFloat(price);
        rowValues.push(isNaN(val) ? '' : val);
        if (!isNaN(val) && val > 0) allPrices.push(val);
      } else {
        rowValues.push('');
      }
    }
    sheet.getRange(rowIndex, minCol, 1, rowValues.length).setValues([rowValues]);

    // Best price = lowest vendor price, written to col C (not contiguous with E+)
    var bestPrice = allPrices.length > 0 ? Math.min.apply(null, allPrices) : '';
    sheet.getRange(rowIndex, 3).setValue(bestPrice);

    invalidatePricingCache_();
    return { success: true, bestPrice: bestPrice };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

// ─── Private Helpers ─────────────────────────────────────────────────────────

function getOrCreateErrorLogSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('ErrorLog');
  if (!sheet) {
    sheet = ss.insertSheet('ErrorLog');
    sheet.appendRow(['Timestamp', 'Action', 'CallerEmail', 'Error', 'PayloadSummary']);
  }
  return sheet;
}

function sanitizePayloadForLog_(payload) {
  var clone = {};
  var keys = Object.keys(payload || {});
  for (var i = 0; i < keys.length; i++) {
    var k = keys[i];
    if (k === 'base64Data') {
      clone[k] = '[omitted, ' + ((payload[k] || '').length) + ' chars]';
    } else {
      clone[k] = payload[k];
    }
  }
  try {
    return JSON.stringify(clone).slice(0, 500);
  } catch (e) {
    return '[unserializable]';
  }
}

function logError_(action, errorText, payload) {
  try {
    var sheet = getOrCreateErrorLogSheet_();
    sheet.appendRow([
      new Date(),
      action,
      (payload && payload.callerEmail) || '',
      errorText,
      sanitizePayloadForLog_(payload)
    ]);
  } catch (e) {
    // never let logging break the response
  }
}

function getSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error("Sheet '" + SHEET_NAME + "' not found.");
  return sheet;
}

function isValidPONumber(s) {
  return /^\d{2}-\d{2}-\d{3,4}[A-Z]?$/.test(s);
}

function formatDateCell(cell, tz) {
  if (!cell) return "";
  if (cell instanceof Date && !isNaN(cell)) {
    return Utilities.formatDate(cell, tz, "MM/dd/yyyy");
  }
  return cell.toString();
}

/** Formats an ISO-8601 timestamp string (as sent over gasCall's JSON bridge, e.g. Invoice Approved/Uploaded At) to MM/dd/yyyy. Not a Sheets Date cell -- see formatDateCell for that. */
function formatIsoTimestamp_(isoStr, tz) {
  if (!isoStr) return "";
  var d = new Date(isoStr);
  if (isNaN(d)) return isoStr.toString();
  return Utilities.formatDate(d, tz, "MM/dd/yyyy");
}

function str(val) {
  return val !== null && val !== undefined ? val.toString() : "";
}

// ─── Photo Upload ─────────────────────────────────────────────────────────────

/**
 * Receives a base64-encoded file from the web app, saves it into the
 * appropriate typed subfolder under "Purchasing" (or under the matching
 * job's own Drive folder -- see resolveBaseFolder) and returns the
 * shareable URL.
 *
 * Called client-side via gasCall('savePhotoToDrive', ...)
 */

/**
 * Returns the child folder of `parentFolder` named `name`, creating it
 * if it doesn't already exist.
 */
function getOrCreateChildFolder(parentFolder, name) {
  var existing = parentFolder.getFoldersByName(name);
  if (existing.hasNext()) return existing.next();
  return parentFolder.createFolder(name);
}

/**
 * The top-level "Purchasing" folder at Drive root, auto-created on first
 * use. This is the default destination for uploads whose Builder+Job
 * doesn't match a row in the "Projects" sheet.
 */
function getPurchasingRootFolder() {
  return getOrCreateChildFolder(DriveApp.getRootFolder(), "Purchasing");
}

/**
 * Resolves the base folder an upload's typed subfolders should live under:
 * the matching job's own Drive folder (Projects sheet lookup) if one
 * exists and is accessible, else the global "Purchasing" folder.
 * isProjectFolder tells the caller whether the returned folder is the
 * job's own (so a later write failure there should retry into Purchasing
 * rather than erroring out -- see savePhotoToDrive). noJobFolder tells the
 * caller whether this upload is landing in the Purchasing fallback because
 * the job has no working Drive folder registered -- surfaced to the user
 * (and, for office/admin, offered as a one-paste fix) rather than silently
 * swallowed.
 */
function resolveBaseFolder(builder, jobRef) {
  var match = getProjectFolderId(builder, jobRef);
  if (match.matched && match.folderId) {
    try {
      return { folder: DriveApp.getFolderById(match.folderId), isProjectFolder: true, noJobFolder: false };
    } catch (folderErr) {
      // Drive ID is set but doesn't resolve to an accessible folder --
      // fall through to the Purchasing fallback below.
    }
  }
  return { folder: getPurchasingRootFolder(), isProjectFolder: false, noJobFolder: true };
}

/**
 * Given a base folder (from resolveBaseFolder) and a document type,
 * returns/creates the folder the file should actually be written to:
 *   'issuedPO'      -> <base>/Issued POs
 *   'invoice'       -> <base>/Invoices
 *   'receivedPhoto' -> <base>/Received Photos/<poNum>
 * Unrecognized/missing docType falls back to the base folder itself.
 */
function getTypedUploadFolder(baseFolder, docType, poNum, jobRef) {
  if (docType === 'issuedPO') return getOrCreateChildFolder(baseFolder, 'Issued POs');
  if (docType === 'invoice')  return getOrCreateChildFolder(baseFolder, 'Invoices');
  if (docType === 'receivedPhoto') {
    var photosFolder = getOrCreateChildFolder(baseFolder, 'Received Photos');
    return poNum ? getOrCreateChildFolder(photosFolder, poNum) : photosFolder;
  }
  if (docType === 'qualityWalk') {
    var qwFolder = getOrCreateChildFolder(baseFolder, 'Quality Walks');
    var safeJobRef = (jobRef || '').toString().trim().slice(0, 100);
    return safeJobRef ? getOrCreateChildFolder(qwFolder, safeJobRef) : qwFolder;
  }
  return baseFolder;
}

// "Projects" sheet in the PO Database maps each Contractor + Job Name pair
// to the Shared Drive folder for that job (columns: A Contractor, B Job
// Name, C Drive folder URL/ID). Used so uploads land in the job's own
// folder instead of the global "Purchasing" folder whenever a match exists.
//
// The New Project form (createProjectAndTask) appends rows here with just
// A Contractor, B Job Name, C Drive folder ID, D Asana Task GID -- the
// remaining form fields (address, maps link, due date, etc.) only go into
// the Asana task's notes, not this sheet. getProjectFolderId only ever
// reads A-C.
//
// Column E (QB Customer/Sub-customer Id) is a later addition, filled in
// manually per job once that job's QuickBooks sub-customer exists --
// createProjectAndTask does not write it. getProjectQuickBooksId_ is the
// column-E counterpart to getProjectFolderId, same Contractor+JobName key.
var PROJECTS_SHEET_NAME = "Projects";

/**
 * Looks up the Drive folder ID for a given Contractor + Job Name pair in
 * the "Projects" sheet. Returns { matched, folderId }: matched is true
 * only when a Contractor+Job row was found (folderId may still be null
 * if that row's Drive ID cell is blank/unparseable) -- this lets
 * resolveBaseFolder tell "job not set up at all" (fall back to
 * Purchasing) apart from "job set up but Drive ID missing/broken"
 * (block the upload). Never throws; failures read as no match.
 */
function getProjectFolderId(builder, jobRef) {
  try {
    var wantBuilder = (builder || "").toString().trim().toLowerCase();
    var wantJob     = (jobRef  || "").toString().trim().toLowerCase();
    if (!wantBuilder || !wantJob) return { matched: false, folderId: null };

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PROJECTS_SHEET_NAME);
    if (!sheet) return { matched: false, folderId: null };

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { matched: false, folderId: null };

    var data = sheet.getRange(2, 1, lastRow - 1, 3).getValues(); // A:Contractor, B:Job Name, C:Drive ID
    for (var i = 0; i < data.length; i++) {
      var rowBuilder = (data[i][0] || "").toString().trim().toLowerCase();
      var rowJob     = (data[i][1] || "").toString().trim().toLowerCase();
      if (rowBuilder === wantBuilder && rowJob === wantJob) {
        return { matched: true, folderId: extractDriveFolderId(data[i][2]) };
      }
    }
    return { matched: false, folderId: null };
  } catch (e) {
    return { matched: false, folderId: null };
  }
}

/**
 * Looks up the QuickBooks Customer/Sub-customer Id for a given Contractor +
 * Job Name pair in column E of the "Projects" sheet, same exact-match scan
 * as getProjectFolderId. Returns { matched, qbCustomerId }: matched is true
 * only when a Contractor+Job row was found (qbCustomerId may still be null
 * if that job hasn't been linked to a QuickBooks job yet).
 */
function getProjectQuickBooksId_(builder, jobRef) {
  try {
    var wantBuilder = (builder || "").toString().trim().toLowerCase();
    var wantJob     = (jobRef  || "").toString().trim().toLowerCase();
    if (!wantBuilder || !wantJob) return { matched: false, qbCustomerId: null };

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PROJECTS_SHEET_NAME);
    if (!sheet) return { matched: false, qbCustomerId: null };

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { matched: false, qbCustomerId: null };

    var data = sheet.getRange(2, 1, lastRow - 1, 5).getValues(); // A:Contractor, B:Job Name, ... E:QB Customer/Sub-customer Id
    for (var i = 0; i < data.length; i++) {
      var rowBuilder = (data[i][0] || "").toString().trim().toLowerCase();
      var rowJob     = (data[i][1] || "").toString().trim().toLowerCase();
      if (rowBuilder === wantBuilder && rowJob === wantJob) {
        var qbId = (data[i][4] || "").toString().trim();
        return { matched: true, qbCustomerId: qbId || null };
      }
    }
    return { matched: false, qbCustomerId: null };
  } catch (e) {
    return { matched: false, qbCustomerId: null };
  }
}

/**
 * Upserts the QuickBooks Customer/Sub-customer Id (column E) for an
 * existing Contractor + Job Name row in the Projects sheet -- the write
 * counterpart to getProjectQuickBooksId_ above. Owner-gated, same as the
 * rest of the QuickBooks-facing pipeline: this is what the invoice-upload
 * link gate (irCheckLinkGate/irSaveJobLink in index.html) uses to fill in a
 * missing job link inline instead of editing the sheet by hand. Deliberately
 * does NOT create a new Projects row if none matches -- the job is expected
 * to already exist (it's tied to a real PO), so a missing row means the
 * project itself was never set up, not that this Id is wrong.
 */
function saveProjectQuickBooksCustomerId(payload) {
  var auth = authorizeInvoiceReviewer_(payload);
  if (!auth.ok) return { success: false, error: auth.error, code: auth.code };

  var builder = (payload.builder || '').toString().trim();
  var jobRef = (payload.jobRef || '').toString().trim();
  var qbCustomerId = (payload.qbCustomerId || '').toString().trim();
  if (!builder || !jobRef) return { success: false, error: 'builder and jobRef are required' };
  if (!qbCustomerId) return { success: false, error: 'qbCustomerId is required' };

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { success: false, error: 'Server is busy - try again in a moment.' };
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PROJECTS_SHEET_NAME);
    if (!sheet) return { success: false, error: "Sheet '" + PROJECTS_SHEET_NAME + "' not found." };
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, error: 'No project found for ' + builder + ' / ' + jobRef + '.' };

    var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues(); // A:Contractor, B:Job Name
    var wantBuilder = builder.toLowerCase(), wantJob = jobRef.toLowerCase();
    for (var i = 0; i < data.length; i++) {
      var rowBuilder = (data[i][0] || '').toString().trim().toLowerCase();
      var rowJob = (data[i][1] || '').toString().trim().toLowerCase();
      if (rowBuilder === wantBuilder && rowJob === wantJob) {
        sheet.getRange(i + 2, 5).setValue(qbCustomerId); // column E
        return { success: true };
      }
    }
    return { success: false, error: 'No project found for ' + builder + ' / ' + jobRef + ' in the Projects sheet.' };
  } finally {
    lock.releaseLock();
  }
}

// ─── QuickBooks vendor ID mapping ────────────────────────────────────────────
// No existing sheet ties a vendor to its QuickBooks Vendor Id -- VENDOR_OPTIONS
// (line 103) is just a name list. This small sheet is the mapping, keyed by
// the same vendor-name strings so it stays in sync with the PO form's vendor
// dropdown without any code change when a vendor is added there.
var QB_VENDOR_MAP_SHEET = "QB Vendor Map";
var QB_VENDOR_MAP_HEADERS = ['Vendor Name', 'QB Vendor Id', 'QB Vendor DisplayName'];

/** Exact-match lookup, same style as getProjectFolderId/getProjectQuickBooksId_. */
function getQuickBooksVendorId_(vendorName) {
  try {
    var want = (vendorName || "").toString().trim().toLowerCase();
    if (!want) return { matched: false, qbVendorId: null };

    var sheet = ensureSheetWithHeaders_(QB_VENDOR_MAP_SHEET, QB_VENDOR_MAP_HEADERS);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { matched: false, qbVendorId: null };

    var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues(); // A:Vendor Name, B:QB Vendor Id
    for (var i = 0; i < data.length; i++) {
      var rowName = (data[i][0] || "").toString().trim().toLowerCase();
      if (rowName === want) {
        var qbId = (data[i][1] || "").toString().trim();
        return { matched: true, qbVendorId: qbId || null };
      }
    }
    return { matched: false, qbVendorId: null };
  } catch (e) {
    return { matched: false, qbVendorId: null };
  }
}

/** Lists the full mapping (for an admin screen to review/fill in gaps against VENDOR_OPTIONS). */
function getQuickBooksVendorMap(payload) {
  try {
    var auth = authorizeCaller(payload, ['admin', 'office']);
    if (!auth.ok) return { error: auth.error, code: auth.code };

    var sheet = ensureSheetWithHeaders_(QB_VENDOR_MAP_SHEET, QB_VENDOR_MAP_HEADERS);
    var lastRow = sheet.getLastRow();
    var rows = lastRow < 2 ? [] : sheet.getRange(2, 1, lastRow - 1, 3).getValues().map(function(r) {
      return { vendorName: r[0] || '', qbVendorId: r[1] || '', qbVendorDisplayName: r[2] || '' };
    });
    return { vendors: VENDOR_OPTIONS.filter(function(v){ return v !== 'Other'; }), rows: rows };
  } catch (e) {
    return { error: e.toString() };
  }
}

/** Upserts a single vendor's QB Vendor Id by exact vendorName match, appending a row if none exists yet. */
function saveQuickBooksVendorMapping(payload) {
  try {
    var auth = authorizeCaller(payload, ['admin', 'office']);
    if (!auth.ok) return { success: false, error: auth.error, code: auth.code };

    var vendorName = (payload.vendorName || '').toString().trim();
    var qbVendorId = (payload.qbVendorId || '').toString().trim();
    var qbDisplayName = (payload.qbVendorDisplayName || '').toString().trim();
    if (!vendorName || !qbVendorId) return { success: false, error: 'vendorName and qbVendorId are required' };

    var sheet = ensureSheetWithHeaders_(QB_VENDOR_MAP_SHEET, QB_VENDOR_MAP_HEADERS);
    var lastRow = sheet.getLastRow();
    var want = vendorName.toLowerCase();
    if (lastRow >= 2) {
      var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (var i = 0; i < data.length; i++) {
        if ((data[i][0] || '').toString().trim().toLowerCase() === want) {
          sheet.getRange(i + 2, 2, 1, 2).setValues([[qbVendorId, qbDisplayName]]);
          return { success: true };
        }
      }
    }
    sheet.appendRow([vendorName, qbVendorId, qbDisplayName]);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Client-facing status check for the invoice-upload link gate
 * (irCheckLinkGate in index.html): does this PO's vendor have a QuickBooks
 * Vendor Id, and does its Contractor+Job have a QuickBooks Customer/
 * Sub-customer Id? Read-only, cheap -- called before the upload dropzone is
 * ever shown so a missing link can be fixed inline instead of only
 * surfacing once Bill creation fails at the very end of the review flow.
 * Owner-gated, same as the rest of the QuickBooks-facing pipeline.
 */
function getQuickBooksLinkStatusForInvoice(payload) {
  var auth = authorizeInvoiceReviewer_(payload);
  if (!auth.ok) return { error: auth.error, code: auth.code };
  var vendorLookup = getQuickBooksVendorId_(payload.vendor);
  var customerLookup = getProjectQuickBooksId_(payload.builder, payload.jobRef);
  return {
    success: true,
    qbVendorId: vendorLookup.qbVendorId || '',
    qbCustomerId: customerLookup.qbCustomerId || '',
    projectFound: customerLookup.matched
  };
}

// ─── QuickBooks item mapping (learns from reviewer picks) ────────────────────
// Same self-creating-sheet, upsert-by-key pattern as QB Vendor Map above, but
// keyed by qboNormalizeItemText_(description) instead of vendor name --
// deliberately NOT vendor-scoped, since a generic product ("Aluminum
// Fascia") should map the same way regardless of which vendor's invoice it
// came from. Read side (getQBItemMap_, checked before the fuzzy matcher) is
// in QuickBooks_OAuth.gs next to matchLineItemToQBOItem_; this is just the
// write side, called from saveInvoiceStagingReview on Approve.
var QB_ITEM_MAP_SHEET = "QB Item Map";
var QB_ITEM_MAP_HEADERS = ['Normalized Description', 'Description Sample', 'QB Item Id', 'QB Item Name', 'Lane'];

/**
 * Upserts one description -> QBO Item mapping, keyed by normalized
 * description + lane ('stock' | 'direct'). Lane-keyed (not description
 * alone) so a material that's sometimes bought direct-for-a-job and
 * sometimes into stock can hold both mappings without one clobbering the
 * other -- see getQBItemMap_ in QuickBooks_OAuth.gs for the read side.
 * Idempotent -- safe to call every Approve even when unchanged.
 */
function saveQBItemMapping_(description, qboItemId, qboItemName, lane) {
  var key = qboNormalizeItemText_(description);
  if (!key || !qboItemId) return;
  lane = lane === 'stock' ? 'stock' : 'direct';
  var sheet = ensureSheetWithHeaders_(QB_ITEM_MAP_SHEET, QB_ITEM_MAP_HEADERS);
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    var data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    for (var i = 0; i < data.length; i++) {
      var rowLane = (data[i][4] || 'direct').toString().trim() || 'direct';
      if ((data[i][0] || '').toString().trim() === key && rowLane === lane) {
        sheet.getRange(i + 2, 2, 1, 4).setValues([[description, qboItemId, qboItemName || '', lane]]);
        return;
      }
    }
  }
  sheet.appendRow([key, description, qboItemId, qboItemName || '', lane]);
}

// ─── Purchase Line Item Log (for PPV / spend analytics) ──────────────────────
// Flat, append-only, one row per purchased line item -- written the moment a
// QuickBooks Bill is created (see logPurchaseLineItems_ call in
// createQuickBooksBill, QuickBooks_OAuth.gs), never at Approve, since that's
// the point the data becomes immutable and financially real (a real QB Bill
// Id, and saveInvoiceStagingReview refuses to edit a Posted row). This is
// deliberately forward-only -- no backfill of Bills posted before this
// existed. QB Item Id/Name is the intended join key for later analytics
// (a clean, already-confirmed vocabulary) rather than the free-text
// Description; no variance/expected-price column is computed here since the
// Pricing sheet's "Best Price" is only matched by fragile exact-string
// Description (getMaterialUnitPrice_) -- left for a real analysis pass once
// the logged data exists to look at.
var PURCHASE_LOG_SHEET = "Purchase Line Item Log";
var PURCHASE_LOG_HEADERS = [
  'Posted At', 'PO Number', 'Vendor', 'Builder', 'Job Ref', 'Vendor Invoice#',
  'Line Type', 'Description', 'Qty', 'Unit', 'Rate', 'Amount',
  'QB Item Id', 'QB Item Name', 'QB Bill Id', 'QB Customer Id', 'QB Vendor Id',
  'Staging Id', 'Extraction Method', 'Match Confidence', 'Posted By'
];

/**
 * Appends one row per non-skipped line item from a just-Posted staging row.
 * Best-effort: returns an error string (never throws) so a logging hiccup
 * surfaces as a warning on the Bill-creation response rather than undoing or
 * blocking a Bill that has already posted to QuickBooks. Skipped lines are
 * excluded since they never made it onto the Bill either.
 */
function logPurchaseLineItems_(staging, qbBillId, postedByEmail, postedAt) {
  try {
    var sheet = ensureSheetWithHeaders_(PURCHASE_LOG_SHEET, PURCHASE_LOG_HEADERS);
    var rows = (staging.lineItems || [])
      .filter(function(li) { return !li.skip; })
      .map(function(li) {
        return [
          postedAt, staging.poNumber, staging.vendor, staging.builder, staging.jobRef, staging.vendorInvoice,
          li.lineType || 'material', li.description || '', li.qty, li.unit || '', li.rate, li.amount,
          li.qboItemId || '', li.qboItemName || '', qbBillId, staging.qbCustomerId, staging.qbVendorId,
          staging.stagingId, staging.extractionMethod || '', li.matchConfidence != null ? li.matchConfidence : '',
          postedByEmail
        ];
      });
    if (rows.length) sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, PURCHASE_LOG_HEADERS.length).setValues(rows);
    return null;
  } catch (e) {
    return e.toString();
  }
}

// ─── QuickBooks invoice staging (review-before-post) ─────────────────────────
// One row per invoice extraction attempt -- the durable "extracted but not
// yet posted to QuickBooks" record. A sheet (not Cache/PropertiesService)
// because review can happen well after upload, across sessions, and needs an
// audit trail; Cache's TTL/size limits make it unsuitable as the source of
// truth here. Status moves Pending -> Approved -> Posted (or Rejected).
// Each line item in the JSON blob carries lineType: 'material' | 'tax' |
// 'freight' -- tax/freight lines skip QBO Item-matching (Phase 2) and map to
// fixed dedicated QBO Items instead, and are excluded from the material
// balance check in the review UI.
var QB_STAGING_SHEET = "QB Invoice Staging";
var QB_STAGING_HEADERS = [
  'Staging Id', 'PO Number', 'Row Index', 'Status', 'Vendor', 'Vendor Invoice#',
  'Invoice File URL', 'Builder', 'Job Ref', 'QB Customer Id', 'QB Vendor Id',
  'Line Items JSON', 'Invoice Total', 'Extracted At', 'Reviewed By', 'Approved At', 'QB Bill Id', 'Posted At',
  'Extraction Method', 'Uploaded By'
];
var QB_STAGING_COL = {}; // 0-based index by header name, built once below
QB_STAGING_HEADERS.forEach(function(h, i) { QB_STAGING_COL[h] = i; });

function stagingRowToObject_(row) {
  var lineItems = [];
  try { lineItems = JSON.parse(row[QB_STAGING_COL['Line Items JSON']] || '[]'); } catch (e) { lineItems = []; }
  return {
    stagingId:      row[QB_STAGING_COL['Staging Id']] || '',
    poNumber:       row[QB_STAGING_COL['PO Number']] || '',
    rowIndex:       row[QB_STAGING_COL['Row Index']] || '',
    status:         row[QB_STAGING_COL['Status']] || '',
    vendor:         row[QB_STAGING_COL['Vendor']] || '',
    vendorInvoice:  row[QB_STAGING_COL['Vendor Invoice#']] || '',
    invoiceFileUrl: row[QB_STAGING_COL['Invoice File URL']] || '',
    builder:        row[QB_STAGING_COL['Builder']] || '',
    jobRef:         row[QB_STAGING_COL['Job Ref']] || '',
    // .toString() here matters: Google Sheets stores a numeric-looking value
    // (e.g. a QBO Customer/Vendor/Bill Id like "799729352") as a Number cell
    // regardless of how it was written, so getValues() hands these back as
    // JS numbers. Left uncoerced, JSON.stringify(CustomerRef: {value: ...})
    // in createQuickBooksBill sends a bare JSON number where QBO's API
    // schema requires a string -- producing a confusing "element id not
    // found" fault even when the id itself is completely correct.
    qbCustomerId:   (row[QB_STAGING_COL['QB Customer Id']] || '').toString(),
    qbVendorId:     (row[QB_STAGING_COL['QB Vendor Id']] || '').toString(),
    lineItems:      lineItems,
    invoiceTotal:   row[QB_STAGING_COL['Invoice Total']] || '',
    extractedAt:    row[QB_STAGING_COL['Extracted At']] || '',
    reviewedBy:     row[QB_STAGING_COL['Reviewed By']] || '',
    approvedAt:     row[QB_STAGING_COL['Approved At']] || '',
    qbBillId:       (row[QB_STAGING_COL['QB Bill Id']] || '').toString(),
    postedAt:       row[QB_STAGING_COL['Posted At']] || '',
    extractionMethod: row[QB_STAGING_COL['Extraction Method']] || '',
    uploadedBy:     row[QB_STAGING_COL['Uploaded By']] || ''
  };
}

/**
 * Internal helper: appends a new Pending staging row. Called by
 * extractInvoiceLineItems (automatic, on upload) -- not exposed directly to
 * the client. fields: {poNumber, rowIndex, vendor, vendorInvoice,
 * invoiceFileUrl, builder, jobRef, qbCustomerId, qbVendorId, lineItems[],
 * extractionMethod: 'gemini'|'code-parser'|'manual', uploadedBy}.
 */
function createStagingRow_(fields) {
  var sheet = ensureSheetWithHeaders_(QB_STAGING_SHEET, QB_STAGING_HEADERS);
  var stagingId = Utilities.getUuid();
  var now = new Date();
  sheet.appendRow([
    stagingId,
    fields.poNumber || '',
    fields.rowIndex || '',
    'Pending',
    fields.vendor || '',
    fields.vendorInvoice || '',
    fields.invoiceFileUrl || '',
    fields.builder || '',
    fields.jobRef || '',
    fields.qbCustomerId || '',
    fields.qbVendorId || '',
    JSON.stringify(fields.lineItems || []),
    fields.invoiceTotal || '',
    now,
    '',
    '',
    '',
    '',
    fields.extractionMethod || '',
    fields.uploadedBy || ''
  ]);
  return stagingId;
}

/** Finds a staging row's sheet row index (1-based) by Staging Id, or -1 if not found. Not exposed to the client. */
function findStagingRowIndex_(sheet, stagingId) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if ((ids[i][0] || '').toString() === stagingId) return i + 2;
  }
  return -1;
}

/**
 * Lists staging rows, optionally filtered by status and/or PO number.
 * Owner-gated, same as the rest of the QuickBooks-facing workflow.
 */
function getInvoiceStaging(payload) {
  var auth = authorizeInvoiceReviewer_(payload);
  if (!auth.ok) return { error: auth.error, code: auth.code };
  try {
    var sheet = ensureSheetWithHeaders_(QB_STAGING_SHEET, QB_STAGING_HEADERS);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { rows: [] };

    var data = sheet.getRange(2, 1, lastRow - 1, QB_STAGING_HEADERS.length).getValues();
    var statusFilter = payload.status ? [].concat(payload.status) : null;
    var poFilter = (payload.poNumber || '').toString().trim();

    var rows = data.map(stagingRowToObject_).filter(function(r) {
      if (statusFilter && statusFilter.indexOf(r.status) === -1) return false;
      if (poFilter && r.poNumber.toString().trim() !== poFilter) return false;
      return true;
    });
    return { rows: rows };
  } catch (e) {
    return { error: e.toString() };
  }
}

/**
 * Persists reviewer edits to a staging row (line item overrides, matched QBO
 * item ids, skip flags, manual customer/vendor id overrides) and/or
 * transitions its Status. Owner-gated. Does not touch QuickBooks -- that
 * only happens in createQuickBooksBill once Status is 'Approved'.
 */
function saveInvoiceStagingReview(payload) {
  var auth = authorizeInvoiceReviewer_(payload);
  if (!auth.ok) return { success: false, error: auth.error, code: auth.code };
  try {
    var stagingId = payload.stagingId;
    if (!stagingId) return { success: false, error: 'Missing stagingId' };

    var sheet = ensureSheetWithHeaders_(QB_STAGING_SHEET, QB_STAGING_HEADERS);
    var rowIdx = findStagingRowIndex_(sheet, stagingId);
    if (rowIdx === -1) return { success: false, error: 'Staging row not found' };

    var currentStatus = sheet.getRange(rowIdx, QB_STAGING_COL['Status'] + 1).getValue();
    if (currentStatus === 'Posted') return { success: false, error: 'This invoice has already been posted to QuickBooks and can no longer be edited.' };

    if (payload.lineItems) {
      sheet.getRange(rowIdx, QB_STAGING_COL['Line Items JSON'] + 1).setValue(JSON.stringify(payload.lineItems));
      // Keep Invoice Total in sync with whatever the reviewer edited the lines to --
      // this is what gets pushed back onto the PO row on approve.
      var recomputedTotal = payload.lineItems.reduce(function(s, li) { return s + (parseFloat(li.amount) || 0); }, 0);
      sheet.getRange(rowIdx, QB_STAGING_COL['Invoice Total'] + 1).setValue(recomputedTotal);
    }

    // Only learn from a real Approve, never a Save Draft -- an unreviewed
    // draft's item pairing hasn't been confirmed yet. Upsert is idempotent,
    // so already-correct auto-matches just get reinforced alongside fresh
    // manual overrides -- no need to detect which case this is. Learns from
    // every line type, not just material -- tax/freight lines are read back
    // via matchInvoiceLineItems's learnedOnly mode (never the fuzzy
    // matcher), so this is what makes a manual "Sales Tax" pick stick for
    // next time instead of needing to be repicked on every invoice.
    if (payload.approve === true && payload.lineItems) {
      var jobRefForLane = sheet.getRange(rowIdx, QB_STAGING_COL['Job Ref'] + 1).getValue();
      var lane = isWarehouseJob_(jobRefForLane) ? 'stock' : 'direct';
      payload.lineItems.forEach(function(li) {
        if (li.qboItemId && li.description) {
          saveQBItemMapping_(li.description, li.qboItemId, li.qboItemName, lane);
        }
      });
    }
    if (payload.qbCustomerId !== undefined) {
      sheet.getRange(rowIdx, QB_STAGING_COL['QB Customer Id'] + 1).setValue(payload.qbCustomerId);
    }
    if (payload.qbVendorId !== undefined) {
      sheet.getRange(rowIdx, QB_STAGING_COL['QB Vendor Id'] + 1).setValue(payload.qbVendorId);
    }
    if (payload.vendorInvoice !== undefined) {
      sheet.getRange(rowIdx, QB_STAGING_COL['Vendor Invoice#'] + 1).setValue(payload.vendorInvoice);
    }

    var nextStatus = payload.approve ? 'Approved' : (payload.reject ? 'Rejected' : currentStatus);
    var approvedAt = null;
    sheet.getRange(rowIdx, QB_STAGING_COL['Status'] + 1).setValue(nextStatus);
    sheet.getRange(rowIdx, QB_STAGING_COL['Reviewed By'] + 1).setValue(auth.email || '');
    if (payload.approve) {
      approvedAt = new Date();
      sheet.getRange(rowIdx, QB_STAGING_COL['Approved At'] + 1).setValue(approvedAt);
    }

    return { success: true, status: nextStatus, reviewedBy: auth.email || '', approvedAt: approvedAt ? approvedAt.toISOString() : '' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/** Finds an already-Posted staging row for this PO+Vendor Invoice# pair, or null. Not exposed to the client. */
function findPostedStagingRow_(poNumber, vendorInvoice) {
  if (!poNumber || !vendorInvoice) return null;
  var sheet = ensureSheetWithHeaders_(QB_STAGING_SHEET, QB_STAGING_HEADERS);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  var data = sheet.getRange(2, 1, lastRow - 1, QB_STAGING_HEADERS.length).getValues();
  var wantPo = poNumber.toString().trim();
  var wantInv = vendorInvoice.toString().trim().toLowerCase();
  for (var i = 0; i < data.length; i++) {
    var obj = stagingRowToObject_(data[i]);
    if (obj.status === 'Posted' && obj.poNumber.toString().trim() === wantPo &&
        obj.vendorInvoice.toString().trim().toLowerCase() === wantInv) {
      return obj;
    }
  }
  return null;
}

// ─── Deterministic (no-AI) invoice line-item extraction ──────────────────────
// Per-vendor parsers, keyed by the exact VENDOR_OPTIONS name (line 103).
// Each parser takes the array of visual invoice-text lines reconstructed
// client-side (qbExtractInvoiceLines in index.html, which clusters pdf.js
// text items by y-position -- NOT the same as mrExtractPdfText's single
// flattened string, which throws away row/column structure) and returns an
// array of line items, or null/throws if it can't confidently parse this
// invoice. No parser registered for a vendor (or a parser that can't
// confidently parse) falls back to an empty line-item list -- the review
// screen starts blank for manual entry. This is the intended, first-class
// path, not a degraded one: it's what every vendor uses until a real
// invoice sample is available to build and verify that vendor's parser
// against. None are implemented yet -- populate this object as samples
// arrive, one parser function per vendor.
function qboClassifyLineType_(description) {
  return /freight|delivery|shipping|fuel surcharge/i.test(description || '') ? 'freight' : 'material';
}

/**
 * Scans reconstructed lines for a "(Sales) Tax(es) ... $amount" style line
 * and returns a tax-type line item, or null. Uses a lazy `.*?` between
 * "tax" and the trailing amount (not a no-digit class) since real tax
 * lines often have digits in between -- a percentage in parens ("Sales Tax
 * (7.3%) 84.42") or an unrelated code column glued onto the same visual
 * row ("SALES TAX 91-0000-00 14.36") -- only the LAST number on the line
 * is taken as the amount.
 */
/**
 * Best-effort scan for a "Invoice[#/Number/No] <value>" style line, used to
 * pre-fill the Vendor Invoice # field on upload instead of leaving it for
 * the user to retype from the PDF. Not anchored to a specific vendor
 * layout -- covers label-adjacent formats (ABC's "Invoice Number
 * 1018814305-001", Alside's "INVOICE # 156009990726", Transcending's
 * "Invoice# INV-027632", Rustic's "INVOICE 8066"). Vendors whose invoice
 * number sits in a table cell below a separate header row (Castalite,
 * Harristone, Timberline, Leak Tech, Lansing) aren't covered by this --
 * the field just stays blank and editable in that case, same as before
 * this existed. Returns '' if nothing confidently matches.
 */
function qboDetectInvoiceNumber_(lines) {
  for (var i = 0; i < lines.length; i++) {
    var m = lines[i].match(/invoice\s*(?:#|number|no\.?)?\s*:?\s*([A-Za-z0-9][A-Za-z0-9-]{2,19})\s*$/i);
    if (m) return m[1];
  }
  return '';
}

function qboExtractTaxLine_(lines) {
  for (var i = 0; i < lines.length; i++) {
    var m = lines[i].match(/(?:sales\s*)?tax(?:es)?\b.*?([\d,]+\.\d{2})\s*$/i);
    if (m) {
      var amt = parseFloat(m[1].replace(/,/g, ''));
      if (amt > 0) return { description: 'Sales Tax', qty: '', unit: '', rate: '', amount: amt, lineType: 'tax' };
    }
  }
  return null;
}

/**
 * Castalite (Castalite Brickyard) and Harristone (G.S. Harris Co./
 * Merrillstone) render from the same underlying invoice template: a
 * "Quantity [Item Code] Description Price Each Amount[T]" row, T marking a
 * taxable line. Harristone's item code is folded into the description (see
 * header comment above INVOICE_PARSERS); Castalite has no item code
 * column at all, so the same pattern matches both.
 */
function parseQtyDescPriceAmount_(lines) {
  var rowRe = /^(\d+(?:\.\d+)?)\s+(.+?)\s+([\d,]+\.\d{2,5})\s+([\d,]+\.\d{2})T?$/;
  var items = [];
  lines.forEach(function(line) {
    var m = line.match(rowRe);
    if (!m) return;
    var description = m[2].trim();
    items.push({
      description: description,
      qty: parseFloat(m[1]),
      unit: '',
      rate: parseFloat(m[3].replace(/,/g, '')),
      amount: parseFloat(m[4].replace(/,/g, '')),
      lineType: qboClassifyLineType_(description)
    });
  });
  var tax = qboExtractTaxLine_(lines);
  if (tax) items.push(tax);
  return items;
}

/**
 * Timberline Exteriors: "Item Code Description Quantity Price Each
 * Amount[T]" -- item code leads instead of quantity, so the description
 * capture comes first here and the trailing three fields are qty/price/
 * amount instead of price/amount alone. Qty allows a comma (Rustic Lumber's
 * quantities run into the thousands, e.g. "1,396").
 */
function parseTimberlineInvoice_(lines) {
  var rowRe = /^(.+?)\s+([\d,]+(?:\.\d+)?)\s+([\d,]+\.\d{2,5})\s+([\d,]+\.\d{2})T?$/;
  var items = [];
  lines.forEach(function(line) {
    var m = line.match(rowRe);
    if (!m) return;
    var description = m[1].trim();
    items.push({
      description: description,
      qty: parseFloat(m[2].replace(/,/g, '')),
      unit: '',
      rate: parseFloat(m[3].replace(/,/g, '')),
      amount: parseFloat(m[4].replace(/,/g, '')),
      lineType: qboClassifyLineType_(description)
    });
  });
  var tax = qboExtractTaxLine_(lines);
  if (tax) items.push(tax);
  return items;
}

/**
 * Lansing Building Products: "LN# Item# Description OrdQty BOQty ShipQty
 * U/M UnitPrice Amount". Uses Ship Qty (what was actually shipped/billed),
 * not Ordered Qty.
 */
function parseLansingInvoice_(lines) {
  var rowRe = /^(\d+)\s+(\S+)\s+(.+?)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\S{1,6})\s+([\d,]+\.\d{1,5})\s+([\d,]+\.\d{2})$/;
  var items = [];
  lines.forEach(function(line) {
    var m = line.match(rowRe);
    if (!m) return;
    var description = m[3].trim();
    items.push({
      description: description,
      qty: parseFloat(m[6]),
      unit: m[7],
      rate: parseFloat(m[8].replace(/,/g, '')),
      amount: parseFloat(m[9].replace(/,/g, '')),
      lineType: qboClassifyLineType_(description)
    });
  });
  var tax = qboExtractTaxLine_(lines);
  if (tax) items.push(tax);
  return items;
}

/**
 * ABC Supply Interiors: "ItemCode Description QtyOrdered QtyShipped UOM
 * ExtendedQty PriceUOM PricePerUOM ExtendedPrice" -- the most column-heavy
 * of these templates. Anchored on the specific 7-field trailing shape
 * (int, int, letters, decimal, letters, decimal, decimal) so embedded
 * measurements in the description (e.g. `40" x 49'`) don't get mistaken
 * for it -- those carry symbols (", ') that keep them from matching a
 * pure-digit token.
 */
function parseABCInteriorsInvoice_(lines) {
  var rowRe = /^(\S+)\s+(.+?)\s+(\d+)\s+(\d+)\s+([A-Za-z]{1,4})\s+([\d.]+)\s+([A-Za-z]{1,4})\s+([\d,]+\.\d{2,5})\s+([\d,]+\.\d{2})$/;
  var items = [];
  lines.forEach(function(line) {
    var m = line.match(rowRe);
    if (!m) return;
    var description = (m[1] + ' ' + m[2]).trim();
    items.push({
      description: description,
      qty: parseFloat(m[4]),
      unit: m[5],
      rate: parseFloat(m[8].replace(/,/g, '')),
      amount: parseFloat(m[9].replace(/,/g, '')),
      lineType: qboClassifyLineType_(description)
    });
  });
  var tax = qboExtractTaxLine_(lines);
  if (tax) items.push(tax);
  return items;
}

/**
 * Associated Materials (Alside Supply Center): "Quantity UOM
 * ItemDescription [Color] ProductCode UnitPrice ExtendedPrice". Some rows
 * (delivery fee, fuel surcharge) omit Quantity/UOM/UnitPrice entirely --
 * those don't match this pattern and are intentionally left for manual
 * entry rather than guessed at.
 */
function parseAssociatedMaterialsInvoice_(lines) {
  var rowRe = /^([\d.]+)\s+([A-Za-z]{1,4})\s+(.+?)\s+([\d.]+)\s+([\d.]+)$/;
  var items = [];
  lines.forEach(function(line) {
    var m = line.match(rowRe);
    if (!m) return;
    var description = m[3].trim();
    items.push({
      description: description,
      qty: parseFloat(m[1]),
      unit: m[2],
      rate: parseFloat(m[4]),
      amount: parseFloat(m[5]),
      lineType: qboClassifyLineType_(description)
    });
  });
  var tax = qboExtractTaxLine_(lines);
  if (tax) items.push(tax);
  return items;
}

/**
 * Scans reconstructed lines for a "Freight ... $amount" style line and
 * returns a freight-type line item, or null. Same lazy-match shape as
 * qboExtractTaxLine_ -- Interstate Brick's "Total Freight 353.00" is a
 * footer line, not a per-line charge like other vendors' freight rows.
 */
function qboExtractFreightLine_(lines) {
  for (var i = 0; i < lines.length; i++) {
    var m = lines[i].match(/freight\b.*?([\d,]+\.\d{2})\s*$/i);
    if (m) {
      var amt = parseFloat(m[1].replace(/,/g, ''));
      if (amt > 0) return { description: 'Freight', qty: '', unit: '', rate: '', amount: amt, lineType: 'freight' };
    }
  }
  return null;
}

/**
 * Interstate Brick (Basalite Building Products): each line item spans TWO
 * physical rows -- row 1 is "ItemNum ProductCode Qty[glued-or-spaced-UOM]
 * UnitPrice Amount", row 2 is the actual product description plus a
 * "/1,000 EA"-style per-unit-basis suffix (stripped), with an optional
 * third "Batch: ..." row that's simply not consumed by either pattern and
 * therefore ignored. Row 1's Qty+UOM are inconsistently glued ("4,240EA")
 * or spaced ("2 RL") across real samples -- \s* handles both. Tax and
 * freight are separate footer totals (multi-page invoices put them on
 * page 2), not per-line charges, so they're pulled via the shared footer
 * extractors instead of by line-item classification.
 */
function parseInterstateBrickInvoice_(lines) {
  var rowRe = /^\d{4}\s+\d+\s+([\d,]+)\s*([A-Za-z]{1,4})\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s*$/;
  var items = [];
  for (var i = 0; i < lines.length; i++) {
    var m = lines[i].match(rowRe);
    if (!m) continue;
    var qty = parseFloat(m[1].replace(/,/g, ''));
    var unit = m[2];
    var rate = parseFloat(m[3].replace(/,/g, ''));
    var amount = parseFloat(m[4].replace(/,/g, ''));
    var next = lines[i + 1];
    var description = '';
    if (next && !rowRe.test(next) && !/^batch:/i.test(next)) {
      description = next.replace(/\/\s*[\d,]+\s*[A-Za-z]+\s*$/i, '').trim();
    }
    items.push({
      description: description || 'Material',
      qty: qty, unit: unit, rate: rate, amount: amount,
      lineType: qboClassifyLineType_(description)
    });
  }
  var tax = qboExtractTaxLine_(lines);
  if (tax) items.push(tax);
  var freight = qboExtractFreightLine_(lines);
  if (freight) items.push(freight);
  return items;
}

/**
 * LKL Associates: each line item also spans TWO physical rows -- row 1 is
 * "QtyOrdered QtyShipped UOM ItemCode ConvertedQty/UOM Price/UOM Amount",
 * row 2 is the real product description. Only the item code (first token
 * after UOM) is kept from row 1; the ConvertedQty/Price-per-UOM fields in
 * between are dropped rather than captured, since they're glued to their
 * UOM with a "/" (e.g. "110.90/RL") and add no value once Amount is known
 * directly. Rate is deliberately left blank rather than guessed from that
 * glued field -- Amount and Qty are what the Bill actually needs.
 */
function parseLKLAssociatesInvoice_(lines) {
  var rowRe = /^(\d+)\s+(\d+)\s+([A-Za-z]{1,4})\s+(\S+).*?([\d,]+\.\d{2})\s*$/;
  var items = [];
  for (var i = 0; i < lines.length; i++) {
    var m = lines[i].match(rowRe);
    if (!m) continue;
    var qtyShipped = parseFloat(m[2]);
    var unit = m[3];
    var itemCode = m[4];
    var amount = parseFloat(m[5].replace(/,/g, ''));
    var next = lines[i + 1];
    var description = (next && !rowRe.test(next)) ? (itemCode + ' ' + next).trim() : itemCode;
    items.push({
      description: description, qty: qtyShipped, unit: unit, rate: '', amount: amount,
      lineType: qboClassifyLineType_(description)
    });
  }
  var tax = qboExtractTaxLine_(lines);
  if (tax) items.push(tax);
  return items;
}

/**
 * Transcending Barriers Metal: "# Item&Description Qty Rate Amount", where
 * Item&Description contains a nested "Quantity | Trim Type" sub-table for
 * the specific trims ordered (e.g. "3 Custom Drip .5/.5 Drip"). Only ONE
 * sample invoice available -- this parser folds the nested sub-table into
 * the outer description (same lazy-capture technique as Timberline/Leak
 * Tech/Rustic) rather than attempting to parse it, since a single-item,
 * single-sample invoice isn't enough to confidently generalize a nested-
 * table parser. Reviewer should expect the specific trim spec to still be
 * worth double-checking against the PDF for a while until more samples
 * confirm this holds for multi-item invoices too.
 */
function parseTranscendingBarriersMetalInvoice_(lines) {
  var rowRe = /^\d+\s+(.+?)\s+([\d.]+)\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})$/;
  var items = [];
  lines.forEach(function(line) {
    var m = line.match(rowRe);
    if (!m) return;
    var description = m[1].trim();
    items.push({
      description: description,
      qty: parseFloat(m[2]),
      unit: '',
      rate: parseFloat(m[3].replace(/,/g, '')),
      amount: parseFloat(m[4].replace(/,/g, '')),
      lineType: qboClassifyLineType_(description)
    });
  });
  var tax = qboExtractTaxLine_(lines);
  if (tax) items.push(tax);
  return items;
}

// ─── Gemini AI invoice extraction (primary) ──────────────────────────────────
// Primary extraction method for every vendor, including ones with a working
// regex parser below -- Gemini's vision + structured-output handles messier
// multi-row templates (Interstate Brick, LKL Associates) and scanned/photo
// invoices (no text layer, so the regex parsers can never handle them) that
// the deterministic parsers above can't -- and also reads the vendor invoice
// number directly off the document, closing a gap qboDetectInvoiceNumber_'s
// label-scan never covered for some vendors (e.g. Timberline, whose invoice
// number sits in a table cell rather than next to an "Invoice #" label).
// Requires a GEMINI_API_KEY Script Property (Apps Script editor -> Project
// Settings -> Script Properties); get a free key at aistudio.google.com. Any
// failure (missing key, non-2xx, quota/429, malformed JSON, safety block,
// empty lineItems) returns null uniformly -- extractInvoiceLineItems falls
// back to INVOICE_PARSERS below, then manual entry, exactly as it did before
// Gemini existed.
var GEMINI_MODEL_ = 'gemini-3.6-flash';

var GEMINI_INVOICE_SCHEMA_ = {
  type: 'object',
  properties: {
    vendorInvoice: { type: 'string' },
    lineItems: {
      type: 'array',
      items: {
        type: 'object',
        properties: {
          description: { type: 'string' },
          qty:         { type: 'number' },
          unit:        { type: 'string' },
          rate:        { type: 'number' },
          amount:      { type: 'number' },
          lineType:    { type: 'string', enum: ['material', 'tax', 'freight'] }
        },
        required: ['description', 'amount', 'lineType']
      }
    }
  },
  required: ['lineItems']
};

/**
 * Calls Gemini with the raw invoice file (PDF or photo) as inline data and
 * asks for the vendor's invoice number plus structured JSON line items in
 * the same shape the regex parsers return: {description, qty, unit, rate,
 * amount, lineType}. Returns {vendorInvoice, lineItems} on success, or null
 * on ANY failure -- see section comment above. vendorInvoice may be '' even
 * on success if Gemini didn't find one; qboDetectInvoiceNumber_ is still the
 * fallback for that case (see extractInvoiceLineItems).
 */
function callGeminiForInvoiceExtraction_(base64Data, mimeType, vendor) {
  try {
    var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) return null;

    var prompt = 'This is a vendor invoice' + (vendor ? ' from ' + vendor : '') + '. ' +
      'Extract the vendor\'s own invoice/document number (vendorInvoice, string, omit if ' +
      'not shown) and every line item as JSON. For each line item: description (string), ' +
      'qty (number, omit if not shown), unit (string, omit if not shown), rate (number, ' +
      'omit if not shown), amount (number, required), lineType ("material" for products/' +
      'services, "tax" for sales tax lines, "freight" for shipping/delivery/freight charges). ' +
      'Include footer tax and freight totals as their own line items. Do not include ' +
      'subtotal or grand-total rows as line items.';

    var resp = UrlFetchApp.fetch(
      'https://generativelanguage.googleapis.com/v1beta/models/' + GEMINI_MODEL_ + ':generateContent',
      {
        method: 'POST',
        headers: { 'x-goog-api-key': apiKey, 'Content-Type': 'application/json' },
        payload: JSON.stringify({
          contents: [{ parts: [
            { text: prompt },
            { inlineData: { mimeType: mimeType, data: base64Data } }
          ] }],
          generationConfig: {
            responseMimeType: 'application/json',
            responseSchema: GEMINI_INVOICE_SCHEMA_
          }
        }),
        muteHttpExceptions: true
      }
    );

    if (resp.getResponseCode() !== 200) return null;
    var raw = JSON.parse(resp.getContentText());
    var textOut = raw.candidates && raw.candidates[0] && raw.candidates[0].content &&
                  raw.candidates[0].content.parts && raw.candidates[0].content.parts[0] &&
                  raw.candidates[0].content.parts[0].text;
    if (!textOut) return null;
    textOut = textOut.replace(/^```json\s*/m, '').replace(/^```\s*/m, '').replace(/```\s*$/m, '').trim();

    var parsed = JSON.parse(textOut);
    if (!parsed || !Array.isArray(parsed.lineItems) || !parsed.lineItems.length) return null;

    return {
      vendorInvoice: (parsed.vendorInvoice || '').toString().trim(),
      lineItems: parsed.lineItems.map(function(li) {
        return {
          description: (li.description || '').toString(),
          qty:      li.qty  !== undefined && li.qty  !== null && li.qty  !== '' ? parseFloat(li.qty)  : '',
          unit:     (li.unit || '').toString(),
          rate:     li.rate !== undefined && li.rate !== null && li.rate !== '' ? parseFloat(li.rate) : '',
          amount:   parseFloat(li.amount) || 0,
          lineType: ['material', 'tax', 'freight'].indexOf(li.lineType) !== -1 ? li.lineType : 'material'
        };
      })
    };
  } catch (e) {
    return null;
  }
}

/**
 * Leak Tech Technologies: "Description QTY RATE AMOUNT[T]", same trailing
 * shape as Timberline. The description here is a short bold label ("Step
 * Fascia", "J - Channel", "L - Metal") with a dimensional spec that wraps
 * onto its own line below in the source PDF -- that wrapped continuation
 * has no qty/rate/amount on its row, so it just won't match this pattern
 * and is dropped, leaving the short label as the captured description.
 * That's the better outcome anyway: a short normalized label is a closer
 * match to a QBO Item name than the full dimensional spec would be.
 */
var INVOICE_PARSERS = {
  'Castalite': parseQtyDescPriceAmount_,
  'Harristone': parseQtyDescPriceAmount_,
  'Timberline': parseTimberlineInvoice_,
  'Leak Tech': parseTimberlineInvoice_,
  'Rustic': parseTimberlineInvoice_,
  'Lansing': parseLansingInvoice_,
  'ABC Interiors': parseABCInteriorsInvoice_,
  'Alside': parseAssociatedMaterialsInvoice_,
  'Interstate Brick': parseInterstateBrickInvoice_,
  'LKL Associates': parseLKLAssociatesInvoice_,
  'Transcending': parseTranscendingBarriersMetalInvoice_
};

/**
 * Extracts line items from an already-uploaded invoice PDF and stages them
 * for review. payload: {poNumber, rowIndex, vendor, vendorInvoice,
 * invoiceFileUrl, builder, jobRef, invoiceLines: string[], invoiceTotal,
 * base64Data, mimeType}. Tries Gemini first (base64Data/mimeType of the raw
 * file, works for PDFs and photos alike), falls back to INVOICE_PARSERS
 * (invoiceLines, PDF-only), then empty/manual entry -- see
 * callGeminiForInvoiceExtraction_ above for the Gemini failure contract.
 * Broad role gate (matches who can upload an invoice at all) since this
 * only stages data -- it never touches QuickBooks. Blocks re-extraction if
 * this PO+Vendor Invoice# was already posted, per the locked decision.
 */
function extractInvoiceLineItems(payload) {
  var auth = authorizeCaller(payload, ['admin', 'office', 'site_manager']);
  if (!auth.ok) return { error: auth.error, code: auth.code };
  try {
    var vendor        = (payload.vendor || '').toString().trim();
    var poNumber      = (payload.poNumber || '').toString().trim();
    var vendorInvoice = (payload.vendorInvoice || '').toString().trim();
    var invoiceLines  = payload.invoiceLines || [];

    var dup = findPostedStagingRow_(poNumber, vendorInvoice);
    if (dup) {
      return {
        success: false, blocked: true,
        error: 'This invoice (PO ' + poNumber + (vendorInvoice ? ', Vendor Invoice# ' + vendorInvoice : '') +
               ') was already posted to QuickBooks as Bill ' + dup.qbBillId + '. It was not re-extracted.',
        existingStagingId: dup.stagingId, qbBillId: dup.qbBillId
      };
    }

    var lineItems = [];
    var extractionMethod = 'manual';
    var geminiVendorInvoice = '';

    if (payload.base64Data && payload.mimeType) {
      var geminiResult = callGeminiForInvoiceExtraction_(payload.base64Data, payload.mimeType, vendor);
      if (geminiResult) {
        lineItems = geminiResult.lineItems;
        geminiVendorInvoice = geminiResult.vendorInvoice || '';
        extractionMethod = 'gemini';
      }
    }

    if (!lineItems.length) {
      var parser = INVOICE_PARSERS[vendor];
      if (parser && invoiceLines.length) {
        try {
          var parsed = parser(invoiceLines, payload.invoiceTotal);
          if (parsed && parsed.length) { lineItems = parsed; extractionMethod = 'code-parser'; }
        } catch (parseErr) {
          lineItems = []; // parser failure -> manual entry, never a half-guessed result
        }
      }
    }

    // Auto-fill what the PO itself needs, so the reviewer is confirming
    // values rather than typing them from scratch: Vendor Invoice # from
    // Gemini's read of the document (falling back to the best-effort label
    // scan when Gemini didn't find one, e.g. it wasn't attempted or the
    // number sits somewhere the model missed), Invoice Total as the sum of
    // extracted lines (the actual number that matters for the balance check
    // and the Bill).
    if (!vendorInvoice) vendorInvoice = geminiVendorInvoice || qboDetectInvoiceNumber_(invoiceLines);
    var computedTotal = lineItems.reduce(function(s, li) { return s + (parseFloat(li.amount) || 0); }, 0);
    var invoiceTotal = payload.invoiceTotal || (lineItems.length ? computedTotal : '');

    var customerLookup = getProjectQuickBooksId_(payload.builder, payload.jobRef);
    var vendorLookup    = getQuickBooksVendorId_(vendor);

    var stagingId = createStagingRow_({
      poNumber: poNumber,
      rowIndex: payload.rowIndex,
      vendor: vendor,
      vendorInvoice: vendorInvoice,
      invoiceFileUrl: payload.invoiceFileUrl,
      builder: payload.builder,
      jobRef: payload.jobRef,
      qbCustomerId: customerLookup.qbCustomerId || '',
      qbVendorId: vendorLookup.qbVendorId || '',
      lineItems: lineItems,
      invoiceTotal: invoiceTotal,
      extractionMethod: extractionMethod,
      uploadedBy: auth.email || ''
    });

    return {
      success: true, stagingId: stagingId, lineItems: lineItems, vendorInvoice: vendorInvoice, invoiceTotal: invoiceTotal,
      needsManualEntry: lineItems.length === 0, extractionMethod: extractionMethod,
      customerMatched: customerLookup.matched, vendorMatched: vendorLookup.matched
    };
  } catch (e) {
    return { error: e.toString() };
  }
}

/**
 * Pulls a folder ID out of a Drive folder URL (or passes through a bare ID).
 */
function extractDriveFolderId(driveUrlOrId) {
  var s = (driveUrlOrId || "").toString().trim();
  if (!s) return null;
  var m = s.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  if (m) return m[1];
  if (/^[a-zA-Z0-9_-]{10,}$/.test(s)) return s;
  return null;
}

/**
 * Pulls a file ID out of a Drive file URL (the file.getUrl() format,
 * https://drive.google.com/file/d/<id>/view...), or passes through a bare
 * ID. Used to re-fetch an already-uploaded invoice file's bytes (e.g. for
 * attaching it to a QuickBooks Bill) from just the URL stored on a row.
 */
function extractDriveFileId_(driveUrlOrId) {
  var s = (driveUrlOrId || "").toString().trim();
  if (!s) return null;
  var m = s.match(/\/file\/d\/([a-zA-Z0-9_-]+)/) || s.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (m) return m[1];
  if (/^[a-zA-Z0-9_-]{10,}$/.test(s)) return s;
  return null;
}

/**
 * Writes (or dedupes against) a single file in `folder`. If a file with
 * this exact name already exists there, reuses it instead of uploading a
 * second copy -- callers' filenames are either fully deterministic
 * (buildDocFileName) or day-granular (ISO date, no time), so a same-day
 * retry after a timeout still produces the same name.
 */
function writeUploadFile_(folder, filename, base64Data, mimeType) {
  var existing = folder.getFilesByName(filename);
  if (existing.hasNext()) {
    var existingFile = existing.next();
    try {
      if (existingFile.getSharingAccess() !== DriveApp.Access.ANYONE_WITH_LINK) {
        existingFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      }
    } catch (sharingErr) {}
    return { file: existingFile, duplicate: true };
  }
  var bytes = Utilities.base64Decode(base64Data);
  var blob  = Utilities.newBlob(bytes, mimeType, filename);
  return { file: folder.createFile(blob), duplicate: false };
}

function savePhotoToDrive(payload) {
  try {
    var auth = requireVerifiedEmail_(payload);
    if (auth.error) return { success: false, error: auth.error, code: auth.code };

    var base64Data = payload.base64Data, mimeType = payload.mimeType, filename = payload.filename,
        builder = payload.builder, jobRef = payload.jobRef, docType = payload.docType, poNum = payload.poNum;

    var base = resolveBaseFolder(builder, jobRef);
    var noJobFolder = !!base.noJobFolder;

    // A project's Drive ID can be syntactically valid (resolveBaseFolder's
    // DriveApp.getFolderById call succeeds) but still not fully writable --
    // wrong folder pasted in, or shared read-only. That failure only
    // surfaces here, inside getTypedUploadFolder/createFile. Rather than
    // erroring out, fall back to the Purchasing folder the same as when no
    // job folder is registered at all -- the upload isn't lost, and the
    // "no Drive folder" notice below still points office/admin at the fix.
    var written;
    try {
      var folder = getTypedUploadFolder(base.folder, docType, poNum, jobRef);
      written = writeUploadFile_(folder, filename, base64Data, mimeType);
    } catch (driveErr) {
      if (!base.isProjectFolder) throw driveErr;
      noJobFolder = true;
      var fallbackFolder = getTypedUploadFolder(getPurchasingRootFolder(), docType, poNum, jobRef);
      written = writeUploadFile_(fallbackFolder, filename, base64Data, mimeType);
    }

    var file = written.file;

    // The app's own <img> thumbnail requests (drive.google.com/thumbnail)
    // are anonymous -- login here is email/password, not Google OAuth, so
    // the browser never carries a Google session. Shared Drive membership
    // does nothing for that anonymous request; only ANYONE_WITH_LINK
    // sharing on the file itself makes the thumbnail load. This must run
    // for project-folder uploads too, not just the Purchasing fallback.
    try {
      if (file.getSharingAccess() !== DriveApp.Access.ANYONE_WITH_LINK) {
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      }
    } catch (sharingErr) {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    }

    return { success: true, url: file.getUrl(), duplicate: written.duplicate, noJobFolder: noJobFolder };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Uploads a base64-encoded file directly into a Drive folder resolved from
 * a folder ID or a pasted Drive folder link (via extractDriveFolderId).
 * Unlike savePhotoToDrive, this does not depend on a Contractor+Job Name
 * match already existing in the "Projects" sheet -- used by the New Project
 * form's Home Plans upload, which happens before that sheet row exists.
 */
function saveFileToFolderById(payload) {
  try {
    var auth = requireVerifiedEmail_(payload);
    if (auth.error) return { success: false, error: auth.error, code: auth.code };

    var base64Data = payload.base64Data, mimeType = payload.mimeType, filename = payload.filename,
        folderIdOrLink = payload.folderId;

    var folderId = extractDriveFolderId(folderIdOrLink);
    if (!folderId) {
      return { success: false, error: 'Could not read a folder ID from that Drive link.' };
    }

    var folder = DriveApp.getFolderById(folderId);

    // If a file with this exact name already exists in the target folder,
    // reuse it instead of uploading a second copy. Callers pass a stable
    // (non-timestamped) filename for this reason.
    var existing = folder.getFilesByName(filename);
    if (existing.hasNext()) {
      var existingFile = existing.next();
      return { success: true, url: existingFile.getUrl(), folderId: folderId, duplicate: true };
    }

    var bytes = Utilities.base64Decode(base64Data);
    var blob  = Utilities.newBlob(bytes, mimeType, filename);
    var file  = folder.createFile(blob);

    try {
      if (file.getSharingAccess() !== DriveApp.Access.ANYONE_WITH_LINK) {
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      }
    } catch (sharingErr) {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    }

    return { success: true, url: file.getUrl(), folderId: folderId, duplicate: false };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Registers (or fixes) the Drive folder link for a Builder+Job pair in the
 * "Projects" sheet -- the same sheet resolveBaseFolder/getProjectFolderId
 * read from. Lets office/admin close the loop right from the "no Drive
 * folder" upload notice instead of editing the spreadsheet by hand.
 * Updates the existing row if the pair is already there (covers a blank or
 * broken Drive ID cell), appends a new row otherwise. Confirms the folder
 * is actually accessible before saving so a bad paste doesn't just move
 * the same failure to the next upload.
 */
function setProjectDriveLink(payload) {
  try {
    var auth = authorizeCaller(payload, ['admin', 'office']);
    if (!auth.ok) return { success: false, error: auth.error, code: auth.code };

    var builder   = (payload.builder   || '').toString().trim();
    var jobRef    = (payload.jobRef    || '').toString().trim();
    var driveLink = (payload.driveLink || '').toString().trim();
    if (!builder || !jobRef) return { success: false, error: 'Builder and Job are required.' };

    var folderId = extractDriveFolderId(driveLink);
    if (!folderId) return { success: false, error: 'Could not read a folder ID from that Drive link.' };

    try {
      DriveApp.getFolderById(folderId);
    } catch (folderErr) {
      return { success: false, error: 'That Drive folder is not accessible to this app - check the link and its sharing settings.' };
    }

    var lock = LockService.getScriptLock();
    var haveLock = lock.tryLock(10000);
    if (!haveLock) return { success: false, error: 'Server is busy - try again in a moment.' };
    try {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PROJECTS_SHEET_NAME);
      if (!sheet) return { success: false, error: "Sheet '" + PROJECTS_SHEET_NAME + "' not found." };

      var wantBuilder = builder.toLowerCase();
      var wantJob     = jobRef.toLowerCase();
      var lastRow = sheet.getLastRow();
      var rowIndex = -1;
      if (lastRow >= 2) {
        var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
        for (var i = 0; i < data.length; i++) {
          var rb = (data[i][0] || '').toString().trim().toLowerCase();
          var rj = (data[i][1] || '').toString().trim().toLowerCase();
          if (rb === wantBuilder && rj === wantJob) { rowIndex = i + 2; break; }
        }
      }

      if (rowIndex === -1) {
        rowIndex = sheet.getLastRow() + 1;
        sheet.getRange(rowIndex, 1).setValue(builder);
        sheet.getRange(rowIndex, 2).setValue(jobRef);
      }
      sheet.getRange(rowIndex, 3).setValue(folderId);
    } finally {
      if (haveLock) lock.releaseLock();
    }

    invalidateConfigOptionsCache_();
    return { success: true, folderId: folderId };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Saves an Office Notes photo into Purchasing/Office Notes (created on first
 * use). Filenames are timestamped since, unlike Home Plans, there's no
 * per-note stable name to dedupe against -- every upload is a distinct file.
 */
function saveOfficeNotePhoto(payload) {
  try {
    var auth = requireVerifiedEmail_(payload);
    if (auth.error) return { success: false, error: auth.error, code: auth.code };

    var base64Data = payload.base64Data, mimeType = payload.mimeType, filename = payload.filename;

    var folder = getOrCreateChildFolder(getPurchasingRootFolder(), 'Office Notes');

    var bytes = Utilities.base64Decode(base64Data);
    var blob  = Utilities.newBlob(bytes, mimeType, filename);
    var file  = folder.createFile(blob);

    try {
      if (file.getSharingAccess() !== DriveApp.Access.ANYONE_WITH_LINK) {
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      }
    } catch (sharingErr) {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    }

    return { success: true, url: file.getUrl() };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Saves a mileage/commission PDF (or any file) an employee attaches from the
 * period-review screen, into Purchasing/Mileage & Commission/{their email}
 * (created on first use, one subfolder per employee for organization).
 * Auth-gated (unlike saveOfficeNotePhoto above) since the destination folder
 * depends on knowing who the caller is.
 */
function saveMileageCommissionPdf(payload) {
  try {
    var auth = requireVerifiedEmail_(payload);
    if (auth.error) return auth;
    var base64Data = (payload.base64Data || '').toString();
    var mimeType   = (payload.mimeType || 'application/pdf').toString();
    var filename   = (payload.filename || 'document.pdf').toString();
    if (!base64Data) return { error: 'No file data received.' };

    var root = getOrCreateChildFolder(getPurchasingRootFolder(), 'Mileage & Commission');
    var folder = getOrCreateChildFolder(root, auth.email);

    var bytes = Utilities.base64Decode(base64Data);
    var blob  = Utilities.newBlob(bytes, mimeType, filename);
    var file  = folder.createFile(blob);

    try {
      if (file.getSharingAccess() !== DriveApp.Access.ANYONE_WITH_LINK) {
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      }
    } catch (sharingErr) {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    }

    return { success: true, url: file.getUrl() };
  } catch (e) {
    return { error: e.toString() };
  }
}

// One-time setup helper: run this once from the Apps Script editor's Run
// menu so new files created under the "Purchasing" folder (and its Issued
// POs / Invoices / Received Photos subfolders) inherit link-sharing and
// savePhotoToDrive can skip the per-file setSharing() call above. Safe to
// re-run; safe to leave in place.
function oneTimeSetFolderSharing() {
  getPurchasingRootFolder()
    .setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
}

function authorizeDrive() {
  DriveApp.getRootFolder();
  Logger.log("Drive authorized!");
}

// ─── Contacts ─────────────────────────────────────────────────────────────────

/**
 * Reads the Contacts sheet. Row 1 = headers, rows 2+ = data.
 * Returns an array of objects keyed by header name.
 */
function getContacts(payload) {
  try {
    var auth = authorizeCaller(payload, ['admin', 'office', 'site_manager']);
    if (!auth.ok) return { headers: [], contacts: [], error: auth.error, code: auth.code };

    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Contacts');
    if (!sheet) return { headers: [], contacts: [] };
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return { headers: [], contacts: [] };
    var headers  = data[0].map(function(h){ return h.toString().trim(); }).filter(Boolean);
    var contacts = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var obj = { _rowIndex: i + 1 }; var hasData = false;
      headers.forEach(function(h, j) {
        obj[h] = (row[j] || '').toString().trim();
        if (obj[h]) hasData = true;
      });
      if (hasData) contacts.push(obj);
    }
    return { headers: headers, contacts: contacts };
  } catch(e) { return { headers: [], contacts: [], error: e.toString() }; }
}

/**
 * Updates a single contact row. `values` is an object keyed by column header.
 */
function updateContact(payload) {
  try {
    var auth = authorizeCaller(payload, ['admin', 'office', 'site_manager']);
    if (!auth.ok) return { success: false, error: auth.error, code: auth.code };

    var rowIndex = payload.rowIndex;
    var values   = payload.values;

    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Contacts');
    if (!sheet) return { success: false, error: 'Contacts sheet not found' };

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    headers.forEach(function(h, i) {
      var key = h.toString().trim();
      if (key && values[key] !== undefined) {
        sheet.getRange(rowIndex, i + 1).setValue(values[key]);
      }
    });
    SpreadsheetApp.flush();
    return { success: true };
  } catch(e) { return { error: e.toString() }; }
}

/**
 * Appends a new contact row. `values` is an object keyed by column header.
 */
function addContact(payload) {
  try {
    var auth = authorizeCaller(payload, ['admin', 'office', 'site_manager']);
    if (!auth.ok) return { success: false, error: auth.error, code: auth.code };

    var values = payload.values || {};

    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Contacts');
    if (!sheet) return { success: false, error: 'Contacts sheet not found' };

    var lastCol = Math.max(sheet.getLastColumn(), 1);
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var row = headers.map(function(h) {
      var key = h.toString().trim();
      return key && values[key] !== undefined ? values[key] : '';
    });
    if (!row.some(function(v) { return v !== '' && v !== null; })) {
      return { success: false, error: 'No contact data provided' };
    }

    sheet.appendRow(row);
    SpreadsheetApp.flush();
    return { success: true, rowIndex: sheet.getLastRow() };
  } catch(e) { return { success: false, error: e.toString() }; }
}

/**
 * Deletes a single contact row.
 */
function deleteContact(payload) {
  try {
    var auth = authorizeCaller(payload, ['admin', 'office', 'site_manager']);
    if (!auth.ok) return { success: false, error: auth.error, code: auth.code };

    var rowIndex = payload.rowIndex;
    if (!rowIndex) return { success: false, error: 'Missing rowIndex' };

    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Contacts');
    if (!sheet) return { success: false, error: 'Contacts sheet not found' };
    if (rowIndex < 2 || rowIndex > sheet.getLastRow()) return { success: false, error: 'Contact not found' };

    sheet.deleteRow(rowIndex);
    SpreadsheetApp.flush();
    return { success: true };
  } catch(e) { return { success: false, error: e.toString() }; }
}

// --- Inventory (fleet + field assets) ------------------------------------------
function ensureSheetWithHeaders_(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
  }
  return sheet;
}

var ASSET_HEADERS = ['Asset Name', 'Type', 'VIN / Serial #', 'Assigned To', 'Status', 'Next Due Date', 'Last Service Date', 'Notes'];
var ASSET_LOG_HEADERS = ['Asset Row', 'Asset Name', 'Event Type', 'Date Performed', 'Performed By', 'Next Due Date', 'Notes'];

function getInventory(payload) {
  try {
    var auth = authorizeCaller(payload, ['admin', 'office']);
    if (!auth.ok) return { headers: [], assets: [], error: auth.error, code: auth.code };

    var sheet = ensureSheetWithHeaders_('Assets', ASSET_HEADERS);
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return { headers: ASSET_HEADERS, assets: [] };
    var headers = data[0].map(function(h){ return h.toString().trim(); }).filter(Boolean);
    var assets = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var obj = { _rowIndex: i + 1 }; var hasData = false;
      headers.forEach(function(h, j) {
        obj[h] = (row[j] || '').toString().trim();
        if (obj[h]) hasData = true;
      });
      if (hasData) assets.push(obj);
    }
    return { headers: headers, assets: assets };
  } catch(e) { return { headers: [], assets: [], error: e.toString() }; }
}

function addAsset(payload) {
  try {
    var auth = authorizeCaller(payload, ['admin', 'office']);
    if (!auth.ok) return { success: false, error: auth.error, code: auth.code };

    var values = payload.values || {};
    var sheet = ensureSheetWithHeaders_('Assets', ASSET_HEADERS);
    var lastCol = Math.max(sheet.getLastColumn(), 1);
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var row = headers.map(function(h) {
      var key = h.toString().trim();
      return key && values[key] !== undefined ? values[key] : '';
    });
    if (!row.some(function(v) { return v !== '' && v !== null; })) {
      return { success: false, error: 'No asset data provided' };
    }

    sheet.appendRow(row);
    SpreadsheetApp.flush();
    return { success: true, rowIndex: sheet.getLastRow() };
  } catch(e) { return { success: false, error: e.toString() }; }
}

function updateAsset(payload) {
  try {
    var auth = authorizeCaller(payload, ['admin', 'office']);
    if (!auth.ok) return { success: false, error: auth.error, code: auth.code };

    var rowIndex = payload.rowIndex;
    var values   = payload.values || {};

    var sheet = ensureSheetWithHeaders_('Assets', ASSET_HEADERS);
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    headers.forEach(function(h, i) {
      var key = h.toString().trim();
      if (key && values[key] !== undefined) {
        sheet.getRange(rowIndex, i + 1).setValue(values[key]);
      }
    });
    SpreadsheetApp.flush();
    return { success: true };
  } catch(e) { return { success: false, error: e.toString() }; }
}

function deleteAsset(payload) {
  try {
    var auth = authorizeCaller(payload, ['admin', 'office']);
    if (!auth.ok) return { success: false, error: auth.error, code: auth.code };

    var rowIndex = payload.rowIndex;
    if (!rowIndex) return { success: false, error: 'Missing rowIndex' };

    var sheet = ensureSheetWithHeaders_('Assets', ASSET_HEADERS);
    if (rowIndex < 2 || rowIndex > sheet.getLastRow()) return { success: false, error: 'Asset not found' };
    sheet.deleteRow(rowIndex);
    SpreadsheetApp.flush();
    return { success: true };
  } catch(e) { return { success: false, error: e.toString() }; }
}

function getAssetMaintenanceLog(payload) {
  try {
    var auth = authorizeCaller(payload, ['admin', 'office']);
    if (!auth.ok) return { logs: [], error: auth.error, code: auth.code };

    var assetRowIndex = payload.assetRowIndex;
    var sheet = ensureSheetWithHeaders_('Asset Maintenance Log', ASSET_LOG_HEADERS);
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return { logs: [] };
    var headers = data[0].map(function(h){ return h.toString().trim(); }).filter(Boolean);
    var logs = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (String(row[0]) !== String(assetRowIndex)) continue;
      var obj = { _rowIndex: i + 1 };
      headers.forEach(function(h, j) { obj[h] = (row[j] || '').toString().trim(); });
      logs.push(obj);
    }
    logs.sort(function(a, b) { return (b['Date Performed'] || '').localeCompare(a['Date Performed'] || ''); });
    return { logs: logs };
  } catch(e) { return { logs: [], error: e.toString() }; }
}

function addMaintenanceLog(payload) {
  try {
    var auth = authorizeCaller(payload, ['admin', 'office']);
    if (!auth.ok) return { success: false, error: auth.error, code: auth.code };

    var assetRowIndex = payload.assetRowIndex;
    var assetName = payload.assetName || '';
    var values = payload.values || {};
    if (!assetRowIndex) return { success: false, error: 'Missing assetRowIndex' };

    var logSheet = ensureSheetWithHeaders_('Asset Maintenance Log', ASSET_LOG_HEADERS);
    var row = ASSET_LOG_HEADERS.map(function(h) {
      if (h === 'Asset Row') return assetRowIndex;
      if (h === 'Asset Name') return assetName;
      return values[h] !== undefined ? values[h] : '';
    });
    logSheet.appendRow(row);

    var assetUpdates = {};
    if (values['Date Performed']) assetUpdates['Last Service Date'] = values['Date Performed'];
    if (values['Next Due Date'])  assetUpdates['Next Due Date']     = values['Next Due Date'];
    if (Object.keys(assetUpdates).length) {
      var assetSheet = ensureSheetWithHeaders_('Assets', ASSET_HEADERS);
      var assetHeaders = assetSheet.getRange(1, 1, 1, assetSheet.getLastColumn()).getValues()[0];
      assetHeaders.forEach(function(h, i) {
        var key = h.toString().trim();
        if (key && assetUpdates[key] !== undefined) {
          assetSheet.getRange(assetRowIndex, i + 1).setValue(assetUpdates[key]);
        }
      });
    }

    SpreadsheetApp.flush();
    return { success: true };
  } catch(e) { return { success: false, error: e.toString() }; }
}

// ── Material Inventory (log material in / take material out) ─────────────────
// On-hand quantity is not stored -- it's computed by summing In/Out rows from
// the log sheet on every read (same aggregate-on-read approach as
// getJobCostSummary_). Material names are drawn from the existing Pricing
// sheet catalog rather than a separate item list, per Aidan's call.

var MATERIAL_LOG_HEADERS = ['Date', 'Material', 'Unit', 'Type', 'Qty', 'Job / Reference', 'Notes', 'Logged By'];

/** On-hand balances per material plus the most recent log activity. */
function getMaterialInventory(payload) {
  try {
    var auth = authorizeCaller(payload, ['admin', 'office', 'runner']);
    if (!auth.ok) return { materials: [], recentLog: [], error: auth.error, code: auth.code };

    var sheet = ensureSheetWithHeaders_('Material Inventory Log', MATERIAL_LOG_HEADERS);
    var data = sheet.getDataRange().getValues();

    var totals = {}; // material name -> { name, unit, qtyIn, qtyOut }
    var log = [];
    for (var i = 1; i < data.length; i++) {
      var row  = data[i];
      var name = (row[1] || '').toString().trim();
      if (!name) continue;
      var unit = (row[2] || '').toString().trim();
      var type = (row[3] || '').toString().trim();
      var qty  = parseFloat(row[4]) || 0;

      if (!totals[name]) totals[name] = { name: name, unit: unit, qtyIn: 0, qtyOut: 0 };
      if (!totals[name].unit && unit) totals[name].unit = unit;
      if (type === 'In') totals[name].qtyIn += qty;
      else if (type === 'Out') totals[name].qtyOut += qty;

      log.push({
        _rowIndex: i + 1,
        date:      row[0] instanceof Date ? row[0].toISOString() : (row[0] || '').toString(),
        material:  name,
        unit:      unit,
        type:      type,
        qty:       qty,
        jobRef:    (row[5] || '').toString(),
        notes:     (row[6] || '').toString(),
        loggedBy:  (row[7] || '').toString()
      });
    }

    var materials = Object.keys(totals).sort().map(function(name) {
      var t = totals[name];
      return { name: t.name, unit: t.unit, qtyIn: t.qtyIn, qtyOut: t.qtyOut, onHand: t.qtyIn - t.qtyOut, source: 'log' };
    });

    // QBO is authoritative for on-hand quantity on any material mapped to a
    // real Inventory Item, per the locked design -- overrides the Sheets-log
    // aggregate for that material rather than trusting the local In/Out sum,
    // which never sees quantity that moved via a QBO Bill (receiving a stock
    // PO) directly. Best-effort: a QuickBooks hiccup just falls back to
    // whatever the Sheets log already showed, same as before this existed.
    var qboItems = getWarehouseItemsOnHand_();
    qboItems.forEach(function(qi) {
      var key = (qi.name || '').toString().trim().toLowerCase();
      var existing = key && materials.find(function(m) { return m.name.toString().trim().toLowerCase() === key; });
      if (existing) {
        existing.onHand = qi.onHand;
        existing.source = 'quickbooks';
        existing.qboItemId = qi.qboItemId;
        existing.qboItemName = qi.qboItemName;
      } else {
        materials.push({ name: qi.qboItemName, unit: '', qtyIn: null, qtyOut: null, onHand: qi.onHand, source: 'quickbooks', qboItemId: qi.qboItemId, qboItemName: qi.qboItemName });
      }
    });
    materials.sort(function(a, b) { return a.name.localeCompare(b.name); });

    log.sort(function(a, b) { return b._rowIndex - a._rowIndex; });

    return { materials: materials, recentLog: log.slice(0, 50) };
  } catch(e) { return { materials: [], recentLog: [], error: e.toString() }; }
}

/**
 * Pulls stocked material for a job -- the QBO-backed replacement for the
 * old manual "Take Material Out" log entry. Accepts multiple materials in
 * one call (one job per submission, any number of materials) so they land
 * as a single InventoryAdjustment rather than one per material. The actual
 * on-hand check and posting happens in pushMaterialPullToQuickBooks_
 * (QuickBooks_OAuth.gs), which re-reads QBO's live quantity immediately
 * before posting and hard-stops if any line's requested quantity exceeds
 * it -- this function never performs that check itself against a
 * client-supplied number. On success, also appends one audit-trail row per
 * material to the Material Inventory Log sheet -- that sheet is no longer
 * the source of truth for on-hand quantity (QBO is, per the locked
 * design), just a record of who pulled what, when, for which job.
 * payload: {lines: [{qboItemId, qty}], builder, jobRef, notes?}.
 */
function pullMaterialForJob(payload) {
  try {
    var auth = authorizeCaller(payload, ['admin', 'office', 'runner']);
    if (!auth.ok) return { success: false, error: auth.error, code: auth.code };

    var lines   = Array.isArray(payload.lines) ? payload.lines : [];
    var builder = (payload.builder || '').toString().trim();
    var jobRef  = (payload.jobRef || '').toString().trim();
    var notes   = (payload.notes || '').toString().trim();

    lines = lines.map(function(l) {
      return { qboItemId: (l.qboItemId || '').toString().trim(), qty: parseFloat(l.qty) };
    }).filter(function(l) { return l.qboItemId && l.qty > 0; });

    if (!lines.length) return { success: false, error: 'Add at least one material.' };
    if (!builder || !jobRef) return { success: false, error: 'Job is required.' };

    var lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) return { success: false, error: 'Server is busy - try again in a moment.' };
    try {
      // Idempotency guard -- see createPO() for the full rationale. A cache
      // hit here means this exact submission already succeeded (and already
      // posted to QuickBooks), so it must not be re-posted on a client retry.
      var idemKey = (payload.idempotencyKey || '').toString().trim();
      var cache = CacheService.getScriptCache();
      var cacheKey = idemKey ? ('idem_matpull_' + idemKey) : null;
      if (cacheKey) {
        var cached = null;
        try { cached = cache.get(cacheKey); } catch (e) {}
        if (cached) return JSON.parse(cached);
      }

      var customerLookup = getProjectQuickBooksId_(builder, jobRef);
      if (!customerLookup.qbCustomerId) {
        return { success: false, error: 'No QuickBooks Customer/Job Id linked for "' + builder + ' / ' + jobRef + '" -- add it to the Projects sheet before pulling material for this job.' };
      }

      var pullRes = pushMaterialPullToQuickBooks_(lines, customerLookup.qbCustomerId, builder, jobRef);
      if (!pullRes.success) return pullRes;

      var callerName = getRoleByEmail(auth.email).name || auth.email;
      var sheet = ensureSheetWithHeaders_('Material Inventory Log', MATERIAL_LOG_HEADERS);
      var now = new Date();
      var rows = pullRes.lines.map(function(l) {
        return [now, l.materialName, '', 'Out', l.qty, builder + ' / ' + jobRef, notes, callerName];
      });
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, MATERIAL_LOG_HEADERS.length).setValues(rows);
      SpreadsheetApp.flush();

      var result = { success: true, lines: pullRes.lines, cogsAccountName: pullRes.cogsAccountName };
      if (cacheKey) { try { cache.put(cacheKey, JSON.stringify(result), 300); } catch (e) {} }
      return result;
    } finally {
      lock.releaseLock();
    }
  } catch(e) { return { success: false, error: e.toString() }; }
}

/**
 * Returns material to stock from a job at known costs -- for material
 * coming back with no vendor invoice attached (e.g. leftover from a job).
 * Accepts multiple materials in one call so they post as a single Bill +
 * Vendor Credit pair rather than one pair per material. The actual
 * QuickBooks posting happens in pushMaterialReturnToQuickBooks_
 * (QuickBooks_OAuth.gs). Only offers materials already tracked as a QBO
 * Inventory Item (same list the Pull-for-job form uses) -- a material that
 * isn't stocked yet needs to become one via a normal stock PO first, not
 * through this flow. payload: {lines: [{qboItemId, materialName, qty,
 * unitCost}], builder, jobRef, notes?}.
 */
function returnMaterialFromJob(payload) {
  try {
    var auth = authorizeCaller(payload, ['admin', 'office', 'runner']);
    if (!auth.ok) return { success: false, error: auth.error, code: auth.code };

    var lines   = Array.isArray(payload.lines) ? payload.lines : [];
    var builder = (payload.builder || '').toString().trim();
    var jobRef  = (payload.jobRef || '').toString().trim();
    var notes   = (payload.notes || '').toString().trim();

    lines = lines.map(function(l) {
      return {
        qboItemId: (l.qboItemId || '').toString().trim(),
        materialName: (l.materialName || '').toString().trim(),
        qty: parseFloat(l.qty),
        unitCost: parseFloat(l.unitCost)
      };
    }).filter(function(l) { return l.qboItemId && l.qty > 0 && !isNaN(l.unitCost) && l.unitCost >= 0; });

    if (!lines.length) return { success: false, error: 'Add at least one material with a quantity and cost.' };
    if (!builder || !jobRef) return { success: false, error: 'Job is required.' };

    var lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) return { success: false, error: 'Server is busy - try again in a moment.' };
    try {
      // Idempotency guard -- see createPO() for the full rationale. A cache
      // hit here means this exact submission already succeeded (and already
      // posted to QuickBooks), so it must not be re-posted on a client retry.
      var idemKey = (payload.idempotencyKey || '').toString().trim();
      var cache = CacheService.getScriptCache();
      var cacheKey = idemKey ? ('idem_matreturn_' + idemKey) : null;
      if (cacheKey) {
        var cached = null;
        try { cached = cache.get(cacheKey); } catch (e) {}
        if (cached) return JSON.parse(cached);
      }

      var customerLookup = getProjectQuickBooksId_(builder, jobRef);
      if (!customerLookup.qbCustomerId) {
        return { success: false, error: 'No QuickBooks Customer/Job Id linked for "' + builder + ' / ' + jobRef + '" -- add it to the Projects sheet before crediting a return from this job.' };
      }

      var returnRes = pushMaterialReturnToQuickBooks_(lines, customerLookup.qbCustomerId, builder, jobRef);
      if (!returnRes.success) return returnRes;

      var callerName = getRoleByEmail(auth.email).name || auth.email;
      var sheet = ensureSheetWithHeaders_('Material Inventory Log', MATERIAL_LOG_HEADERS);
      var now = new Date();
      var rows = lines.map(function(l) {
        return [now, l.materialName, '', 'In', l.qty, builder + ' / ' + jobRef, notes, callerName];
      });
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, MATERIAL_LOG_HEADERS.length).setValues(rows);
      SpreadsheetApp.flush();

      var result = { success: true, amount: returnRes.amount, billId: returnRes.billId, vendorCreditId: returnRes.vendorCreditId, creditWarning: returnRes.creditWarning };
      if (cacheKey) { try { cache.put(cacheKey, JSON.stringify(result), 300); } catch (e) {} }
      return result;
    } finally {
      lock.releaseLock();
    }
  } catch(e) { return { success: false, error: e.toString() }; }
}

/** Removes a mistaken log entry. Admin-only -- deleting alters historical on-hand math. */
function deleteMaterialLogEntry(payload) {
  try {
    var auth = authorizeCaller(payload, ['admin']);
    if (!auth.ok) return { success: false, error: auth.error, code: auth.code };

    var rowIndex = payload.rowIndex;
    if (!rowIndex) return { success: false, error: 'Missing rowIndex' };

    var sheet = ensureSheetWithHeaders_('Material Inventory Log', MATERIAL_LOG_HEADERS);
    if (rowIndex < 2 || rowIndex > sheet.getLastRow()) return { success: false, error: 'Log entry not found' };
    sheet.deleteRow(rowIndex);
    SpreadsheetApp.flush();
    return { success: true };
  } catch(e) { return { success: false, error: e.toString() }; }
}

// ── Internal Invoices / Credit Memos (material-to-job costing) ────────────
// Material leaving the warehouse for a job is invoiced to that job; material
// coming back is credited. These are internal-only documents -- nothing is
// sent to QuickBooks or any outside party. PDF generation reuses the same
// approach as getPayrollPdf(): render an HTML string via HtmlService and
// hand back base64 for direct download.

var INTERNAL_INVOICE_SHEET_NAME = 'Internal Invoices';
var INTERNAL_INVOICE_HEADERS = ['Doc Number', 'Type', 'Date', 'Job / Reference', 'Material', 'Unit', 'Qty', 'Unit Price', 'Amount', 'Material Log Row', 'Created By'];

var COMPANY_NAME = 'Panoramic Building LLC';
var COMPANY_ADDRESS_LINE1 = '1460 N Stonecrest Ln';
var COMPANY_ADDRESS_LINE2 = 'Logan, UT 84341';
var COMPANY_PHONE = '801-689-3553';
var COMPANY_LOGO_URL = 'https://ops.panoramicbuildingllc.com/panoramic-logo.png';

/** Looks up a material's Best Price (Pricing sheet column C) by exact description match. Returns 0 if not found. */
function getMaterialUnitPrice_(materialDescription) {
  try {
    var want = (materialDescription || '').toString().trim().toLowerCase();
    if (!want) return 0;
    var data = getPricingSheetRaw_().data; // A=Description, B=U/M, C=Best Price
    for (var i = 0; i < data.length; i++) {
      var desc = (data[i][0] || '').toString().trim().toLowerCase();
      if (desc === want) return parseFloat(data[i][2]) || 0;
    }
    return 0;
  } catch (e) { return 0; }
}

/** Next sequential doc number for a type, e.g. 'INV-0001' / 'CM-0001'. Scans existing rows -- fine at this volume, same no-extra-state approach the PO numbering already uses. */
function getNextInternalInvoiceNumber_(sheet, type) {
  var prefix = type === 'Credit Memo' ? 'CM-' : 'INV-';
  var lastRow = sheet.getLastRow();
  var maxN = 0;
  if (lastRow >= 2) {
    var rows = sheet.getRange(2, 1, lastRow - 1, 2).getValues(); // Doc Number, Type
    rows.forEach(function(r) {
      if (r[1] !== type) return;
      var n = parseInt((r[0] || '').toString().replace(prefix, ''), 10);
      if (!isNaN(n) && n > maxN) maxN = n;
    });
  }
  return prefix + ('0000' + (maxN + 1)).slice(-4);
}

/** Builds the invoice/credit-memo PDF HTML, styled after the paper vendor invoice this internal doc replaces. */
function buildInternalInvoiceHtml_(doc) {
  var esc = function(s) { return (s == null ? '' : s.toString()).replace(/[&<>]/g, function(c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]; }); };
  var titleColor = doc.type === 'Credit Memo' ? '#b02a2a' : '#1a1a1a';
  var amountLabel = doc.type === 'Credit Memo' ? 'Credit' : 'Amount';

  return '<!DOCTYPE html><html><head><meta charset="utf-8"><style>' +
    'body{font-family:Arial,Helvetica,sans-serif;font-size:12px;color:#222;margin:32px}' +
    '.header{display:flex;justify-content:space-between;align-items:flex-start}' +
    '.company img{max-width:220px;max-height:70px}' +
    '.company .name{font-weight:bold;font-size:14px;margin-top:4px}' +
    '.doctitle{font-size:26px;font-weight:bold;color:' + titleColor + ';text-align:right}' +
    'table.meta{border-collapse:collapse;margin-top:8px}' +
    'table.meta th, table.meta td{border:1px solid #333;padding:5px 12px;font-size:11px}' +
    'table.meta th{background:#f4f4f4;text-align:left}' +
    '.billto{border:1px solid #333;padding:10px 14px;width:280px;margin-top:24px}' +
    '.billto .label{font-weight:bold;border-bottom:1px solid #333;margin:-10px -14px 8px;padding:6px 14px;background:#f4f4f4}' +
    'table.items{width:100%;border-collapse:collapse;margin-top:24px}' +
    'table.items th{background:#f4f4f4;text-align:left;padding:6px 8px;font-size:11px;border:1px solid #333}' +
    'table.items td{padding:6px 8px;border:1px solid #333;font-size:11px}' +
    '.num{text-align:right}' +
    '.totals{margin-top:12px;text-align:right;font-size:13px;font-weight:bold}' +
    '</style></head><body>' +
    '<div class="header">' +
      '<div class="company"><img src="' + COMPANY_LOGO_URL + '"><div class="name">' + esc(COMPANY_NAME) + '</div>' +
        '<div>' + esc(COMPANY_ADDRESS_LINE1) + '</div><div>' + esc(COMPANY_ADDRESS_LINE2) + '</div><div>' + esc(COMPANY_PHONE) + '</div></div>' +
      '<div><div class="doctitle">' + esc(doc.type) + '</div>' +
        '<table class="meta"><tr><th>Date</th><th>Doc #</th></tr><tr><td>' + esc(doc.dateStr) + '</td><td>' + esc(doc.docNumber) + '</td></tr></table></div>' +
    '</div>' +
    '<div class="billto"><div class="label">Job / Reference</div>' + esc(doc.jobRef) + '</div>' +
    '<table class="items"><tr><th>Item</th><th>Description</th><th class="num">Quantity</th><th class="num">Price Each</th><th class="num">' + amountLabel + '</th></tr>' +
    '<tr><td>' + esc(doc.material) + '</td><td>' + esc(doc.material) + '</td><td class="num">' + doc.qty + ' ' + esc(doc.unit) + '</td>' +
    '<td class="num">' + doc.unitPrice.toFixed(2) + '</td><td class="num">' + doc.amount.toFixed(2) + '</td></tr>' +
    '</table>' +
    '<div class="totals">Total ' + amountLabel + ': $' + doc.amount.toFixed(2) + '</div>' +
    (doc.notes ? '<div style="margin-top:16px;font-size:11px;color:#555">Notes: ' + esc(doc.notes) + '</div>' : '') +
    '</body></html>';
}

/**
 * Creates one Internal Invoice or Credit Memo for a single material log
 * transaction: writes a row to the "Internal Invoices" sheet and returns the
 * generated PDF as base64. Called from logMaterialTransaction() right after
 * the log row is appended.
 */
function createInternalInvoiceDoc_(type, materialRow, callerName) {
  var sheet = ensureSheetWithHeaders_(INTERNAL_INVOICE_SHEET_NAME, INTERNAL_INVOICE_HEADERS);
  var docNumber = getNextInternalInvoiceNumber_(sheet, type);
  var unitPrice = getMaterialUnitPrice_(materialRow.material);
  var amount = Math.round(unitPrice * materialRow.qty * 100) / 100;
  var now = new Date();
  var tz = Session.getScriptTimeZone();
  var dateStr = Utilities.formatDate(now, tz, 'M/d/yyyy');

  sheet.appendRow([docNumber, type, now, materialRow.jobRef, materialRow.material, materialRow.unit, materialRow.qty, unitPrice, amount, materialRow.logRowIndex, callerName]);

  var html = buildInternalInvoiceHtml_({
    type: type, docNumber: docNumber, dateStr: dateStr, jobRef: materialRow.jobRef,
    material: materialRow.material, unit: materialRow.unit, qty: materialRow.qty,
    unitPrice: unitPrice, amount: amount, notes: materialRow.notes
  });
  var pdfBlob = HtmlService.createHtmlOutput(html).getAs('application/pdf');
  var filename = docNumber + ' - ' + materialRow.jobRef + '.pdf';
  return { docNumber: docNumber, type: type, amount: amount, filename: filename, base64: Utilities.base64Encode(pdfBlob.getBytes()) };
}

/** Recent internal invoices/credit memos, most recent first. */
function getInternalInvoices(payload) {
  try {
    var auth = authorizeCaller(payload, ['admin', 'office', 'site_manager', 'runner']);
    if (!auth.ok) return { docs: [], error: auth.error, code: auth.code };

    var sheet = ensureSheetWithHeaders_(INTERNAL_INVOICE_SHEET_NAME, INTERNAL_INVOICE_HEADERS);
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return { docs: [] };

    var docs = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      docs.push({
        _rowIndex: i + 1,
        docNumber: row[0], type: row[1],
        date: row[2] instanceof Date ? row[2].toISOString() : (row[2] || '').toString(),
        jobRef: row[3], material: row[4], unit: row[5], qty: row[6], unitPrice: row[7], amount: row[8],
        materialLogRow: row[9], createdBy: row[10]
      });
    }
    docs.sort(function(a, b) { return b._rowIndex - a._rowIndex; });
    return { docs: docs.slice(0, 50) };
  } catch (e) { return { docs: [], error: e.toString() }; }
}

/** Regenerates the PDF for an existing doc number (reprint). */
function getInternalInvoicePdf(payload) {
  try {
    var auth = authorizeCaller(payload, ['admin', 'office', 'site_manager', 'runner']);
    if (!auth.ok) return { error: auth.error, code: auth.code };

    var docNumber = (payload.docNumber || '').toString().trim();
    if (!docNumber) return { error: 'Missing docNumber' };

    var sheet = ensureSheetWithHeaders_(INTERNAL_INVOICE_SHEET_NAME, INTERNAL_INVOICE_HEADERS);
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if ((row[0] || '').toString() !== docNumber) continue;
      var tz = Session.getScriptTimeZone();
      var html = buildInternalInvoiceHtml_({
        type: row[1], docNumber: row[0], dateStr: Utilities.formatDate(row[2], tz, 'M/d/yyyy'),
        jobRef: row[3], material: row[4], unit: row[5], qty: row[6], unitPrice: row[7], amount: row[8], notes: ''
      });
      var pdfBlob = HtmlService.createHtmlOutput(html).getAs('application/pdf');
      return { success: true, filename: docNumber + ' - ' + row[3] + '.pdf', base64: Utilities.base64Encode(pdfBlob.getBytes()) };
    }
    return { error: 'Document not found' };
  } catch (e) { return { error: e.toString() }; }
}

// ── Portal Storage (website credential vault) ────────────────────────────────
// Stores login credentials for external websites/portals. Passwords are
// AES-256-CBC encrypted (see Aes_Helper.gs) before ever touching the sheet --
// unlike the plain-text HR password column, this is a purpose-built vault.

var PORTAL_STORAGE_ROLES = ['admin', 'office', 'human_resources'];
var PORTAL_STORAGE_HEADERS = ['Site Name', 'Link', 'Username', 'Password', 'Notes', 'Added By', 'Date Added', 'Last Updated By', 'Last Updated'];

/** Full credential list, passwords decrypted server-side for display. */
function getPortalCredentials(payload) {
  try {
    var auth = authorizeCaller(payload, PORTAL_STORAGE_ROLES);
    if (!auth.ok) return { credentials: [], error: auth.error, code: auth.code };

    var keyBytes = getPortalStorageKeyBytes_();
    var sheet = ensureSheetWithHeaders_('Portal Storage', PORTAL_STORAGE_HEADERS);
    var data = sheet.getDataRange().getValues();
    var credentials = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var siteName = (row[0] || '').toString().trim();
      if (!siteName) continue;
      var password = '';
      try { password = row[3] ? aesDecrypt_(row[3].toString(), keyBytes) : ''; } catch (e) { password = ''; }
      credentials.push({
        _rowIndex: i + 1,
        siteName: siteName,
        link: (row[1] || '').toString(),
        username: (row[2] || '').toString(),
        password: password,
        notes: (row[4] || '').toString(),
        addedBy: (row[5] || '').toString(),
        dateAdded: row[6] instanceof Date ? row[6].toISOString() : (row[6] || '').toString(),
        updatedBy: (row[7] || '').toString(),
        lastUpdated: row[8] instanceof Date ? row[8].toISOString() : (row[8] || '').toString()
      });
    }
    credentials.sort(function(a, b) { return a.siteName.localeCompare(b.siteName); });
    return { credentials: credentials };
  } catch(e) { return { credentials: [], error: e.toString() }; }
}

/** Adds a new credential entry. */
function savePortalCredential(payload) {
  try {
    var auth = authorizeCaller(payload, PORTAL_STORAGE_ROLES);
    if (!auth.ok) return { success: false, error: auth.error, code: auth.code };

    var siteName = (payload.siteName || '').toString().trim();
    var link     = (payload.link || '').toString().trim();
    var username = (payload.username || '').toString().trim();
    var password = (payload.password || '').toString();
    var notes    = (payload.notes || '').toString().trim();

    if (!siteName) return { success: false, error: 'Site name is required' };
    if (!username) return { success: false, error: 'Username is required' };
    if (!password) return { success: false, error: 'Password is required' };

    var lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) return { success: false, error: 'Server is busy - try again in a moment.' };
    try {
      var idemKey = (payload.idempotencyKey || '').toString().trim();
      var cache = CacheService.getScriptCache();
      var cacheKey = idemKey ? ('idem_portalsave_' + idemKey) : null;
      if (cacheKey) {
        var cached = null;
        try { cached = cache.get(cacheKey); } catch (e) {}
        if (cached) return JSON.parse(cached);
      }

      var callerName = getRoleByEmail(auth.email).name || auth.email;
      var keyBytes = getPortalStorageKeyBytes_();
      var encrypted = aesEncrypt_(password, keyBytes);
      var now = new Date();

      var sheet = ensureSheetWithHeaders_('Portal Storage', PORTAL_STORAGE_HEADERS);
      sheet.appendRow([siteName, link, username, encrypted, notes, callerName, now, callerName, now]);
      SpreadsheetApp.flush();

      var result = { success: true, rowIndex: sheet.getLastRow() };
      if (cacheKey) { try { cache.put(cacheKey, JSON.stringify(result), 300); } catch (e) {} }
      return result;
    } finally {
      lock.releaseLock();
    }
  } catch(e) { return { success: false, error: e.toString() }; }
}

/** Updates an existing credential entry (password only re-encrypted if changed). */
function updatePortalCredential(payload) {
  try {
    var auth = authorizeCaller(payload, PORTAL_STORAGE_ROLES);
    if (!auth.ok) return { success: false, error: auth.error, code: auth.code };

    var rowIndex = payload.rowIndex;
    if (!rowIndex) return { success: false, error: 'Missing rowIndex' };

    var siteName = (payload.siteName || '').toString().trim();
    var link     = (payload.link || '').toString().trim();
    var username = (payload.username || '').toString().trim();
    var password = (payload.password || '').toString();
    var notes    = (payload.notes || '').toString().trim();

    if (!siteName) return { success: false, error: 'Site name is required' };
    if (!username) return { success: false, error: 'Username is required' };

    var lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) return { success: false, error: 'Server is busy - try again in a moment.' };
    try {
      var sheet = ensureSheetWithHeaders_('Portal Storage', PORTAL_STORAGE_HEADERS);
      if (rowIndex < 2 || rowIndex > sheet.getLastRow()) return { success: false, error: 'Credential not found' };

      var callerName = getRoleByEmail(auth.email).name || auth.email;
      var now = new Date();

      var encryptedPassword;
      if (password) {
        encryptedPassword = aesEncrypt_(password, getPortalStorageKeyBytes_());
      } else {
        encryptedPassword = sheet.getRange(rowIndex, 4).getValue(); // unchanged
      }

      sheet.getRange(rowIndex, 1, 1, 9).setValues([[
        siteName, link, username, encryptedPassword, notes,
        sheet.getRange(rowIndex, 6).getValue(), sheet.getRange(rowIndex, 7).getValue(),
        callerName, now
      ]]);
      SpreadsheetApp.flush();
      return { success: true };
    } finally {
      lock.releaseLock();
    }
  } catch(e) { return { success: false, error: e.toString() }; }
}

/** Removes a credential entry. Admin-only. */
function deletePortalCredential(payload) {
  try {
    var auth = authorizeCaller(payload, ['admin']);
    if (!auth.ok) return { success: false, error: auth.error, code: auth.code };

    var rowIndex = payload.rowIndex;
    if (!rowIndex) return { success: false, error: 'Missing rowIndex' };

    var sheet = ensureSheetWithHeaders_('Portal Storage', PORTAL_STORAGE_HEADERS);
    if (rowIndex < 2 || rowIndex > sheet.getLastRow()) return { success: false, error: 'Credential not found' };
    sheet.deleteRow(rowIndex);
    SpreadsheetApp.flush();
    return { success: true };
  } catch(e) { return { success: false, error: e.toString() }; }
}

// ── Reconcile Statement ───────────────────────────────────────────────────────
function reconcileStatement(payload) {
  try {
    var auth = authorizeCaller(payload, ['admin', 'office']);
    if (!auth.ok) return { error: auth.error, code: auth.code };

    var invoiceNumbers = payload.invoiceNumbers;

    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return { error: 'PO Database sheet not found' };
    var data = sheet.getDataRange().getValues();
    var colPoNum   = 0;  // A
    var colJob     = 3;  // D
    var colVendor  = 4;  // E
    var colInvoice = 5;  // F - Vendor Invoice
    var colStatus  = 6;  // G
    var dbMap = {};
    for (var i = 1; i < data.length; i++) {
      var inv = (data[i][colInvoice] || '').toString().trim();
      if (!inv) continue;
      dbMap[inv.toLowerCase()] = {
        poNum:  data[i][colPoNum],
        vendor: data[i][colVendor],
        job:    data[i][colJob],
        status: data[i][colStatus],
        invNum: inv
      };
    }
    var matched = [], unmatched = [];
    (invoiceNumbers || []).forEach(function(inv) {
      var key = inv.toString().trim().toLowerCase();
      var found = dbMap[key];
      if (!found) {
        var keys = Object.keys(dbMap);
        for (var k = 0; k < keys.length; k++) {
          if (keys[k].indexOf(key) === 0 || key.indexOf(keys[k]) === 0) {
            found = dbMap[keys[k]]; break;
          }
        }
      }
      if (found) matched.push({ invoiceNumber: inv, poNum: found.poNum, vendor: found.vendor, job: found.job, status: found.status });
      else unmatched.push(inv);
    });
    return { success: true, matched: matched, unmatched: unmatched };
  } catch(e) { return { error: e.toString() }; }
}

// ── Job List ─────────────────────────────────────────────────────────────────
function getJobList() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) return { error: 'PO Database sheet not found' };
    var data = sheet.getDataRange().getValues();
    var jobs = {};
    for (var i = 0; i < data.length; i++) {
      if (!isValidPONumber((data[i][0] || '').toString().trim())) continue;
      var job = (data[i][3] || '').toString().trim();
      if (job) jobs[job] = true;
    }
    return { success: true, jobs: Object.keys(jobs).sort() };
  } catch(e) { return { error: e.toString() }; }
}

// ── Job Cost Summary ──────────────────────────────────────────────────────────
/** Public, authorized entry point for the client. */
function getJobCostSummary(payload) {
  var auth = authorizeCaller(payload, ['admin']);
  if (!auth.ok) return { error: auth.error, code: auth.code };
  return getJobCostSummary_(payload.jobRef);
}

/** Unauthenticated helper - only call from other server-side functions that have already authorized the caller. */
function getJobCostSummary_(jobRef) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) return { error: 'PO Database sheet not found' };
    var data = sheet.getDataRange().getValues();
    var rows = [], totalSpend = 0;
    var target = (jobRef || '').toString().trim().toLowerCase();
    for (var i = 0; i < data.length; i++) {
      if (!isValidPONumber((data[i][0] || '').toString().trim())) continue;
      var job = (data[i][3] || '').toString().trim();
      if (job.toLowerCase() !== target) continue;
      var total = parseFloat(data[i][7]) || 0;
      totalSpend += total;
      rows.push({
        poNum:      data[i][0],
        dateIssued: data[i][1] instanceof Date ? Utilities.formatDate(data[i][1], Session.getScriptTimeZone(), 'MM/dd/yy') : '',
        vendor:     data[i][4],
        invoiceNum: data[i][5],
        status:     data[i][6],
        total:      total
      });
    }
    return { success: true, rows: rows, totalSpend: totalSpend };
  } catch(e) { return { error: e.toString() }; }
}

/**
 * Sums this job's actual QuickBooks Bill totals so the Job Dashboard can be
 * checked against the PO Database's own total (cost.totalSpend) for drift --
 * a Bill edited or voided directly in QuickBooks after posting would go
 * unnoticed otherwise. Re-queries QBO live for each Bill's *current*
 * TotalAmt rather than trusting the Purchase Line Item Log, which is only a
 * point-in-time copy taken at post time.
 *
 * Bill Ids come from the QB Invoice Staging sheet's per-PO record (one row
 * per invoice review; 'QB Bill Id' is set once createQuickBooksBill posts
 * it) filtered to this job's Job Ref. One batched query covers every Bill
 * Id at once via QBO's query language IN-list rather than N separate calls.
 *
 * Unauthenticated helper -- only called from getJobDashboard, which has
 * already authorized the caller. Returns { connected:false } if QuickBooks
 * was never connected, so the dashboard can just hide the comparison.
 */
function getQuickBooksJobTotal_(jobRef) {
  try {
    if (!getQuickBooksService_().hasAccess()) return { connected: false };

    var target = (jobRef || '').toString().trim().toLowerCase();
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(QB_STAGING_SHEET);
    if (!sheet) return { connected: true, total: 0, billCount: 0, poCount: 0 };
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { connected: true, total: 0, billCount: 0, poCount: 0 };

    var data = sheet.getRange(2, 1, lastRow - 1, QB_STAGING_HEADERS.length).getValues();
    var billIds = {}, poCount = 0;
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var rowJob = (row[QB_STAGING_COL['Job Ref']] || '').toString().trim().toLowerCase();
      if (rowJob !== target) continue;
      var billId = (row[QB_STAGING_COL['QB Bill Id']] || '').toString().trim();
      if (!billId) continue;
      billIds[billId] = true;
      poCount++;
    }
    var ids = Object.keys(billIds);
    if (!ids.length) return { connected: true, total: 0, billCount: 0, poCount: 0 };

    var query = "SELECT Id, TotalAmt FROM Bill WHERE Id IN (" +
      ids.map(function(id) { return "'" + id.replace(/'/g, "") + "'"; }).join(',') + ")";
    var res = quickbooksApiGet_('/query?query=' + encodeURIComponent(query));
    if (!res.success) return { connected: true, error: res.error };

    var bills = (res.data && res.data.QueryResponse && res.data.QueryResponse.Bill) || [];
    var total = 0;
    bills.forEach(function(b) { total += parseFloat(b.TotalAmt) || 0; });

    // billCount < expectedBillCount means a Bill this app posted no longer
    // exists in QuickBooks (voided/deleted there directly) -- surfaced to
    // the caller rather than silently under-totaling.
    return { connected: true, total: total, billCount: bills.length, expectedBillCount: ids.length, poCount: poCount };
  } catch (e) {
    return { connected: true, error: e.toString() };
  }
}

// ── Missing Invoices ──────────────────────────────────────────────────────────
/** Public, authorized entry point for the client. */
function getMissingInvoices(payload) {
  var auth = authorizeCaller(payload, ['admin']);
  if (!auth.ok) return { error: auth.error, code: auth.code };
  return getMissingInvoices_();
}

/** Unauthenticated helper - only call from other server-side functions that have already authorized the caller. */
function getMissingInvoices_() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) return { error: 'PO Database sheet not found' };
    var data = sheet.getDataRange().getValues();
    var missing = [];
    // Statuses where we don't yet expect an invoice
    var skipStatuses = { 'draft': true, 'ordered': true, 'being made': true,
                         'pending pickup': true, 'pending delivery': true,
                         'pending delivery to supplier': true, 'currently picking up': true,
                         'canceled': true };
    for (var i = 0; i < data.length; i++) {
      var poNum = (data[i][0] || '').toString().trim();
      if (!isValidPONumber(poNum)) continue;
      var status  = (data[i][6] || '').toString().trim();
      var invoice = (data[i][5] || '').toString().trim();
      if (skipStatuses[status.toLowerCase()]) continue;
      if (!invoice) {
        missing.push({
          poNum:      poNum,
          dateIssued: data[i][1] instanceof Date ? Utilities.formatDate(data[i][1], Session.getScriptTimeZone(), 'MM/dd/yy') : '',
          vendor:     data[i][4],
          job:        data[i][3],
          status:     status
        });
      }
    }
    return { success: true, missing: missing };
  } catch(e) { return { error: e.toString() }; }
}

// ── Job Dashboard (Jobs Registry + Quality Walk log) ──────────────────────────
// The "Projects" sheet's A-D columns (Contractor, Job Name, Drive folder,
// Asana GID) are the only ones this feature still reads -- getProjectFolderId/
// getRecentJobs only ever read that same fixed-width range, so nothing else
// touches it. Status/Start Date/End Date are NOT stored on this sheet at all;
// they're read live from Asana (see getJobDashboard).

/**
 * Best-effort split of an Asana task name formatted "Builder, Job Name,
 * Address" (the convention createProjectAndTask writes) down to the bare
 * Job Name segment, for matching against PO Database jobRef / Projects
 * sheet Job Name (neither of which know about Asana's longer task names).
 * Falls back to the raw string for tasks that don't follow the convention
 * (manually created Asana tasks, older jobs, etc.) -- cost/invoice lookups
 * for those will simply come back empty rather than erroring.
 */
function parseAsanaJobName_(rawTaskName) {
  var s = (rawTaskName || '').toString().trim();
  var parts = s.split(',');
  return parts.length >= 2 ? parts[1].trim() : s;
}

/**
 * Combined payload for the Job Dashboard panel: job meta, cost summary
 * (reuses getJobCostSummary as-is, spend only), missing-invoice count, and
 * quality-walk history. One round trip, matching this app's existing
 * one-action-per-panel convention.
 *
 * The job picker is sourced from Asana (getAsanaJobs, same list Quality
 * Check already uses) -- payload.jobGid is the Asana task gid (primary
 * key), payload.jobName is that task's full "Builder, Job Name, Address"
 * display string. The "Projects" sheet (Contractor, Drive folder) is
 * looked up by exact Asana Task GID match (column D) when the job has
 * been linked via the New Project intake flow; otherwise meta falls back
 * to a parsed short name with no Drive folder (nowhere to look one up).
 * Status, Start Date, and End Date all come live from the job's Asana
 * task itself: Status is whichever section/bucket the task currently sits
 * in on the Exterior Master Schedule board (ASANA_EXT_SCHED) -- the same
 * board getAsanaJobs sources the job picker from, so every job here is
 * guaranteed to be a task on it -- and Start/End Date are the task's own
 * start_on/due_on fields. None of the three are stored anywhere in this
 * app; moving a task's bucket or dates in Asana is reflected immediately.
 * Cost/invoice matching uses the resolved short job name; quality-walk
 * history is read live from Asana (getQualityWalkHistory_) -- submitQualityCheck
 * writes each check as a subtask of the job's Asana task, so there's
 * nothing to look up by name at all.
 */
function getJobDashboard(payload) {
  try {
    var auth = authorizeCaller(payload, ['admin', 'office', 'site_manager']);
    if (!auth.ok) return { error: auth.error, code: auth.code };

    var jobGid = (payload.jobGid || '').toString().trim();
    if (!jobGid) return { error: 'A job must be selected from the list.' };
    var rawJobName = (payload.jobName || '').toString().trim();

    var meta = null;
    var pSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PROJECTS_SHEET_NAME);
    if (pSheet) {
      var pLastRow = pSheet.getLastRow();
      if (pLastRow >= 2) {
        var pData = pSheet.getRange(2, 1, pLastRow - 1, 4).getValues();
        for (var i = 0; i < pData.length; i++) {
          var rowGid = (pData[i][3] || '').toString().trim();
          if (!rowGid || rowGid !== jobGid) continue;
          meta = {
            contractor:    (pData[i][0] || '').toString().trim(),
            jobName:       (pData[i][1] || '').toString().trim(),
            driveFolderId: extractDriveFolderId(pData[i][2]),
            asanaTaskGid:  rowGid
          };
          break;
        }
      }
    }
    var shortJobName = meta ? meta.jobName : parseAsanaJobName_(rawJobName);
    if (!meta) {
      meta = { contractor: '', jobName: shortJobName, driveFolderId: null, asanaTaskGid: jobGid };
    }
    var shortJobNameLower = shortJobName.toLowerCase();

    var taskInfo = asanaRequest('get', '/tasks/' + jobGid +
      '?opt_fields=start_on,due_on,memberships.section.name,memberships.project.gid');
    meta.startDate = (taskInfo.data && taskInfo.data.start_on) || '';
    meta.endDate   = (taskInfo.data && taskInfo.data.due_on)   || '';
    var membership = ((taskInfo.data && taskInfo.data.memberships) || []).filter(function(m) {
      return m.project && m.project.gid === ASANA_EXT_SCHED;
    })[0];
    meta.status = (membership && membership.section && membership.section.name) || '';

    var cost = getJobCostSummary_(shortJobName);
    var qb = getQuickBooksJobTotal_(shortJobName);

    var missingAll = getMissingInvoices_();
    var missingRows = [];
    if (missingAll.missing) {
      missingRows = missingAll.missing.filter(function(m) {
        return (m.job || '').toString().trim().toLowerCase() === shortJobNameLower;
      });
    }

    var quality = getQualityWalkHistory_(jobGid);

    return { success: true, meta: meta, cost: cost, qb: qb, missingCount: missingRows.length, missingRows: missingRows, quality: quality };
  } catch (e) { return { error: e.toString() }; }
}

// ── Vendor Spend ──────────────────────────────────────────────────────────────
function getVendorSpend(payload) {
  try {
    var auth = authorizeCaller(payload, ['admin', 'office']);
    if (!auth.ok) return { error: auth.error, code: auth.code };

    var startDate = payload.startDate;
    var endDate   = payload.endDate;
    var wantTrend = !!payload.includeTrend;

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) return { error: 'PO Database sheet not found' };
    var data = sheet.getDataRange().getValues();
    var tz    = Session.getScriptTimeZone();
    var start = startDate ? new Date(startDate + 'T00:00:00') : null;
    var end   = endDate   ? new Date(endDate   + 'T23:59:59') : null;
    var vendors = {}, grandTotal = 0, vendorRows = {};
    var monthTotals = {};
    var vendorMonthTotals = {};
    for (var i = 0; i < data.length; i++) {
      if (!isValidPONumber((data[i][0] || '').toString().trim())) continue;
      var vendor = (data[i][4] || '').toString().trim();
      var total  = parseFloat(data[i][7]) || 0;
      if (!vendor || total === 0) continue;
      var d = data[i][1] instanceof Date ? data[i][1] : null;
      if (start || end) {
        if (!d || isNaN(d.getTime())) continue;
        if (start && d < start) continue;
        if (end   && d > end)   continue;
      }
      vendors[vendor] = (vendors[vendor] || 0) + total;
      grandTotal += total;
      // Track top rows per vendor for debugging
      if (!vendorRows[vendor]) vendorRows[vendor] = [];
      vendorRows[vendor].push({ poNum: data[i][0], total: total, row: i + 1 });

      if (wantTrend && d && !isNaN(d.getTime())) {
        var mk = Utilities.formatDate(d, tz, 'yyyy-MM');
        monthTotals[mk] = (monthTotals[mk] || 0) + total;
        if (!vendorMonthTotals[vendor]) vendorMonthTotals[vendor] = {};
        vendorMonthTotals[vendor][mk] = (vendorMonthTotals[vendor][mk] || 0) + total;
      }
    }
    var result = Object.keys(vendors).map(function(v) {
      var rows = (vendorRows[v] || []).sort(function(a,b){return b.total-a.total;}).slice(0,3);
      return { vendor: v, total: vendors[v], topRows: rows };
    }).sort(function(a, b) { return b.total - a.total; });

    var out = { success: true, vendors: result, grandTotal: grandTotal, gasVersion: 4 };

    if (wantTrend) {
      var months = Object.keys(monthTotals).sort();
      out.trend = {
        months: months,
        overall: months.map(function(m) { return monthTotals[m] || 0; }),
        vendors: result.map(function(v) {
          var byMonth = vendorMonthTotals[v.vendor] || {};
          return { vendor: v.vendor, values: months.map(function(m) { return byMonth[m] || 0; }) };
        })
      };
    }

    return out;
  } catch(e) { return { error: e.toString() }; }
}


// ─── Material Report ─────────────────────────────────────────────────────────
function categorizeInvoices(payload) {
  var auth = authorizeCaller(payload, ['admin', 'office']);
  if (!auth.ok) return { error: auth.error, code: auth.code };
  var apiKey = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');
  if (!apiKey) return { error: 'CLAUDE_API_KEY not set in Script Properties' };

  var systemPrompt = [
    'You are a building materials invoice categorizer for Panoramic Building LLC, an exterior siding contractor in Utah.',
    '',
    'CATEGORIES - use exactly these names:',
    '  Siding Lap      : LP SmartSide lap siding (3/8x8x16), 5/4 cedar trim boards',
    '  Siding B&B      : LP SmartSide panels 4x10 (any groove), battens 19/32x3, 4/4 cedar trim boards - only panels used as WALL SIDING, not wrap',
    '  Siding Flashing : Panel Union Flashing, Z-flashing, brick flashing angles/strips',
    '  Metal           : Coil stock, touch-up paint, metal accessories (non-soffit/fascia)',
    '  Soffit & Fascia : Aluminum soffit panels (solid or vented), fascia trim, J-channel, drip edge, coil wrap',
    '  Masonry         : Stone, brick, Lueders, building paper, metal lath, mortar (Type S/N), pallet charges from masonry vendors, lime',
    '  Vinyl           : Vinyl lap or board-and-batten siding panels (any color)',
    '  Vinyl Accessories    : Vinyl starter/finish strips, outside corners, J-channel for vinyl, outlet boxes, light boxes',
    '  Stucco          : Stucco base/finish coat, dryvit, mesh, stucco accessories',
    '  Angle Iron      : Steel angle iron, wide flange beams, structural steel, plasma cutting, steel delivery',
    '  Beam/Post/Garage Wrap : Hardboard/B&B panels used specifically for wrapping beams, posts, columns, or garage openings (NOT wall siding). If B&B panels are ordered and some are clearly for wrapping, classify those here.',
    '',
    'IMPORTANT: Do NOT assign a category. Return an empty string "" for the category field on every line item.',
    'Your job is ONLY to extract and structure the line items with correct amounts, tax shares, and shipping shares.',
    'The user will assign categories themselves.',
    '',
    'INPUT: JSON array of invoice objects, each with fileName and text (raw PDF text, may be messy).',
    '',
    'OUTPUT: Return ONLY a valid JSON array - no prose, no markdown fences. Each element:',
    '{',
    '  "fileName": "...",',
    '  "invoiceNum": "...",',
    '  "vendor": "...",',
    '  "subtotal": 0.00,',
    '  "tax": 0.00,',
    '  "shipping": 0.00,',
    '  "invoiceTotal": 0.00,',
    '  "lineItems": [',
    '    {',
    '      "description": "...",',
    '      "qty": 0,',
    '      "unit": "SqF",',
    '      "amount": 0.00,',
    '      "category": "",',
    '      "taxShare": 0.00,',
    '      "shippingShare": 0.00,',
    '      "uncertain": false',
    '    }',
    '  ],',
    '  "lineItemsSum": 0.00,',
    '  "balanceCheck": true,',
    '  "notes": ""',
    '}',
    '',
    'RULES:',
    '1. Extract invoice number, vendor, subtotal, tax, shipping from each invoice.',
    '2. Tax split: item_taxShare = (item_amount / subtotal) * total_tax. If subtotal=0, split evenly.',
    '3. Shipping split: item_shippingShare = (item_amount / subtotal) * total_shipping.',
    '4. Pallet charges go to Masonry.',
    '5. A delivery line item (not footer total) is treated as shipping - distribute its cost proportionally.',
    '6. Returns/credits use negative amounts.',
    '7. Set uncertain:true if category is genuinely unclear.',
    '8. lineItemsSum = sum of all lineItem amounts (not including tax/shipping).',
    '9. balanceCheck = (Math.abs(lineItemsSum - subtotal) < 0.10).',
    '10. If invoice text is unreadable (scanned PDF), set lineItems:[] and notes:"Scanned - manual entry required".',
    '11. Do not include tax rows or shipping rows as separate line items - they belong in the tax/shipping fields.'
  ].join('\n');

  var invoices = payload.invoices || [];

  // Process in batches of 10 to stay within Claude token limits
  var allCategorized = [];
  var batchSize = 10;
  for (var b = 0; b < invoices.length; b += batchSize) {
    var batch = invoices.slice(b, b + batchSize);
    try {
      var resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
        method: 'POST',
        headers: {
          'x-api-key': apiKey,
          'anthropic-version': '2023-06-01',
          'content-type': 'application/json'
        },
        payload: JSON.stringify({
          model: 'claude-haiku-4-5-20251001',
          max_tokens: 8192,
          system: systemPrompt,
          messages: [{ role: 'user', content: JSON.stringify(batch) }]
        }),
        muteHttpExceptions: true
      });
      var raw = JSON.parse(resp.getContentText());
      if (raw.error) return { error: raw.error.message };
      var text = (raw.content && raw.content[0]) ? raw.content[0].text : '';
      // Strip any accidental markdown fences
      text = text.replace(/^```json\s*/m, '').replace(/^```\s*/m, '').replace(/```\s*$/m, '').trim();
      var parsed = JSON.parse(text);
      allCategorized = allCategorized.concat(Array.isArray(parsed) ? parsed : [parsed]);
    } catch(e) {
      return { error: 'Batch ' + (b/batchSize+1) + ' failed: ' + e.toString() };
    }
  }
  return { success: true, categorized: allCategorized };
}

// ─── Suggest Categories (lightweight) ────────────────────────────────────────
function suggestCategories(payload) {
  var auth = authorizeCaller(payload, ['admin', 'office']);
  if (!auth.ok) return { error: auth.error, code: auth.code };
  var apiKey = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');
  if (!apiKey) return { error: 'CLAUDE_API_KEY not set in Script Properties' };

  var items = payload.items || []; // [{idx, description, vendor, amount}]
  if (!items.length) return { suggestions: [] };

  var catList = [
    'Siding Lap      : LP SmartSide lap siding, 5/4 cedar trim boards',
    'Siding B&B      : LP SmartSide panels 4x10, battens 19/32x3, 4/4 cedar trim - wall siding only, not wrap',
    'Siding Flashing : Panel Union Flashing, Z-flashing, brick flashing',
    'Metal           : Coil stock, touch-up paint, metal accessories (non-soffit/fascia)',
    'Soffit & Fascia : Aluminum soffit panels (solid/vented), fascia trim, J-channel, drip edge, coil wrap',
    'Masonry         : Stone, brick, building paper, metal lath, mortar, pallet charges, lime',
    'Vinyl           : Vinyl siding panels',
    'Vinyl Accessories    : Vinyl starter strips, corners, J-channel for vinyl, outlet/light boxes',
    'Stucco          : Stucco base/finish, dryvit, mesh',
    'Angle Iron      : Steel angle iron, wide flange beams, structural steel',
    'Beam/Post/Garage Wrap : Hardboard/B&B panels for wrapping beams, posts, columns, or garage openings (not wall siding)'
  ].join('\n');

  var productList = [
    'LP 3/8x8x16 Lap',
    'Hardboard 4x10 Panel','Hardboard 4x8 Panel','Hardboard Cedar Shake','LP 19/32x3 Battens',
    '5/4 2" Trim','5/4 4" Trim','5/4 6" Trim','5/4 8" Trim','5/4 10" Trim','5/4 12" Trim',
    '4/4 2" Trim','4/4 4" Trim','4/4 6" Trim','4/4 8" Trim','4/4 10" Trim','4/4 12" Trim',
    'Panel Union Flashing','Window Flashing',
    'Coil Stock','Metal Accessories',
    'Alum Soffit Solid','Alum Soffit Vented','Alum Fascia','J-Channel','Touch-Up Paint',
    'Stone Veneer','Modular Brick','King Size Brick','Mortar Type S','Mortar Type N','Metal Lath','Building Paper','Pallet Charge',
    'Vinyl Lap Panel','Vinyl B&B Panel',
    'Starter Strip','Outside Corner','J-Channel Vinyl','Outlet Box','Light Box','Finish Trim',
    'Stucco Base Coat','Stucco Finish Coat','Stucco Mesh','Stucco Accessories',
    'Angle Iron',
    'Hardboard 4x10 Panel'
  ].join(', ');

  var systemPrompt = 'You are a building materials categorizer. Given a list of invoice line items, assign each to exactly one category AND suggest a canonical product name.\n\n'
    + 'Categories:\n' + catList + '\n\n'
    + 'Canonical product names (pick the closest match, or null if none fit):\n' + productList + '\n\n'
    + 'Return ONLY a JSON array: [{"idx":0,"category":"Metal","suggestedProduct":"Coil Stock"}, ...]\n'
    + 'Use exact category names. suggestedProduct must be one of the canonical names above, or null.';

  try {
    var resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01',
        'content-type': 'application/json'
      },
      payload: JSON.stringify({
        model: 'claude-haiku-4-5-20251001',
        max_tokens: 2048,
        system: systemPrompt,
        messages: [{ role: 'user', content: JSON.stringify(items) }]
      }),
      muteHttpExceptions: true
    });
    var raw = JSON.parse(resp.getContentText());
    if (raw.error) return { error: raw.error.message };
    var text = (raw.content[0].text || '').replace(/```json\s*/g,'').replace(/```/g,'').trim();
    return { suggestions: JSON.parse(text) };
  } catch(e) {
    return { error: e.toString() };
  }
}

// ── Process estimate PO + match to invoice line items ──
function processEstimateWithMatching(payload) {
  try {
    var auth = authorizeCaller(payload, ['admin', 'office']);
    if (!auth.ok) return { error: auth.error, code: auth.code };
    var estimateRows = payload.estimateRows || [];
    var invoiceItems = payload.invoiceItems || [];
    var apiKey = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');
    var categories = ['Siding Lap','Siding B&B','Siding Flashing','Metal','Soffit & Fascia','Masonry','Vinyl','Vinyl Accessories','Stucco','Angle Iron','Beam/Post/Garage Wrap'];

    var invSummary = invoiceItems.slice(0, 60).map(function(it) {
      return (it.description || '') + (it.qty ? ' | qty:' + it.qty : '') + (it.unit ? ' ' + it.unit : '') + (it.category ? ' [' + it.category + ']' : '');
    }).join('\n');

    // Support both spreadsheet rows (xlsx) and raw text (pdf)
    var estimateContent;
    if (payload.estimateText) {
      estimateContent = 'ESTIMATE TEXT (from PDF):\n' + payload.estimateText;
    } else {
      estimateContent = 'ESTIMATE ROWS (tab-separated):\n'
        + (payload.estimateRows || []).slice(0, 70).map(function(r){ return r.join('\t'); }).join('\n');
    }

    var prompt = 'You are analyzing a construction estimate and matching it to actual invoice line items.\n\n'
      + 'CATEGORIES: ' + categories.join(', ') + '\n\n'
      + estimateContent
      + '\n\nINVOICE LINE ITEMS (for matching):\n' + (invSummary || '(none)')
      + '\n\nFor each estimate material line item (skip headers/totals/blank/SqF summary rows):\n'
      + '1. Extract: description, ogQty (ordered qty), unit, estWastePct (waste factor %, as a number like 7 for 7%)\n'
      + '2. Assign one category from the list above\n'
      + '3. Find the best matching invoice line item(s) and sum their qty as actualQty (0 if no match)\n\n'
      + 'Return ONLY valid JSON:\n'
      + '{"items":[{"description":"...","category":"...","ogQty":0,"unit":"SqF","estWastePct":0,"actualQty":0}]}';

    var resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post',
      headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01', 'content-type': 'application/json' },
      payload: JSON.stringify({
        model: 'claude-haiku-4-5-20251001',
        max_tokens: 2048,
        messages: [{ role: 'user', content: prompt }]
      }),
      muteHttpExceptions: true
    });
    var body = JSON.parse(resp.getContentText());
    if (body.error) return { error: body.error.message };
    var text = (body.content[0].text || '').replace(/```json\s*/g,'').replace(/```/g,'').trim();
    var m = text.match(/\{[\s\S]*\}/);
    if (m) return JSON.parse(m[0]);
    return { items: [] };
  } catch(e) {
    return { error: e.toString() };
  }
}

// ── Append approved rows to Material Report History tab ──
function saveMaterialHistory(payload) {
  try {
    var auth = authorizeCaller(payload, ['admin', 'office']);
    if (!auth.ok) return { error: auth.error, code: auth.code };
    var rows = payload.rows || [];
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Material Report History');
    if (!sheet) return { error: 'Sheet "Material Report History" not found in this spreadsheet' };

    var HEADERS = ['Date','Job','Tier','Contractor','Category','Description','OG Qty','Est. Waste%','Unit','Product','Invoiced Qty','Return Qty','Actual Qty','Actual Waste%'];
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.getRange(1,1,1,HEADERS.length).setFontWeight('bold').setBackground('#1F3971').setFontColor('#ffffff');
    } else {
      // Add Product column if missing from existing sheet
      var existingHdrs = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      if (existingHdrs.indexOf('Product') === -1) {
        var unitCol = existingHdrs.indexOf('Unit') + 1; // 1-indexed
        if (unitCol > 0) {
          sheet.insertColumnAfter(unitCol);
          var prodCell = sheet.getRange(1, unitCol + 1);
          prodCell.setValue('Product').setFontWeight('bold').setBackground('#1F3971').setFontColor('#ffffff');
        }
      }
    }

    rows.forEach(function(r) {
      sheet.appendRow([
        r.date, r.job, r.tier || '', r.contractor, r.category, r.description || '',
        r.ogQty || '', r.estWastePct || '', r.unit || '',
        r.product || '',
        r.invoicedQty !== undefined ? r.invoicedQty : '',
        r.returnQty   !== undefined ? r.returnQty   : '',
        r.actualQty   !== undefined ? r.actualQty   : '',
        r.actualWastePct !== '' && r.actualWastePct !== undefined ? r.actualWastePct : ''
      ]);
    });

    return { saved: rows.length };
  } catch(e) {
    return { error: e.toString() };
  }
}

// -- SOPs ---------------------------------------------------------------------
var SOP_SHEET = "SOPs";

function getSopData() {
  try {
    var ss      = SpreadsheetApp.getActiveSpreadsheet();
    var sheet   = ss.getSheetByName(SOP_SHEET);
    if (!sheet) return { sops: [] };
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { sops: [] };
    var numRows  = lastRow - 1;
    var data     = sheet.getRange(2, 1, numRows, 6).getValues();
    var richCol    = sheet.getRange(2, 6, numRows, 1).getRichTextValues();
    var formulaCol = sheet.getRange(2, 6, numRows, 1).getFormulas();
    var sops = [];
    data.forEach(function(row, i) {
      if (!row[0]) return;
      var updated = '';
      if (row[3]) {
        try { updated = Utilities.formatDate(new Date(row[3]), Session.getScriptTimeZone(), 'MM/dd/yyyy'); } catch(e) { updated = String(row[3]); }
      }
      var pdfLink = String(row[5] || '');
      if (!pdfLink.match(/^https?:\/\//)) {
        var runs = richCol[i][0] ? richCol[i][0].getRuns() : [];
        for (var r = 0; r < runs.length; r++) {
          var u = runs[r].getLinkUrl();
          if (u && u.match(/^https?:\/\//)) { pdfLink = u; break; }
        }
      }
      if (!pdfLink.match(/^https?:\/\//)) {
        var formula = formulaCol[i][0] || '';
        var fm = formula.match(/=HYPERLINK\(\s*"([^"]+)"/i);
        if (fm) pdfLink = fm[1];
      }
      sops.push({
        title:       String(row[0] || ''),
        category:    String(row[1] || ''),
        role:        String(row[2] || ''),
        lastUpdated: updated,
        notes:       String(row[4] || ''),
        pdfLink:     pdfLink
      });
    });
    return { sops: sops };
  } catch(e) {
    return { error: e.toString() };
  }
}

// -- Asana Integration --------------------------------------------------------

var ASANA_API          = 'https://app.asana.com/api/1.0';
var ASANA_EXT_SCHED    = '1208049422174439';
var ASANA_OFFICE_TASKS = '1208049422174458';
var ASANA_PTO_PROJECT  = '1210392177822419';

function getAsanaPAT() {
  return PropertiesService.getScriptProperties().getProperty('ASANA_PAT');
}

/** Escapes text for use inside an Asana html_notes <body> -- Asana's rich-text field only accepts a small tag whitelist (b/i/u/s/a/code/ol/ul/li/etc.), so any &, <, > from user-entered note/item text must be entitized or the request is rejected / mis-rendered. */
function escapeAsanaHtml_(str) {
  return String(str == null ? '' : str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function asanaRequest(method, endpoint, payload) {
  var options = {
    method: method,
    headers: {
      'Authorization': 'Bearer ' + getAsanaPAT(),
      'Content-Type':  'application/json'
    },
    muteHttpExceptions: true
  };
  if (payload) options.payload = JSON.stringify({ data: payload });
  var resp = UrlFetchApp.fetch(ASANA_API + endpoint, options);
  return JSON.parse(resp.getContentText());
}

/**
 * Uploads a base64-encoded file as a native Asana attachment on a task or
 * subtask. Asana's attachments endpoint takes multipart/form-data, not the
 * JSON body asanaRequest() sends, so this builds its own UrlFetchApp call --
 * passing a Blob in `payload` makes Apps Script generate the multipart body
 * and boundary automatically (no Content-Type header needed here). Any
 * failure (bad gid, oversized file, Asana outage) is swallowed into
 * {success:false} rather than thrown, since callers upload several photos
 * in a loop and one bad attachment shouldn't block the rest -- the photo is
 * already safe in Drive regardless of this call's outcome.
 */
var ASANA_ATTACHMENT_ALLOWED_MIME_TYPES = ['image/jpeg', 'image/png', 'image/webp', 'image/heic', 'image/heif'];
var ASANA_ATTACHMENT_MAX_BASE64_LEN = 12000000; // ~9MB decoded -- comfortably above the client's compressed JPEGs, well under UrlFetchApp/Asana limits

function asanaUploadAttachment(taskGid, base64Data, mimeType, filename) {
  try {
    if (!base64Data || base64Data.length > ASANA_ATTACHMENT_MAX_BASE64_LEN) {
      return { success: false, error: 'Photo too large to attach.' };
    }
    var safeMimeType = ASANA_ATTACHMENT_ALLOWED_MIME_TYPES.indexOf(mimeType) !== -1 ? mimeType : 'image/jpeg';
    var safeFilename = (filename || 'photo.jpg').toString().replace(/[^A-Za-z0-9_.\- ]/g, '').slice(0, 120) || 'photo.jpg';
    var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), safeMimeType, safeFilename);
    var resp = UrlFetchApp.fetch(ASANA_API + '/tasks/' + taskGid + '/attachments', {
      method: 'post',
      headers: { 'Authorization': 'Bearer ' + getAsanaPAT() },
      payload: { file: blob },
      muteHttpExceptions: true
    });
    var json = JSON.parse(resp.getContentText());
    if (json.errors) return { success: false, error: json.errors[0].message };
    return { success: true, gid: json.data && json.data.gid };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Every caller of getAsanaJobs() pages through the full Exterior Master
 * Schedule project (up to 10 sequential Asana API round-trips) just to
 * populate a job picker -- that's the single slowest call in the app, and
 * with several UI panels (Quality Check, Job Dashboard) all hitting it
 * independently, that latency added up often enough to blow past the
 * client's fetch timeout even though the request would have succeeded a
 * few seconds later. A short script-cache means only the first caller in
 * any 3-minute window pays the Asana pagination cost; everyone else gets
 * an instant response. Job lists don't change second-to-second, so a few
 * minutes of staleness is an easy trade.
 */
function getAsanaJobs() {
  var cache = CacheService.getScriptCache();
  var cacheKey = 'asana_jobs_v1';
  try {
    var cached = cache.get(cacheKey);
    if (cached) return JSON.parse(cached);
  } catch (e) { /* fall through and fetch fresh */ }

  try {
    var jobs   = [];
    var offset = null;
    var maxPages = 10;
    for (var page = 0; page < maxPages; page++) {
      var url = '/projects/' + ASANA_EXT_SCHED +
        '/tasks?opt_fields=gid,name,completed&limit=100' +
        (offset ? '&offset=' + encodeURIComponent(offset) : '');
      var result = asanaRequest('get', url);
      if (result.errors) return { error: result.errors[0].message };
      (result.data || []).forEach(function(t) {
        if (!t.completed && t.name) jobs.push({ gid: t.gid, name: t.name });
      });
      if (result.next_page && result.next_page.offset) {
        offset = result.next_page.offset;
      } else {
        break;
      }
    }
    var response = { jobs: jobs };
    try { cache.put(cacheKey, JSON.stringify(response), 180); } catch (e) { /* e.g. over the 100KB cache limit -- fine, just skip caching */ }
    return response;
  } catch(e) { return { error: e.toString() }; }
}

/**
 * Bulk-fetches incomplete Exterior Master Schedule tasks together with their
 * current section membership, and groups them by section. Section names on
 * this Asana board (Estimate Requested, Siding, Masonry, etc.) ARE the job
 * phases -- this reuses the exact pagination shape of getAsanaJobs() above,
 * just with memberships.section fields added, so it's one bulk call instead
 * of the N+1 per-job section lookup getJobDashboard() does for a single job.
 * The phase list/order comes from the sections endpoint itself (the same
 * call getSectionGidByName() below makes) so nothing is hardcoded here.
 */
function getJobsByPhase(payload) {
  try {
    var auth = authorizeCaller(payload, ['admin']);
    if (!auth.ok) return { error: auth.error, code: auth.code };

    // Same slow up-to-10-page Asana pagination as getAsanaJobs() (plus a
    // sections lookup on top) -- cached the same way so repeat callers don't
    // re-pay that cost every time.
    var cache = CacheService.getScriptCache();
    var cacheKey = 'asana_jobs_by_phase_v1';
    try {
      var cached = cache.get(cacheKey);
      if (cached) return JSON.parse(cached);
    } catch (e) { /* fall through and fetch fresh */ }

    var sectionsResult = asanaRequest('get', '/projects/' + ASANA_EXT_SCHED + '/sections?opt_fields=gid,name');
    if (sectionsResult.errors) return { error: sectionsResult.errors[0].message };
    var sections = (sectionsResult.data || []).map(function(s) { return { gid: s.gid, name: s.name, jobs: [] }; });
    var sectionByGid = {};
    sections.forEach(function(s) { sectionByGid[s.gid] = s; });

    var offset = null;
    var maxPages = 10;
    for (var page = 0; page < maxPages; page++) {
      var url = '/projects/' + ASANA_EXT_SCHED +
        '/tasks?opt_fields=gid,name,completed,memberships.section.gid,memberships.section.name&limit=100' +
        (offset ? '&offset=' + encodeURIComponent(offset) : '');
      var result = asanaRequest('get', url);
      if (result.errors) return { error: result.errors[0].message };
      (result.data || []).forEach(function(t) {
        if (t.completed || !t.name) return;
        var membership = (t.memberships || []).filter(function(m) { return m.section && sectionByGid[m.section.gid]; })[0];
        if (membership) sectionByGid[membership.section.gid].jobs.push({ gid: t.gid, name: t.name });
      });
      if (result.next_page && result.next_page.offset) {
        offset = result.next_page.offset;
      } else {
        break;
      }
    }
    var response = { sections: sections };
    try { cache.put(cacheKey, JSON.stringify(response), 180); } catch (e) { /* over cache size limit -- fine, just skip caching */ }
    return response;
  } catch(e) { return { error: e.toString() }; }
}

/**
 * Splits a quality-check subtask's notes body into its per-item results --
 * one {status, text, note} per "[PASS]"/"[FLAG]"/"[N/A] item - note" line --
 * so the Job Dashboard can show the full breakdown for a walk, not just
 * the pass/flag/n-a counts. Item text never contains " - " (checked against
 * the checklist content), so splitting each line on the first occurrence
 * of it to separate item text from an optional flag note is safe.
 */
function parseQualityCheckItems_(notes) {
  var lines = (notes || '').split('\n').filter(function(l) {
    return /^\[(PASS|FLAG|N\/A)\]/.test(l);
  });
  return lines.map(function(l) {
    var m      = l.match(/^\[(PASS|FLAG|N\/A)\]\s*(.*)$/);
    var status = m[1] === 'PASS' ? 'pass' : (m[1] === 'N/A' ? 'na' : 'flag');
    var rest   = m[2] || '';
    var sepIdx = rest.indexOf(' - ');
    return {
      status: status,
      text:   sepIdx === -1 ? rest : rest.slice(0, sepIdx),
      note:   sepIdx === -1 ? ''   : rest.slice(sepIdx + 3)
    };
  });
}

/**
 * Reads quality-walk history straight from Asana -- submitQualityCheck logs
 * each check as a subtask of the job's Asana task (name "Quality Check
 * [<Walk Type>] - <date>", notes starting with "Walk Type:"/"Submitted
 * by:"/"Trade(s):" then one "[PASS]"/"[FLAG]"/"[N/A]" line per item), so
 * there's no separate sheet to keep in sync. Subtasks predating the walk
 * -type change just come back with walkType: ''.
 */
function getQualityWalkHistory_(jobGid) {
  var result = { count: 0, recent: [] };
  try {
    var subtasks = [];
    var offset = null;
    var maxPages = 5;
    for (var page = 0; page < maxPages; page++) {
      var url = '/tasks/' + jobGid + '/subtasks?opt_fields=name,notes,created_at&limit=100' +
        (offset ? '&offset=' + encodeURIComponent(offset) : '');
      var resp = asanaRequest('get', url);
      if (resp.errors) break;
      subtasks = subtasks.concat(resp.data || []);
      if (resp.next_page && resp.next_page.offset) offset = resp.next_page.offset;
      else break;
    }

    var matches = subtasks.filter(function(t) {
      return (t.name || '').indexOf('Quality Check') === 0;
    }).map(function(t) {
      var notes     = t.notes || '';
      var walkType  = (notes.match(/Walk Type:\s*(.+)/)     || [])[1] || '';
      var submitter = (notes.match(/Submitted by:\s*(.+)/)  || [])[1] || '';
      var trades    = (notes.match(/Trade\(s\):\s*(.+)/)    || [])[1] || '';
      var items     = parseQualityCheckItems_(notes);
      var passCount = 0, flagCount = 0, naCount = 0;
      items.forEach(function(it) {
        if (it.status === 'pass') passCount++;
        else if (it.status === 'flag') flagCount++;
        else naCount++;
      });
      var created = t.created_at ? new Date(t.created_at) : null;
      return {
        gid:       t.gid,
        timestamp: created ? Utilities.formatDate(created, Session.getScriptTimeZone(), 'MM/dd/yy') : '',
        ts:        created ? created.getTime() : 0,
        walkType:  walkType,
        trades:    trades,
        submitter: submitter,
        passCount: passCount,
        flagCount: flagCount,
        naCount:   naCount,
        items:     items
      };
    });

    matches.sort(function(a, b) { return b.ts - a.ts; });
    result.count = matches.length;
    result.recent = matches.slice(0, 8);
  } catch (e) {
    // swallow -- a Quality Walks read failure shouldn't break the rest of the Job Dashboard
  }
  return result;
}

/**
 * Fetches the photos attached to a single Quality Check subtask, for the
 * read-only walk detail view. Kept separate from getQualityWalkHistory_/
 * getRecentQualityWalks (rather than bulk-fetched alongside them) since
 * attachments need their own Asana API call per subtask -- doing that for
 * every walk in a history list would be N+1 and slow the list down for a
 * gallery most walks don't even have. Called only when a user actually
 * opens a walk's detail view.
 */
function getQualityWalkPhotos(payload) {
  try {
    var auth = authorizeCaller(payload, ['admin', 'office', 'site_manager']);
    if (!auth.ok) return { photos: [], error: auth.error, code: auth.code };
    var taskGid = (payload && payload.taskGid || '').toString().trim();
    if (!taskGid) return { photos: [] };
    var result = asanaRequest('get', '/tasks/' + taskGid + '/attachments?opt_fields=name,view_url,download_url');
    if (result.errors) return { photos: [], error: result.errors[0].message };
    var photos = (result.data || []).map(function(a) {
      // download_url is a direct-binary link (embeddable as <img src>); view_url is
      // Asana's authenticated page (not embeddable, but a nicer "open in Asana" link).
      return { name: a.name || '', url: a.download_url || a.view_url || '', viewUrl: a.view_url || '' };
    }).filter(function(p) { return p.url; });
    return { photos: photos };
  } catch (e) { return { photos: [], error: e.toString() }; }
}

/**
 * The 5 most recent Quality Check walks across ALL jobs, for the aidan-only
 * Dashboard. Quality Check walks are subtasks scattered across every job
 * task in ASANA_EXT_SCHED, so getting "most recent across all jobs" cheaply
 * means using Asana's workspace task search (one call, sorted server-side)
 * rather than crawling every job's subtasks one by one (N+1, slow). Falls
 * back to an empty list (not an error) if search comes back empty so one
 * missing/renamed workspace doesn't break the rest of the Dashboard.
 */
function getRecentQualityWalks(payload) {
  try {
    var auth = authorizeCaller(payload, ['admin']);
    if (!auth.ok) return { error: auth.error, code: auth.code };
    var limit = Math.min(Math.max(parseInt((payload && payload.limit), 10) || 5, 1), 50);

    var wsResult = asanaRequest('get', '/workspaces?opt_fields=gid&limit=1');
    if (wsResult.errors || !wsResult.data || !wsResult.data.length) return { walks: [] };
    var workspaceGid = wsResult.data[0].gid;

    // Fetch more than `limit` up front -- some recent results turn out to be
    // older free-text-summary tasks (pre-dating the current Walk Type/
    // Trade(s)/[PASS]-[FLAG]-[N/A] format and never created as a job
    // subtask), which get filtered out below. Without the headroom, a few of
    // those near the top of the date-sorted results would silently leave
    // fewer than `limit` real walks instead of backfilling from further back.
    var searchFetch = Math.min(Math.max(limit * 4, 20), 100);
    var searchUrl = '/workspaces/' + workspaceGid +
      '/tasks/search?text=Quality Check&sort_by=created_at&sort_ascending=false&limit=' + searchFetch +
      '&opt_fields=name,notes,created_at,parent.name,parent.gid,permalink_url';
    var searchResult = asanaRequest('get', searchUrl);
    if (searchResult.errors) return { walks: [], error: searchResult.errors[0].message };

    var walks = (searchResult.data || []).filter(function(t) {
      return (t.name || '').indexOf('Quality Check') === 0;
    }).map(function(t) {
      var notes     = t.notes || '';
      var walkType  = (notes.match(/Walk Type:\s*(.+)/)    || [])[1] || '';
      var submitter = (notes.match(/Submitted by:\s*(.+)/) || [])[1] || '';
      var trades    = (notes.match(/Trade\(s\):\s*(.+)/)   || [])[1] || '';
      var created   = t.created_at ? new Date(t.created_at) : null;
      var items     = parseQualityCheckItems_(notes);
      var passCount = 0, flagCount = 0, naCount = 0;
      items.forEach(function(it) {
        if (it.status === 'pass') passCount++;
        else if (it.status === 'flag') flagCount++;
        else naCount++;
      });
      return {
        gid:          t.gid,
        jobName:      t.parent ? t.parent.name : '',
        jobGid:       t.parent ? t.parent.gid : '',
        walkType:     walkType,
        submitter:    submitter,
        trades:       trades,
        timestamp:    created ? Utilities.formatDate(created, Session.getScriptTimeZone(), 'MM/dd/yy') : '',
        items:        items,
        passCount:    passCount,
        flagCount:    flagCount,
        naCount:      naCount,
        notes:        notes,
        permalinkUrl: t.permalink_url || ''
      };
    }).filter(function(w) {
      // Only genuine structured walks -- older free-text-summary tasks (no
      // parsed checklist items) are a different, pre-current-format thing
      // and are excluded here rather than shown as "Unknown job".
      return w.items.length > 0;
    }).slice(0, limit);

    return { walks: walks };
  } catch (e) { return { error: e.toString() }; }
}

/**
 * All of the caller's own assigned, incomplete Asana tasks across every
 * project/list they're in -- for the "My Asana Tasks" Dashboard card and
 * Other-tab panel. Two things make this different from getAsanaJobs()/
 * getRecentQualityWalks() above:
 *
 * 1. The whole app authenticates to Asana with one shared script PAT
 *    (getAsanaPAT()), which belongs to a single service-account identity --
 *    not the person using the app. So `assignee=me` would return that
 *    service account's tasks, not the caller's. The caller's own email
 *    (verified via requireVerifiedEmail_) has to be translated to their
 *    Asana user gid first via GET /users/{email}, which only works because
 *    this is a domain-restricted Asana org. That gid is cached for 6 hours
 *    (an Asana identity doesn't change minute to minute) so repeat loads
 *    only pay for it once per workday, not once per dashboard paint.
 * 2. "All of a person's tasks" spans however many projects they're a member
 *    of -- crawling each one (like getAsanaJobs()'s up-to-10-page pagination
 *    per project) would multiply that cost per project per user. Asana's
 *    workspace task search does it in one call instead, same tradeoff
 *    getRecentQualityWalks() already makes above. Results are capped at a
 *    single page of 100 (sorted soonest-due-first) rather than paginated
 *    further -- a person realistically has far fewer open assigned tasks
 *    than that, and paginating would reintroduce the exact per-call latency
 *    this design is trying to avoid.
 *
 * The task list itself (unlike the gid) is cached for only 120s -- someone
 * actively working through their tasks expects to see changes reflected
 * soon, so staleness tolerance here is much lower than the identity lookup.
 */
function getMyAsanaTasks(payload) {
  try {
    var auth = requireVerifiedEmail_(payload);
    if (auth.error) return { error: auth.error, code: auth.code };
    var email = auth.email;

    var cache = CacheService.getScriptCache();
    var gidCacheKey = 'asana_user_gid_v1_' + email;
    var userGid = null;
    try { userGid = cache.get(gidCacheKey); } catch (e) { /* fall through and resolve fresh */ }

    if (userGid === null) {
      // Not cached at all (distinct from a cached '' sentinel, which means
      // "already confirmed this email has no Asana account").
      var userResult = asanaRequest('get', '/users/' + encodeURIComponent(email) + '?opt_fields=gid');
      userGid = (!userResult.errors && userResult.data && userResult.data.gid) ? userResult.data.gid : '';
      try { cache.put(gidCacheKey, userGid, userGid ? 21600 : 900); } catch (e) { /* over cache size limit -- fine, just skip caching */ }
    }

    if (!userGid) return { tasks: [], noAsanaAccount: true };

    var tasksCacheKey = 'asana_my_tasks_v1_' + userGid;
    try {
      var cachedTasks = cache.get(tasksCacheKey);
      if (cachedTasks) return JSON.parse(cachedTasks);
    } catch (e) { /* fall through and fetch fresh */ }

    var wsResult = asanaRequest('get', '/workspaces?opt_fields=gid&limit=1');
    if (wsResult.errors || !wsResult.data || !wsResult.data.length) return { tasks: [] };
    var workspaceGid = wsResult.data[0].gid;

    var searchUrl = '/workspaces/' + workspaceGid + '/tasks/search' +
      '?assignee.any=' + userGid + '&completed=false' +
      '&sort_by=due_date&sort_ascending=true&limit=100' +
      '&opt_fields=name,due_on,permalink_url,memberships.project.gid,memberships.project.name';
    var searchResult = asanaRequest('get', searchUrl);
    if (searchResult.errors) return { tasks: [], error: searchResult.errors[0].message };

    var tasks = (searchResult.data || []).map(function(t) {
      var projects = (t.memberships || [])
        .map(function(m) { return m.project; })
        .filter(function(p) { return p && p.gid; })
        .map(function(p) { return { gid: p.gid, name: p.name || '' }; });
      return {
        gid:          t.gid,
        name:         t.name || '',
        dueOn:        t.due_on || '',
        permalinkUrl: t.permalink_url || '',
        projects:     projects
      };
    });

    var response = { tasks: tasks };
    try { cache.put(tasksCacheKey, JSON.stringify(response), 120); } catch (e) { /* over cache size limit -- fine, just skip caching */ }
    return response;
  } catch (e) { return { error: e.toString() }; }
}

var QC_SUBMIT_MAX_PHOTOS = 30; // generous upper bound (legitimate UI max is 3 per flagged item + 3 general) -- just a backstop against a malformed/hostile payload

function submitQualityCheck(payload) {
  try {
    var auth = requireVerifiedEmail_(payload);
    if (auth.error) return { error: auth.error, code: auth.code };

    // Idempotency guard -- see createPO() for the full rationale. Only the
    // check + the primary subtask-create call below are held under the
    // lock (mirroring submitOfficeNote's single-Asana-call pattern) -- the
    // lock is released before the slower photo-upload/flagged-item work so
    // a multi-photo submission doesn't hold the global lock for its whole
    // duration and block unrelated actions (createPO, clockIn, etc.).
    var idemKey = (payload.idempotencyKey || '').toString().trim();
    var cache = CacheService.getScriptCache();
    var cacheKey = idemKey ? ('idem_qualitycheck_' + idemKey) : null;
    var lock = cacheKey ? LockService.getScriptLock() : null;
    var haveLock = lock ? lock.tryLock(10000) : false;
    if (haveLock) {
      var cached = null;
      try { cached = cache.get(cacheKey); } catch (e) {}
      if (cached) { lock.releaseLock(); return JSON.parse(cached); }
    }

    var jobGid        = payload.jobGid;
    var jobName       = payload.jobName;
    var sections      = payload.sections;
    var submitter     = payload.submitter || 'Field';
    var walkTypeLabel = payload.walkTypeLabel || 'Quality Check';
    var tradesStr     = (payload.trades || []).join(', ') || 'General';
    var tz            = Session.getScriptTimeZone();
    var date          = Utilities.formatDate(new Date(), tz, 'MM/dd/yyyy');

    var lines = ['Walk Type: ' + walkTypeLabel, 'Submitted by: ' + submitter, 'Trade(s): ' + tradesStr];
    if (payload.equipmentOnsite === true || payload.equipmentOnsite === false) {
      lines.push('Equipment On Site: ' + (payload.equipmentOnsite ? 'Yes' : 'No'));
    }
    lines.push('');
    var flagged = [];
    var allPhotos = []; // flat list of {base64Data, mimeType, filename} across every item + general, for the Asana attachment pass below
    function photoNote(photos) {
      var n = (photos || []).length;
      return n ? (n + ' photo' + (n > 1 ? 's' : '') + ' attached') : '';
    }
    sections.forEach(function(s) {
      var icon = s.status === 'flag' ? 'FLAG' : (s.status === 'na' ? 'N/A' : 'PASS');
      var noteParts = [];
      if (s.notes) noteParts.push(s.notes);
      var pNote = photoNote(s.photos);
      if (pNote) noteParts.push(pNote);
      lines.push('[' + icon + '] ' + s.name + (noteParts.length ? ' - ' + noteParts.join(' — ') : ''));
      if (s.status === 'flag') flagged.push(s);
      if (s.photos && s.photos.length) allPhotos = allPhotos.concat(s.photos);
    });
    var generalNotes = (payload.generalNotes || '').toString().trim();
    var generalPhotos = payload.generalPhotos || [];
    var genPNote = photoNote(generalPhotos);
    var genNoteParts = [];
    if (generalNotes) genNoteParts.push(generalNotes);
    if (genPNote) genNoteParts.push(genPNote);
    if (genNoteParts.length) lines.push('', 'General Notes: ' + genNoteParts.join(' — '));
    if (generalPhotos.length) allPhotos = allPhotos.concat(generalPhotos);

    var sub = asanaRequest('post', '/tasks/' + jobGid + '/subtasks', {
      name:      'Quality Check [' + walkTypeLabel + '] - ' + date,
      notes:     lines.join('\n'),
      completed: true
    });
    if (sub.errors) {
      if (haveLock) lock.releaseLock();
      return { error: sub.errors[0].message };
    }
    var subtaskGid = sub.data && sub.data.gid;

    // The subtask now exists -- cache a success marker and release the lock
    // before the slower photo-upload/flagged-item work below, which doesn't
    // need lock protection since a retry is now recognized via this cache
    // entry regardless of how long the rest of this request takes.
    if (cacheKey) {
      try { cache.put(cacheKey, JSON.stringify({ success: true, flagged: 0, photosAttached: 0, photosFailed: 0 }), 300); } catch (e) {}
    }
    if (haveLock) { lock.releaseLock(); haveLock = false; }

    var photosAttached = 0, photosFailed = 0;
    if (subtaskGid) {
      allPhotos.slice(0, QC_SUBMIT_MAX_PHOTOS).forEach(function(p) {
        if (!p || !p.base64Data) return;
        var res = asanaUploadAttachment(subtaskGid, p.base64Data, p.mimeType || 'image/jpeg', p.filename || 'photo.jpg');
        if (res.success) photosAttached++; else photosFailed++;
      });
    }

    if (flagged.length > 0) {
      // Bold the item name (the "question") via html_notes so it's scannable at a glance in Asana,
      // leaving the field-entered note itself in normal text right after it.
      var offLines = ['Quality check flagged items for: ' + escapeAsanaHtml_(jobName) + ' (' + date + ')', ''];
      var flaggedPhotos = []; // photos on flagged items only -- also attached to this Office Task copy, not just the subtask
      flagged.forEach(function(f) {
        var fNoteParts = [];
        if (f.notes) fNoteParts.push(escapeAsanaHtml_(f.notes));
        var fPNote = photoNote(f.photos);
        if (fPNote) fNoteParts.push(fPNote);
        offLines.push('- <b>' + escapeAsanaHtml_(f.name) + '</b>' + (fNoteParts.length ? ': ' + fNoteParts.join(' — ') : ''));
        if (f.photos && f.photos.length) flaggedPhotos = flaggedPhotos.concat(f.photos);
      });
      var officeTask = asanaRequest('post', '/tasks', {
        projects:   [ASANA_OFFICE_TASKS],
        name:       'Quality Check - ' + jobName + ' - ' + date,
        html_notes: '<body>' + offLines.join('\n') + '</body>',
        assignee:   auth.email
      });
      var officeTaskGid = officeTask.data && officeTask.data.gid;
      if (officeTaskGid) {
        flaggedPhotos.slice(0, QC_SUBMIT_MAX_PHOTOS).forEach(function(p) {
          if (!p || !p.base64Data) return;
          asanaUploadAttachment(officeTaskGid, p.base64Data, p.mimeType || 'image/jpeg', p.filename || 'photo.jpg');
        });
      }
    }

    var result = { success: true, flagged: flagged.length, photosAttached: photosAttached, photosFailed: photosFailed };
    if (cacheKey) { try { cache.put(cacheKey, JSON.stringify(result), 300); } catch (e) {} }
    return result;
  } catch(e) { return { error: e.toString() }; }
  finally { if (haveLock) lock.releaseLock(); }
}

/**
 * Office Notes intake: creates a task directly in the ASANA_OFFICE_TASKS
 * project, optionally with a due date and an assignee. Replaces the external
 * Asana-hosted form, which had no way to assign the resulting task to
 * someone. Asana's task-create API accepts a plain email address for
 * `assignee`, so no separate Asana-user-gid mapping is needed here.
 *
 * Task name is the note text itself (trimmed/truncated) rather than a
 * generic "Office Note - <date>" label, so the task list is scannable.
 * Photos come in as base64 (client already saved them to Drive via
 * saveOfficeNotePhoto for backup and kept the bytes in memory) and are
 * attached to the created task as real Asana attachments, matching the
 * pattern in submitQualityCheck below.
 */
var OFFICE_NOTE_MAX_PHOTOS = 10; // backstop against a malformed/hostile payload
var OFFICE_NOTE_TASK_NAME_MAX_LEN = 100;

function submitOfficeNote(payload) {
  var auth = requireVerifiedEmail_(payload);
  if (auth.error) return { success: false, error: auth.error, code: auth.code };

  var lock = LockService.getScriptLock();
  var haveLock = lock.tryLock(10000);
  try {
    var note         = (payload.note || '').toString().trim();
    var dueDate       = (payload.dueDate || '').toString().trim();
    var assigneeEmail = (payload.assigneeEmail || '').toString().trim();
    var photos         = payload.photos || [];
    var submittedBy    = (payload.submittedBy || '').toString().trim();

    if (!note) return { success: false, error: 'Note is required' };

    // Idempotency guard -- see createPO() for the full rationale. A cache
    // hit here means this exact submission already created its Asana task.
    var idemKey = (payload.idempotencyKey || '').toString().trim();
    var cache = CacheService.getScriptCache();
    var cacheKey = (haveLock && idemKey) ? ('idem_officenote_' + idemKey) : null;
    if (cacheKey) {
      var cached = null;
      try { cached = cache.get(cacheKey); } catch (e) { /* ignore, fall through */ }
      if (cached) return JSON.parse(cached);
    }

    var tz   = Session.getScriptTimeZone();
    var date = Utilities.formatDate(new Date(), tz, 'MM/dd/yyyy');

    var lines = ['Note: ' + note, 'Submitted by: ' + (submittedBy || 'N/A'), 'Submitted: ' + date];
    if (photos.length) lines.push(photos.length + ' photo' + (photos.length > 1 ? 's' : '') + ' attached');

    var taskName = note.replace(/\s+/g, ' ').trim();
    if (taskName.length > OFFICE_NOTE_TASK_NAME_MAX_LEN) {
      taskName = taskName.slice(0, OFFICE_NOTE_TASK_NAME_MAX_LEN - 1) + '…';
    }

    var taskPayload = {
      projects: [ASANA_OFFICE_TASKS],
      name:     taskName,
      notes:    lines.join('\n')
    };
    if (dueDate) taskPayload.due_on = dueDate;
    // Default to whoever submitted the note (the verified session email) when the
    // "Assigned To" picker was left unset; an explicit pick there still wins.
    taskPayload.assignee = assigneeEmail || auth.email;

    var created = asanaRequest('post', '/tasks', taskPayload);
    if (created.errors) return { success: false, error: created.errors[0].message };

    var taskGid = created.data.gid;
    var result = {
      success:      true,
      asanaTaskGid: taskGid,
      asanaTaskUrl: 'https://app.asana.com/0/' + ASANA_OFFICE_TASKS + '/' + taskGid
    };

    // The task now exists -- cache a success marker and release the lock
    // before the slower photo-attachment work below, which doesn't need
    // lock protection since a retry is now recognized via this cache entry
    // regardless of how long the rest of this request takes.
    if (cacheKey) {
      try { cache.put(cacheKey, JSON.stringify(result), 300); } catch (e) { /* fine to skip */ }
    }
    if (haveLock) { lock.releaseLock(); haveLock = false; }

    var photosAttached = 0, photosFailed = 0;
    photos.slice(0, OFFICE_NOTE_MAX_PHOTOS).forEach(function(p) {
      if (!p || !p.base64Data) return;
      var res = asanaUploadAttachment(taskGid, p.base64Data, p.mimeType || 'image/jpeg', p.filename || 'photo.jpg');
      if (res.success) photosAttached++; else photosFailed++;
    });
    result.photosAttached = photosAttached;
    result.photosFailed   = photosFailed;

    return result;
  } catch(e) { return { success: false, error: e.toString() }; }
  finally { if (haveLock) lock.releaseLock(); }
}

/**
 * New Project intake: creates an Asana task in ASANA_EXT_SCHED (moved into
 * the "Estimate Requested" section), then appends a row to the "Projects"
 * sheet linking Contractor + Job Name to the Drive folder ID and the new
 * Asana task GID. Mirrors the external "New Exteriors Project" Asana form.
 *
 * The Asana task is created before the sheet row is written. If the sheet
 * write then fails, the Asana task still exists -- return its link/GID in
 * the error response so the task isn't silently orphaned.
 */
function createProjectAndTask(payload) {
  try {
    var auth = authorizeCaller(payload, ['admin', 'office']);
    if (!auth.ok) return { success: false, error: auth.error, code: auth.code };

    var builder        = (payload.builder || '').toString().trim();
    var jobName         = (payload.jobName || '').toString().trim();
    var address         = (payload.address || '').toString().trim();
    var googleMaps      = (payload.googleMaps || '').toString().trim();
    var driveLink       = (payload.driveLink || '').toString().trim();
    var estimateDueDate = (payload.estimateDueDate || '').toString().trim();
    var longLead        = (payload.longLead || '').toString().trim();
    var senderNotes      = (payload.senderNotes || '').toString().trim();
    var homePlansUrl     = (payload.homePlansUrl || '').toString().trim();
    var submittedBy      = (payload.submittedBy || '').toString().trim();

    if (!builder || !jobName || !address || !googleMaps || !driveLink || !estimateDueDate) {
      return { success: false, error: 'Builder Name, Job Name, Address, Google Maps, Google Drive Project Link, and Estimate Due Date are required.' };
    }

    var folderId = extractDriveFolderId(driveLink);
    if (!folderId) {
      return { success: false, error: 'Could not read a folder ID from the Google Drive Project Link.' };
    }

    // Idempotency guard -- see createPO() for the full rationale. A cache
    // hit here means this exact submission already created its Asana task
    // and Projects-sheet row.
    var idemKey = (payload.idempotencyKey || '').toString().trim();
    var cache = CacheService.getScriptCache();
    var cacheKey = idemKey ? ('idem_newproject_' + idemKey) : null;
    var lock = LockService.getScriptLock();
    var haveLock = lock.tryLock(10000);
    try {
      if (cacheKey && haveLock) {
        var cached = null;
        try { cached = cache.get(cacheKey); } catch (e) {}
        if (cached) return JSON.parse(cached);
      }

      var tz    = Session.getScriptTimeZone();
      var today = Utilities.formatDate(new Date(), tz, 'MM/dd/yyyy');
      var taskName = builder + ', ' + jobName + ', ' + address;
      var notes = [
        'Builder Name: '   + builder,
        'Job Name: '       + jobName,
        'Address: '        + address,
        'Google Maps: '    + googleMaps,
        'Google Drive: '   + driveLink,
        'Home Plans: '     + (homePlansUrl || 'None uploaded'),
        'Long Lead-time for Materials: ' + (longLead || 'N/A'),
        "Senders Email & Notes: " + (senderNotes || 'N/A'),
        'Estimate Due Date: ' + estimateDueDate,
        'Submitted by: '   + (submittedBy || 'N/A'),
        'Submitted: '      + today
      ].join('\n');

      var created = asanaRequest('post', '/tasks', {
        projects: [ASANA_EXT_SCHED],
        name:     taskName,
        notes:    notes,
        due_on:   estimateDueDate // input type="date" already gives YYYY-MM-DD, what Asana expects
      });
      if (created.errors) return { success: false, error: created.errors[0].message };

      var asanaTaskGid = created.data.gid;

      var sectionGid = getSectionGidByName(ASANA_EXT_SCHED, 'Estimate Requested');
      if (sectionGid) {
        asanaRequest('post', '/sections/' + sectionGid + '/addTask', { task: asanaTaskGid });
      }

      try {
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PROJECTS_SHEET_NAME);
        if (!sheet) throw new Error("Sheet '" + PROJECTS_SHEET_NAME + "' not found.");
        var nextRow = sheet.getLastRow() + 1;
        sheet.getRange(nextRow, 1).setValue(builder);
        sheet.getRange(nextRow, 2).setValue(jobName);
        sheet.getRange(nextRow, 3).setValue(folderId);
        sheet.getRange(nextRow, 4).setValue(asanaTaskGid);
      } catch (sheetErr) {
        return {
          success: false,
          error: 'Asana task was created but the Projects sheet row failed to save: ' + sheetErr.toString(),
          asanaTaskGid: asanaTaskGid,
          asanaTaskUrl: 'https://app.asana.com/0/' + ASANA_EXT_SCHED + '/' + asanaTaskGid
        };
      }

      var result = {
        success: true,
        driveFolderId: folderId,
        asanaTaskGid: asanaTaskGid,
        asanaTaskUrl: 'https://app.asana.com/0/' + ASANA_EXT_SCHED + '/' + asanaTaskGid
      };
      if (cacheKey) { try { cache.put(cacheKey, JSON.stringify(result), 300); } catch (e) {} }
      invalidateConfigOptionsCache_(); // this builder/job pair is new
      return result;
    } finally {
      if (haveLock) lock.releaseLock();
    }
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// -- PTO / HR Functions ───────────────────────────────────────────────────────
// HR sheet columns: A=Name, B=Email, C=Phone, D=Role, E=Password, F=Allotted, G=Used

/**
 * Gets PTO balance + request history for an employee (and pending queue for admins).
 */
function getPTOData(payload) {
  try {
    var auth = requireVerifiedEmail_(payload);
    if (auth.error) return auth;
    var email = auth.email;
    var callerRoles = getRoleByEmail(email).effRoles;
    var canSeeQueue = hasAnyRole_(callerRoles, ['admin', 'human_resources']);

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hrSheet = ss.getSheetByName(ROLES_SHEET);
    var balance = { allotted: 0, used: 0, remaining: 0, name: '' };

    if (hrSheet) {
      var lastRow = hrSheet.getLastRow();
      if (lastRow >= 2) {
        var data = hrSheet.getRange(2, 1, lastRow - 1, 7).getValues();
        for (var i = 0; i < data.length; i++) {
          var rowEmail = (data[i][1] || '').toString().toLowerCase().trim();
          if (rowEmail === email.toLowerCase().trim()) {
            balance.name      = (data[i][0] || '').toString().trim();
            balance.allotted  = parseFloat(data[i][5]) || 0;
            balance.used      = parseFloat(data[i][6]) || 0;
            balance.remaining = balance.allotted - balance.used;
            break;
          }
        }
      }
    }

    // Fetch all tasks from PTO project
    var result = asanaRequest('get',
      '/projects/' + ASANA_PTO_PROJECT +
      '/tasks?opt_fields=gid,name,notes,completed,memberships.section.name&limit=100');
    if (result.errors) return { error: result.errors[0].message };

    var myRequests   = [];
    var pendingQueue = [];

    (result.data || []).forEach(function(task) {
      var notes = task.notes || '';
      var section = '';
      if (task.memberships && task.memberships[0] && task.memberships[0].section) {
        section = task.memberships[0].section.name || '';
      }

      var parseField = function(label) {
        var m = notes.match(new RegExp(label + ':\\s*([^\\n]+)'));
        return m ? m[1].trim() : '';
      };

      var taskEmail = parseField('Requester');
      var status = section === 'Approved' ? 'approved'
                 : section === 'Denied' ? 'denied'
                 : 'pending';

      var req = {
        gid:            task.gid,
        requesterEmail: taskEmail,
        requesterName:  parseField('Name') || task.name,
        dates:          parseField('Dates'),
        days:           parseFloat(parseField('Days')) || 0,
        reason:         parseField('Reason'),
        status:         status
      };

      if (taskEmail.toLowerCase() === email.toLowerCase().trim()) {
        myRequests.push(req);
      }
      if (status === 'pending' && canSeeQueue) {
        pendingQueue.push(req);
      }
    });

    return { balance: balance, myRequests: myRequests, pendingQueue: pendingQueue };
  } catch(e) { return { error: e.toString() }; }
}

/**
 * Creates a PTO request task in Asana under New Requests.
 */
function submitPTORequest(payload) {
  try {
    var auth = requireVerifiedEmail_(payload);
    if (auth.error) return auth;
    var email  = auth.email;
    var name   = payload.name || email;
    var start  = payload.startDate;
    var end    = payload.endDate;
    var days   = payload.days;
    var reason = payload.reason || 'N/A';
    var tz     = Session.getScriptTimeZone();
    var today  = Utilities.formatDate(new Date(), tz, 'MM/dd/yyyy');

    var taskName = 'PTO - ' + name + ' (' + start + (start !== end ? ' to ' + end : '') + ')';
    var notes = [
      'Name: '      + name,
      'Requester: ' + email,
      'Dates: '     + start + (start !== end ? ' - ' + end : ''),
      'Days: '      + days,
      'Reason: '    + reason,
      'Submitted: ' + today
    ].join('\n');

    // Create task
    var created = asanaRequest('post', '/tasks', {
      projects: [ASANA_PTO_PROJECT],
      name:     taskName,
      notes:    notes
    });
    if (created.errors) return { error: created.errors[0].message };

    // Move to New Requests section
    var sectionGid = getPTOSectionGid('New Requests');
    if (sectionGid && created.data && created.data.gid) {
      asanaRequest('post', '/sections/' + sectionGid + '/addTask', { task: created.data.gid });
    }

    return { success: true };
  } catch(e) { return { error: e.toString() }; }
}

/**
 * Returns all pending (non-completed, non-denied) PTO requests for admin view.
 */
function getPTOQueue(payload) {
  try {
    var auth = authorizeCaller(payload, ['admin', 'human_resources']);
    if (!auth.ok) return { error: auth.error, code: auth.code };

    var result = asanaRequest('get',
      '/projects/' + ASANA_PTO_PROJECT +
      '/tasks?opt_fields=gid,name,notes,completed,memberships.section.name&limit=100');
    if (result.errors) return { error: result.errors[0].message };

    var queue = [];
    (result.data || []).forEach(function(task) {
      if (task.completed) return;
      var section = '';
      if (task.memberships && task.memberships[0] && task.memberships[0].section) {
        section = task.memberships[0].section.name || '';
      }
      if (section === 'Approved' || section === 'Denied') return;

      var notes = task.notes || '';
      var parseField = function(label) {
        var m = notes.match(new RegExp(label + ':\\s*([^\\n]+)'));
        return m ? m[1].trim() : '';
      };

      queue.push({
        gid:            task.gid,
        requesterName:  parseField('Name') || task.name,
        requesterEmail: parseField('Requester'),
        dates:          parseField('Dates'),
        days:           parseFloat(parseField('Days')) || 0,
        reason:         parseField('Reason')
      });
    });

    return { queue: queue };
  } catch(e) { return { error: e.toString() }; }
}

/**
 * Marks a PTO request as approved: completes the Asana task + increments Used days on HR sheet.
 */
function approvePTO(payload) {
  try {
    var auth = authorizeCaller(payload, ['admin', 'human_resources']);
    if (!auth.ok) return { error: auth.error, code: auth.code };

    var taskGid  = payload.taskGid;
    var empEmail = payload.employeeEmail;
    var days     = parseFloat(payload.days) || 0;

    // Move to Approved section (triggers Asana email rule)
    var approvedGid = getPTOSectionGid('Approved');
    if (!approvedGid) return { error: 'Approved section not found in Asana project' };
    var moved = asanaRequest('post', '/sections/' + approvedGid + '/addTask', { task: taskGid });
    if (moved.errors) return { error: moved.errors[0].message };

    if (empEmail && days > 0) updatePTOUsed(empEmail, days);
    return { success: true };
  } catch(e) { return { error: e.toString() }; }
}

/**
 * Marks a PTO request as denied: renames it [Denied] and completes it.
 */
function denyPTO(payload) {
  try {
    var auth = authorizeCaller(payload, ['admin', 'human_resources']);
    if (!auth.ok) return { error: auth.error, code: auth.code };

    var taskGid = payload.taskGid;

    // Move to Denied section (triggers Asana email rule)
    var deniedGid = getPTOSectionGid('Denied');
    if (!deniedGid) return { error: 'Denied section not found in Asana project' };
    var moved = asanaRequest('post', '/sections/' + deniedGid + '/addTask', { task: taskGid });
    if (moved.errors) return { error: moved.errors[0].message };
    return { success: true };
  } catch(e) { return { error: e.toString() }; }
}

/**
 * Cancels/withdraws a PTO request the caller submitted themselves, as long as
 * it hasn't already been approved or denied.
 */
function cancelPTORequest(payload) {
  try {
    var auth = requireVerifiedEmail_(payload);
    if (auth.error) return auth;
    var email = auth.email;
    var taskGid = payload.taskGid;
    if (!taskGid) return { error: 'Missing taskGid' };

    var result = asanaRequest('get', '/tasks/' + taskGid + '?opt_fields=notes,memberships.section.name');
    if (result.errors) return { error: result.errors[0].message };

    var notes = (result.data && result.data.notes) || '';
    var m = notes.match(/Requester:\s*([^\n]+)/);
    var requester = m ? m[1].trim().toLowerCase() : '';
    if (requester !== email) return { error: 'You can only cancel your own requests.', code: 'FORBIDDEN' };

    var section = '';
    if (result.data && result.data.memberships && result.data.memberships[0] && result.data.memberships[0].section) {
      section = result.data.memberships[0].section.name || '';
    }
    if (section === 'Approved' || section === 'Denied') {
      return { error: 'This request has already been decided and can no longer be cancelled.' };
    }

    var del = asanaRequest('delete', '/tasks/' + taskGid);
    if (del.errors) return { error: del.errors[0].message };
    return { success: true };
  } catch(e) { return { error: e.toString() }; }
}

/**
 * Adds daysToAdd to the Used column (G) for the given employee email.
 */
function updatePTOUsed(email, daysToAdd) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hrSheet = ss.getSheetByName(ROLES_SHEET);
    if (!hrSheet) return;
    var lastRow = hrSheet.getLastRow();
    if (lastRow < 2) return;
    var data = hrSheet.getRange(2, 1, lastRow - 1, 7).getValues();
    for (var i = 0; i < data.length; i++) {
      var rowEmail = (data[i][1] || '').toString().toLowerCase().trim();
      if (rowEmail === email.toLowerCase().trim()) {
        var currentUsed = parseFloat(data[i][6]) || 0;
        hrSheet.getRange(i + 2, 7).setValue(currentUsed + daysToAdd);
        return;
      }
    }
  } catch(e) { /* silent */ }
}

/**
 * Looks up a section GID by name in the given Asana project. Returns null
 * (not throw) on any failure or no match.
 */
function getSectionGidByName(projectGid, sectionName) {
  try {
    var result = asanaRequest('get', '/projects/' + projectGid + '/sections?opt_fields=gid,name');
    if (result.errors || !result.data) return null;
    for (var i = 0; i < result.data.length; i++) {
      if (result.data[i].name === sectionName) return result.data[i].gid;
    }
    return null;
  } catch(e) { return null; }
}

/**
 * Looks up a section GID by name in the PTO project.
 */
function getPTOSectionGid(sectionName) {
  return getSectionGidByName(ASANA_PTO_PROJECT, sectionName);
}

// -- Time Tracking ------------------------------------------------------------
// Sheet: "Time Tracking"  cols: A=Name, B=Email, C=Date, D=ClockIn, E=ClockOut, F=Hours
var TIME_SHEET = 'Time Tracking';

// Semi-monthly pay periods: 1st-15th and 16th-end of month
function getPeriodBounds(d) {
  var tz    = Session.getScriptTimeZone();
  var year  = parseInt(Utilities.formatDate(d, tz, 'yyyy'));
  var month = parseInt(Utilities.formatDate(d, tz, 'M')) - 1; // 0-indexed
  var day   = parseInt(Utilities.formatDate(d, tz, 'd'));
  var start, end;
  if (day <= 15) {
    start = new Date(year, month, 1);
    end   = new Date(year, month, 15);
  } else {
    start = new Date(year, month, 16);
    end   = new Date(year, month + 1, 0); // day 0 of next month = last day of this month
  }
  return { start: start, end: end };
}

/**
 * Bounds for the semi-monthly period `offset` periods before the period
 * containing refDate (offset <= 0; 0 = current). Steps period-by-period
 * since boundaries alternate 1st-15th / 16th-end-of-month, so month/year
 * rollover needs the Date constructor's overflow normalization rather than
 * simple date math.
 */
function getPeriodBoundsOffset_(refDate, offset) {
  var bounds = getPeriodBounds(refDate);
  var start = bounds.start, end = bounds.end;
  var steps = Math.abs(Math.min(0, offset || 0));
  for (var i = 0; i < steps; i++) {
    if (start.getDate() === 1) {
      var y = start.getFullYear(), m = start.getMonth() - 1;
      start = new Date(y, m, 16);
      end   = new Date(y, m + 1, 0);
    } else {
      start = new Date(start.getFullYear(), start.getMonth(), 1);
      end   = new Date(start.getFullYear(), start.getMonth(), 15);
    }
  }
  return { start: start, end: end };
}

/**
 * True when the CURRENT (still-open) period's last day is a Saturday or
 * Sunday and `today` falls on/after the last business day (Friday) before
 * that close -- the window where payroll must run before the period
 * technically ends, so early self-approval of the still-open period is
 * allowed. periodEnd must be the bounds.end of getPeriodBounds(today).
 */
function isEarlyApprovalWindow_(periodEnd, today) {
  var dow = periodEnd.getDay(); // 0 = Sun, 6 = Sat
  if (dow !== 0 && dow !== 6) return false;
  var lastBizDay = new Date(periodEnd.getFullYear(), periodEnd.getMonth(), periodEnd.getDate() - (dow === 6 ? 1 : 2));
  var t = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  return t.getTime() >= lastBizDay.getTime() && t.getTime() <= periodEnd.getTime();
}

var MONTH_ABBRS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

/** Formats a period's bounds as e.g. "Jul 1 - Jul 15". */
function formatPeriodLabel_(pStart, pEnd) {
  return MONTH_ABBRS[pStart.getMonth()] + ' ' + pStart.getDate() + ' - ' + MONTH_ABBRS[pEnd.getMonth()] + ' ' + pEnd.getDate();
}

// Overtime threshold. This split is advisory (for admin visibility before
// payroll runs), not a certified payroll engine -- a week that straddles a
// pay-period boundary is only evaluated against the days present in the
// current period, same simplification as the documented pay-period-boundary
// limitation for hour bucketing.
var OT_WEEKLY_THRESHOLD = 40;

function getWeekKey_(date) {
  var d = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  d.setDate(d.getDate() - d.getDay()); // back up to the Sunday that starts this week
  return d.getTime();
}

/**
 * Splits an array of { date, hours } day-entries (one period, any order)
 * into { regular, overtime } using only the weekly (>40h/week) threshold --
 * no daily threshold, applied chronologically so weekly accumulation makes
 * sense.
 */
function splitRegularOvertime_(dayEntries) {
  var sorted = dayEntries.slice().sort(function(a, b) { return a.date - b.date; });
  var weekTotals = {};
  var regular = 0, overtime = 0;
  sorted.forEach(function(entry) {
    var weekKey       = getWeekKey_(entry.date);
    var weekSoFar      = weekTotals[weekKey] || 0;
    var roomLeftInWeek = Math.max(0, OT_WEEKLY_THRESHOLD - weekSoFar);
    var weeklyRegular   = Math.min(entry.hours, roomLeftInWeek);
    var weeklyOT        = entry.hours - weeklyRegular;

    regular  += weeklyRegular;
    overtime += weeklyOT;
    weekTotals[weekKey] = weekSoFar + entry.hours;
  });
  return { regular: Math.round(regular * 100) / 100, overtime: Math.round(overtime * 100) / 100 };
}

var TIME_SHEET_HEADERS = ['Employee Name','Email','Date','Clock In','Clock Out','Hours','Clock In Location','Clock Out Location'];

function getTimeSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(TIME_SHEET);
  if (!sh) {
    sh = ss.insertSheet(TIME_SHEET);
    sh.getRange(1, 1, 1, TIME_SHEET_HEADERS.length).setValues([TIME_SHEET_HEADERS]);
    sh.getRange(1, 1, 1, TIME_SHEET_HEADERS.length).setFontWeight('bold').setBackground('#1F3971').setFontColor('#ffffff');
  } else if (sh.getLastColumn() < TIME_SHEET_HEADERS.length) {
    // Additive migration for sheets created before the location columns existed --
    // never touches/reorders the existing A-F columns or their data.
    var startCol   = sh.getLastColumn() + 1;
    var newHeaders = TIME_SHEET_HEADERS.slice(sh.getLastColumn());
    sh.getRange(1, startCol, 1, newHeaders.length).setValues([newHeaders]);
    sh.getRange(1, startCol, 1, newHeaders.length).setFontWeight('bold').setBackground('#1F3971').setFontColor('#ffffff');
  }
  return sh;
}

/** Formats lat/lng into a Google Maps link, or '' if either is missing/invalid. */
function formatLocation_(lat, lng) {
  var la = parseFloat(lat), lo = parseFloat(lng);
  if (isNaN(la) || isNaN(lo)) return '';
  return 'https://maps.google.com/?q=' + la + ',' + lo;
}

function clockIn(payload) {
  try {
    var auth = requireVerifiedEmail_(payload);
    if (auth.error) return auth;
    var email = auth.email;
    var name  = (payload.name || email).toString();
    var sh    = getTimeSheet_();
    var tz    = Session.getScriptTimeZone();
    var now   = new Date();
    var loc   = formatLocation_(payload.lat, payload.lng);

    if (isPeriodApprovedForEmail_(email, formatPeriodLabel_(getPeriodBounds(now).start, getPeriodBounds(now).end))) {
      return { error: 'Your timesheet for this period has already been approved. Contact your manager if you need to log more time.' };
    }

    var lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) return { error: 'Server is busy - try again in a moment.' };
    try {
      // Check for open record
      var lastRow = sh.getLastRow();
      if (lastRow >= 2) {
        var data = sh.getRange(2, 1, lastRow - 1, 5).getValues();
        for (var i = data.length - 1; i >= 0; i--) {
          if ((data[i][1] || '').toString().toLowerCase() === email && !data[i][4]) {
            return { error: 'Already clocked in at ' + Utilities.formatDate(new Date(data[i][3]), tz, 'h:mm a') };
          }
        }
      }

      var today = Utilities.formatDate(now, tz, 'MM/dd/yyyy');
      sh.appendRow([name, email, today, now, '', '', loc, '']);
      return { success: true, clockIn: Utilities.formatDate(now, tz, 'h:mm a') };
    } finally {
      lock.releaseLock();
    }
  } catch(e) { return { error: e.toString() }; }
}

function clockOut(payload) {
  try {
    var auth = requireVerifiedEmail_(payload);
    if (auth.error) return auth;
    var email = auth.email;
    var sh    = getTimeSheet_();
    var tz    = Session.getScriptTimeZone();
    var now   = new Date();
    var loc   = formatLocation_(payload.lat, payload.lng);

    if (isPeriodApprovedForEmail_(email, formatPeriodLabel_(getPeriodBounds(now).start, getPeriodBounds(now).end))) {
      return { error: 'Your timesheet for this period has already been approved. Contact your manager to reopen it if you need to log more time.' };
    }

    var lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) return { error: 'Server is busy - try again in a moment.' };
    try {
      var lastRow = sh.getLastRow();
      if (lastRow < 2) return { error: 'No clock-in record found' };

      var data = sh.getRange(2, 1, lastRow - 1, 5).getValues();
      for (var i = data.length - 1; i >= 0; i--) {
        if ((data[i][1] || '').toString().toLowerCase() === email && !data[i][4]) {
          var clockInTime = new Date(data[i][3]);
          var rawHours = (now - clockInTime) / 3600000;
          // Clamp a negative duration (clock skew / bad manual edit) to 0 rather than
          // writing a bad value into payroll totals -- Clock In > Clock Out is still
          // visible in the raw D/E cells, so findFlaggableShifts_ still catches it.
          var hours = Math.round(Math.max(0, rawHours) * 100) / 100;
          var rowNum = i + 2;
          sh.getRange(rowNum, 5).setValue(now);
          sh.getRange(rowNum, 6).setValue(hours);
          sh.getRange(rowNum, 8).setValue(loc);
          return { success: true, clockOut: Utilities.formatDate(now, tz, 'h:mm a'), hours: hours };
        }
      }
      return { error: 'No open clock-in found' };
    } finally {
      lock.releaseLock();
    }
  } catch(e) { return { error: e.toString() }; }
}

function getClockStatus(payload) {
  try {
    var auth = requireVerifiedEmail_(payload);
    if (auth.error) return auth;
    var email = auth.email;
    var sh    = getTimeSheet_();
    var tz    = Session.getScriptTimeZone();
    var lastRow = sh.getLastRow();

    if (lastRow >= 2) {
      var data = sh.getRange(2, 1, lastRow - 1, 5).getValues();
      for (var i = data.length - 1; i >= 0; i--) {
        if ((data[i][1] || '').toString().toLowerCase() === email && !data[i][4]) {
          return { clockedIn: true, since: Utilities.formatDate(new Date(data[i][3]), tz, 'h:mm a') };
        }
      }
    }
    return { clockedIn: false };
  } catch(e) { return { error: e.toString() }; }
}

function getTimesheet(payload) {
  try {
    var auth = requireVerifiedEmail_(payload);
    if (auth.error) return auth;
    var email      = auth.email;
    var callerRoles = getRoleByEmail(email).effRoles;
    var canSeeAll  = hasAnyRole_(callerRoles, ['admin', 'human_resources']);
    var sh     = getTimeSheet_();
    var tz     = Session.getScriptTimeZone();
    var now    = new Date();
    var bounds = getPeriodBounds(now);
    var pStart = bounds.start;
    var pEnd   = bounds.end;
    var periodLabel = formatPeriodLabel_(pStart, pEnd);

    var myHours      = 0;
    var myDays       = {};
    var myDayEntries = []; // [{ date, hours }] for regular/OT split
    var empMap       = {}; // email -> { name, hours }

    var lastRow = sh.getLastRow();
    if (lastRow >= 2) {
      var data = sh.getRange(2, 1, lastRow - 1, 6).getValues();
      for (var i = 0; i < data.length; i++) {
        var rowEmail = (data[i][1] || '').toString().toLowerCase().trim();
        var rowDate  = data[i][2] ? new Date(data[i][2]) : null;
        var rowHours = parseFloat(data[i][5]) || 0;
        if (!rowDate) continue;
        // Normalize rowDate to midnight for comparison
        var rd = new Date(rowDate.getFullYear(), rowDate.getMonth(), rowDate.getDate());
        if (rd < pStart || rd > pEnd) continue;

        // My rows
        if (rowEmail === email) {
          myHours += rowHours;
          var dayLabel = MONTH_ABBRS[rd.getMonth()] + ' ' + rd.getDate();
          myDays[dayLabel] = Math.round(((myDays[dayLabel] || 0) + rowHours) * 100) / 100;
          var existingEntry = myDayEntries.filter(function(e) { return e.date.getTime() === rd.getTime(); })[0];
          if (existingEntry) { existingEntry.hours += rowHours; }
          else { myDayEntries.push({ date: rd, hours: rowHours }); }
        }

        // All employees (admin / HR)
        if (canSeeAll) {
          if (!empMap[rowEmail]) empMap[rowEmail] = { name: (data[i][0] || rowEmail).toString().trim(), hours: 0 };
          empMap[rowEmail].hours = Math.round((empMap[rowEmail].hours + rowHours) * 100) / 100;
        }
      }
    }

    var allEmployees = [];
    if (canSeeAll) {
      allEmployees = Object.keys(empMap).map(function(e) {
        return { name: empMap[e].name, hours: empMap[e].hours };
      }).sort(function(a, b) { return b.hours - a.hours; });
    }

    var otSplit = splitRegularOvertime_(myDayEntries);

    return {
      myHours:      Math.round(myHours * 100) / 100,
      myRegularHours:  otSplit.regular,
      myOvertimeHours: otSplit.overtime,
      myDays:       myDays,
      periodLabel:  periodLabel,
      allEmployees: allEmployees
    };
  } catch(e) { return { error: e.toString() }; }
}

/**
 * Lean {name, email} list for populating the Office Notes "Assigned To"
 * dropdown. Unlike getEmployees, this isn't gated to admin/HR -- Office
 * Notes is visible to every role, and this only exposes name+email, none
 * of getEmployees' PTO/role data.
 */
function getAssignableEmployees(payload) {
  try {
    var auth = requireVerifiedEmail_(payload);
    if (auth.error) return auth;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(ROLES_SHEET);
    if (!sh) return { error: 'HR sheet not found' };
    var lastRow = sh.getLastRow();
    if (lastRow < 2) return { employees: [] };
    var data = sh.getRange(2, 1, lastRow - 1, 2).getValues();
    var employees = [];
    for (var i = 0; i < data.length; i++) {
      var email = (data[i][1] || '').toString().trim();
      if (!email) continue;
      employees.push({ name: (data[i][0] || '').toString().trim(), email: email });
    }
    return { employees: employees };
  } catch(e) { return { error: e.toString() }; }
}

// ── Admin: Employee Manager ───────────────────────────────────────────────────
function getEmployees(payload) {
  try {
    var auth = authorizeCaller(payload, ['admin', 'human_resources']);
    if (!auth.ok) return { error: auth.error, code: auth.code };
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(ROLES_SHEET);
    if (!sh) return { error: 'HR sheet not found' };
    var lastRow = sh.getLastRow();
    if (lastRow < 2) return { employees: [] };
    var data = sh.getRange(2, 1, lastRow - 1, 7).getValues();
    var employees = [];
    for (var i = 0; i < data.length; i++) {
      if (!data[i][1]) continue;
      employees.push({
        rowIndex:  i + 2,
        name:      (data[i][0] || '').toString().trim(),
        email:     (data[i][1] || '').toString().trim(),
        phone:     (data[i][2] || '').toString().trim(),
        role:      (data[i][3] || '').toString().trim(),
        allotted:  parseFloat(data[i][5]) || 0,
        used:      parseFloat(data[i][6]) || 0,
        remaining: (parseFloat(data[i][5]) || 0) - (parseFloat(data[i][6]) || 0)
      });
    }
    return { employees: employees };
  } catch(e) { return { error: e.toString() }; }
}

function addEmployee(payload) {
  try {
    var auth = authorizeCaller(payload, ['admin', 'human_resources']);
    if (!auth.ok) return { error: auth.error, code: auth.code };
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(ROLES_SHEET);
    if (!sh) return { error: 'HR sheet not found' };
    var email = (payload.email || '').toLowerCase().trim();
    if (!email || !payload.name) return { error: 'Name and email are required' };
    var newRoleList = filterValidRoles_(parseRoleList_(payload.role));
    if (!newRoleList.length) newRoleList = ['runner'];
    if (isOwnerEmail(email) && newRoleList.indexOf('aidan') === -1) {
      return { error: 'This account is protected and must be added as aidan.', code: 'OWNER_PROTECTED' };
    }
    // Lock around the existing-email check + append -- without this, two
    // near-simultaneous submissions for the same new hire could both pass
    // the check before either had appended, creating two rows.
    var lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) return { error: 'Server is busy - try again in a moment.' };
    try {
      var lastRow = sh.getLastRow();
      if (lastRow >= 2) {
        var existing = sh.getRange(2, 1, lastRow - 1, 2).getValues();
        for (var i = 0; i < existing.length; i++) {
          if ((existing[i][1] || '').toLowerCase().trim() === email) {
            return { error: 'An employee with that email already exists' };
          }
        }
      }
      sh.appendRow([
        payload.name.trim(),
        email,
        (payload.phone || '').trim(),
        newRoleList.join(','),
        (payload.password || '').trim(),
        parseFloat(payload.allotted) || 0,
        0
      ]);
      invalidateRolesCache_();
      return { success: true };
    } finally {
      lock.releaseLock();
    }
  } catch(e) { return { error: e.toString() }; }
}

function updateEmployee(payload) {
  try {
    var auth = authorizeCaller(payload, ['admin', 'human_resources']);
    if (!auth.ok) return { error: auth.error, code: auth.code };
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(ROLES_SHEET);
    if (!sh) return { error: 'HR sheet not found' };
    var email = (payload.email || '').toLowerCase().trim();
    var lastRow = sh.getLastRow();
    if (lastRow < 2) return { error: 'Employee not found' };
    var data = sh.getRange(2, 1, lastRow - 1, 7).getValues();
    for (var i = 0; i < data.length; i++) {
      if ((data[i][1] || '').toLowerCase().trim() === email) {
        var newRoleList = payload.role !== undefined ? filterValidRoles_(parseRoleList_(payload.role)) : undefined;
        if (newRoleList !== undefined) {
          if (isOwnerEmail(email) && newRoleList.indexOf('aidan') === -1) {
            return { error: 'This account is protected and must remain aidan.', code: 'OWNER_PROTECTED' };
          }
          var currentEffRoles = normalizeRoleList_(parseRoleList_(data[i][3]));
          var newEffRoles     = normalizeRoleList_(newRoleList);
          if (currentEffRoles.indexOf('admin') !== -1 && newEffRoles.indexOf('admin') === -1 && countAdminRows(data) <= 1) {
            return { error: 'Cannot demote the last remaining admin.', code: 'LAST_ADMIN_PROTECTED' };
          }
          if (!newRoleList.length) newRoleList = ['runner'];
        }
        var row = i + 2;
        if (payload.name     !== undefined) sh.getRange(row, 1).setValue(payload.name);
        if (payload.phone    !== undefined) sh.getRange(row, 3).setValue(payload.phone);
        if (newRoleList      !== undefined) sh.getRange(row, 4).setValue(newRoleList.join(','));
        if (payload.password !== undefined && payload.password !== '') sh.getRange(row, 5).setValue(payload.password);
        if (payload.allotted !== undefined) sh.getRange(row, 6).setValue(parseFloat(payload.allotted) || 0);
        invalidateRolesCache_();
        return { success: true };
      }
    }
    return { error: 'Employee not found' };
  } catch(e) { return { error: e.toString() }; }
}

function removeEmployee(payload) {
  try {
    var auth = authorizeCaller(payload, ['admin', 'human_resources']);
    if (!auth.ok) return { error: auth.error, code: auth.code };
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(ROLES_SHEET);
    if (!sh) return { error: 'HR sheet not found' };
    var email = (payload.email || '').toLowerCase().trim();
    var lastRow = sh.getLastRow();
    if (lastRow < 2) return { error: 'Employee not found' };
    var data = sh.getRange(2, 1, lastRow - 1, 7).getValues();
    for (var i = 0; i < data.length; i++) {
      if ((data[i][1] || '').toLowerCase().trim() === email) {
        if (isOwnerEmail(email)) {
          return { error: 'This account is protected and cannot be removed.', code: 'OWNER_PROTECTED' };
        }
        var currentEffRoles = normalizeRoleList_(parseRoleList_(data[i][3]));
        if (currentEffRoles.indexOf('admin') !== -1 && countAdminRows(data) <= 1) {
          return { error: 'Cannot remove the last remaining admin.', code: 'LAST_ADMIN_PROTECTED' };
        }
        sh.deleteRow(i + 2);
        invalidateRolesCache_();
        return { success: true };
      }
    }
    return { error: 'Employee not found' };
  } catch(e) { return { error: e.toString() }; }
}

// ── Admin: PTO Overview ───────────────────────────────────────────────────────
function getPTOOverview(payload) {
  try {
    var auth = authorizeCaller(payload, ['admin', 'human_resources']);
    if (!auth.ok) return { error: auth.error, code: auth.code };
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(ROLES_SHEET);
    var balances = [];
    if (sh && sh.getLastRow() >= 2) {
      var data = sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues();
      for (var i = 0; i < data.length; i++) {
        if (!data[i][1]) continue;
        var allotted  = parseFloat(data[i][5]) || 0;
        var used      = parseFloat(data[i][6]) || 0;
        balances.push({ name: (data[i][0] || '').toString().trim(), email: (data[i][1] || '').toString().trim(), allotted: allotted, used: used, remaining: allotted - used });
      }
    }
    var result = asanaRequest('get', '/projects/' + ASANA_PTO_PROJECT + '/tasks?opt_fields=gid,name,notes,memberships.section.name&limit=100');
    var requests = [];
    if (!result.errors) {
      (result.data || []).forEach(function(task) {
        var notes = task.notes || '';
        var section = (task.memberships && task.memberships[0] && task.memberships[0].section) ? task.memberships[0].section.name : '';
        var parseField = function(label) { var m = notes.match(new RegExp(label + ':\s*([^\n]+)')); return m ? m[1].trim() : ''; };
        requests.push({
          gid:    task.gid,
          name:   parseField('Name') || task.name,
          email:  parseField('Requester'),
          dates:  parseField('Dates'),
          days:   parseFloat(parseField('Days')) || 0,
          reason: parseField('Reason'),
          status: section === 'Approved' ? 'approved' : section === 'Denied' ? 'denied' : 'pending'
        });
      });
    }
    return { balances: balances, requests: requests };
  } catch(e) { return { error: e.toString() }; }
}

// ── Payroll: shift flagging ───────────────────────────────────────────────────
// A shift that shouldn't be silently summed into payroll totals: still-open
// (forgot to clock out) or closed with an abnormal/negative duration.
var FLAG_OPEN_SHIFT_HOURS = 16; // an unclosed shift open this long needs review
var FLAG_LONG_SHIFT_HOURS = 16; // a closed shift this long needs review

/**
 * Scans the full Time Tracking sheet and returns flaggable rows, each with
 * rowIndex so callers can exclude that exact row from totals. Open shifts
 * are checked sheet-wide (not just the current period) so a forgotten
 * clock-out from an earlier period doesn't quietly disappear once its
 * period rolls off getPayrollSummary's date filter. Closed-shift anomalies
 * are scoped to [pStart, pEnd] to match the totals view.
 */
function findFlaggableShifts_(sh, pStart, pEnd) {
  var flagged = [];
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return flagged;
  var tz  = Session.getScriptTimeZone();
  var now = new Date();
  var data = sh.getRange(2, 1, lastRow - 1, 6).getValues();
  for (var i = 0; i < data.length; i++) {
    var rowIndex = i + 2;
    var name     = (data[i][0] || '').toString().trim();
    var email    = (data[i][1] || '').toString().trim();
    var rowDate  = data[i][2] ? new Date(data[i][2]) : null;
    var clockIn  = data[i][3] ? new Date(data[i][3]) : null;
    var clockOut = data[i][4] ? new Date(data[i][4]) : null;
    if (!email || !clockIn) continue;

    if (!clockOut) {
      var hoursOpen = (now - clockIn) / 3600000;
      if (hoursOpen > FLAG_OPEN_SHIFT_HOURS) {
        flagged.push({
          rowIndex: rowIndex, name: name, email: email,
          date: rowDate ? Utilities.formatDate(rowDate, tz, 'MM/dd/yyyy') : '',
          clockIn: Utilities.formatDate(clockIn, tz, 'MM/dd/yyyy h:mm a'),
          reason: 'Still clocked in - forgot to clock out?'
        });
      }
      continue;
    }

    if (!rowDate) continue;
    var rd = new Date(rowDate.getFullYear(), rowDate.getMonth(), rowDate.getDate());
    if (rd < pStart || rd > pEnd) continue;

    if (clockOut < clockIn) {
      flagged.push({
        rowIndex: rowIndex, name: name, email: email,
        date: Utilities.formatDate(rd, tz, 'MM/dd/yyyy'), clockIn: Utilities.formatDate(clockIn, tz, 'h:mm a'),
        reason: 'Clock out is before clock in - check for a bad edit'
      });
      continue;
    }

    var hours = (clockOut - clockIn) / 3600000;
    if (hours > FLAG_LONG_SHIFT_HOURS) {
      flagged.push({
        rowIndex: rowIndex, name: name, email: email,
        date: Utilities.formatDate(rd, tz, 'MM/dd/yyyy'), clockIn: Utilities.formatDate(clockIn, tz, 'h:mm a'),
        reason: 'Shift over ' + FLAG_LONG_SHIFT_HOURS + ' hours - verify'
      });
    }
  }
  return flagged;
}

// ── Payroll: period approval/lock ─────────────────────────────────────────────
var PAYROLL_APPROVALS_SHEET  = 'Payroll Approvals';
var PAYROLL_APPROVALS_HEADERS = ['Period Label', 'Employee Email', 'Approved By', 'Approved At', 'Employee Approved At', 'Employee Note', 'Employee PDF URL'];

/**
 * True if an ADMIN/HR has approved this period for this employee (column D),
 * OR the employee has self-approved it (column E). Column E is normally
 * informational-only and must never block clocking in/out -- but this
 * function is only ever called with the CALLER'S CURRENT (still-open)
 * period's label (see clockIn/clockOut below), and approveMyTimesheet only
 * ever writes a row under the current period's label when
 * isEarlyApprovalWindow_ let it through. So a column-E hit here can only mean
 * an eligible early self-approval, which is intentionally locking.
 */
function isPeriodApprovedForEmail_(email, periodLabel) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(PAYROLL_APPROVALS_SHEET);
  if (!sheet) return false;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  email = (email || '').toString().toLowerCase().trim();
  var data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  for (var i = 0; i < data.length; i++) {
    if ((data[i][0] || '') === periodLabel && (data[i][1] || '').toString().toLowerCase().trim() === email && (data[i][3] || data[i][4])) return true;
  }
  return false;
}

/** email(lowercase) -> { approvedBy, approvedAt, employeeApprovedAt, employeeNote } for every approval on file for periodLabel. */
function getApprovalMap_(periodLabel) {
  var map = {};
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(PAYROLL_APPROVALS_SHEET);
  if (!sheet) return map;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return map;
  var tz = Session.getScriptTimeZone();
  var data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  for (var i = 0; i < data.length; i++) {
    if ((data[i][0] || '') !== periodLabel) continue;
    var email = (data[i][1] || '').toString().toLowerCase().trim();
    if (!email) continue;
    map[email] = {
      approvedBy:         (data[i][2] || '').toString(),
      approvedAt:         data[i][3] ? Utilities.formatDate(new Date(data[i][3]), tz, 'MM/dd/yyyy h:mm a') : '',
      employeeApprovedAt: data[i][4] ? Utilities.formatDate(new Date(data[i][4]), tz, 'MM/dd/yyyy h:mm a') : '',
      employeeNote:       (data[i][5] || '').toString(),
      employeePdfUrl:     (data[i][6] || '').toString()
    };
  }
  return map;
}

/**
 * Approves (or re-approves) a pay period for one employee -- the current
 * period by default, or a past one via periodOffset (never a future one).
 * Idempotent per period+employee.
 */
function approveTimesheet(payload) {
  try {
    var auth = authorizeCaller(payload, ['admin', 'human_resources']);
    if (!auth.ok) return { error: auth.error, code: auth.code };
    var employeeEmail = (payload.employeeEmail || '').toString().toLowerCase().trim();
    if (!employeeEmail) return { error: 'Missing employeeEmail.' };

    var offset = parseInt(payload.periodOffset, 10);
    if (isNaN(offset) || offset > 0) offset = 0;
    var bounds = getPeriodBoundsOffset_(new Date(), offset);
    var periodLabel = formatPeriodLabel_(bounds.start, bounds.end);
    var sheet = ensureSheetWithHeaders_(PAYROLL_APPROVALS_SHEET, PAYROLL_APPROVALS_HEADERS);

    var lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) return { error: 'Server is busy - try again in a moment.' };
    try {
      var now = new Date();
      var lastRow = sheet.getLastRow();
      if (lastRow >= 2) {
        var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
        for (var i = 0; i < data.length; i++) {
          if ((data[i][0] || '') === periodLabel && (data[i][1] || '').toString().toLowerCase().trim() === employeeEmail) {
            sheet.getRange(i + 2, 3).setValue(auth.email);
            sheet.getRange(i + 2, 4).setValue(now);
            return { success: true, approvedBy: auth.email, approvedAt: Utilities.formatDate(now, Session.getScriptTimeZone(), 'MM/dd/yyyy h:mm a') };
          }
        }
      }
      sheet.appendRow([periodLabel, employeeEmail, auth.email, now]);
      return { success: true, approvedBy: auth.email, approvedAt: Utilities.formatDate(now, Session.getScriptTimeZone(), 'MM/dd/yyyy h:mm a') };
    } finally {
      lock.releaseLock();
    }
  } catch(e) { return { error: e.toString() }; }
}

/**
 * Employee self-approval of their OWN hours for a CLOSED pay period, with an
 * optional note. Unlike approveTimesheet this is never admin/HR-gated -- it's
 * only ever keyed to the caller's own email (never accepts an employeeEmail
 * param), which is what makes it safe to expose to any signed-in user.
 * Purely informational for payroll -- doesn't block or affect
 * approveTimesheet/unapproveTimesheet in any way. Idempotent/re-callable so
 * an employee can update their note later without "unapproving" first.
 *
 * The currently-open period (offset 0) is normally rejected (falls back to
 * -1, the most recently closed period) -- except when isEarlyApprovalWindow_
 * says the open period's last day is a Sat/Sun and today is on/after the
 * Friday before it, e.g. payroll must run before a weekend close. In that
 * case the resulting row DOES gate clock in/out (see isPeriodApprovedForEmail_).
 */
function approveMyTimesheet(payload) {
  try {
    var auth = requireVerifiedEmail_(payload);
    if (auth.error) return auth;
    var email = auth.email;
    var note = (payload.note || '').toString().trim();
    var pdfUrl = (payload.pdfUrl || '').toString().trim();
    var now = new Date();

    var offset = parseInt(payload.periodOffset, 10);
    if (isNaN(offset) || offset > 0) offset = -1;
    if (offset === 0 && !isEarlyApprovalWindow_(getPeriodBounds(now).end, now)) offset = -1;
    var bounds = getPeriodBoundsOffset_(now, offset);
    var periodLabel = formatPeriodLabel_(bounds.start, bounds.end);
    var sheet = ensureSheetWithHeaders_(PAYROLL_APPROVALS_SHEET, PAYROLL_APPROVALS_HEADERS);

    var lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) return { error: 'Server is busy - try again in a moment.' };
    try {
      var lastRow = sheet.getLastRow();
      if (lastRow >= 2) {
        var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
        for (var i = 0; i < data.length; i++) {
          if ((data[i][0] || '') === periodLabel && (data[i][1] || '').toString().toLowerCase().trim() === email) {
            sheet.getRange(i + 2, 5).setValue(now);
            sheet.getRange(i + 2, 6).setValue(note);
            sheet.getRange(i + 2, 7).setValue(pdfUrl);
            return { success: true, approvedAt: Utilities.formatDate(now, Session.getScriptTimeZone(), 'MM/dd/yyyy h:mm a') };
          }
        }
      }
      sheet.appendRow([periodLabel, email, '', '', now, note, pdfUrl]);
      return { success: true, approvedAt: Utilities.formatDate(now, Session.getScriptTimeZone(), 'MM/dd/yyyy h:mm a') };
    } finally {
      lock.releaseLock();
    }
  } catch(e) { return { error: e.toString() }; }
}

/**
 * Day-by-day detail of the CALLER'S OWN hours for a given pay period --
 * powers both the "Approve Hours" period-review screen (usually the most
 * recently-closed period, offset -1, but offset 0 when earlyApprovalEligible
 * lets the still-open period be reviewed/approved early) and the read-only
 * "My Timesheets" history browser (any offset <= 0, including the current
 * open period). Every calendar day in the period is included (blank
 * clockIn/clockOut for days with no shift on record), so a missed punch is
 * visible and correctable, not just days that already have a row.
 */
function getMyPeriodDetail(payload) {
  try {
    var auth = requireVerifiedEmail_(payload);
    if (auth.error) return auth;
    var email = auth.email;
    var tz = Session.getScriptTimeZone();

    var offset = parseInt(payload.periodOffset, 10);
    if (isNaN(offset)) offset = -1;
    if (offset > 0) offset = 0; // never allow a future period
    var bounds = getPeriodBoundsOffset_(new Date(), offset);
    var pStart = bounds.start, pEnd = bounds.end;
    var periodLabel = formatPeriodLabel_(pStart, pEnd);

    var byDate = {}; // 'yyyy-MM-dd' -> { clockIn, clockOut, hours }
    var dayEntries = []; // for splitRegularOvertime_
    var sh = getTimeSheet_();
    var lastRow = sh.getLastRow();
    if (lastRow >= 2) {
      var data = sh.getRange(2, 1, lastRow - 1, 6).getValues();
      for (var i = 0; i < data.length; i++) {
        var rowEmail = (data[i][1] || '').toString().toLowerCase().trim();
        if (rowEmail !== email) continue;
        var rowDate = data[i][2] ? new Date(data[i][2]) : null;
        if (!rowDate) continue;
        var rd = new Date(rowDate.getFullYear(), rowDate.getMonth(), rowDate.getDate());
        if (rd < pStart || rd > pEnd) continue;
        var key = Utilities.formatDate(rd, tz, 'yyyy-MM-dd');
        var rowHours = parseFloat(data[i][5]) || 0;
        byDate[key] = {
          clockIn:  data[i][3] ? Utilities.formatDate(new Date(data[i][3]), tz, 'HH:mm') : '',
          clockOut: data[i][4] ? Utilities.formatDate(new Date(data[i][4]), tz, 'HH:mm') : '',
          hours: rowHours
        };
        dayEntries.push({ date: rd, hours: rowHours });
      }
    }

    var days = [];
    var totalHours = 0;
    for (var d = new Date(pStart); d.getTime() <= pEnd.getTime(); d.setDate(d.getDate() + 1)) {
      var key = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
      var entry = byDate[key] || { clockIn: '', clockOut: '', hours: 0 };
      totalHours += entry.hours;
      days.push({
        date: key,
        dayLabel: MONTH_ABBRS[d.getMonth()] + ' ' + d.getDate(),
        clockIn: entry.clockIn,
        clockOut: entry.clockOut,
        hours: entry.hours
      });
    }

    var otSplit = splitRegularOvertime_(dayEntries);
    var myApproval = getApprovalMap_(periodLabel)[email] || {};
    var isCurrentPeriod = offset === 0;

    return {
      periodLabel: periodLabel,
      periodOffset: offset,
      isCurrentPeriod: isCurrentPeriod,
      earlyApprovalEligible: isCurrentPeriod && isEarlyApprovalWindow_(pEnd, new Date()),
      totalHours: Math.round(totalHours * 100) / 100,
      regularHours: otSplit.regular,
      overtimeHours: otSplit.overtime,
      days: days,
      approved: !!myApproval.approvedAt,
      approvedAt: myApproval.approvedAt || '',
      employeeApproved: !!myApproval.employeeApprovedAt,
      employeeApprovedAt: myApproval.employeeApprovedAt || '',
      employeeNote: myApproval.employeeNote || '',
      employeePdfUrl: myApproval.employeePdfUrl || ''
    };
  } catch(e) { return { error: e.toString() }; }
}

/**
 * Reopens whichever period (current or a past periodOffset) for one
 * employee (admin approval only), letting them clock in/out again and
 * admins re-approve later. Clears the admin approval columns (C/D) instead
 * of deleting the row, so an employee's own self-approval/note in columns
 * E/F -- written independently via approveMyTimesheet -- survives an admin
 * reopen.
 */
function unapproveTimesheet(payload) {
  try {
    var auth = authorizeCaller(payload, ['admin', 'human_resources']);
    if (!auth.ok) return { error: auth.error, code: auth.code };
    var employeeEmail = (payload.employeeEmail || '').toString().toLowerCase().trim();
    if (!employeeEmail) return { error: 'Missing employeeEmail.' };

    var offset = parseInt(payload.periodOffset, 10);
    if (isNaN(offset) || offset > 0) offset = 0;
    var bounds = getPeriodBoundsOffset_(new Date(), offset);
    var periodLabel = formatPeriodLabel_(bounds.start, bounds.end);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PAYROLL_APPROVALS_SHEET);
    if (!sheet) return { success: true };

    var lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) return { error: 'Server is busy - try again in a moment.' };
    try {
      var lastRow = sheet.getLastRow();
      if (lastRow >= 2) {
        var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
        for (var i = data.length - 1; i >= 0; i--) {
          if ((data[i][0] || '') === periodLabel && (data[i][1] || '').toString().toLowerCase().trim() === employeeEmail) {
            sheet.getRange(i + 2, 3).setValue('');
            sheet.getRange(i + 2, 4).setValue('');
          }
        }
      }
      return { success: true };
    } finally {
      lock.releaseLock();
    }
  } catch(e) { return { error: e.toString() }; }
}

// ── Pay-period approval reminder (time-based trigger) ─────────────────────────
// Fires daily but only actually sends a push on the day a new pay period
// starts (i.e. the day right after one just closed), about the period that
// just ended, to whichever employees haven't yet self-approved it (via
// approveMyTimesheet). The in-app "Approve Hours" card (see
// getMyPeriodDetail) is computed independently on every Account-tab load,
// so a missed/failed push doesn't hide the prompt.
function checkPayPeriodApprovalReminder() {
  try {
    var today = new Date();
    today = new Date(today.getFullYear(), today.getMonth(), today.getDate());
    var yesterday = new Date(today.getFullYear(), today.getMonth(), today.getDate() - 1);
    var todayBounds = getPeriodBounds(today);
    var yesterdayBounds = getPeriodBounds(yesterday);
    if (todayBounds.start.getTime() === yesterdayBounds.start.getTime()) return; // not a period-boundary day

    var periodLabel = formatPeriodLabel_(yesterdayBounds.start, yesterdayBounds.end);
    var approvalMap = getApprovalMap_(periodLabel);
    var roster = getAssignableEmployees().employees || [];
    var targets = roster.map(function(r) { return (r.email || '').toString().toLowerCase().trim(); })
      .filter(function(e) { return e && !(approvalMap[e] && approvalMap[e].employeeApprovedAt); });

    if (targets.length) {
      sendPushNotification(targets, 'Approve your hours', 'Please check and approve your hours for ' + periodLabel + '.', '/');
    }
  } catch (e) {
    // Trigger context -- no caller to report an error to.
  }
}

/**
 * Run this ONCE from the Apps Script editor (or `clasp run
 * createPayPeriodReminderTrigger`) after deploying, to install the daily
 * check. clasp push never installs triggers on its own. Safe to re-run --
 * clears any prior trigger on the same handler first.
 */
function createPayPeriodReminderTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'checkPayPeriodApprovalReminder') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('checkPayPeriodApprovalReminder').timeBased().everyDays(1).atHour(8).create();
}

// ── Time Corrections ──────────────────────────────────────────────────────────
// Sheet: "Time Corrections" -- lets an employee ask for a past clock in/out to
// be fixed (clocked in late, forgot to clock out, forgot to punch at all) and
// an admin/HR reviewer to approve (writes the fix into Time Tracking) or deny it.
var TIME_CORRECTIONS_SHEET = 'Time Corrections';
var TIME_CORRECTIONS_HEADERS = [
  'Request ID', 'Submitted At', 'Employee Name', 'Employee Email', 'Shift Date',
  'Original Clock In', 'Original Clock Out', 'Requested Clock In', 'Requested Clock Out',
  'Reason', 'Status', 'Reviewed By', 'Reviewed At', 'Review Note'
];

function getTimeCorrectionsSheet_() {
  return ensureSheetWithHeaders_(TIME_CORRECTIONS_SHEET, TIME_CORRECTIONS_HEADERS);
}

/** Combines a 'YYYY-MM-DD' date string and an 'HH:MM' (24h) time string into one Date. */
function combineDateTime_(dateStr, timeStr) {
  var dp = dateStr.split('-');
  var tp = timeStr.split(':');
  return new Date(parseInt(dp[0], 10), parseInt(dp[1], 10) - 1, parseInt(dp[2], 10), parseInt(tp[0], 10), parseInt(tp[1], 10));
}

/**
 * Sheets silently auto-converts plain "YYYY-MM-DD" / "HH:MM" strings written
 * into a cell into real Date objects if they look like a date/time. Format
 * those properly instead of calling .toString() on a raw Date, which prints
 * the full "Thu Jul 30 2026 00:00:00 GMT-0600 (...)" form.
 */
function formatCorrectionDate_(v, tz) {
  return (v instanceof Date) ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : (v || '').toString();
}
function formatCorrectionTime_(v, tz) {
  return (v instanceof Date) ? Utilities.formatDate(v, tz, 'h:mm a') : (v || '').toString();
}
/** Same Date-vs-string handling as formatCorrectionTime_, but in the 24h "HH:mm" shape combineDateTime_() expects. */
function formatCorrectionTime24_(v, tz) {
  return (v instanceof Date) ? Utilities.formatDate(v, tz, 'HH:mm') : (v || '').toString();
}

function formatCorrectionRow_(row, tz) {
  return {
    id:                (row[0] || '').toString(),
    submittedAt:       row[1] ? Utilities.formatDate(new Date(row[1]), tz, 'MM/dd/yyyy h:mm a') : '',
    employeeName:      (row[2] || '').toString(),
    employeeEmail:     (row[3] || '').toString(),
    date:              formatCorrectionDate_(row[4], tz),
    originalClockIn:   formatCorrectionTime_(row[5], tz),
    originalClockOut:  formatCorrectionTime_(row[6], tz),
    requestedClockIn:  formatCorrectionTime_(row[7], tz),
    requestedClockOut: formatCorrectionTime_(row[8], tz),
    reason:            (row[9]  || '').toString(),
    status:            (row[10] || 'pending').toString(),
    reviewedBy:        (row[11] || '').toString(),
    reviewedAt:        row[12] ? Utilities.formatDate(new Date(row[12]), tz, 'MM/dd/yyyy h:mm a') : '',
    reviewNote:        (row[13] || '').toString()
  };
}

/** Finds a Time Corrections row by its Request ID. Returns { rowNum, row } or null. */
function findCorrectionRow_(sheet, requestId) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  var data = sheet.getRange(2, 1, lastRow - 1, TIME_CORRECTIONS_HEADERS.length).getValues();
  for (var i = 0; i < data.length; i++) {
    if ((data[i][0] || '').toString() === requestId) return { rowNum: i + 2, row: data[i] };
  }
  return null;
}

/** Finds the Time Tracking row for email+date (midnight-normalized). Returns 1-based row number or -1. */
function findTimeTrackingRowForDate_(email, targetMidnight) {
  var sh = getTimeSheet_();
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return -1;
  var data = sh.getRange(2, 1, lastRow - 1, 6).getValues();
  for (var i = 0; i < data.length; i++) {
    var rowEmail = (data[i][1] || '').toString().toLowerCase().trim();
    if (rowEmail !== email) continue;
    var rowDate = data[i][2] ? new Date(data[i][2]) : null;
    if (!rowDate) continue;
    var rd = new Date(rowDate.getFullYear(), rowDate.getMonth(), rowDate.getDate());
    if (rd.getTime() === targetMidnight.getTime()) return i + 2;
  }
  return -1;
}

/**
 * Looks up what's currently on record for the caller on a given date, so the
 * correction form can prefill "what's there now" instead of starting blank.
 * Returns 24h 'HH:mm' strings (or '' if that side was never punched).
 */
function getShiftForDate(payload) {
  try {
    var auth = requireVerifiedEmail_(payload);
    if (auth.error) return auth;
    var email = auth.email;
    var dateStr = (payload.date || '').toString().trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) return { error: 'Invalid date' };

    var dp = dateStr.split('-');
    var targetMidnight = new Date(parseInt(dp[0], 10), parseInt(dp[1], 10) - 1, parseInt(dp[2], 10));
    var matchRow = findTimeTrackingRowForDate_(email, targetMidnight);
    if (matchRow === -1) return { clockIn: '', clockOut: '' };

    var tz = Session.getScriptTimeZone();
    var vals = getTimeSheet_().getRange(matchRow, 4, 1, 2).getValues()[0];
    return {
      clockIn:  vals[0] ? Utilities.formatDate(new Date(vals[0]), tz, 'HH:mm') : '',
      clockOut: vals[1] ? Utilities.formatDate(new Date(vals[1]), tz, 'HH:mm') : ''
    };
  } catch(e) { return { error: e.toString() }; }
}

/** Employee-submitted request to fix a past clock in/out. Requires at least one corrected time and a reason. */
function submitTimeCorrection(payload) {
  try {
    var auth = requireVerifiedEmail_(payload);
    if (auth.error) return auth;
    var email    = auth.email;
    var name     = (payload.name || email).toString();
    var dateStr  = (payload.date || '').toString().trim();   // 'YYYY-MM-DD' from <input type=date>
    var clockIn  = (payload.clockIn  || '').toString().trim(); // 'HH:MM' 24h from <input type=time>
    var clockOut = (payload.clockOut || '').toString().trim();
    var reason   = (payload.reason || '').toString().trim();

    if (!/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) return { error: 'Please choose a valid shift date.' };
    if (!clockIn && !clockOut) return { error: 'Enter a corrected clock in and/or clock out time.' };
    if (!reason) return { error: 'Please explain why this correction is needed.' };

    var dp = dateStr.split('-');
    var targetMidnight = new Date(parseInt(dp[0], 10), parseInt(dp[1], 10) - 1, parseInt(dp[2], 10));
    var todayMidnight = new Date();
    todayMidnight = new Date(todayMidnight.getFullYear(), todayMidnight.getMonth(), todayMidnight.getDate());
    if (targetMidnight.getTime() > todayMidnight.getTime()) return { error: 'Shift date cannot be in the future.' };

    var tz = Session.getScriptTimeZone();

    // Best-effort snapshot of what's currently on record, purely for the reviewer's context.
    var originalIn = '', originalOut = '';
    var matchRow = findTimeTrackingRowForDate_(email, targetMidnight);
    if (matchRow !== -1) {
      var tsSheet = getTimeSheet_();
      var existing = tsSheet.getRange(matchRow, 4, 1, 2).getValues()[0];
      originalIn  = existing[0] ? Utilities.formatDate(new Date(existing[0]), tz, 'h:mm a') : '';
      originalOut = existing[1] ? Utilities.formatDate(new Date(existing[1]), tz, 'h:mm a') : '';
    }

    var sheet = getTimeCorrectionsSheet_();
    var lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) return { error: 'Server is busy - try again in a moment.' };
    try {
      // Idempotency guard -- see createPO() for the full rationale.
      var idemKey = (payload.idempotencyKey || '').toString().trim();
      var cache = CacheService.getScriptCache();
      var cacheKey = idemKey ? ('idem_timecorrection_' + idemKey) : null;
      if (cacheKey) {
        var cached = null;
        try { cached = cache.get(cacheKey); } catch (e) {}
        if (cached) return JSON.parse(cached);
      }

      sheet.appendRow([
        Utilities.getUuid(), new Date(), name, email, dateStr,
        originalIn, originalOut, clockIn, clockOut, reason,
        'pending', '', '', ''
      ]);
      var result = { success: true };
      if (cacheKey) { try { cache.put(cacheKey, JSON.stringify(result), 300); } catch (e) {} }
      return result;
    } finally {
      lock.releaseLock();
    }
  } catch(e) { return { error: e.toString() }; }
}

/**
 * Submits corrections for MULTIPLE days at once (e.g. from the period-review
 * screen where an employee edits several days' times before submitting),
 * sharing one reason across the whole batch. Each item becomes its own
 * 'pending' row in the same Time Corrections sheet submitTimeCorrection
 * writes to -- reviewed/approved individually by admin exactly like a
 * single-day request, no changes needed on the admin side.
 */
function submitTimeCorrectionsBatch(payload) {
  try {
    var auth = requireVerifiedEmail_(payload);
    if (auth.error) return auth;
    var email = auth.email;
    var name  = (payload.name || email).toString();
    var reason = (payload.reason || '').toString().trim();
    var corrections = Array.isArray(payload.corrections) ? payload.corrections : [];

    if (!reason) return { error: 'Please explain why these corrections are needed.' };
    if (!corrections.length) return { error: 'No changed days to submit.' };

    var tz = Session.getScriptTimeZone();
    var todayMidnight = new Date();
    todayMidnight = new Date(todayMidnight.getFullYear(), todayMidnight.getMonth(), todayMidnight.getDate());

    var items = [];
    for (var c = 0; c < corrections.length; c++) {
      var item = corrections[c] || {};
      var dateStr  = (item.date || '').toString().trim();
      var clockIn  = (item.clockIn  || '').toString().trim();
      var clockOut = (item.clockOut || '').toString().trim();
      if (!/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) return { error: 'Invalid date in corrections: ' + dateStr };
      if (!clockIn && !clockOut) continue; // nothing changed for this day, skip

      var dp = dateStr.split('-');
      var targetMidnight = new Date(parseInt(dp[0], 10), parseInt(dp[1], 10) - 1, parseInt(dp[2], 10));
      if (targetMidnight.getTime() > todayMidnight.getTime()) return { error: 'Shift date cannot be in the future: ' + dateStr };

      items.push({ dateStr: dateStr, targetMidnight: targetMidnight, clockIn: clockIn, clockOut: clockOut });
    }
    if (!items.length) return { error: 'No changed days to submit.' };

    // Read the Time Tracking sheet ONCE, before the lock, and match every
    // correction item against this one in-memory pass, instead of calling
    // findTimeTrackingRowForDate_() (a full sheet re-read) once per item --
    // that used to happen *inside* the lock, so a multi-day batch held the
    // global script lock for N full-sheet scans, blocking every other
    // lock-guarded action app-wide (createPO, clockIn/clockOut, etc.) for
    // the duration. This read is safe to do outside the lock since it's
    // read-only; only the appendRow writes below need lock protection.
    var tsSheet = getTimeSheet_();
    var tsLastRow = tsSheet.getLastRow();
    var tsRows = tsLastRow >= 2 ? tsSheet.getRange(2, 1, tsLastRow - 1, 6).getValues() : [];
    var origByKey = {}; // 'email|midnightTimestamp' -> { in, out }
    for (var t = 0; t < tsRows.length; t++) {
      var rowEmail = (tsRows[t][1] || '').toString().toLowerCase().trim();
      var rowDate  = tsRows[t][2] ? new Date(tsRows[t][2]) : null;
      if (!rowEmail || !rowDate) continue;
      var rd = new Date(rowDate.getFullYear(), rowDate.getMonth(), rowDate.getDate());
      origByKey[rowEmail + '|' + rd.getTime()] = { in: tsRows[t][3], out: tsRows[t][4] };
    }

    var sheet = getTimeCorrectionsSheet_();
    var lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) return { error: 'Server is busy - try again in a moment.' };
    try {
      // Idempotency guard -- see createPO() for the full rationale. One key
      // covers the whole batch (all-or-nothing).
      var idemKey = (payload.idempotencyKey || '').toString().trim();
      var cache = CacheService.getScriptCache();
      var cacheKey = idemKey ? ('idem_timecorrectionbatch_' + idemKey) : null;
      if (cacheKey) {
        var cached = null;
        try { cached = cache.get(cacheKey); } catch (e) {}
        if (cached) return JSON.parse(cached);
      }

      items.forEach(function(it) {
        var originalIn = '', originalOut = '';
        var match = origByKey[email + '|' + it.targetMidnight.getTime()];
        if (match) {
          originalIn  = match.in  ? Utilities.formatDate(new Date(match.in),  tz, 'h:mm a') : '';
          originalOut = match.out ? Utilities.formatDate(new Date(match.out), tz, 'h:mm a') : '';
        }
        sheet.appendRow([
          Utilities.getUuid(), new Date(), name, email, it.dateStr,
          originalIn, originalOut, it.clockIn, it.clockOut, reason,
          'pending', '', '', ''
        ]);
      });
      var result = { success: true, count: items.length };
      if (cacheKey) { try { cache.put(cacheKey, JSON.stringify(result), 300); } catch (e) {} }
      return result;
    } finally {
      lock.releaseLock();
    }
  } catch(e) { return { error: e.toString() }; }
}

/** The caller's own correction requests (any status), newest submitted first. */
function getMyTimeCorrections(payload) {
  try {
    var auth = requireVerifiedEmail_(payload);
    if (auth.error) return auth;
    var email = auth.email;
    var sheet = getTimeCorrectionsSheet_();
    var lastRow = sheet.getLastRow();
    var requests = [];
    if (lastRow >= 2) {
      var data = sheet.getRange(2, 1, lastRow - 1, TIME_CORRECTIONS_HEADERS.length).getValues();
      var tz = Session.getScriptTimeZone();
      for (var i = 0; i < data.length; i++) {
        if ((data[i][3] || '').toString().toLowerCase().trim() !== email) continue;
        requests.push(formatCorrectionRow_(data[i], tz));
      }
    }
    requests.reverse();
    return { requests: requests };
  } catch(e) { return { error: e.toString() }; }
}

/** All pending correction requests across employees, for the admin/HR review queue. */
function getTimeCorrectionQueue(payload) {
  try {
    var auth = authorizeCaller(payload, ['admin', 'human_resources']);
    if (!auth.ok) return { error: auth.error, code: auth.code };
    var sheet = getTimeCorrectionsSheet_();
    var lastRow = sheet.getLastRow();
    var queue = [];
    if (lastRow >= 2) {
      var data = sheet.getRange(2, 1, lastRow - 1, TIME_CORRECTIONS_HEADERS.length).getValues();
      var tz = Session.getScriptTimeZone();
      for (var i = 0; i < data.length; i++) {
        if ((data[i][10] || '').toString() !== 'pending') continue;
        queue.push(formatCorrectionRow_(data[i], tz));
      }
    }
    return { queue: queue };
  } catch(e) { return { error: e.toString() }; }
}

/**
 * Approves a correction request: writes the requested time(s) into the matching
 * Time Tracking row (only the field(s) actually requested), recomputes Hours if
 * both Clock In and Clock Out are now present, or appends a brand-new row if no
 * shift existed for that date at all (e.g. the employee forgot to punch in).
 */
function approveTimeCorrection(payload) {
  try {
    var auth = authorizeCaller(payload, ['admin', 'human_resources']);
    if (!auth.ok) return { error: auth.error, code: auth.code };
    var requestId = (payload.requestId || '').toString();
    if (!requestId) return { error: 'Missing requestId' };

    var sheet = getTimeCorrectionsSheet_();
    var lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) return { error: 'Server is busy - try again in a moment.' };
    try {
      var found = findCorrectionRow_(sheet, requestId);
      if (!found) return { error: 'Request not found' };
      if ((found.row[10] || '').toString() !== 'pending') return { error: 'This request has already been reviewed.' };

      var tz      = Session.getScriptTimeZone();
      var email   = (found.row[3] || '').toString().toLowerCase().trim();
      var name    = (found.row[2] || '').toString();
      var dateStr = formatCorrectionDate_(found.row[4], tz);
      var reqIn   = formatCorrectionTime24_(found.row[7], tz);
      var reqOut  = formatCorrectionTime24_(found.row[8], tz);

      var dp = dateStr.split('-');
      var targetMidnight = new Date(parseInt(dp[0], 10), parseInt(dp[1], 10) - 1, parseInt(dp[2], 10));
      var tsSheet = getTimeSheet_();
      var matchRow = findTimeTrackingRowForDate_(email, targetMidnight);

      var newIn  = reqIn  ? combineDateTime_(dateStr, reqIn)  : null;
      var newOut = reqOut ? combineDateTime_(dateStr, reqOut) : null;

      if (matchRow !== -1) {
        if (newIn)  tsSheet.getRange(matchRow, 4).setValue(newIn);
        if (newOut) tsSheet.getRange(matchRow, 5).setValue(newOut);
        var finalInVal  = tsSheet.getRange(matchRow, 4).getValue();
        var finalOutVal = tsSheet.getRange(matchRow, 5).getValue();
        if (finalInVal && finalOutVal) {
          var hrs = Math.round(Math.max(0, (new Date(finalOutVal) - new Date(finalInVal)) / 3600000) * 100) / 100;
          tsSheet.getRange(matchRow, 6).setValue(hrs);
        }
      } else {
        var hours = (newIn && newOut) ? Math.round(Math.max(0, (newOut - newIn) / 3600000) * 100) / 100 : '';
        tsSheet.appendRow([name, email, Utilities.formatDate(targetMidnight, tz, 'MM/dd/yyyy'), newIn || '', newOut || '', hours, '', '']);
      }

      sheet.getRange(found.rowNum, 11).setValue('approved');
      sheet.getRange(found.rowNum, 12).setValue(auth.email);
      sheet.getRange(found.rowNum, 13).setValue(new Date());
      return { success: true };
    } finally {
      lock.releaseLock();
    }
  } catch(e) { return { error: e.toString() }; }
}

/** Denies a correction request, leaving Time Tracking untouched. Optional note is shown to the employee. */
function denyTimeCorrection(payload) {
  try {
    var auth = authorizeCaller(payload, ['admin', 'human_resources']);
    if (!auth.ok) return { error: auth.error, code: auth.code };
    var requestId = (payload.requestId || '').toString();
    if (!requestId) return { error: 'Missing requestId' };
    var note = (payload.note || '').toString().trim();

    var sheet = getTimeCorrectionsSheet_();
    var lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) return { error: 'Server is busy - try again in a moment.' };
    try {
      var found = findCorrectionRow_(sheet, requestId);
      if (!found) return { error: 'Request not found' };
      if ((found.row[10] || '').toString() !== 'pending') return { error: 'This request has already been reviewed.' };
      sheet.getRange(found.rowNum, 11).setValue('denied');
      sheet.getRange(found.rowNum, 12).setValue(auth.email);
      sheet.getRange(found.rowNum, 13).setValue(new Date());
      sheet.getRange(found.rowNum, 14).setValue(note);
      return { success: true };
    } finally {
      lock.releaseLock();
    }
  } catch(e) { return { error: e.toString() }; }
}

// ── Admin: Manual Time Entry ───────────────────────────────────────────────────
// Lets an admin/HR reviewer add or correct an employee's clock in/out (or set
// a flat hours total) directly from the Payroll panel, without routing through
// the employee-submitted Time Corrections queue (submitTimeCorrection). Every
// write is ALSO logged as an already-'approved' row in that same Time
// Corrections sheet purely for an audit trail -- it shows up in the admin
// queue history and in the employee's own getMyTimeCorrections list, so there's
// a visible record of who changed what even though no request was ever pending.
function adminSetTimeEntry(payload) {
  try {
    var auth = authorizeCaller(payload, ['admin', 'human_resources']);
    if (!auth.ok) return { error: auth.error, code: auth.code };

    var email = (payload.employeeEmail || '').toString().toLowerCase().trim();
    if (!email) return { error: 'Choose an employee.' };
    var dateStr = (payload.date || '').toString().trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) return { error: 'Choose a valid date.' };

    var clockIn  = (payload.clockIn  || '').toString().trim();
    var clockOut = (payload.clockOut || '').toString().trim();
    var hoursRaw = (payload.hoursOverride || '').toString().trim();
    var hasHoursOverride = hoursRaw !== '';
    var hoursOverride = hasHoursOverride ? parseFloat(hoursRaw) : null;
    if (hasHoursOverride && (isNaN(hoursOverride) || hoursOverride < 0 || hoursOverride > 24)) {
      return { error: 'Hours must be a number between 0 and 24.' };
    }
    if (!clockIn && !clockOut && !hasHoursOverride) {
      return { error: 'Enter a clock in/out time or a total hours value.' };
    }

    var rolesMap = getRolesMap_();
    var name = (rolesMap[email] && rolesMap[email].name) || email;
    var tz = Session.getScriptTimeZone();
    var dp = dateStr.split('-');
    var targetMidnight = new Date(parseInt(dp[0], 10), parseInt(dp[1], 10) - 1, parseInt(dp[2], 10));

    var lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) return { error: 'Server is busy - try again in a moment.' };
    try {
      var tsSheet = getTimeSheet_();
      var matchRow = findTimeTrackingRowForDate_(email, targetMidnight);

      var originalIn = '', originalOut = '';
      if (matchRow !== -1) {
        var existing = tsSheet.getRange(matchRow, 4, 1, 2).getValues()[0];
        originalIn  = existing[0] ? Utilities.formatDate(new Date(existing[0]), tz, 'h:mm a') : '';
        originalOut = existing[1] ? Utilities.formatDate(new Date(existing[1]), tz, 'h:mm a') : '';
      }

      var newIn  = clockIn  ? combineDateTime_(dateStr, clockIn)  : null;
      var newOut = clockOut ? combineDateTime_(dateStr, clockOut) : null;

      if (matchRow !== -1) {
        if (newIn)  tsSheet.getRange(matchRow, 4).setValue(newIn);
        if (newOut) tsSheet.getRange(matchRow, 5).setValue(newOut);
        if (hasHoursOverride) {
          tsSheet.getRange(matchRow, 6).setValue(hoursOverride);
        } else {
          var finalInVal  = tsSheet.getRange(matchRow, 4).getValue();
          var finalOutVal = tsSheet.getRange(matchRow, 5).getValue();
          if (finalInVal && finalOutVal) {
            var hrs = Math.round(Math.max(0, (new Date(finalOutVal) - new Date(finalInVal)) / 3600000) * 100) / 100;
            tsSheet.getRange(matchRow, 6).setValue(hrs);
          }
        }
      } else {
        var hours = hasHoursOverride ? hoursOverride
          : ((newIn && newOut) ? Math.round(Math.max(0, (newOut - newIn) / 3600000) * 100) / 100 : '');
        tsSheet.appendRow([name, email, Utilities.formatDate(targetMidnight, tz, 'MM/dd/yyyy'), newIn || '', newOut || '', hours, '', '']);
      }

      // Audit trail -- see comment above the function.
      var corrSheet = getTimeCorrectionsSheet_();
      var note = (payload.note || '').toString().trim() || ('Manual entry by ' + auth.email);
      if (hasHoursOverride) note += ' [hours set to ' + hoursOverride + ']';
      corrSheet.appendRow([
        Utilities.getUuid(), new Date(), name, email, dateStr,
        originalIn, originalOut, clockIn, clockOut, note,
        'approved', auth.email, new Date(), ''
      ]);

      return { success: true };
    } finally {
      lock.releaseLock();
    }
  } catch(e) { return { error: e.toString() }; }
}

/**
 * Deletes an employee's Time Tracking row for one date entirely (e.g. a
 * duplicate punch, or an entry that should never have existed). Logs the same
 * kind of audit row as adminSetTimeEntry, with blank requested times marking
 * a removal rather than a change.
 */
function adminDeleteTimeEntry(payload) {
  try {
    var auth = authorizeCaller(payload, ['admin', 'human_resources']);
    if (!auth.ok) return { error: auth.error, code: auth.code };

    var email = (payload.employeeEmail || '').toString().toLowerCase().trim();
    if (!email) return { error: 'Missing employeeEmail.' };
    var dateStr = (payload.date || '').toString().trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) return { error: 'Invalid date.' };

    var dp = dateStr.split('-');
    var targetMidnight = new Date(parseInt(dp[0], 10), parseInt(dp[1], 10) - 1, parseInt(dp[2], 10));
    var tz = Session.getScriptTimeZone();

    var lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) return { error: 'Server is busy - try again in a moment.' };
    try {
      var tsSheet = getTimeSheet_();
      var matchRow = findTimeTrackingRowForDate_(email, targetMidnight);
      if (matchRow === -1) return { error: 'No time entry found for that date.' };

      var existing = tsSheet.getRange(matchRow, 1, 1, 6).getValues()[0];
      var name = (existing[0] || email).toString();
      var originalIn  = existing[3] ? Utilities.formatDate(new Date(existing[3]), tz, 'h:mm a') : '';
      var originalOut = existing[4] ? Utilities.formatDate(new Date(existing[4]), tz, 'h:mm a') : '';

      tsSheet.deleteRow(matchRow);

      var corrSheet = getTimeCorrectionsSheet_();
      var note = (payload.note || '').toString().trim() || ('Entry deleted by ' + auth.email);
      corrSheet.appendRow([
        Utilities.getUuid(), new Date(), name, email, dateStr,
        originalIn, originalOut, '', '', note,
        'approved', auth.email, new Date(), ''
      ]);

      return { success: true };
    } finally {
      lock.releaseLock();
    }
  } catch(e) { return { error: e.toString() }; }
}

// ── Admin: Payroll Summary ────────────────────────────────────────────────────
function getPayrollSummary(payload) {
  try {
    var auth = authorizeCaller(payload, ['admin', 'human_resources']);
    if (!auth.ok) return { error: auth.error, code: auth.code };
    var sh  = getTimeSheet_();
    var offset = parseInt(payload.periodOffset, 10);
    if (isNaN(offset) || offset > 0) offset = 0; // never allow future periods
    var bounds = getPeriodBoundsOffset_(new Date(), offset);
    var pStart = bounds.start;
    var pEnd   = bounds.end;
    var periodLabel = formatPeriodLabel_(pStart, pEnd);

    // Build email->name lookup from the cached HR roles map (authoritative
    // source) instead of a fresh full-sheet scan on every summary request.
    var hrRolesMap = getRolesMap_();
    var hrNameMap = {};
    for (var hrEmail in hrRolesMap) {
      hrNameMap[hrEmail] = hrRolesMap[hrEmail].name;
    }

    var flagged = findFlaggableShifts_(sh, pStart, pEnd);
    var flaggedRowSet = {};
    flagged.forEach(function(f) { flaggedRowSet[f.rowIndex] = true; });

    var empMap = {};
    var lastRow = sh.getLastRow();
    if (lastRow >= 2) {
      var tz = Session.getScriptTimeZone();
      var data = sh.getRange(2, 1, lastRow - 1, 8).getValues();
      for (var i = 0; i < data.length; i++) {
        if (flaggedRowSet[i + 2]) continue; // excluded from totals until resolved -- see needsReview
        var rowEmail = (data[i][1] || '').toString().toLowerCase().trim();
        var rowDate  = data[i][2] ? new Date(data[i][2]) : null;
        var rowHours = parseFloat(data[i][5]) || 0;
        if (!rowDate || !rowEmail) continue;
        var rd = new Date(rowDate.getFullYear(), rowDate.getMonth(), rowDate.getDate());
        if (rd < pStart || rd > pEnd) continue;
        if (!empMap[rowEmail]) {
          var resolvedName = hrNameMap[rowEmail] || (data[i][0] || '').toString().trim() || rowEmail;
          empMap[rowEmail] = { name: resolvedName, total: 0, days: {}, dayEntries: [], shifts: [] };
        }
        empMap[rowEmail].total = Math.round((empMap[rowEmail].total + rowHours) * 100) / 100;
        var dayLabel = MONTH_ABBRS[rd.getMonth()] + ' ' + rd.getDate();
        empMap[rowEmail].days[dayLabel] = Math.round(((empMap[rowEmail].days[dayLabel] || 0) + rowHours) * 100) / 100;
        var existingEntry = empMap[rowEmail].dayEntries.filter(function(e) { return e.date.getTime() === rd.getTime(); })[0];
        if (existingEntry) { existingEntry.hours += rowHours; }
        else { empMap[rowEmail].dayEntries.push({ date: rd, hours: rowHours }); }
        empMap[rowEmail].shifts.push({
          date: dayLabel,
          dateISO:    Utilities.formatDate(rd, tz, 'yyyy-MM-dd'),
          clockIn:    data[i][3] ? Utilities.formatDate(new Date(data[i][3]), tz, 'h:mm a') : '',
          clockOut:   data[i][4] ? Utilities.formatDate(new Date(data[i][4]), tz, 'h:mm a') : '',
          clockIn24:  data[i][3] ? Utilities.formatDate(new Date(data[i][3]), tz, 'HH:mm') : '',
          clockOut24: data[i][4] ? Utilities.formatDate(new Date(data[i][4]), tz, 'HH:mm') : '',
          hours:      rowHours,
          clockInLoc:  (data[i][6] || '').toString(),
          clockOutLoc: (data[i][7] || '').toString()
        });
      }
    }

    var approvalMap = getApprovalMap_(periodLabel);
    var employees = Object.keys(empMap).map(function(e) {
      var otSplit = splitRegularOvertime_(empMap[e].dayEntries);
      var approval = approvalMap[e];
      // NOTE: a map entry can now exist purely from the employee's own
      // self-approval (approveMyTimesheet), so "admin approved" must check
      // approvedAt specifically rather than just truthiness of the entry.
      return {
        email: e, name: empMap[e].name, total: empMap[e].total,
        regularHours: otSplit.regular, overtimeHours: otSplit.overtime,
        days: empMap[e].days, shifts: empMap[e].shifts,
        approved: !!(approval && approval.approvedAt), approvedBy: approval ? approval.approvedBy : '', approvedAt: approval ? approval.approvedAt : '',
        employeeApproved: !!(approval && approval.employeeApprovedAt), employeeApprovedAt: approval ? approval.employeeApprovedAt : '', employeeNote: approval ? approval.employeeNote : '',
        employeePdfUrl: approval ? approval.employeePdfUrl : ''
      };
    }).sort(function(a, b) { return a.name.localeCompare(b.name); });

    return { employees: employees, periodLabel: periodLabel, needsReview: flagged, periodOffset: offset };
  } catch(e) { return { error: e.toString() }; }
}

// Builds the same payroll summary as emailPayroll(), but renders it to a PDF
// and returns it as base64 for the client to download directly -- no email
// sent. Replaces the old "Email" button on the Payroll panel.
function getPayrollPdf(payload) {
  try {
    var auth = authorizeCaller(payload, ['admin', 'human_resources']);
    if (!auth.ok) return { error: auth.error, code: auth.code };
    var summary = getPayrollSummary(payload);
    if (summary.error) return { error: summary.error };

    var esc = function(s) { return (s == null ? '' : s.toString()).replace(/[&<>]/g, function(c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]; }); };
    var grandTotal = 0;
    var rows = summary.employees.map(function(e) {
      grandTotal += e.total;
      var dayLines = Object.keys(e.days).map(function(d) { return esc(d) + ': ' + e.days[d] + ' hrs'; }).join('<br>');
      var status = e.approved ? 'Approved by ' + esc(e.approvedBy) : 'PENDING APPROVAL';
      return '<tr>' +
        '<td>' + esc(e.name) + '</td>' +
        '<td>' + e.total + ' hrs (' + e.regularHours + ' reg / ' + e.overtimeHours + ' OT)</td>' +
        '<td>' + status + '</td>' +
        '<td>' + dayLines + '</td>' +
      '</tr>';
    }).join('');

    var needsReviewHtml = '';
    if (summary.needsReview && summary.needsReview.length) {
      var reviewRows = summary.needsReview.map(function(r) {
        return '<tr><td>' + esc(r.name) + ' (' + esc(r.email) + ')</td><td>' + esc(r.date) + '</td><td>' + esc(r.reason) + '</td></tr>';
      }).join('');
      needsReviewHtml = '<h2>Needs Review (excluded from totals above)</h2>' +
        '<table><tr><th>Employee</th><th>Date</th><th>Reason</th></tr>' + reviewRows + '</table>';
    }

    var html = '<!DOCTYPE html><html><head><meta charset="utf-8"><style>' +
      'body{font-family:Arial,Helvetica,sans-serif;font-size:12px;color:#222}' +
      'h1{font-size:18px;margin:0 0 4px}' +
      'h2{font-size:13px;margin:20px 0 6px}' +
      'table{width:100%;border-collapse:collapse;margin-top:8px}' +
      'th{text-align:left;background:#f4f4f4;padding:6px 8px;font-size:11px;border-bottom:2px solid #ddd}' +
      'td{padding:6px 8px;border-bottom:1px solid #eee;vertical-align:top}' +
      '.total{margin-top:14px;font-weight:bold}' +
      '</style></head><body>' +
      '<h1>Payroll Summary</h1><div>' + esc(summary.periodLabel) + '</div>' +
      '<table><tr><th>Employee</th><th>Hours</th><th>Status</th><th>Daily Breakdown</th></tr>' + rows + '</table>' +
      '<div class="total">Grand Total: ' + (Math.round(grandTotal * 100) / 100) + ' hrs</div>' +
      needsReviewHtml +
      '</body></html>';

    var pdfBlob = HtmlService.createHtmlOutput(html).getAs('application/pdf');
    var filename = 'Payroll Summary - ' + summary.periodLabel + '.pdf';
    return { success: true, filename: filename, base64: Utilities.base64Encode(pdfBlob.getBytes()) };
  } catch(e) { return { error: e.toString() }; }
}

function emailPayroll(payload) {
  try {
    var auth = authorizeCaller(payload, ['admin', 'human_resources']);
    if (!auth.ok) return { error: auth.error, code: auth.code };
    var to = payload.to || Session.getActiveUser().getEmail();
    var summary = getPayrollSummary(payload);
    if (summary.error) return { error: summary.error };
    var lines = ['Payroll Summary - ' + summary.periodLabel, '===========================', ''];
    var grandTotal = 0;
    summary.employees.forEach(function(e) {
      var approvalNote = e.approved ? ' [Approved by ' + e.approvedBy + ']' : ' [PENDING APPROVAL]';
      lines.push(e.name + ': ' + e.total + ' hrs (' + e.regularHours + ' reg / ' + e.overtimeHours + ' OT)' + approvalNote);
      var dayKeys = Object.keys(e.days);
      dayKeys.forEach(function(d) { lines.push('  ' + d + ': ' + e.days[d] + ' hrs'); });
      lines.push('');
      grandTotal += e.total;
    });
    lines.push('---------------------------');
    lines.push('Grand Total: ' + Math.round(grandTotal * 100) / 100 + ' hrs');
    if (summary.needsReview && summary.needsReview.length) {
      lines.push('');
      lines.push('NEEDS REVIEW (excluded from totals above):');
      summary.needsReview.forEach(function(r) {
        lines.push('  ' + r.name + ' (' + r.email + ') - ' + r.date + ' - ' + r.reason);
      });
    }
    GmailApp.sendEmail(to, 'Payroll Summary - ' + summary.periodLabel, lines.join('\n'));
    return { success: true };
  } catch(e) { return { error: e.toString() }; }
}