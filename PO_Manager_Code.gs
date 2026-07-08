/**
 * PO Manager Web App - Panoramic Building
 * =========================================
 * Paste this into your Google Apps Script project (Extensions > Apps Script).
 * Also paste the contents of PO_Manager_index.html into a new HTML file named "index".
 * Then deploy: Deploy > New Deployment > Web App.
 */

var SHEET_NAME  = "PO Database";
var ROLES_SHEET = "HR";

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
  "Other"
];

var VENDOR_OPTIONS = [
  "ABC Interiors",
  "Lansing",
  "Timberline",
  "Castalite",
  "Harristone",
  "Tresselwood",
  "Leak Tech",
  "Plaster",
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

    if      (action === 'getConfig')        result = getConfig(payload.email);
    else if (action === 'verifyLogin')       result = verifyLogin(payload.email, payload.password);
    else if (action === 'getSheetData')      result = getSheetData();
    else if (action === 'createPO')          result = createPO(payload);
    else if (action === 'updatePO')          result = updatePO(payload.rowIndex, payload.updates);
    else if (action === 'findPOByNumber')    result = findPOByNumber(payload.poNum);
    else if (action === 'savePhotoToDrive')  result = savePhotoToDrive(payload.base64Data, payload.mimeType, payload.filename);
    else if (action === 'getPricingData')    result = getPricingData();
    else if (action === 'updatePricing')     result = updatePricing(payload.rowIndex, payload.vendorPrices);
    else if (action === 'getContacts')         result = getContacts();
    else if (action === 'updateContact')       result = updateContact(payload.rowIndex, payload.values);
    else if (action === 'reconcileStatement')  result = reconcileStatement(payload.invoiceNumbers);
    else if (action === 'getJobList')          result = getJobList();
    else if (action === 'getJobCostSummary')   result = getJobCostSummary(payload.jobRef);
    else if (action === 'getMissingInvoices')  result = getMissingInvoices();
    else if (action === 'getVendorSpend')      result = getVendorSpend(payload.startDate, payload.endDate);
    else if (action === 'categorizeInvoices')  result = categorizeInvoices(payload);
    else if (action === 'suggestCategories')   result = suggestCategories(payload);
    else if (action === 'processEstimateWithMatching') result = processEstimateWithMatching(payload);
    else if (action === 'getSopData')                  result = getSopData();
    else if (action === 'saveMaterialHistory')          result = saveMaterialHistory(payload);
    else if (action === 'getAsanaJobs')                result = getAsanaJobs();
    else if (action === 'submitQualityCheck')           result = submitQualityCheck(payload);
    else if (action === 'getPTOData')                  result = getPTOData(payload.email, payload.role);
    else if (action === 'submitPTORequest')             result = submitPTORequest(payload);
    else if (action === 'getPTOQueue')                  result = getPTOQueue();
    else if (action === 'approvePTO')                   result = approvePTO(payload);
    else if (action === 'denyPTO')                      result = denyPTO(payload);
    else                                        result = { error: 'Unknown action: ' + action };

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
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
 */
function getSheetData() {
  var sheet = getSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var numRows = lastRow - 1;
  var data     = sheet.getRange(2, 1, numRows, 12).getValues();
  var tz       = Session.getScriptTimeZone();
  var pos      = [];

  // getRichTextValues lets us read hyperlinks that getValues() strips out.
  // Column A (index 1) holds the invoice hyperlink on the PO number cell.
  // Column J (index 10) holds the issued-PO link.
  var colARich = sheet.getRange(2, 1,  numRows, 1).getRichTextValues();
  var colJRich = sheet.getRange(2, 10, numRows, 1).getRichTextValues();

  data.forEach(function(row, i) {
    var poNum = row[0] ? row[0].toString().trim() : "";
    if (!isValidPONumber(poNum)) return; // skip header / input rows

    var dateIssued   = formatDateCell(row[1], tz);
    var deliveryDate = formatDateCell(row[8], tz);

    // Extract hyperlink URLs from rich-text cells
    var invoiceLink  = "";
    var issuedPOLink = "";
    try { invoiceLink  = colARich[i][0].getLinkUrl() || ""; } catch(e) {}
    try { issuedPOLink = colJRich[i][0].getLinkUrl() || ""; } catch(e) {}

    // Column J may also just contain a plain-text URL
    if (!issuedPOLink) issuedPOLink = str(row[9]);

    pos.push({
      rowIndex:     i + 2,
      poNum:        poNum,
      dateIssued:   dateIssued,
      builder:      str(row[2]),
      jobRef:       str(row[3]),
      vendor:       str(row[4]),
      vendorInvoice:str(row[5]),
      status:       str(row[6]).trim(),
      invoiceTotal: str(row[7]),
      deliveryDate: deliveryDate,
      issuedPO:     str(row[9]),
      issuedPOLink: issuedPOLink,
      invoiceLink:  invoiceLink,
      receivedNote: str(row[10]),
      notes:        str(row[11])
    });
  });

  return pos;
}

/**
 * Creates a new PO row and returns { success, poNumber } or { success: false, error }.
 */
function createPO(data) {
  try {
    if (!data.jobRef || !data.vendor) {
      return { success: false, error: "Job Reference and Vendor are required." };
    }

    var sheet = getSheet();
    var now   = new Date();
    var tz    = Session.getScriptTimeZone();
    var year  = Utilities.formatDate(now, tz, "yy");
    var qtr   = Math.ceil((now.getMonth() + 1) / 3);
    var paddedQtr = ("0" + qtr).slice(-2);

    var nextRow  = sheet.getLastRow() + 1;
    var poNumber = year + "-" + paddedQtr + "-" + Utilities.formatString("%03d", nextRow);
    var today    = Utilities.formatDate(now, tz, "MM/dd/yyyy");

    sheet.getRange(nextRow, 1).setValue(poNumber);
    sheet.getRange(nextRow, 2).setValue(today);
    sheet.getRange(nextRow, 3).setValue(data.builder       || "");
    sheet.getRange(nextRow, 4).setValue(data.jobRef        || "");
    sheet.getRange(nextRow, 5).setValue(data.vendor        || "");
    sheet.getRange(nextRow, 6).setValue(data.vendorInvoice || "");
    sheet.getRange(nextRow, 7).setValue(data.status        || "Pending Pickup");
    sheet.getRange(nextRow, 8).setValue(data.invoiceTotal  || "");
    sheet.getRange(nextRow, 12).setValue(data.notes        || "");

    return { success: true, poNumber: poNumber };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Updates specific fields on an existing PO row.
 * Only fields present in `updates` are written.
 */
function updatePO(rowIndex, updates) {
  try {
    var sheet = getSheet();

    if (updates.builder       !== undefined) sheet.getRange(rowIndex, 3).setValue(updates.builder);
    if (updates.jobRef        !== undefined) sheet.getRange(rowIndex, 4).setValue(updates.jobRef);
    if (updates.vendor        !== undefined) sheet.getRange(rowIndex, 5).setValue(updates.vendor);
    if (updates.vendorInvoice !== undefined) sheet.getRange(rowIndex, 6).setValue(updates.vendorInvoice);
    if (updates.status        !== undefined) sheet.getRange(rowIndex, 7).setValue(updates.status);
    if (updates.invoiceTotal  !== undefined) sheet.getRange(rowIndex, 8).setValue(updates.invoiceTotal);
    if (updates.deliveryDate  !== undefined) sheet.getRange(rowIndex, 9).setValue(updates.deliveryDate);
    if (updates.issuedPO      !== undefined) sheet.getRange(rowIndex, 10).setValue(updates.issuedPO);
    if (updates.receivedNote  !== undefined) sheet.getRange(rowIndex, 11).setValue(updates.receivedNote);
    if (updates.notes         !== undefined) sheet.getRange(rowIndex, 12).setValue(updates.notes);

    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Looks up a single PO by number. Returns the PO object or null.
 */
function findPOByNumber(poNum) {
  try {
    if (!poNum) return null;
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
      var row = sheet.getRange(rowIndex, 1, 1, 12).getValues()[0];
      var invoiceLink = '', issuedPOLink = '';
      try { invoiceLink  = sheet.getRange(rowIndex, 1,  1, 1).getRichTextValues()[0][0].getLinkUrl() || ''; } catch(e2) {}
      try { issuedPOLink = sheet.getRange(rowIndex, 10, 1, 1).getRichTextValues()[0][0].getLinkUrl() || ''; } catch(e2) {}
      if (!issuedPOLink) issuedPOLink = str(row[9]);
      return {
        rowIndex:      rowIndex,
        poNum:         (row[0] || '').toString().trim(),
        dateIssued:    formatDateCell(row[1], tz),
        builder:       str(row[2]),
        jobRef:        str(row[3]),
        vendor:        str(row[4]),
        vendorInvoice: str(row[5]),
        status:        str(row[6]).trim(),
        invoiceTotal:  str(row[7]),
        deliveryDate:  formatDateCell(row[8], tz),
        issuedPO:      str(row[9]),
        issuedPOLink:  issuedPOLink,
        invoiceLink:   invoiceLink,
        receivedNote:  str(row[10]),
        notes:         str(row[11])
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
      if (rowEmail === email) {
        if (rowPass && rowPass === password) {
          return {
            success: true, role: rowRole, email: email,
            config: { statusOptions: STATUS_OPTIONS, vendorOptions: VENDOR_OPTIONS, userRole: rowRole, userEmail: email }
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
 * Returns config (status/vendor lists) + role for a cached/returning user.
 * Only called after a successful verifyLogin() - email is trusted from localStorage.
 */
function getConfig(email) {
  var roleData = getRoleByEmail(email || '');
  return {
    statusOptions: STATUS_OPTIONS,
    vendorOptions: VENDOR_OPTIONS,
    userRole:      roleData.role,
    userEmail:     roleData.email
  };
}

/**
 * Looks up a role by a caller-supplied email address.
 * Returns { role, email }. Falls back to 'runner' if not found.
 */
function getRoleByEmail(email) {
  try {
    if (!email) return { role: 'runner', email: '' };
    email = email.toLowerCase().trim();

    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(ROLES_SHEET);
    if (!sheet) return { role: 'runner', email: email };

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var rowEmail = (data[i][1] || '').toString().toLowerCase().trim(); // Column B
      var rowRole  = (data[i][3] || '').toString().toLowerCase().trim(); // Column D
      if (rowEmail === email) return { role: rowRole, email: email };
    }
    return { role: 'runner', email: email };
  } catch(e) {
    return { role: 'runner', email: email };
  }
}

/**
 * Looks up the active user's email in the Roles sheet and returns their role.
 * Roles sheet columns: A = Email, B = Role
 * Valid roles: admin | office | site_manager | runner | accountant
 * Falls back to 'runner' (most restricted) if email not found.
 */
function getUserRole() {
  try {
    var email = Session.getActiveUser().getEmail();
    if (!email) return { role: 'runner', email: 'unknown' };
    email = email.toLowerCase().trim();

    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(ROLES_SHEET);

    // No Roles sheet yet? Grant admin to script owner, runner to everyone else.
    if (!sheet) {
      var owner = Session.getEffectiveUser().getEmail().toLowerCase().trim();
      return { role: (email === owner ? 'admin' : 'runner'), email: email };
    }

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {          // row 0 = header
      var rowEmail = (data[i][0] || '').toString().toLowerCase().trim();
      var rowRole  = (data[i][1] || '').toString().toLowerCase().trim();
      if (rowEmail === email) return { role: rowRole, email: email };
    }

    // Not in the Roles sheet - default to runner (most restricted)
    return { role: 'runner', email: email };
  } catch(e) {
    return { role: 'runner', email: '' };
  }
}

// ─── Pricing ─────────────────────────────────────────────────────────────────

var PRICING_SHEET = "Pricing";

/**
 * Reads the Pricing sheet and returns { vendors, items }.
 * Vendor columns are read dynamically from the header row (E onwards),
 * so adding a new vendor column to the sheet requires no code changes.
 *
 * Layout: A=Description, B=U/M, C=Best Price, D=empty, E+=Vendors
 * Category header rows: description in A, everything else blank - no U/M and no prices.
 */
function getPricingData() {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(PRICING_SHEET);
    if (!sheet) return { vendors: [], items: [] };

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 5) return { vendors: [], items: [] };

    // Read header row to discover vendor columns (E onwards = index 4+)
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var vendorCols = []; // [{ name, colIndex }]
    for (var c = 4; c < headers.length; c++) {
      var h = (headers[c] || '').toString().trim();
      if (h) vendorCols.push({ name: h, colIndex: c });
    }

    // Read all data rows
    var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
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

    var lastUpdated = DriveApp.getFileById(ss.getId()).getLastUpdated();
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
function updatePricing(rowIndex, vendorPrices) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(PRICING_SHEET);
    if (!sheet) return { success: false, error: 'Pricing sheet not found' };

    var allPrices = [];
    // Vendor columns E-J are 1-based cols 5-10
    PRICING_VENDORS.forEach(function(vendor, i) {
      var price = vendorPrices[vendor];
      var col   = i + 5; // col E = 5, F = 6 ... J = 10
      if (price !== '' && price !== null && price !== undefined) {
        var val = parseFloat(price);
        sheet.getRange(rowIndex, col).setValue(isNaN(val) ? '' : val);
        if (!isNaN(val) && val > 0) allPrices.push(val);
      } else {
        sheet.getRange(rowIndex, col).setValue('');
      }
    });

    // Best price = lowest vendor price, written to col C
    var bestPrice = allPrices.length > 0 ? Math.min.apply(null, allPrices) : '';
    sheet.getRange(rowIndex, 3).setValue(bestPrice);

    return { success: true, bestPrice: bestPrice };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

// ─── Private Helpers ─────────────────────────────────────────────────────────

function getSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error("Sheet '" + SHEET_NAME + "' not found.");
  return sheet;
}

function isValidPONumber(s) {
  return /^\d{2}-\d{2}-\d{3,4}$/.test(s);
}

function formatDateCell(cell, tz) {
  if (!cell) return "";
  if (cell instanceof Date && !isNaN(cell)) {
    return Utilities.formatDate(cell, tz, "MM/dd/yyyy");
  }
  return cell.toString();
}

function str(val) {
  return val !== null && val !== undefined ? val.toString() : "";
}

// ─── Photo Upload ─────────────────────────────────────────────────────────────

/**
 * Receives a base64-encoded image from the web app, saves it to the
 * "PO Received Photos" folder in Drive and returns the shareable URL.
 *
 * ⚠️  SETUP: Create a folder called "PO Received Photos" in your Google Drive,
 * then paste its ID below (the long string from the folder's URL).
 *
 * Called client-side via google.script.run.savePhotoToDrive(...)
 */
var PO_PHOTOS_FOLDER_ID = "1SYFetk5XolUv9oIpJjBPhGDj-0SBqRJI";

function savePhotoToDrive(base64Data, mimeType, filename) {
  try {
    var folder = DriveApp.getFolderById(PO_PHOTOS_FOLDER_ID);

    var bytes = Utilities.base64Decode(base64Data);
    var blob  = Utilities.newBlob(bytes, mimeType, filename);
    var file  = folder.createFile(blob);

    // Anyone with the link can view (needed so the link is useful in the sheet)
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return { success: true, url: file.getUrl() };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
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
function getContacts() {
  try {
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
function updateContact(rowIndex, values) {
  try {
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

// ── Reconcile Statement ───────────────────────────────────────────────────────
function reconcileStatement(invoiceNumbers) {
  try {
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
function getJobCostSummary(jobRef) {
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

// ── Missing Invoices ──────────────────────────────────────────────────────────
function getMissingInvoices() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) return { error: 'PO Database sheet not found' };
    var data = sheet.getDataRange().getValues();
    var missing = [];
    // Statuses where we don't yet expect an invoice
    var skipStatuses = { 'draft': true, 'ordered': true, 'being made': true,
                         'pending pickup': true, 'pending delivery': true,
                         'pending delivery to supplier': true, 'currently picking up': true };
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

// ── Vendor Spend ──────────────────────────────────────────────────────────────
function getVendorSpend(startDate, endDate) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) return { error: 'PO Database sheet not found' };
    var data = sheet.getDataRange().getValues();
    var tz    = Session.getScriptTimeZone();
    var start = startDate ? new Date(startDate + 'T00:00:00') : null;
    var end   = endDate   ? new Date(endDate   + 'T23:59:59') : null;
    var vendors = {}, grandTotal = 0, vendorRows = {};
    for (var i = 0; i < data.length; i++) {
      if (!isValidPONumber((data[i][0] || '').toString().trim())) continue;
      var vendor = (data[i][4] || '').toString().trim();
      var total  = parseFloat(data[i][7]) || 0;
      if (!vendor || total === 0) continue;
      if (start || end) {
        var d = data[i][1] instanceof Date ? data[i][1] : null;
        if (!d || isNaN(d.getTime())) continue;
        if (start && d < start) continue;
        if (end   && d > end)   continue;
      }
      vendors[vendor] = (vendors[vendor] || 0) + total;
      grandTotal += total;
      // Track top rows per vendor for debugging
      if (!vendorRows[vendor]) vendorRows[vendor] = [];
      vendorRows[vendor].push({ poNum: data[i][0], total: total, row: i + 1 });
    }
    var result = Object.keys(vendors).map(function(v) {
      var rows = (vendorRows[v] || []).sort(function(a,b){return b.total-a.total;}).slice(0,3);
      return { vendor: v, total: vendors[v], topRows: rows };
    }).sort(function(a, b) { return b.total - a.total; });
    return { success: true, vendors: result, grandTotal: grandTotal, gasVersion: 3 };
  } catch(e) { return { error: e.toString() }; }
}


// ─── Material Report ─────────────────────────────────────────────────────────
function categorizeInvoices(payload) {
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

function getAsanaJobs() {
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
    return { jobs: jobs };
  } catch(e) { return { error: e.toString() }; }
}

function submitQualityCheck(payload) {
  try {
    var jobGid    = payload.jobGid;
    var jobName   = payload.jobName;
    var sections  = payload.sections;
    var submitter = payload.submitter || 'Field';
    var tz        = Session.getScriptTimeZone();
    var date      = Utilities.formatDate(new Date(), tz, 'MM/dd/yyyy');

    var lines = ['Quality Check - ' + date, 'Submitted by: ' + submitter, ''];
    var flagged = [];
    sections.forEach(function(s) {
      var icon = s.status === 'pass' ? 'PASS' : 'FLAG';
      lines.push('[' + icon + '] ' + s.name + (s.notes ? ' - ' + s.notes : ''));
      if (s.status !== 'pass') flagged.push(s);
    });

    var sub = asanaRequest('post', '/tasks/' + jobGid + '/subtasks', {
      name:      'Quality Check - ' + date,
      notes:     lines.join('\n'),
      completed: true
    });
    if (sub.errors) return { error: sub.errors[0].message };

    if (flagged.length > 0) {
      var offLines = ['Quality check flagged items for: ' + jobName + ' (' + date + ')', ''];
      flagged.forEach(function(f) {
        offLines.push('- ' + f.name + (f.notes ? ': ' + f.notes : ''));
      });
      asanaRequest('post', '/tasks', {
        projects: [ASANA_OFFICE_TASKS],
        name:     'Quality Check - ' + jobName + ' - ' + date,
        notes:    offLines.join('\n')
      });
    }

    return { success: true, flagged: flagged.length };
  } catch(e) { return { error: e.toString() }; }
}

// -- PTO / HR Functions ───────────────────────────────────────────────────────
// HR sheet columns: A=Name, B=Email, C=Phone, D=Role, E=Password, F=Allotted, G=Used

/**
 * Gets PTO balance + request history for an employee (and pending queue for admins).
 */
function getPTOData(email, role) {
  try {
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
      if (status === 'pending' && role === 'admin') {
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
    var email  = payload.email;
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
function getPTOQueue() {
  try {
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
 * Looks up a section GID by name in the PTO project.
 */
function getPTOSectionGid(sectionName) {
  try {
    var result = asanaRequest('get', '/projects/' + ASANA_PTO_PROJECT + '/sections?opt_fields=gid,name');
    if (result.errors || !result.data) return null;
    for (var i = 0; i < result.data.length; i++) {
      if (result.data[i].name === sectionName) return result.data[i].gid;
    }
    return null;
  } catch(e) { return null; }
}
