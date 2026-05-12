/**
 * PO Manager Web App — Panoramic Building
 * =========================================
 * Paste this into your Google Apps Script project (Extensions > Apps Script).
 * Also paste the contents of PO_Manager_index.html into a new HTML file named "index".
 * Then deploy: Deploy > New Deployment > Web App.
 */

var SHEET_NAME  = "PO Database";
var ROLES_SHEET = "Roles";

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
  "LW",
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
  var pos = getSheetData();
  for (var i = 0; i < pos.length; i++) {
    if (pos[i].poNum === poNum) return pos[i];
  }
  return null;
}

/**
 * Returns config (status/vendor lists) + current user's role.
 */
function getConfig() {
  var roleData = getUserRole();
  return {
    statusOptions: STATUS_OPTIONS,
    vendorOptions: VENDOR_OPTIONS,
    userRole:      roleData.role,
    userEmail:     roleData.email
  };
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

    // Not in the Roles sheet — default to runner (most restricted)
    return { role: 'runner', email: email };
  } catch(e) {
    return { role: 'runner', email: '' };
  }
}

// ─── Pricing ─────────────────────────────────────────────────────────────────

var PRICING_SHEET   = "Pricing";
var PRICING_VENDORS = ["LW Supply", "Allside", "Timberline", "LKL", "Lansing", "BFS"];

/**
 * Reads the Pricing sheet and returns an array of material objects.
 * Category header rows (no U/M) are captured and attached to subsequent items.
 */
function getPricingData() {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(PRICING_SHEET);
    if (!sheet) return [];

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    // Columns: A=desc, B=um, C=bestPrice, D=empty, E-J=vendors
    var data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
    var items = [];
    var currentCategory = '';

    data.forEach(function(row, rowIdx) {
      var desc = (row[0] || '').toString().trim();
      var um   = (row[1] || '').toString().trim();
      if (!desc) return;

      // Category header — no U/M value
      if (!um) { currentCategory = desc; return; }

      var bestPrice = parseFloat(row[2]) || 0;
      var prices    = {};
      // Vendor columns: E=4, F=5, G=6, H=7, I=8, J=9 (0-based)
      PRICING_VENDORS.forEach(function(vendor, i) {
        var v = row[i + 4];
        if (v !== '' && v !== null && v !== undefined && v !== 0) {
          prices[vendor] = parseFloat(v) || 0;
        }
      });

      items.push({
        description:  desc,
        um:           um,
        bestPrice:    bestPrice,
        prices:       prices,
        category:     currentCategory,
        rowIndex:     rowIdx + 2   // actual sheet row (1-based, +1 for header, +1 for 0-based)
      });
    });

    return items;
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
    return { success: true };
  } catch(e) { return { success: false, error: e.toString() }; }
}
