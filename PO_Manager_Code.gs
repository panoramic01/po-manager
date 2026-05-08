/**
 * PO Manager Web App — Panoramic Building
 * =========================================
 * Paste this into your Google Apps Script project (Extensions > Apps Script).
 * Also paste the contents of PO_Manager_index.html into a new HTML file named "index".
 * Then deploy: Deploy > New Deployment > Web App.
 */

var SHEET_NAME = "PO Database";

var STATUS_OPTIONS = [
  "Pending Pickup",
  "Pending Delivery",
  "Pending Delivery to Supplier",
  "Ordered",
  "Being made",
  "Currently Picking Up",
  "Delivered",
  "Ready to Reconcile",
  "Complete",
  "Draft"
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
    .setTitle("PO Manager — Panoramic Building")
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
 * Returns config (status/vendor lists) for populating dropdowns.
 */
function getConfig() {
  return {
    statusOptions: STATUS_OPTIONS,
    vendorOptions: VENDOR_OPTIONS
  };
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
