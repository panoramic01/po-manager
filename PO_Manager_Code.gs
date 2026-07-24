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
    else if (action === 'verifyGoogleLogin') result = verifyGoogleLogin(payload.credential);
    else if (action === 'getSheetData')      result = getSheetData();
    else if (action === 'createPO')          result = createPO(payload);
    else if (action === 'updatePO')          result = updatePO(payload.rowIndex, payload.updates);
    else if (action === 'findPOByNumber')    result = findPOByNumber(payload.poNum);
    else if (action === 'savePhotoToDrive')  result = savePhotoToDrive(payload.base64Data, payload.mimeType, payload.filename, payload.builder, payload.jobRef, payload.docType, payload.poNum);
    else if (action === 'createProject')       result = createProjectAndTask(payload);
    else if (action === 'saveFileToFolderById') result = saveFileToFolderById(payload.base64Data, payload.mimeType, payload.filename, payload.folderId);
    else if (action === 'getPricingData')    result = getPricingData();
    else if (action === 'updatePricing')     result = updatePricing(payload.rowIndex, payload.vendorPrices);
    else if (action === 'getContacts')         result = getContacts();
    else if (action === 'updateContact')       result = updateContact(payload.rowIndex, payload.values);
    else if (action === 'reconcileStatement')  result = reconcileStatement(payload.invoiceNumbers);
    else if (action === 'getJobList')          result = getJobList();
    else if (action === 'getJobCostSummary')   result = getJobCostSummary(payload.jobRef);
    else if (action === 'getMissingInvoices')  result = getMissingInvoices();
    else if (action === 'getJobsRegistry')     result = getJobsRegistry();
    else if (action === 'getJobDashboard')     result = getJobDashboard(payload.jobName);
    else if (action === 'updateJobMeta')       result = updateJobMeta(payload);
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
  var data     = sheet.getRange(2, 1, numRows, 14).getValues();
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

    var invoiceFile = str(row[13]);

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
      invoiceFile:  invoiceFile,
      invoiceLink:  invoiceFile || legacyInvoiceLink,
      receivedNote: str(row[10]),
      notes:        str(row[11]),
      orderedBy:    str(row[12])
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
    var status   = data.status || "Pending Pickup";

    sheet.getRange(nextRow, 1).setValue(poNumber);
    sheet.getRange(nextRow, 2).setValue(today);
    sheet.getRange(nextRow, 3).setValue(data.builder       || "");
    sheet.getRange(nextRow, 4).setValue(data.jobRef        || "");
    sheet.getRange(nextRow, 5).setValue(data.vendor        || "");
    sheet.getRange(nextRow, 6).setValue(data.vendorInvoice || "");
    sheet.getRange(nextRow, 7).setValue(status);
    sheet.getRange(nextRow, 8).setValue(data.invoiceTotal  || "");
    sheet.getRange(nextRow, 12).setValue(data.notes        || "");
    sheet.getRange(nextRow, 13).setValue(getFirstName(data.orderedBy));

    // Pending Pickup POs are picked up the same day they're created, so
    // default the pickup/delivery date to today rather than leaving it blank.
    if (status === "Pending Pickup") {
      sheet.getRange(nextRow, 9).setValue(today);
    }

    return { success: true, poNumber: poNumber, rowIndex: nextRow };
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
    if (updates.orderedBy     !== undefined) sheet.getRange(rowIndex, 13).setValue(updates.orderedBy);
    if (updates.invoiceFile   !== undefined) sheet.getRange(rowIndex, 14).setValue(updates.invoiceFile);

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
      var row = sheet.getRange(rowIndex, 1, 1, 14).getValues()[0];
      var legacyInvoiceLink = '', issuedPOLink = '';
      try { legacyInvoiceLink = sheet.getRange(rowIndex, 1,  1, 1).getRichTextValues()[0][0].getLinkUrl() || ''; } catch(e2) {}
      try { issuedPOLink      = sheet.getRange(rowIndex, 10, 1, 1).getRichTextValues()[0][0].getLinkUrl() || ''; } catch(e2) {}
      if (!issuedPOLink) issuedPOLink = str(row[9]);
      var invoiceFile = str(row[13]);
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
        invoiceFile:   invoiceFile,
        invoiceLink:   invoiceFile || legacyInvoiceLink,
        receivedNote:  str(row[10]),
        notes:         str(row[11]),
        orderedBy:     str(row[12])
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
          if (isOwnerEmail(email)) rowRole = 'admin';
          return {
            success: true, role: rowRole, email: email,
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
        if (isOwnerEmail(email)) rowRole = 'admin';
        return {
          success: true, role: rowRole, email: email,
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
 * Only called after a successful verifyLogin() - email is trusted from localStorage.
 */
function getConfig(email) {
  var roleData = getRoleByEmail(email || '');
  return {
    statusOptions:  STATUS_OPTIONS,
    vendorOptions:  VENDOR_OPTIONS,
    builderOptions: getBuilderNames(),
    jobOptions:     getRecentJobs(),
    userRole:       roleData.role,
    userEmail:      roleData.email,
    userName:       roleData.name,
    userPhone:      roleData.phone
  };
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
 * De-duplicated case-insensitively; a pair already covered by a Projects
 * row is not repeated from PO Database.
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
          var pKey = b.toLowerCase() + '|' + j.toLowerCase();
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
          var dKey = builder.toLowerCase() + '|' + jobRef.toLowerCase();
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
 * Builds a de-duplicated (case-insensitive), alphabetically sorted list of
 * builder/company names already in use, pulled from the "Projects" sheet
 * (Contractor, col A) and the "PO Database" sheet (Builder, col C). Powers
 * the New Project form's company dropdown so names stay consistent instead
 * of drifting across free-text entries. Always ends with "Other" so a
 * genuinely new company can still be typed in.
 */
function getBuilderNames() {
  try {
    var seen  = {}; // lowercase-trimmed name -> canonical display value
    var names = [];

    function addName(raw) {
      var s = (raw || '').toString().trim();
      if (!s) return;
      var key = s.toLowerCase();
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
    var email = (payload.email || '').toLowerCase().trim();
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
        return { success: true };
      }
    }
    return { error: 'Employee not found' };
  } catch(e) {
    return { error: e.message };
  }
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
    if (!sheet) return { role: isOwnerEmail(email) ? 'admin' : 'runner', email: email };

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var rowEmail = (data[i][1] || '').toString().toLowerCase().trim(); // Column B
      var rowRole  = (data[i][3] || '').toString().toLowerCase().trim(); // Column D
      var rowName  = (data[i][0] || '').toString().trim();               // Column A
      var rowPhone = (data[i][2] || '').toString().trim();               // Column C
      if (rowEmail === email) {
        if (isOwnerEmail(email)) rowRole = 'admin';
        return { role: rowRole, email: email, name: rowName, phone: rowPhone };
      }
    }
    return { role: isOwnerEmail(email) ? 'admin' : 'runner', email: email, name: '', phone: '' };
  } catch(e) {
    return { role: isOwnerEmail(email) ? 'admin' : 'runner', email: email, name: '', phone: '' };
  }
}

/**
 * Server-side authorization gate for privileged actions. Requires payload.callerEmail
 * to resolve (via getRoleByEmail, which applies the owner override above) to one of
 * allowedRoles. Callers must check .ok before proceeding.
 */
function authorizeCaller(payload, allowedRoles) {
  var callerEmail = ((payload && payload.callerEmail) || '').toString().toLowerCase().trim();
  if (!callerEmail) return { ok: false, code: 'AUTH_REQUIRED', error: 'You must be signed in to do this.' };
  var role = getRoleByEmail(callerEmail).role;
  if (allowedRoles.indexOf(role) === -1) {
    return { ok: false, code: 'FORBIDDEN', error: 'You do not have permission to do this.' };
  }
  return { ok: true, role: role, email: callerEmail };
}

/** Counts rows whose role (column D, index 3) is 'admin'. */
function countAdminRows(data) {
  var n = 0;
  for (var i = 0; i < data.length; i++) {
    if ((data[i][3] || '').toString().toLowerCase().trim() === 'admin') n++;
  }
  return n;
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
 * exists, else the global "Purchasing" folder. isProjectFolder tells the
 * caller whether to skip the ANYONE_WITH_LINK sharing fixup (project
 * folders live on a Shared Drive, governed by drive membership instead).
 */
function resolveBaseFolder(builder, jobRef) {
  var projectFolderId = getProjectFolderId(builder, jobRef);
  if (projectFolderId) {
    try {
      return { folder: DriveApp.getFolderById(projectFolderId), isProjectFolder: true };
    } catch (folderErr) {
      // bad/inaccessible ID in the Projects sheet -- fall back below
    }
  }
  return { folder: getPurchasingRootFolder(), isProjectFolder: false };
}

/**
 * Given a base folder (from resolveBaseFolder) and a document type,
 * returns/creates the folder the file should actually be written to:
 *   'issuedPO'      -> <base>/Issued POs
 *   'invoice'       -> <base>/Invoices
 *   'receivedPhoto' -> <base>/Received Photos/<poNum>
 * Unrecognized/missing docType falls back to the base folder itself.
 */
function getTypedUploadFolder(baseFolder, docType, poNum) {
  if (docType === 'issuedPO') return getOrCreateChildFolder(baseFolder, 'Issued POs');
  if (docType === 'invoice')  return getOrCreateChildFolder(baseFolder, 'Invoices');
  if (docType === 'receivedPhoto') {
    var photosFolder = getOrCreateChildFolder(baseFolder, 'Received Photos');
    return poNum ? getOrCreateChildFolder(photosFolder, poNum) : photosFolder;
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
var PROJECTS_SHEET_NAME = "Projects";

/**
 * Looks up the Drive folder ID for a given Contractor + Job Name pair in
 * the "Projects" sheet. Returns null (not throw) on any failure to look
 * up or on no match, so callers can fall back to the default folder.
 */
function getProjectFolderId(builder, jobRef) {
  try {
    var wantBuilder = (builder || "").toString().trim().toLowerCase();
    var wantJob     = (jobRef  || "").toString().trim().toLowerCase();
    if (!wantBuilder || !wantJob) return null;

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PROJECTS_SHEET_NAME);
    if (!sheet) return null;

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return null;

    var data = sheet.getRange(2, 1, lastRow - 1, 3).getValues(); // A:Contractor, B:Job Name, C:Drive ID
    for (var i = 0; i < data.length; i++) {
      var rowBuilder = (data[i][0] || "").toString().trim().toLowerCase();
      var rowJob     = (data[i][1] || "").toString().trim().toLowerCase();
      if (rowBuilder === wantBuilder && rowJob === wantJob) {
        return extractDriveFolderId(data[i][2]);
      }
    }
    return null;
  } catch (e) {
    return null;
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

function savePhotoToDrive(base64Data, mimeType, filename, builder, jobRef, docType, poNum) {
  try {
    var base   = resolveBaseFolder(builder, jobRef);
    var folder = getTypedUploadFolder(base.folder, docType, poNum);

    var bytes = Utilities.base64Decode(base64Data);
    var blob  = Utilities.newBlob(bytes, mimeType, filename);
    var file  = folder.createFile(blob);

    if (!base.isProjectFolder) {
      // New files inherit ANYONE_WITH_LINK from the folder (set once via
      // oneTimeSetFolderSharing) so no extra Drive permissions API call is
      // needed in the common case. Only fix up sharing if inheritance did
      // not actually apply, instead of paying for setSharing() on every
      // upload. Project folders live on a shared drive, where access is
      // governed by drive membership -- no setSharing() call needed there.
      try {
        if (file.getSharingAccess() !== DriveApp.Access.ANYONE_WITH_LINK) {
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        }
      } catch (sharingErr) {
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      }
    }

    return { success: true, url: file.getUrl() };
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
function saveFileToFolderById(base64Data, mimeType, filename, folderIdOrLink) {
  try {
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
// Extends the "Projects" sheet with two appended columns: E=Status,
// F=Start Date. Existing A-D columns (Contractor, Job Name, Drive folder,
// Asana GID) are untouched -- getProjectFolderId/getRecentJobs only ever
// read fixed-width ranges, so appending E/F is safe.
var QUALITY_CHECKS_SHEET_NAME = "Quality Checks";

function getOrCreateQualityChecksSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(QUALITY_CHECKS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(QUALITY_CHECKS_SHEET_NAME);
    sheet.appendRow(['Timestamp', 'Job Name', 'Trade(s)', 'Submitted By', 'Pass Count', 'Flag Count', 'Job GID']);
  }
  return sheet;
}

function ensureProjectsHeaders_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PROJECTS_SHEET_NAME);
  if (!sheet) return;
  if (!sheet.getRange(1, 5).getValue()) sheet.getRange(1, 5).setValue('Status');
  if (!sheet.getRange(1, 6).getValue()) sheet.getRange(1, 6).setValue('Start Date');
}

/**
 * Canonical job picker for the Job Dashboard panel, sourced from the
 * "Projects" sheet -- distinct from getJobList() (free-text jobRef scrape
 * of PO Database), which is left untouched for whatever else uses it.
 * Returns rowIndex so updateJobMeta can write back without re-matching.
 */
function getJobsRegistry() {
  try {
    ensureProjectsHeaders_();
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PROJECTS_SHEET_NAME);
    if (!sheet) return { error: "Sheet '" + PROJECTS_SHEET_NAME + "' not found." };
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, jobs: [] };
    var data = sheet.getRange(2, 1, lastRow - 1, 6).getValues(); // A:F
    var jobs = [];
    for (var i = 0; i < data.length; i++) {
      var jobName = (data[i][1] || '').toString().trim();
      if (!jobName) continue;
      jobs.push({
        rowIndex:   i + 2,
        contractor: (data[i][0] || '').toString().trim(),
        jobName:    jobName,
        status:     (data[i][4] || '').toString().trim(),
        startDate:  data[i][5] instanceof Date ? Utilities.formatDate(data[i][5], Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
        hasDrive:   !!extractDriveFolderId(data[i][2]),
        hasAsana:   !!(data[i][3] || '').toString().trim()
      });
    }
    return { success: true, jobs: jobs };
  } catch (e) { return { error: e.toString() }; }
}

/**
 * Combined payload for the Job Dashboard panel: job meta (Projects row),
 * cost summary (reuses getJobCostSummary as-is, spend only), missing-invoice
 * count for the job, and quality-walk history (reuses the "Quality Checks"
 * log appended to by submitQualityCheck). One round trip, matching this
 * app's existing one-action-per-panel convention.
 *
 * Quality Check's job picker uses Asana task names formatted as
 * "Builder, Job Name, Address" -- not the bare Projects-sheet Job Name --
 * so exact string matching would almost never hit. Quality-walk rows are
 * matched by substring (either direction), same fuzzy style already used
 * in reconcileStatement.
 */
function getJobDashboard(jobName) {
  try {
    var target = (jobName || '').toString().trim();
    if (!target) return { error: 'Job name is required.' };
    var targetLower = target.toLowerCase();

    var meta = null;
    var pSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PROJECTS_SHEET_NAME);
    if (pSheet) {
      var pLastRow = pSheet.getLastRow();
      if (pLastRow >= 2) {
        var pData = pSheet.getRange(2, 1, pLastRow - 1, 6).getValues();
        for (var i = 0; i < pData.length; i++) {
          var rowJob = (pData[i][1] || '').toString().trim();
          if (rowJob.toLowerCase() !== targetLower) continue;
          meta = {
            rowIndex:      i + 2,
            contractor:    (pData[i][0] || '').toString().trim(),
            jobName:       rowJob,
            driveFolderId: extractDriveFolderId(pData[i][2]),
            asanaTaskGid:  (pData[i][3] || '').toString().trim(),
            status:        (pData[i][4] || '').toString().trim(),
            startDate:     pData[i][5] instanceof Date ? Utilities.formatDate(pData[i][5], Session.getScriptTimeZone(), 'yyyy-MM-dd') : ''
          };
          break;
        }
      }
    }
    if (!meta) {
      meta = { rowIndex: null, contractor: '', jobName: target, driveFolderId: null, asanaTaskGid: '', status: '', startDate: '' };
    }

    var cost = getJobCostSummary(target);

    var missingAll = getMissingInvoices();
    var missingRows = [];
    if (missingAll.missing) {
      missingRows = missingAll.missing.filter(function(m) {
        return (m.job || '').toString().trim().toLowerCase() === targetLower;
      });
    }

    var quality = { count: 0, recent: [] };
    var qcSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(QUALITY_CHECKS_SHEET_NAME);
    if (qcSheet) {
      var qLastRow = qcSheet.getLastRow();
      if (qLastRow >= 2) {
        var qData = qcSheet.getRange(2, 1, qLastRow - 1, 7).getValues();
        var matches = [];
        for (var q = 0; q < qData.length; q++) {
          var qJob = (qData[q][1] || '').toString().trim().toLowerCase();
          if (!qJob) continue;
          if (qJob.indexOf(targetLower) === -1 && targetLower.indexOf(qJob) === -1) continue;
          matches.push({
            timestamp: qData[q][0] instanceof Date ? Utilities.formatDate(qData[q][0], Session.getScriptTimeZone(), 'MM/dd/yy') : '',
            ts:        qData[q][0] instanceof Date ? qData[q][0].getTime() : 0,
            trades:    qData[q][2],
            submitter: qData[q][3],
            passCount: qData[q][4],
            flagCount: qData[q][5]
          });
        }
        matches.sort(function(a, b) { return b.ts - a.ts; });
        quality.count = matches.length;
        quality.recent = matches.slice(0, 8);
      }
    }

    return { success: true, meta: meta, cost: cost, missingCount: missingRows.length, missingRows: missingRows, quality: quality };
  } catch (e) { return { error: e.toString() }; }
}

/**
 * Admin edit of a job's Status/Start Date on the Projects sheet, keyed by
 * rowIndex (not name) so two contractors sharing a job name can't collide.
 */
function updateJobMeta(payload) {
  try {
    var rowIndex = payload.rowIndex;
    if (!rowIndex) return { error: 'rowIndex is required.' };
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PROJECTS_SHEET_NAME);
    if (!sheet) return { error: "Sheet '" + PROJECTS_SHEET_NAME + "' not found." };
    sheet.getRange(rowIndex, 5).setValue((payload.status || '').toString().trim());
    var startDate = (payload.startDate || '').toString().trim();
    if (startDate) {
      var parts = startDate.split('-'); // input type=date gives YYYY-MM-DD
      sheet.getRange(rowIndex, 6).setValue(new Date(parts[0], parts[1] - 1, parts[2]));
    } else {
      sheet.getRange(rowIndex, 6).setValue('');
    }
    return { success: true };
  } catch (e) { return { error: e.toString() }; }
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

    // Additive: log this check to the "Quality Checks" sheet so job
    // dashboards can show walk history. Never let a logging failure break
    // the Asana submission above -- that flow must behave exactly as before.
    try {
      var passCount = 0, flagCountForLog = 0;
      sections.forEach(function(s) { if (s.status === 'pass') passCount++; else flagCountForLog++; });
      var qcLogSheet = getOrCreateQualityChecksSheet_();
      var tradesStr = (payload.trades || []).join(', ') || 'General';
      qcLogSheet.appendRow([new Date(), jobName, tradesStr, submitter, passCount, flagCountForLog, jobGid]);
    } catch (logErr) {
      // swallow -- logging must never fail the Quality Check submission
    }

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

    return {
      success: true,
      driveFolderId: folderId,
      asanaTaskGid: asanaTaskGid,
      asanaTaskUrl: 'https://app.asana.com/0/' + ASANA_EXT_SCHED + '/' + asanaTaskGid
    };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
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

function getTimeSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(TIME_SHEET);
  if (!sh) {
    sh = ss.insertSheet(TIME_SHEET);
    sh.getRange(1, 1, 1, 6).setValues([['Employee Name','Email','Date','Clock In','Clock Out','Hours']]);
    sh.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#1F3971').setFontColor('#ffffff');
  }
  return sh;
}

function clockIn(payload) {
  try {
    var email = payload.email;
    var name  = payload.name || email;
    var sh    = getTimeSheet_();
    var tz    = Session.getScriptTimeZone();
    var now   = new Date();

    // Check for open record
    var lastRow = sh.getLastRow();
    if (lastRow >= 2) {
      var data = sh.getRange(2, 1, lastRow - 1, 6).getValues();
      for (var i = data.length - 1; i >= 0; i--) {
        if ((data[i][1] || '').toString().toLowerCase() === email.toLowerCase() && !data[i][4]) {
          return { error: 'Already clocked in at ' + Utilities.formatDate(new Date(data[i][3]), tz, 'h:mm a') };
        }
      }
    }

    var today = Utilities.formatDate(now, tz, 'MM/dd/yyyy');
    sh.appendRow([name, email, today, now, '', '']);
    return { success: true, clockIn: Utilities.formatDate(now, tz, 'h:mm a') };
  } catch(e) { return { error: e.toString() }; }
}

function clockOut(payload) {
  try {
    var email = payload.email;
    var sh    = getTimeSheet_();
    var tz    = Session.getScriptTimeZone();
    var now   = new Date();

    var lastRow = sh.getLastRow();
    if (lastRow < 2) return { error: 'No clock-in record found' };

    var data = sh.getRange(2, 1, lastRow - 1, 6).getValues();
    for (var i = data.length - 1; i >= 0; i--) {
      if ((data[i][1] || '').toString().toLowerCase() === email.toLowerCase() && !data[i][4]) {
        var clockInTime = new Date(data[i][3]);
        var hours = Math.round((now - clockInTime) / 3600000 * 100) / 100;
        var rowNum = i + 2;
        sh.getRange(rowNum, 5).setValue(now);
        sh.getRange(rowNum, 6).setValue(hours);
        return { success: true, clockOut: Utilities.formatDate(now, tz, 'h:mm a'), hours: hours };
      }
    }
    return { error: 'No open clock-in found' };
  } catch(e) { return { error: e.toString() }; }
}

function getClockStatus(payload) {
  try {
    var email = payload.email;
    var sh    = getTimeSheet_();
    var tz    = Session.getScriptTimeZone();
    var lastRow = sh.getLastRow();

    if (lastRow >= 2) {
      var data = sh.getRange(2, 1, lastRow - 1, 6).getValues();
      for (var i = data.length - 1; i >= 0; i--) {
        if ((data[i][1] || '').toString().toLowerCase() === email.toLowerCase() && !data[i][4]) {
          return { clockedIn: true, since: Utilities.formatDate(new Date(data[i][3]), tz, 'h:mm a') };
        }
      }
    }
    return { clockedIn: false };
  } catch(e) { return { error: e.toString() }; }
}

function getTimesheet(payload) {
  try {
    var email  = payload.email;
    var role   = payload.role || 'runner';
    var sh     = getTimeSheet_();
    var tz     = Session.getScriptTimeZone();
    var now    = new Date();
    var bounds = getPeriodBounds(now);
    var pStart = bounds.start;
    var pEnd   = bounds.end;

    // Format label: "Jul 1 - Jul 15" or "Jul 16 - Jul 31"
    var months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    var sm = months[pStart.getMonth()];
    var em = months[pEnd.getMonth()];
    var periodLabel = sm + ' ' + pStart.getDate() + ' - ' + em + ' ' + pEnd.getDate();

    var myHours   = 0;
    var myDays    = {};
    var empMap    = {}; // email -> { name, hours }

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
        if (rowEmail === email.toLowerCase().trim()) {
          myHours += rowHours;
          var dayLabel = months[rd.getMonth()] + ' ' + rd.getDate();
          myDays[dayLabel] = Math.round(((myDays[dayLabel] || 0) + rowHours) * 100) / 100;
        }

        // All employees (admin)
        if (role === 'admin') {
          if (!empMap[rowEmail]) empMap[rowEmail] = { name: (data[i][0] || rowEmail).toString().trim(), hours: 0 };
          empMap[rowEmail].hours = Math.round((empMap[rowEmail].hours + rowHours) * 100) / 100;
        }
      }
    }

    var allEmployees = [];
    if (role === 'admin') {
      allEmployees = Object.keys(empMap).map(function(e) {
        return { name: empMap[e].name, hours: empMap[e].hours };
      }).sort(function(a, b) { return b.hours - a.hours; });
    }

    return {
      myHours:      Math.round(myHours * 100) / 100,
      myDays:       myDays,
      periodLabel:  periodLabel,
      allEmployees: allEmployees
    };
  } catch(e) { return { error: e.toString() }; }
}

// ── Admin: Employee Manager ───────────────────────────────────────────────────
function getEmployees(payload) {
  try {
    var auth = authorizeCaller(payload, ['admin']);
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
    var auth = authorizeCaller(payload, ['admin']);
    if (!auth.ok) return { error: auth.error, code: auth.code };
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(ROLES_SHEET);
    if (!sh) return { error: 'HR sheet not found' };
    var email = (payload.email || '').toLowerCase().trim();
    if (!email || !payload.name) return { error: 'Name and email are required' };
    if (isOwnerEmail(email) && (payload.role || '').toLowerCase().trim() !== 'admin') {
      return { error: 'This account is protected and must be added as admin.', code: 'OWNER_PROTECTED' };
    }
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
      (payload.role  || 'runner').trim(),
      (payload.password || '').trim(),
      parseFloat(payload.allotted) || 0,
      0
    ]);
    return { success: true };
  } catch(e) { return { error: e.toString() }; }
}

function updateEmployee(payload) {
  try {
    var auth = authorizeCaller(payload, ['admin']);
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
        var newRole = payload.role !== undefined ? payload.role.toString().toLowerCase().trim() : undefined;
        if (newRole !== undefined && newRole !== 'admin') {
          if (isOwnerEmail(email)) {
            return { error: 'This account is protected and must remain admin.', code: 'OWNER_PROTECTED' };
          }
          var currentRole = (data[i][3] || '').toString().toLowerCase().trim();
          if (currentRole === 'admin' && countAdminRows(data) <= 1) {
            return { error: 'Cannot demote the last remaining admin.', code: 'LAST_ADMIN_PROTECTED' };
          }
        }
        var row = i + 2;
        if (payload.name     !== undefined) sh.getRange(row, 1).setValue(payload.name);
        if (payload.phone    !== undefined) sh.getRange(row, 3).setValue(payload.phone);
        if (payload.role     !== undefined) sh.getRange(row, 4).setValue(payload.role);
        if (payload.password !== undefined && payload.password !== '') sh.getRange(row, 5).setValue(payload.password);
        if (payload.allotted !== undefined) sh.getRange(row, 6).setValue(parseFloat(payload.allotted) || 0);
        return { success: true };
      }
    }
    return { error: 'Employee not found' };
  } catch(e) { return { error: e.toString() }; }
}

function removeEmployee(payload) {
  try {
    var auth = authorizeCaller(payload, ['admin']);
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
        var currentRole = (data[i][3] || '').toString().toLowerCase().trim();
        if (currentRole === 'admin' && countAdminRows(data) <= 1) {
          return { error: 'Cannot remove the last remaining admin.', code: 'LAST_ADMIN_PROTECTED' };
        }
        sh.deleteRow(i + 2);
        return { success: true };
      }
    }
    return { error: 'Employee not found' };
  } catch(e) { return { error: e.toString() }; }
}

// ── Admin: PTO Overview ───────────────────────────────────────────────────────
function getPTOOverview(payload) {
  try {
    var auth = authorizeCaller(payload, ['admin']);
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

// ── Admin: Payroll Summary ────────────────────────────────────────────────────
function getPayrollSummary(payload) {
  try {
    var auth = authorizeCaller(payload, ['admin']);
    if (!auth.ok) return { error: auth.error, code: auth.code };
    var sh  = getTimeSheet_();
    var tz  = Session.getScriptTimeZone();
    var now = new Date();
    var bounds = getPeriodBounds(now);
    var pStart = bounds.start;
    var pEnd   = bounds.end;
    var months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    var periodLabel = months[pStart.getMonth()] + ' ' + pStart.getDate() + ' - ' + months[pEnd.getMonth()] + ' ' + pEnd.getDate();
    // Build email->name lookup from HR sheet (authoritative source)
    var hrNameMap = {};
    var hrSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('HR');
    if (hrSheet) {
      var hrData = hrSheet.getDataRange().getValues();
      for (var h = 1; h < hrData.length; h++) {
        var hrEmail = (hrData[h][1] || '').toString().toLowerCase().trim();
        var hrName  = (hrData[h][0] || '').toString().trim();
        if (hrEmail) hrNameMap[hrEmail] = hrName;
      }
    }
    var empMap = {};
    var lastRow = sh.getLastRow();
    if (lastRow >= 2) {
      var data = sh.getRange(2, 1, lastRow - 1, 6).getValues();
      for (var i = 0; i < data.length; i++) {
        var rowEmail = (data[i][1] || '').toString().toLowerCase().trim();
        var rowDate  = data[i][2] ? new Date(data[i][2]) : null;
        var rowHours = parseFloat(data[i][5]) || 0;
        if (!rowDate || !rowEmail) continue;
        var rd = new Date(rowDate.getFullYear(), rowDate.getMonth(), rowDate.getDate());
        if (rd < pStart || rd > pEnd) continue;
        if (!empMap[rowEmail]) {
          var resolvedName = hrNameMap[rowEmail] || (data[i][0] || '').toString().trim() || rowEmail;
          empMap[rowEmail] = { name: resolvedName, total: 0, days: {} };
        }
        empMap[rowEmail].total = Math.round((empMap[rowEmail].total + rowHours) * 100) / 100;
        var dayLabel = months[rd.getMonth()] + ' ' + rd.getDate();
        empMap[rowEmail].days[dayLabel] = Math.round(((empMap[rowEmail].days[dayLabel] || 0) + rowHours) * 100) / 100;
      }
    }
    var employees = Object.keys(empMap).map(function(e) {
      return { email: e, name: empMap[e].name, total: empMap[e].total, days: empMap[e].days };
    }).sort(function(a, b) { return a.name.localeCompare(b.name); });
    return { employees: employees, periodLabel: periodLabel };
  } catch(e) { return { error: e.toString() }; }
}

function emailPayroll(payload) {
  try {
    var auth = authorizeCaller(payload, ['admin']);
    if (!auth.ok) return { error: auth.error, code: auth.code };
    var to = payload.to || Session.getActiveUser().getEmail();
    var summary = getPayrollSummary(payload);
    if (summary.error) return { error: summary.error };
    var lines = ['Payroll Summary - ' + summary.periodLabel, '===========================', ''];
    var grandTotal = 0;
    summary.employees.forEach(function(e) {
      lines.push(e.name + ': ' + e.total + ' hrs');
      var dayKeys = Object.keys(e.days);
      dayKeys.forEach(function(d) { lines.push('  ' + d + ': ' + e.days[d] + ' hrs'); });
      lines.push('');
      grandTotal += e.total;
    });
    lines.push('---------------------------');
    lines.push('Grand Total: ' + Math.round(grandTotal * 100) / 100 + ' hrs');
    GmailApp.sendEmail(to, 'Payroll Summary - ' + summary.periodLabel, lines.join('\n'));
    return { success: true };
  } catch(e) { return { error: e.toString() }; }
}