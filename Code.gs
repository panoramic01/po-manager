/*************************************
 * CREATE NEW PO
 *************************************/
function createNewPO() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

 // INPUT ROW (Row 2)
var builder = sheet.getRange("B2").getValue();
var jobRef = sheet.getRange("C2").getValue();
var vendor = sheet.getRange("D2").getValue();
var vendorInvoice = sheet.getRange("E2").getValue();
var status = sheet.getRange("F2").getValue();
var notes = sheet.getRange("G2").getValue();
var invoiceTotal = sheet.getRange("H2").getValue();


  // VALIDATION
if (!jobRef || !vendor) {
  SpreadsheetApp.getActiveSpreadsheet().toast(
    "⚠️ Please enter both Job and Vendor before creating a PO."
  );
  return;
}


  // BUILD PO NUMBER (YY-QQ-###)
  var now = new Date();
  var year = Utilities.formatDate(now, Session.getScriptTimeZone(), "yy");
  var quarter = Math.ceil((now.getMonth() + 1) / 3);
  var paddedQuarter = ("0" + quarter).slice(-2);

  var lastRow = sheet.getLastRow();
  var nextRow = lastRow + 1;

  var poNumber =
    year + "-" + paddedQuarter + "-" +
    Utilities.formatString("%03d", nextRow);

  var today = Utilities.formatDate(
    now,
    Session.getScriptTimeZone(),
    "MM/dd/yyyy"
  );

  // WRITE TO PO DATABASE
  sheet.getRange(nextRow, 1).setValue(poNumber);      // A PO Num
sheet.getRange(nextRow, 2).setValue(today);         // B Date Issued
sheet.getRange(nextRow, 3).setValue(builder);       // C Builder
sheet.getRange(nextRow, 4).setValue(jobRef);        // D Job Reference
sheet.getRange(nextRow, 5).setValue(vendor);        // E Vendor
sheet.getRange(nextRow, 6).setValue(vendorInvoice); // F Vendor Invoice #
sheet.getRange(nextRow, 7).setValue(status);        // G Status
sheet.getRange(nextRow, 8).setValue(invoiceTotal);  // H Invoice Total
sheet.getRange(nextRow, 12).setValue(notes);        // L Notes


  // CLEAR INPUTS
  sheet.getRangeList(["B2","C2","D2","E2","F2","G2","H2"]).clearContent();


  SpreadsheetApp.getActiveSpreadsheet().toast(
    "✅ PO " + poNumber + " added successfully!"
  );
}
/*************************************
 * MORNING DELIVERY EMAIL
 *************************************/
var ADMIN_EMAIL = "aidan@panoramicbuildingllc.com, tyson@panoramicbuildingllc.com";
var SHEET_URL   = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
 
/*************************************
 * MAIN FUNCTION — builds and sends the digest
 *************************************/
function sendMorningPODigest() {
 
  // ── Skip weekends ──────────────────────────────
  var now     = new Date();
  var dayOfWeek = now.getDay(); // 0 = Sunday, 6 = Saturday
  if (dayOfWeek === 0 || dayOfWeek === 6) {
    Logger.log("Weekend — no digest sent.");
    return;
  }
 
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("PO Database");
 
  if (!sheet) {
    Logger.log("ERROR: Sheet 'PO Database' not found.");
    return;
  }
 
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log("No data rows found in PO Database.");
    return;
  }
 
  var data = sheet.getRange(2, 1, lastRow - 1, 12).getValues();
 
  var today = new Date();
  today.setHours(0, 0, 0, 0);
 
  var nextWeek = new Date(today);
  nextWeek.setDate(today.getDate() + 7);
 
  // ── Buckets — split by Delivery vs Pickup ──────
  var todayDelivery    = [];
  var todayPickup      = [];
  var upcomingDelivery = [];
  var upcomingPickup   = [];
  var overdueDelivery  = [];
  var overduePickup    = [];
  var missingDelivery  = [];
  var missingPickup    = [];
  var totalOpenValue   = 0;
  var totalOpenCount   = 0;
 
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
 
    var poNum        = row[0];
    var dateIssued   = row[1];
    var builder      = row[2];
    var jobRef       = row[3];
    var vendor       = row[4];
    var vendorInv    = row[5];
    var status       = row[6] ? row[6].toString().trim() : "";
    var invoiceTotal = row[7];
    var deliveryDate = row[8];
    var notes        = row[11];
 
    if (!poNum) continue;
 
    var statusLower       = status.toLowerCase();
    var isPendingDelivery = statusLower === "pending delivery";
    var isPendingPickup   = statusLower === "pending pickup";
    if (!isPendingDelivery && !isPendingPickup) continue;
 
    totalOpenCount++;
    if (invoiceTotal && !isNaN(invoiceTotal)) {
      totalOpenValue += Number(invoiceTotal);
    }
 
    var po = {
      poNum:        poNum,
      dateIssued:   dateIssued,
      builder:      builder      || "—",
      jobRef:       jobRef       || "—",
      vendor:       vendor       || "—",
      invoiceTotal: invoiceTotal,
      deliveryDate: deliveryDate,
      status:       status       || "—",
      notes:        notes        || ""
    };
 
    var hasValidDate = deliveryDate && (deliveryDate instanceof Date) && !isNaN(deliveryDate);
 
    if (hasValidDate) {
      var delDay = new Date(deliveryDate);
      delDay.setHours(0, 0, 0, 0);
 
      if (delDay.getTime() === today.getTime()) {
        isPendingDelivery ? todayDelivery.push(po) : todayPickup.push(po);
      } else if (delDay > today && delDay <= nextWeek) {
        isPendingDelivery ? upcomingDelivery.push(po) : upcomingPickup.push(po);
      } else if (delDay < today) {
        isPendingDelivery ? overdueDelivery.push(po) : overduePickup.push(po);
      }
    } else {
      isPendingDelivery ? missingDelivery.push(po) : missingPickup.push(po);
    }
  }
 
  // Sort upcoming soonest first, overdue most overdue first
  function byDate(a, b) { return new Date(a.deliveryDate) - new Date(b.deliveryDate); }
  upcomingDelivery.sort(byDate);
  upcomingPickup.sort(byDate);
  overdueDelivery.sort(byDate);
  overduePickup.sort(byDate);
 
  var subject  = buildSubject(todayDelivery, todayPickup, overdueDelivery, overduePickup, missingDelivery, missingPickup);
  var htmlBody = buildHTML(
    today, totalOpenCount, totalOpenValue,
    todayDelivery, todayPickup,
    upcomingDelivery, upcomingPickup,
    overdueDelivery, overduePickup,
    missingDelivery, missingPickup
  );
 
  MailApp.sendEmail({
    to:       ADMIN_EMAIL,
    subject:  subject,
    htmlBody: htmlBody
  });
 
  Logger.log("Morning PO Digest sent to " + ADMIN_EMAIL);
}
 
 
/*************************************
 * EMAIL SUBJECT LINE
 *************************************/
function buildSubject(todayDelivery, todayPickup, overdueDelivery, overduePickup, missingDelivery, missingPickup) {
  var tz      = Session.getScriptTimeZone();
  var dateStr = Utilities.formatDate(new Date(), tz, "EEE, MMM d");
  var flags   = [];
 
  var todayTotal   = todayDelivery.length + todayPickup.length;
  var overdueTotal = overdueDelivery.length + overduePickup.length;
  var missingTotal = missingDelivery.length + missingPickup.length;
 
  if (todayTotal   > 0) flags.push(todayTotal + " today");
  if (overdueTotal > 0) flags.push(overdueTotal + " overdue");
  if (missingTotal > 0) flags.push(missingTotal + " need dates");
 
  var prefix = (overdueTotal + missingTotal > 0) ? "⚠️" : "📦";
  return flags.length > 0
    ? prefix + " PO Digest — " + dateStr + " · " + flags.join(", ")
    : "📦 PO Digest — " + dateStr + " · All clear";
}
 
 
/*************************************
 * FULL HTML EMAIL BODY
 *************************************/
function buildHTML(today, totalOpenCount, totalOpenValue,
                   todayDelivery, todayPickup,
                   upcomingDelivery, upcomingPickup,
                   overdueDelivery, overduePickup,
                   missingDelivery, missingPickup) {
 
  var tz      = Session.getScriptTimeZone();
  var dateStr = Utilities.formatDate(today, tz, "EEEE, MMMM d, yyyy");
 
  var css = [
    "body{font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#333;background:#f4f4f4;margin:0;padding:20px;}",
    ".wrap{max-width:780px;margin:0 auto;background:#fff;border-radius:8px;padding:28px 32px;box-shadow:0 2px 8px rgba(0,0,0,.08);}",
    "h1{color:#1a1a2e;margin:0 0 4px;font-size:22px;}",
    ".subtitle{color:#999;font-size:13px;margin:0 0 20px;}",
    ".stats{display:flex;gap:0;border:1px solid #e8e8e8;border-radius:6px;overflow:hidden;margin-bottom:24px;}",
    ".stat{flex:1;text-align:center;padding:14px 8px;border-right:1px solid #e8e8e8;}",
    ".stat:last-child{border-right:none;}",
    ".stat .n{font-size:24px;font-weight:bold;line-height:1;}",
    ".stat .l{font-size:10px;color:#999;text-transform:uppercase;margin-top:4px;letter-spacing:.5px;}",
    ".n-red{color:#c62828;}.n-amber{color:#e65100;}.n-green{color:#2e7d32;}.n-blue{color:#1565c0;}",
    "h2{font-size:14px;font-weight:bold;margin:24px 0 6px;padding:8px 12px;border-radius:4px;}",
    ".h-green{background:#e8f5e9;color:#2e7d32;}",
    ".h-teal{background:#e0f2f1;color:#00695c;}",
    ".h-blue{background:#e3f2fd;color:#1565c0;}",
    ".h-indigo{background:#e8eaf6;color:#283593;}",
    ".h-red{background:#fbe9e7;color:#bf360c;}",
    ".h-pink{background:#fce4ec;color:#880e4f;}",
    ".h-amber{background:#fff8e1;color:#e65100;}",
    ".h-orange{background:#fff3e0;color:#bf360c;}",
    "table{width:100%;border-collapse:collapse;font-size:12.5px;}",
    "th{background:#fafafa;text-align:left;padding:7px 10px;font-size:11px;color:#888;border-bottom:2px solid #eee;white-space:nowrap;}",
    "td{padding:7px 10px;border-bottom:1px solid #f0f0f0;vertical-align:top;}",
    "tr:last-child td{border-bottom:none;}",
    ".po{font-weight:bold;white-space:nowrap;}",
    ".amt{font-family:monospace;}",
    ".badge{display:inline-block;padding:1px 7px;border-radius:10px;font-size:10px;font-weight:bold;white-space:nowrap;}",
    ".b-today{background:#c8e6c9;color:#1b5e20;}",
    ".b-overdue{background:#ffcdd2;color:#b71c1c;}",
    ".note{color:#aaa;font-size:11px;}",
    ".none{color:#bbb;font-style:italic;padding:10px 0;font-size:13px;}",
    ".footer{margin-top:28px;padding-top:12px;border-top:1px solid #eee;font-size:11px;color:#ccc;}"
  ].join("");
 
  var h = '<!DOCTYPE html><html><head><meta charset="utf-8"><style>' + css + '</style></head><body><div class="wrap">';
 
  h += '<h1>📦 PO Morning Digest</h1><p class="subtitle">' + dateStr + '</p>';
 
  // Stats bar
  var openValueFmt  = "$" + formatCurrency(totalOpenValue);
  var overdueTotal  = overdueDelivery.length + overduePickup.length;
  var missingTotal  = missingDelivery.length + missingPickup.length;
  h += '<div class="stats">';
  h += statBox(todayDelivery.length,    "Deliveries Today",  todayDelivery.length > 0 ? "n-green" : "");
  h += statBox(todayPickup.length,      "Pickups Today",     todayPickup.length > 0 ? "n-teal" : "");
  h += statBox(overdueTotal,            "Overdue",           overdueTotal > 0 ? "n-red" : "n-green");
  h += statBox(missingTotal,            "Missing Date",      missingTotal > 0 ? "n-amber" : "n-green");
  h += statBox(totalOpenCount,          "Open POs",          "");
  h += statBox(openValueFmt,            "Open Value",        "");
  h += '</div>';
 
  // ── TODAY — DELIVERIES ─────────────────────────
  h += '<h2 class="h-green">🚚 Today\'s Deliveries (' + todayDelivery.length + ')</h2>';
  h += todayDelivery.length === 0
    ? '<p class="none">No deliveries scheduled for today.</p>'
    : buildTable(todayDelivery, true, false);
 
  // ── TODAY — PICKUPS ────────────────────────────
  h += '<h2 class="h-teal">🏪 Today\'s Pickups (' + todayPickup.length + ')</h2>';
  h += todayPickup.length === 0
    ? '<p class="none">No pickups scheduled for today.</p>'
    : buildTable(todayPickup, true, false);
 
  // ── UPCOMING — DELIVERIES ──────────────────────
  h += '<h2 class="h-blue">📅 Upcoming Deliveries — This Week (' + upcomingDelivery.length + ')</h2>';
  h += upcomingDelivery.length === 0
    ? '<p class="none">No deliveries scheduled in the next 7 days.</p>'
    : buildTable(upcomingDelivery, true, false);
 
  // ── UPCOMING — PICKUPS ─────────────────────────
  h += '<h2 class="h-indigo">📅 Upcoming Pickups — This Week (' + upcomingPickup.length + ')</h2>';
  h += upcomingPickup.length === 0
    ? '<p class="none">No pickups scheduled in the next 7 days.</p>'
    : buildTable(upcomingPickup, true, false);
 
  // ── OVERDUE — DELIVERIES ───────────────────────
  h += '<h2 class="h-red">🔴 Overdue Deliveries (' + overdueDelivery.length + ')</h2>';
  if (overdueDelivery.length === 0) {
    h += '<p class="none">✅ No overdue deliveries.</p>';
  } else {
    h += '<p style="color:#bf360c;font-size:12px;margin:4px 0 8px;">These deliveries have passed their scheduled date. Follow up with the vendor or update the date.</p>';
    h += buildTable(overdueDelivery, true, true);
  }
 
  // ── OVERDUE — PICKUPS ──────────────────────────
  h += '<h2 class="h-pink">🔴 Overdue Pickups (' + overduePickup.length + ')</h2>';
  if (overduePickup.length === 0) {
    h += '<p class="none">✅ No overdue pickups.</p>';
  } else {
    h += '<p style="color:#880e4f;font-size:12px;margin:4px 0 8px;">These pickups have passed their scheduled date. Arrange pickup or update the date.</p>';
    h += buildTable(overduePickup, true, true);
  }
 
  // ── MISSING DATE — DELIVERIES ──────────────────
  h += '<h2 class="h-amber">🟡 Deliveries — No Date Set (' + missingDelivery.length + ')</h2>';
  if (missingDelivery.length === 0) {
    h += '<p class="none">✅ All pending deliveries have a date.</p>';
  } else {
    h += '<p style="color:#e65100;font-size:12px;margin:4px 0 8px;">Enter an expected delivery date so these show up in the schedule.</p>';
    h += buildTable(missingDelivery, false, false);
  }
 
  // ── MISSING DATE — PICKUPS ─────────────────────
  h += '<h2 class="h-orange">🟡 Pickups — No Date Set (' + missingPickup.length + ')</h2>';
  if (missingPickup.length === 0) {
    h += '<p class="none">✅ All pending pickups have a date.</p>';
  } else {
    h += '<p style="color:#bf360c;font-size:12px;margin:4px 0 8px;">Enter an expected pickup date so these show up in the schedule.</p>';
    h += buildTable(missingPickup, false, false);
  }
 
  h += '<div class="footer">Auto-generated by PO Morning Digest &nbsp;·&nbsp; <a href="' + SHEET_URL + '" style="color:#aaa;">Open PO Database →</a></div>';
  h += '</div></body></html>';
  return h;
}
 
 
/*************************************
 * HTML HELPERS
 *************************************/
function statBox(value, label, colorClass) {
  return '<div class="stat"><div class="n ' + (colorClass || "") + '">' + value + '</div><div class="l">' + label + '</div></div>';
}
 
function buildTable(rows, showDeliveryDate, isOverdue) {
  var html = '<table><tr>'
    + '<th>PO #</th>'
    + '<th>Builder</th>'
    + '<th>Job</th>'
    + '<th>Vendor</th>'
    + (showDeliveryDate ? '<th>Date</th>' : '')
    + '<th>Invoice Total</th>'
    + '<th>Notes</th>'
    + '</tr>';
 
  var tz = Session.getScriptTimeZone();
 
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
 
    var dateCell = "";
    if (showDeliveryDate) {
      if (r.deliveryDate && r.deliveryDate instanceof Date && !isNaN(r.deliveryDate)) {
        var fmtDate = Utilities.formatDate(r.deliveryDate, tz, "MM/dd/yyyy");
        if (isOverdue) {
          dateCell = '<td><span class="badge b-overdue">OVERDUE</span> ' + fmtDate + '</td>';
        } else {
          var isToday = new Date(r.deliveryDate).setHours(0,0,0,0) === new Date().setHours(0,0,0,0);
          dateCell = isToday
            ? '<td><span class="badge b-today">TODAY</span> ' + fmtDate + '</td>'
            : '<td>' + fmtDate + '</td>';
        }
      } else {
        dateCell = '<td style="color:#e65100;font-style:italic;">Not set</td>';
      }
    }
 
    var totalCell = (r.invoiceTotal && !isNaN(r.invoiceTotal) && Number(r.invoiceTotal) > 0)
      ? '<td class="amt">$' + formatCurrency(Number(r.invoiceTotal)) + '</td>'
      : '<td style="color:#ccc;">—</td>';
 
    html += '<tr>'
      + '<td class="po">' + r.poNum + '</td>'
      + '<td>' + r.builder + '</td>'
      + '<td>' + r.jobRef + '</td>'
      + '<td>' + r.vendor + '</td>'
      + (showDeliveryDate ? dateCell : '')
      + totalCell
      + '<td class="note">' + r.notes + '</td>'
      + '</tr>';
  }
 
  return html + '</table>';
}
 
function formatCurrency(num) {
  var parts = Number(num).toFixed(2).split(".");
  parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ",");
  return parts.join(".");
}
 
 
/*************************************
 * SETUP: Schedule daily 7 AM trigger
 * Run this function ONCE from the Apps Script editor.
 *************************************/
function createMorningTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "sendMorningPODigest") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
 
  ScriptApp.newTrigger("sendMorningPODigest")
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .create();
 
  SpreadsheetApp.getActiveSpreadsheet().toast(
    "✅ Morning digest scheduled for 7:00 AM daily (Mon–Fri only).",
    "PO Digest",
    5
  );
}
 
 
/*************************************
 * TEST: Send a digest email right now
 *************************************/
function testMorningDigest() {
  sendMorningPODigest();
  SpreadsheetApp.getActiveSpreadsheet().toast(
    "📧 Test digest sent to " + ADMIN_EMAIL,
    "PO Digest",
    5
  );
}




/*************************************
 * MENU ON OPEN
 *************************************/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📋 PO Tools")
    .addItem("Hide Completed Rows", "hideCompletedRows")
    .addItem("Show All Rows", "showAllRows")
    .addToUi();
}

/*************************************
 * HIDE COMPLETED ROWS
 *************************************/
function hideCompletedRows() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PO Database");
  var statusColumn = 6; // Column F
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var statuses = sheet.getRange(2, statusColumn, lastRow - 1).getValues();
  var hiddenCount = 0;

  for (var i = statuses.length - 1; i >= 0; i--) {
    var value = statuses[i][0];
    if (value && value.toString().trim().toLowerCase() === "complete") {
      sheet.hideRows(i + 2);
      hiddenCount++;
    }
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "📦 " + hiddenCount + " completed row(s) hidden."
  );
}

/*************************************
 * SHOW ALL ROWS
 *************************************/
function showAllRows() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PO Database");
  sheet.showRows(1, sheet.getMaxRows());

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "👀 All rows are now visible."
  );
}
