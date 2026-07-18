function onFormSubmitReceived(e) {

  // Spreadsheet + sheet
  var ss = e.source;
  var sheet = ss.getSheetByName("PO Database");

  var responses = e.namedValues;

  // =============================
  // CONFIG
  // =============================
  var FINAL_RECONCILE_STATUS = "Ready to Reconcile";

  // --- emails ---
  var runnerEmail = (responses["What is your email?"] || [""])[0];
  var adminEmail = "aidan@panoramicbuildingllc.com";
  var recipientList = runnerEmail
    ? adminEmail + "," + runnerEmail
    : adminEmail;

  // --- form values ---
  var hasPO = (responses["Do you have a PO number?"] || [""])[0];
  var poInput = (responses["Enter PO Number (Format ex. 26-01-###)"] || [""])[0];
  var job = (responses["Enter Job Name"] || [""])[0];
  var vendor = (responses["Vendor"] || [""])[0];
  var notes = (responses["Any other Notes"] || [""])[0];
  var status = (responses["Status of PO"] || [""])[0];

  // File upload → Drive URL
  var imageLinks = responses["Add Picture of Received Note"] || [];
  var imageLink = imageLinks.length ? imageLinks[0] : "";

  var now = new Date();
  var today = Utilities.formatDate(
    now,
    Session.getScriptTimeZone(),
    "MM/dd/yyyy"
  );

  // ====================================================
  // PATH A — PO EXISTS
  // ====================================================
  if (hasPO === "Yes" && poInput) {
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return;

    var poValues = sheet.getRange(2, 1, lastRow - 1).getValues();

    for (var i = 0; i < poValues.length; i++) {
      if (poValues[i][0] === poInput) {
        var row = i + 2;

        // 🔍 Check if a received note already exists
        var existingReceivedNote = sheet.getRange(row, 11).getValue(); // Column K

        // ❌ If already exists → block overwrite
        if (existingReceivedNote) {
          MailApp.sendEmail(
            recipientList,
            "Received Note Already Exists: " + poInput,
            "A received note is already attached to this PO.\n\n" +
            "PO Number: " + poInput + "\n\n" +
            "No changes were made.\n" +
            "Please contact Aidan if this needs to be updated."
          );
          return;
        }

        // ✅ Otherwise proceed normally
        var finalStatus = imageLink
          ? FINAL_RECONCILE_STATUS
          : status;

        if (imageLink) {
          sheet.getRange(row, 11).setValue(imageLink); // K Received Note
        }
        sheet.getRange(row, 7).setValue(finalStatus); // G Status

        MailApp.sendEmail(
          recipientList,
          "PO Updated: " + poInput,
          "A received note has been added.\n\n" +
          "PO Number: " + poInput + "\n" +
          "Status: " + finalStatus
        );
        return;
      }
    }

    // PO not found
    MailApp.sendEmail(
      adminEmail,
      "⚠️ PO Not Found",
      "The PO number '" + poInput + "' could not be found in the PO Database."
    );
    return;
  }

  // ====================================================
  // PATH B — NO PO: CREATE NEW PO
  // ====================================================

  // --- build PO number (YY-QQ-ROW) ---
  var year = Utilities.formatDate(now, Session.getScriptTimeZone(), "yy");
  var quarter = Math.ceil((now.getMonth() + 1) / 3);
  var paddedQuarter = ("0" + quarter).slice(-2);

  var nextRow = sheet.getLastRow() + 1;

  var poNumber =
    year + "-" + paddedQuarter + "-" +
    Utilities.formatString("%03d", nextRow);

  var finalStatus = imageLink
    ? FINAL_RECONCILE_STATUS
    : status;

  // Write new PO row
  sheet.getRange(nextRow, 1).setValue(poNumber);     // A PO Num
  sheet.getRange(nextRow, 2).setValue(today);        // B Date Issued
  sheet.getRange(nextRow, 4).setValue(job);          // D Job Reference
  sheet.getRange(nextRow, 5).setValue(vendor);       // E Vendor
  sheet.getRange(nextRow, 7).setValue(finalStatus);  // G Status
  sheet.getRange(nextRow, 11).setValue(imageLink);   // K Received Note
  sheet.getRange(nextRow, 12).setValue(notes);       // L Notes

  // Email confirmation
  MailApp.sendEmail(
    recipientList,
    "New PO Created: " + poNumber,
    "A new PO has been created.\n\n" +
    "PO Number: " + poNumber + "\n" +
    "Job: " + job + "\n" +
    "Vendor: " + vendor + "\n" +
    "Status: " + finalStatus
  );
}
