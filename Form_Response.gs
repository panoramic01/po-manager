function onFormSubmit(e) {
  var sheet = e.source.getSheetByName("PO Database");
  

  // --- get form responses ---
  var builder = e.namedValues["Builder Name"] 
    ? e.namedValues["Builder Name"][0] 
    : "";

  var jobRef = e.namedValues["Job Reference"][0];
  var vendor = e.namedValues["Vendor"] 
    ? e.namedValues["Vendor"][0] 
    : "";

  var status = e.namedValues["Status"] 
    ? e.namedValues["Status"][0] 
    : "Draft";

  var notes = e.namedValues["Notes"] 
    ? e.namedValues["Notes"][0] 
    : "";

  // --- build PO number (YY-QQ-###) ---
  var now = new Date();
  var year = Utilities.formatDate(now, Session.getScriptTimeZone(), "yy");
  var quarter = Math.ceil((now.getMonth() + 1) / 3);
  var paddedQuarter = ("0" + quarter).slice(-2);

  // count existing POs for this year+quarter
  var poCol = sheet.getRange(2, 1, sheet.getLastRow() - 1).getValues();
  var count = 0;

  poCol.forEach(function(row) {
    if (row[0] && row[0].startsWith(year + "-" + paddedQuarter)) {
      count++;
    }
  });

  var poNumber = year + "-" + paddedQuarter + "-" +
    Utilities.formatString("%03d", count + 1);

  var today = Utilities.formatDate(
    now,
    Session.getScriptTimeZone(),
    "MM/dd/yyyy"
  );

  var nextRow = sheet.getLastRow() + 1;

  // --- write to PO Database ---
  sheet.getRange(nextRow, 1).setValue(poNumber); // A PO Num
  sheet.getRange(nextRow, 2).setValue(today);    // B Date Issued
  sheet.getRange(nextRow, 3).setValue(builder);  // C Builder
  sheet.getRange(nextRow, 4).setValue(jobRef);   // D Job Reference
  sheet.getRange(nextRow, 5).setValue(vendor);   // E Vendor
  sheet.getRange(nextRow, 7).setValue(status);   // G Status
  sheet.getRange(nextRow, 12).setValue(notes);   // L Notes

  // --- confirmation email ---
  var recipient = "aidan@panoramicbuildingllc.com";
  var subject = "PO Created: " + poNumber + " (" + jobRef + ")";
  var body =
    "A new Purchase Order has been created.\n\n" +
    "PO Number: " + poNumber + "\n" +
    "Builder: " + builder + "\n" +
    "Job Reference: " + jobRef + "\n" +
    "Vendor: " + vendor + "\n" +
    "Status: " + status + "\n\n" +
    "View it in the PO Database sheet.";

  MailApp.sendEmail(recipient, subject, body);
}
