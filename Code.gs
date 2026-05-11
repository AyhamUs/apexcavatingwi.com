// AP Excavating — Contact Form Handler
// Deploy as a Web App: Execute as Me | Who has access: Anyone

var NOTIFY_EMAIL = 'ayhamissa416@gmail.com';
var SHEET_NAME   = 'Submissions';

function doPost(e) {
  try {
    var p = e.parameter;

    // ── Write to spreadsheet ──────────────────────────────────────────────
    var ss    = getOrCreateSpreadsheet();
    var sheet = getOrCreateSheet(ss);
    sheet.appendRow([
      new Date(),
      p.firstName || '',
      p.lastName  || '',
      p.email     || '',
      p.phone     || '',
      p.service   || '',
      p.message   || ''
    ]);

    // ── Send notification email ───────────────────────────────────────────
    var subject = 'New Quote Request — ' + (p.firstName || '') + ' ' + (p.lastName || '');
    var body =
      'A new quote request was submitted on apexcavating.com.\n\n' +
      'Name:    ' + (p.firstName || '') + ' ' + (p.lastName || '') + '\n' +
      'Email:   ' + (p.email   || '') + '\n' +
      'Phone:   ' + (p.phone   || '') + '\n' +
      'Service: ' + (p.service || '') + '\n\n' +
      'Message:\n' + (p.message || '') + '\n\n' +
      '——\nSubmitted: ' + new Date().toLocaleString();

    MailApp.sendEmail({ to: NOTIFY_EMAIL, subject: subject, body: body });

    return json({ success: true });

  } catch (err) {
    return json({ success: false, error: err.message });
  }
}

// ── Helpers ───────────────────────────────────────────────────────────────

function getOrCreateSpreadsheet() {
  var files = DriveApp.getFilesByName('AP Excavating — Form Submissions');
  if (files.hasNext()) {
    return SpreadsheetApp.open(files.next());
  }
  var ss = SpreadsheetApp.create('AP Excavating — Form Submissions');
  return ss;
}

function getOrCreateSheet(ss) {
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  // Add header row if the sheet is empty
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Timestamp', 'First Name', 'Last Name', 'Email', 'Phone', 'Service', 'Message']);
    sheet.getRange(1, 1, 1, 7).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
