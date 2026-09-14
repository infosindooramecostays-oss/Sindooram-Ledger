// ===== Sindooram ID Vault — Apps Script =====
// A separate, standalone script from the main Ledger one. It only reads
// booking basics (no money, status, or guest contact fields) from the
// Ledger Sheet, and saves uploaded documents to Drive. Deploy this as its
// own Web App with its own URL — do not reuse the Ledger's URL, since
// that one exposes full financial data.

// Paste your Ledger Google Sheet's ID here — open the Ledger Sheet, look
// at its URL: https://docs.google.com/spreadsheets/d/THIS_PART/edit
var LEDGER_SHEET_ID = 'PASTE_YOUR_LEDGER_SHEET_ID_HERE';

// The Drive folder everything gets saved under. Created automatically the
// first time anyone uploads a document — nothing to set up by hand.
var VAULT_FOLDER_NAME = 'Sindooram ID Vault';

// Handles GET requests: returns recent bookings for the picker.
function doGet(e) {
  return jsonResponse({ bookings: getRecentBookings() });
}

// Handles POST requests: saves one uploaded file and logs it.
function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  return jsonResponse(saveUpload(data));
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Only what's needed to identify a booking — no amount, status, remarks,
// or guest contact details are ever read or returned.
function getRecentBookings() {
  var sheet = SpreadsheetApp.openById(LEDGER_SHEET_ID).getSheetByName('Bookings');
  if (!sheet) return [];
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  var headers = values[0];
  var col = {};
  headers.forEach(function (h, i) { col[h] = i; });
  var rows = values.slice(1).filter(function (row) { return row[0] !== '' && row[0] !== null; });
  rows.sort(function (a, b) {
    return String(b[col['Check-in']]).localeCompare(String(a[col['Check-in']]));
  });
  return rows.slice(0, 60).map(function (row) {
    return {
      bookingNumber: cleanText(row[col['Booking Number']]),
      guestName: cleanText(row[col['Guest']]),
      source: cleanText(row[col['Source']]),
      checkIn: cleanText(row[col['Check-in']]),
      checkOut: cleanText(row[col['Check-out']])
    };
  });
}

// Ledger dates/text are stored with a leading apostrophe (forces plain
// text, see the Ledger's own script) — strip it back off, and format any
// stray real Date cell the same way.
function cleanText(val) {
  if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var s = String(val == null ? '' : val);
  return s.charAt(0) === "'" ? s.slice(1) : s;
}

// Saves one uploaded file into Drive under "Sindooram ID Vault/<Booking ID>/",
// and logs the upload as a row in the Ledger Sheet's "ID Vault Uploads" tab
// so there's a record of what's been collected, visible next to the bookings.
function saveUpload(data) {
  var bookingId = String(data.bookingId || '').trim() || 'Unlabeled';
  var folder = getOrCreateFolder(getVaultRootFolder(), bookingId);

  var bytes = Utilities.base64Decode(data.fileBase64);
  var blob = Utilities.newBlob(bytes, data.mimeType || 'application/octet-stream', data.fileName || 'upload');
  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmmss');
  blob.setName(stamp + '_' + (data.fileName || 'upload'));
  var file = folder.createFile(blob);

  logUpload(bookingId, data.bookingType || '', data.checkInDate || '', file.getName(), data.uploadedBy || '', file.getUrl());

  return { status: 'ok', fileUrl: file.getUrl(), folderUrl: folder.getUrl() };
}

function getVaultRootFolder() {
  return getOrCreateFolder(DriveApp.getRootFolder(), VAULT_FOLDER_NAME);
}

function getOrCreateFolder(parent, name) {
  var existing = parent.getFoldersByName(name);
  if (existing.hasNext()) return existing.next();
  return parent.createFolder(name);
}

function logUpload(bookingId, bookingType, checkInDate, fileName, uploadedBy, fileUrl) {
  var ss = SpreadsheetApp.openById(LEDGER_SHEET_ID);
  var sheet = ss.getSheetByName('ID Vault Uploads') || ss.insertSheet('ID Vault Uploads');
  if (sheet.getLastRow() === 0) {
    var headers = ['Booking ID', 'Booking Type', 'Check-in', 'File Name', 'Uploaded By', 'Uploaded At', 'File Link'];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  sheet.appendRow([bookingId, bookingType, checkInDate, fileName, uploadedBy, new Date(), fileUrl]);
  sheet.autoResizeColumns(1, 7);
}
