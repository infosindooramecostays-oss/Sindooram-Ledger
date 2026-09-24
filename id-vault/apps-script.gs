// ===== Sindooram ID Vault — Apps Script =====
// A separate, standalone script from the main Ledger one. It only reads
// booking basics (no money, status, or guest contact fields) from the
// Ledger Sheet, and saves uploaded documents to Drive. Deploy this as its
// own Web App with its own URL — do not reuse the Ledger's URL, since
// that one exposes full financial data.

// Paste your Ledger Google Sheet's ID here — open the Ledger Sheet, look
// at its URL: https://docs.google.com/spreadsheets/d/THIS_PART/edit
var LEDGER_SHEET_ID = '1aUJV7OZzxd9Bzwjr18SbsfYiDIp95nlAk1XhrXOcK34';

// The Drive folder everything gets saved under. Created automatically the
// first time anyone uploads a document — nothing to set up by hand.
var VAULT_FOLDER_NAME = 'Sindooram ID Vault';

// Handles GET requests: returns recent bookings for the picker, or (when
// called with ?bookingNumber=...) the documents already on file for one
// specific booking, so the app can show what's already been uploaded.
function doGet(e) {
  var bookingNumber = e && e.parameter && e.parameter.bookingNumber;
  if (bookingNumber) {
    return jsonResponse({ uploads: getUploadsForBooking(bookingNumber) });
  }
  return jsonResponse({ bookings: getRecentBookings() });
}

// Handles POST requests: saves one uploaded file and logs it, or (action
// 'deleteUpload') removes a file someone just uploaded by mistake.
function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  if (data.action === 'deleteUpload') {
    try { return jsonResponse(deleteUpload(data)); }
    catch (err) { return jsonResponse({ error: String(err.message || err) }); }
  }
  try { return jsonResponse(saveUpload(data)); }
  catch (err) { return jsonResponse({ error: String(err.message || err) }); }
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Only what the manager needs to identify a booking — no amount, status,
// source, or guest contact details are ever read or returned. Status is
// read only to filter to Confirmed bookings; it's never included in what
// gets sent back.
function getRecentBookings() {
  var sheet = SpreadsheetApp.openById(LEDGER_SHEET_ID).getSheetByName('Bookings');
  if (!sheet) return [];
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  var headers = values[0];
  var col = {};
  headers.forEach(function (h, i) { col[h] = i; });
  // Matches 'Confirmed', 'Confirmed-Advance', 'Confirmed-Paid', and any
  // other "Confirmed-..." variant the Ledger uses (or mistypes) — a plain
  // equality check against 'confirmed' was missing every Direct booking
  // the moment its status moved past the bare "Confirmed" stage, which is
  // exactly when a booking is most likely to need documents uploaded.
  var rows = values.slice(1).filter(function (row) {
    if (row[0] === '' || row[0] === null) return false;
    return String(row[col['Status']] || '').trim().toLowerCase().indexOf('confirmed') === 0;
  });
  // No date filtering here — the calendar needs past, present, and future
  // bookings. Hiding stays-already-over from the upload picker happens on
  // the frontend instead, so the calendar isn't starved by the same cutoff.
  // Compare cleaned yyyy-MM-dd strings, not the raw cell values — rows added
  // by an outside automation can store Check-in as a real Date object while
  // rows added through the app store it as force-text ('2026-09-30);
  // sorting the raw values mixes two incomparable formats and looks random.
  // Sort newest-first to pick the 200 most relevant rows, then flip to
  // chronological order for display.
  rows.sort(function (a, b) {
    return cleanText(b[col['Check-in']]).localeCompare(cleanText(a[col['Check-in']]));
  });
  rows = rows.slice(0, 200).reverse();
  return rows.map(function (row) {
    return {
      bookingNumber: cleanText(row[col['Booking Number']]),
      guestName: cleanText(row[col['Guest']]),
      checkIn: cleanText(row[col['Check-in']]),
      checkOut: cleanText(row[col['Check-out']]),
      remarks: cleanText(row[col['Remarks']]),
      guests: cleanText(row[col['Guests']])
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

// Saves one uploaded file into Drive under
// "Sindooram ID Vault/<Check-in date>/<Booking Number>/", and logs the
// upload as a row in the Ledger Sheet's "ID Vault Uploads" tab so there's
// a record of what's been collected, visible next to the bookings.
function saveUpload(data) {
  var bookingNumber = String(data.bookingNumber || '').trim() || 'Unlabeled';
  var checkInDate = String(data.checkInDate || '').trim() || 'Unknown date';
  var dateFolder = getOrCreateFolder(getVaultRootFolder(), checkInDate);
  var folder = getOrCreateFolder(dateFolder, bookingNumber);

  var bytes = Utilities.base64Decode(data.fileBase64);
  var blob = Utilities.newBlob(bytes, data.mimeType || 'application/octet-stream', data.fileName || 'upload');
  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmmss');
  blob.setName(stamp + '_' + (data.fileName || 'upload'));
  var file = folder.createFile(blob);

  logUpload(bookingNumber, data.guestName || '', checkInDate, data.comments || '', file.getName(), data.uploadedBy || '', file.getUrl());

  return { status: 'ok', fileUrl: file.getUrl(), folderUrl: folder.getUrl() };
}

// Undoes a mistaken upload from the same session: trashes the Drive file
// and removes its row from the "ID Vault Uploads" log, matched by the
// file's URL (the same value the app got back when it uploaded it).
function deleteUpload(data) {
  var fileUrl = String(data.fileUrl || '');
  var m = fileUrl.match(/[-\w]{25,}/);
  if (!m) throw new Error('Could not identify the file to remove.');
  DriveApp.getFileById(m[0]).setTrashed(true);
  removeUploadLogRow(fileUrl);
  return { status: 'removed' };
}

function removeUploadLogRow(fileUrl) {
  var ss = SpreadsheetApp.openById(LEDGER_SHEET_ID);
  var sheet = ss.getSheetByName('ID Vault Uploads');
  if (!sheet) return;
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return;
  var headers = values[0];
  var linkCol = headers.indexOf('File Link');
  if (linkCol === -1) return;
  for (var i = values.length - 1; i >= 1; i--) {
    if (values[i][linkCol] === fileUrl) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
}

// Reads the "ID Vault Uploads" log and returns just the rows for one
// booking, newest first — this is what lets the app show "documents on
// file" for a booking even after leaving and coming back, or reloading.
function getUploadsForBooking(bookingNumber) {
  var ss = SpreadsheetApp.openById(LEDGER_SHEET_ID);
  var sheet = ss.getSheetByName('ID Vault Uploads');
  if (!sheet) return [];
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  var headers = values[0];
  var col = {};
  headers.forEach(function (h, i) { col[h] = i; });
  var wanted = String(bookingNumber).trim();
  var rows = values.slice(1).filter(function (row) {
    return String(row[col['Booking Number']] || '').trim() === wanted;
  });
  return rows.map(function (row) {
    var uploadedAt = row[col['Uploaded At']];
    return {
      fileName: String(row[col['File Name']] || ''),
      uploadedBy: String(row[col['Uploaded By']] || ''),
      uploadedAt: uploadedAt instanceof Date ? Utilities.formatDate(uploadedAt, Session.getScriptTimeZone(), 'dd MMM, h:mm a') : String(uploadedAt || ''),
      fileUrl: String(row[col['File Link']] || '')
    };
  }).reverse();
}

function getVaultRootFolder() {
  return getOrCreateFolder(DriveApp.getRootFolder(), VAULT_FOLDER_NAME);
}

function getOrCreateFolder(parent, name) {
  var existing = parent.getFoldersByName(name);
  if (existing.hasNext()) return existing.next();
  return parent.createFolder(name);
}

function logUpload(bookingNumber, guestName, checkInDate, comments, fileName, uploadedBy, fileUrl) {
  var ss = SpreadsheetApp.openById(LEDGER_SHEET_ID);
  var sheet = ss.getSheetByName('ID Vault Uploads') || ss.insertSheet('ID Vault Uploads');
  if (sheet.getLastRow() === 0) {
    var headers = ['Booking Number', 'Guest Name', 'Check-in', 'Comments', 'File Name', 'Uploaded By', 'Uploaded At', 'File Link'];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  sheet.appendRow([bookingNumber, guestName, checkInDate, comments, fileName, uploadedBy, new Date(), fileUrl]);
  sheet.autoResizeColumns(1, 8);
}
