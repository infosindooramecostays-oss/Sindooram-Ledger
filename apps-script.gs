// Handles GET requests from the app: just returns everything in the Sheet.
function doGet(e) {
  return jsonResponse(readAll());
}

// Handles POST requests from the app: overwrites both sheets with the data
// sent, then returns the freshly saved copy back to the app.
function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  writeTransactions(ss, data.transactions || []);
  writeBookings(ss, data.bookings || []);
  return jsonResponse(readAll());
}

// Reads both sheets and returns them as plain JS objects/arrays.
function readAll() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return {
    transactions: readSheet(ss, 'Transactions', ['id','type','date','category','subcategory','description','amount','addedBy','recurrence','createdAt','updatedAt']),
    bookings: readSheet(ss, 'Bookings', ['id','bookingNumber','guestName','checkIn','checkOut','source','amount','status','addedBy','createdAt','updatedAt'])
  };
}

// Wraps an object as the JSON response the app expects.
function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Strips a leading apostrophe (our own "force text" marker — see writeRow)
// and converts any stray Date-typed cell back to a plain yyyy-mm-dd string.
function cleanCellValue(val) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
  }
  if (typeof val === 'string' && val.charAt(0) === "'") {
    return val.slice(1);
  }
  return val;
}

// Reads one sheet's rows into an array of objects keyed by "keys", skipping
// blank rows.
function readSheet(ss, name, keys) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) return [];
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  return values.slice(1)
    .filter(function (row) { return row[0] !== '' && row[0] !== null; })
    .map(function (row) {
      var obj = {};
      keys.forEach(function (key, i) { obj[key] = cleanCellValue(row[i]); });
      return obj;
    });
}

// A leading apostrophe is Sheets' own "treat this as literal text" marker —
// it works the same whether typed by hand or set via the API, and (unlike
// setNumberFormat) doesn't depend on a format change landing before the
// value does. Sheets strips it on input; cleanCellValue() strips it too as
// a fallback, in case a value ever comes back with it still attached.
function forceText(value) {
  return "'" + value;
}

// Replaces the whole Transactions sheet with the given rows.
function writeTransactions(ss, transactions) {
  var sheet = ss.getSheetByName('Transactions') || ss.insertSheet('Transactions');
  sheet.clear();
  var headers = ['ID', 'Type', 'Date', 'Category', 'Subcategory', 'Description', 'Amount', 'Added By', 'Recurrence', 'Created At', 'Updated At'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (transactions.length > 0) {
    var rows = transactions.map(function (t) {
      return [t.id, t.type, forceText(t.date), t.category, t.subcategory || '', t.description || '', t.amount, t.addedBy, t.recurrence || 'One-off', forceText(t.createdAt), forceText(t.updatedAt)];
    });
    sheet.getRange(2, 3, rows.length, 1).setNumberFormat('@');
    sheet.getRange(2, 10, rows.length, 2).setNumberFormat('@');
    SpreadsheetApp.flush();
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.autoResizeColumns(1, headers.length);
}

// Replaces the whole Bookings sheet with the given rows.
function writeBookings(ss, bookings) {
  var sheet = ss.getSheetByName('Bookings') || ss.insertSheet('Bookings');
  sheet.clear();
  var headers = ['ID', 'Booking Number', 'Guest', 'Check-in', 'Check-out', 'Source', 'Amount', 'Status', 'Added By', 'Created At', 'Updated At'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (bookings.length > 0) {
    var rows = bookings.map(function (b) {
      return [b.id, b.bookingNumber, b.guestName, forceText(b.checkIn), forceText(b.checkOut), b.source, b.amount, b.status, b.addedBy, forceText(b.createdAt), forceText(b.updatedAt)];
    });
    sheet.getRange(2, 4, rows.length, 2).setNumberFormat('@');
    sheet.getRange(2, 10, rows.length, 2).setNumberFormat('@');
    SpreadsheetApp.flush();
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.autoResizeColumns(1, headers.length);
}
