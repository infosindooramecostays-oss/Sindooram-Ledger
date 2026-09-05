function doGet(e) {
  return jsonResponse(readAll());
}

function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  writeTransactions(ss, data.transactions || []);
  writeBookings(ss, data.bookings || []);
  return jsonResponse(readAll());
}

function readAll() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return {
    transactions: readSheet(ss, 'Transactions', ['id','type','date','category','subcategory','description','amount','addedBy','recurrence','createdAt','updatedAt']),
    bookings: readSheet(ss, 'Bookings', ['id','bookingNumber','guestName','checkIn','checkOut','source','amount','status','addedBy','createdAt','updatedAt'])
  };
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function readSheet(ss, name, keys) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) return [];
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  return values.slice(1)
    .filter(function (row) { return row[0] !== '' && row[0] !== null; })
    .map(function (row) {
      var obj = {};
      keys.forEach(function (key, i) {
        var val = row[i];
        if (val instanceof Date) {
          val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
        }
        obj[key] = val;
      });
      return obj;
    });
}

function writeTransactions(ss, transactions) {
  var sheet = ss.getSheetByName('Transactions') || ss.insertSheet('Transactions');
  sheet.clear();
  var headers = ['ID', 'Type', 'Date', 'Category', 'Subcategory', 'Description', 'Amount', 'Added By', 'Recurrence', 'Created At', 'Updated At'];
  sheet.appendRow(headers);
  if (transactions.length > 0) {
    sheet.getRange(2, 3, transactions.length, 1).setNumberFormat('@');
    sheet.getRange(2, 10, transactions.length, 2).setNumberFormat('@');
  }
  transactions.forEach(function (t) {
    sheet.appendRow([t.id, t.type, t.date, t.category, t.subcategory || '', t.description || '', t.amount, t.addedBy, t.recurrence || 'One-off', t.createdAt, t.updatedAt]);
  });
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.autoResizeColumns(1, headers.length);
}

function writeBookings(ss, bookings) {
  var sheet = ss.getSheetByName('Bookings') || ss.insertSheet('Bookings');
  sheet.clear();
  var headers = ['ID', 'Booking Number', 'Guest', 'Check-in', 'Check-out', 'Source', 'Amount', 'Status', 'Added By', 'Created At', 'Updated At'];
  sheet.appendRow(headers);
  if (bookings.length > 0) {
    sheet.getRange(2, 4, bookings.length, 2).setNumberFormat('@');
    sheet.getRange(2, 10, bookings.length, 2).setNumberFormat('@');
  }
  bookings.forEach(function (b) {
    sheet.appendRow([b.id, b.bookingNumber, b.guestName, b.checkIn, b.checkOut, b.source, b.amount, b.status, b.addedBy, b.createdAt, b.updatedAt]);
  });
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.autoResizeColumns(1, headers.length);
}
