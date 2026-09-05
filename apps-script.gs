// Handles GET requests from the app: just returns everything in the Sheet.
function doGet(e) {
  return jsonResponse(readAll());
}

// Handles POST requests from the app: overwrites both sheets with the data
// sent, then returns the freshly saved copy back to the app.
function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  if (data.action === 'sendReport') {
    sendMonthlyReport();
    return jsonResponse({ status: 'sent' });
  }
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
    bookings: readSheet(ss, 'Bookings', ['id','bookingNumber','guestName','guestEmail','guestPhone','checkIn','checkOut','source','amount','status','remarks','addedBy','createdAt','updatedAt'])
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
  var headers = ['ID', 'Booking Number', 'Guest', 'Guest Email', 'Guest Phone', 'Check-in', 'Check-out', 'Source', 'Amount', 'Status', 'Remarks', 'Added By', 'Created At', 'Updated At'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (bookings.length > 0) {
    var rows = bookings.map(function (b) {
      return [b.id, b.bookingNumber, b.guestName, b.guestEmail || '', b.guestPhone || '', forceText(b.checkIn), forceText(b.checkOut), b.source, b.amount, b.status, b.remarks || '', b.addedBy, forceText(b.createdAt), forceText(b.updatedAt)];
    });
    sheet.getRange(2, 6, rows.length, 2).setNumberFormat('@');
    sheet.getRange(2, 13, rows.length, 2).setNumberFormat('@');
    SpreadsheetApp.flush();
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.autoResizeColumns(1, headers.length);
}

// Where the monthly Profit & Loss email gets sent. Change this any time.
var REPORT_RECIPIENT = 'chinnoos.pr@gmail.com';

// Emails a P&L summary for the calendar month that just ended. Safe to
// call any time (manually, or from the scheduled trigger below) — it only
// reads data, never changes the Sheet.
function sendMonthlyReport() {
  var now = new Date();
  var periodStart = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  var periodEnd = new Date(now.getFullYear(), now.getMonth(), 0, 23, 59, 59);
  var prevStart = new Date(periodStart.getFullYear(), periodStart.getMonth() - 1, 1);
  var prevEnd = new Date(periodStart.getFullYear(), periodStart.getMonth(), 0, 23, 59, 59);
  var data = readAll();
  var current = summarizePeriod(data, periodStart, periodEnd);
  var previous = summarizePeriod(data, prevStart, prevEnd);
  var monthLabel = Utilities.formatDate(periodStart, Session.getScriptTimeZone(), 'MMMM yyyy');
  MailApp.sendEmail({
    to: REPORT_RECIPIENT,
    subject: 'Sindooram Ledger — P&L for ' + monthLabel,
    htmlBody: buildReportHtml(monthLabel, current, previous)
  });
}

// Totals income/expenses/bookings for transactions and bookings whose date
// falls between start and end (inclusive).
function summarizePeriod(data, start, end) {
  var income = 0, expenses = 0, byCategory = {};
  data.transactions.forEach(function (t) {
    var d = new Date(t.date + 'T00:00:00');
    if (d < start || d > end) return;
    var amt = Number(t.amount) || 0;
    if (t.type === 'income') { income += amt; }
    else { expenses += amt; byCategory[t.category] = (byCategory[t.category] || 0) + amt; }
  });
  var bookingsCount = 0, nights = 0, bookingsRevenue = 0;
  data.bookings.forEach(function (b) {
    var d = new Date(b.checkIn + 'T00:00:00');
    if (d < start || d > end) return;
    bookingsCount += 1;
    bookingsRevenue += Number(b.amount) || 0;
    nights += Math.round((new Date(b.checkOut + 'T00:00:00') - d) / 86400000);
  });
  return { income: income, expenses: expenses, net: income - expenses, byCategory: byCategory, bookingsCount: bookingsCount, nights: nights, bookingsRevenue: bookingsRevenue };
}

function formatMoney(n) {
  return '₹' + Math.round(n).toLocaleString('en-IN');
}

function pctChange(curr, prev) {
  if (!prev) return null;
  return ((curr - prev) / Math.abs(prev)) * 100;
}

function deltaHtml(curr, prev) {
  var pct = pctChange(curr, prev);
  if (pct === null) return '';
  var arrow = pct >= 0 ? '▲' : '▼';
  return ' <span style="font-size:12px;color:#898781;">(' + arrow + ' ' + Math.abs(Math.round(pct)) + '% vs last month)</span>';
}

// Builds the HTML email body: a header banner, a big net profit/loss
// figure, income vs expenses, an expense-by-category table, and a
// bookings summary — the same numbers the dashboard already tracks.
function buildReportHtml(monthLabel, curr, prev) {
  var netColor = curr.net >= 0 ? '#0ca30c' : '#d03b3b';
  var netLabel = curr.net >= 0 ? 'Net Profit' : 'Net Loss';
  var categories = Object.keys(curr.byCategory).sort(function (a, b) { return curr.byCategory[b] - curr.byCategory[a]; });
  var categoryRows = categories.map(function (cat) {
    return '<tr><td style="padding:6px 0;color:#52514e;">' + cat + '</td><td style="padding:6px 0;text-align:right;font-weight:600;">' + formatMoney(curr.byCategory[cat]) + '</td></tr>';
  }).join('');
  return '<div style="font-family:system-ui,-apple-system,sans-serif;max-width:480px;margin:0 auto;">' +
    '<div style="background:linear-gradient(135deg,#4c3fd7,#8b3fea,#ec4899);padding:24px;border-radius:16px;color:#fff;text-align:center;">' +
    '<div style="font-size:14px;opacity:.85;">Sindooram Ledger</div>' +
    '<div style="font-size:18px;font-weight:700;margin-top:4px;">' + monthLabel + ' — Profit &amp; Loss</div></div>' +
    '<div style="background:#fff;border:1px solid #eee;border-radius:16px;padding:20px;margin-top:12px;text-align:center;">' +
    '<div style="font-size:12px;color:#898781;text-transform:uppercase;">' + netLabel + '</div>' +
    '<div style="font-size:32px;font-weight:800;color:' + netColor + ';margin-top:4px;">' + formatMoney(Math.abs(curr.net)) + '</div>' +
    '<div>' + deltaHtml(curr.net, prev.net) + '</div></div>' +
    '<div style="background:#fff;border:1px solid #eee;border-radius:16px;padding:20px;margin-top:12px;">' +
    '<table style="width:100%;border-collapse:collapse;">' +
    '<tr><td style="padding:6px 0;color:#52514e;">Total Income</td><td style="padding:6px 0;text-align:right;font-weight:700;color:#006300;">' + formatMoney(curr.income) + deltaHtml(curr.income, prev.income) + '</td></tr>' +
    '<tr><td style="padding:6px 0;color:#52514e;">Total Expenses</td><td style="padding:6px 0;text-align:right;font-weight:700;">' + formatMoney(curr.expenses) + deltaHtml(curr.expenses, prev.expenses) + '</td></tr>' +
    '</table></div>' +
    '<div style="background:#fff;border:1px solid #eee;border-radius:16px;padding:20px;margin-top:12px;">' +
    '<div style="font-size:13px;font-weight:700;margin-bottom:8px;">Expenses by category</div>' +
    '<table style="width:100%;border-collapse:collapse;">' + (categoryRows || '<tr><td style="color:#898781;">No expenses recorded.</td></tr>') + '</table></div>' +
    '<div style="background:#fff;border:1px solid #eee;border-radius:16px;padding:20px;margin-top:12px;">' +
    '<div style="font-size:13px;font-weight:700;margin-bottom:8px;">Bookings</div>' +
    '<div style="color:#52514e;font-size:14px;">' + curr.bookingsCount + ' booking(s) · ' + curr.nights + ' night(s) · ' + formatMoney(curr.bookingsRevenue) + ' revenue' + deltaHtml(curr.bookingsRevenue, prev.bookingsRevenue) + '</div></div>' +
    '<div style="text-align:center;margin-top:16px;font-size:12px;color:#898781;">Generated automatically from your Sindooram Ledger Sheet.</div></div>';
}

// Run this once yourself (select it above, then click Run) to schedule
// sendMonthlyReport() for 7am on the 1st of every month. Re-running it is
// safe — it clears any previous schedule for this function first so you
// never end up with duplicate emails.
function createMonthlyReportTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sendMonthlyReport') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendMonthlyReport').timeBased().onMonthDay(1).atHour(7).create();
}
