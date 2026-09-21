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
  if (data.action === 'previewReceipt') {
    try { return jsonResponse(handlePreviewReceipt(data)); }
    catch (err) { return jsonResponse({ error: String(err.message || err) }); }
  }
  if (data.action === 'sendReceipt') {
    try { return jsonResponse(handleSendReceipt(data)); }
    catch (err) { return jsonResponse({ error: String(err.message || err) }); }
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  writeTransactions(ss, data.transactions || []);
  writeBookings(ss, data.bookings || []);
  return jsonResponse(readAll());
}

// Reads both sheets and returns them as plain JS objects/arrays.
function readAll() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  backfillMissingBookingIds(ss);
  return {
    transactions: readSheet(ss, 'Transactions', ['id','type','date','category','subcategory','description','amount','addedBy','recurrence','createdAt','updatedAt','endsOn']),
    bookings: readSheet(ss, 'Bookings', ['id','bookingNumber','guestName','guestEmail','guestPhone','checkIn','checkOut','source','amount','status','remarks','addedBy','createdAt','updatedAt','guests'])
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

// Assigns a temp ID to any Bookings row that has data but no ID — e.g. a
// row an outside automation (Airbnb email → Sheet) appended directly,
// which has no concept of our ID format. Without an ID the row is invisible
// to the app (readSheet skips it); once it has any ID it shows up and can
// be edited normally, including filling in the real Booking Number by hand.
function backfillMissingBookingIds(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Bookings');
  if (!sheet) return;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  var range = sheet.getRange(2, 1, lastRow - 1, 15);
  var values = range.getValues();
  var changed = false;
  values.forEach(function (row, i) {
    var hasData = row.slice(1).some(function (v) { return v !== '' && v !== null; });
    if ((row[0] === '' || row[0] === null) && hasData) {
      row[0] = 'temp_' + new Date().getTime().toString(36) + Math.random().toString(36).slice(2, 7);
      values[i] = row;
      changed = true;
    }
  });
  if (changed) range.setValues(values);
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
  var headers = ['ID', 'Type', 'Date', 'Category', 'Subcategory', 'Description', 'Amount', 'Added By', 'Recurrence', 'Created At', 'Updated At', 'Ends On'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (transactions.length > 0) {
    var rows = transactions.map(function (t) {
      return [t.id, t.type, forceText(t.date), t.category, t.subcategory || '', t.description || '', t.amount, t.addedBy, t.recurrence || 'One-off', forceText(t.createdAt), forceText(t.updatedAt), forceText(t.endsOn || '')];
    });
    sheet.getRange(2, 3, rows.length, 1).setNumberFormat('@');
    sheet.getRange(2, 10, rows.length, 3).setNumberFormat('@');
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
  var headers = ['ID', 'Booking Number', 'Guest', 'Guest Email', 'Guest Phone', 'Check-in', 'Check-out', 'Source', 'Amount', 'Status', 'Remarks', 'Added By', 'Created At', 'Updated At', 'Guests'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (bookings.length > 0) {
    var rows = bookings.map(function (b) {
      return [b.id, b.bookingNumber, b.guestName, b.guestEmail || '', b.guestPhone || '', forceText(b.checkIn), forceText(b.checkOut), b.source, b.amount, b.status, b.remarks || '', b.addedBy, forceText(b.createdAt), forceText(b.updatedAt), b.guests || ''];
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

// Categories that reduce Net — matches the app's own NET_REVENUE_CATEGORIES.
// Maintenance, Property Upgrades, and Digital Promotions still count toward
// `expenses`/byCategory (for audit) but aren't subtracted from net.
var NET_REVENUE_CATEGORIES = ['Operational Expenses'];

// Totals income/expenses/bookings for transactions and bookings whose date
// falls between start and end (inclusive).
// TODO: unlike the app's own dashboard, this does NOT treat Weekly/Monthly/
// Yearly transactions as recurring — a fixed expense entered once will only
// show in the one month it was dated, not in every later month's email.
// Deliberately left as-is for now; revisit if the email should match.
function summarizePeriod(data, start, end) {
  var income = 0, expenses = 0, netExpenses = 0, byCategory = {};
  data.transactions.forEach(function (t) {
    var d = new Date(t.date + 'T00:00:00');
    if (d < start || d > end) return;
    var amt = Number(t.amount) || 0;
    if (t.type === 'income') { income += amt; }
    else {
      expenses += amt;
      byCategory[t.category] = (byCategory[t.category] || 0) + amt;
      if (NET_REVENUE_CATEGORIES.indexOf(t.category) !== -1) netExpenses += amt;
    }
  });
  var bookingsCount = 0, nights = 0, bookingsRevenue = 0;
  data.bookings.forEach(function (b) {
    var d = new Date(b.checkIn + 'T00:00:00');
    if (d < start || d > end) return;
    bookingsCount += 1;
    bookingsRevenue += Number(b.amount) || 0;
    nights += Math.round((new Date(b.checkOut + 'T00:00:00') - d) / 86400000);
  });
  return { income: income, expenses: expenses, net: income - netExpenses, byCategory: byCategory, bookingsCount: bookingsCount, nights: nights, bookingsRevenue: bookingsRevenue };
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

// ===== Payment receipts (Confirmed-Advance / Confirmed-Paid, Direct bookings only) =====

// Logo embedded as base64 directly in the script — avoids needing the
// script.external_request permission (UrlFetchApp) just to fetch an
// image; the bytes live in the script itself instead.
var LOGO_BASE64 = '/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAYEBAUEBAYFBQUGBgYHCQ4JCQgICRINDQoOFRIWFhUSFBQXGiEcFxgfGRQUHScdHyIjJSUlFhwpLCgkKyEkJST/2wBDAQYGBgkICREJCREkGBQYJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCQkJCT/wAARCAErAcIDASIAAhEBAxEB/8QAHQABAAIDAQEBAQAAAAAAAAAAAAYHBAUIAwIBCf/EAFMQAAEDAwEFAwcFCwcJCQAAAAABAgMEBREGBxIhMUETUWEIFCJxgZGhFTJCdLEWFyM1NlJUYrKz0jM3VXKVlsElQ3WSlMLT4vEkNFNjc4KDotH/xAAaAQEAAgMBAAAAAAAAAAAAAAAABAUBAgMG/8QANREBAAIBAgMECAYCAgMAAAAAAAECAwQREiExBRNBURQiYXGBkaHwFTKxwdHhUvEjQjNygv/aAAwDAQACEQMRAD8A6pAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAI9qDaHpTSzlZeL9QUkqf5pZN6T/Ubl3wMTMR1bVpa07VjdIQVdP5SOgIXq1lXcJkT6UdE/HxwZlu8oHZ7XuRrr26kcvLzqnkjT34x8TTvaebvOjzxG80n5LFBh2u9W290/nNsr6Wug/8AEp5WyN96KZh0R5iYnaQABgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAANbqLUlr0paZrreKuOlpIU4vdzcvRrU5ucvREPS+Xqh07aKu7XKZIaSkjWSV69EToneqrhETqqocypU3XbtqSqu+oK75I0jZ8yyvc5EZSxr9FF5LK5Obl5J4YReWTJw8o6pmk0ve73vO1Y6z+0e1sr1tE19tku01p0RTVlvtLPRe5jkjcqcfSmlT5uU+g1c+s2ln8ney2KCSXUtc+93VY+3bbqOVYEkai+lhy+k9fdyLRp7ZQ6LjtlRZX1DLG2FKZltoYGvZI+RyKk7n/OVcJjOeOTA1NWT1VVO2Golc/zqOmp+3oVjWle926r45eGeCO9+SJmmKRM252WWHNa0xTDHDT6z75+/eh9VoXZ9QNrqZtts8b3ywSUqyo+SSONd1XsfzVFTC8+PE9I9lmgLw1VWyxIypqZnNqbfUujZSU7eTnrndz1wqZ4+Bs/O6a2USVC19ZbqF08kFPDQxsWSRGLh0kj3c1VeOPcfckS09we2oSCd7amGmqHdlux1kMqb0b3xphN9qoi+JBjNO+/35/fKOXwTJrbh2i0/X2R7PHbxnn7N0CvGwbU+lJfl/Z5fZayPHaRsik7KoVnNMKnoypjouM9ykp2YbforhK3T+tv8nXdj+xbUyR9nHK7ON16f5t/rwi+HIndNfZIrG6nWpuE9TLMtIyoht/Zebuc3DXNY5MK1vPjkh21TZbZNR0FrpKu9U8Wr5IFgpqqdGxuur2M4te1OC9+U+bnqnAsadOLH8kC2WMv/AB6qPdbbn+nRcAKU2BbSK2pkm0JqVZWXW37zKZZ/nvYzg6J3e5nTvb6uN1kql4vG8KvUYLYbzS3+wAG7gAAAAAAAAAAAAAAAAAAABkZAAZGQAGRkAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAOgFA+UVf62/wB+suz20O3pamSOWdqLwdI5d2Jq+CJvPX2dxva7VGkNF2GTZzY7vBRXWKJKds9RRrJTvqVxlsznJuZevBc8t7mmCMaFZ91HlI365TfhG251QsarxRNzdhb8FU2eo6fX161jer1cK2ssWj7EssjUw1qVUUSZVGsX5++qL6TkVEReGSJEzO9vOdvgvZpWIphnpEb++Z+E7+5K9F7Tb3qK+wWBdKyt8zhZHdq5JUZDS1KMy+NiYXeVHcMIuefRMrtNW2yTtap75qpkCsSo89qaj8DSPauWNjjROK5RM+C9SuajTV41lqmmSuifZqiqoWXew3SlRWpTSei99NKjcNfhXr6Soj1RM5XKlz11ws77hRWO5S0z66qa6ogpntysnZ4VzkTwVUFsc5aTWyLa1cOStqR4c4j9fvy5x518+qqKF8sc29QLLJ20tLU29aqDtV5yROTlnng+6SnqKqqik7OeZz5XViLVKkE1xnYmERicdxrEXKIv/Tf6tq9PaQtNVJd7862pcK1s6yuXMjnbzfRa1vFWoiIi9ycVNxT26zQw1F5WeKekkl+Umyucjo4l3OMjHJ0VEzkiRo7cW0z9/fs+m6TOurFOKK9fZ7vj9Z6eezyjluln0rPV0lPXXetbGs0NHWytjlXl+DV+MZTjxUrSTa3VXmrtt5r46fS9qtMskd0StgSaq85ym7TQoqbyK9npZRqLjPcmZdtDs9Jr/StHWR3iRNOxNdcqlKNyo6thbGrmMa5OSKvHkvTqQO30Oobhpu02GtnbpW53pk1RauxjRFp2xIiMpnudmRVcx6u3ldvJjHehNneu1Y6I+CtLRN77bzPy6/tz6cvojW1a5W+quVp2q6LmfllU2nq0dE6NzKhiZYr2rx9JmWr0VMHR+nb3TajsVBeKRcwVsDJ2J+bvJnC+KLlPYU821asumxjVli1pTVDrla0dLT1M2HrMxiJI1yPT5+FRyZ544LxN35NF1fcNm7aZ797zCsmgancxcPT9tRjmYv72+rrFsG8f9J28+U845rWMS7XSkstuqLjXSpFTU7FfI9eif4qvJE8TLKr8oK4ywWC3ULHKjKqpVz8dUY3KJ71RfYWGmxd7lrj83n9Xn7jDbL5IlqLbpqC4VD22dsVspUXDFViSSuTvcq5RPUie1TTUm1/WdPLvpeu3TmrJYY3IvuRF+JJNiOibVf8Az263SnZVpSyNihgkTLEdjeVyp15oiZ4cy3LnovT13on0lVaKJY3JhFZE1jmeLVRMovqLfLn0uC/dd3vt1UGn0ut1VO/73bfpHP8AZEdnW1+HVNW21XaCKjuD/wCRdGq9lP4Jni13hxz8DI1rqq9WTUsNNQTMSKSFipFKxFa9yucnPmmcInMpi96VvOl9R1NNR01dM6inR0FRHA9d7GHNciomM8i7tV6cn1hYrfdqWLdr0ga9YXejvtciKrePJUXOM+JT9u6XhxRk0s9efJJ0ep1GXHfHf89fq2dq1xS1thmuVRE6Camd2U1P1R/REz3+PLj3Gsj1rdqm5U9K2mpoUlka3cVFc5EVevHguDWWmz1TrbNUTPhiuD1YiwTStRXSMVzUevHnuqi+tDeaY0jLbKlbldJGNfEiua3eyjV6ucvvPL5La/Lkx0rvEct56fP2+xa4r3tWN0wB+Ncjmo5qoqLxRU6n4+RkTVc97WNTq5cIeiSn0D8Y9r2o5jkc1eSouUU/QABhXC92y0oi3C4UlJlMp28rWZ9WVMxEzyhibREbyzSk9rW1C60t6msNlqnUUVKiNnnj4SPeqZVEX6KIipy45yXDb7vb7rGr6Cupqtqc1gla/HrwpzNrdrZdol2a5MtdcVaqd6byIWfZeGtss8cdIU3bOovTBWMU7cU7bw29q0htKvlGyvpprmkMybzHTV6xq9F6oiuzgzPvc7UP0iq/tT/mOgGNaxqNaiI1EwiJ0Q/VXCZUxPad9+VY+TNexcW3rXtv73P33udqH6RVf2ov8Q+9ztQ/Sar+1F/iLlq9daYoZ+wqL/bY5E4K1Z2qqevC8Da0VfSXGBJ6KqgqYl5SQvR7fehme0M0RvNI+TWvZWmtO1ckzP8A7KF+9ztQ/Sar+1F/iH3udqH6RVf2ov8AEdAg0/E8n+MfJ0/BsP8Alb5/0oD73u1N3Baqrx43T/mJVozZLeKW4wXHU15lqEhckjKSOoke1zk5b7lXiic8InHvLEuuo7NZPxldKOjVeKNmla1y+zmfFr1TY70/ct12oap/5kUzVd7s5F9bnvTlWIjziDF2dpseSN7TM+Uz+zaAArluAGPDcaOondTw1dPJMzi6NkjVcnrRFyZ2YmYhkAGBcNQWi0uRlwudFSOXk2aZrF9yqIiZ5QWtFY3mWeDHoblRXKLtqGrp6qP8+GRHp70PtKymV+4k8SuzjG+mcjaehFomN4l6g+XyMjRFe9rEVUaiuXGVVcIh9GGQH4r2o5GqqZXkmeJ+gAMoAAAAAAAAAAAAAADmrZjW1di2r6783gSerYlQrYnIvpIlU1XcE4/NXJYG1y/1dx2eXmn8zdTQ1bqWjY92cvWWZqOTuXhnkQTU+/oHyiI67f8ANqW9N4SrwRvbMWNXZ/VkRrj4+S77R7Fpbneb5LXpR11JG2hc3CUPYVWJEd1c9Xc1Xoicyu4bbzEW5Rvy+/evMuHjzY9Rxcp4eXtTDWGudRt1Klk0tBO+C1+g2np2bz6t8bWrI9/XsY0cxu63DpHruovA2erNqkWzjTlOt7rKO9aknjV9PDTU606bjuKOe1yudG1OGc8VVMInPGe6OV21K8MgrI6StrrDB8nVD4u0axrZZO1VG5TeVFdGuM9UyVO3YfLqXaLqW03HVFVUz0VNDVOrZIUWSeWVFVN5M4RqY6eGMEi03j8vi54aYL7d7yiIiZ85+PvaK2SJtCfLqDUkstbXSvqInSuic+GnRsWWQMjavo5y56S/Nj3fSyqnvojazPsx1BV2Vyy3LTPbOY6lcu8+lz85I1Xg7dXKL9F+MpjJ7bCNltm10691F5lrE8yVtMxlNKsWd9Hbyq5OKpw5clzxyanZ5skg1tqvUFjfdJaSK077WSsiRzpFSRWNyi8ET0cr8CPHHymvWVrecG+THkn1axHLblG/TZceqtXahqaagrNF18Fws9cnaU0dtpMT0scLN6RXOcqxrjgnZOa1XZ3UwqZNdqHU0upaHQd5qGQsuFHqWmp5HQZ7KVsrFVsjM8dx7Fa5EXimVReKGP5Pejqq1WZ2pI9R1FPTvnnirKB0bXU8iRPVu/leLXejneT1cUMOB75dGbLoHMRrKvUqTwoqYXsUlldH7N1zfYd95mN58VZwUreaV58M7b+PSd9/l+6xdWaluPyFcaSe2dkzzSsSpkVHbrY2xO4p6+HPv4Z6Q7yUo3t0bdXuzuuuCInrSJmTQ7Ubve9LSa0t9bfZLlS1lHS0dJvIjXRdrLI/s3Y+c5saSelzVrm5LI2B2F9i2Z2ztWqyWuV9c9qpy7RfR/8AojTTDW0X2tO+zh3U4dJbed+KY2+W6wyEbXNIT6r0z/2KNZK6if28TE5yJhUc1PFU5eKITcj9419pmwVzqC53aGlqWtRyxva5VwvJeCFnp5vW8WxxvMKTVVx2xTTLO0Tyc8aO1xd9DV0z6JGPjkVGz0s6KjXKnxa5OPH3oWra/KAsdQ1EuVurqN/VY8Ss+1F+BKajSujdeUsd2dQUtayobllVEixuemcfOTCry6kTvXk/2iojc+z3Gqo5cZRk2JY1Xu6OT3qWt9RpM8/81Zrb7++ikxaXXaau2ntFq+Eff8p3p/Wdg1QipabnDUSNTLouLZET+quFN10OTrzZrxoe+pTVLnUtdT7ssU0D+aLyc13dwX7FOj9Aakk1VpShuc6NSoeixzbqYRXtXdVU9eM+0i6zRRhrGTHO9ZTezu0rai04stdrw5pq2N+6eb0W/jF3T/zlOqr5+JLh9Wl/ZU5Wq/ynm/0g798p1TfPxLcPq0v7KkrtTri+/JA7E6ZvvzUToLbFV6Zs0ttuMT66KKBfMnZ9JjscI3L+Z4805eqJ3Cv1HrSslq6hK+5SZyqRRveyNO5Gt4NQwtO2z5avNttiydmlXPHCr/zUVcKvuOsrTaaKyW+G32+BkFPC3daxqfFe9V6qSdVlxaS/FSu9rImhwZtfj4L32pVypZ9RXrTFaktvramjljd6USqu6vg5i8F9SodKaC1hDrTT8VxaxsU7V7KoiRcoyROePBUVFTwUgflAWShZb6G8shRla6dKd8jeG+zdcqb3eqKnBfFTE8neaTt75Bvfg92F+7+tl6Z9xw1XBqdN6REbTCToe80es9Fm29Z/2mO1XXrtGWiOKiVq3Osy2HeTKRNT50ip1xlERO9fAoa2WPUGt7lM6kp6m5VSrvTTPdndz1c93BPUSfbjVyVGu5IXqu5T00TGJ4Lly/FS2tktsprdoS1up2t3qmPziV6c3Pdzz6uCewzjvGj01clY3tYyY7doay2K1tq1UvV6H1toFW3tlO+n7D0lqaSVH9n/AF0T6PflMEfqrnLedRuuU7Wslqqtsr2t5IquTODraSNkrHRyNa9jkVrmuTKKi80U5SvdvgtWs6ygpv5CnuCxxp3NSTgns5ew7aDVzqJtxx60R19iN2loI0sV7u08Mz0nzdXnPm1TaZV3241FntlS+G1U7ljesa4Wpci8VVfzUXknXmpc2u6yWg0beqmBVSVlJJuqnNFwqZ9mTnnZlaqa8a5tVJWI18PaOkcx3FHqxquRF7+KIQuzMVIrbPeN+FZdsZ8k2ppsc7cXX9HpadlmrrvQtrKW0qyB6bzO2kbEr070aq59+DCoLhqHZ5fF7NJ7fWRKnawSJ6Mre5ycnNXvT2KdWIVbt/tVLLpyjuatalVT1LYWvxxcx6LlvvRFO2n7Stmyd1krG0o2q7Hrp8U5sNp4q8010Vq2l1lYo7nTsWJ28sc0KrlYpE5pnqnFFRe5SMbXdokuk6OK2Wx6NudW1XdpjPYR8t5E/OVconqVSM+TxVypUXuk4rErIZfU7Lk+z7CHbXauWp19d1kz+BVkTEXo1GJj7VX2nPDoqemWxz0jm66jtHJ6BXLHK1uX6/w1lk0tqLWlVNLQUk9dJvZmqJX4ajl/Oe5efhzPS/aG1JpJGVVyt8tPGjkRtTE9HNa7p6TV9FfXg6P0TaqWzaVtlJSNajEp2PVyfTc5EVzl8VVTZ3G3011oKihq42y09QxY5GuTmim1u17Rk2iscLWnYNLYuKbTxz8labHdpFTfldYbzMstbGxX09Q750zE5td3uTnnqnq4+m3C819hjsNfbaqSmqY6mTD2LzTc5KnJU8FKm0RLJbdeWnsXq5Y69sOU+k1XKxfeilmeUP8Aiyy/WJP2DfJp6U1lOGOVvD5tMWryZOz8nFPrV25/GET1RtmvOorJBbYIkt73t3auWBy5m7kZ1ai9U59M457jYFaKmnv9yrKijnhb5o1jHyRK1Hbz0Xgqpx+abDYJp22VFtq71PSsmro6lYY5Hpns2o1q+inRVzz5lwcjjq9TjxRbTYq7ecu+g0eXNamrzX3nwhXO1/aHNpWkitdrkRlyq2q5ZeawR8t5P1lXKJ3YVSl7FpPUOtKmaS3UktY9FzNUSvw1FXve5eK+HFTZbXKySq1/d1fn8CrImJ3NaxP8VX3l/wChbXS2jSVqpqVrUZ5syRzk+m5zUc53tVTtGSNFp62pHrWR+6ntHV3rknatP9KFl01rXZjUxXlIH0zGOTemgkSSJyZ+bIidF5cU9XE1OmKhazXVrqXNRrprnFKqJyRXSouPidT1lHBX0k1JUxtlgmYscjHJwc1eCoctaap20uvbZTsdvshukcbXd6NlREX4HXSar0il5tHrRDhr9D6LkxxS0zWZ6T4TyXvticrdn1zc1VRUWFUVFwqL2rSs7Ztsu1JpOptk+9LdGtbHS1q8V3V4Kr883InJevXlxsvbH/N5dPXF+9aUjs30zBqvVlLb6t2KZqOnmanN7W49H2qqJ6snDQY8VtNa2WN4id/0Su08uausrTDO02jb6y0VRUXGof8AKFRLWSOkdnzmRz13neD16+0nVq2yXm3aQqLW+Z81yR7WU1XJ6To4lRd5VVebkVERM9/HkW7tDoKVNn14p0p4khho3LHGjURrFamW4TpjCFJ7HrXTXTXdIyrjbLHBHJUNY5Mormom7n1KufYSKajFqcNsl6cqol9Lm0morix5Od/H3o7cG3+NzbjcEurHSLltTUdo3eXnwcpcuxTXtbf46myXWd1RU0rElhneuXPjzhUcvVUVU49y+BMdoNJDWaJvcc7Ee1KSR6IvRzUVUX1oqIUvsLVfu6Txo5ftacrZa6vTXtNdpq70wX0OspWt94t1dEgAoHqQAAAAAAAAAAVN5RuiF1Ho75YpmKtbZt6bDU4vgX+UTh3YR3/tUpm2awhvViutBedT3iB13Vq3ClpbKyoZIrURGyI5HoqOVGtyuEVVTjnmdfPY2RjmPajmuTCtcmUVO5TlXars4qtluoGXq0yVrdP102FbSTvgfDlcrAr28kXC7qr3Yxw4xM9JieOPivOzM9b17i/WOcdPlz397Jq9cQ1tVaaubWmpfOLOqrRys0w1rmIrd1UVUk9JqoiZReCnjLq6GS9Vd7j15q+mr61kcc8tNpxsfaNYio1FTfxwypsrNZ11BbYblbLTrqopJ0VWP+6+BqrhcLlHORUXPRUM37jbj/QOu/74038ZptM/cpHFjrO3w/6fx5vLZXrzRuzOG5xtqNTXJa+SORXOs6xbitRU6PXOckbprtZLVe7ndrJrDWNqluMz5JUgsOeDnq5Gqqv44VSU/cbcf6B13/fGm/jH3G3H+gdd/wB8ab+MztO0Rt097EXx8Vr785686o9R6porfpmp0zTay1RHbKpJUlYmmm77u0VVf6faZTOV9561msKKvkssk+rdQ/5DeklA1mlmNbCqNRqcEk4oiInM3n3HXD+gdd/3xpv4yF67uC6dclrgj1dQXZ6NennGo0qmtYqr9GJV9JeiKv8AgazvWOf7umOK5LbV6/8Az8fB50FHU7VdodNZI7tW3KhqKp1ZUVVRTNgeqK1vauVjVXHosaxvHhw5ZOvoII6aGOCFjY442oxjGphGtRMIiewrTYVsxdoaxOuFzi3b3cmo6ZHc4I+bYvX1d48OhZxIwUmsbz1lU9o6iuS8Up+WvL+ZCoNvelZJ4KbUlMxXebt83qUROTFXLXepFVUX1oW+fE0MdRE+GaNkkcjVa9j0yjkXmip1QnabPOHJGSFLq9NXUYpxT4udNnG1Ko0W19BWQPq7ZI/f3WKiSQuXmrc8FReqL1495Zq7dNIJF2iSV7nY/k0pl3vtx8TT6k2BUVXO+osNf5jvLnzaZqvjT+qqcUTwXJG2eT/qRZN11xtTW/nbz1+G6W9/Qc895adpUGL8S0sd1WvFEdPH7+KM7QtZ/dvfkuDKZ1PBFEkMUblRXbqKqqq46qq8i/tnFik09oy2UM7d2fs+1lTue9d5U9mcewjejNidt0/Vx3C6VPynVRKjo2bm5FG5OuOKuVOmeHgWSRNdqsdqVw4fywndmaLLS9tRqPzWckVf5Tzf6Qd++U6pvn4luH1aX9lSpptg1xlu0lcl7pEa6pWfd7B2cK/exzLfr6ZaygqaZrkas0T40cqZxlFTPxNtfqMeTu+Cd9v6a9laTLhjL3ldt+n1cs7P/wAs9P8A12H7Tq4p7Tewuvsd8ttykvVLKyjnZMrGwORXI1eSLkuE17Tz4816zjnfk27F0uXBjtGWNt5Vf5QX5LUH15P3bzR+Tv8A99vn/pQ/tPLA2kaKn1xaKaggrIqR0NQkyvkYrkVN1yYwip3mv2Z7N6nQk9fLUXCGrSqZG1EjjVm7uqq9VXvNqajHGinFM+t/bGTS5Z7RrmiPV8/ghW3/AE9JDc6K/RtVYZ4/NpVT6L25VufWir/qmLsy2uQ6XtyWa8wzSUcblWCaFN50SKuVarc8UyqqmOKZLwutqor3b5rfcKdlRTTN3Xxu6/8A4viVDefJ7n84c+y3iLsFXKR1jF3m+G83n7jpptTgyYe41HLbpLjq9HqcOonU6Xnv1hu9SbdbHSUEiWRJq2tc38Gr41ZHGve5V4rjuT4FHRSTTXaOWpV7ppKhkj3PTi5znoqqvrzkuDTGwKOlq2VOoLhHVxxqi+a07VRj/Bzl4qngiJ6yQa02R0OprpDdqKpS31jFZ2idnvRyo1Uxw4YXCYynRE4HXDqdLp5mmPnv1lwz6PXausZMu0THSP1Te40EN0t9TQVCKsNTE6J6Jz3XIqLj3nLNZT3HQOrXxsduVltqN6N6pwe3m1fFHN+1Tq4i+t9nlp1vTt87R1PWRJiKriRN9qdypyc3wX2YIOg1cYbTW/5ZWfamgtqKxbHytXo0Np266ZqqFJLh5zQ1LW+nD2SyIq/qubzT14Kz2nbSF1vUQ09JFJT2ylcr2Nk+fK/GN5yJwTCck8VNtU+T/qGOZW01xts0WeD3q9i48Uwv2kp0fsLorTUx1t9qmXGWNd5tOxithRU5K7PF3q4J6ydSdFgnvaTvPhH3+6svXtLVV7i9do8Z+/2ZWw3TE1m05NcqqN0c1yej2scmFSJuUbw8VVy+pUIbt403JQ6ghvsbVWnr2JG9cfNlYmML6249yl9omEwhh3ez0N9t81vuNOyoppkw5jvgqL0VOioQMWutXUTnnx6+5aZuza30kaas9Ok+1UGznbLR2q0wWfUDZmpTNSOGqjbvorE5NeiccpyymeHx2Gs9uNtW2z0enEnmqpWrGlS9isZEi83Jniq93DBrb55PtY2d77JdYZIVXLYqtFa9vhvNRUX3IeVp8n25yTNW7Xalgh+k2mar3r4IrkRE+JPmNBNu9mfh/Srie0607iK+zf8Avdpdi2mpL1q6Kuc1fNbYnbvd3yLlGNz35yvsJh5Q/wCLLL9Yk/YLK09p226YtsdutdOkMDOK8cue7q5y9VUj20vQVRrqkoYKeuipFpZXSKskav3stxjgqEb02uTV1y25VhM/DrYtDbDXnaf5hpPJ/wDyTrfrzv2GFnEU2caLn0RZp7fPWR1bpahZkfGxWoiK1ExhVXuJWQtXet81rV6SsdBjtj09KXjaYhz7t00/JbtUtuzWr5vcY09LHBJGJhU9qbq+82mzjbJSWW0xWa/sn7OmTcgqYm7/AKHRrk58OipngWprCjsVdYKmLUawst3BXySu3ezdnDXI7ouV4L4lGVGziy1tyZS2DW9nqlnejIop1VJFVeSZblFX2IWOHV6fLhjDqJ2mOil1Wmz6bUzm00xPF4f0m2sNuVrjt01Np3tqislarG1D2KxkOfpceLl7k5FS6Pa5usbK16OR6XCBHI7nntE5+JPdLaH0fYNRRwaj1Rbqy4QyIjaKNVbEkmeCPcvNUX6PDjz7iQa40lpe0auoNS1N+p7PO6pjqZaeVu8k6tciq5qJxaq44rxTJtj1ekwRbHjnlMdfa5ZdPqtTtnyzHqzHLeOUfNv9sf8AN5dPXF+9aVXsM/Ltv1Sb/dLl1JbKbaBpCajt1xh7CtRjmVLE7RuGvRei+GCC6f0datk9+iu171TRtSSF8TIXRKxzs44pxVcJjuOGn1OKmkvS0853TdZp721mPPEerG287x5ynm0T8hr79Sl/ZKX2Hfl4z6rN/ulv6rvmn7rpCdsl+oaakukL6eGqc9FYrlRc448VTuILs101YbBqCa602sbbcUpaSV0scTN1WR8MvVd5cIn+JpptTipp8mO1uc9GdZitk1mLJXbaOvOFla4/I2+fUZv2FKQ2F/l036nL/ulxXvUOnr5pueCO+0McNzbJQw1G+itWRUxhO9UynAiOgdntJo3Wm6/UlJV1zaV+aJsW7Juux6Xzl4cPiNLqcdNPkpaec9G2sw2y6rFkpziOvOFrAArl0AAAAAAAAAAAYF+sVBqW0VVoucCT0lUxWSMX4Ki9FRcKi9FQzwNmYmYneHE2uNFV2zHVy0dXTU9bC13bUstTAj4qqLP0mrzVOTk6L7FJNpS6WHUkE7qm1bL7JLC5E7K40MzO0RfpNckmF8U5odE7SNn9BtD07JbKvdiqWZkpKnGXQSY5+LV5KnVPFEOMr9Yrhpq71VoukDqespX7kjF5eCovVqpxReqKV2Sk4p3jo9Vo9RXW4+G07Xj73XD8l6a/Sti3+zzf8Q/fkvTX6VsW/wBnm/4hRwNO+9iT6DP+cpZrq726SpltFvtOkGxwvRVuFmonxpKuOKNc9yruovVMZx3FreT5sfYscOsb9SpxVH22menLjwncn7KL6+4h+w3ZO7XN1+VrrC75Con+k1eCVcqcezT9VObvYnVcdZsY2NiMY1GtamEREwiId8GLinjsrO0tZGKvo2KefjL9ABNeeAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEJ2z5+9zdsc/wP71pHtIX3Zo6K0U9NHb0vnZRNa5KRyP7dGcfS3cZznjkkO2f+bm7f/D+9aavSOtNC1FFaLXBJSfKz4IoGolIqO7bcRPnbvPPXJHt/5PBW5ZiNTzmI5R1989EX0bomHV2ym4tp6Skfd56x6xVM3ByKjm834VUTG97zaWKyRVm1iWgvcENc+islO1WzJ2jEejWIqpnnxV3vI5Y9bR6X2bXWwxVdRRajZWPZFE2Nd9FV7c4XGOip/wBTeUl8Zo/aW246oklpUrLJAxZpGK7elRrN7OE57zXJ6/WcqzXaPhv9UPHbFw09m2/l49fa3ew1OztN8p2KqRQ3WVsbOjUwnBDV6jntFv2y+capZClultyNp31bN6Le9vD872qbbYbHI+xXauWN7IKy5Sywq5Mb7cImfflPYpkaq1lZ6XVqad1Va6JLW+m7eGsqW76K7kqImOHJUyi93edI27uN0qsR6NTeducdej4tehrJatGXjzapp7xQzpPW0qvjY9kCrGqegvHuTjz4EOtaUds2D1NxZTQMrqpr6RahGIkj2umxuq7mqYTl4G32Vtc3Qeq3RdolvdNUrSb2cbnZryz7PaRR75K3ZTpOxQORJbldHtaq+D3Jn3vRTSZjaJiPCUe9q8EWrG3qz+sQxbXTz1OzC/0uN2qslyirWp1blN132KWHsvqE1TrDUmrML2b0hpIMp0RqK77G+8jmnbRUWDVGsNOXirZWyVtofUPla3dSVcZzjv8ASX3Ey2FQRRbP6aVnz55pZJF73b279iIMVfWiPvl/s0eOe8pE+G+/viZiP1WAACYvAAAAAAAAAAAAAAKq25bJW63tnyxaYUS+0TPRRvDzuNMr2a/rJzavs68LVBresWjaXXBmthvF6dYfz6ex0b3Me1zXNVUc1yYVFTmip0UlGzjZ/cNoeoY7ZSb0VMzElXVYykEeefi5eSJ1XwRS79sewWbU90S/6YSGKuqZGtrKZ6oxkiquFmReionFydcZTjzsnZ9oS27P9PRWmgbvyfPqalUw6ol6uXuToidEINNNPHtbo9Hn7Yp3EWx/mn6ffg2un7Bb9MWeltFrgSCkpmbjGpzXvVV6qq8VXqqmwALCI2eYmZmd5AAGAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAB41lFTXCndTVlPDUwPxvRysRzXYXKZReHMwafSthpZ2T09ltsM0a7zJGUzGuaveionA2gMbQ1mlZneYa2XTNkmuKXKS00L61FRyVDoGq/Kclzjn4npdLFa72xjLnb6WtbGuWJPEj931Z5GcBtBwV6bPOmpoKOBlPTQxwwxpusjjajWtTuRE5GFedOWjUDI2XW3Utaka5Z2zEcre/C80NiBtHRmaxMbTHJj01toqOhbQU9LBFSNarEgYxEYjV5pjkeDNP2iJtM1lroWtpXK+BEgaiQuVcqreHornuM8DaGOGPJiyWm3y1bqyShpn1Lo1idM6JqvVi825xnHgfdFQUltp201FTQ0sDVVUjhYjGpnnwQ9wNmeGI5gAMsgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD/9k=';
// The exact "Payment receipt" folder the user already created — using its
// ID (from the folder's own URL) instead of searching by name, since a
// name search only looks at the top level of My Drive and this folder
// isn't there.
var RECEIPTS_FOLDER_ID = '1bMOGMDEEn9aaP5p_Z1wBbsoUNdMEOkrg';
var RECEIPT_FONT = 'Arial';
var RECEIPT_FILL_COLOR = '#F1E9DA';

function getOrCreateFolder(parent, name) {
  var existing = parent.getFoldersByName(name);
  if (existing.hasNext()) return existing.next();
  return parent.createFolder(name);
}

// The two receipt "stages" only differ in this copy. Amount Received/Total
// Amount/dates are filled in by the caller either way.
function getReceiptContent(input) {
  if (input.stage === 'paid') {
    return {
      statusLine: 'Your booking is confirmed and complete.',
      balanceLines: ['Full payment received.', 'Booking confirmed and complete — no balance due.'],
      nonRefundable: 'All payments received are non-refundable.'
    };
  }
  var remaining = Math.max(0, Number(input.totalAmount) - Number(input.amountReceived));
  var dueDate = new Date(input.checkIn + 'T00:00:00');
  dueDate.setDate(dueDate.getDate() - 1);
  var dueDateStr = Utilities.formatDate(dueDate, Session.getScriptTimeZone(), 'd MMM yyyy');
  return {
    statusLine: 'Your booking has been confirmed, pending the remaining balance.',
    balanceLines: ['Booking confirmed, balance of ' + formatMoney(remaining) + ',', 'Due one day before check-in, i.e. ' + dueDateStr + '.'],
    nonRefundable: 'Advance payment is non-refundable.'
  };
}

// Sets bold/color/size on a paragraph's text without relying on Paragraph
// having its own style setters — editAsText() always works.
function styleParaText(paragraph, opts) {
  var t = paragraph.editAsText();
  t.setFontFamily(RECEIPT_FONT);
  if (opts.bold) t.setBold(true);
  if (opts.color) t.setForegroundColor(opts.color);
  if (opts.size) t.setFontSize(opts.size);
}

// A borderless 2-column [label, value] row, value cell highlighted — the
// same "fill-in line" look as the paper template.
function appendFieldRow(body, label, value) {
  var row = body.appendTable([[label, String(value || '')]]);
  row.setBorderWidth(0);
  row.getCell(0, 0).setWidth(150);
  styleParaText(row.getCell(0, 0).getChild(0).asParagraph(), { bold: true, size: 10 });
  var valueCell = row.getCell(0, 1);
  valueCell.setBackgroundColor(RECEIPT_FILL_COLOR);
  valueCell.setPaddingTop(5).setPaddingBottom(5).setPaddingLeft(8);
  styleParaText(valueCell.getChild(0).asParagraph(), { bold: true, size: 10 });
  return row;
}

// Builds the payment receipt as a temporary Google Doc, exports it as a PDF
// blob, then trashes the Doc — only the PDF bytes are kept/returned. `input`
// needs: bookingNumber, guestName, checkIn, checkOut, guests, totalAmount,
// amountReceived, receivedDate, rateIncludes, stage ('advance' or 'paid').
function buildReceiptPdf(input) {
  var content = getReceiptContent(input);
  var doc = DocumentApp.create('Receipt - ' + input.bookingNumber + ' - ' + input.stage + ' - TEMP');
  var docId = doc.getId();
  var body = doc.getBody();
  body.setMarginTop(30).setMarginBottom(30).setMarginLeft(40).setMarginRight(40);

  // Letterhead: logo left, business info right.
  var head = body.appendTable([['', '']]);
  head.setBorderWidth(0);
  var logoCell = head.getCell(0, 0);
  logoCell.clear();
  try {
    var logoBlob = Utilities.newBlob(Utilities.base64Decode(LOGO_BASE64), 'image/jpeg', 'logo.jpeg');
    var blankPara = logoCell.getChild(0).asParagraph();
    var img = logoCell.appendImage(logoBlob);
    blankPara.removeFromParent();
    var ratio = img.getHeight() / img.getWidth();
    img.setWidth(150);
    img.setHeight(Math.round(150 * ratio));
  } catch (imgErr) {
    Logger.log('Logo embed failed: ' + imgErr);
    var logoFallback = logoCell.getChild(0).asParagraph();
    logoFallback.setText('Sindooram Ecostay');
    styleParaText(logoFallback, { bold: true });
  }
  var infoCell = head.getCell(0, 1);
  infoCell.clear();
  var infoP1 = infoCell.getChild(0).asParagraph();
  infoP1.setText('Sindooram Ecostay');
  infoP1.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  styleParaText(infoP1, { bold: true, size: 10 });
  var infoP2 = infoCell.appendParagraph('Opposite Thakadi Temple, Edava, Varkala, Kerala 695311');
  infoP2.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  styleParaText(infoP2, { size: 9 });
  var infoP3 = infoCell.appendParagraph('+91 98460 22350  |  sindooramecostays.com');
  infoP3.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  styleParaText(infoP3, { size: 9 });

  body.appendParagraph('');

  // Banner.
  var bannerTable = body.appendTable([['PAYMENT RECEIPT']]);
  bannerTable.setBorderWidth(0);
  var bannerCell = bannerTable.getCell(0, 0);
  bannerCell.setBackgroundColor('#2F4739');
  bannerCell.setPaddingTop(10).setPaddingBottom(10);
  var bannerP = bannerCell.getChild(0).asParagraph();
  bannerP.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  styleParaText(bannerP, { bold: true, color: '#FFFFFF', size: 15 });

  body.appendParagraph('');
  styleParaText(body.appendParagraph('Greetings from Sindooram Ecostay!'), { bold: true, size: 11 });
  body.appendParagraph('');

  appendFieldRow(body, 'Booking Number:', input.bookingNumber);
  appendFieldRow(body, 'Received From:', input.guestName);
  appendFieldRow(body, 'Amount Received:', formatMoney(input.amountReceived));
  appendFieldRow(body, 'Total Booking Amount:', formatMoney(input.totalAmount));
  appendFieldRow(body, 'Received Date:', input.receivedDate);

  body.appendParagraph('');
  styleParaText(body.appendParagraph(content.statusLine), { bold: true, color: '#8F4128', size: 11 });

  body.appendParagraph('');
  body.appendHorizontalRule();
  body.appendParagraph('');

  styleParaText(body.appendParagraph('Booking Details'), { bold: true, color: '#8F4128', size: 12 });
  appendFieldRow(body, 'Check-in', input.checkIn);
  appendFieldRow(body, 'Check-out', input.checkOut);
  appendFieldRow(body, 'Guests', String(input.guests || ''));
  appendFieldRow(body, 'Rate includes', input.rateIncludes || '');

  body.appendParagraph('');
  styleParaText(body.appendParagraph('Payment Terms & Conditions'), { bold: true, color: '#8F4128', size: 12 });

  var balanceTable = body.appendTable([['']]);
  balanceTable.setBorderWidth(0);
  var balanceCell = balanceTable.getCell(0, 0);
  balanceCell.setBackgroundColor(RECEIPT_FILL_COLOR);
  balanceCell.setPaddingTop(8).setPaddingBottom(8).setPaddingLeft(10).setPaddingRight(10);
  balanceCell.getChild(0).asParagraph().removeFromParent();
  content.balanceLines.forEach(function (line) {
    var li = balanceCell.appendListItem(line);
    li.setGlyphType(DocumentApp.GlyphType.BULLET);
    styleParaText(li, { size: 10 });
  });

  var bullet = body.appendListItem('Early check-in is subject to availability');
  bullet.setGlyphType(DocumentApp.GlyphType.BULLET);
  bullet.editAsText().setFontFamily(RECEIPT_FONT).setFontSize(10);
  styleParaText(body.appendParagraph(content.nonRefundable), { bold: true, size: 10 });

  body.appendParagraph('');
  styleParaText(body.appendParagraph('If you have any questions before your stay, feel free to reach out to us directly on WhatsApp: +91 98460 22350.'), { size: 10 });
  styleParaText(body.appendParagraph('Feel free to check out our website for things to do nearby and more to help you plan your stay: sindooramecostays.com.'), { size: 10 });
  body.appendParagraph('');
  styleParaText(body.appendParagraph('Looking forward to welcoming you all soon!'), { size: 10 });
  styleParaText(body.appendParagraph('Thanks,'), { size: 10 });
  body.appendParagraph('');
  styleParaText(body.appendParagraph('warm regards'), { size: 10 });
  styleParaText(body.appendParagraph('Team Sindooram'), { bold: true, size: 10 });

  doc.saveAndClose();
  var pdfBlob = DriveApp.getFileById(docId).getAs('application/pdf');
  pdfBlob.setName('Receipt-' + input.bookingNumber + '-' + input.stage + '.pdf');
  DriveApp.getFileById(docId).setTrashed(true);
  return pdfBlob;
}

// Turns a doPost payload into everything buildReceiptPdf/receiptEmailContent
// need, pulling the booking's stored fields and layering the caller-supplied
// amount/date/rate-includes on top (those three aren't tracked on the
// booking itself — they're specific to this one receipt).
function findBookingInput(data) {
  var all = readAll();
  var booking = all.bookings.filter(function (b) { return b.id === data.bookingId; })[0];
  if (!booking) throw new Error('Booking not found — try saving the booking again first.');
  var stage = booking.status === 'Confirmed-Paid' ? 'paid' : 'advance';
  return {
    bookingId: booking.id,
    bookingNumber: booking.bookingNumber,
    guestName: booking.guestName,
    guestEmail: booking.guestEmail,
    checkIn: booking.checkIn,
    checkOut: booking.checkOut,
    guests: booking.guests,
    totalAmount: Number(booking.amount) || 0,
    amountReceived: Number(data.amountReceived) || 0,
    receivedDate: Utilities.formatDate(data.receivedDate ? new Date(data.receivedDate + 'T00:00:00') : new Date(), Session.getScriptTimeZone(), 'd MMM yyyy'),
    rateIncludes: data.rateIncludes || '',
    stage: stage
  };
}

function receiptEmailContent(input) {
  var stagePretty = input.stage === 'paid' ? 'Payment Received in Full' : 'Advance Payment Received';
  var content = getReceiptContent(input);
  var bodyText = 'Hi ' + input.guestName + ',\n\n' +
    'Thank you! Please find attached your payment receipt for booking ' + input.bookingNumber + ' at Sindooram Ecostay.\n\n' +
    content.statusLine + '\n\n' +
    'Check-in: ' + input.checkIn + '\nCheck-out: ' + input.checkOut + '\n\n' +
    'If you have any questions before your stay, reach us on WhatsApp: +91 98460 22350.\n\n' +
    'Looking forward to welcoming you all soon!\n\nRegards,\nTeam Sindooram';
  return { subject: 'Payment Receipt — Booking ' + input.bookingNumber + ' (' + stagePretty + ')', bodyText: bodyText };
}

// doPost action 'previewReceipt' — builds the PDF and email text but sends
// nothing; the app shows this to the manager first.
function handlePreviewReceipt(data) {
  var input = findBookingInput(data);
  var pdfBlob = buildReceiptPdf(input);
  var email = receiptEmailContent(input);
  return {
    pdfBase64: Utilities.base64Encode(pdfBlob.getBytes()),
    fileName: pdfBlob.getName(),
    subject: email.subject,
    bodyText: email.bodyText,
    toEmail: input.guestEmail
  };
}

// doPost action 'sendReceipt' — rebuilds the same PDF (from the same inputs
// the app just showed in Preview), actually emails it, saves a copy to
// Drive, and logs it to the Receipts sheet tab for audit/reminder lookups.
function handleSendReceipt(data) {
  var input = findBookingInput(data);
  if (!input.guestEmail) throw new Error('This booking has no Guest Email — add one before sending a receipt.');
  var pdfBlob = buildReceiptPdf(input);
  var email = receiptEmailContent(input);
  MailApp.sendEmail({ to: input.guestEmail, subject: email.subject, body: email.bodyText, attachments: [pdfBlob] });
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var receiptsFolder = DriveApp.getFolderById(RECEIPTS_FOLDER_ID);
  var bookingFolder = getOrCreateFolder(receiptsFolder, input.bookingNumber);
  var savedFile = bookingFolder.createFile(pdfBlob);
  logReceipt(ss, input, savedFile.getUrl());
  return { status: 'sent' };
}

function logReceipt(ss, input, fileUrl) {
  var sheet = ss.getSheetByName('Receipts') || ss.insertSheet('Receipts');
  if (sheet.getLastRow() === 0) {
    var headers = ['ID', 'Booking ID', 'Booking Number', 'Guest Name', 'Guest Email', 'Stage', 'Amount Received', 'Total Amount', 'Received Date', 'Rate Includes', 'Check-in', 'Check-out', 'Guests', 'Sent At', 'PDF Link'];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  var id = 'rc_' + new Date().getTime().toString(36) + Math.random().toString(36).slice(2, 7);
  sheet.appendRow([id, input.bookingId, input.bookingNumber, input.guestName, input.guestEmail, input.stage, input.amountReceived, input.totalAmount, input.receivedDate, input.rateIncludes, input.checkIn, input.checkOut, input.guests, new Date(), fileUrl]);
  sheet.autoResizeColumns(1, 15);
}

function latestReceiptForBooking(ss, bookingId, stage) {
  var sheet = ss.getSheetByName('Receipts');
  if (!sheet) return null;
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return null;
  var headers = values[0];
  var col = {};
  headers.forEach(function (h, i) { col[h] = i; });
  var matches = values.slice(1).filter(function (row) {
    return row[col['Booking ID']] === bookingId && row[col['Stage']] === stage;
  });
  if (!matches.length) return null;
  var last = matches[matches.length - 1];
  return { amountReceived: last[col['Amount Received']] };
}

// Daily reminder for Direct bookings still on Confirmed-Advance (balance
// outstanding) whose check-in is tomorrow. Plain text, no attachment.
// Skips anything already Confirmed-Paid — nothing owed, nothing to remind.
function sendPaymentReminders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var data = readAll();
  var tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  var tomorrowStr = Utilities.formatDate(tomorrow, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  data.bookings.forEach(function (b) {
    if (b.source !== 'Direct') return;
    if (b.status !== 'Confirmed-Advance') return;
    if (String(b.checkIn).slice(0, 10) !== tomorrowStr) return;
    if (!b.guestEmail) return;
    var receipt = latestReceiptForBooking(ss, b.id, 'advance');
    var amountReceived = receipt ? Number(receipt.amountReceived) : (Number(b.amount) || 0) / 2;
    var remaining = Math.max(0, (Number(b.amount) || 0) - amountReceived);
    var checkInLabel = Utilities.formatDate(new Date(b.checkIn + 'T00:00:00'), Session.getScriptTimeZone(), 'd MMM yyyy');
    var dueDateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'd MMM yyyy');
    var body = 'Hi ' + b.guestName + ',\n\n' +
      'Just a friendly reminder — your stay at Sindooram Ecostay (Booking ' + b.bookingNumber + ') checks in tomorrow, ' + checkInLabel + '.\n\n' +
      'The remaining balance of ' + formatMoney(remaining) + ' is due by ' + dueDateStr + ' (one day before check-in).\n\n' +
      'If you have already paid this, please disregard this message. For any questions, reach us on WhatsApp: +91 98460 22350.\n\n' +
      'Looking forward to welcoming you!\n\nRegards,\nTeam Sindooram';
    MailApp.sendEmail({ to: b.guestEmail, subject: 'Reminder: Balance due for your upcoming stay — ' + b.bookingNumber, body: body });
  });
}

// Run this once yourself (select it above, then click Run) to schedule
// sendPaymentReminders() daily at 9am. Re-running it is safe — it clears
// any previous schedule for this function first.
function createReminderTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sendPaymentReminders') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendPaymentReminders').timeBased().everyDays(1).atHour(9).create();
}

// Run this once yourself to sanity-check receipt generation without sending
// a real email — saves a sample PDF to your Drive root so you can check it
// looks right before relying on it for real guests.
function testBuildReceiptPdf() {
  var input = {
    bookingId: 'test', bookingNumber: 'SE-2026-TEST', guestName: 'Test Guest',
    guestEmail: '', checkIn: '2026-09-30', checkOut: '2026-10-04', guests: 7,
    totalAmount: 20574, amountReceived: 10287,
    receivedDate: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'd MMM yyyy'),
    rateIncludes: 'Breakfast from our set menu', stage: 'advance'
  };
  var pdfBlob = buildReceiptPdf(input);
  DriveApp.getRootFolder().createFile(pdfBlob);
  Logger.log('Saved test receipt to your Drive root: ' + pdfBlob.getName());
}
