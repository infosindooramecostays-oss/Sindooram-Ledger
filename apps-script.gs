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
  if (data.action === 'getReceiptInfo') {
    try { return jsonResponse({ receipt: latestReceiptSentInfo(SpreadsheetApp.getActiveSpreadsheet(), data.bookingId) }); }
    catch (err) { return jsonResponse({ error: String(err.message || err) }); }
  }
  // Everything above this point is an action with no transactions/bookings
  // of its own — a receipt request, a report send, etc. Reaching here with
  // an action this deployed version doesn't recognize (the exact way a
  // stale deployment causes damage: a newly-added action like
  // 'getReceiptInfo' hits an OLDER doPost that has no branch for it) used
  // to silently fall through to `data.transactions || []` — quietly
  // treating "these fields aren't here because this wasn't a save at all"
  // as "here's an empty replacement for everything you have". That's what
  // caused the 2026-09-26 5:36-5:38am wipe (and, we now believe, the
  // original one too) — both traced back to a receipt click. Requiring
  // both fields to genuinely be arrays closes that off for good: an
  // unrecognized action now fails loudly instead of wiping anything,
  // regardless of which version of this file is actually deployed.
  if (!Array.isArray(data.transactions) || !Array.isArray(data.bookings)) {
    throw new Error('Unrecognized request: action "' + data.action + '" was not matched, and no transactions/bookings arrays were sent. Refusing to touch the Sheet — this deployment may be out of date.');
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var incomingTransactions = data.transactions;
  var incomingBookings = data.bookings;
  // The app never has a "delete everything" feature — removals always take
  // out one row at a time, so a genuine save should never send back a
  // fully empty list for a sheet that currently has real data. Confirmed
  // live on 2026-09-26: a client somehow synced with empty local state and
  // this path silently wiped both sheets down to just their headers, with
  // no confirmation and no trace — recovered only via Sheets version
  // history. Refuse the write instead of repeating that.
  guardAgainstEmptyOverwrite(ss, 'Transactions', incomingTransactions);
  guardAgainstEmptyOverwrite(ss, 'Bookings', incomingBookings);
  assertDirectBookingsComplete(incomingBookings);
  writeTransactions(ss, incomingTransactions);
  writeBookings(ss, incomingBookings);
  return jsonResponse(readAll());
}

// The app's own form already requires these fields for every Direct
// booking via HTML "required" attributes — but that only protects saves
// made through the app's UI. A booking written straight to this endpoint
// (an automation, a manual API call) skips that entirely. This is the
// same rule enforced server-side, so it can't be bypassed that way either.
// Scoped to Confirmed/Confirmed-Advance/Confirmed-Paid specifically,
// matching what these statuses actually mean (only Direct bookings use
// them) — NOT every Direct booking regardless of status, since that would
// retroactively break saves on older Direct rows that are Pending,
// Cancelled, etc. and were never required to be complete.
var CONFIRMED_STATUSES = ['Confirmed', 'Confirmed-Advance', 'Confirmed-Paid'];
function assertDirectBookingsComplete(bookings) {
  var requiredFields = ['bookingNumber', 'guestName', 'guestEmail', 'checkIn', 'checkOut', 'amount', 'guests'];
  bookings.forEach(function (b) {
    if (b.source !== 'Direct') return;
    if (CONFIRMED_STATUSES.indexOf(b.status) === -1) return;
    var missing = requiredFields.filter(function (field) {
      var v = b[field];
      return v === undefined || v === null || String(v).trim() === '';
    });
    if (missing.length) {
      throw new Error('Booking ' + (b.bookingNumber || b.guestName || b.id) + ' is missing required field(s) for a ' + b.status + ' Direct booking: ' + missing.join(', '));
    }
  });
}

// Throws (uncaught, on purpose) if asked to overwrite a sheet that
// currently has real data with an empty list — see the note in doPost()
// above for why. Throwing here means the client's fetch sees a failed
// request and falls into its existing "Could not save" error handling,
// rather than the request quietly succeeding with an empty result (the
// app has no code path today that inspects a JSON {error: ...} body from
// this endpoint, so returning one instead of throwing would be treated as
// a successful, empty save — exactly the failure mode this exists to
// prevent). Also emails the team, since a legitimate save should never
// hit this — it means something upstream failed silently.
function guardAgainstEmptyOverwrite(ss, sheetName, incomingRows) {
  if (incomingRows.length > 0) return;
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return;
  if (sheet.getLastRow() <= 1) return; // already empty — nothing to lose
  MailApp.sendEmail({
    to: TEAM_ALERT_EMAILS,
    subject: '[Blocked] A save tried to wipe the "' + sheetName + '" sheet',
    body: 'A save request tried to overwrite "' + sheetName + '" — which currently has ' + (sheet.getLastRow() - 1) + ' row(s) of real data — with an empty list.\n\n' +
      'This has been blocked automatically. Nothing was changed in the Sheet.\n\n' +
      'This should never happen during normal use (the app only ever removes one row at a time, never all of them). It most likely means someone\'s copy of the app failed to load data properly before it tried to save. Ask them to reload the app and check their connection before saving again.'
  });
  throw new Error('Refused to save — "' + sheetName + '" currently has data, but this save would have wiped it. Nothing was changed. Reload the app and try again.');
}

// Bump this any time doGet/doPost or anything they call changes, together
// with the matching EXPECTED_SCRIPT_VERSION constant in index.html. A
// second data wipe (2026-09-26, 5:36-5:38am) traced back to the Web App
// silently still running an OLD deployed version -- one that predated the
// empty-overwrite guard entirely, so the guard never actually ran, and
// Manage Deployments gives no obvious warning when a redeploy didn't
// stick. Sending this back on every response is what lets the app (and
// anyone checking) tell definitively whether a redeploy actually took
// effect, instead of trusting the deployments UI alone.
var SCRIPT_VERSION = '2026-09-26-3';

// Reads both sheets and returns them as plain JS objects/arrays.
function readAll() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  backfillMissingBookingIds(ss);
  return {
    transactions: readSheet(ss, 'Transactions', ['id','type','date','category','subcategory','description','amount','addedBy','recurrence','createdAt','updatedAt','endsOn']),
    bookings: readSheet(ss, 'Bookings', ['id','bookingNumber','guestName','guestEmail','guestPhone','checkIn','checkOut','source','amount','status','remarks','addedBy','createdAt','updatedAt','guests']),
    scriptVersion: SCRIPT_VERSION
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
var LOGO_BASE64 = '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';
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

// A bullet line built as plain text with a manually-prefixed "•" instead of
// a native list item — Apps Script's list items always use a plain black
// glyph with no way to recolor just the bullet, which reads jarring next
// to this warm palette. Coloring only the bullet character keeps the body
// text itself in the normal ink color.
var BULLET_COLOR = '#8F4128';
function appendBulletLine(cell, text, opts) {
  opts = opts || {};
  var bullet = '•  ';
  var p = cell.appendParagraph(bullet + text);
  var t = p.editAsText();
  t.setFontFamily(RECEIPT_FONT);
  if (opts.size) t.setFontSize(opts.size);
  if (opts.bold) t.setBold(true);
  t.setForegroundColor(0, 0, BULLET_COLOR);
  return p;
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
  // DocumentApp.create always drops the new Doc in Drive root — move it into
  // the Payment receipt folder immediately so it never sits visible in root,
  // even briefly, before it gets trashed below.
  var tempFile = DriveApp.getFileById(docId);
  DriveApp.getFolderById(RECEIPTS_FOLDER_ID).addFile(tempFile);
  DriveApp.getRootFolder().removeFile(tempFile);
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
  infoP1.setSpacingBefore(0).setSpacingAfter(2).setLineSpacing(1);
  styleParaText(infoP1, { bold: true, size: 10 });
  var infoP2 = infoCell.appendParagraph('Opposite Thakadi Temple, Edava, Varkala, Kerala 695311');
  infoP2.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  infoP2.setSpacingBefore(0).setSpacingAfter(2).setLineSpacing(1);
  styleParaText(infoP2, { size: 9 });
  var infoP3 = infoCell.appendParagraph('+91 98460 22350  |  sindooramecostays.com');
  infoP3.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  infoP3.setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1);
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

  body.appendPageBreak();
  styleParaText(body.appendParagraph('Payment Terms & Conditions'), { bold: true, color: '#8F4128', size: 12 });

  var termsTable = body.appendTable([['']]);
  termsTable.setBorderWidth(0);
  var termsCell = termsTable.getCell(0, 0);
  termsCell.setBackgroundColor(RECEIPT_FILL_COLOR);
  termsCell.setPaddingTop(10).setPaddingBottom(10).setPaddingLeft(12).setPaddingRight(12);
  termsCell.getChild(0).asParagraph().removeFromParent();
  content.balanceLines.forEach(function (line) {
    appendBulletLine(termsCell, line, { size: 10 });
  });
  appendBulletLine(termsCell, 'Early check-in is subject to availability', { size: 10 });
  appendBulletLine(termsCell, content.nonRefundable, { size: 10, bold: true });

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
  logReceipt(ss, input, savedFile.getUrl(), String(data.sentBy || '').trim());
  return { status: 'sent' };
}

// Writes columns by header NAME, not fixed position — self-healing if the
// sheet already exists from before 'Sent By' was added (appends the header
// once, then every write since lines up under it), same lesson as the
// ID Vault "stale headers" bug: a fixed-position appendRow() silently goes
// out of sync the moment the header row and the code's write order drift.
function logReceipt(ss, input, fileUrl, sentBy) {
  var sheet = ss.getSheetByName('Receipts') || ss.insertSheet('Receipts');
  var baseHeaders = ['ID', 'Booking ID', 'Booking Number', 'Guest Name', 'Guest Email', 'Stage', 'Amount Received', 'Total Amount', 'Received Date', 'Rate Includes', 'Check-in', 'Check-out', 'Guests', 'Sent At', 'Sent By', 'PDF Link'];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(baseHeaders);
    sheet.getRange(1, 1, 1, baseHeaders.length).setFontWeight('bold');
  }
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headers.indexOf('Sent By') === -1) {
    sheet.getRange(1, headers.length + 1, 1, 1).setValue('Sent By').setFontWeight('bold');
    headers = headers.concat(['Sent By']);
  }
  var id = 'rc_' + new Date().getTime().toString(36) + Math.random().toString(36).slice(2, 7);
  var values = {
    'ID': id, 'Booking ID': input.bookingId, 'Booking Number': input.bookingNumber,
    'Guest Name': input.guestName, 'Guest Email': input.guestEmail, 'Stage': input.stage,
    'Amount Received': input.amountReceived, 'Total Amount': input.totalAmount,
    'Received Date': input.receivedDate, 'Rate Includes': input.rateIncludes,
    'Check-in': input.checkIn, 'Check-out': input.checkOut, 'Guests': input.guests,
    'Sent At': new Date(), 'Sent By': sentBy || '', 'PDF Link': fileUrl
  };
  var row = headers.map(function (h) { return values[h] !== undefined ? values[h] : ''; });
  sheet.getRange(sheet.getLastRow() + 1, 1, 1, row.length).setValues([row]);
  sheet.autoResizeColumns(1, headers.length);
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

// Powers the "already sent by X on Y" notice in the app — lets a second
// person opening the same booking see a receipt went out before they
// double-send one. Any stage counts (advance or paid), most recent wins.
function latestReceiptSentInfo(ss, bookingId) {
  var sheet = ss.getSheetByName('Receipts');
  if (!sheet) return null;
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return null;
  var headers = values[0];
  var col = {};
  headers.forEach(function (h, i) { col[h] = i; });
  var matches = values.slice(1).filter(function (row) {
    return row[col['Booking ID']] === bookingId;
  });
  if (!matches.length) return null;
  var last = matches[matches.length - 1];
  var sentAt = last[col['Sent At']];
  return {
    stage: last[col['Stage']],
    sentBy: col['Sent By'] !== undefined ? String(last[col['Sent By']] || '') : '',
    sentAt: sentAt instanceof Date ? sentAt.toISOString() : String(sentAt || '')
  };
}

// Team address used for BCC on every scheduled guest email below, and as
// the direct recipient for the missing-info alert.
var TEAM_ALERT_EMAILS = 'chinnoos.pr@gmail.com,chandusrinivasan@yahoo.co.in';

// A booking matched one of the scheduled emails below (right source,
// right status, right date) but is missing something needed to actually
// send it — most commonly no Guest Email on file. Rather than silently
// skipping it (the old behavior — the booking would just vanish with no
// email and no trace), this tells the team directly so someone can add
// the missing field and, if needed, send it by hand.
function alertMissingGuestInfo(emailLabel, b, reason) {
  MailApp.sendEmail({
    to: TEAM_ALERT_EMAILS,
    subject: '[Action needed] ' + emailLabel + ' not sent — Booking ' + (b.bookingNumber || '(no number)'),
    body: 'Booking ' + (b.bookingNumber || '(no number)') + ' (' + (b.guestName || 'unknown guest') + ') matched today\'s ' +
      emailLabel.toLowerCase() + ', but ' + reason + ', so nothing was sent.\n\n' +
      'Check-in: ' + (b.checkIn || '—') + '\nCheck-out: ' + (b.checkOut || '—') + '\n\n' +
      'Fix the booking in the Ledger, then run retrySendFailed() (or wait for it to catch it automatically if you\'ve scheduled it) — no need to resend the whole day\'s batch by hand.'
  });
  logFailedSend(emailLabel, b, reason);
}

// Logs a row to the "Failed Sends" tab for every booking alertMissingGuestInfo
// fires for, so retrySendFailed() has something targeted to retry later —
// only this specific booking/email pair, never the whole day's batch (which
// would risk double-emailing guests who already got theirs).
function logFailedSend(emailLabel, b, reason) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Failed Sends') || ss.insertSheet('Failed Sends');
  if (sheet.getLastRow() === 0) {
    var headers = ['ID', 'Booking ID', 'Booking Number', 'Guest Name', 'Email Type', 'Reason', 'Check-in', 'Check-out', 'Detected At', 'Resolved At'];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  var id = 'fs_' + new Date().getTime().toString(36) + Math.random().toString(36).slice(2, 7);
  sheet.appendRow([id, b.id, b.bookingNumber || '', b.guestName || '', emailLabel, reason, b.checkIn || '', b.checkOut || '', new Date(), '']);
  sheet.autoResizeColumns(1, 10);
}

// Retries every unresolved row in "Failed Sends" — safe to run as often as
// you like, since it only ever touches the specific booking/email pairs
// that previously failed and are still unresolved, never bookings that
// already succeeded. Run this manually (Apps Script editor → select →
// Run) any time after fixing a booking's missing info.
function retrySendFailed() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Failed Sends');
  if (!sheet) { Logger.log('No Failed Sends sheet yet — nothing to retry.'); return; }
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return;
  var headers = values[0];
  var col = {};
  headers.forEach(function (h, i) { col[h] = i; });

  var data = readAll();
  var bookingsById = {};
  data.bookings.forEach(function (b) { bookingsById[b.id] = b; });

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    if (row[col['Resolved At']]) continue;
    var b = bookingsById[row[col['Booking ID']]];
    if (!b || !b.guestEmail) continue; // still not fixed — leave unresolved, try again later

    var isConfirmed = b.source === 'Direct' && (b.status === 'Confirmed-Advance' || b.status === 'Confirmed-Paid');
    var msg = null;
    var emailLabel = row[col['Email Type']];
    if (emailLabel === 'Balance reminder' && b.status === 'Confirmed-Advance' && isConfirmed) {
      msg = buildPaymentReminderEmail(ss, b);
    } else if (emailLabel === 'Arrival guide' && isConfirmed) {
      msg = buildArrivalGuideEmail(b);
    } else if (emailLabel === 'Pre-checkout email' && isConfirmed) {
      msg = buildPreCheckoutEmail(b);
    } else if (emailLabel === 'Feedback email' && isConfirmed) {
      msg = buildFeedbackEmail(b);
    }
    // A status that's since moved on (e.g. Cancelled, or Confirmed-Paid for
    // a balance reminder) means there's nothing left to send for this row —
    // resolved below either way, since there's nothing more to retry.
    if (msg) {
      MailApp.sendEmail({ to: b.guestEmail, subject: msg.subject, body: msg.body, bcc: TEAM_ALERT_EMAILS });
    }
    sheet.getRange(i + 1, col['Resolved At'] + 1).setValue(new Date());
  }
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
    if (!b.guestEmail) { alertMissingGuestInfo('Balance reminder', b, 'it has no Guest Email'); return; }
    var msg = buildPaymentReminderEmail(ss, b);
    MailApp.sendEmail({ to: b.guestEmail, subject: msg.subject, body: msg.body, bcc: TEAM_ALERT_EMAILS });
  });
}

// Builds the balance-reminder subject/body for one booking — shared by
// sendPaymentReminders() (the daily scan) and retrySendFailed() (targeted
// resend), so the wording only lives in one place.
function buildPaymentReminderEmail(ss, b) {
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
  return { subject: 'Reminder: Balance due for your upcoming stay — ' + b.bookingNumber, body: body };
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

// Arrival-guide email for Direct bookings checking in tomorrow — check-in
// details, house rules, and the arrival guide, sent once the day before.
// Plain text, no attachment. Airbnb bookings are handled through Airbnb's
// own guest messaging, so this is Direct-only.
function sendArrivalGuideEmails() {
  var data = readAll();
  var tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  var tomorrowStr = Utilities.formatDate(tomorrow, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  data.bookings.forEach(function (b) {
    if (b.source !== 'Direct') return;
    if (b.status !== 'Confirmed-Advance' && b.status !== 'Confirmed-Paid') return;
    if (String(b.checkIn).slice(0, 10) !== tomorrowStr) return;
    if (!b.guestEmail) { alertMissingGuestInfo('Arrival guide', b, 'it has no Guest Email'); return; }
    var msg = buildArrivalGuideEmail(b);
    MailApp.sendEmail({ to: b.guestEmail, subject: msg.subject, body: msg.body, bcc: TEAM_ALERT_EMAILS });
  });
}

// Builds the arrival-guide subject/body for one booking — shared by
// sendArrivalGuideEmails() and retrySendFailed().
function buildArrivalGuideEmail(b) {
    var checkInLabel = Utilities.formatDate(new Date(b.checkIn + 'T00:00:00'), Session.getScriptTimeZone(), 'd MMM yyyy');
    var body = 'Hi ' + b.guestName + ',\n\n' +
      'Booking: ' + b.bookingNumber + '\n\n' +
      'Looking forward to hosting you tomorrow, ' + checkInLabel + '. A few details for your arrival:\n\n' +
      'CHECK-IN\n' +
      '- Check-in from 2:00 PM.\n' +
      '- You\'ll be looked after by Mr. Mansoor, our property manager — his number is +91 96567 71881. Feel free to reach him directly for anything during your stay.\n' +
      '- Every adult in your group needs a valid government-approved ID (driving licence, Aadhaar, etc.) — collected at check-in per Tourism Department rules.\n' +
      '- Let Mansoor know your breakfast preference at check-in; it\'s prepared by Swadish Catering, a local family-run kitchen.\n' +
      '- WiFi details are shared at check-in, and Mansoor will walk you through the house on arrival.\n\n' +
      'GETTING HERE\n' +
      'Sindooram Ecostay, opposite Thakadi Temple, P.O, Edava, Varkala, Kerala 695311\n' +
      'Maps: https://www.google.com/maps/search/?api=1&query=Sindooram+Ecostay+opposite+Thakadi+Temple+Edava+Varkala+Kerala+695311\n\n' +
      'GETTING AROUND\n' +
      'Want to rent a scooter for beach hopping? Contact Shiraz at Venad Bike Rentals OPS Pvt Ltd — +91 98956 82274.\n' +
      'Maps: https://www.google.com/maps/search/?api=1&query=PPRC%2B4Q8,+Bus+Stand+Rd,+Varkala,+Kerala,+695141\n\n' +
      'HOUSE RULES, BRIEFLY\n' +
      '- Waste goes in the 3 bins provided (recyclable, non-recyclable, general).\n' +
      '- Meals and beverages¹ stay in the dining area rather than the bedrooms.\n' +
      '- Please keep the kitchen clean after use.\n' +
      '- If you\'re not using a light, fan, or AC, please switch it off.\n' +
      '- The switch under the TV is shared between the TV, WiFi, and CCTV cameras — please leave it on at all times, even when you\'re not using the TV or WiFi.\n' +
      '- The villa is smoke-free; there\'s a designated smoking/vaping spot in the backyard.\n' +
      '- Shoes off at the entrance (poomugham).\n' +
      '- Rinse sand off at the well area, not in the toilets — sandy clothes get hand-washed there too, not in the machine.\n' +
      '- Quiet hours after 10 PM.\n' +
      '- Drugs² are not permitted on the property, no exceptions.\n' +
      '- Travelling with a pet? They stay in the covered outdoor kennel beside Thira, not inside the villa or common areas.\n\n' +
      'Legal Compliance\n' +
      '¹ Selling, serving, or consuming alcohol for anyone under the age of 23 is prohibited in Kerala under the Abkari Act.\n' +
      '² Possession, sale, or consumption of drugs is prohibited in India for all ages under the NDPS Act.\n\n' +
      'For more on things to do or places to see, visit https://sindooramecostays.com\n\n' +
      'Anything before you arrive, WhatsApp us: https://wa.me/919846022350\n\n' +
      'Warm regards,\nTeam Sindooram';
    return { subject: 'Your stay at Sindooram Ecostay starts tomorrow — arrival details inside', body: body };
}

// Run this once yourself to schedule sendArrivalGuideEmails() daily at
// 9am, alongside the balance-due reminder. Re-running it is safe — it
// clears any previous schedule for this function first.
function createArrivalGuideTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sendArrivalGuideEmails') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendArrivalGuideEmails').timeBased().everyDays(1).atHour(9).create();
}

// Pre-checkout email for Direct bookings checking out tomorrow — a "how's
// the stay been" check-in plus checkout details, sent the day before.
// Plain text, no attachment. Direct-only, same reasoning as the arrival
// guide (Airbnb bookings are handled through Airbnb's own messaging).
function sendPreCheckoutEmails() {
  var data = readAll();
  var tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  var tomorrowStr = Utilities.formatDate(tomorrow, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  data.bookings.forEach(function (b) {
    if (b.source !== 'Direct') return;
    if (b.status !== 'Confirmed-Advance' && b.status !== 'Confirmed-Paid') return;
    if (String(b.checkOut).slice(0, 10) !== tomorrowStr) return;
    if (!b.guestEmail) { alertMissingGuestInfo('Pre-checkout email', b, 'it has no Guest Email'); return; }
    var msg = buildPreCheckoutEmail(b);
    MailApp.sendEmail({ to: b.guestEmail, subject: msg.subject, body: msg.body, bcc: TEAM_ALERT_EMAILS });
  });
}

// Builds the pre-checkout subject/body for one booking — shared by
// sendPreCheckoutEmails() and retrySendFailed().
function buildPreCheckoutEmail(b) {
    var checkOutLabel = Utilities.formatDate(new Date(b.checkOut + 'T00:00:00'), Session.getScriptTimeZone(), 'd MMM yyyy');
    var body = 'Hi ' + b.guestName + ',\n\n' +
      'Booking: ' + b.bookingNumber + '\n\n' +
      'Hope the stay\'s been a good one so far. You\'re checking out tomorrow, ' + checkOutLabel + ' — a few details for the morning:\n\n' +
      'CHECK-OUT\n' +
      '- By 11:00 AM.\n' +
      '- Leave the keys with Mansoor, or wherever he\'s asked you to.\n\n' +
      'BEFORE YOU GO\n' +
      '- Wet clothes/towels (including the thorthu) and anything sandy go in the laundry bag in the kitchen.\n' +
      '- Carrying wet or sandy clothes home instead? Ask Mansoor for a take-away laundry bag — available on request.\n' +
      '- Please wash and rack any dishes used.\n' +
      '- Segregate garbage/kitchen waste into the 3 bins in the kitchen.\n' +
      '- Turn off AC, fans, and lights in all rooms — except the switch under the TV (shared with WiFi and CCTV), which should stay on.\n\n' +
      'Anything not quite right, or need anything before you leave — let Mansoor know, or reach us on WhatsApp: https://wa.me/919846022350\n\n' +
      'Hope you\'ve had a good stay.\n\n' +
      'Warm regards,\nTeam Sindooram';
    return { subject: 'Checking out tomorrow — a few details', body: body };
}

// Run this once yourself to schedule sendPreCheckoutEmails() daily at 9am,
// alongside the other reminders. Re-running it is safe — it clears any
// previous schedule for this function first.
function createPreCheckoutTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sendPreCheckoutEmails') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendPreCheckoutEmails').timeBased().everyDays(1).atHour(9).create();
}

// Feedback email for Direct bookings that checked out yesterday — thanks
// the guest and asks for a Google review. Direct-only: an Airbnb-booked
// guest has no Airbnb reservation to attach a review to, so that ask
// doesn't apply here. Plain text, no attachment.
function sendFeedbackEmails() {
  var data = readAll();
  var yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  var yesterdayStr = Utilities.formatDate(yesterday, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  data.bookings.forEach(function (b) {
    if (b.source !== 'Direct') return;
    if (b.status !== 'Confirmed-Advance' && b.status !== 'Confirmed-Paid') return;
    if (String(b.checkOut).slice(0, 10) !== yesterdayStr) return;
    if (!b.guestEmail) { alertMissingGuestInfo('Feedback email', b, 'it has no Guest Email'); return; }
    var msg = buildFeedbackEmail(b);
    MailApp.sendEmail({ to: b.guestEmail, subject: msg.subject, body: msg.body, bcc: TEAM_ALERT_EMAILS });
  });
}

// Builds the feedback/review-ask subject/body for one booking — shared by
// sendFeedbackEmails() and retrySendFailed().
function buildFeedbackEmail(b) {
    var body = 'Hi ' + b.guestName + ',\n\n' +
      'Booking: ' + b.bookingNumber + '\n\n' +
      'Thank you for staying at Sindooram Ecostay — it was a pleasure having you, and we hope you had a good time.\n\n' +
      'If you enjoyed your stay, the one thing that would mean the world to us is a quick Google review — for a small, family-run place like ours, it genuinely makes a difference: https://g.page/r/CWPQRpdJLtFuEBM/review\n\n' +
      'We\'d also love to feature your stay on our socials — if you\'re up for it, share a few pictures with us on WhatsApp: https://wa.me/919846022350 (we make stunning reels 😂). And do follow us on Instagram (https://www.instagram.com/sindooramecostay) and Facebook (https://www.facebook.com/share/14roNs3DeBc/) so we can tag you.\n\n' +
      'If anything wasn\'t quite right, we\'d rather hear it directly — just reply to this email or WhatsApp us.\n\n' +
      'Warm regards,\nTeam Sindooram';
    return { subject: 'Thank you for staying with us', body: body };
}

// Run this once yourself to schedule sendFeedbackEmails() daily at 9am,
// alongside the other reminders. Re-running it is safe — it clears any
// previous schedule for this function first.
function createFeedbackTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sendFeedbackEmails') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendFeedbackEmails').timeBased().everyDays(1).atHour(9).create();
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
