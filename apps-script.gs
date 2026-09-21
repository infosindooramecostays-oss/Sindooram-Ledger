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

var LOGO_URL = 'https://infosindooramecostays-oss.github.io/Sindooram-Ledger/assets/brand/sindooram-logo.jpeg';
// The exact "Payment receipt" folder the user already created — using its
// ID (from the folder's own URL) instead of searching by name, since a
// name search only looks at the top level of My Drive and this folder
// isn't there.
var RECEIPTS_FOLDER_ID = '1bMOGMDEEn9aaP5p_Z1wBbsoUNdMEOkrg';
var RECEIPT_FONT = 'Arial';

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
      balanceLines: ['- Full payment received.', 'Booking confirmed and complete — no balance due.'],
      nonRefundable: 'All payments received are non-refundable.'
    };
  }
  var remaining = Math.max(0, Number(input.totalAmount) - Number(input.amountReceived));
  var dueDate = new Date(input.checkIn + 'T00:00:00');
  dueDate.setDate(dueDate.getDate() - 1);
  var dueDateStr = Utilities.formatDate(dueDate, Session.getScriptTimeZone(), 'd MMM yyyy');
  return {
    statusLine: 'Your booking has been confirmed, pending the remaining balance.',
    balanceLines: ['- Booking confirmed, balance of Rs ' + formatMoney(remaining) + ',', 'due one day before check-in, i.e. ' + dueDateStr + '.'],
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
  valueCell.setBackgroundColor('#FBD3C5');
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
    var logoBlob = UrlFetchApp.fetch(LOGO_URL).getBlob();
    var blankPara = logoCell.getChild(0).asParagraph();
    var img = logoCell.appendImage(logoBlob);
    blankPara.removeFromParent();
    var ratio = img.getHeight() / img.getWidth();
    img.setWidth(150);
    img.setHeight(Math.round(150 * ratio));
  } catch (imgErr) {
    Logger.log('Logo fetch failed: ' + imgErr);
    var logoFallback = logoCell.getChild(0).asParagraph();
    logoFallback.setText('Sindooram Ecostay (logo error: ' + imgErr.message + ')');
    styleParaText(logoFallback, { bold: true, size: 8 });
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
  balanceCell.setBackgroundColor('#FBD3C5');
  balanceCell.setPaddingTop(8).setPaddingBottom(8).setPaddingLeft(10).setPaddingRight(10);
  content.balanceLines.forEach(function (line, i) {
    var p = i === 0 ? balanceCell.getChild(0).asParagraph() : balanceCell.appendParagraph('');
    p.setText(line);
    styleParaText(p, { size: 10 });
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
  styleParaText(body.appendParagraph('Regards,'), { size: 10 });
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
