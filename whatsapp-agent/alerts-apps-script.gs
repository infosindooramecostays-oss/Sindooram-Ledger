// ===== Sindooram WhatsApp Alerts — Apps Script =====
// A small standalone Web App the WhatsApp function calls whenever Maya hands
// a chat over to a person. It emails the alert list and logs the escalation
// to a "WhatsApp Escalations" tab in the Ledger Sheet.
//
// Script Properties (Project Settings → Script Properties) — never in code,
// this file is public on GitHub:
//   ALERT_SECRET  the same random string as ALERT_SECRET in Netlify
//   ALERT_EMAILS  who gets the emails, comma-separated
//   LOG_SHEET_ID  the Ledger Sheet's ID (optional — skip to email only)

function doPost(e) {
  try {
    var props = PropertiesService.getScriptProperties();
    var data = JSON.parse(e.postData.contents);
    if (!data.secret || data.secret !== props.getProperty('ALERT_SECRET')) return out({ error: 'forbidden' });
    sendAlert(data, props);
    return out({ status: 'ok' });
  } catch (err) {
    return out({ error: String(err.message || err) });
  }
}

function sendAlert(data, props) {
  var phone = String(data.phone || '').replace(/[^0-9]/g, '');
  var who = data.name ? data.name + ' (+' + phone + ')' : '+' + phone;
  var messages = data.messages || [];
  var transcript = messages.map(function (m) {
    return (m.role === 'user' ? 'Guest: ' : 'Maya: ') + m.content;
  }).join('\n\n');

  MailApp.sendEmail({
    to: props.getProperty('ALERT_EMAILS'),
    subject: '[WhatsApp] ' + who + ' — ' + String(data.reason || '').slice(0, 80),
    body: 'Maya has handed this WhatsApp chat to you.\n\n' +
      'Guest: ' + who + '\n' +
      'Reply on WhatsApp: https://wa.me/' + phone + '\n\n' +
      'Why: ' + data.reason + '\n\n' +
      'Recent messages:\n\n' + (transcript || '(none)')
  });

  var sheetId = props.getProperty('LOG_SHEET_ID');
  if (!sheetId) return;
  var ss = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheetByName('WhatsApp Escalations') || ss.insertSheet('WhatsApp Escalations');
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Time', 'Guest', 'Number', 'Reason', 'Last guest message', 'Maya replied']);
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
  }
  var lastGuest = messages.filter(function (m) { return m.role === 'user'; }).pop();
  var lastMaya = messages.filter(function (m) { return m.role === 'assistant'; }).pop();
  // Leading apostrophe keeps the number as text (no scientific notation).
  sheet.appendRow([new Date(), data.name || '', "'+" + phone, data.reason || '',
    lastGuest ? lastGuest.content : '', lastMaya ? lastMaya.content : '']);
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Run once from the editor: grants the email/Sheet permissions and sends a
// sample alert, so you can see what Negin will receive.
function testAlert() {
  sendAlert({
    phone: '910000000000', name: 'Test Guest', reason: 'price for 15–18 Nov, 6 adults (TEST)',
    messages: [
      { role: 'user', content: 'Hi, is 15–18 Nov available for 6 adults? What is the price?' },
      { role: 'assistant', content: 'Breakfast is included 🍳 Let me get Negin to share the exact price for 15–18 Nov, 6 adults — he will message you shortly! 🙌' }
    ]
  }, PropertiesService.getScriptProperties());
  Logger.log('Test alert sent to ' + PropertiesService.getScriptProperties().getProperty('ALERT_EMAILS'));
}
