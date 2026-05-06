// ===== КОНФИГ =====
var SPREADSHEET_ID = '13Ue0WwAkXE7NGY-Mrn7LCB6HIlI41df2pUgu86wMYjg';
var SHEET_NAME = ''; // пусто = первый лист (Answers)

var HEADERS = ['timestamp', 'email', 'name', 'topic', 'details'];

// Человеческие имена тем
var TOPIC_LABELS = {
  alpha: 'Alpha',
  invoice: 'Invoice',
  payment: 'Payment details change',
  other: 'Other question'
};

// Человеческие имена полей в details
var FIELD_LABELS = {
  alpha_primary_type: 'Alpha type',
  invoice_issue_type: 'Issue type',
  invoice_description: 'Description',
  payment_method: 'Payment method',
  payment_paypal_email: 'PayPal email',
  payment_skrill_email: 'Skrill email',
  payment_payoneer_email: 'Payoneer email',
  payment_wire_beneficiary_name: 'Wire — beneficiary',
  payment_wire_bank_name: 'Wire — bank',
  payment_wire_swift: 'Wire — SWIFT',
  payment_wire_routing_number: 'Wire — routing (ABA)',
  payment_wire_account_number: 'Wire — account',
  payment_owner_confirm: 'Owner confirmed',
  payment_additional_info: 'Additional info',
  other_question_text: 'Question'
};

// Что НЕ показывать в details (они идут в отдельные колонки или служебные)
var SKIP_IN_DETAILS = {
  your_email: true,
  your_name: true,
  main_topic: true,
  timestamp: true,
  submitted_at: true
};


// ===== MAIN =====

function doPost(e) {
  try {
    var payload = parsePayload_(e);
    if (!payload || Object.keys(payload).length === 0) {
      return jsonOut_('error', 'no data');
    }

    if (payload.website) {
      return jsonOut_('ok');
    }

    var errors = validatePayload_(payload);
    if (errors.length > 0) {
      return jsonOut_('error', errors.join('; '));
    }

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = SHEET_NAME ? ss.getSheetByName(SHEET_NAME) : ss.getActiveSheet();
    if (!sheet) return jsonOut_('error', 'sheet not found');

    ensureHeaders_(sheet);

    var topicRaw = payload.main_topic || '';
    var topic = TOPIC_LABELS[topicRaw] || topicRaw || '—';
    var email = payload.your_email || '';
    var name = payload.your_name || '';
    var details = formatDetails_(payload);

    var row = [new Date(), email, name, topic, details];
    sheet.appendRow(row);

    // Перенос по словам в details (5-я колонка)
    var lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 5).setWrap(true).setVerticalAlignment('top');

    try { sendConfirmationEmail_(payload, details); } catch (mailErr) {
      Logger.log('Email error: ' + mailErr);
    }

    return jsonOut_('ok');
  } catch (err) {
    return jsonOut_('error', err.toString());
  }
}

function doGet() {
  return ContentService.createTextOutput('Form endpoint is active. Use POST.')
    .setMimeType(ContentService.MimeType.TEXT);
}


// ===== HELPERS =====

function parsePayload_(e) {
  var payload = {};
  if (e && e.postData && e.postData.contents) {
    var raw = e.postData.contents.trim();
    var ct = (e.postData.type || '').toLowerCase();

    if (ct.indexOf('application/x-www-form-urlencoded') !== -1) {
      var parts = raw.split('&');
      for (var i = 0; i < parts.length; i++) {
        var pair = parts[i];
        var eq = pair.indexOf('=');
        if (eq < 0) continue;
        var k = decodeURIComponent(pair.substring(0, eq).replace(/\+/g, ' '));
        var v = decodeURIComponent(pair.substring(eq + 1).replace(/\+/g, ' '));
        if (k === 'payload') {
          try { payload = JSON.parse(v); } catch (er) {}
        }
      }
    }

    if (Object.keys(payload).length === 0) {
      try { payload = JSON.parse(raw); } catch (er) {}
    }
  }

  if (Object.keys(payload).length === 0 && e && e.parameter && e.parameter.payload) {
    try { payload = JSON.parse(e.parameter.payload); } catch (er) {}
  }

  return payload;
}

function ensureHeaders_(sheet) {
  var lastCol = sheet.getLastColumn();
  var firstCellEmpty = sheet.getLastRow() === 0 ||
    String(sheet.getRange(1, 1).getValue()).trim() === '';

  if (firstCellEmpty || lastCol < HEADERS.length) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 160); // timestamp
    sheet.setColumnWidth(2, 220); // email
    sheet.setColumnWidth(3, 180); // name
    sheet.setColumnWidth(4, 180); // topic
    sheet.setColumnWidth(5, 520); // details
  }
}

function formatDetails_(payload) {
  var keys = Object.keys(payload);
  var ordered = [];
  for (var k in FIELD_LABELS) {
    if (payload.hasOwnProperty(k)) ordered.push(k);
  }
  for (var i = 0; i < keys.length; i++) {
    if (!FIELD_LABELS.hasOwnProperty(keys[i]) && !SKIP_IN_DETAILS[keys[i]]) {
      ordered.push(keys[i]);
    }
  }

  var lines = [];
  for (var j = 0; j < ordered.length; j++) {
    var key = ordered[j];
    if (SKIP_IN_DETAILS[key]) continue;
    var val = payload[key];
    if (val === null || val === undefined) continue;
    var str = (typeof val === 'object') ? JSON.stringify(val) : String(val);
    if (str.trim() === '') continue;

    var label = FIELD_LABELS[key] || key;
    lines.push(label + ': ' + str);
  }

  return lines.join('\n');
}

// Отправитель должен быть добавлен как псевдоним в настройках Gmail аккаунта,
// от имени которого развёрнут скрипт (Gmail → Settings → Accounts → Send mail as).
var CONFIRMATION_FROM    = 'noreplyvendors@alconost.com';
var CONFIRMATION_REPLYTO = 'vendors@alconost.com';

function sendConfirmationEmail_(payload, details) {
  var to = payload.your_email;
  if (!to) return;

  var topic = TOPIC_LABELS[payload.main_topic] || payload.main_topic || '—';

  var copyText = 'Name: '  + (payload.your_name  || '') + '\n' +
                 'Email: ' + (payload.your_email || '') + '\n' +
                 'Topic: ' + topic;
  if (details) copyText += '\n' + details;

  var plain = 'Hello,\n\n' +
    'Thank you for filling out the form. We\'ve successfully received your submission.\n\n' +
    'This is an automated email, so please do not reply to this message.\n\n' +
    'Here is a copy of the information you provided:\n\n' +
    copyText + '\n\n' +
    'Our team will review your request and get back to you as soon as possible.\n\n' +
    'If you have any additional questions, feel free to contact us at ' + CONFIRMATION_REPLYTO + '.\n\n' +
    'Best regards,\n' +
    'Vendor Management Team\n' +
    'alconost.com';

  var copyHtml = copyText.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/\n/g,'<br>');

  var html = '<p>Hello,</p>' +
    '<p>Thank you for filling out the form. We\'ve successfully received your submission.</p>' +
    '<p><em>This is an automated email, so please do not reply to this message.</em></p>' +
    '<p>Here is a copy of the information you provided:</p>' +
    '<p style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:6px;padding:12px 16px;font-family:monospace;font-size:13px;line-height:1.6">' + copyHtml + '</p>' +
    '<p>Our team will review your request and get back to you as soon as possible.</p>' +
    '<p>If you have any additional questions, feel free to contact us at ' +
    '<a href="mailto:' + CONFIRMATION_REPLYTO + '">' + CONFIRMATION_REPLYTO + '</a>.</p>' +
    '<p>Best regards,<br>Vendor Management Team<br>alconost.com</p>';

  GmailApp.sendEmail(to, 'We\'ve Received Your Submission', plain, {
    from:    CONFIRMATION_FROM,
    name:    'Vendor Management Team',
    replyTo: CONFIRMATION_REPLYTO,
    htmlBody: html
  });
}

function str_(v) {
  return (v !== null && v !== undefined && String(v).trim() !== '') ? String(v).trim() : null;
}

function validatePayload_(payload) {
  var errors = [];

  if (!str_(payload.your_name))  errors.push('your_name is required');
  if (!str_(payload.your_email)) errors.push('your_email is required');
  if (str_(payload.your_email) && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(payload.your_email)) {
    errors.push('your_email is invalid');
  }

  var topic = str_(payload.main_topic);
  if (!topic) {
    errors.push('main_topic is required');
    return errors;
  }

  if (topic === 'alpha') {
    if (!str_(payload.alpha_primary_type)) errors.push('alpha_primary_type is required');

  } else if (topic === 'invoice') {
    if (!str_(payload.invoice_issue_type)) errors.push('invoice_issue_type is required');
    if (payload.invoice_issue_type === 'other' && !str_(payload.invoice_description)) errors.push('invoice_description is required');

  } else if (topic === 'payment') {
    var method = str_(payload.payment_method);
    if (!method) {
      errors.push('payment_method is required');
    } else {
      if (method === 'paypal'   && !str_(payload.payment_paypal_email))   errors.push('payment_paypal_email is required');
      if (method === 'skrill'   && !str_(payload.payment_skrill_email))   errors.push('payment_skrill_email is required');
      if (method === 'payoneer' && !str_(payload.payment_payoneer_email)) errors.push('payment_payoneer_email is required');
      if (method === 'wire_transfer') {
        if (!str_(payload.payment_wire_beneficiary_name)) errors.push('payment_wire_beneficiary_name is required');
        if (!str_(payload.payment_wire_bank_name))        errors.push('payment_wire_bank_name is required');
        if (!str_(payload.payment_wire_account_number))   errors.push('payment_wire_account_number is required');
      }
      if (payload.payment_owner_confirm !== 'yes') errors.push('payment_owner_confirm is required');
    }

  } else if (topic === 'other') {
    if (!str_(payload.other_question_text)) errors.push('other_question_text is required');
  }

  return errors;
}

function textOut_(msg) {
  return ContentService.createTextOutput(msg)
    .setMimeType(ContentService.MimeType.TEXT);
}

function jsonOut_(status, message) {
  var obj = { status: status };
  if (message) obj.message = message;
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Запусти один раз вручную — пересоздаст шапку и почистит лист
function initializeHeaders() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = SHEET_NAME ? ss.getSheetByName(SHEET_NAME) : ss.getActiveSheet();
  sheet.clear();
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidth(2, 220);
  sheet.setColumnWidth(3, 180);
  sheet.setColumnWidth(4, 180);
  sheet.setColumnWidth(5, 520);
  return 'Headers reset. 5 columns: timestamp, email, name, topic, details';
}
