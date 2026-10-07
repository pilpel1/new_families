/**
 * טופס גוגל → הודעה לטלגרם עם כפתור וואטסאפ.
 *
 * סודות (לא בקובץ הזה):
 *   Project Settings → Script properties
 *   TELEGRAM_TOKEN
 *   CHAT_ID
 *
 * הבוט לא עונה למי שכותב לו. הוא רק שולח אל CHAT_ID.
 * אין כאן webhook ואין תשובות נכנסות.
 *
 * משנים רק את CONFIG, ואז מריצים setup ואחר כך installTrigger.
 * ברשימת הפונקציות לא מריצים handleFormSubmit ידנית.
 */

var CONFIG = {
  // השמות חייבים להיות זהים לכותרות השאלות בטופס, כולל רווחים ונקודתיים.
  NAME_QUESTION: 'שם המשפחה שלכם:',
  PHONE_QUESTION: 'מספר טלפון נייד של אחד מכם:',

  // {{name}} מוחלף בתשובה של NAME_QUESTION. הנקודתיים נשארות רק בכותרת השאלה.
  WA_TEMPLATE: [
    'אהלן, משפחת {{name}} יקרים!',
    'תודה שעניתם על השאלון שלנו, ותודה שאתם עוזרים לנו לשכלל את הקהילה.',
    'תמיד תוכלו להוסיף עוד בעריכת השאלון, או לשלוח אליי בפרטי.',
    'שמחים שהצטרפתם, תהנו מכל רגע!',
    'וועדת משפחות חדשות'
  ].join('\n')
};

var LOG_SHEET_NAME = 'TelegramLog';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('טלגרם')
    .addItem('בדיקת חיבור', 'setup')
    .addItem('התקנת טריגר', 'installTrigger')
    .addItem('שליחת בדיקה', 'testWithSample')
    .addItem('בדיקת טלפון שגוי', 'testBadPhone')
    .addItem('שליחה מחדש של השורה האחרונה', 'resendLast')
    .addToUi();
}

/** פעם אחת: מוחק webhook (כדי שהבוט ישתוק) ושולח הודעת בדיקה. */
function setup() {
  var cfg = props_();
  telegram_(cfg.token, 'deleteWebhook', { drop_pending_updates: true });
  var me = telegram_(cfg.token, 'getMe', {});
  var username = me.result && me.result.username ? '@' + me.result.username : 'הבוט';
  postMessage_(
    cfg,
    username +
      ' מחובר.\n' +
      'הודעות מהטופס יגיעו רק לכאן.\n' +
      'מי שפותח את הבוט וכותב לו לא מקבל תשובה.',
    '',
    '',
    ''
  );
  toast_('הבוט מחובר. הודעת בדיקה נשלחה לטלגרם.');
}

/** פעם אחת: טריגר על שליחת טופס. לא לקרוא לפונקציה onFormSubmit. */
function installTrigger() {
  var ss = SpreadsheetApp.getActive();
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    var name = trigger.getHandlerFunction();
    if (name === 'handleFormSubmit' || name === 'onFormSubmit') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  ScriptApp.newTrigger('handleFormSubmit').forSpreadsheet(ss).onFormSubmit().create();
  toast_('הטריגר הותקן. מלאו את הטופס כדי לבדוק.');
}

function testWithSample() {
  var named = { 'חותמת זמן': ['בדיקה'] };
  named[CONFIG.NAME_QUESTION] = ['ישראל ישראלי'];
  named[CONFIG.PHONE_QUESTION] = ['0501234567'];
  named['הערה'] = ['זו הודעת בדיקה, אפשר למחוק אותה בטלגרם'];
  processSubmission_(fakeEvent_(named), { skipDedupe: true });
  toast_('נשלחה הודעת בדיקה לטלגרם.');
}

function testBadPhone() {
  var named = { 'חותמת זמן': ['בדיקה'] };
  named[CONFIG.NAME_QUESTION] = ['ישראל ישראלי'];
  named[CONFIG.PHONE_QUESTION] = ['12345'];
  processSubmission_(fakeEvent_(named), { skipDedupe: true });
  toast_('נשלחה לטלגרם דוגמה של שגיאת טלפון.');
}

/** אחרי תיקון שורה בגיליון (או תיקון שמות השאלות ב-CONFIG). */
function resendLast() {
  var sh = responseSheet_();
  var row = sh.getLastRow();
  if (row < 2) {
    throw new Error('אין תשובות בגיליון.');
  }
  processSubmission_(eventFromRow_(sh, row), { skipDedupe: true });
  toast_('שורה ' + row + ' נשלחה לטלגרם.');
}

function handleFormSubmit(e) {
  if (!e || !e.namedValues) {
    throw new Error(
      'אין כאן תשובה של טופס. לא מריצים את handleFormSubmit ידנית. מריצים testWithSample, או ממלאים את הטופס.'
    );
  }
  try {
    processSubmission_(e, {});
  } catch (err) {
    if (!safeReport_(e, err)) {
      throw new Error(scrub_(err.message));
    }
  }
}

function processSubmission_(e, opt) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    processSubmissionUnlocked_(e, opt || {});
  } finally {
    lock.releaseLock();
  }
}

function processSubmissionUnlocked_(e, opt) {
  var cfg = props_();
  var sub = collect_(e);
  var key = submissionKey_(e);

  if (!opt.skipDedupe && alreadySent_(key)) {
    console.log('כבר נשלח, דילוג. שורה ' + sub.row);
    return;
  }

  var problems = [];
  var name = '';
  var phone = '';
  var draft = '';
  var keys = questionKeys_(sub);

  if (!hasQuestion_(sub.named, CONFIG.NAME_QUESTION)) {
    problems.push(
      'לא נמצאה שאלה בשם "' +
        CONFIG.NAME_QUESTION +
        '". השאלות שהגיעו: ' +
        (keys.join(', ') || 'אין')
    );
  } else {
    name = answer_(sub.named, CONFIG.NAME_QUESTION);
    if (!name) problems.push('השאלה "' + CONFIG.NAME_QUESTION + '" הגיעה ריקה.');
  }

  if (!hasQuestion_(sub.named, CONFIG.PHONE_QUESTION)) {
    problems.push(
      'לא נמצאה שאלה בשם "' +
        CONFIG.PHONE_QUESTION +
        '". השאלות שהגיעו: ' +
        (keys.join(', ') || 'אין')
    );
  } else {
    var rawPhone = answer_(sub.named, CONFIG.PHONE_QUESTION);
    phone = toWhatsAppPhone(rawPhone);
    if (!phone) {
      problems.push(
        'מספר הטלפון לא זוהה כנייד ישראלי (05 או 07). הערך שהתקבל: "' +
          (rawPhone || 'ריק') +
          '"'
      );
    }
  }

  if (name) {
    var applied = applyTemplate_(CONFIG.WA_TEMPLATE, sub.named, name);
    draft = applied.text;
    if (applied.missing.length) {
      problems.push(
        'בתבנית הוואטסאפ נשארו שדות שלא קיימים בטופס: ' + applied.missing.join(', ')
      );
    }
  }

  if (problems.length) {
    deliver_(cfg, renderProblem_(problems, sub, draft), '', '');
    markSent_(key, 'warning', problems.join(' | '));
    console.log('אזהרה בשורה ' + sub.row + ': ' + scrub_(problems.join(' | ')));
    return;
  }

  var wa = buildWaUrl_(phone, draft);
  deliver_(cfg, renderOk_(sub, wa), wa.url, buttonLabel_(name));
  markSent_(key, 'ok', 'נשלח');
  console.log('נשלח. שורה ' + sub.row);
}

function renderProblem_(problems, sub, draft) {
  var lines = [{ value: 'מה לא הצליח:' }];
  for (var i = 0; i < problems.length; i++) {
    lines.push({ label: i + 1 + '.', value: problems[i] });
  }
  lines.push({ value: 'פרטים שנכנסו:' });
  lines = lines.concat(fieldLines_(sub));
  lines.push({
    value: 'השורה כבר שמורה בגיליון. אחרי תיקון מריצים resendLast (או מהתפריט: שליחה מחדש של השורה האחרונה).'
  });
  var pre = draft ? { label: 'טיוטה (בלי קישור):', value: draft } : null;
  return render_('⚠️ הטופס נשמר בגיליון, וואטסאפ לא הוכן', lines, pre);
}

function renderOk_(sub, wa) {
  var lines = fieldLines_(sub);
  if (wa.trimmed) {
    lines.push({
      value: 'הטיוטה בקישור קוצרה כי היא ארוכה מדי. למטה מופיע הנוסח המלא.'
    });
  }
  return render_('📩 התקבלה קליטה חדשה', lines, {
    label: 'טיוטה לוואטסאפ:',
    value: wa.fullText
  });
}

function render_(title, lines, preBlock) {
  var html = [esc_(title)];
  var plain = [title];

  for (var i = 0; i < lines.length; i++) {
    var line = lines[i];
    if (line.label) {
      html.push('<b>' + esc_(line.label) + '</b> ' + esc_(line.value));
      plain.push(line.label + ' ' + line.value);
    } else {
      html.push(esc_(line.value));
      plain.push(line.value);
    }
  }

  if (preBlock) {
    html.push('<b>' + esc_(preBlock.label) + '</b>\n<pre>' + esc_(preBlock.value) + '</pre>');
    plain.push(preBlock.label + '\n' + preBlock.value);
  }

  return {
    html: clip_(html.join('\n'), 3900),
    plain: clip_(plain.join('\n'), 3900)
  };
}

function fieldLines_(sub) {
  var lines = [];
  if (sub.row >= 2) lines.push({ label: 'שורה בגיליון:', value: String(sub.row) });
  for (var i = 0; i < sub.fields.length; i++) {
    lines.push({
      label: sub.fields[i].key + ':',
      value: sub.fields[i].value || '—'
    });
  }
  return lines;
}

function deliver_(cfg, rendered, waUrl, buttonText) {
  try {
    postMessage_(cfg, rendered.html, 'HTML', waUrl, buttonText);
  } catch (err) {
    var extra = waUrl ? '\n' + waUrl : '';
    postMessage_(
      cfg,
      clip_(
        'ההודעה נשלחה בלי עיצוב, כי הניסיון הראשון נדחה.\n' +
          scrub_(err.message) +
          '\n\n' +
          rendered.plain +
          extra,
        4000
      ),
      '',
      '',
      ''
    );
  }
}

function safeReport_(e, err) {
  try {
    var cfg = props_();
    var sub = collect_(e || {});
    var lines = [
      { value: 'מה לא הצליח:' },
      { value: scrub_(err && err.message ? err.message : err) },
      { value: 'פרטים שנכנסו:' }
    ].concat(fieldLines_(sub));
    lines.push({
      value: 'אחרי שטלגרם חוזר לעבוד, מריצים resendLast כדי לשלוח שוב את השורה האחרונה.'
    });
    var rendered = render_('⚠️ השליחה נכשלה', lines, null);
    postMessage_(cfg, rendered.plain, '', '', '');
    return true;
  } catch (reportErr) {
    console.error(scrub_(reportErr.message));
    return false;
  }
}

function postMessage_(cfg, text, parseMode, waUrl, buttonText) {
  var body = {
    chat_id: cfg.chatId,
    text: text
  };
  if (parseMode) body.parse_mode = parseMode;
  if (waUrl) {
    body.reply_markup = {
      inline_keyboard: [[{ text: buttonText || '📲 פתח וואטסאפ', url: waUrl }]]
    };
  }
  telegram_(cfg.token, 'sendMessage', body);
}

function telegram_(token, method, payload) {
  var res;
  try {
    res = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/' + method, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload || {}),
      muteHttpExceptions: true
    });
  } catch (err) {
    throw new Error(method + ' לא הגיע לטלגרם. ' + scrub_(err.message));
  }

  var raw = res.getContentText();
  var json = {};
  try {
    json = JSON.parse(raw);
  } catch (parseErr) {
    json = { ok: false, description: raw };
  }
  if (res.getResponseCode() !== 200 || !json.ok) {
    var desc = json && json.description ? json.description : raw;
    var hint = hintForTelegram_(desc);
    throw new Error(method + ' נכשל (' + res.getResponseCode() + '): ' + desc + (hint ? '\n' + hint : ''));
  }
  return json;
}

function hintForTelegram_(message) {
  var text = String(message || '');
  if (text.indexOf('chat not found') !== -1) {
    return 'פותחים את הבוט בטלגרם ולוחצים Start פעם אחת (הוא לא יענה), ובודקים ש-CHAT_ID הוא המספר של הצ\'אט הזה.';
  }
  if (text.indexOf('blocked') !== -1) {
    return 'הבוט חסום בטלגרם. מבטלים את החסימה ולוחצים Start.';
  }
  return '';
}

function props_() {
  var stored = PropertiesService.getScriptProperties();
  var token = String(stored.getProperty('TELEGRAM_TOKEN') || '').trim();
  var chatId = String(stored.getProperty('CHAT_ID') || '').trim();
  if (!token || !chatId) {
    throw new Error(
      'חסרים TELEGRAM_TOKEN או CHAT_ID. מוסיפים אותם ב-Project Settings → Script properties. לא בתוך הקוד.'
    );
  }
  if (!/^-?\d+$/.test(chatId)) {
    throw new Error('CHAT_ID צריך להיות מספר (למשל 123456789), לא טוקן ולא שם משתמש.');
  }
  if (token.indexOf(':') === -1) {
    throw new Error('TELEGRAM_TOKEN לא נראה כמו טוקן של בוט. הוא אמור להכיל נקודתיים.');
  }
  return { token: token, chatId: chatId };
}

function toWhatsAppPhone(raw) {
  var digits = String(raw || '').replace(/\D/g, '');
  if (digits.indexOf('00972') === 0) digits = digits.slice(2);
  if (digits.indexOf('9720') === 0) digits = '972' + digits.slice(4);
  if (digits.indexOf('0') === 0) digits = '972' + digits.slice(1);
  if (/^[57]\d{8}$/.test(digits)) digits = '972' + digits;
  if (!/^972[57]\d{8}$/.test(digits)) return '';
  return digits;
}

function applyTemplate_(template, named, name) {
  var text = String(template).split('{{name}}').join(name);
  var keys = Object.keys(named);
  for (var i = 0; i < keys.length; i++) {
    text = text.split('{{' + keys[i] + '}}').join(answer_(named, keys[i]));
  }
  var missing = text.match(/\{\{[^}]+\}\}/g);
  return { text: text, missing: missing || [] };
}

function buildWaUrl_(phone, text) {
  var body = text;
  var url = waUrl_(phone, body);
  var trimmed = false;
  while (url.length > 1900 && body.length > 40) {
    trimmed = true;
    body = body.slice(0, body.length - 40).trim() + '…';
    url = waUrl_(phone, body);
  }
  return { url: url, text: body, fullText: text, trimmed: trimmed };
}

function waUrl_(phone, text) {
  return 'https://wa.me/' + phone + '?text=' + encodeURIComponent(text);
}

function buttonLabel_(name) {
  var clean = String(name || '').replace(/[<>&]/g, '').trim();
  if (clean.length > 20) clean = clean.slice(0, 20) + '…';
  return clean ? '📲 וואטסאפ אל ' + clean : '📲 פתח וואטסאפ';
}

function collect_(e) {
  var named = e && e.namedValues ? e.namedValues : {};
  var row = e && e.range && e.range.getRow ? e.range.getRow() : 0;
  var keys = Object.keys(named);
  var fields = [];
  for (var i = 0; i < keys.length; i++) {
    fields.push({ key: keys[i], value: answer_(named, keys[i]) });
  }
  return { named: named, row: row, fields: fields };
}

function questionKeys_(sub) {
  var keys = [];
  for (var i = 0; i < sub.fields.length; i++) {
    var key = sub.fields[i].key;
    if (key === 'חותמת זמן' || key === 'Timestamp') continue;
    keys.push(key);
  }
  return keys;
}

function hasQuestion_(named, key) {
  return Object.prototype.hasOwnProperty.call(named, key);
}

function answer_(named, key) {
  if (!hasQuestion_(named, key)) return '';
  var value = named[key];
  var parts = Array.isArray(value) ? value : [value];
  var out = [];
  for (var i = 0; i < parts.length; i++) {
    var item = parts[i];
    if (item == null || item === '') continue;
    if (Object.prototype.toString.call(item) === '[object Date]') {
      out.push(
        Utilities.formatDate(
          item,
          SpreadsheetApp.getActive().getSpreadsheetTimeZone(),
          'dd/MM/yyyy HH:mm'
        )
      );
    } else {
      out.push(String(item).trim());
    }
  }
  return out.join(', ');
}

function fakeEvent_(named) {
  return {
    namedValues: named,
    range: {
      getRow: function () {
        return 0;
      }
    }
  };
}

function eventFromRow_(sh, row) {
  var width = sh.getLastColumn();
  var headers = sh.getRange(1, 1, 1, width).getValues()[0];
  var values = sh.getRange(row, 1, 1, width).getValues()[0];
  var tz = sh.getParent().getSpreadsheetTimeZone();
  var named = {};
  for (var i = 0; i < headers.length; i++) {
    if (!headers[i]) continue;
    named[String(headers[i])] = [cellText_(values[i], tz)];
  }
  return {
    namedValues: named,
    range: sh.getRange(row, 1)
  };
}

function cellText_(value, tz) {
  if (value == null || value === '') return '';
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return Utilities.formatDate(value, tz, 'dd/MM/yyyy HH:mm');
  }
  return String(value).trim();
}

function responseSheet_() {
  var sheets = SpreadsheetApp.getActive().getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getFormUrl()) return sheets[i];
  }
  throw new Error('לא מצאתי גיליון שמחובר לטופס. מדביקים את הסקריפט בגיליון התשובות של הטופס, לא בעותק.');
}

function logSheet_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(LOG_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(LOG_SHEET_NAME);
    sh.appendRow(['key', 'status', 'detail', 'sentAt']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function alreadySent_(key) {
  var sh = logSheet_();
  var last = sh.getLastRow();
  if (last < 2) return false;
  var values = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (values[i][0] === key) return true;
  }
  return false;
}

function markSent_(key, status, detail) {
  logSheet_().appendRow([key, status, String(detail).slice(0, 500), new Date()]);
}

function submissionKey_(e) {
  var row = e.range && e.range.getRow ? e.range.getRow() : 0;
  return String(row) + ':' + digest_(JSON.stringify(e.namedValues || {}));
}

function digest_(text) {
  var bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    text,
    Utilities.Charset.UTF_8
  );
  var hex = '';
  for (var i = 0; i < bytes.length; i++) {
    var value = (bytes[i] + 256) % 256;
    hex += (value < 16 ? '0' : '') + value.toString(16);
  }
  return hex;
}

function esc_(text) {
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function clip_(text, max) {
  var value = String(text);
  if (value.length <= max) return value;
  return value.slice(0, max - 30).replace(/\s+\S*$/, '') + '\n… (קוצר)';
}

function scrub_(text) {
  return String(text).replace(/bot\d{6,}:[A-Za-z0-9_-]+/gi, 'bot***');
}

function toast_(message) {
  try {
    SpreadsheetApp.getActive().toast(message, 'טלגרם', 8);
  } catch (err) {
    console.log(message);
  }
}
