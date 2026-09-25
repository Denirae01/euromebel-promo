/**
 * EuroMebel · Промо «Испытай удачу» — бэкенд промокодов
 * Google Таблица + Apps Script (бесплатно, без своего сервера).
 *
 * Что делает:
 *  - принимает email с лендинга (POST от GitHub Pages);
 *  - выдаёт каждому email один свободный промокод из загруженного пула;
 *  - ставит заявку в очередь отправки;
 *  - отправляет письмо через ваш сервис (WEBHOOK_URL) — сразу или позже, пачкой;
 *  - даёт салонам проверить и погасить код (страница check.html).
 *
 * Настройки — в «Настройки проекта → Свойства скрипта» (см. README).
 */

// ---------- названия листов и колонок ----------
var SH_CODES = 'Промокоды';
var SH_LEADS = 'Заявки';
var SH_IMPORT = 'Загрузка';

var CODES_HEAD = ['Код', 'Статус', 'Email', 'Выдан', 'Погашен', 'Салон'];
var LEADS_HEAD = ['Дата', 'Email', 'Код', 'Кампания', 'Отправка', 'Отправлено', 'Ошибка', 'Страница', 'UTM'];

var ST_FREE = 'свободен', ST_ISSUED = 'выдан', ST_REDEEMED = 'погашен';
var SEND_PENDING = 'ожидает', SEND_OK = 'отправлено', SEND_ERR = 'ошибка';

// ============================================================
//  Настройки
// ============================================================
function cfg_() {
  var p = PropertiesService.getScriptProperties().getProperties();
  return {
    spreadsheetId: p.SPREADSHEET_ID || '',
    webhookUrl: p.WEBHOOK_URL || '',          // URL вашего сервиса рассылки
    webhookSecret: p.WEBHOOK_SECRET || '',    // передаётся в заголовке X-Promo-Secret
    sendOnClaim: p.SEND_ON_CLAIM !== 'false', // отправлять сразу при заявке
    sendViaGmail: p.SEND_VIA_GMAIL === 'true',// запасной вариант: письмо из Gmail владельца
    fromName: p.FROM_NAME || 'EuroMebel',
    replyTo: p.REPLY_TO || '',
    staffPin: p.STAFF_PIN || '',
    discount: p.DISCOUNT || '−20%',
    scope: p.SCOPE || 'на всю мягкую мебель',
    validUntil: p.VALID_UNTIL || '',
    catalogUrl: p.CATALOG_URL || 'https://euromebel.kz/'
  };
}

function ss_() {
  var id = cfg_().spreadsheetId;
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}

function sheet_(name, head) {
  var ss = ss_();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (head) {
      sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
  }
  return sh;
}

// ============================================================
//  Меню в таблице
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Промо')
    .addItem('1. Подготовить таблицу', 'setup')
    .addItem('2. Загрузить коды из листа «Загрузка»', 'importCodes')
    .addItem('Сгенерировать коды…', 'generateCodesPrompt')
    .addSeparator()
    .addItem('Отправить всем, кому ещё не отправлено', 'sendPending')
    .addItem('Включить автоотправку (каждые 10 мин)', 'installAutoSend')
    .addItem('Выключить автоотправку', 'removeAutoSend')
    .addSeparator()
    .addItem('Статистика', 'showStats')
    .addToUi();
}

function setup() {
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', SpreadsheetApp.getActiveSpreadsheet().getId());
  sheet_(SH_CODES, CODES_HEAD);
  sheet_(SH_LEADS, LEADS_HEAD);
  var imp = sheet_(SH_IMPORT, ['Вставьте промокоды в колонку A (по одному в строке), затем меню Промо → «Загрузить коды»']);
  imp.setColumnWidth(1, 420);
  SpreadsheetApp.getUi().alert('Готово. Листы «Промокоды», «Заявки» и «Загрузка» созданы.');
}

// ============================================================
//  Загрузка промокодов
// ============================================================
/** Переносит коды из листа «Загрузка» (колонка A) в пул. Дубликаты пропускаются. */
function importCodes() {
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var imp = sheet_(SH_IMPORT);
    var codesSh = sheet_(SH_CODES, CODES_HEAD);
    var last = imp.getLastRow();
    if (last < 2) { SpreadsheetApp.getUi().alert('Лист «Загрузка» пуст: вставьте коды в колонку A начиная со 2-й строки.'); return; }
    var raw = imp.getRange(2, 1, last - 1, 1).getValues();
    var existing = {};
    if (codesSh.getLastRow() > 1) {
      codesSh.getRange(2, 1, codesSh.getLastRow() - 1, 1).getValues().forEach(function (r) { existing[String(r[0]).trim().toUpperCase()] = 1; });
    }
    var rows = [], dup = 0, empty = 0;
    raw.forEach(function (r) {
      var c = String(r[0]).trim().toUpperCase();
      if (!c) { empty++; return; }
      if (existing[c]) { dup++; return; }
      existing[c] = 1;
      rows.push([c, ST_FREE, '', '', '', '']);
    });
    if (rows.length) codesSh.getRange(codesSh.getLastRow() + 1, 1, rows.length, CODES_HEAD.length).setValues(rows);
    imp.getRange(2, 1, last - 1, 1).clearContent();
    SpreadsheetApp.getUi().alert('Загружено: ' + rows.length + '\nДубликаты пропущены: ' + dup);
  } finally { lock.releaseLock(); }
}

function generateCodesPrompt() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt('Сколько кодов сгенерировать?', 'Например 1000. Формат: SOFT20-XXXXXX', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var n = parseInt(r.getResponseText(), 10);
  if (!(n > 0 && n <= 50000)) { ui.alert('Введите число от 1 до 50000'); return; }
  var added = generateCodes(n, 'SOFT20');
  ui.alert('Сгенерировано ' + added + ' кодов. Не забудьте завести их в 1С / на сайте, чтобы касса их принимала.');
}

/** Генерирует n уникальных кодов вида PREFIX-XXXXXX и добавляет в пул. */
function generateCodes(n, prefix) {
  var ALPH = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789'; // без 0/O, 1/I
  var codesSh = sheet_(SH_CODES, CODES_HEAD);
  var existing = {};
  if (codesSh.getLastRow() > 1) {
    codesSh.getRange(2, 1, codesSh.getLastRow() - 1, 1).getValues().forEach(function (r) { existing[r[0]] = 1; });
  }
  var rows = [];
  while (rows.length < n) {
    var s = '';
    for (var i = 0; i < 6; i++) s += ALPH.charAt(Math.floor(Math.random() * ALPH.length));
    var code = prefix + '-' + s;
    if (existing[code]) continue;
    existing[code] = 1;
    rows.push([code, ST_FREE, '', '', '', '']);
  }
  codesSh.getRange(codesSh.getLastRow() + 1, 1, rows.length, CODES_HEAD.length).setValues(rows);
  return rows.length;
}

// ============================================================
//  API для лендинга и салонов
// ============================================================
function doPost(e) {
  var body = {};
  try { body = JSON.parse(e.postData.contents || '{}'); } catch (err) { return json_({ ok: false, error: 'bad_request' }); }
  try {
    if (body.action === 'claim') return json_(claim_(body));
    if (body.action === 'check') return json_(staff_(body, false));
    if (body.action === 'redeem') return json_(staff_(body, true));
    return json_({ ok: false, error: 'unknown_action' });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: 'server' });
  }
}

function doGet() {
  return json_({ ok: true, service: 'euromebel-promo' });
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

var EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/;

function claim_(b) {
  var email = String(b.email || '').trim().toLowerCase();
  if (b.company) return { ok: true };                       // бот заполнил ловушку
  if (!EMAIL_RE.test(email) || email.length > 254) return { ok: false, error: 'invalid_email' };

  var cache = CacheService.getScriptCache();
  var rk = 'rl:' + email;
  var hits = Number(cache.get(rk) || 0);
  if (hits >= 5) return { ok: false, error: 'rate_limited' };
  cache.put(rk, String(hits + 1), 60);

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  var leadRow;
  try {
    var leads = sheet_(SH_LEADS, LEADS_HEAD);
    // один email — один код
    var found = leads.getRange('B:B').createTextFinder(email).matchEntireCell(true).findNext();
    if (found) return { ok: true, repeat: true };

    var codesSh = sheet_(SH_CODES, CODES_HEAD);
    var free = codesSh.getRange('B:B').createTextFinder(ST_FREE).matchEntireCell(true).findNext();
    if (!free) return { ok: false, error: 'no_codes' };

    var r = free.getRow();
    var code = codesSh.getRange(r, 1).getValue();
    var now = new Date();
    codesSh.getRange(r, 2, 1, 3).setValues([[ST_ISSUED, email, now]]);

    leads.appendRow([now, email, code, String(b.campaign || ''), SEND_PENDING, '', '',
      String(b.page || '').slice(0, 300), JSON.stringify(b.utm || {}).slice(0, 500)]);
    leadRow = leads.getLastRow();
  } finally {
    lock.releaseLock();
  }

  // Отправка сразу (если настроена). Ошибка отправки не мешает посетителю —
  // заявка останется в очереди со статусом «ошибка»/«ожидает».
  if (cfg_().sendOnClaim && deliveryConfigured_()) {
    try { sendRow_(leadRow); } catch (err) { console.error(err); }
  }
  return { ok: true };
}

// ---------- салоны: проверка и погашение ----------
function staff_(b, redeem) {
  var c = cfg_();
  if (!c.staffPin) return { ok: false, error: 'staff_disabled' };
  var cache = CacheService.getScriptCache();
  var fails = Number(cache.get('pinfail') || 0);
  if (fails >= 20) return { ok: false, error: 'rate_limited' };
  if (String(b.pin || '') !== c.staffPin) { cache.put('pinfail', String(fails + 1), 600); return { ok: false, error: 'bad_pin' }; }

  var code = String(b.code || '').trim().toUpperCase();
  if (!code) return { ok: false, error: 'not_found' };
  var sh = sheet_(SH_CODES, CODES_HEAD);
  var cell = sh.getRange('A:A').createTextFinder(code).matchEntireCell(true).findNext();
  if (!cell) return { ok: false, error: 'not_found' };
  var r = cell.getRow();
  var row = sh.getRange(r, 1, 1, CODES_HEAD.length).getValues()[0];
  var status = row[1];

  if (redeem) {
    if (status !== ST_ISSUED) return { ok: false, error: status === ST_REDEEMED ? 'already_redeemed' : 'not_issued', status: status };
    var lock = LockService.getScriptLock(); lock.waitLock(20000);
    try {
      sh.getRange(r, 2).setValue(ST_REDEEMED);
      sh.getRange(r, 5, 1, 2).setValues([[new Date(), String(b.salon || '').slice(0, 100)]]);
    } finally { lock.releaseLock(); }
    return { ok: true, status: ST_REDEEMED, discount: c.discount, scope: c.scope };
  }
  return {
    ok: true, status: status, discount: c.discount, scope: c.scope,
    redeemedAt: row[4] ? Utilities.formatDate(new Date(row[4]), 'Asia/Almaty', 'dd.MM.yyyy HH:mm') : '',
    salon: row[5] || ''
  };
}

// ============================================================
//  Отправка писем
// ============================================================
function deliveryConfigured_() {
  var c = cfg_();
  return !!(c.webhookUrl || c.sendViaGmail);
}

/** Отправить всем, у кого «ожидает» или «ошибка». Можно запускать из меню или по таймеру. */
function sendPending() {
  if (!deliveryConfigured_()) {
    var msg = 'Отправка не настроена: задайте WEBHOOK_URL (ваш сервис) или SEND_VIA_GMAIL=true в свойствах скрипта.';
    try { SpreadsheetApp.getUi().alert(msg); } catch (e) { console.warn(msg); }
    return;
  }
  var sh = sheet_(SH_LEADS, LEADS_HEAD);
  var last = sh.getLastRow();
  if (last < 2) return;
  var statuses = sh.getRange(2, 5, last - 1, 1).getValues();
  var started = Date.now(), sent = 0, failed = 0;
  for (var i = 0; i < statuses.length; i++) {
    if (Date.now() - started > 5 * 60 * 1000) break;        // лимит Apps Script — 6 мин на запуск
    var s = statuses[i][0];
    if (s !== SEND_PENDING && s !== SEND_ERR) continue;
    if (sendRow_(i + 2)) sent++; else failed++;
  }
  var res = 'Отправлено: ' + sent + ', ошибок: ' + failed;
  try { SpreadsheetApp.getUi().alert(res); } catch (e) { console.log(res); }
}

function sendRow_(row) {
  var c = cfg_();
  var sh = sheet_(SH_LEADS, LEADS_HEAD);
  var v = sh.getRange(row, 1, 1, LEADS_HEAD.length).getValues()[0];
  var email = v[1], code = v[2], campaign = v[3];
  var payload = {
    email: email, code: code, campaign: campaign,
    discount: c.discount, scope: c.scope, validUntil: c.validUntil, catalogUrl: c.catalogUrl,
    subject: 'Ваш промокод ' + c.discount + ' ' + c.scope,
    html: emailHtml_(code, c)
  };
  try {
    if (c.webhookUrl) {
      var resp = UrlFetchApp.fetch(c.webhookUrl, {
        method: 'post',
        contentType: 'application/json',
        headers: { 'X-Promo-Secret': c.webhookSecret },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });
      var status = resp.getResponseCode();
      if (status < 200 || status >= 300) throw new Error('HTTP ' + status + ': ' + resp.getContentText().slice(0, 200));
    } else if (c.sendViaGmail) {
      var opts = { htmlBody: payload.html, name: c.fromName };
      if (c.replyTo) opts.replyTo = c.replyTo;
      MailApp.sendEmail(email, payload.subject, 'Ваш промокод: ' + code, opts);
    } else {
      return false;
    }
    sh.getRange(row, 5, 1, 3).setValues([[SEND_OK, new Date(), '']]);
    return true;
  } catch (err) {
    sh.getRange(row, 5, 1, 3).setValues([[SEND_ERR, '', String(err).slice(0, 300)]]);
    return false;
  }
}

function installAutoSend() {
  removeAutoSend();
  ScriptApp.newTrigger('sendPending').timeBased().everyMinutes(10).create();
  SpreadsheetApp.getUi().alert('Автоотправка включена: каждые 10 минут.');
}

function removeAutoSend() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sendPending') ScriptApp.deleteTrigger(t);
  });
}

// ---------- шаблон письма (можно не использовать, если сервис шлёт свой) ----------
function emailHtml_(code, c) {
  var esc = function (s) { return String(s).replace(/[&<>"]/g, function (ch) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[ch]; }); };
  var valid = c.validUntil ? '<p style="margin:0 0 16px;color:#3D3D3D">Промокод действует до ' + esc(c.validUntil) + ' на euromebel.kz и в салонах EuroMebel.</p>' : '';
  return '' +
    '<div style="background:#F4F7FB;padding:32px 16px;font-family:Arial,sans-serif;color:#111111">' +
    '<div style="max-width:520px;margin:0 auto;background:#FFFFFF;border-radius:20px;padding:32px">' +
    '<div style="font-weight:800;font-size:26px;margin-bottom:24px"><span style="color:#E30016">Euro</span><span style="color:#00509E">Mebel</span></div>' +
    '<h1 style="margin:0 0 12px;font-size:28px">Скидка ' + esc(c.discount) + ' ваша</h1>' +
    '<p style="margin:0 0 20px;color:#3D3D3D">Ваш промокод ' + esc(c.scope) + ':</p>' +
    '<div style="font-size:30px;font-weight:800;letter-spacing:3px;color:#E30016;border:2px dashed #E30016;border-radius:14px;padding:18px;text-align:center;margin-bottom:20px">' + esc(code) + '</div>' +
    valid +
    '<a href="' + esc(c.catalogUrl) + '" style="display:inline-block;background:#E30016;color:#FFFFFF;text-decoration:none;font-weight:700;padding:16px 28px;border-radius:999px">Перейти в каталог</a>' +
    '</div></div>';
}

// ---------- статистика ----------
function showStats() {
  var codes = sheet_(SH_CODES, CODES_HEAD), leads = sheet_(SH_LEADS, LEADS_HEAD);
  var cnt = function (sh, col, val) {
    return sh.getLastRow() < 2 ? 0 : sh.getRange(2, col, sh.getLastRow() - 1, 1).getValues().filter(function (r) { return r[0] === val; }).length;
  };
  SpreadsheetApp.getUi().alert(
    'Коды: свободно ' + cnt(codes, 2, ST_FREE) + ', выдано ' + cnt(codes, 2, ST_ISSUED) + ', погашено ' + cnt(codes, 2, ST_REDEEMED) +
    '\nЗаявки: всего ' + Math.max(leads.getLastRow() - 1, 0) + ', ждут отправки ' + cnt(leads, 5, SEND_PENDING) +
    ', отправлено ' + cnt(leads, 5, SEND_OK) + ', ошибок ' + cnt(leads, 5, SEND_ERR));
}
