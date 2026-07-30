// ============================================================
//  ОПВ / Единый сервис — ДОБАВЛЕНИЯ К Code.gs
//  Подписка со сроком + отдача встроенных инструментов (iframe)
//
//  КАК ПРИМЕНИТЬ:
//  1) Скопируй ВЕСЬ этот файл в конец своего Code.gs.
//  2) Три функции — _getClientsSheet_, apiLogin, apiMe —
//     ЗАМЕНЯЮТ одноимённые в старом Code.gs. Удали старые версии
//     (иначе Apps Script выдаст ошибку "повторное объявление").
//  3) Создай в проекте Apps Script два HTML-файла:
//     DepositCalc  и  DepositBlanks (вставь одноимённые файлы).
//  4) Впиши свои данные в блок НАСТРОЙКИ ниже.
// ============================================================

// ─────────────── НАСТРОЙКИ (впиши свои значения) ───────────────
var SUB_PRICE           = 25000;                 // цена подписки, ₸/мес (для экрана оплаты)
var KASPI_PHONE         = '+7 771 128 51 35';    // номер Kaspi для перевода
var KASPI_NAME          = 'Салихалы';            // получатель Kaspi
var PARSER_DOWNLOAD_URL = '';                    // ссылка на архив парсера ПКБ/ГКБ ('' = не задана)
var COL_UNTIL           = 5;                      // колонка F "Действует до" в листе Клиенты

// Секрет для подписи лицензии парсера (HMAC). ДОЛЖЕН совпадать с LICENSE_SECRET
// в license_gate.py. Придумай длинную случайную строку (60+ символов) и НИКОМУ
// не показывай — на нём держится защита от подделки срока подписки.
var LICENSE_SECRET      = 'ЗАМЕНИ_НА_ДЛИННЫЙ_СЛУЧАЙНЫЙ_СЕКРЕТ_например_kZ9x7Qm2...';
// ───────────────────────────────────────────────────────────────


// ЗАМЕНЯЕТ старую _getClientsSheet_ — добавляет колонку F "Действует до"
function _getClientsSheet_() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName('Клиенты');
  if (!sh) {
    sh = ss.insertSheet('Клиенты');
    sh.appendRow(['Ключ (пароль)', 'Имя клиента', 'Активен (да/нет)', 'DeviceId', 'Дата привязки', 'Действует до']);
    sh.setFrozenRows(1);
    sh.getRange('A1:F1').setFontWeight('bold').setBackground('#3B0764').setFontColor('#fff');
    sh.setColumnWidths(1, 6, [140, 180, 150, 300, 140, 130]);
    sh.appendRow(['EXAMPLE', 'Пример клиента', 'нет', '', '', '']);
  } else if (sh.getLastColumn() < 6) {
    // миграция старого листа: дорисовать колонку "Действует до"
    sh.getRange(1, 6).setValue('Действует до')
      .setFontWeight('bold').setBackground('#3B0764').setFontColor('#fff');
    sh.setColumnWidth(6, 130);
  }
  return sh;
}

// Возвращает полную строку клиента по ключу (имя, активность, устройство, срок)
function _clientRow_(key) {
  var data = _getClientsSheet_().getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (!data[i]) continue;
    if (String(data[i][COL.KEY]).trim() === key) {
      return {
        row:    i + 1,
        name:   String(data[i][COL.NAME] || '').trim(),
        active: String(data[i][COL.ACTIVE] || '').trim().toLowerCase() === 'да',
        device: String(data[i][COL.DEVICE] || '').trim() || null,
        until:  data[i][COL_UNTIL] || null
      };
    }
  }
  return null;
}

// Разбирает «Действует до» в Date. Понимает ячейку-дату Sheets и текст
// «дд.мм.гггг» / «гггг-мм-дд». Возвращает Date или null (дата не задана/битая).
function _untilDate_(until) {
  if (until === '' || until === null || until === undefined) return null;
  if (until instanceof Date) return new Date(until.getTime());
  var s = String(until).trim();
  var m = s.match(/^(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{4})$/);      // дд.мм.гггг
  if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
  var m2 = s.match(/^(\d{4})[.\/-](\d{1,2})[.\/-](\d{1,2})$/);     // гггг-мм-дд
  if (m2) return new Date(Number(m2[1]), Number(m2[2]) - 1, Number(m2[3]));
  var d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

// Сколько дней осталось до конца подписки. null = дата не задана (бессрочно).
function _subDaysLeft_(until) {
  var end = _untilDate_(until);
  if (!end) return null;
  var today = new Date(); today.setHours(0, 0, 0, 0);
  end.setHours(0, 0, 0, 0);
  return Math.round((end - today) / 86400000);
}

// Подписанный токен лицензии: клиент не сможет подделать срок без LICENSE_SECRET.
function _signLicense_(key, deviceId, expISO, name) {
  var payloadObj = { k: key, d: deviceId, exp: expISO, n: name };
  var payload = Utilities.base64Encode(JSON.stringify(payloadObj));
  var sig = Utilities.base64Encode(
    Utilities.computeHmacSha256Signature(payload, LICENSE_SECRET));
  return { token: payload, sig: sig };
}

function _fmtDate_(d) {
  try { return Utilities.formatDate(new Date(d), 'Asia/Almaty', 'dd.MM.yyyy'); }
  catch (e) { return String(d); }
}

// Ответ для экрана оплаты (подписка истекла)
function _payInfo_(key, name, msg) {
  return {
    ok: false, expired: true, message: msg,
    key: key, clientName: name,
    kaspiPhone: KASPI_PHONE, kaspiName: KASPI_NAME, price: SUB_PRICE
  };
}

// ЗАМЕНЯЕТ старую apiLogin — вход с проверкой срока подписки
function apiLogin(key, deviceId) {
  key = String(key || '').trim();
  deviceId = String(deviceId || '').trim();
  var c = _clientRow_(key);
  if (!c || !c.name)       return { ok: false, message: 'Неверный ключ доступа' };
  if (!c.active)           return { ok: false, message: 'Ключ отключён. Обратитесь к администратору' };
  if (!deviceId)           return { ok: false, message: 'Ошибка идентификации устройства' };
  if (c.device && c.device !== deviceId)
                           return { ok: false, message: '⛔ Ключ привязан к другому устройству' };

  var days = _subDaysLeft_(c.until);
  if (days !== null && days < 0)
    return _payInfo_(key, c.name, 'Срок подписки истёк ' + _fmtDate_(c.until) + '.');

  if (!c.device) _setDevice_(key, deviceId);
  var token = _newToken_();
  _setSession_(token, { client: c.name, key: key, device: deviceId });
  _log_(c.name, key, 'LOGIN', '', '', 0, 0);
  return {
    ok: true, token: token, clientName: c.name,
    daysLeft: days, parserUrl: PARSER_DOWNLOAD_URL
  };
}

// ЗАМЕНЯЕТ старую apiMe — восстановление сессии + проверка подписки
function apiMe(t) {
  var s = _refreshSession_(t);
  if (!s) return { ok: false };
  var c = _clientRow_(s.key);
  if (c) {
    if (!c.active) return { ok: false, message: 'Ключ отключён' };
    var days = _subDaysLeft_(c.until);
    if (days !== null && days < 0)
      return _payInfo_(s.key, c.name, 'Срок подписки истёк ' + _fmtDate_(c.until) + '.');
    return { ok: true, clientName: s.client, daysLeft: days, parserUrl: PARSER_DOWNLOAD_URL };
  }
  return { ok: true, clientName: s.client };
}

// Отдаёт HTML встроенного инструмента (грузится в изолированный iframe)
function getToolHtml(fileName, token) {
  _requireSession_(token);
  var allowed = { DepositCalc: 1, DepositBlanks: 1 };
  if (!allowed[fileName]) throw new Error('Unknown tool: ' + fileName);
  return HtmlService.createHtmlOutputFromFile(fileName).getContent();
}

// ============================================================
//  JSON-API для десктоп-парсера ПКБ/ГКБ (проверка подписки «по времени»)
//  Десктоп-программа шлёт POST {action:'login', key, deviceId}
//  и получает статус подписки. Требует деплой «Все, у кого есть ссылка».
// ============================================================
function doPost(e) {
  var out;
  try {
    var body = (e && e.postData && e.postData.contents)
      ? JSON.parse(e.postData.contents) : {};
    var action = body.action || 'login';
    if (action === 'checkToken') {
      // серверный парсер проверяет сессию веб-хаба
      out = apiCheckToken_(String(body.token || ''));
    } else {
      // десктоп-парсер: вход по ключу + привязка к устройству
      out = apiLicense_(String(body.key || ''), String(body.deviceId || ''));
    }
  } catch (err) {
    out = { ok: false, reason: 'error', message: String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

// Проверка активной сессии веб-хаба (для серверного парсера на Cloud Run)
function apiCheckToken_(token) {
  var s = _refreshSession_(token);
  if (!s) return { ok: false, message: 'Сессия истекла — войдите в сервис заново' };
  return { ok: true, clientName: s.client };
}

// Проверка ключа + срока подписки + привязки к устройству (для десктопа)
function apiLicense_(key, deviceId) {
  key = String(key || '').trim();
  deviceId = String(deviceId || '').trim();
  var c = _clientRow_(key);
  if (!c || !c.name) return { ok: false, reason: 'badkey',   message: 'Неверный ключ доступа' };
  if (!c.active)     return { ok: false, reason: 'disabled', message: 'Ключ отключён администратором' };
  if (!deviceId)     return { ok: false, reason: 'nodevice', message: 'Нет идентификатора устройства' };
  if (c.device && c.device !== deviceId)
                     return { ok: false, reason: 'device',   message: 'Ключ привязан к другому устройству' };

  var days = _subDaysLeft_(c.until);
  if (days !== null && days < 0)
    return {
      ok: false, reason: 'expired',
      message: 'Подписка истекла ' + _fmtDate_(c.until),
      until: _fmtDate_(c.until),
      kaspiPhone: KASPI_PHONE, kaspiName: KASPI_NAME, price: SUB_PRICE, key: key
    };

  if (!c.device) _setDevice_(key, deviceId);

  // Подписанный токен для офлайн-проверки на десктопе (защита от подделки срока).
  var end = _untilDate_(c.until);
  var expISO = end ? Utilities.formatDate(end, 'Asia/Almaty', 'yyyy-MM-dd') : '2999-12-31';
  var lic = _signLicense_(key, deviceId, expISO, c.name);

  return {
    ok: true, clientName: c.name, daysLeft: days,
    until: c.until ? _fmtDate_(c.until) : '',   // '' = бессрочно
    token: lic.token, sig: lic.sig,
    kaspiPhone: KASPI_PHONE, kaspiName: KASPI_NAME, price: SUB_PRICE
  };
}
