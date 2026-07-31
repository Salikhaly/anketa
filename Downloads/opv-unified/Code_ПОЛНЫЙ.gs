// ============================================================
//  ОПВ / Единый сервис — ПОЛНЫЙ Code.gs (одним файлом)
//
//  Этот файл ЗАМЕНЯЕТ собой и старый Код.gs, и Code_additions.gs.
//  Просто вставь его целиком в Код.gs, а файл Code_additions.gs
//  удали из проекта (или очисти) — иначе будут дубли функций.
// ============================================================

var SPREADSHEET_ID = '1yDAFMqOldkNRvZnh0A6FBJAnr8tYo-129wlA08nPmMM';
var SESSION_SEC    = 1800;
var LOG_LOCK_MS    = 3000;
var COEFF          = { 3:0.1688, 4:0.1519, 5:0.1381, 6:0.1266 };
var COL            = { KEY:0, NAME:1, ACTIVE:2, DEVICE:3, DATE:4 };

// ─────────────── НАСТРОЙКИ ПОДПИСКИ ───────────────
var SUB_PRICE           = 25000;                 // цена подписки, ₸/мес
var KASPI_PHONE         = '+7 771 128 51 35';    // номер Kaspi для перевода
var KASPI_NAME          = 'Салихалы';            // получатель Kaspi
var PARSER_DOWNLOAD_URL = '';                    // ссылка на архив десктоп-парсера
var COL_UNTIL           = 5;                     // колонка F «Действует до»

// Секрет подписи лицензии десктоп-парсера (HMAC).
// ДОЛЖЕН совпадать с LICENSE_SECRET в license_gate.py.
var LICENSE_SECRET      = 'ЗАМЕНИ_НА_ДЛИННЫЙ_СЛУЧАЙНЫЙ_СЕКРЕТ_например_kZ9x7Qm2...';
// ───────────────────────────────────────────────────

var MORTGAGE_PROGRAMS = {
  '5050':     { key:'5050',     name:'Ипотека 50/50',       icon:'🏛️', downRatio:0.50, desc:'Взнос 50%', variants:[{label:'8.5% — 8 лет',coeff:0.0080522},{label:'5% — 6 лет',coeff:0.0080525}] },
  '3070':     { key:'3070',     name:'Программа 30/70',     icon:'🏠', downRatio:0.30, desc:'Взнос 30%', variants:[{label:'~10-12 лет',  coeff:0.00886788}] },
  'nauryz20': { key:'nauryz20', name:'Наурыз (взнос 20%)',  icon:'🌸', downRatio:0.20, desc:'Взнос 20%', variants:[{label:'7% — 19 лет', coeff:0.00843335},{label:'9% — 19 лет',coeff:0.0101}] },
  'nauryz10': { key:'nauryz10', name:'Наурыз (взнос 10%)',  icon:'🌷', downRatio:0.10, desc:'Взнос 10%', variants:[{label:'7% — 19 лет', coeff:0.00943335},{label:'9% — 19 лет',coeff:0.0111}] },
  'jasyl':    { key:'jasyl',    name:'Жасыл Ипотека',       icon:'🌿', downRatio:0.20, desc:'Взнос 20%', variants:[{label:'7% очередники',coeff:0.00783333},{label:'11% военные',coeff:0.01116667},{label:'15% все',coeff:0.01241667}] },
  'askeri':   { key:'askeri',   name:'Наурыз Аскери',       icon:'🎖️', downRatio:0.00, desc:'Взнос 0%',  variants:[{label:'1-8 лет',     coeff:0.0127},     {label:'9-19 лет',  coeff:0.00376}] }
};

// Лист «Клиенты» с колонкой F «Действует до»
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
    sh.getRange(1, 6).setValue('Действует до')
      .setFontWeight('bold').setBackground('#3B0764').setFontColor('#fff');
    sh.setColumnWidth(6, 130);
  }
  return sh;
}

function _initSettingsSheet_(ss) {
  var sh = ss.getSheetByName('Настройки');
  if (!sh) {
    sh = ss.insertSheet('Настройки');
    sh.appendRow(['Параметр', 'Значение', 'Описание']);
    sh.getRange('A1:C1').setFontWeight('bold').setBackground('#1e3a5f').setFontColor('#fff');
    sh.setFrozenRows(1);
    sh.setColumnWidths(1, 3, [160, 120, 300]);
    var rows = [
      ['МРП',        4325,  'МРП текущего года (₸)'],
      ['ПМ_НАУРЫЗ',  10,    'Прожит.мин для Наурыз: кол-во МРП на 1 чел.'],
      ['ПМ_ДРУГИЕ',  13,    'Прожит.мин для остальных: кол-во МРП на 1 чел.'],
      ['КД_ПОРОГ_1', 40,    'Если доход/МРП ≤ этого → КД = значение 1'],
      ['КД_ПОРОГ_2', 65,    'Если доход/МРП ≤ этого → КД = значение 2'],
      ['КД_ПОРОГ_3', 90,    'Если доход/МРП ≤ этого → КД = значение 3'],
      ['КД_1',       0.40,  'КД при доходе ≤ порог 1'],
      ['КД_2',       0.50,  'КД при доходе ≤ порог 2'],
      ['КД_3',       0.60,  'КД при доходе ≤ порог 3'],
      ['КД_4',       0.70,  'КД при доходе > порог 3'],
    ];
    rows.forEach(function(r) { sh.appendRow(r); });
    sh.getRange('A2:A11').setBackground('#EDE9FE').setFontWeight('bold');
    sh.getRange('B2:B11').setBackground('#F5F3FF');
  }
}

function _initProgramsSheet_(ss) {
  var sh = ss.getSheetByName('Программы');
  if (!sh) {
    sh = ss.insertSheet('Программы');
    sh.appendRow(['Ключ','Название','Иконка','Взнос %','Вар1 Название','Вар1 Коэфф','Вар2 Название','Вар2 Коэфф','Вар3 Название','Вар3 Коэфф','Описание']);
    sh.getRange('A1:K1').setFontWeight('bold').setBackground('#065F46').setFontColor('#fff');
    sh.setFrozenRows(1);
    sh.setColumnWidths(1, 11, [90,170,50,65,150,100,150,100,150,100,200]);
    var progs = [
      ['5050',     'Ипотека 50/50',      '🏛️', 50, '8.5% — 8 лет',        0.0080522,  '5% — 6 лет',          0.0080525,  '',                  '',          'Первоначальный взнос 50%'],
      ['3070',     'Программа 30/70',    '🏠', 30, '~10-12 лет',          0.00886788, '',                    '',         '',                  '',          'Первоначальный взнос 30%'],
      ['nauryz20', 'Наурыз (взнос 20%)', '🌸', 20, '7% годовых — 19 лет', 0.00843335, '9% годовых — 19 лет', 0.0101,     '',                  '',          'Государственная программа, взнос 20%, срок 19 лет'],
      ['nauryz10', 'Наурыз (взнос 10%)', '🌷', 10, '7% годовых — 19 лет', 0.00943335, '9% годовых — 19 лет', 0.0111,     '',                  '',          'Государственная программа, взнос 10%, срок 19 лет'],
      ['jasyl',    'Жасыл Ипотека',      '🌿', 20, '7% — очередники',     0.00783333, '11% — военные',       0.01116667, '15% — все граждане',0.01241667,  'Экопрограмма, взнос 20%'],
      ['askeri',   'Наурыз Аскери',      '🎖️',  0, 'Первые 8 лет (фикс)', 0.0127,     'После 8 лет (ост)',   0.00376,    '',                  '',          'Военная ипотека, взнос 0%'],
    ];
    progs.forEach(function(r, i) {
      sh.appendRow(r);
      sh.getRange(i+2, 1, 1, 11).setBackground(i%2===0?'#F0FDF4':'#FFFFFF');
    });
    sh.getRange('A2:A7').setFontWeight('bold').setBackground('#D1FAE5');
    sh.getRange('D2:D7').setHorizontalAlignment('center').setFontWeight('bold');
    sh.getRange('F2:J7').setNumberFormat('0.0000000');
    sh.getRange('A9').setValue('⚠️ Коэффициент = ежемесячный платёж ÷ сумма займа');
    sh.getRange('A9:K9').merge().setFontStyle('italic').setFontColor('#666').setFontSize(9);
  }
}

var _memCache_={};
function _readSettings_() {
  if(_memCache_.cfg && (Date.now()-_memCache_.ts)<120000) return _memCache_.cfg;
  var cache = CacheService.getScriptCache();
  var cached = cache.get('CFG_V3');
  if (cached) { try { var p=JSON.parse(cached); _memCache_={cfg:p,ts:Date.now()}; return p; } catch(e) {} }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  _initSettingsSheet_(ss);
  _initProgramsSheet_(ss);

  var shS = ss.getSheetByName('Настройки');
  var sData = shS.getDataRange().getValues();
  var cfg = {};
  for (var i = 1; i < sData.length; i++) {
    var k = String(sData[i][0]||'').trim();
    if (k) cfg[k] = Number(sData[i][1])||0;
  }
  var mrp=cfg['МРП']||4325, pmNauryz=cfg['ПМ_НАУРЫЗ']||10, pmOther=cfg['ПМ_ДРУГИЕ']||13;
  var kd={p1:cfg['КД_ПОРОГ_1']||40,p2:cfg['КД_ПОРОГ_2']||65,p3:cfg['КД_ПОРОГ_3']||90,
          v1:cfg['КД_1']||0.40,v2:cfg['КД_2']||0.50,v3:cfg['КД_3']||0.60,v4:cfg['КД_4']||0.70};

  var shP = ss.getSheetByName('Программы');
  var pData = shP.getDataRange().getValues();
  var programs = {};
  for (var j = 1; j < pData.length; j++) {
    var row = pData[j];
    var pKey = String(row[0]||'').trim();
    if (!pKey || pKey.charAt(0)==='⚠') continue;
    var variants = [];
    for (var v = 0; v < 3; v++) {
      var lbl=String(row[4+v*2]||'').trim(), coeff=Number(row[5+v*2])||0;
      if (lbl&&coeff>0) variants.push({label:lbl,coeff:coeff});
    }
    if (!variants.length) continue;
    programs[pKey]={key:pKey,name:String(row[1]||''),icon:String(row[2]||'🏠'),
                    downRatio:(Number(row[3])||0)/100,desc:String(row[10]||''),variants:variants};
  }
  if (!Object.keys(programs).length) programs = MORTGAGE_PROGRAMS;

  var result={mrp:mrp,pmNauryz:pmNauryz,pmOther:pmOther,kd:kd,programs:programs,nauryzKeys:['nauryz10','nauryz20']};
  _memCache_={cfg:result,ts:Date.now()}; cache.put('CFG_V3', JSON.stringify(result), 600);
  return result;
}

function _getKdFromSettings_(income,cfg){
  var r=income/cfg.mrp;
  if(r<=cfg.kd.p1)return cfg.kd.v1;
  if(r<=cfg.kd.p2)return cfg.kd.v2;
  if(r<=cfg.kd.p3)return cfg.kd.v3;
  return cfg.kd.v4;
}

function _getPm_(programKey,cfg,members){
  var isN=cfg.nauryzKeys.indexOf(programKey)>=0;
  return cfg.mrp*(isN?cfg.pmNauryz:cfg.pmOther)*(members||1);
}

function _calcMaxPayment_(orgs,members,programKey,cfg){
  var totalIncome=0,totalOldCredit=0;
  orgs.forEach(function(o){
    totalIncome+=Number(o.income)||0;
    totalOldCredit+=Number(o.oldCredit)||0;
  });
  var pm=_getPm_(programKey,cfg,members);

  var kd=0.70;
  for(var iter=0; iter<10; iter++){
    var m1=Math.round(totalIncome*kd-totalOldCredit);
    var m2=Math.round(totalIncome-totalIncome*0.10-pm-totalOldCredit);
    var max=Math.min(m1,m2);

    var testIncome=totalIncome>0?totalIncome:(max+totalOldCredit)/kd;
    if(testIncome<=0) testIncome=1;
    var newKd=_getKdFromSettings_(testIncome,cfg);
    if(Math.abs(newKd-kd)<0.001) break;
    kd=newKd;
  }

  var method1=Math.round(totalIncome*kd-totalOldCredit);
  var method2=Math.round(totalIncome-totalIncome*0.10-pm-totalOldCredit);
  var maxPayment=Math.min(method1,method2);

  return {totalIncome:totalIncome,totalOldCredit:totalOldCredit,kd:kd,pm:pm,
          method1:method1,method2:method2,maxPayment:maxPayment,approved:maxPayment>0};
}

var _clientsCache_=null;
function getClients(){
  if(_clientsCache_) return _clientsCache_;
  var data=_getClientsSheet_().getDataRange().getValues(),out={};
  for(var i=1; i < data.length; i++){
    if (!data[i]) continue;
    var k=String(data[i][COL.KEY]||'').trim(),
        n=String(data[i][COL.NAME]||'').trim(),
        a=String(data[i][COL.ACTIVE]||'').trim().toLowerCase();
    if(k&&n&&a==='да') out[k]=n;
  }
  _clientsCache_=out;
  return out;
}

// Полная строка клиента: имя, активность, устройство, срок подписки
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

function _getDevice_(key){
  var data=_getClientsSheet_().getDataRange().getValues();
  for(var i=1; i < data.length; i++){
    if (!data[i]) continue;
    if(String(data[i][COL.KEY]).trim()===key) return String(data[i][COL.DEVICE]||'').trim()||null;
  }
  return null;
}

function _setDevice_(key,deviceId){
  var sh=_getClientsSheet_(),data=sh.getDataRange().getValues();
  for(var i=1; i < data.length; i++){
    if (!data[i]) continue;
    if(String(data[i][COL.KEY]).trim()===key){
      sh.getRange(i+1,COL.DEVICE+1).setValue(deviceId);
      sh.getRange(i+1,COL.DATE+1).setValue(new Date());
      return;
    }
  }
}

// Дата «Действует до»: понимает ячейку-дату и текст «дд.мм.гггг» / «гггг-мм-дд»
function _untilDate_(until) {
  if (until === '' || until === null || until === undefined) return null;
  if (until instanceof Date) return new Date(until.getTime());
  var s = String(until).trim();
  var m = s.match(/^(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{4})$/);
  if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
  var m2 = s.match(/^(\d{4})[.\/-](\d{1,2})[.\/-](\d{1,2})$/);
  if (m2) return new Date(Number(m2[1]), Number(m2[2]) - 1, Number(m2[3]));
  var d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

// Сколько дней осталось. null = дата не задана (бессрочно)
function _subDaysLeft_(until) {
  var end = _untilDate_(until);
  if (!end) return null;
  var today = new Date(); today.setHours(0, 0, 0, 0);
  end.setHours(0, 0, 0, 0);
  return Math.round((end - today) / 86400000);
}

// Подписанный токен лицензии для десктоп-парсера
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

function doGet(){
  return HtmlService.createTemplateFromFile('Index')
    .evaluate().setTitle('Единый сервис')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function _newToken_(){ return 'T'+Utilities.getUuid().replace(/-/g,''); }
function _setSession_(t,d){ d._ts=Date.now(); CacheService.getScriptCache().put('S_'+t,JSON.stringify(d),SESSION_SEC); }
function _getSession_(t){
  if(!t) return null;
  try {
    var r = CacheService.getScriptCache().get('S_'+t);
    return r ? JSON.parse(r) : null;
  } catch(e) {
    return null;
  }
}
function _refreshSession_(t){
  var s=_getSession_(t);
  if(!s) return null;
  if((Date.now()-(s._ts||0))/1000>SESSION_SEC){
    CacheService.getScriptCache().remove('S_'+t);
    return null;
  }
  _setSession_(t,s);
  return s;
}
function _requireSession_(t){
  var s=_refreshSession_(t);
  if(!s) throw new Error('NO_SESSION');
  return s;
}

// Вход с проверкой срока подписки
function apiLogin(key, deviceId){
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

function apiLogout(t){
  if(t) CacheService.getScriptCache().remove('S_'+t);
  return {ok:true};
}

// Восстановление сессии + проверка подписки
function apiMe(t){
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

function _log_(client,key,mode,type,summary,avg,total){
  try{
    var lock=LockService.getScriptLock();
    if(!lock.tryLock(LOG_LOCK_MS)) return;
    try{
      var ss=SpreadsheetApp.openById(SPREADSHEET_ID);
      var sh=ss.getSheetByName('Логи');
      if(!sh){
        sh=ss.insertSheet('Логи');
        sh.appendRow(['Дата','Клиент','Ключ','Режим','Тип','Данные','Ср.ЗП','ИТОГО']);
        sh.setFrozenRows(1);
        sh.getRange('A1:H1').setFontWeight('bold').setBackground('#1e3a5f').setFontColor('#fff');
        sh.setColumnWidths(1,8,[140,160,80,60,80,300,120,120]);
      }
      sh.appendRow([new Date(),client,key,mode,type,summary,avg||'',total||'']);
    } finally {
      lock.releaseLock();
    }
  } catch(e) {}
}

function apiGetHistory(token){
  var s=_requireSession_(token);
  try{
    var ss=SpreadsheetApp.openById(SPREADSHEET_ID);
    var sh=ss.getSheetByName('Логи');
    if(!sh) return {ok:true,rows:[]};
    var data=sh.getDataRange().getValues(),rows=[];
    for(var i=data.length-1; i>=1 && rows.length<30; i--){
      if(String(data[i][2]).trim()===s.key && String(data[i][3]).trim()!=='LOGIN'){
        rows.push({
          date: Utilities.formatDate(new Date(data[i][0]),'Asia/Almaty','dd.MM.yyyy HH:mm'),
          mode: String(data[i][3]||''),
          type: String(data[i][4]||''),
          summary: String(data[i][5]||'').slice(0,60),
          avg: Number(data[i][6])||0,
          total: Number(data[i][7])||0
        });
      }
    }
    return {ok:true,rows:rows};
  } catch(e) {
    return {ok:false,message:String(e)};
  }
}

// Средняя зарплата из ОПВ.
// ВАЖНО: знаменатель max(6, n), а не жёсткая «6». Раньше при 12 заполненных
// месяцах результат завышался в полтора раза, потому что вторая половина
// формулы делится на n−2 и масштабируется правильно, а первая — нет.
function _avgSalary_(list){
  var n = list.length, sum = list.reduce(function(a,b){ return a+b; }, 0);
  var s1 = sum * 7.9 / Math.max(6, n);
  if (n < 3) return Math.round(s1);
  var mx = Math.max.apply(null, list), mn = Math.min.apply(null, list);
  return Math.round((s1 + (sum - mx - mn) * 7.9 / (n - 2)) / 2);
}

function _calcTax_(s,type){
  var opv=Math.round(s*0.10), vosms=Math.min(Math.round(s*0.02),34000),
      so=Math.min(Math.round(s*0.045),29750), work=15000, ipn, opvr, oosms;
  if(type==='Трудовой'){
    ipn=s<=147444?0:Math.max(0,Math.round((s*0.9-s*0.02-129750)*0.10));
    opvr=Math.round(s*0.035);
    oosms=Math.round(s*0.03);
  } else {
    ipn=Math.max(0,Math.round((s*0.9-vosms-so)*0.10));
    opvr=0;
    oosms=0;
  }
  return {opv:opv,vosms:vosms,so:so,ipn:ipn,opvr:opvr,oosms:oosms,work:work,
          total:opv+vosms+so+ipn+opvr+oosms+work};
}

function _buildBreakdown_(salary,type){
  var t=_calcTax_(salary,type), isT=type==='Трудовой';
  return [
    {label:'ОПВ', val:t.opv},
    {label:'ВОСМС', val:t.vosms},
    {label:'СО', val:t.so},
    {label:'ИПН', val:t.ipn},
    {label:'ОПВР (работодатель)', val:isT?t.opvr:null},
    {label:'ООСМС (работодатель)', val:isT?t.oosms:null},
    {label:'Работа / услуга', val:t.work},
    {label:'ИТОГО к бухгалтеру', val:t.total, isTotal:true}
  ];
}

function _findNeededOpv_(existing,target,addMonths){
  if(target<=0 || addMonths<=0) return null;
  var lo=0, hi=target*0.25;
  for(var i=0; i<80; i++){
    var mid=(lo+hi)/2, list=existing.slice();
    for(var j=0; j<addMonths; j++) list.push(mid);
    if(_avgSalary_(list)<target) lo=mid; else hi=mid;
    if(Math.abs(hi-lo)<0.5) break;
  }
  var opv=Math.round((lo+hi)/2), check=existing.slice();
  for(var k=0; k<addMonths; k++) check.push(opv);
  return {opvPerMonth:opv, achievedSalary:_avgSalary_(check)};
}

function apiGetSettingsMeta(token){
  _requireSession_(token);
  var cfg=_readSettings_();
  return {
    ok:true,
    mrp:cfg.mrp,
    programs:Object.keys(cfg.programs).map(function(k){
      var p=cfg.programs[k];
      return {
        key:p.key, name:p.name, icon:p.icon, desc:p.desc,
        downRatio:p.downRatio,
        variantCount:p.variants.length,
        variantLabels:p.variants.map(function(v){return v.label;})
      };
    })
  };
}

// Полный конфиг для ипотечного калькулятора в браузере.
// Расчёт идёт на клиенте (мгновенно), цифры — всегда из таблицы.
function apiGetCalcConfig(token){
  _requireSession_(token);
  var cfg = _readSettings_();
  var programs = Object.keys(cfg.programs).map(function(k){
    var p = cfg.programs[k];
    return {
      key: p.key, name: p.name, icon: p.icon, desc: p.desc,
      downRatio: p.downRatio,
      variants: p.variants.map(function(v){ return { label: v.label, coeff: v.coeff }; })
    };
  });
  return {
    ok: true,
    mrp: cfg.mrp,
    pmNauryz: cfg.pmNauryz,
    pmOther: cfg.pmOther,
    kd: cfg.kd,
    nauryzKeys: cfg.nauryzKeys,
    incomeReservePct: 0.10,
    programs: programs
  };
}

function apiCalcVar(token,payload){
  var s=_requireSession_(token);
  payload=payload||{};
  var type=String(payload.type||''), salary=Number(payload.salary)||0,
      target=Number(payload.target)||0, raw=payload.opv12||[];
  var opvList=raw.map(Number).filter(function(v){return isFinite(v)&&v>0;});
  var avgSalary=opvList.length?_avgSalary_(opvList):0;
  var targetPlan=null;
  if(target>0 && target>avgSalary){
    targetPlan=[];
    for(var n=1; n<=6; n++){
      var r=_findNeededOpv_(opvList,target,n);
      targetPlan.push({months:n, opvPerMonth:r?r.opvPerMonth:null, achieved:r?r.achievedSalary:null});
    }
  }
  var breakdown=(salary>0&&type)?_buildBreakdown_(salary,type):null;
  _log_(s.client,s.key,'VAR',type,'sal='+salary+' n='+opvList.length,avgSalary,
        breakdown?breakdown[breakdown.length-1].val:0);
  return {ok:true, avgSalary:avgSalary, breakdown:breakdown, targetPlan:targetPlan};
}

function apiCalcEq(token,payload){
  var s=_requireSession_(token);
  payload=payload||{};
  var type=String(payload.type||''), income=Number(payload.income)||0,
      months=Number(payload.months)||0, salary=Number(payload.salary)||0;
  var payments=[], avgSalary=0;
  if(income>0 && COEFF[months]){
    var mo=Math.round(income*COEFF[months]);
    for(var i=0; i<months; i++) payments.push(mo);
    avgSalary=_avgSalary_(payments);
  }
  var breakdown=(salary>0&&type)?_buildBreakdown_(salary,type):null;
  _log_(s.client,s.key,'EQ',type,'inc='+income+' mo='+months,avgSalary,
        breakdown?breakdown[breakdown.length-1].val:0);
  return {ok:true, payments:payments, avgSalary:avgSalary, breakdown:breakdown};
}

function apiCalcBuh(token,payload){
  var s=_requireSession_(token);
  payload=payload||{};
  var type=String(payload.type||''), salary=Number(payload.salary)||0;
  if(!type || !salary) return {ok:false, message:'Укажите тип и зарплату'};
  var bd=_buildBreakdown_(salary,type);
  _log_(s.client,s.key,'BUH',type,'sal='+salary,0,bd[bd.length-1].val);
  return {ok:true, breakdown:bd};
}

function apiScenario(token,payload){
  var s=_requireSession_(token);
  payload=payload||{};
  var scenario=String(payload.scenario||''), type=String(payload.type||'Трудовой'),
      existing=(payload.existing||[]).map(Number).filter(function(v){return isFinite(v)&&v>0;});
  var target=Number(payload.target)||300000;
  var cur=existing.length?_avgSalary_(existing):0;
  var plan=[];
  for(var n=1; n<=6; n++){
    var r=_findNeededOpv_(existing,target,n);
    plan.push({
      months:n,
      opvPerMonth:r?r.opvPerMonth:null,
      achieved:r?r.achievedSalary:null,
      totalCost:r?r.opvPerMonth*n:null
    });
  }
  var bd=_buildBreakdown_(target,type);
  var best=plan.reduce(function(b,p){return(!b||p.totalCost<b.totalCost)?p:b;},null);
  _log_(s.client,s.key,'SCENARIO',scenario,'tgt='+target,0,0);
  return {ok:true, scenario:scenario,
          result:{target:target,currentAvg:cur,gap:Math.max(0,target-cur),
                  plan:plan,breakdown:bd,best:best}};
}

// ══════════════════════════════════════════════════════════
//  ИПОТЕКА (серверные расчёты — используются старыми экранами)
// ══════════════════════════════════════════════════════════

function _calcReqSalary_(monthly, oldCredit, pm, cfg) {
  var reqCSD = Math.round((monthly + pm + oldCredit) / 0.9);
  var kd = 0.70;
  for(var iter = 0; iter < 20; iter++) {
    var reqOD0 = Math.round((monthly + oldCredit) / kd);
    var newKd = _getKdFromSettings_(reqOD0, cfg);
    if(Math.abs(newKd - kd) < 0.001) break;
    kd = newKd;
  }
  var reqOD = Math.round((monthly + oldCredit) / kd);
  return { reqCSD: reqCSD, reqOD: reqOD, reqSalary: Math.max(reqCSD, reqOD), kd: kd };
}

function apiCalcMortgage(token,payload){
  _requireSession_(token);
  payload = payload || {};

  var progKey = String(payload.program || '');
  var price = Number(payload.price) || 0;
  var orgs = payload.orgs || [{income: Number(payload.salary) || 0, oldCredit: 0}];
  var members = Number(payload.members) || 1;
  var existingOpv = (payload.existingOpv || []).map(Number).filter(function(v){ return isFinite(v) && v > 0; });

  var cfg = _readSettings_();
  var prog = cfg.programs[progKey];
  if(!prog) return {ok: false, message: 'Программа не найдена'};

  if(orgs.length > 0 && orgs[0].oldCredit !== undefined) {
    orgs[0].oldCredit = Number(orgs[0].oldCredit) || 0;
  }

  var calc = _calcMaxPayment_(orgs, members, progKey, cfg);

  var variantsByPrice = null;
  if(price > 0){
    variantsByPrice = prog.variants.map(function(v, idx){
      var monthly = Math.round(price * v.coeff);
      var down = Math.round(price * prog.downRatio);
      var pm = _getPm_(progKey, cfg, members);
      var req = _calcReqSalary_(monthly, calc.totalOldCredit, pm, cfg);

      var opvPlan = [];
      for(var n = 1; n <= 6; n++){
        var plan = _findNeededOpv_(existingOpv, req.reqSalary, n);
        opvPlan.push({
          months: n,
          opvPerMonth: plan ? plan.opvPerMonth : null,
          achieved: plan ? plan.achievedSalary : null,
          totalCost: plan ? plan.opvPerMonth * n : null
        });
      }

      return {
        label: v.label,
        monthly: monthly,
        downPayment: down,
        loanAmount: price - down,
        requiredSalary: req.reqSalary,
        reqByOD: req.reqOD,
        reqByCSD: req.reqCSD,
        kd: req.kd,
        canAfford: calc.approved && monthly <= calc.maxPayment,
        opvPlan: opvPlan
      };
    });
  }

  var variantsBySalary = null;
  if(calc.approved && calc.maxPayment > 0){
    variantsBySalary = prog.variants.map(function(v){
      var maxLoan = Math.round(calc.maxPayment / v.coeff);
      var maxPrice = Math.round(maxLoan / (1 - prog.downRatio));
      return {
        label: v.label,
        maxMonthly: calc.maxPayment,
        maxLoan: maxLoan,
        maxPrice: maxPrice,
        downPayment: Math.round(maxPrice * prog.downRatio),
        method1: calc.method1,
        method2: calc.method2,
        kd: calc.kd,
        pm: calc.pm
      };
    });
  }

  var comparison = null;
  if(price > 0){
    comparison = Object.keys(cfg.programs).map(function(key){
      var p = cfg.programs[key];
      var v0 = p.variants[0];
      var monthly = Math.round(price * v0.coeff);
      var pmCompare = _getPm_(key, cfg, members);
      var req = _calcReqSalary_(monthly, calc.totalOldCredit, pmCompare, cfg);

      return {
        key: key,
        name: p.name,
        icon: p.icon,
        downPayment: Math.round(price * p.downRatio),
        monthly: monthly,
        requiredSalary: req.reqSalary,
        reqByOD: req.reqOD,
        reqByCSD: req.reqCSD,
        canAfford: calc.approved && monthly <= calc.maxPayment
      };
    });
    comparison.sort(function(a, b){ return a.monthly - b.monthly; });
  }

  return {
    ok: true,
    calc: calc,
    variantsByPrice: variantsByPrice,
    variantsBySalary: variantsBySalary,
    comparison: comparison,
    prog: { name: prog.name, icon: prog.icon, downRatio: prog.downRatio }
  };
}

function apiCalcMortgageBySalary(token,payload){
  _requireSession_(token);
  payload=payload||{};
  var progKey=String(payload.program||''), salary=Number(payload.salary)||0,
      members=Number(payload.members)||1, oldCredit=Number(payload.oldCredit)||0;
  var cfg=_readSettings_(), prog=cfg.programs[progKey];
  if(!prog) return {ok:false, message:'Программа не найдена'};

  var orgs=[{income:salary, oldCredit:oldCredit}];
  var calc=_calcMaxPayment_(orgs, members, progKey, cfg);

  if(!calc.approved || calc.maxPayment<=0){
    return {
      ok:true, approved:false, maxPayment:0, maxLoan:0, maxPrice:0, down:0, payment:0,
      rate:prog.variants[0]?prog.variants[0].label:'',
      members:members, oldCredit:oldCredit, pm:calc.pm, kd:calc.kd,
      totalIncome:calc.totalIncome, method1:calc.method1, method2:calc.method2
    };
  }

  var v0=prog.variants[0];
  var maxLoan=Math.round(calc.maxPayment/v0.coeff);
  var maxPrice=Math.round(maxLoan/(1-prog.downRatio));
  var down=Math.round(maxPrice*prog.downRatio);

  return {
    ok:true, approved:true, program:progKey, programName:prog.name, icon:prog.icon,
    maxPayment:calc.maxPayment, maxLoan:maxLoan, maxPrice:maxPrice, down:down,
    payment:calc.maxPayment, rate:v0.label,
    members:members, oldCredit:oldCredit, pm:calc.pm, kd:calc.kd,
    totalIncome:calc.totalIncome, method1:calc.method1, method2:calc.method2,
    downRatio:prog.downRatio
  };
}

function apiCalcBankApproval(token,payload){
  var s=_requireSession_(token);
  payload=payload||{};
  var rawOrgs=payload.orgs||[{income:Number(payload.income)||0, oldCredit:0}],
      members=Number(payload.members)||1, program=String(payload.program||'');
  var cfg=_readSettings_();

  var orgs=rawOrgs.map(function(o){
    if(o.mode==='opv' && o.opvMonths && o.opvMonths.length){
      var list=o.opvMonths.map(Number).filter(function(v){return isFinite(v)&&v>0;});
      var sals=list.map(function(v){return v/0.10;});
      return {income:sals.length?_avgSalary_(sals):0, oldCredit:Number(o.oldCredit)||0};
    }
    return {income:Number(o.income)||0, oldCredit:Number(o.oldCredit)||0};
  });

  var calc=_calcMaxPayment_(orgs, members, program, cfg);
  var allPrograms=[];
  if(calc.approved){
    Object.keys(cfg.programs).forEach(function(key){
      var p=cfg.programs[key], v0=p.variants[0];
      var maxLoan=Math.round(calc.maxPayment/v0.coeff);
      var maxPrice=Math.round(maxLoan/(1-p.downRatio));
      allPrograms.push({
        key:key, name:p.name, icon:p.icon, downRatio:p.downRatio,
        maxPrice:maxPrice, maxLoan:maxLoan,
        downPayment:Math.round(maxPrice*p.downRatio)
      });
    });
    allPrograms.sort(function(a,b){return b.maxPrice-a.maxPrice;});
  }
  _log_(s.client,s.key,'BANK',program,'inc='+calc.totalIncome+' old='+calc.totalOldCredit+' mem='+members,calc.totalIncome,calc.maxPayment);
  return {
    ok:true,
    totalIncome:calc.totalIncome, totalOldCredit:calc.totalOldCredit,
    kd:calc.kd, pm:calc.pm, mrp:cfg.mrp,
    method1:calc.method1, method2:calc.method2,
    maxPayment:calc.maxPayment, approved:calc.approved,
    members:members, allPrograms:allPrograms
  };
}

function apiGetPdfData(token,payload){
  var s=_requireSession_(token);
  return {
    ok:true,
    clientName:s.client,
    date:Utilities.formatDate(new Date(),'Asia/Almaty','dd.MM.yyyy HH:mm'),
    payload:payload
  };
}

// ============================================================
//  JSON-API для десктоп-парсера и серверного парсера на Vercel
//  Требует деплой веб-приложения с доступом «Все».
// ============================================================
function doPost(e) {
  var out;
  try {
    var body = (e && e.postData && e.postData.contents)
      ? JSON.parse(e.postData.contents) : {};
    var action = body.action || 'login';
    if (action === 'checkToken') {
      out = apiCheckToken_(String(body.token || ''));
    } else {
      out = apiLicense_(String(body.key || ''), String(body.deviceId || ''));
    }
  } catch (err) {
    out = { ok: false, reason: 'error', message: String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

// Проверка активной сессии веб-хаба (для серверного парсера)
function apiCheckToken_(token) {
  var s = _refreshSession_(token);
  if (!s) return { ok: false, message: 'Сессия истекла — войдите в сервис заново' };
  return { ok: true, clientName: s.client };
}

// Проверка ключа + срока + привязки к устройству (для десктоп-парсера)
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

  var end = _untilDate_(c.until);
  var expISO = end ? Utilities.formatDate(end, 'Asia/Almaty', 'yyyy-MM-dd') : '2999-12-31';
  var lic = _signLicense_(key, deviceId, expISO, c.name);

  return {
    ok: true, clientName: c.name, daysLeft: days,
    until: c.until ? _fmtDate_(c.until) : '',
    token: lic.token, sig: lic.sig,
    kaspiPhone: KASPI_PHONE, kaspiName: KASPI_NAME, price: SUB_PRICE
  };
}

// РАЗОВАЯ функция: записывает в лист «Программы» актуальные ставки
// и коэффициенты. Запусти один раз: выбери setActualPrograms в списке
// функций вверху редактора → «Выполнить».
//
// ВНИМАНИЕ: перезаписывает названия вариантов и коэффициенты у шести
// программ. Прежние значения пишутся в журнал выполнения.
function setActualPrograms(){
  var ACTUAL = [
    ['5050',     'Ипотека 50/50',      '🏛️', 50, '8.5% — 8 лет',       0.0080522,  '5% — 6 лет',          0.0080525,  '',                   0],
    ['3070',     'Программа 30/70',    '🏠', 30, '~10-12 лет',          0.00886788, '',                    0,          '',                   0],
    ['nauryz20', 'Наурыз (взнос 20%)', '🌸', 20, '7% — 19 лет',         0.00843335, '9% — 19 лет',         0.0101,     '',                   0],
    ['nauryz10', 'Наурыз (взнос 10%)', '🌷', 10, '7% — 19 лет',         0.00943335, '9% — 19 лет',         0.0111,     '',                   0],
    ['jasyl',    'Жасыл Ипотека',      '🌿', 20, '7% очередники',       0.00783333, '11% военные',         0.01116667, '15% все граждане',   0.01241667],
    ['askeri',   'Наурыз Аскери',      '🎖️',  0, '1-8 лет',             0.0127,     '9-19 лет',            0.00376,    '',                   0]
  ];

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  _initProgramsSheet_(ss);
  var sh = ss.getSheetByName('Программы');
  var data = sh.getDataRange().getValues();
  var updated = 0, added = 0;

  ACTUAL.forEach(function(row){
    var key = row[0], foundRow = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').trim() === key) { foundRow = i + 1; break; }
    }
    if (foundRow > 0) {
      Logger.log('БЫЛО  %s: %s', key, JSON.stringify(data[foundRow - 1].slice(0, 10)));
      sh.getRange(foundRow, 1, 1, 10).setValues([row]);
      updated++;
    } else {
      sh.appendRow(row.concat(['']));
      added++;
    }
    Logger.log('СТАЛО %s: %s', key, JSON.stringify(row));
  });

  sh.getRange(2, 6, sh.getLastRow() - 1, 5).setNumberFormat('0.0000000');
  CacheService.getScriptCache().remove('CFG_V3');
  _memCache_ = {};
  Logger.log('Готово: обновлено %s, добавлено %s. Кэш сброшен.', updated, added);
  return 'Обновлено: ' + updated + ', добавлено: ' + added;
}
