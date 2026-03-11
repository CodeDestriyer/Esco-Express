// ================================================================
// EscoExpress_Posylki.gs — CRM Посилки (менеджери)
// Живе в таблиці: Posylki_crm_v3
// Deploy: Web App → доступ "Будь-хто"
// ================================================================
// Архітектура: 4 окремі скрипти (Passengers, Posylki, Driver, Client)
// Кожен має свій doPost, свій URL, свою чергу запитів.
// ================================================================

var HEADER_ROW = 1;
var DATA_START = 2;

// ── ВСІ ТАБЛИЦІ СИСТЕМИ (SpreadsheetApp.openById) ──
var DB = {
  POSYLKI:    '1_vfEhdLEM2SVTBiu_3eDilMs1HlKxvPrJBbiHYjgrJo',
  MARHRUT:    '10SZhKV08BJyvWoMwhT0iddtWzYrDYFjCM8xgqViuE3Y',
  KLIYENTU:   '1KW2Vh_E7OxggNB_NOzWmVM8siHzHr_mG8C939YXDC38',
  FINANCE:    '1AhID7Ust45sA4PCAUjWJz515qnxzQGSj5wGQ7K8Jbu0',
  CONFIG:     '1hZ67tuQYukugO_TjNsOS3IjovBR5hWMg-JmGAq5udBE',
  ARCHIVE:    '19Ftljah5eX07RLHJaBrvYV7hStxspxcJVi6VATGZvF0'
};

// Головна таблиця цього скрипта
var SS_ID = DB.POSYLKI;

// ── АРКУШІ в Posylki_crm_v3 ──
var SHEETS = {
  PKG_UE: 'Реєстрація ТТН УК-єв',
  PKG_EU: 'Виклик Курєра ЄВ-ук',
  PHOTOS: 'Фото посилок'
};

// ── COLUMNS УК→ЄВ (52 колонки) ──
var PKG_UE_COLS = [
  'PKG_ID','Ід_смарт','Напрям','SOURCE_SHEET','Дата створення',
  'Піб відправника','Телефон реєстратора','Адреса відправки',
  'Піб отримувача','Телефон отримувача','Адреса в Європі','Внутрішній №',
  'Номер ТТН','Опис','Деталі','Кількість позицій','Кг','Оціночна вартість',
  'Сума НП','Валюта НП','Форма НП','Статус НП',
  'Сума','Валюта оплати','Завдаток','Валюта завдатку','Форма оплати',
  'Статус оплати','Борг','Примітка оплати',
  'Дата відправки','Таймінг','Номер авто','RTE_ID',
  'Дата отримання','Статус посилки','Статус ліда','Статус CRM',
  'Контроль перевірки','Дата перевірки','Фото посилки',
  'Рейтинг','Коментар рейтингу','Тег','Примітка','Примітка СМС',
  'CLI_ID','ORDER_ID','DATE_ARCHIVE','ARCHIVED_BY','ARCHIVE_REASON','ARCHIVE_ID'
];

// ── COLUMNS ЄВ→УК (51 колонка) ──
var PKG_EU_COLS = [
  'PKG_ID','Ід_смарт','Напрям','SOURCE_SHEET','Дата створення',
  'Піб відправника','Телефон реєстратора','Адреса відправки',
  'Піб отримувача','Телефон отримувача','Місто Нова Пошта','Внутрішній №',
  'Опис','Деталі','Кількість позицій','Кг','Оціночна вартість',
  'НП активна','Сума НП','Валюта НП','Статус НП',
  'Сума','Валюта оплати','Завдаток','Валюта завдатку','Форма оплати',
  'Статус оплати','Борг','Примітка оплати',
  'Дата відправки','Таймінг','Номер авто','RTE_ID',
  'Дата отримання','Статус посилки','Статус ліда','Статус CRM',
  'Контроль перевірки','Дата перевірки','Фото посилки',
  'Рейтинг','Коментар рейтингу','Тег','Примітка','Примітка СМС',
  'CLI_ID','ORDER_ID','DATE_ARCHIVE','ARCHIVED_BY','ARCHIVE_REASON','ARCHIVE_ID'
];

// ── COLUMNS Фото посилок (12 колонок) ──
var PHOTO_COLS = [
  'PHOTO_ID','PKG_ID','Номер ТТН','Штрих-код ТТН',
  'Тип фото','Фото посилки','Хто завантажив','Роль',
  'Коментар','Статус перевірки','Ід реєстратора','Час'
];

// ── FINANCE ──
var FINANCE_SHEET_NAME = 'Платежі';
var FINANCE_COLS = [
  'PAY_ID','Дата створення','Хто вніс','Роль',
  'CLI_ID','PAX_ID','PKG_ID','RTE_ID','CAL_ID',
  'Ід_смарт','Тип платежу','Сума','Валюта',
  'Форма оплати','Статус платежу','Борг сума','Борг валюта',
  'Дата погашення','Примітка','DATE_ARCHIVE','ARCHIVED_BY'
];


// ══════════════════════════════════════════════════════════════
// HELPERS
// ══════════════════════════════════════════════════════════════

function getSheet(name) {
  return SpreadsheetApp.openById(SS_ID).getSheetByName(name);
}

function getSheetFromDb(dbId, name) {
  return SpreadsheetApp.openById(dbId).getSheetByName(name);
}

function genId(prefix) {
  var d = Utilities.formatDate(new Date(), 'Europe/Kiev', 'yyyyMMdd');
  var r = Math.random().toString(36).substr(2, 4).toUpperCase();
  return prefix + '-' + d + '-' + r;
}

function now() {
  return Utilities.formatDate(new Date(), 'Europe/Kiev', 'dd.MM.yyyy HH:mm');
}

function today() {
  return Utilities.formatDate(new Date(), 'Europe/Kiev', 'dd.MM.yyyy');
}

function sheetAlias(alias) {
  if (alias === 'ue' || alias === 'ua-eu') return SHEETS.PKG_UE;
  if (alias === 'eu' || alias === 'eu-ua') return SHEETS.PKG_EU;
  return alias;
}

function resolveSheet(params) {
  if (params.sheet) return sheetAlias(params.sheet);
  if (params.pkg_id) {
    var sh1 = getSheet(SHEETS.PKG_UE);
    if (sh1 && findRow(sh1, 'PKG_ID', params.pkg_id)) return SHEETS.PKG_UE;
    var sh2 = getSheet(SHEETS.PKG_EU);
    if (sh2 && findRow(sh2, 'PKG_ID', params.pkg_id)) return SHEETS.PKG_EU;
  }
  return SHEETS.PKG_UE;
}

function isUeSheet(shName) {
  return shName === SHEETS.PKG_UE;
}

function getHeaders(sheet) {
  return sheet.getRange(HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
}

function getAllData(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < DATA_START) return { headers: getHeaders(sheet), data: [] };
  var headers = getHeaders(sheet);
  var data = sheet.getRange(DATA_START, 1, lastRow - DATA_START + 1, headers.length).getValues();
  return { headers: headers, data: data };
}

function rowToObj(headers, row) {
  var obj = {};
  for (var i = 0; i < headers.length; i++) {
    obj[headers[i]] = row[i] !== undefined ? row[i] : '';
  }
  return obj;
}

function objToRow(headers, obj) {
  return headers.map(function(h) { return obj[h] !== undefined ? obj[h] : ''; });
}

function findRow(sheet, colName, value) {
  var info = getAllData(sheet);
  var colIdx = info.headers.indexOf(colName);
  if (colIdx === -1) return null;
  for (var i = 0; i < info.data.length; i++) {
    if (String(info.data[i][colIdx]) == String(value)) {
      return { rowNum: DATA_START + i, headers: info.headers, data: info.data[i] };
    }
  }
  return null;
}

function findAllRows(sheet, colName, value) {
  var info = getAllData(sheet);
  var colIdx = info.headers.indexOf(colName);
  if (colIdx === -1) return [];
  var results = [];
  for (var i = 0; i < info.data.length; i++) {
    if (String(info.data[i][colIdx]) == String(value)) {
      results.push({ rowNum: DATA_START + i, headers: info.headers, data: info.data[i] });
    }
  }
  return results;
}

// Борг = max(0, Сума − Завдаток) для обох напрямів
// УК→ЄВ: Сума = колонка W (idx 22), Завдаток = колонка Y (idx 24)
// ЄВ→УК: Сума = колонка V (idx 21), Завдаток = колонка X (idx 23)
function calcDebt(obj) {
  var sum = parseFloat(obj['Сума']) || 0;
  var dep = parseFloat(obj['Завдаток']) || 0;
  return Math.max(0, sum - dep);
}

function pkgObjFromData(headers, data, shName, rowNum) {
  var obj = rowToObj(headers, data);
  obj._sheet = shName;
  obj._rowNum = rowNum;
  obj['Борг'] = calcDebt(obj);
  return obj;
}


// ══════════════════════════════════════════════════════════════
// 1. PARCELS — READ
// ══════════════════════════════════════════════════════════════

function apiGetAll(params) {
  var shAlias = params.sheet || 'all';
  var results = [];

  function loadSheet(name) {
    var sh = getSheet(name);
    if (!sh) return;
    var info = getAllData(sh);
    for (var i = 0; i < info.data.length; i++) {
      if (!info.data[i][0] && !info.data[i][5]) continue;
      var obj = pkgObjFromData(info.headers, info.data[i], name, DATA_START + i);

      if (params.filter) {
        if (params.filter.statusPkg && params.filter.statusPkg !== 'all') {
          if (obj['Статус посилки'] !== params.filter.statusPkg) continue;
        }
        if (params.filter.statusLid && params.filter.statusLid !== 'all') {
          if (obj['Статус ліда'] !== params.filter.statusLid) continue;
        }
        if (params.filter.statusOplata && params.filter.statusOplata !== 'all') {
          if (obj['Статус оплати'] !== params.filter.statusOplata) continue;
        }
        if (params.filter.statusCrm && params.filter.statusCrm !== 'all') {
          if (obj['Статус CRM'] !== params.filter.statusCrm) continue;
        }
        if (params.filter.tag && params.filter.tag !== 'all') {
          if (obj['Тег'] !== params.filter.tag) continue;
        }
        if (params.filter.search) {
          var s = params.filter.search.toLowerCase();
          if (String(obj['Піб відправника'] || '').toLowerCase().indexOf(s) === -1 &&
              String(obj['Піб отримувача'] || '').toLowerCase().indexOf(s) === -1 &&
              String(obj['Телефон реєстратора'] || '').indexOf(s) === -1 &&
              String(obj['Телефон отримувача'] || '').indexOf(s) === -1 &&
              String(obj['PKG_ID'] || '').toLowerCase().indexOf(s) === -1 &&
              String(obj['Номер ТТН'] || '').toLowerCase().indexOf(s) === -1) continue;
        }
      }
      results.push(obj);
    }
  }

  if (shAlias === 'all' || shAlias === 'ue') loadSheet(SHEETS.PKG_UE);
  if (shAlias === 'all' || shAlias === 'eu') loadSheet(SHEETS.PKG_EU);

  return { ok: true, count: results.length, data: results };
}

function apiGetOne(params) {
  var pkgId = params.pkg_id;
  if (!pkgId) return { ok: false, error: 'pkg_id не вказано' };

  var sheets = [SHEETS.PKG_UE, SHEETS.PKG_EU];
  for (var s = 0; s < sheets.length; s++) {
    var sh = getSheet(sheets[s]);
    if (!sh) continue;
    var found = findRow(sh, 'PKG_ID', pkgId);
    if (found) {
      var obj = pkgObjFromData(found.headers, found.data, sheets[s], found.rowNum);
      return { ok: true, data: obj };
    }
  }
  return { ok: false, error: 'Посилка не знайдена: ' + pkgId };
}

function apiGetStats(params) {
  var all = 0, ue = 0, eu = 0;
  var byStatus = {}, byPay = {}, totalDebt = 0;
  var byPkgStatus = {};

  function countSheet(name, dir) {
    var sh = getSheet(name);
    if (!sh) return;
    var info = getAllData(sh);
    var statusLidIdx = info.headers.indexOf('Статус ліда');
    var payIdx = info.headers.indexOf('Статус оплати');
    var crmIdx = info.headers.indexOf('Статус CRM');
    var pkgStatusIdx = info.headers.indexOf('Статус посилки');

    for (var i = 0; i < info.data.length; i++) {
      if (!info.data[i][0] && !info.data[i][5]) continue;
      var crm = String(info.data[i][crmIdx] || 'Активний');
      if (crm === 'Архів') continue;

      all++;
      if (dir === 'ue') ue++; else eu++;

      var st = String(info.data[i][statusLidIdx] || 'Новий');
      byStatus[st] = (byStatus[st] || 0) + 1;

      var pay = String(info.data[i][payIdx] || 'Не оплачено');
      byPay[pay] = (byPay[pay] || 0) + 1;

      var pkgSt = String(info.data[i][pkgStatusIdx] || '');
      if (pkgSt) byPkgStatus[pkgSt] = (byPkgStatus[pkgSt] || 0) + 1;

      var obj = rowToObj(info.headers, info.data[i]);
      totalDebt += calcDebt(obj);
    }
  }

  countSheet(SHEETS.PKG_UE, 'ue');
  countSheet(SHEETS.PKG_EU, 'eu');

  return {
    ok: true,
    total: all, ue: ue, eu: eu,
    byStatus: byStatus, byPay: byPay, byPkgStatus: byPkgStatus,
    totalDebt: totalDebt
  };
}


// ══════════════════════════════════════════════════════════════
// 2. PARCELS — CREATE
// ══════════════════════════════════════════════════════════════

function apiAddParcel(params) {
  var shName = sheetAlias(params.sheet || 'ue');
  var sh = getSheet(shName);
  if (!sh) return { ok: false, error: 'Аркуш не знайдено: ' + shName };
  var headers = getHeaders(sh);
  var d = params.data || {};
  var isUE = isUeSheet(shName);
  var cols = isUE ? PKG_UE_COLS : PKG_EU_COLS;

  var pkgId = genId('PKG');
  var obj = {};
  cols.forEach(function(c) { obj[c] = ''; });

  obj['PKG_ID'] = pkgId;
  obj['Дата створення'] = today();
  obj['SOURCE_SHEET'] = shName;
  obj['Напрям'] = isUE ? 'УК→ЄВ' : 'ЄВ→УК';
  obj['Піб відправника'] = d.sender || '';
  obj['Телефон реєстратора'] = d.phone || '';
  obj['Адреса відправки'] = d.addressFrom || '';
  obj['Піб отримувача'] = d.receiver || '';
  obj['Телефон отримувача'] = d.phoneReceiver || '';

  if (isUE) {
    obj['Адреса в Європі'] = d.addressTo || '';
    obj['Номер ТТН'] = d.ttn || '';
  } else {
    obj['Місто Нова Пошта'] = d.cityNP || '';
  }

  obj['Опис'] = d.description || '';
  obj['Деталі'] = d.details || '';
  obj['Кількість позицій'] = d.qty || '';
  obj['Кг'] = d.weight || '';
  obj['Оціночна вартість'] = d.estimatedValue || '';
  obj['Сума'] = d.sum || '';
  obj['Валюта оплати'] = d.currency || 'UAH';
  obj['Завдаток'] = d.deposit || '';
  obj['Валюта завдатку'] = d.currencyDeposit || d.currency || 'UAH';
  obj['Форма оплати'] = d.payForm || '';
  obj['Статус оплати'] = d.payStatus || 'Не оплачено';
  obj['Статус ліда'] = 'Новий';
  obj['Статус CRM'] = 'Активний';
  obj['Статус посилки'] = d.pkgStatus || '';
  obj['Тег'] = d.tag || '';
  obj['Примітка'] = d.note || '';

  var row = objToRow(headers, obj);
  sh.appendRow(row);

  return { ok: true, pkg_id: pkgId };
}

function apiCheckDuplicates(params) {
  function checkSheet(shName) {
    var sh = getSheet(shName);
    if (!sh) return null;
    var info = getAllData(sh);
    var pibIdx = info.headers.indexOf('Піб відправника');
    var phoneIdx = info.headers.indexOf('Телефон реєстратора');
    var idIdx = info.headers.indexOf('PKG_ID');

    var pib = (params.pib || '').toLowerCase().trim();
    var phone = (params.phone || '').trim();

    for (var i = 0; i < info.data.length; i++) {
      var rPib = String(info.data[i][pibIdx] || '').toLowerCase().trim();
      var rPhone = String(info.data[i][phoneIdx] || '').trim();

      if (rPhone === phone && rPib === pib && phone && pib) {
        return { exact: false, soft: true, match: {
          pkg_id: info.data[i][idIdx], pib: info.data[i][pibIdx], phone: rPhone
        }};
      }
    }
    return null;
  }

  var r1 = checkSheet(SHEETS.PKG_UE);
  if (r1) return r1;
  var r2 = checkSheet(SHEETS.PKG_EU);
  if (r2) return r2;
  return { exact: false, soft: false };
}


// ══════════════════════════════════════════════════════════════
// 3. PARCELS — UPDATE
// ══════════════════════════════════════════════════════════════

function apiUpdateField(params) {
  var shName = resolveSheet(params);
  var sh = getSheet(shName);
  if (!sh) return { ok: false, error: 'Аркуш не знайдено: ' + shName };
  var found = findRow(sh, 'PKG_ID', params.pkg_id);
  if (!found) return { ok: false, error: 'Запис не знайдено: ' + params.pkg_id };

  var colIdx = found.headers.indexOf(params.col);
  if (colIdx === -1) return { ok: false, error: 'Колонка не знайдена: ' + params.col };

  sh.getRange(found.rowNum, colIdx + 1).setValue(params.value);

  // Перерахунок боргу + автооновлення статусу оплати
  if (['Сума','Завдаток'].indexOf(params.col) !== -1) {
    var obj = rowToObj(found.headers, found.data);
    obj[params.col] = params.value;
    var debt = calcDebt(obj);
    var debtIdx = found.headers.indexOf('Борг');
    if (debtIdx !== -1) {
      sh.getRange(found.rowNum, debtIdx + 1).setValue(debt);
    }

    var dep = parseFloat(obj['Завдаток']) || 0;
    var sum = parseFloat(obj['Сума']) || 0;
    var newPayStatus = 'Не оплачено';
    if (dep > 0 && debt > 0) newPayStatus = 'Частково';
    if (dep > 0 && debt === 0) newPayStatus = 'Оплачено';
    var payStatusIdx = found.headers.indexOf('Статус оплати');
    if (payStatusIdx !== -1) {
      sh.getRange(found.rowNum, payStatusIdx + 1).setValue(newPayStatus);
    }

    // Автозапис платежу в Finance_crm при зміні Завдаток
    if (params.col === 'Завдаток') {
      var oldDep = parseFloat(found.data[found.headers.indexOf('Завдаток')]) || 0;
      var newDep = parseFloat(params.value) || 0;
      var delta = newDep - oldDep;
      if (delta !== 0) {
        var updatedRow = sh.getRange(found.rowNum, 1, 1, found.headers.length).getValues()[0];
        var pkgData = rowToObj(found.headers, updatedRow);
        pkgData._sheet = shName;
        addPayment(pkgData, params.manager || '', delta);
      }
    }
  }

  // При зміні статусу посилки — оновити Kliyentu якщо є ORDER_ID
  if (params.col === 'Статус посилки') {
    var orderId = String(found.data[found.headers.indexOf('ORDER_ID')] || '').trim();
    if (orderId) {
      updateKliyentuOrderStatus(orderId, params.value);
    }
  }

  return { ok: true };
}


// ══════════════════════════════════════════════════════════════
// 4. PARCELS — DELETE
// ══════════════════════════════════════════════════════════════

function apiDeleteParcel(params) {
  var shName = resolveSheet(params);
  var sh = getSheet(shName);
  if (!sh) return { ok: false, error: 'Аркуш не знайдено' };
  var found = findRow(sh, 'PKG_ID', params.pkg_id);
  if (!found) return { ok: false, error: 'Запис не знайдено' };

  sh.deleteRow(found.rowNum);
  return { ok: true };
}


// ══════════════════════════════════════════════════════════════
// 5. FINANCE — Платежі (Finance_crm_v2)
// ══════════════════════════════════════════════════════════════

function addPayment(pkgData, managerName, delta) {
  var finSS = SpreadsheetApp.openById(DB.FINANCE);
  var finSheet = finSS.getSheetByName(FINANCE_SHEET_NAME);
  if (!finSheet) {
    finSheet = finSS.insertSheet(FINANCE_SHEET_NAME);
    finSheet.getRange(1, 1, 1, FINANCE_COLS.length).setValues([FINANCE_COLS]);
  }

  var payId = genId('PAY');
  var absDelta = Math.abs(delta);
  var debt = calcDebt(pkgData);

  var payObj = {};
  FINANCE_COLS.forEach(function(c) { payObj[c] = ''; });

  payObj['PAY_ID'] = payId;
  payObj['Дата створення'] = now();
  payObj['Хто вніс'] = managerName || '';
  payObj['Роль'] = 'Менеджер';
  payObj['CLI_ID'] = pkgData['CLI_ID'] || '';
  payObj['PAX_ID'] = '';
  payObj['PKG_ID'] = pkgData['PKG_ID'] || '';
  payObj['RTE_ID'] = pkgData['RTE_ID'] || '';
  payObj['Ід_смарт'] = pkgData['Ід_смарт'] || '';
  payObj['Тип платежу'] = delta > 0 ? 'Завдаток' : 'Повернення';
  payObj['Сума'] = absDelta;
  payObj['Валюта'] = pkgData['Валюта завдатку'] || pkgData['Валюта оплати'] || 'UAH';
  payObj['Статус платежу'] = delta > 0 ? 'Отримано' : 'Повернено';
  payObj['Борг сума'] = debt;
  payObj['Борг валюта'] = pkgData['Валюта оплати'] || 'UAH';

  var finHeaders = finSheet.getRange(1, 1, 1, finSheet.getLastColumn()).getValues()[0];
  var row = finHeaders.map(function(h) { return payObj[h] !== undefined ? payObj[h] : ''; });
  finSheet.appendRow(row);

  return { ok: true, pay_id: payId };
}

function apiGetPayments(params) {
  var pkgId = params.pkg_id;
  if (!pkgId) return { ok: false, error: 'pkg_id не вказано' };

  var finSS = SpreadsheetApp.openById(DB.FINANCE);
  var finSheet = finSS.getSheetByName(FINANCE_SHEET_NAME);
  if (!finSheet) return { ok: true, data: [] };

  var lastRow = finSheet.getLastRow();
  if (lastRow < 2) return { ok: true, data: [] };

  var headers = finSheet.getRange(1, 1, 1, finSheet.getLastColumn()).getValues()[0];
  var data = finSheet.getRange(2, 1, lastRow - 1, headers.length).getValues();

  var pkgIdx = headers.indexOf('PKG_ID');
  if (pkgIdx === -1) return { ok: true, data: [] };

  var results = [];
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][pkgIdx]) === String(pkgId)) {
      var obj = {};
      for (var j = 0; j < headers.length; j++) {
        obj[headers[j]] = data[i][j] !== undefined ? data[i][j] : '';
      }
      results.push(obj);
    }
  }

  results.sort(function(a, b) {
    function toSortable(s) {
      var m = String(s).match(/(\d{2})\.(\d{2})\.(\d{4})\s*(\d{2}):?(\d{2})?/);
      if (m) return m[3] + m[2] + m[1] + m[4] + (m[5] || '00');
      return String(s);
    }
    return toSortable(b['Дата створення']).localeCompare(toSortable(a['Дата створення']));
  });

  return { ok: true, data: results };
}


// ══════════════════════════════════════════════════════════════
// 6. PHOTOS — Фото посилок
// ══════════════════════════════════════════════════════════════

function apiAddPhoto(params) {
  var sh = getSheet(SHEETS.PHOTOS);
  if (!sh) return { ok: false, error: 'Аркуш Фото посилок не знайдений' };

  var headers = getHeaders(sh);
  var photoId = genId('PHOTO');

  var obj = {};
  PHOTO_COLS.forEach(function(c) { obj[c] = ''; });

  obj['PHOTO_ID'] = photoId;
  obj['PKG_ID'] = params.pkg_id || '';
  obj['Номер ТТН'] = params.ttn || '';
  obj['Штрих-код ТТН'] = params.barcode || '';
  obj['Тип фото'] = params.type || 'Загальне';
  obj['Фото посилки'] = params.url || '';
  obj['Хто завантажив'] = params.uploaded_by || '';
  obj['Роль'] = params.role || 'Менеджер';
  obj['Коментар'] = params.comment || '';
  obj['Статус перевірки'] = '';
  obj['Ід реєстратора'] = params.registrar_id || '';
  obj['Час'] = now();

  sh.appendRow(objToRow(headers, obj));

  // Оновити поле Фото посилки в основній таблиці
  if (params.pkg_id && params.url) {
    var shName = resolveSheet({ pkg_id: params.pkg_id });
    var mainSh = getSheet(shName);
    if (mainSh) {
      var found = findRow(mainSh, 'PKG_ID', params.pkg_id);
      if (found) {
        var photoIdx = found.headers.indexOf('Фото посилки');
        if (photoIdx !== -1) {
          mainSh.getRange(found.rowNum, photoIdx + 1).setValue(params.url);
        }
      }
    }
  }

  return { ok: true, photo_id: photoId };
}

function apiGetPhotos(params) {
  var pkgId = params.pkg_id;
  if (!pkgId) return { ok: false, error: 'pkg_id не вказано' };

  var sh = getSheet(SHEETS.PHOTOS);
  if (!sh) return { ok: true, data: [] };

  var rows = findAllRows(sh, 'PKG_ID', pkgId);
  var results = [];
  for (var i = 0; i < rows.length; i++) {
    results.push(rowToObj(rows[i].headers, rows[i].data));
  }

  return { ok: true, data: results };
}


// ══════════════════════════════════════════════════════════════
// 7. ROUTE — Прив'язка посилки до маршруту
// ══════════════════════════════════════════════════════════════

function apiAddToRoute(params) {
  var pkgId = params.pkg_id;
  var sheetName = params.sheet_name;
  if (!pkgId || !sheetName) return { ok: false, error: 'pkg_id та sheet_name обов\'язкові' };

  // Отримати дані посилки
  var pkgRes = apiGetOne({ pkg_id: pkgId });
  if (!pkgRes.ok) return pkgRes;
  var pkg = pkgRes.data;

  // Записати в Marhrut_crm_v6
  var ss = SpreadsheetApp.openById(DB.MARHRUT);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { ok: false, error: 'Аркуш "' + sheetName + '" не знайдено' };

  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return { ok: false, error: 'Аркуш порожній' };
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h).trim(); });

  var lead = {};
  lead['Тип запису'] = 'Посилка';
  lead['Напрям'] = pkg['Напрям'] || '';
  lead['PAX_ID/PKG_ID'] = pkgId;
  lead['Піб відправника'] = pkg['Піб відправника'] || '';
  lead['Телефон'] = pkg['Телефон реєстратора'] || '';
  lead['Піб отримувача'] = pkg['Піб отримувача'] || '';
  lead['Телефон отримувача'] = pkg['Телефон отримувача'] || '';
  lead['Адреса'] = pkg['Адреса відправки'] || '';
  lead['Кг'] = pkg['Кг'] || '';
  lead['Сума'] = pkg['Сума'] || '';
  lead['Завдаток'] = pkg['Завдаток'] || '';
  lead['Борг'] = pkg['Борг'] || '';
  lead['Статус оплати'] = pkg['Статус оплати'] || '';

  var row = headers.map(function(h) { return lead[h] || ''; });
  sheet.appendRow(row);

  // Оновити RTE_ID в посилці
  var rteId = params.rte_id || sheetName;
  apiUpdateField({ pkg_id: pkgId, col: 'RTE_ID', value: rteId });

  if (params.auto_id) {
    apiUpdateField({ pkg_id: pkgId, col: 'Номер авто', value: params.auto_id });
  }
  if (params.date) {
    apiUpdateField({ pkg_id: pkgId, col: 'Дата відправки', value: params.date });
  }

  return { ok: true };
}

function apiRemoveFromRoute(params) {
  var pkgId = params.pkg_id;
  if (!pkgId) return { ok: false, error: 'pkg_id не вказано' };

  apiUpdateField({ pkg_id: pkgId, col: 'RTE_ID', value: '' });
  apiUpdateField({ pkg_id: pkgId, col: 'Номер авто', value: '' });
  apiUpdateField({ pkg_id: pkgId, col: 'Дата відправки', value: '' });

  return { ok: true };
}

// Отримати доступні маршрути
function apiGetRoutesList(params) {
  var ss = SpreadsheetApp.openById(DB.MARHRUT);
  var allSheets = ss.getSheets();
  var result = [];

  for (var s = 0; s < allSheets.length; s++) {
    var sheet = allSheets[s];
    var sheetName = sheet.getName();
    if (/^(Лог|Конфіг|Config|Log|Шаблон|Template)/i.test(sheetName)) continue;

    var lastRow = sheet.getLastRow();
    var rowCount = lastRow >= 2 ? lastRow - 1 : 0;
    result.push({ sheetName: sheetName, rowCount: rowCount });
  }

  return { ok: true, data: result };
}


// ══════════════════════════════════════════════════════════════
// 8. KLIYENTU — Синхронізація статусу
// ══════════════════════════════════════════════════════════════

function updateKliyentuOrderStatus(orderId, newStatus) {
  try {
    var ss = SpreadsheetApp.openById(DB.KLIYENTU);
    var sheet = ss.getSheetByName('Замовлення');
    if (!sheet) return;

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var orderIdx = headers.indexOf('ORDER_ID');
    var statusIdx = headers.indexOf('Статус посилки');
    if (orderIdx === -1 || statusIdx === -1) return;

    var data = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][orderIdx]) === String(orderId)) {
        sheet.getRange(DATA_START + i, statusIdx + 1).setValue(newStatus);
        break;
      }
    }
  } catch (e) {
    // Kliyentu може бути недоступний — не критично
  }
}

function apiGetOrderInfo(params) {
  var orderId = params.order_id;
  if (!orderId) return { ok: false, error: 'order_id не вказано' };

  try {
    var ss = SpreadsheetApp.openById(DB.KLIYENTU);
    var sheet = ss.getSheetByName('Замовлення');
    if (!sheet) return { ok: false, error: 'Аркуш Замовлення не знайдений' };

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { ok: false, error: 'Замовлення не знайдено' };

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var data = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
    var orderIdx = headers.indexOf('ORDER_ID');
    if (orderIdx === -1) return { ok: false, error: 'Колонка ORDER_ID не знайдена' };

    for (var i = 0; i < data.length; i++) {
      if (String(data[i][orderIdx]) === String(orderId)) {
        var obj = {};
        for (var j = 0; j < headers.length; j++) {
          obj[headers[j]] = data[i][j] !== undefined ? data[i][j] : '';
        }
        return { ok: true, data: obj };
      }
    }
    return { ok: false, error: 'Замовлення не знайдено: ' + orderId };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}


// ══════════════════════════════════════════════════════════════
// 9. NOVA POSHTA — Трекінг
// ══════════════════════════════════════════════════════════════

function apiTrackParcel(params) {
  var pkgId = params.pkg_id;
  var ttn = params.ttn;
  if (!ttn) return { ok: false, error: 'Номер ТТН не вказано' };

  // Перевіряємо чи є API ключ в Config
  try {
    var configSS = SpreadsheetApp.openById(DB.CONFIG);
    var configSheet = configSS.getSheetByName('Налаштування');
    if (!configSheet) return { ok: false, error: 'Конфігурація не знайдена' };

    var configData = getAllData(configSheet);
    var paramIdx = configData.headers.indexOf('Параметр');
    var valIdx = configData.headers.indexOf('Значення');
    if (paramIdx === -1 || valIdx === -1) return { ok: false, error: 'Конфігурація не налаштована' };

    var apiKey = '';
    for (var i = 0; i < configData.data.length; i++) {
      if (String(configData.data[i][paramIdx]).trim() === 'nova_poshta_api_key') {
        apiKey = String(configData.data[i][valIdx]).trim();
        break;
      }
    }

    if (!apiKey) {
      return { ok: false, error: 'API ключ Нової Пошти не налаштований', manual: true };
    }

    // Запит до API Нової Пошти
    var payload = {
      apiKey: apiKey,
      modelName: 'TrackingDocument',
      calledMethod: 'getStatusDocuments',
      methodProperties: {
        Documents: [{ DocumentNumber: ttn }]
      }
    };

    var response = UrlFetchApp.fetch('https://api.novaposhta.ua/v2.0/json/', {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    var result = JSON.parse(response.getContentText());
    if (result.success && result.data && result.data.length > 0) {
      var trackData = result.data[0];
      var status = trackData.Status || '';

      // Оновити статус в таблиці
      if (pkgId && status) {
        apiUpdateField({ pkg_id: pkgId, col: 'Статус посилки', value: status });
      }

      return {
        ok: true,
        status: status,
        statusCode: trackData.StatusCode || '',
        cityFrom: trackData.CitySender || '',
        cityTo: trackData.CityRecipient || '',
        weight: trackData.DocumentWeight || '',
        estimatedDelivery: trackData.ScheduledDeliveryDate || '',
        actualDelivery: trackData.ActualDeliveryDate || ''
      };
    }

    return { ok: false, error: 'НП не повернула дані для ТТН: ' + ttn };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// Перевірка наявності API ключа
function apiCheckNpApiKey(params) {
  try {
    var configSS = SpreadsheetApp.openById(DB.CONFIG);
    var configSheet = configSS.getSheetByName('Налаштування');
    if (!configSheet) return { ok: true, hasKey: false };

    var configData = getAllData(configSheet);
    var paramIdx = configData.headers.indexOf('Параметр');
    var valIdx = configData.headers.indexOf('Значення');
    if (paramIdx === -1 || valIdx === -1) return { ok: true, hasKey: false };

    for (var i = 0; i < configData.data.length; i++) {
      if (String(configData.data[i][paramIdx]).trim() === 'nova_poshta_api_key') {
        var key = String(configData.data[i][valIdx]).trim();
        return { ok: true, hasKey: !!key };
      }
    }
    return { ok: true, hasKey: false };
  } catch (e) {
    return { ok: true, hasKey: false };
  }
}


// ══════════════════════════════════════════════════════════════
// 10. VERIFICATION — Перевірка посилок (УК→Європа)
// ══════════════════════════════════════════════════════════════

// Статуси перевірки
var VERIFICATION_STATUSES = [
  'Нова',
  'В перевірці',
  'Готова до маршруту',
  'Відмова',
  'В маршруті',
  'Доставлено'
];

// Сканування ТТН — пошук посилки по ТТН або створення "невідомої"
function apiScanTTN(params) {
  var ttn = String(params.ttn || '').trim();
  if (!ttn) return { ok: false, error: 'ТТН не вказано' };

  var verifier = params.verifier || '';
  var sh = getSheet(SHEETS.PKG_UE);
  if (!sh) return { ok: false, error: 'Аркуш не знайдено' };

  var info = getAllData(sh);
  var ttnIdx = info.headers.indexOf('Номер ТТН');
  var dirIdx = info.headers.indexOf('Напрям');
  if (ttnIdx === -1) return { ok: false, error: 'Колонка Номер ТТН не знайдена' };

  var found = null;
  var foundRow = -1;
  for (var i = 0; i < info.data.length; i++) {
    if (String(info.data[i][ttnIdx]).trim() === ttn) {
      var dir = String(info.data[i][dirIdx] || '');
      if (dir === 'УК→ЄВ' || dir === '') {
        found = info.data[i];
        foundRow = DATA_START + i;
        break;
      }
    }
  }

  if (found) {
    // Тип А — посилка знайдена в базі
    var obj = pkgObjFromData(info.headers, found, SHEETS.PKG_UE, foundRow);

    // Оновити статус контролю перевірки та дату
    var ctrlIdx = info.headers.indexOf('Контроль перевірки');
    var dateIdx = info.headers.indexOf('Дата перевірки');
    if (ctrlIdx !== -1) sh.getRange(foundRow, ctrlIdx + 1).setValue('В перевірці');
    if (dateIdx !== -1) sh.getRange(foundRow, dateIdx + 1).setValue(now());

    // Шукаємо дублі по отримувачу
    var duplicates = findDuplicatesByRecipient(info, obj, foundRow);

    return {
      ok: true,
      found: true,
      type: 'A',
      package_id: obj['PKG_ID'],
      ttn: ttn,
      recipient_name: obj['Піб отримувача'] || '',
      recipient_phone: obj['Телефон отримувача'] || '',
      address: obj['Адреса в Європі'] || '',
      description: obj['Опис'] || '',
      weight: obj['Кг'] || '',
      estimated_value: obj['Оціночна вартість'] || '',
      internal_number: obj['Внутрішній №'] || '',
      status: 'В перевірці',
      verification_user: verifier,
      duplicates: duplicates,
      data: obj
    };
  } else {
    // Тип Б — ТТН не знайдено, створюємо нову "невідому" посилку
    var pkgId = genId('PKG');
    var cols = PKG_UE_COLS;
    var newObj = {};
    cols.forEach(function(c) { newObj[c] = ''; });

    newObj['PKG_ID'] = pkgId;
    newObj['Дата створення'] = today();
    newObj['SOURCE_SHEET'] = SHEETS.PKG_UE;
    newObj['Напрям'] = 'УК→ЄВ';
    newObj['Номер ТТН'] = ttn;
    newObj['Контроль перевірки'] = 'В перевірці';
    newObj['Дата перевірки'] = now();
    newObj['Статус ліда'] = 'Новий';
    newObj['Статус CRM'] = 'Активний';

    var headers = getHeaders(sh);
    var row = objToRow(headers, newObj);
    sh.appendRow(row);

    return {
      ok: true,
      found: false,
      type: 'B',
      package_id: pkgId,
      ttn: ttn,
      status: 'В перевірці',
      verification_user: verifier,
      duplicates: [],
      data: newObj
    };
  }
}

// Пошук дублів по отримувачу (телефон + адреса + ПІБ)
function findDuplicatesByRecipient(info, pkg, excludeRow) {
  var phone = String(pkg['Телефон отримувача'] || '').trim();
  var address = String(pkg['Адреса в Європі'] || '').trim().toLowerCase();
  var name = String(pkg['Піб отримувача'] || '').trim().toLowerCase();

  if (!phone && !address && !name) return [];

  var phoneIdx = info.headers.indexOf('Телефон отримувача');
  var addrIdx = info.headers.indexOf('Адреса в Європі');
  var nameIdx = info.headers.indexOf('Піб отримувача');
  var ttnIdx = info.headers.indexOf('Номер ТТН');
  var idIdx = info.headers.indexOf('PKG_ID');
  var crmIdx = info.headers.indexOf('Статус CRM');

  var duplicates = [];
  for (var i = 0; i < info.data.length; i++) {
    var rowNum = DATA_START + i;
    if (rowNum === excludeRow) continue;
    if (String(info.data[i][crmIdx] || '') === 'Архів') continue;

    var rPhone = String(info.data[i][phoneIdx] || '').trim();
    var rAddr = String(info.data[i][addrIdx] || '').trim().toLowerCase();
    var rName = String(info.data[i][nameIdx] || '').trim().toLowerCase();

    var match = false;
    // Збіг по телефону + адресі
    if (phone && rPhone && phone === rPhone && address && rAddr && address === rAddr) match = true;
    // Збіг по телефону + ПІБ
    if (phone && rPhone && phone === rPhone && name && rName && name === rName) match = true;
    // Збіг по адресі + ПІБ
    if (address && rAddr && address === rAddr && name && rName && name === rName) match = true;

    if (match) {
      duplicates.push({
        ttn: String(info.data[i][ttnIdx] || ''),
        package_id: String(info.data[i][idIdx] || ''),
        recipient_name: String(info.data[i][nameIdx] || ''),
        recipient_phone: rPhone
      });
    }
  }
  return duplicates;
}

// API для пошуку дублів по отримувачу
function apiFindDuplicatesByRecipient(params) {
  var pkgId = params.pkg_id;
  if (!pkgId) return { ok: false, error: 'pkg_id не вказано' };

  var sh = getSheet(SHEETS.PKG_UE);
  if (!sh) return { ok: false, error: 'Аркуш не знайдено' };

  var found = findRow(sh, 'PKG_ID', pkgId);
  if (!found) return { ok: false, error: 'Посилка не знайдена' };

  var obj = rowToObj(found.headers, found.data);
  var info = getAllData(sh);
  var duplicates = findDuplicatesByRecipient(info, obj, found.rowNum);

  return { ok: true, duplicates: duplicates };
}

// Призначення маршруту + авто-генерація Внутрішнього №
// Формат: послідовні номери — маршрут 200 → 200,201,...299,900,901,...
//                              маршрут 500 → 500,501,...599,800,801,...
function apiAssignRouteNumber(params) {
  var pkgId = params.pkg_id;
  var routeNum = String(params.route_number || '').trim();
  if (!pkgId) return { ok: false, error: 'pkg_id не вказано' };
  if (!routeNum) return { ok: false, error: 'route_number не вказано' };

  var sh = getSheet(SHEETS.PKG_UE);
  if (!sh) return { ok: false, error: 'Аркуш не знайдено' };

  var found = findRow(sh, 'PKG_ID', pkgId);
  if (!found) return { ok: false, error: 'Посилка не знайдена' };

  // Визначаємо діапазони для маршруту
  var baseNum = parseInt(routeNum);
  var overflowStart;
  if (baseNum === 200) overflowStart = 900;
  else if (baseNum === 500) overflowStart = 800;
  else overflowStart = baseNum + 100; // fallback для інших маршрутів

  var rangeEnd = baseNum + 99; // 200→299, 500→599

  // Збираємо всі існуючі внутрішні номери що належать цьому маршруту
  var info = getAllData(sh);
  var intIdx = info.headers.indexOf('Внутрішній №');
  var maxNum = baseNum - 1; // щоб перший був baseNum (200 або 500)
  if (intIdx !== -1) {
    for (var i = 0; i < info.data.length; i++) {
      var val = parseInt(String(info.data[i][intIdx] || ''));
      if (isNaN(val)) continue;
      // Перевіряємо чи номер належить цьому маршруту
      if ((val >= baseNum && val <= rangeEnd) || (val >= overflowStart && val <= overflowStart + 99)) {
        if (val > maxNum) maxNum = val;
      }
    }
  }

  // Визначаємо наступний номер
  var nextNum;
  if (maxNum < baseNum) {
    nextNum = baseNum; // перший номер в маршруті
  } else if (maxNum < rangeEnd) {
    nextNum = maxNum + 1; // ще є місце в основному діапазоні
  } else if (maxNum === rangeEnd) {
    nextNum = overflowStart; // переходимо в overflow діапазон
  } else {
    nextNum = maxNum + 1; // продовжуємо в overflow діапазоні
  }

  var internalNumber = String(nextNum);

  // Оновити поля
  var intNumIdx = found.headers.indexOf('Внутрішній №');
  if (intNumIdx !== -1) sh.getRange(found.rowNum, intNumIdx + 1).setValue(internalNumber);

  return { ok: true, internal_number: internalNumber, route_number: routeNum };
}

// Завершити перевірку — позначити посилку як "Готова до маршруту"
function apiCompleteVerification(params) {
  var pkgId = params.pkg_id;
  if (!pkgId) return { ok: false, error: 'pkg_id не вказано' };

  var sh = getSheet(SHEETS.PKG_UE);
  if (!sh) return { ok: false, error: 'Аркуш не знайдено' };

  var found = findRow(sh, 'PKG_ID', pkgId);
  if (!found) return { ok: false, error: 'Посилка не знайдена' };

  var obj = rowToObj(found.headers, found.data);

  // Перевірки обов'язкових полів
  var errors = [];
  if (!obj['Кг'] && !params.skip_validation) errors.push('Вага не заповнена');
  if (!obj['Оціночна вартість'] && !params.skip_validation) errors.push('Ціна не заповнена');
  if (!obj['Фото посилки'] && !params.skip_validation) errors.push('Фото не завантажено');
  if (!obj['Внутрішній №'] && !params.skip_validation) errors.push('Маршрут не призначено');
  if (errors.length > 0) return { ok: false, error: errors.join('; ') };

  // Оновити контроль перевірки
  var ctrlIdx = found.headers.indexOf('Контроль перевірки');
  if (ctrlIdx !== -1) sh.getRange(found.rowNum, ctrlIdx + 1).setValue('Готова до маршруту');

  // Оновити дату перевірки
  var dateIdx = found.headers.indexOf('Дата перевірки');
  if (dateIdx !== -1) sh.getRange(found.rowNum, dateIdx + 1).setValue(now());

  return { ok: true };
}

// Відмова посилки
function apiRejectVerification(params) {
  var pkgId = params.pkg_id;
  var reason = params.reason || '';
  if (!pkgId) return { ok: false, error: 'pkg_id не вказано' };

  var sh = getSheet(SHEETS.PKG_UE);
  if (!sh) return { ok: false, error: 'Аркуш не знайдено' };

  var found = findRow(sh, 'PKG_ID', pkgId);
  if (!found) return { ok: false, error: 'Посилка не знайдена' };

  var ctrlIdx = found.headers.indexOf('Контроль перевірки');
  if (ctrlIdx !== -1) sh.getRange(found.rowNum, ctrlIdx + 1).setValue('Відмова');

  var dateIdx = found.headers.indexOf('Дата перевірки');
  if (dateIdx !== -1) sh.getRange(found.rowNum, dateIdx + 1).setValue(now());

  if (reason) {
    var noteIdx = found.headers.indexOf('Примітка');
    if (noteIdx !== -1) {
      var existing = String(found.data[noteIdx] || '');
      var newNote = (existing ? existing + '; ' : '') + 'Відмова: ' + reason;
      sh.getRange(found.rowNum, noteIdx + 1).setValue(newNote);
    }
  }

  return { ok: true };
}

// Отримати статистику перевірки
function apiGetVerificationStats(params) {
  var sh = getSheet(SHEETS.PKG_UE);
  if (!sh) return { ok: true, counts: {} };

  var info = getAllData(sh);
  var ctrlIdx = info.headers.indexOf('Контроль перевірки');
  var crmIdx = info.headers.indexOf('Статус CRM');
  var dirIdx = info.headers.indexOf('Напрям');

  var counts = { 'all': 0, 'Нова': 0, 'В перевірці': 0, 'Готова до маршруту': 0, 'Відмова': 0 };

  for (var i = 0; i < info.data.length; i++) {
    if (String(info.data[i][crmIdx] || '') === 'Архів') continue;
    var dir = String(info.data[i][dirIdx] || '');
    if (dir !== 'УК→ЄВ' && dir !== '') continue;

    counts['all']++;
    var ctrl = String(info.data[i][ctrlIdx] || '').trim();
    if (!ctrl || ctrl === 'Нова') {
      counts['Нова']++;
    } else if (counts[ctrl] !== undefined) {
      counts[ctrl]++;
    }
  }

  return { ok: true, counts: counts };
}


// ══════════════════════════════════════════════════════════════
// doGet / doPost — UNIVERSAL ROUTER
// ══════════════════════════════════════════════════════════════

function doGet(e) {
  var action = (e && e.parameter) ? e.parameter.action || '' : '';
  var result = { ok: false, error: 'Unknown action' };

  try {
    switch (action) {
      case 'ping':
        result = { ok: true, message: 'EscoExpress Posylki CRM v3 API', version: '3.0', timestamp: new Date().toISOString() };
        break;
      case 'getAll':
        result = apiGetAll({ sheet: e.parameter.sheet || 'all', filter: {} });
        break;
      case 'getStats':
        result = apiGetStats({});
        break;
      default:
        result = { ok: false, error: 'Unknown GET action: ' + action };
    }
  } catch (err) {
    result = { ok: false, error: err.message };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var body = {};
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'Invalid JSON: ' + err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var action = body.action || '';
  var result = { ok: false, error: 'Unknown action: ' + action };

  try {
    switch (action) {
      // ── PARCELS READ ──
      case 'getAll':             result = apiGetAll(body); break;
      case 'getOne':             result = apiGetOne(body); break;
      case 'getStats':           result = apiGetStats(body); break;
      case 'checkDuplicates':    result = apiCheckDuplicates(body); break;

      // ── PARCELS CREATE ──
      case 'addParcel':          result = apiAddParcel(body); break;

      // ── PARCELS UPDATE ──
      case 'updateField':        result = apiUpdateField(body); break;

      // ── PARCELS DELETE ──
      case 'deleteParcel':       result = apiDeleteParcel(body); break;

      // ── FINANCE ──
      case 'getPayments':        result = apiGetPayments(body); break;

      // ── PHOTOS ──
      case 'addPhoto':           result = apiAddPhoto(body); break;
      case 'getPhotos':          result = apiGetPhotos(body); break;

      // ── ROUTES ──
      case 'addToRoute':         result = apiAddToRoute(body); break;
      case 'removeFromRoute':    result = apiRemoveFromRoute(body); break;
      case 'getRoutesList':      result = apiGetRoutesList(body); break;

      // ── KLIYENTU ──
      case 'getOrderInfo':       result = apiGetOrderInfo(body); break;

      // ── NOVA POSHTA ──
      case 'trackParcel':        result = apiTrackParcel(body); break;
      case 'checkNpApiKey':      result = apiCheckNpApiKey(body); break;

      // ── VERIFICATION (Перевірка посилок) ──
      case 'scanTTN':                  result = apiScanTTN(body); break;
      case 'findDuplicatesByRecipient': result = apiFindDuplicatesByRecipient(body); break;
      case 'assignRouteNumber':        result = apiAssignRouteNumber(body); break;
      case 'completeVerification':     result = apiCompleteVerification(body); break;
      case 'rejectVerification':       result = apiRejectVerification(body); break;
      case 'getVerificationStats':     result = apiGetVerificationStats(body); break;

      default:
        result = { ok: false, error: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { ok: false, error: err.message };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}
