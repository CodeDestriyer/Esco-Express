// ================================================================
// EscoExpress CRM v3.0 — UNIVERSAL GAS Backend (Final)
// Таблиця: Passengers_crm_v4
// ID: 1lgaCHqWBIa6oFjFWfD8m58sLwbvQjmeje2gx3YAnBCo
// ================================================================
// Цей скрипт містить ВСІ можливі дії CRM.
// Після deploy — більше не потрібно змінювати.
// ================================================================

var SS_ID = '1lgaCHqWBIa6oFjFWfD8m58sLwbvQjmeje2gx3YAnBCo';
var HEADER_ROW = 1;
var DATA_START = 2;

// ── SHEETS ──
var SHEETS = {
  PAX_UE: 'Україна-ЄВ',
  PAX_EU: 'Європа-УК',
  AUTOPARK: 'Автопарк',
  CALENDAR: 'Календар',
  SEATING: 'Розсадка по авто'
};

// ── COLUMNS ──
var PAX_COLS = [
  'PAX_ID','Ід_смарт','Напрям','SOURCE_SHEET','Дата створення',
  'Піб','Телефон пасажира','Телефон реєстратора','Кількість місць',
  'Адреса відправки','Адреса прибуття','Дата виїзду','Таймінг',
  'Номер авто','Місце в авто','RTE_ID','Ціна квитка','Валюта квитка',
  'Завдаток','Валюта завдатку','Вага багажу','Ціна багажу','Валюта багажу',
  'Борг','Статус оплати','Статус ліда','Статус CRM','Тег',
  'Примітка','Примітка СМС','CLI_ID','BOOKING_ID',
  'DATE_ARCHIVE','ARCHIVED_BY','ARCHIVE_REASON','ARCHIVE_ID','CAL_ID'
];

var AUTO_COLS = [
  'AUTO_ID','Назва авто','Держ. номер','Тип розкладки','Місткість',
  'Місце','Тип місця','Ціна UAH','Ціна CHF','Ціна EUR',
  'Ціна PLN','Ціна CZK','Ціна USD','Статус місця','Статус авто','Примітка'
];

var CAL_COLS = [
  'CAL_ID','RTE_ID','AUTO_ID','Назва авто','Тип розкладки',
  'Дата рейсу','Напрямок','Місто','Макс. місць','Вільні місця',
  'Зайняті місця','Список вільних','Список зайнятих','PAIRED_CAL_ID','Статус рейсу'
];

var SEAT_COLS = [
  'SEAT_ID','RTE_ID','CAL_ID','AUTO_ID','PAX_ID',
  'Дата','Напрям','Назва авто','Тип розкладки','Місце',
  'Тип місця','Ціна місця','Валюта','Піб','Телефон пасажира',
  'Статус','DATE_RESERVED'
];

var LAYOUTS = {
  '1-3-3': [
    {seat:'V1',type:'Водій'},
    {seat:'A1',type:'Пасажир'},{seat:'A2',type:'Пасажир'},{seat:'A3',type:'Пасажир'},
    {seat:'B1',type:'Пасажир'},{seat:'B2',type:'Пасажир'},{seat:'B3',type:'Пасажир'}
  ],
  '2-2-3': [
    {seat:'A1',type:'Пасажир'},{seat:'A2',type:'Пасажир'},
    {seat:'B1',type:'Пасажир'},{seat:'B2',type:'Пасажир'},
    {seat:'C1',type:'Пасажир'},{seat:'C2',type:'Пасажир'},{seat:'C3',type:'Пасажир'}
  ],
  '2-2-2': [
    {seat:'A1',type:'Пасажир'},{seat:'A2',type:'Пасажир'},
    {seat:'B1',type:'Пасажир'},{seat:'B2',type:'Пасажир'},
    {seat:'C1',type:'Пасажир'},{seat:'C2',type:'Пасажир'}
  ]
};


// ══════════════════════════════════════════════════════════════
// HELPERS
// ══════════════════════════════════════════════════════════════

function getSheet(name) {
  return SpreadsheetApp.openById(SS_ID).getSheetByName(name);
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
  if (alias === 'ue' || alias === 'ua-eu') return SHEETS.PAX_UE;
  if (alias === 'eu' || alias === 'eu-ua') return SHEETS.PAX_EU;
  return alias;
}

function resolveSheet(params) {
  // Спочатку пробуємо alias, потім шукаємо пасажира в обох аркушах
  if (params.sheet) return sheetAlias(params.sheet);
  if (params.pax_id) {
    var sh1 = getSheet(SHEETS.PAX_UE);
    if (sh1 && findRow(sh1, 'PAX_ID', params.pax_id)) return SHEETS.PAX_UE;
    var sh2 = getSheet(SHEETS.PAX_EU);
    if (sh2 && findRow(sh2, 'PAX_ID', params.pax_id)) return SHEETS.PAX_EU;
  }
  return SHEETS.PAX_UE;
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

function calcDebt(obj) {
  var price = parseFloat(obj['Ціна квитка']) || 0;
  var wp = parseFloat(obj['Ціна багажу']) || 0;
  var dep = parseFloat(obj['Завдаток']) || 0;
  return Math.max(0, price + wp - dep);
}

function paxObjFromData(headers, data, shName, rowNum) {
  var obj = rowToObj(headers, data);
  obj._sheet = shName;
  obj._rowNum = rowNum;
  obj['Борг'] = calcDebt(obj);
  return obj;
}


// ══════════════════════════════════════════════════════════════
// 1. PASSENGERS — READ
// ══════════════════════════════════════════════════════════════

// getAll — Отримати всіх пасажирів (з фільтрами)
function apiGetAll(params) {
  var shAlias = params.sheet || 'all';
  var results = [];

  function loadSheet(name) {
    var sh = getSheet(name);
    if (!sh) return;
    var info = getAllData(sh);
    for (var i = 0; i < info.data.length; i++) {
      if (!info.data[i][0] && !info.data[i][5]) continue;
      var obj = paxObjFromData(info.headers, info.data[i], name, DATA_START + i);

      if (params.filter) {
        if (params.filter.dir && params.filter.dir !== 'all') {
          var rawDir = String(obj['Напрям'] || '').toLowerCase();
          if (params.filter.dir === 'ua-eu' && !rawDir.match(/ук|ua/)) continue;
          if (params.filter.dir === 'eu-ua' && !rawDir.match(/єв|eu/)) continue;
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
        if (params.filter.cal_id) {
          if (params.filter.cal_id === 'none') {
            if (obj['CAL_ID'] && String(obj['CAL_ID']).trim() !== '') continue;
          } else {
            if (obj['CAL_ID'] !== params.filter.cal_id) continue;
          }
        }
        if (params.filter.date_from) {
          if (String(obj['Дата виїзду']) < params.filter.date_from) continue;
        }
        if (params.filter.date_to) {
          if (String(obj['Дата виїзду']) > params.filter.date_to) continue;
        }
        if (params.filter.search) {
          var s = params.filter.search.toLowerCase();
          if (String(obj['Піб'] || '').toLowerCase().indexOf(s) === -1 &&
              String(obj['Телефон пасажира'] || '').indexOf(s) === -1 &&
              String(obj['PAX_ID'] || '').toLowerCase().indexOf(s) === -1) continue;
        }
      }
      results.push(obj);
    }
  }

  if (shAlias === 'all' || shAlias === 'ue') loadSheet(SHEETS.PAX_UE);
  if (shAlias === 'all' || shAlias === 'eu') loadSheet(SHEETS.PAX_EU);

  return { ok: true, count: results.length, data: results };
}

// getOne — Отримати одного пасажира по PAX_ID
function apiGetOne(params) {
  var paxId = params.pax_id;
  if (!paxId) return { ok: false, error: 'pax_id не вказано' };

  var sheets = [SHEETS.PAX_UE, SHEETS.PAX_EU];
  for (var s = 0; s < sheets.length; s++) {
    var sh = getSheet(sheets[s]);
    if (!sh) continue;
    var found = findRow(sh, 'PAX_ID', paxId);
    if (found) {
      var obj = paxObjFromData(found.headers, found.data, sheets[s], found.rowNum);
      return { ok: true, data: obj };
    }
  }
  return { ok: false, error: 'Пасажир не знайдений: ' + paxId };
}

// getPassengersByTrip — Всі пасажири прив'язані до рейсу
function apiGetPassengersByTrip(params) {
  var calId = params.cal_id || '';
  if (!calId) return { ok: false, error: 'cal_id не вказано' };
  var results = [];

  [SHEETS.PAX_UE, SHEETS.PAX_EU].forEach(function(shName) {
    var sh = getSheet(shName);
    if (!sh) return;
    var info = getAllData(sh);
    var calIdx = info.headers.indexOf('CAL_ID');
    if (calIdx === -1) return;
    for (var i = 0; i < info.data.length; i++) {
      if (String(info.data[i][calIdx]) === String(calId)) {
        results.push(paxObjFromData(info.headers, info.data[i], shName, DATA_START + i));
      }
    }
  });

  return { ok: true, count: results.length, data: results };
}

// getStats — Статистика (лічильники)
function apiGetStats(params) {
  var all = 0, ue = 0, eu = 0;
  var byStatus = {}, byPay = {}, noTrip = 0, withTrip = 0;
  var totalDebt = 0;

  function countSheet(name, dir) {
    var sh = getSheet(name);
    if (!sh) return;
    var info = getAllData(sh);
    var statusIdx = info.headers.indexOf('Статус ліда');
    var payIdx = info.headers.indexOf('Статус оплати');
    var crmIdx = info.headers.indexOf('Статус CRM');
    var calIdx = info.headers.indexOf('CAL_ID');

    for (var i = 0; i < info.data.length; i++) {
      if (!info.data[i][0] && !info.data[i][5]) continue;
      var crm = String(info.data[i][crmIdx] || 'Активний');
      if (crm === 'Архів') continue;

      all++;
      if (dir === 'ue') ue++; else eu++;

      var st = String(info.data[i][statusIdx] || 'Новий');
      byStatus[st] = (byStatus[st] || 0) + 1;

      var pay = String(info.data[i][payIdx] || 'Не оплачено');
      byPay[pay] = (byPay[pay] || 0) + 1;

      var cal = String(info.data[i][calIdx] || '').trim();
      if (cal) withTrip++; else noTrip++;

      var obj = rowToObj(info.headers, info.data[i]);
      totalDebt += calcDebt(obj);
    }
  }

  countSheet(SHEETS.PAX_UE, 'ue');
  countSheet(SHEETS.PAX_EU, 'eu');

  var tripCount = 0;
  var calSheet = getSheet(SHEETS.CALENDAR);
  if (calSheet) {
    var calInfo = getAllData(calSheet);
    for (var i = 0; i < calInfo.data.length; i++) {
      if (calInfo.data[i][0]) tripCount++;
    }
  }

  return {
    ok: true,
    total: all, ue: ue, eu: eu,
    byStatus: byStatus, byPay: byPay,
    noTrip: noTrip, withTrip: withTrip,
    totalDebt: totalDebt, trips: tripCount
  };
}


// ══════════════════════════════════════════════════════════════
// 2. PASSENGERS — CREATE
// ══════════════════════════════════════════════════════════════

// addPassenger — Додати нового пасажира
function apiAddPassenger(params) {
  var shName = sheetAlias(params.sheet || 'ue');
  var sh = getSheet(shName);
  if (!sh) return { ok: false, error: 'Аркуш не знайдено: ' + shName };
  var headers = getHeaders(sh);
  var d = params.data || {};

  var paxId = genId('PAX');
  var obj = {};
  PAX_COLS.forEach(function(c) { obj[c] = ''; });

  obj['PAX_ID'] = paxId;
  obj['Дата створення'] = today();
  obj['SOURCE_SHEET'] = shName;
  obj['Напрям'] = shName === SHEETS.PAX_EU ? 'Європа-УК' : 'Україна-ЄВ';
  obj['Піб'] = d.name || '';
  obj['Телефон пасажира'] = d.phone || '';
  obj['Телефон реєстратора'] = d.phoneReg || '';
  obj['Кількість місць'] = d.seats || 1;
  obj['Адреса відправки'] = d.from || '';
  obj['Адреса прибуття'] = d.to || '';
  obj['Дата виїзду'] = d.date || '';
  obj['Таймінг'] = d.timing || '';
  obj['Ціна квитка'] = d.price || '';
  obj['Валюта квитка'] = d.currency || 'UAH';
  obj['Завдаток'] = d.deposit || '';
  obj['Валюта завдатку'] = d.currencyDeposit || d.currency || 'UAH';
  obj['Вага багажу'] = d.weight || '';
  obj['Ціна багажу'] = d.weightPrice || '';
  obj['Валюта багажу'] = d.currencyWeight || d.currency || 'UAH';
  obj['Статус оплати'] = d.payStatus || 'Не оплачено';
  obj['Статус ліда'] = d.leadStatus || 'Новий';
  obj['Статус CRM'] = 'Активний';
  obj['Тег'] = d.tag || '';
  obj['Примітка'] = d.note || '';
  obj['Примітка СМС'] = d.noteSms || '';

  var row = objToRow(headers, obj);
  sh.appendRow(row);

  // Автопідказка рейсу (suggestTrip) — якщо є дата
  var suggested = [];
  if (d.date) {
    suggested = findMatchingTrips(d.date, obj['Напрям']);
  }

  return { ok: true, pax_id: paxId, suggestedTrips: suggested };
}

// clonePassenger — Клонувати ліда (дубль для іншої дати)
function apiClonePassenger(params) {
  var paxId = params.pax_id;
  if (!paxId) return { ok: false, error: 'pax_id не вказано' };

  var shName = resolveSheet(params);
  var sh = getSheet(shName);
  if (!sh) return { ok: false, error: 'Аркуш не знайдено' };

  var found = findRow(sh, 'PAX_ID', paxId);
  if (!found) return { ok: false, error: 'Пасажир не знайдений' };

  var obj = rowToObj(found.headers, found.data);
  var newId = genId('PAX');

  // Копіюємо все крім системних
  obj['PAX_ID'] = newId;
  obj['Дата створення'] = today();
  obj['Статус ліда'] = 'Новий';
  obj['Статус оплати'] = 'Не оплачено';
  obj['Статус CRM'] = 'Активний';
  obj['CAL_ID'] = '';
  obj['Місце в авто'] = '';
  obj['Номер авто'] = '';
  obj['Завдаток'] = '';
  obj['Борг'] = '';
  obj['BOOKING_ID'] = '';
  obj['DATE_ARCHIVE'] = '';
  obj['ARCHIVED_BY'] = '';
  obj['ARCHIVE_REASON'] = '';
  obj['ARCHIVE_ID'] = '';

  // Нова дата якщо передана
  if (params.new_date) obj['Дата виїзду'] = params.new_date;

  sh.appendRow(objToRow(found.headers, obj));

  return { ok: true, pax_id: newId, cloned_from: paxId };
}

// checkDuplicates — Перевірка дублікатів
function apiCheckDuplicates(params) {
  function checkSheet(shName) {
    var sh = getSheet(shName);
    if (!sh) return null;
    var info = getAllData(sh);
    var pibIdx = info.headers.indexOf('Піб');
    var phoneIdx = info.headers.indexOf('Телефон пасажира');
    var dateIdx = info.headers.indexOf('Дата виїзду');
    var idIdx = info.headers.indexOf('PAX_ID');

    var pib = (params.pib || '').toLowerCase().trim();
    var phone = (params.phone || '').trim();
    var date = (params.date || '').trim();

    for (var i = 0; i < info.data.length; i++) {
      var rPib = String(info.data[i][pibIdx] || '').toLowerCase().trim();
      var rPhone = String(info.data[i][phoneIdx] || '').trim();
      var rDate = String(info.data[i][dateIdx] || '').trim();

      if (rPhone === phone && rPib === pib && rDate === date && phone && pib && date) {
        return { exact: true, soft: false, match: {
          pax_id: info.data[i][idIdx], pib: info.data[i][pibIdx], phone: rPhone
        }};
      }
      if (rPhone === phone && rPib === pib && phone && pib) {
        return { exact: false, soft: true, match: {
          pax_id: info.data[i][idIdx], pib: info.data[i][pibIdx], phone: rPhone
        }};
      }
    }
    return null;
  }

  var r1 = checkSheet(SHEETS.PAX_UE);
  if (r1) return r1;
  var r2 = checkSheet(SHEETS.PAX_EU);
  if (r2) return r2;
  return { exact: false, soft: false };
}


// ══════════════════════════════════════════════════════════════
// 3. PASSENGERS — UPDATE
// ══════════════════════════════════════════════════════════════

// updateField — Оновити одне поле
function apiUpdateField(params) {
  var shName = resolveSheet(params);
  var sh = getSheet(shName);
  if (!sh) return { ok: false, error: 'Аркуш не знайдено: ' + shName };
  var found = findRow(sh, 'PAX_ID', params.pax_id);
  if (!found) return { ok: false, error: 'Запис не знайдено: ' + params.pax_id };

  var colIdx = found.headers.indexOf(params.col);
  if (colIdx === -1) return { ok: false, error: 'Колонка не знайдена: ' + params.col };

  sh.getRange(found.rowNum, colIdx + 1).setValue(params.value);

  // Перерахунок боргу
  if (['Ціна квитка','Ціна багажу','Завдаток'].indexOf(params.col) !== -1) {
    var obj = rowToObj(found.headers, found.data);
    obj[params.col] = params.value;
    var debtIdx = found.headers.indexOf('Борг');
    if (debtIdx !== -1) {
      sh.getRange(found.rowNum, debtIdx + 1).setValue(calcDebt(obj));
    }
  }

  return { ok: true };
}

// updatePassenger — Оновити ВСІ поля пасажира за раз
function apiUpdatePassenger(params) {
  var shName = resolveSheet(params);
  var sh = getSheet(shName);
  if (!sh) return { ok: false, error: 'Аркуш не знайдено' };
  var found = findRow(sh, 'PAX_ID', params.pax_id);
  if (!found) return { ok: false, error: 'Запис не знайдено' };

  var obj = rowToObj(found.headers, found.data);
  var d = params.data || {};

  // Оновлюємо тільки передані поля
  var fieldMap = {
    name:'Піб', phone:'Телефон пасажира', phoneReg:'Телефон реєстратора',
    seats:'Кількість місць', from:'Адреса відправки', to:'Адреса прибуття',
    date:'Дата виїзду', timing:'Таймінг', price:'Ціна квитка', currency:'Валюта квитка',
    deposit:'Завдаток', currencyDeposit:'Валюта завдатку',
    weight:'Вага багажу', weightPrice:'Ціна багажу', currencyWeight:'Валюта багажу',
    payStatus:'Статус оплати', leadStatus:'Статус ліда', crmStatus:'Статус CRM',
    tag:'Тег', note:'Примітка', noteSms:'Примітка СМС',
    vehicle:'Номер авто', seatInCar:'Місце в авто', calId:'CAL_ID'
  };

  for (var key in d) {
    var col = fieldMap[key] || key;
    if (obj.hasOwnProperty(col)) {
      obj[col] = d[key];
    }
  }

  obj['Борг'] = calcDebt(obj);
  var row = objToRow(found.headers, obj);
  sh.getRange(found.rowNum, 1, 1, row.length).setValues([row]);

  return { ok: true };
}

// bulkUpdateField — Масове оновлення одного поля для N пасажирів
function apiBulkUpdateField(params) {
  var paxIds = params.pax_ids || [];
  var col = params.col || '';
  var value = params.value;
  if (!col) return { ok: false, error: 'Не вказано колонку' };

  var updated = 0;

  [SHEETS.PAX_UE, SHEETS.PAX_EU].forEach(function(shName) {
    var sh = getSheet(shName);
    if (!sh) return;
    var info = getAllData(sh);
    var idIdx = info.headers.indexOf('PAX_ID');
    var colIdx = info.headers.indexOf(col);
    if (colIdx === -1) return;

    for (var i = 0; i < info.data.length; i++) {
      if (paxIds.indexOf(String(info.data[i][idIdx])) !== -1) {
        sh.getRange(DATA_START + i, colIdx + 1).setValue(value);
        updated++;
      }
    }
  });

  return { ok: true, updated: updated };
}


// ══════════════════════════════════════════════════════════════
// 4. PASSENGERS — TRIP ASSIGNMENT
// ══════════════════════════════════════════════════════════════

// suggestTrips — Автопідказка рейсів по даті + напряму пасажира
function apiSuggestTrips(params) {
  var date = params.date || '';
  var direction = params.direction || '';
  if (!date) return { ok: true, data: [] };

  var suggested = findMatchingTrips(date, direction);
  return { ok: true, data: suggested };
}

// Внутрішня: пошук рейсів що збігаються по даті
function findMatchingTrips(date, direction) {
  var calSheet = getSheet(SHEETS.CALENDAR);
  if (!calSheet) return [];

  var info = getAllData(calSheet);
  var results = [];
  var dirLower = String(direction || '').toLowerCase();
  var isUE = dirLower.indexOf('ук') !== -1 || dirLower.indexOf('ua') !== -1 || dirLower.indexOf('україна') !== -1;
  var isEU = dirLower.indexOf('єв') !== -1 || dirLower.indexOf('eu') !== -1 || dirLower.indexOf('європа') !== -1;

  for (var i = 0; i < info.data.length; i++) {
    var obj = rowToObj(info.headers, info.data[i]);
    if (!obj['CAL_ID']) continue;
    if (obj['Статус рейсу'] === 'Архів' || obj['Статус рейсу'] === 'Виконано') continue;

    // Порівнюємо дату
    if (String(obj['Дата рейсу']).trim() !== String(date).trim()) continue;

    // Порівнюємо напрям (якщо вказано)
    if (direction) {
      var tDir = String(obj['Напрямок'] || '').toLowerCase();
      var tIsUE = tDir.indexOf('ук') !== -1 || tDir.indexOf('ua') !== -1 || tDir.indexOf('україна') !== -1;
      var tIsEU = tDir.indexOf('єв') !== -1 || tDir.indexOf('eu') !== -1 || tDir.indexOf('європа') !== -1;
      if (isUE && !tIsUE) continue;
      if (isEU && !tIsEU) continue;
    }

    results.push({
      cal_id: obj['CAL_ID'],
      auto_id: obj['AUTO_ID'] || '',
      auto_name: obj['Назва авто'] || '',
      layout: obj['Тип розкладки'] || '',
      date: obj['Дата рейсу'] || '',
      direction: obj['Напрямок'] || '',
      city: obj['Місто'] || '',
      max_seats: parseInt(obj['Макс. місць']) || 0,
      free_seats: parseInt(obj['Вільні місця']) || 0,
      occupied: parseInt(obj['Зайняті місця']) || 0,
      status: obj['Статус рейсу'] || ''
    });
  }

  return results;
}

// assignTrip — Призначити рейс (з валідацією місць + авто-статус)
function apiAssignTrip(params) {
  var calId = params.cal_id || '';
  var paxIds = params.pax_ids || [];
  var seatChoice = params.seat || '';  // конкретне місце або '' (вільна розсадка)
  if (!calId || paxIds.length === 0) return { ok: false, error: 'Не вказано cal_id або pax_ids' };

  // Перевіряємо рейс
  var calSheet = getSheet(SHEETS.CALENDAR);
  if (!calSheet) return { ok: false, error: 'Аркуш Календар не знайдений' };

  var calRow = findRow(calSheet, 'CAL_ID', calId);
  if (!calRow) return { ok: false, error: 'Рейс не знайдено: ' + calId };
  var calObj = rowToObj(calRow.headers, calRow.data);
  var freeSeats = parseInt(calObj['Вільні місця']);
  if (!isNaN(freeSeats) && freeSeats < paxIds.length) {
    return { ok: false, error: 'Недостатньо місць! Вільних: ' + freeSeats + ', потрібно: ' + paxIds.length };
  }

  var updated = 0;

  [SHEETS.PAX_UE, SHEETS.PAX_EU].forEach(function(shName) {
    var sh = getSheet(shName);
    if (!sh) return;
    var info = getAllData(sh);
    var idIdx = info.headers.indexOf('PAX_ID');
    var calIdx = info.headers.indexOf('CAL_ID');
    var statusIdx = info.headers.indexOf('Статус ліда');
    var seatIdx = info.headers.indexOf('Місце в авто');
    var vehicleIdx = info.headers.indexOf('Номер авто');
    if (idIdx === -1 || calIdx === -1) return;

    for (var i = 0; i < info.data.length; i++) {
      if (paxIds.indexOf(String(info.data[i][idIdx])) !== -1) {
        sh.getRange(DATA_START + i, calIdx + 1).setValue(calId);

        // Автоматичний статус
        if (statusIdx !== -1 && String(info.data[i][statusIdx]) === 'Новий') {
          sh.getRange(DATA_START + i, statusIdx + 1).setValue('В роботі');
        }

        // Записуємо авто з рейсу
        if (vehicleIdx !== -1 && calObj['Назва авто']) {
          sh.getRange(DATA_START + i, vehicleIdx + 1).setValue(calObj['Назва авто']);
        }

        // Місце: конкретне або "Вільна розсадка"
        if (seatIdx !== -1) {
          if (seatChoice) {
            sh.getRange(DATA_START + i, seatIdx + 1).setValue(seatChoice);
          } else {
            sh.getRange(DATA_START + i, seatIdx + 1).setValue('Вільна розсадка');
          }
        }

        updated++;
      }
    }
  });

  updateCalendarOccupancy(calId);

  return { ok: true, updated: updated };
}

// unassignTrip — Зняти пасажира з рейсу
function apiUnassignTrip(params) {
  var paxIds = params.pax_ids || [];
  if (params.pax_id) paxIds.push(params.pax_id);
  if (paxIds.length === 0) return { ok: false, error: 'pax_ids не вказано' };

  var affectedCalIds = {};

  [SHEETS.PAX_UE, SHEETS.PAX_EU].forEach(function(shName) {
    var sh = getSheet(shName);
    if (!sh) return;
    var info = getAllData(sh);
    var idIdx = info.headers.indexOf('PAX_ID');
    var calIdx = info.headers.indexOf('CAL_ID');
    var seatIdx = info.headers.indexOf('Місце в авто');
    var vehicleIdx = info.headers.indexOf('Номер авто');
    if (idIdx === -1 || calIdx === -1) return;

    for (var i = 0; i < info.data.length; i++) {
      if (paxIds.indexOf(String(info.data[i][idIdx])) !== -1) {
        var oldCalId = String(info.data[i][calIdx]);
        if (oldCalId) affectedCalIds[oldCalId] = true;

        sh.getRange(DATA_START + i, calIdx + 1).setValue('');
        if (seatIdx !== -1) sh.getRange(DATA_START + i, seatIdx + 1).setValue('');
        if (vehicleIdx !== -1) sh.getRange(DATA_START + i, vehicleIdx + 1).setValue('');
      }
    }
  });

  // Оновлюємо лічильники для всіх задіяних рейсів
  for (var cid in affectedCalIds) {
    updateCalendarOccupancy(cid);
  }

  return { ok: true };
}

// reassignTrip — Пересадити пасажира на інший рейс/авто
function apiReassignTrip(params) {
  var paxId = params.pax_id || '';
  var newCalId = params.new_cal_id || '';
  var newSeat = params.seat || '';
  if (!paxId || !newCalId) return { ok: false, error: 'pax_id та new_cal_id обов\'язкові' };

  // Спочатку знімаємо
  var unRes = apiUnassignTrip({ pax_ids: [paxId] });
  if (!unRes.ok) return unRes;

  // Потім призначаємо
  var asRes = apiAssignTrip({ cal_id: newCalId, pax_ids: [paxId], seat: newSeat });
  return asRes;
}

// Оновити зайнятість рейсу
function updateCalendarOccupancy(calId) {
  var calSheet = getSheet(SHEETS.CALENDAR);
  if (!calSheet) return;

  var found = findRow(calSheet, 'CAL_ID', calId);
  if (!found) return;

  var count = 0;
  var paxNames = [];

  [SHEETS.PAX_UE, SHEETS.PAX_EU].forEach(function(shName) {
    var sh = getSheet(shName);
    if (!sh) return;
    var info = getAllData(sh);
    var calIdx = info.headers.indexOf('CAL_ID');
    var pibIdx = info.headers.indexOf('Піб');
    if (calIdx === -1) return;
    for (var i = 0; i < info.data.length; i++) {
      if (String(info.data[i][calIdx]) === String(calId)) {
        count++;
        if (pibIdx !== -1 && info.data[i][pibIdx]) {
          paxNames.push(String(info.data[i][pibIdx]));
        }
      }
    }
  });

  var obj = rowToObj(found.headers, found.data);
  var maxSeats = parseInt(obj['Макс. місць']) || 0;
  var freeCount = Math.max(0, maxSeats - count);

  var occIdx = found.headers.indexOf('Зайняті місця');
  var freeIdx = found.headers.indexOf('Вільні місця');
  var occListIdx = found.headers.indexOf('Список зайнятих');
  var statusIdx = found.headers.indexOf('Статус рейсу');

  if (occIdx !== -1) calSheet.getRange(found.rowNum, occIdx + 1).setValue(count);
  if (freeIdx !== -1) calSheet.getRange(found.rowNum, freeIdx + 1).setValue(freeCount);
  if (occListIdx !== -1) calSheet.getRange(found.rowNum, occListIdx + 1).setValue(paxNames.join(', '));

  // Автоматично ставимо статус "Повний" якщо місць 0
  if (statusIdx !== -1 && freeCount <= 0 && maxSeats > 0) {
    calSheet.getRange(found.rowNum, statusIdx + 1).setValue('Повний');
  } else if (statusIdx !== -1 && freeCount > 0 && obj['Статус рейсу'] === 'Повний') {
    calSheet.getRange(found.rowNum, statusIdx + 1).setValue('Відкритий');
  }
}


// ══════════════════════════════════════════════════════════════
// 5. PASSENGERS — DELETE / ARCHIVE
// ══════════════════════════════════════════════════════════════

// deletePassenger — Повне видалення
function apiDeletePassenger(params) {
  var shName = resolveSheet(params);
  var sh = getSheet(shName);
  if (!sh) return { ok: false, error: 'Аркуш не знайдено' };
  var found = findRow(sh, 'PAX_ID', params.pax_id);
  if (!found) return { ok: false, error: 'Запис не знайдено' };

  var obj = rowToObj(found.headers, found.data);
  var calId = obj['CAL_ID'];

  sh.deleteRow(found.rowNum);

  if (calId) updateCalendarOccupancy(calId);

  return { ok: true };
}

// bulkDelete — Масове видалення
function apiBulkDelete(params) {
  var paxIds = params.pax_ids || [];
  if (paxIds.length === 0) return { ok: false, error: 'pax_ids порожній' };

  var deleted = 0;
  var calIds = {};

  // Видаляємо з кінця щоб не зсувались рядки
  [SHEETS.PAX_UE, SHEETS.PAX_EU].forEach(function(shName) {
    var sh = getSheet(shName);
    if (!sh) return;
    var info = getAllData(sh);
    var idIdx = info.headers.indexOf('PAX_ID');
    var calIdx = info.headers.indexOf('CAL_ID');

    var rowsToDelete = [];
    for (var i = 0; i < info.data.length; i++) {
      if (paxIds.indexOf(String(info.data[i][idIdx])) !== -1) {
        rowsToDelete.push(DATA_START + i);
        var c = String(info.data[i][calIdx] || '');
        if (c) calIds[c] = true;
      }
    }

    // Видаляємо з кінця
    rowsToDelete.sort(function(a,b) { return b - a; });
    for (var r = 0; r < rowsToDelete.length; r++) {
      sh.deleteRow(rowsToDelete[r]);
      deleted++;
    }
  });

  for (var cid in calIds) { updateCalendarOccupancy(cid); }

  return { ok: true, deleted: deleted };
}

// archivePassenger — Перенос в архів (статус CRM = Архів)
function apiArchivePassenger(params) {
  var paxIds = params.pax_ids || [];
  if (params.pax_id) paxIds.push(params.pax_id);
  if (paxIds.length === 0) return { ok: false, error: 'pax_ids не вказано' };

  var reason = params.reason || '';
  var archivedBy = params.archived_by || 'Менеджер';
  var archived = 0;

  [SHEETS.PAX_UE, SHEETS.PAX_EU].forEach(function(shName) {
    var sh = getSheet(shName);
    if (!sh) return;
    var info = getAllData(sh);
    var idIdx = info.headers.indexOf('PAX_ID');
    var crmIdx = info.headers.indexOf('Статус CRM');
    var dateArchIdx = info.headers.indexOf('DATE_ARCHIVE');
    var byIdx = info.headers.indexOf('ARCHIVED_BY');
    var reasonIdx = info.headers.indexOf('ARCHIVE_REASON');
    var archiveIdIdx = info.headers.indexOf('ARCHIVE_ID');

    for (var i = 0; i < info.data.length; i++) {
      if (paxIds.indexOf(String(info.data[i][idIdx])) !== -1) {
        if (crmIdx !== -1) sh.getRange(DATA_START + i, crmIdx + 1).setValue('Архів');
        if (dateArchIdx !== -1) sh.getRange(DATA_START + i, dateArchIdx + 1).setValue(now());
        if (byIdx !== -1) sh.getRange(DATA_START + i, byIdx + 1).setValue(archivedBy);
        if (reasonIdx !== -1) sh.getRange(DATA_START + i, reasonIdx + 1).setValue(reason);
        if (archiveIdIdx !== -1) sh.getRange(DATA_START + i, archiveIdIdx + 1).setValue(genId('ARC'));
        archived++;
      }
    }
  });

  return { ok: true, archived: archived };
}

// restorePassenger — Відновити з архіву
function apiRestorePassenger(params) {
  var paxIds = params.pax_ids || [];
  if (params.pax_id) paxIds.push(params.pax_id);
  if (paxIds.length === 0) return { ok: false, error: 'pax_ids не вказано' };

  var restored = 0;

  [SHEETS.PAX_UE, SHEETS.PAX_EU].forEach(function(shName) {
    var sh = getSheet(shName);
    if (!sh) return;
    var info = getAllData(sh);
    var idIdx = info.headers.indexOf('PAX_ID');
    var crmIdx = info.headers.indexOf('Статус CRM');
    var dateArchIdx = info.headers.indexOf('DATE_ARCHIVE');
    var byIdx = info.headers.indexOf('ARCHIVED_BY');
    var reasonIdx = info.headers.indexOf('ARCHIVE_REASON');

    for (var i = 0; i < info.data.length; i++) {
      if (paxIds.indexOf(String(info.data[i][idIdx])) !== -1) {
        if (crmIdx !== -1) sh.getRange(DATA_START + i, crmIdx + 1).setValue('Активний');
        if (dateArchIdx !== -1) sh.getRange(DATA_START + i, dateArchIdx + 1).setValue('');
        if (byIdx !== -1) sh.getRange(DATA_START + i, byIdx + 1).setValue('');
        if (reasonIdx !== -1) sh.getRange(DATA_START + i, reasonIdx + 1).setValue('');
        restored++;
      }
    }
  });

  return { ok: true, restored: restored };
}

// moveDirection — Перенос пасажира між аркушами UE ↔ EU
function apiMoveDirection(params) {
  var paxId = params.pax_id;
  var targetDir = params.target_dir || '';
  if (!paxId || !targetDir) return { ok: false, error: 'pax_id та target_dir обов\'язкові' };

  var fromName = targetDir === 'eu-ua' ? SHEETS.PAX_UE : SHEETS.PAX_EU;
  var toName = targetDir === 'eu-ua' ? SHEETS.PAX_EU : SHEETS.PAX_UE;

  var fromSh = getSheet(fromName);
  var toSh = getSheet(toName);
  if (!fromSh || !toSh) return { ok: false, error: 'Аркуші не знайдені' };

  var found = findRow(fromSh, 'PAX_ID', paxId);
  if (!found) return { ok: false, error: 'Пасажир не знайдений в ' + fromName };

  var obj = rowToObj(found.headers, found.data);
  obj['Напрям'] = targetDir === 'eu-ua' ? 'Європа-УК' : 'Україна-ЄВ';
  obj['SOURCE_SHEET'] = toName;

  var toHeaders = getHeaders(toSh);
  toSh.appendRow(objToRow(toHeaders, obj));
  fromSh.deleteRow(found.rowNum);

  return { ok: true, moved_to: toName };
}


// ══════════════════════════════════════════════════════════════
// 6. TRIPS — CRUD
// ══════════════════════════════════════════════════════════════

// getTrips
function apiGetTrips(params) {
  var calSheet = getSheet(SHEETS.CALENDAR);
  if (!calSheet) return { ok: true, data: [] };

  var info = getAllData(calSheet);
  var results = [];

  for (var i = 0; i < info.data.length; i++) {
    if (!info.data[i][0]) continue;
    var obj = rowToObj(info.headers, info.data[i]);

    if (params.filter) {
      if (params.filter.status && params.filter.status !== 'all') {
        if (obj['Статус рейсу'] !== params.filter.status) continue;
      }
      if (params.filter.dir && params.filter.dir !== 'all') {
        var d = String(obj['Напрямок'] || '').toLowerCase();
        if (params.filter.dir === 'ua-eu' && !d.match(/ук|ua/)) continue;
        if (params.filter.dir === 'eu-ua' && !d.match(/єв|eu/)) continue;
      }
      if (params.filter.date) {
        if (String(obj['Дата рейсу']).trim() !== String(params.filter.date).trim()) continue;
      }
      if (params.filter.auto_id) {
        if (obj['AUTO_ID'] !== params.filter.auto_id) continue;
      }
    }

    results.push({
      cal_id: obj['CAL_ID'] || '',
      rte_id: obj['RTE_ID'] || '',
      auto_id: obj['AUTO_ID'] || '',
      auto_name: obj['Назва авто'] || '',
      layout: obj['Тип розкладки'] || '',
      date: obj['Дата рейсу'] || '',
      direction: obj['Напрямок'] || '',
      city: obj['Місто'] || '',
      max_seats: parseInt(obj['Макс. місць']) || 0,
      free_seats: parseInt(obj['Вільні місця']) || 0,
      occupied: parseInt(obj['Зайняті місця']) || 0,
      free_list: obj['Список вільних'] || '',
      occupied_list: obj['Список зайнятих'] || '',
      paired_id: obj['PAIRED_CAL_ID'] || '',
      status: obj['Статус рейсу'] || 'Відкритий',
      _rowNum: DATA_START + i
    });
  }

  return { ok: true, count: results.length, data: results };
}

// getTrip — Один рейс
function apiGetTrip(params) {
  var calId = params.cal_id;
  if (!calId) return { ok: false, error: 'cal_id не вказано' };

  var calSheet = getSheet(SHEETS.CALENDAR);
  if (!calSheet) return { ok: false, error: 'Аркуш Календар не знайдений' };
  var found = findRow(calSheet, 'CAL_ID', calId);
  if (!found) return { ok: false, error: 'Рейс не знайдено' };

  var obj = rowToObj(found.headers, found.data);

  // Додатково: хто в рейсі
  var paxRes = apiGetPassengersByTrip({ cal_id: calId });

  return {
    ok: true,
    trip: {
      cal_id: obj['CAL_ID'], auto_id: obj['AUTO_ID'], auto_name: obj['Назва авто'],
      layout: obj['Тип розкладки'], date: obj['Дата рейсу'], direction: obj['Напрямок'],
      city: obj['Місто'], max_seats: parseInt(obj['Макс. місць']) || 0,
      free_seats: parseInt(obj['Вільні місця']) || 0, occupied: parseInt(obj['Зайняті місця']) || 0,
      status: obj['Статус рейсу'], paired_id: obj['PAIRED_CAL_ID'] || ''
    },
    passengers: paxRes.data || []
  };
}

// createTrip
function apiCreateTrip(params) {
  var calSheet = getSheet(SHEETS.CALENDAR);
  var autoSheet = getSheet(SHEETS.AUTOPARK);
  if (!calSheet || !autoSheet) return { ok: false, error: 'Аркуші не знайдені' };

  var calHeaders = getHeaders(calSheet);
  var autoHeaders = getHeaders(autoSheet);

  var city = params.city || '';
  var dir = params.dir || 'ua-eu';
  var vehicles = params.vehicles || [];
  var dates = params.dates || [];
  var calIds = [];

  var dirText = dir === 'eu-ua' ? 'Європа-УК' : dir === 'bt' ? 'Загальний' : 'Україна-ЄВ';

  for (var v = 0; v < vehicles.length; v++) {
    var veh = vehicles[v];
    var autoId = genId('AUTO');
    var layout = veh.layout || '1-3-3';
    var seats = parseInt(veh.seats) || 7;
    var name = veh.name || 'Авто ' + (v + 1);
    var plate = veh.plate || '';

    var seatList = [];
    if (layout === 'bus') {
      for (var s = 1; s <= seats; s++) seatList.push({ seat: String(s), type: 'Пасажир' });
    } else {
      var layoutDef = LAYOUTS[layout];
      if (layoutDef) {
        for (var s = 0; s < layoutDef.length; s++) seatList.push(layoutDef[s]);
      }
    }
    if (veh.reserve) seatList.push({ seat: 'R1', type: 'Резервне' });

    // Autopark rows
    for (var s = 0; s < seatList.length; s++) {
      var autoObj = {};
      AUTO_COLS.forEach(function(c) { autoObj[c] = ''; });
      autoObj['AUTO_ID'] = autoId;
      autoObj['Назва авто'] = name;
      autoObj['Держ. номер'] = plate;
      autoObj['Тип розкладки'] = layout;
      autoObj['Місткість'] = seats;
      autoObj['Місце'] = seatList[s].seat;
      autoObj['Тип місця'] = seatList[s].type;
      autoObj['Статус місця'] = 'Вільне';
      autoObj['Статус авто'] = 'Активний';

      // Prices if provided
      if (veh.prices) {
        if (veh.prices.UAH) autoObj['Ціна UAH'] = veh.prices.UAH;
        if (veh.prices.CHF) autoObj['Ціна CHF'] = veh.prices.CHF;
        if (veh.prices.EUR) autoObj['Ціна EUR'] = veh.prices.EUR;
        if (veh.prices.PLN) autoObj['Ціна PLN'] = veh.prices.PLN;
        if (veh.prices.CZK) autoObj['Ціна CZK'] = veh.prices.CZK;
        if (veh.prices.USD) autoObj['Ціна USD'] = veh.prices.USD;
      }

      autoSheet.appendRow(objToRow(autoHeaders, autoObj));
    }

    var freeList = seatList.filter(function(x) { return x.type !== 'Водій'; }).map(function(x) { return x.seat; }).join(', ');
    var maxPaxSeats = seatList.filter(function(x) { return x.type !== 'Водій'; }).length;

    if (dir === 'bt') {
      for (var d = 0; d < dates.length; d++) {
        var calIdUe = genId('CAL');
        var calIdEu = genId('CAL');

        function makeCalRow(cid, dirTxt, paired) {
          var o = {};
          CAL_COLS.forEach(function(c) { o[c] = ''; });
          o['CAL_ID'] = cid; o['AUTO_ID'] = autoId; o['Назва авто'] = name;
          o['Тип розкладки'] = layout; o['Дата рейсу'] = dates[d];
          o['Напрямок'] = dirTxt; o['Місто'] = city;
          o['Макс. місць'] = maxPaxSeats; o['Вільні місця'] = maxPaxSeats;
          o['Зайняті місця'] = 0; o['Список вільних'] = freeList;
          o['PAIRED_CAL_ID'] = paired; o['Статус рейсу'] = 'Відкритий';
          return o;
        }

        calSheet.appendRow(objToRow(calHeaders, makeCalRow(calIdUe, 'Україна-ЄВ', calIdEu)));
        calSheet.appendRow(objToRow(calHeaders, makeCalRow(calIdEu, 'Європа-УК', calIdUe)));
        calIds.push(calIdUe, calIdEu);
      }
    } else {
      for (var d = 0; d < dates.length; d++) {
        var calId = genId('CAL');
        var calObj = {};
        CAL_COLS.forEach(function(c) { calObj[c] = ''; });
        calObj['CAL_ID'] = calId; calObj['AUTO_ID'] = autoId; calObj['Назва авто'] = name;
        calObj['Тип розкладки'] = layout; calObj['Дата рейсу'] = dates[d];
        calObj['Напрямок'] = dirText; calObj['Місто'] = city;
        calObj['Макс. місць'] = maxPaxSeats; calObj['Вільні місця'] = maxPaxSeats;
        calObj['Зайняті місця'] = 0; calObj['Список вільних'] = freeList;
        calObj['Статус рейсу'] = 'Відкритий';
        calSheet.appendRow(objToRow(calHeaders, calObj));
        calIds.push(calId);
      }
    }
  }

  return { ok: true, cal_ids: calIds };
}

// updateTrip
function apiUpdateTrip(params) {
  var calSheet = getSheet(SHEETS.CALENDAR);
  if (!calSheet) return { ok: false, error: 'Аркуш не знайдений' };

  var found = findRow(calSheet, 'CAL_ID', params.cal_id);
  if (!found) return { ok: false, error: 'Рейс не знайдено: ' + params.cal_id };

  var obj = rowToObj(found.headers, found.data);

  if (params.city !== undefined) obj['Місто'] = params.city;
  if (params.dir) {
    if (params.dir === 'ua-eu') obj['Напрямок'] = 'Україна-ЄВ';
    else if (params.dir === 'eu-ua') obj['Напрямок'] = 'Європа-УК';
    else obj['Напрямок'] = 'Загальний';
  }
  if (params.date) obj['Дата рейсу'] = params.date;
  if (params.dates && params.dates.length > 0) obj['Дата рейсу'] = params.dates[0];
  if (params.auto_name !== undefined) obj['Назва авто'] = params.auto_name;
  if (params.status) obj['Статус рейсу'] = params.status;
  if (params.max_seats !== undefined) obj['Макс. місць'] = params.max_seats;

  var row = objToRow(found.headers, obj);
  calSheet.getRange(found.rowNum, 1, 1, row.length).setValues([row]);

  return { ok: true };
}

// archiveTrip
function apiArchiveTrip(params) {
  var calSheet = getSheet(SHEETS.CALENDAR);
  if (!calSheet) return { ok: false, error: 'Аркуш не знайдений' };

  var found = findRow(calSheet, 'CAL_ID', params.cal_id);
  if (!found) return { ok: false, error: 'Рейс не знайдено' };

  var statusIdx = found.headers.indexOf('Статус рейсу');
  if (statusIdx !== -1) calSheet.getRange(found.rowNum, statusIdx + 1).setValue('Архів');

  // Архівувати і пасажирів рейсу якщо потрібно
  if (params.archive_passengers) {
    var paxRes = apiGetPassengersByTrip({ cal_id: params.cal_id });
    if (paxRes.ok && paxRes.data.length > 0) {
      var ids = paxRes.data.map(function(p) { return p['PAX_ID']; });
      apiArchivePassenger({ pax_ids: ids, reason: 'Рейс архівовано', archived_by: params.archived_by || 'Система' });
    }
  }

  return { ok: true };
}

// deleteTrip
function apiDeleteTrip(params) {
  var calSheet = getSheet(SHEETS.CALENDAR);
  if (!calSheet) return { ok: false, error: 'Аркуш не знайдений' };

  var found = findRow(calSheet, 'CAL_ID', params.cal_id);
  if (!found) return { ok: false, error: 'Рейс не знайдено' };

  // Знімаємо пасажирів з рейсу
  clearCalIdInPassengers(params.cal_id);

  calSheet.deleteRow(found.rowNum);

  return { ok: true };
}

// duplicateTrip — Дублювання рейсу на нову дату
function apiDuplicateTrip(params) {
  var calId = params.cal_id;
  var newDates = params.dates || [];
  if (!calId || newDates.length === 0) return { ok: false, error: 'cal_id та dates обов\'язкові' };

  var calSheet = getSheet(SHEETS.CALENDAR);
  if (!calSheet) return { ok: false, error: 'Аркуш не знайдений' };

  var found = findRow(calSheet, 'CAL_ID', calId);
  if (!found) return { ok: false, error: 'Рейс не знайдено' };

  var calHeaders = getHeaders(calSheet);
  var obj = rowToObj(found.headers, found.data);
  var newCalIds = [];

  for (var i = 0; i < newDates.length; i++) {
    var newId = genId('CAL');
    var newObj = {};
    CAL_COLS.forEach(function(c) { newObj[c] = obj[c] || ''; });
    newObj['CAL_ID'] = newId;
    newObj['Дата рейсу'] = newDates[i];
    newObj['Зайняті місця'] = 0;
    newObj['Вільні місця'] = parseInt(obj['Макс. місць']) || 0;
    newObj['Список зайнятих'] = '';
    newObj['Статус рейсу'] = 'Відкритий';
    newObj['PAIRED_CAL_ID'] = '';

    calSheet.appendRow(objToRow(calHeaders, newObj));
    newCalIds.push(newId);
  }

  return { ok: true, cal_ids: newCalIds };
}

function clearCalIdInPassengers(calId) {
  [SHEETS.PAX_UE, SHEETS.PAX_EU].forEach(function(shName) {
    var sh = getSheet(shName);
    if (!sh) return;
    var info = getAllData(sh);
    var calIdx = info.headers.indexOf('CAL_ID');
    var seatIdx = info.headers.indexOf('Місце в авто');
    var vehicleIdx = info.headers.indexOf('Номер авто');
    if (calIdx === -1) return;

    for (var i = 0; i < info.data.length; i++) {
      if (String(info.data[i][calIdx]) === String(calId)) {
        sh.getRange(DATA_START + i, calIdx + 1).setValue('');
        if (seatIdx !== -1) sh.getRange(DATA_START + i, seatIdx + 1).setValue('');
        if (vehicleIdx !== -1) sh.getRange(DATA_START + i, vehicleIdx + 1).setValue('');
      }
    }
  });
}


// ══════════════════════════════════════════════════════════════
// 7. AUTOPARK
// ══════════════════════════════════════════════════════════════

// getAutopark — Список всіх авто
function apiGetAutopark(params) {
  var sh = getSheet(SHEETS.AUTOPARK);
  if (!sh) return { ok: true, data: [] };

  var info = getAllData(sh);
  var results = [];
  var autoMap = {};

  for (var i = 0; i < info.data.length; i++) {
    var obj = rowToObj(info.headers, info.data[i]);
    if (!obj['AUTO_ID']) continue;

    var aid = obj['AUTO_ID'];
    if (!autoMap[aid]) {
      autoMap[aid] = {
        auto_id: aid,
        name: obj['Назва авто'] || '',
        plate: obj['Держ. номер'] || '',
        layout: obj['Тип розкладки'] || '',
        capacity: parseInt(obj['Місткість']) || 0,
        status: obj['Статус авто'] || '',
        seats: []
      };
    }
    autoMap[aid].seats.push({
      seat: obj['Місце'] || '',
      type: obj['Тип місця'] || '',
      status: obj['Статус місця'] || '',
      prices: {
        UAH: obj['Ціна UAH'] || '',
        CHF: obj['Ціна CHF'] || '',
        EUR: obj['Ціна EUR'] || '',
        PLN: obj['Ціна PLN'] || '',
        CZK: obj['Ціна CZK'] || '',
        USD: obj['Ціна USD'] || ''
      }
    });
  }

  for (var k in autoMap) results.push(autoMap[k]);

  return { ok: true, data: results };
}

// getAutoSeats — Місця конкретного авто (для вибору місця менеджером)
function apiGetAutoSeats(params) {
  var autoId = params.auto_id;
  if (!autoId) return { ok: false, error: 'auto_id не вказано' };

  var sh = getSheet(SHEETS.AUTOPARK);
  if (!sh) return { ok: false, error: 'Аркуш Автопарк не знайдений' };

  var rows = findAllRows(sh, 'AUTO_ID', autoId);
  var seats = [];

  for (var i = 0; i < rows.length; i++) {
    var obj = rowToObj(rows[i].headers, rows[i].data);
    seats.push({
      seat: obj['Місце'] || '',
      type: obj['Тип місця'] || '',
      status: obj['Статус місця'] || '',
      prices: {
        UAH: obj['Ціна UAH'] || '',
        CHF: obj['Ціна CHF'] || '',
        EUR: obj['Ціна EUR'] || ''
      }
    });
  }

  return { ok: true, auto_id: autoId, seats: seats };
}


// ══════════════════════════════════════════════════════════════
// 8. SEATING — Розсадка по авто
// ══════════════════════════════════════════════════════════════

// getSeating — Розсадка для конкретного рейсу
function apiGetSeating(params) {
  var calId = params.cal_id;
  if (!calId) return { ok: false, error: 'cal_id не вказано' };

  var sh = getSheet(SHEETS.SEATING);
  if (!sh) return { ok: true, data: [] };

  var rows = findAllRows(sh, 'CAL_ID', calId);
  var results = [];
  for (var i = 0; i < rows.length; i++) {
    results.push(rowToObj(rows[i].headers, rows[i].data));
  }

  return { ok: true, data: results };
}

// assignSeat — Конкретне місце пасажиру (менеджер обрав)
function apiAssignSeat(params) {
  var calId = params.cal_id;
  var paxId = params.pax_id;
  var seat = params.seat;
  if (!calId || !paxId || !seat) return { ok: false, error: 'cal_id, pax_id та seat обов\'язкові' };

  // Отримуємо дані рейсу
  var calSheet = getSheet(SHEETS.CALENDAR);
  if (!calSheet) return { ok: false, error: 'Календар не знайдений' };
  var calRow = findRow(calSheet, 'CAL_ID', calId);
  if (!calRow) return { ok: false, error: 'Рейс не знайдено' };
  var calObj = rowToObj(calRow.headers, calRow.data);

  // Отримуємо дані пасажира
  var paxData = apiGetOne({ pax_id: paxId });
  if (!paxData.ok) return paxData;

  // Записуємо в розсадку
  var seatSheet = getSheet(SHEETS.SEATING);
  if (seatSheet) {
    var seatHeaders = getHeaders(seatSheet);
    var seatObj = {};
    SEAT_COLS.forEach(function(c) { seatObj[c] = ''; });
    seatObj['SEAT_ID'] = genId('SEAT');
    seatObj['CAL_ID'] = calId;
    seatObj['AUTO_ID'] = calObj['AUTO_ID'] || '';
    seatObj['PAX_ID'] = paxId;
    seatObj['Дата'] = calObj['Дата рейсу'] || '';
    seatObj['Напрям'] = calObj['Напрямок'] || '';
    seatObj['Назва авто'] = calObj['Назва авто'] || '';
    seatObj['Тип розкладки'] = calObj['Тип розкладки'] || '';
    seatObj['Місце'] = seat;
    seatObj['Піб'] = paxData.data['Піб'] || '';
    seatObj['Телефон пасажира'] = paxData.data['Телефон пасажира'] || '';
    seatObj['Статус'] = 'Зайняте';
    seatObj['DATE_RESERVED'] = now();
    seatSheet.appendRow(objToRow(seatHeaders, seatObj));
  }

  // Оновлюємо поле "Місце в авто" у пасажира
  apiUpdateField({ pax_id: paxId, col: 'Місце в авто', value: seat });

  return { ok: true, seat_id: seatObj['SEAT_ID'] };
}

// freeSeat — Звільнити місце
function apiFreeSeat(params) {
  var seatId = params.seat_id;
  if (!seatId) return { ok: false, error: 'seat_id не вказано' };

  var sh = getSheet(SHEETS.SEATING);
  if (!sh) return { ok: false, error: 'Аркуш Розсадка не знайдений' };

  var found = findRow(sh, 'SEAT_ID', seatId);
  if (!found) return { ok: false, error: 'Місце не знайдено' };

  var obj = rowToObj(found.headers, found.data);
  var paxId = obj['PAX_ID'];

  sh.deleteRow(found.rowNum);

  // Очистити "Місце в авто" у пасажира
  if (paxId) {
    apiUpdateField({ pax_id: paxId, col: 'Місце в авто', value: '' });
  }

  return { ok: true };
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
        result = { ok: true, message: 'EscoExpress CRM v3 API', version: '3.0', timestamp: new Date().toISOString() };
        break;
      case 'getAll':
        result = apiGetAll({ sheet: e.parameter.sheet || 'all', filter: {} });
        break;
      case 'getTrips':
        result = apiGetTrips({ filter: {} });
        break;
      case 'getStats':
        result = apiGetStats({});
        break;
      case 'getAutopark':
        result = apiGetAutopark({});
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
      // ── PASSENGERS READ ──
      case 'getAll':             result = apiGetAll(body); break;
      case 'getOne':             result = apiGetOne(body); break;
      case 'getPassengersByTrip':result = apiGetPassengersByTrip(body); break;
      case 'getStats':           result = apiGetStats(body); break;
      case 'checkDuplicates':    result = apiCheckDuplicates(body); break;
      case 'suggestTrips':       result = apiSuggestTrips(body); break;

      // ── PASSENGERS CREATE ──
      case 'addPassenger':       result = apiAddPassenger(body); break;
      case 'clonePassenger':     result = apiClonePassenger(body); break;

      // ── PASSENGERS UPDATE ──
      case 'updateField':        result = apiUpdateField(body); break;
      case 'updatePassenger':    result = apiUpdatePassenger(body); break;
      case 'bulkUpdateField':    result = apiBulkUpdateField(body); break;

      // ── PASSENGERS TRIP ──
      case 'assignTrip':         result = apiAssignTrip(body); break;
      case 'unassignTrip':       result = apiUnassignTrip(body); break;
      case 'reassignTrip':       result = apiReassignTrip(body); break;

      // ── PASSENGERS DELETE/ARCHIVE ──
      case 'deletePassenger':    result = apiDeletePassenger(body); break;
      case 'bulkDelete':         result = apiBulkDelete(body); break;
      case 'archivePassenger':   result = apiArchivePassenger(body); break;
      case 'restorePassenger':   result = apiRestorePassenger(body); break;
      case 'moveDirection':      result = apiMoveDirection(body); break;

      // ── TRIPS ──
      case 'getTrips':           result = apiGetTrips(body); break;
      case 'getTrip':            result = apiGetTrip(body); break;
      case 'createTrip':         result = apiCreateTrip(body); break;
      case 'updateTrip':         result = apiUpdateTrip(body); break;
      case 'archiveTrip':        result = apiArchiveTrip(body); break;
      case 'deleteTrip':         result = apiDeleteTrip(body); break;
      case 'duplicateTrip':      result = apiDuplicateTrip(body); break;

      // ── AUTOPARK ──
      case 'getAutopark':        result = apiGetAutopark(body); break;
      case 'getAutoSeats':       result = apiGetAutoSeats(body); break;

      // ── SEATING ──
      case 'getSeating':         result = apiGetSeating(body); break;
      case 'assignSeat':         result = apiAssignSeat(body); break;
      case 'freeSeat':           result = apiFreeSeat(body); break;

      default:
        result = { ok: false, error: 'Unknown action: ' + action + '. Available: getAll, getOne, getPassengersByTrip, getStats, checkDuplicates, suggestTrips, addPassenger, clonePassenger, updateField, updatePassenger, bulkUpdateField, assignTrip, unassignTrip, reassignTrip, deletePassenger, bulkDelete, archivePassenger, restorePassenger, moveDirection, getTrips, getTrip, createTrip, updateTrip, archiveTrip, deleteTrip, duplicateTrip, getAutopark, getAutoSeats, getSeating, assignSeat, freeSeat' };
    }
  } catch (err) {
    result = { ok: false, error: err.message };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}
