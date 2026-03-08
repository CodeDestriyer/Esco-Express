// ================================================================
// EscoExpress CRM v2.0 — GAS Backend
// Таблиця: Passengers_crm_v4
// ID: 1lgaCHqWBIa6oFjFWfD8m58sLwbvQjmeje2gx3YAnBCo
// HEADER_ROW: 1 | DATA_START_ROW: 2
// ================================================================

const SS_ID = '1lgaCHqWBIa6oFjFWfD8m58sLwbvQjmeje2gx3YAnBCo';
const HEADER_ROW = 1;
const DATA_START = 2;

// Назви аркушів
const SHEETS = {
  PAX_UE: 'Україна-ЄВ',
  PAX_EU: 'Європа-УК',
  AUTOPARK: 'Автопарк',
  CALENDAR: 'Календар',
  SEATING: 'Розсадка по авто'
};

// 37 колонок пасажирів (A–AK)
const PAX_COLS = [
  'PAX_ID','Ід_смарт','Напрям','SOURCE_SHEET','Дата створення',
  'Піб','Телефон пасажира','Телефон реєстратора','Кількість місць',
  'Адреса відправки','Адреса прибуття','Дата виїзду','Таймінг',
  'Номер авто','Місце в авто','RTE_ID','Ціна квитка','Валюта квитка',
  'Завдаток','Валюта завдатку','Вага багажу','Ціна багажу','Валюта багажу',
  'Борг','Статус оплати','Статус ліда','Статус CRM','Тег',
  'Примітка','Примітка СМС','CLI_ID','BOOKING_ID',
  'DATE_ARCHIVE','ARCHIVED_BY','ARCHIVE_REASON','ARCHIVE_ID','CAL_ID'
];

// 16 колонок автопарку
const AUTO_COLS = [
  'AUTO_ID','Назва авто','Держ. номер','Тип розкладки','Місткість',
  'Місце','Тип місця','Ціна UAH','Ціна CHF','Ціна EUR',
  'Ціна PLN','Ціна CZK','Ціна USD','Статус місця','Статус авто','Примітка'
];

// 15 колонок календаря
const CAL_COLS = [
  'CAL_ID','RTE_ID','AUTO_ID','Назва авто','Тип розкладки',
  'Дата рейсу','Напрямок','Місто','Макс. місць','Вільні місця',
  'Зайняті місця','Список вільних','Список зайнятих','PAIRED_CAL_ID','Статус рейсу'
];

// 17 колонок розсадки
const SEAT_COLS = [
  'SEAT_ID','RTE_ID','CAL_ID','AUTO_ID','PAX_ID',
  'Дата','Напрям','Назва авто','Тип розкладки','Місце',
  'Тип місця','Ціна місця','Валюта','Піб','Телефон пасажира',
  'Статус','DATE_RESERVED'
];

// Типи розкладок
const LAYOUTS = {
  '1-3-3': [
    {seat:'V1', type:'Водій'},
    {seat:'A1', type:'Пасажир'},{seat:'A2', type:'Пасажир'},{seat:'A3', type:'Пасажир'},
    {seat:'B1', type:'Пасажир'},{seat:'B2', type:'Пасажир'},{seat:'B3', type:'Пасажир'}
  ],
  '2-2-3': [
    {seat:'A1', type:'Пасажир'},{seat:'A2', type:'Пасажир'},
    {seat:'B1', type:'Пасажир'},{seat:'B2', type:'Пасажир'},
    {seat:'C1', type:'Пасажир'},{seat:'C2', type:'Пасажир'},{seat:'C3', type:'Пасажир'}
  ],
  '2-2-2': [
    {seat:'A1', type:'Пасажир'},{seat:'A2', type:'Пасажир'},
    {seat:'B1', type:'Пасажир'},{seat:'B2', type:'Пасажир'},
    {seat:'C1', type:'Пасажир'},{seat:'C2', type:'Пасажир'}
  ]
};


// ════════════════════════════════════════════
// HELPERS
// ════════════════════════════════════════════

function getSheet(name) {
  return SpreadsheetApp.openById(SS_ID).getSheetByName(name);
}

function genId(prefix) {
  var d = Utilities.formatDate(new Date(), 'Europe/Kiev', 'yyyyMMdd');
  var r = Math.random().toString(36).substr(2, 3).toUpperCase();
  return prefix + '-' + d + '-' + r;
}

function sheetAlias(alias) {
  if (alias === 'ue' || alias === 'ua-eu') return SHEETS.PAX_UE;
  if (alias === 'eu' || alias === 'eu-ua') return SHEETS.PAX_EU;
  return alias;
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


// ════════════════════════════════════════════
// API — PASSENGERS
// ════════════════════════════════════════════

function apiGetAll(params) {
  var sheetAlias_ = params.sheet || 'all';
  var results = [];

  function loadSheet(name) {
    var sh = getSheet(name);
    if (!sh) return;
    var info = getAllData(sh);
    for (var i = 0; i < info.data.length; i++) {
      if (!info.data[i][0] && !info.data[i][5]) continue; // skip empty rows
      var obj = rowToObj(info.headers, info.data[i]);
      obj._rowNum = DATA_START + i;
      obj._sheet = name;

      // Apply filters
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
        if (params.filter.cal_id) {
          if (obj['CAL_ID'] !== params.filter.cal_id) continue;
        }
        if (params.filter.search) {
          var s = params.filter.search.toLowerCase();
          if (String(obj['Піб'] || '').toLowerCase().indexOf(s) === -1 &&
              String(obj['Телефон пасажира'] || '').indexOf(s) === -1) continue;
        }
      }

      // Calc debt
      var price = parseFloat(obj['Ціна квитка']) || 0;
      var wp = parseFloat(obj['Ціна багажу']) || 0;
      var dep = parseFloat(obj['Завдаток']) || 0;
      obj['Борг'] = Math.max(0, price + wp - dep);

      results.push(obj);
    }
  }

  if (sheetAlias_ === 'all' || sheetAlias_ === 'ue') loadSheet(SHEETS.PAX_UE);
  if (sheetAlias_ === 'all' || sheetAlias_ === 'eu') loadSheet(SHEETS.PAX_EU);

  return { ok: true, data: results };
}

function apiAddPassenger(params) {
  var shName = sheetAlias(params.sheet || 'ue');
  var sh = getSheet(shName);
  var headers = getHeaders(sh);
  var d = params.data || {};

  var paxId = genId('PAX');
  var obj = {};
  PAX_COLS.forEach(function(c) { obj[c] = ''; });

  obj['PAX_ID'] = paxId;
  obj['Дата створення'] = Utilities.formatDate(new Date(), 'Europe/Kiev', 'dd.MM.yyyy');
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
  obj['Валюта завдатку'] = d.currency || 'UAH';
  obj['Вага багажу'] = d.weight || '';
  obj['Ціна багажу'] = d.weightPrice || '';
  obj['Валюта багажу'] = d.currency || 'UAH';
  obj['Статус оплати'] = 'Не оплачено';
  obj['Статус ліда'] = 'Новий';
  obj['Статус CRM'] = 'Активний';
  obj['Примітка'] = d.note || '';
  // CAL_ID, RTE_ID, CLI_ID, BOOKING_ID — порожньо

  var row = objToRow(headers, obj);
  sh.appendRow(row);

  return { ok: true, pax_id: paxId };
}

function apiUpdateField(params) {
  var shName = sheetAlias(params.sheet || 'ue');
  var sh = getSheet(shName);
  var found = findRow(sh, 'PAX_ID', params.pax_id);
  if (!found) return { ok: false, error: 'Запис не знайдено: ' + params.pax_id };

  var colIdx = found.headers.indexOf(params.col);
  if (colIdx === -1) return { ok: false, error: 'Колонка не знайдена: ' + params.col };

  sh.getRange(found.rowNum, colIdx + 1).setValue(params.value);

  // Recalc debt if financial field changed
  if (['Ціна квитка','Ціна багажу','Завдаток'].indexOf(params.col) !== -1) {
    var obj = rowToObj(found.headers, found.data);
    obj[params.col] = params.value;
    var price = parseFloat(obj['Ціна квитка']) || 0;
    var wp = parseFloat(obj['Ціна багажу']) || 0;
    var dep = parseFloat(obj['Завдаток']) || 0;
    var debtIdx = found.headers.indexOf('Борг');
    if (debtIdx !== -1) {
      sh.getRange(found.rowNum, debtIdx + 1).setValue(Math.max(0, price + wp - dep));
    }
  }

  return { ok: true };
}

function apiDeletePassenger(params) {
  var shName = sheetAlias(params.sheet || 'ue');
  var sh = getSheet(shName);
  var found = findRow(sh, 'PAX_ID', params.pax_id);
  if (!found) return { ok: false, error: 'Запис не знайдено' };

  sh.deleteRow(found.rowNum);
  return { ok: true };
}

function apiCheckDuplicates(params) {
  var shName = sheetAlias(params.sheet || 'ue');
  var sh = getSheet(shName);
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
  return { exact: false, soft: false };
}


// ════════════════════════════════════════════
// API — TRIPS (Calendar + Autopark)
// ════════════════════════════════════════════

function apiGetTrips(params) {
  var calSheet = getSheet(SHEETS.CALENDAR);
  if (!calSheet) return { ok: true, data: [] };

  var info = getAllData(calSheet);
  var results = [];

  for (var i = 0; i < info.data.length; i++) {
    if (!info.data[i][0]) continue; // skip empty
    var obj = rowToObj(info.headers, info.data[i]);

    // Apply filters
    if (params.filter) {
      if (params.filter.status && obj['Статус рейсу'] !== params.filter.status) continue;
      if (params.filter.dir && params.filter.dir !== 'all') {
        var d = String(obj['Напрямок'] || '').toLowerCase();
        if (params.filter.dir === 'ua-eu' && !d.match(/ук|ua/)) continue;
        if (params.filter.dir === 'eu-ua' && !d.match(/єв|eu/)) continue;
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
      max_seats: obj['Макс. місць'] || 0,
      free_seats: obj['Вільні місця'] || 0,
      occupied: obj['Зайняті місця'] || 0,
      free_list: obj['Список вільних'] || '',
      occupied_list: obj['Список зайнятих'] || '',
      paired_id: obj['PAIRED_CAL_ID'] || '',
      status: obj['Статус рейсу'] || 'Відкритий',
      _rowNum: DATA_START + i
    });
  }

  return { ok: true, data: results };
}

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

  // Determine direction text
  var dirText = dir === 'eu-ua' ? 'Європа-УК' : dir === 'bt' ? 'Загальний' : 'Україна-ЄВ';

  for (var v = 0; v < vehicles.length; v++) {
    var veh = vehicles[v];
    var autoId = genId('AUTO');
    var layout = veh.layout || '1-3-3';
    var seats = parseInt(veh.seats) || 7;
    var name = veh.name || 'Авто ' + (v + 1);
    var plate = veh.plate || '';

    // Generate seats in Autopark
    var seatList = [];
    if (layout === 'bus') {
      for (var s = 1; s <= seats; s++) {
        seatList.push({ seat: String(s), type: 'Пасажир' });
      }
    } else {
      var layoutDef = LAYOUTS[layout];
      if (layoutDef) {
        for (var s = 0; s < layoutDef.length; s++) {
          seatList.push(layoutDef[s]);
        }
      }
    }
    // Add reserve seat if requested
    if (veh.reserve) {
      seatList.push({ seat: 'R1', type: 'Резервне' });
    }

    // Write autopark rows
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
      autoSheet.appendRow(objToRow(autoHeaders, autoObj));
    }

    // Generate free list
    var freeList = seatList.filter(function(x) { return x.type !== 'Водій'; }).map(function(x) { return x.seat; }).join(', ');
    var maxPaxSeats = seatList.filter(function(x) { return x.type !== 'Водій'; }).length;

    // Create calendar entries (one per date)
    if (dir === 'bt') {
      // Загальний: два рядки per date (UE + EU) linked via PAIRED_CAL_ID
      for (var d = 0; d < dates.length; d++) {
        var calIdUe = genId('CAL');
        var calIdEu = genId('CAL');

        var calObjUe = {};
        CAL_COLS.forEach(function(c) { calObjUe[c] = ''; });
        calObjUe['CAL_ID'] = calIdUe;
        calObjUe['AUTO_ID'] = autoId;
        calObjUe['Назва авто'] = name;
        calObjUe['Тип розкладки'] = layout;
        calObjUe['Дата рейсу'] = dates[d];
        calObjUe['Напрямок'] = 'Україна-ЄВ';
        calObjUe['Місто'] = city;
        calObjUe['Макс. місць'] = maxPaxSeats;
        calObjUe['Вільні місця'] = maxPaxSeats;
        calObjUe['Зайняті місця'] = 0;
        calObjUe['Список вільних'] = freeList;
        calObjUe['PAIRED_CAL_ID'] = calIdEu;
        calObjUe['Статус рейсу'] = 'Відкритий';
        calSheet.appendRow(objToRow(calHeaders, calObjUe));

        var calObjEu = {};
        CAL_COLS.forEach(function(c) { calObjEu[c] = ''; });
        calObjEu['CAL_ID'] = calIdEu;
        calObjEu['AUTO_ID'] = autoId;
        calObjEu['Назва авто'] = name;
        calObjEu['Тип розкладки'] = layout;
        calObjEu['Дата рейсу'] = dates[d];
        calObjEu['Напрямок'] = 'Європа-УК';
        calObjEu['Місто'] = city;
        calObjEu['Макс. місць'] = maxPaxSeats;
        calObjEu['Вільні місця'] = maxPaxSeats;
        calObjEu['Зайняті місця'] = 0;
        calObjEu['Список вільних'] = freeList;
        calObjEu['PAIRED_CAL_ID'] = calIdUe;
        calObjEu['Статус рейсу'] = 'Відкритий';
        calSheet.appendRow(objToRow(calHeaders, calObjEu));

        calIds.push(calIdUe, calIdEu);
      }
    } else {
      // Single direction
      for (var d = 0; d < dates.length; d++) {
        var calId = genId('CAL');
        var calObj = {};
        CAL_COLS.forEach(function(c) { calObj[c] = ''; });
        calObj['CAL_ID'] = calId;
        calObj['AUTO_ID'] = autoId;
        calObj['Назва авто'] = name;
        calObj['Тип розкладки'] = layout;
        calObj['Дата рейсу'] = dates[d];
        calObj['Напрямок'] = dirText;
        calObj['Місто'] = city;
        calObj['Макс. місць'] = maxPaxSeats;
        calObj['Вільні місця'] = maxPaxSeats;
        calObj['Зайняті місця'] = 0;
        calObj['Список вільних'] = freeList;
        calObj['Статус рейсу'] = 'Відкритий';
        calSheet.appendRow(objToRow(calHeaders, calObj));
        calIds.push(calId);
      }
    }
  }

  return { ok: true, cal_ids: calIds };
}

function apiUpdateTrip(params) {
  var calSheet = getSheet(SHEETS.CALENDAR);
  if (!calSheet) return { ok: false, error: 'Аркуш Календар не знайдений' };

  var found = findRow(calSheet, 'CAL_ID', params.cal_id);
  if (!found) return { ok: false, error: 'Рейс не знайдено: ' + params.cal_id };

  var obj = rowToObj(found.headers, found.data);

  // Update fields
  if (params.city) obj['Місто'] = params.city;
  if (params.dir) {
    if (params.dir === 'ua-eu') obj['Напрямок'] = 'Україна-ЄВ';
    else if (params.dir === 'eu-ua') obj['Напрямок'] = 'Європа-УК';
    else obj['Напрямок'] = 'Загальний';
  }
  if (params.dates && params.dates.length > 0) {
    obj['Дата рейсу'] = params.dates[0];
  }
  if (params.vehicles && params.vehicles.length > 0) {
    var v = params.vehicles[0];
    if (v.name) obj['Назва авто'] = v.name;
    if (v.layout) obj['Тип розкладки'] = v.layout;
    if (v.seats) obj['Макс. місць'] = v.seats;
  }

  var row = objToRow(found.headers, obj);
  calSheet.getRange(found.rowNum, 1, 1, row.length).setValues([row]);

  return { ok: true };
}

function apiArchiveTrip(params) {
  var calSheet = getSheet(SHEETS.CALENDAR);
  if (!calSheet) return { ok: false, error: 'Аркуш не знайдений' };

  var found = findRow(calSheet, 'CAL_ID', params.cal_id);
  if (!found) return { ok: false, error: 'Рейс не знайдено' };

  var statusIdx = found.headers.indexOf('Статус рейсу');
  if (statusIdx !== -1) {
    calSheet.getRange(found.rowNum, statusIdx + 1).setValue('Архів');
  }

  return { ok: true };
}

function apiDeleteTrip(params) {
  var calSheet = getSheet(SHEETS.CALENDAR);
  if (!calSheet) return { ok: false, error: 'Аркуш не знайдений' };

  var found = findRow(calSheet, 'CAL_ID', params.cal_id);
  if (!found) return { ok: false, error: 'Рейс не знайдено' };

  // Clear CAL_ID in passengers
  clearCalIdInPassengers(params.cal_id);

  calSheet.deleteRow(found.rowNum);

  return { ok: true };
}

function clearCalIdInPassengers(calId) {
  [SHEETS.PAX_UE, SHEETS.PAX_EU].forEach(function(shName) {
    var sh = getSheet(shName);
    if (!sh) return;
    var info = getAllData(sh);
    var calIdx = info.headers.indexOf('CAL_ID');
    if (calIdx === -1) return;

    for (var i = 0; i < info.data.length; i++) {
      if (String(info.data[i][calIdx]) === String(calId)) {
        sh.getRange(DATA_START + i, calIdx + 1).setValue('');
      }
    }
  });
}


// ════════════════════════════════════════════
// doGet / doPost — Unified API
// ════════════════════════════════════════════

function doGet(e) {
  var action = e.parameter.action || '';
  var result = { ok: false, error: 'Unknown action' };

  try {
    if (action === 'getAll') {
      result = apiGetAll({ sheet: e.parameter.sheet, filter: {} });
    } else if (action === 'getTrips') {
      result = apiGetTrips({ filter: {} });
    }
  } catch (err) {
    result = { ok: false, error: err.message };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var body = JSON.parse(e.postData.contents);
  var action = body.action || '';
  var result = { ok: false, error: 'Unknown action' };

  try {
    switch (action) {
      case 'getAll':
        result = apiGetAll(body);
        break;
      case 'addPassenger':
        result = apiAddPassenger(body);
        break;
      case 'updateField':
        result = apiUpdateField(body);
        break;
      case 'deletePassenger':
        result = apiDeletePassenger(body);
        break;
      case 'checkDuplicates':
        result = apiCheckDuplicates(body);
        break;
      case 'getTrips':
        result = apiGetTrips(body);
        break;
      case 'createTrip':
        result = apiCreateTrip(body);
        break;
      case 'updateTrip':
        result = apiUpdateTrip(body);
        break;
      case 'archiveTrip':
        result = apiArchiveTrip(body);
        break;
      case 'deleteTrip':
        result = apiDeleteTrip(body);
        break;
      default:
        result = { ok: false, error: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { ok: false, error: err.message };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}
