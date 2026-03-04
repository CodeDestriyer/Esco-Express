// ============================================================
// CRM BACKEND — Google Apps Script
// Структура таблиць + Логіка рейтингів + Архів + Клієнти
// ============================================================

// ── КОНФІГУРАЦІЯ ─────────────────────────────────────────────
const CONFIG = {
  SPREADSHEETS: {
    PASSENGERS: "ID_ФАЙЛУ",  // Passengers_crm
    PARCELS:    "ID_ФАЙЛУ",  // Posylki_crm_v2
    ROUTES:     "ID_ФАЙЛУ",  // Marhrut_crm_v3
    CLIENTS:    "ID_ФАЙЛУ",  // Kliyentu_crm_final
    ARCHIVE:    "ID_ФАЙЛУ",  // archive_crm
  },

  // Назви аркушів
  SHEETS: {
    // Passengers_crm
    PAX_UA_EU: "Україна-ЄВ",
    PAX_EU_UA: "Європа-УК",

    // Posylki_crm_v2
    PKG_UA_EU: "Реєстрація ТТН УК-єв",
    PKG_EU_UA: "Виклик Курєра ЄВ-ук",

    // Marhrut_crm_v3
    RTE_ZURICH:    "Цюріх",
    RTE_GENEVA:    "Женева",
    RTE_OPTIM:     "Оптимістичний",

    // archive_crm
    ARC_PASSENGERS: "Пасажири",
    ARC_PARCELS:    "Посилки",
    ARC_ROUTES:     "Маршрути",
    ARC_LOGS:       "Логи",

    // Kliyentu_crm_final
    CLIENTS:         "Клієнти",
    CLIENT_RATINGS:  "Рейтинг клієнтів",
    CLIENT_REVIEWS:  "Відгуки клієнтів",
  },

  // Префікси ID
  PREFIXES: {
    PAX: "PAX",
    PKG: "PKG",
    RTE: "RTE",
    CLI: "CLI",
    RAT: "RAT",
    REV: "REV",
    ARC: "ARC",
    LOG: "LOG",
  },
};

// Шкала балів для зовнішнього рейтингу
const SCORE = { "Супер": 5, "Добре": 3, "Погано": 1 };


// ── ГЕНЕРАЦІЯ УНІКАЛЬНИХ ID ──────────────────────────────────
function generateID(prefix) {
  const date = Utilities.formatDate(new Date(), "UTC", "yyyyMMdd");
  const random = Math.random().toString(36).substr(2, 4).toUpperCase();
  return prefix + "-" + date + "-" + random;
}


// ── ВИЗНАЧЕННЯ КОЛОНОК ТА ЗАГОЛОВКІВ ─────────────────────────

// Таблиця 1 — Passengers_crm (28 колонок, однакові для обох аркушів)
const PASSENGER_COLUMNS = [
  "PAX_ID",
  "Ід_смарт/CRM",
  "Напрям",
  "SOURCE_SHEET",
  "Дата створення",
  "Піб",
  "Телефон пасажира",
  "Телефон реєстратора",
  "Кількість місць",
  "Адреса прибуття",
  "Дата виїзду",
  "Таймінг",
  "Номер авто",
  "Місце в авто",
  "RTE_ID",
  "Ціна квитка",
  "Завдаток",
  "Статус оплати",
  "Форма оплати",
  "Валюта",
  "Вага багажу",
  "Ціна багажу",
  "Статус ліда",
  "Статус CRM",
  "DATE_ARCHIVE",
  "ARCHIVED_BY",
  "ARCHIVE_REASON",
  "ARCHIVE_ID",
];

// Таблиця 2 — Posylki_crm_v2: "Реєстрація ТТН УК-єв" (39 колонок)
const PARCEL_UA_EU_COLUMNS = [
  "PKG_ID",
  "Ід_смарт/CRM",
  "Напрям",
  "SOURCE_SHEET",
  "Дата створення",
  "Піб відправника",
  "Телефон реєстратора",
  "Піб отримувача",
  "Телефон отримувача",
  "Адреса в Європі",
  "Внутрішній номер№",
  "Номер ТТН",
  "Опис",
  "Деталі",
  "Кількість позицій",
  "Кг",
  "Оціночна вартість",
  "Фото посилки",
  "Дата відправки",
  "Таймінг",
  "Номер авто",
  "RTE_ID",
  "Сума",
  "Валюта оплати",
  "Форма оплати",
  "Статус оплати",
  "Статус посилки",
  "Статус",
  "Статус CRM",
  "Дата отримання",
  "Контроль перевірки",
  "Дата перевірки і час",
  "Рейтинг",
  "Коментар до рейтингу",
  "Примітка смс",
  "DATE_ARCHIVE",
  "ARCHIVED_BY",
  "ARCHIVE_REASON",
  "ARCHIVE_ID",
];

// Таблиця 2 — Posylki_crm_v2: "Виклик Курєра ЄВ-ук" (37 колонок)
// Без: "Адреса в Європі", "Номер ТТН", "Оціночна вартість"
// Додано: "Місто Нова Пошта" (після "Телефон отримувача")
const PARCEL_EU_UA_COLUMNS = [
  "PKG_ID",
  "Ід_смарт/CRM",
  "Напрям",
  "SOURCE_SHEET",
  "Дата створення",
  "Піб відправника",
  "Телефон реєстратора",
  "Піб отримувача",
  "Телефон отримувача",
  "Місто Нова Пошта",
  "Внутрішній номер№",
  "Опис",
  "Деталі",
  "Кількість позицій",
  "Кг",
  "Фото посилки",
  "Дата відправки",
  "Таймінг",
  "Номер авто",
  "RTE_ID",
  "Сума",
  "Валюта оплати",
  "Форма оплати",
  "Статус оплати",
  "Статус посилки",
  "Статус",
  "Статус CRM",
  "Дата отримання",
  "Контроль перевірки",
  "Дата перевірки і час",
  "Рейтинг",
  "Коментар до рейтингу",
  "Примітка смс",
  "DATE_ARCHIVE",
  "ARCHIVED_BY",
  "ARCHIVE_REASON",
  "ARCHIVE_ID",
];

// Таблиця 3 — Marhrut_crm_v3 (40 колонок, однакові для всіх аркушів)
const ROUTE_COLUMNS = [
  "RTE_ID",
  "Тип запису",
  "Напрям",
  "SOURCE_SHEET",
  "PAX_ID / PKG_ID",
  "Дата створення",
  "Дата рейсу",
  "Таймінг",
  "Номер авто",
  "Водій",
  "Телефон водія",
  "Місце в авто",
  "Піб пасажира",
  "Телефон пасажира",
  "Адреса прибуття",
  "Кількість місць",
  "Вага багажу",
  "Піб відправника",
  "Піб отримувача",
  "Телефон отримувача",
  "Адреса отримувача",
  "Внутрішній номер№",
  "Номер ТТН",
  "Опис посилки",
  "Кг посилки",
  "Сума",
  "Валюта",
  "Форма оплати",
  "Статус оплати",
  "Статус",
  "Статус CRM",
  "Примітка",
  "Рейтинг водія",
  "Коментар водія",
  "Рейтинг менеджера",
  "Коментар менеджера",
  "DATE_ARCHIVE",
  "ARCHIVED_BY",
  "ARCHIVE_REASON",
  "ARCHIVE_ID",
];

// Таблиця 4 — archive_crm: "Логи" (10 колонок)
const LOG_COLUMNS = [
  "LOG_ID",
  "Дата і час",
  "Хто",
  "Роль",
  "Дія",
  "Таблиця",
  "Аркуш",
  "ID запису",
  "Поле",
  "Було → Стало",
];

// Таблиця 4 — archive_crm: аркуші архіву (загальна структура)
const ARCHIVE_META_COLUMNS = [
  "ARCHIVE_ID",
  "Тип дії",
  "DATE_ARCHIVE",
  "ARCHIVED_BY",
  "ARCHIVE_REASON",
  "SOURCE_TABLE",
  "SOURCE_SHEET",
  "ORIGINAL_ID",
  "Напрям",
  "Статус на момент",
];

const ARCHIVE_RESTORE_COLUMNS = [
  "Відновлено",
  "Дата відновлення",
  "Відновив",
  "Причина відновлення",
];

// Таблиця 5 — Kliyentu_crm_final: "Клієнти" (34 колонки)
const CLIENT_COLUMNS = [
  "CLIENT_ID",
  "Ід_смарт/CRM",
  "Дата реєстрації",
  "Остання активність",
  "Телефон",
  "Піб",
  "Додатковий телефон",
  "Напрям",
  "Тип клієнта",
  "К-сть рейсів",
  "К-сть посилок",
  "Загальна сума UAH",
  "Борг",
  "Остання оплата",
  "Рейт. водія (сер.)",
  "Оцінок від водія",
  "Сума балів водія",
  "Рейт. менеджера (сер.)",
  "Оцінок від менеджера",
  "Сума балів менеджера",
  "Внутрішній рейтинг",
  "Останні 3 коментарі",
  "Рейт. через бот (сер.)",
  "Оцінок через бот",
  "Сума балів бот",
  "Супер 😊",
  "Добре 😐",
  "Погано 😞",
  "Останній відгук",
  "Дата останнього відгуку",
  "VIP",
  "Стоп-лист",
  "Причина стоп-листа",
  "Примітка",
];

// Таблиця 5 — Kliyentu_crm_final: "Рейтинг клієнтів" (14 колонок)
const CLIENT_RATING_COLUMNS = [
  "RATE_ID",
  "Дата оцінки",
  "CLIENT_ID",
  "Телефон клієнта",
  "Піб клієнта",
  "RTE_ID",
  "Дата рейсу",
  "Тип запису",
  "Оцінка водія",
  "Коментар водія",
  "Водій",
  "Оцінка менеджера",
  "Коментар менеджера",
  "Менеджер",
];

// Таблиця 5 — Kliyentu_crm_final: "Відгуки клієнтів" (12 колонок)
const CLIENT_REVIEW_COLUMNS = [
  "ID_смарт",
  "Телефон клієнта",
  "Дата відгуку",
  "Оцінка менеджера",
  "Оцінка водія",
  "Коментар",
  "Статус",
  "Опрацював",
  "Результат",
  "Дублікат",
  "Оригінальний текст A",
  "Оригінальний текст B",
];


// ── ДОПОМІЖНІ ФУНКЦІЇ ────────────────────────────────────────

/**
 * Відкриває аркуш Google Sheets за ID та назвою
 */
function getSheet(spreadsheetId, sheetName) {
  return SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);
}

/**
 * Знаходить рядок за значенням у конкретній колонці
 * Повертає об'єкт { row: номер_рядка, data: масив_значень } або null
 */
function findRowByColumn(sheet, columnName, value) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colIndex = headers.indexOf(columnName);
  if (colIndex === -1) return null;

  const data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][colIndex] == value) {
      return { row: i + 1, data: data[i], headers: headers };
    }
  }
  return null;
}

/**
 * Конвертує рядок масиву в об'єкт з іменованими ключами
 */
function rowToObject(headers, rowData) {
  var obj = {};
  for (var i = 0; i < headers.length; i++) {
    obj[headers[i]] = rowData[i] !== undefined ? rowData[i] : "";
  }
  return obj;
}

/**
 * Конвертує об'єкт назад у масив значень відповідно до заголовків
 */
function objectToRow(headers, obj) {
  return headers.map(function(h) {
    return obj[h] !== undefined ? obj[h] : "";
  });
}

/**
 * Записує оновлений рядок назад в таблицю
 */
function updateRow(sheet, rowNumber, headers, obj) {
  var values = objectToRow(headers, obj);
  sheet.getRange(rowNumber, 1, 1, values.length).setValues([values]);
}

/**
 * Додає новий рядок в кінець таблиці
 */
function appendRow(sheet, headers, obj) {
  var values = objectToRow(headers, obj);
  sheet.appendRow(values);
}


// ── ЛОГУВАННЯ ────────────────────────────────────────────────

/**
 * Записує запис у аркуш "Логи"
 */
function writeLog(who, role, action, tableName, sheetName, recordId, field, change) {
  var logSheet = getSheet(CONFIG.SPREADSHEETS.ARCHIVE, CONFIG.SHEETS.ARC_LOGS);

  var logEntry = {};
  logEntry["LOG_ID"] = generateID(CONFIG.PREFIXES.LOG);
  logEntry["Дата і час"] = new Date();
  logEntry["Хто"] = who || "";
  logEntry["Роль"] = role || "Система";
  logEntry["Дія"] = action || "";
  logEntry["Таблиця"] = tableName || "";
  logEntry["Аркуш"] = sheetName || "";
  logEntry["ID запису"] = recordId || "";
  logEntry["Поле"] = field || "";
  logEntry["Було → Стало"] = change || "";

  appendRow(logSheet, LOG_COLUMNS, logEntry);
}


// ══════════════════════════════════════════════════════════════
// РЕЙТИНГИ — ВНУТРІШНІЙ (водій/менеджер → клієнту)
// ══════════════════════════════════════════════════════════════

/**
 * Спрацьовує при зміні статусу маршруту на "Виконано"
 * Ставить авто-рейтинг 5 якщо оцінка ще не виставлена вручну
 */
function onRouteComplete(rteId, routeSheetName) {
  var routeSheet = getSheet(CONFIG.SPREADSHEETS.ROUTES, routeSheetName);
  var found = findRowByColumn(routeSheet, "RTE_ID", rteId);
  if (!found) return;

  var headers = found.headers;
  var row = rowToObject(headers, found.data);

  var changed = false;

  // Якщо рейтинг ще не виставлено вручну — ставимо 5
  if (!row["Рейтинг водія"]) {
    row["Рейтинг водія"] = 5;
    row["Коментар водія"] = "";
    changed = true;
  }
  if (!row["Рейтинг менеджера"]) {
    row["Рейтинг менеджера"] = 5;
    row["Коментар менеджера"] = "";
    changed = true;
  }

  if (changed) {
    updateRow(routeSheet, found.row, headers, row);
  }

  // Записуємо в аркуш "Рейтинг клієнтів"
  addClientRating(row);

  writeLog("Система", "Система", "Рейтинг", "Marhrut_crm_v3", routeSheetName,
    rteId, "Статус", "→ Виконано (авто-рейтинг)");
}

/**
 * Ручний рейтинг — коли водій/менеджер натискає кнопку "⭐ Рейтинг"
 * @param {string} rteId — ID маршруту
 * @param {string} routeSheetName — назва аркуша маршруту
 * @param {string} role — "driver" або "manager"
 * @param {number} rating — число 1-5
 * @param {string} comment — коментар (обов'язковий якщо < 5)
 * @param {string} raterName — ПІБ того хто оцінює
 */
function setManualRating(rteId, routeSheetName, role, rating, comment, raterName) {
  // Валідація
  rating = Number(rating);
  if (rating < 1 || rating > 5) {
    throw new Error("Оцінка має бути від 1 до 5");
  }
  if (rating < 5 && !comment) {
    throw new Error("Коментар обов'язковий якщо оцінка менше 5");
  }

  var routeSheet = getSheet(CONFIG.SPREADSHEETS.ROUTES, routeSheetName);
  var found = findRowByColumn(routeSheet, "RTE_ID", rteId);
  if (!found) {
    throw new Error("Маршрут " + rteId + " не знайдено");
  }

  var headers = found.headers;
  var row = rowToObject(headers, found.data);

  var oldRating;
  if (role === "driver") {
    oldRating = row["Рейтинг водія"];
    row["Рейтинг водія"] = rating;
    row["Коментар водія"] = comment || "";
  } else {
    oldRating = row["Рейтинг менеджера"];
    row["Рейтинг менеджера"] = rating;
    row["Коментар менеджера"] = comment || "";
  }

  updateRow(routeSheet, found.row, headers, row);

  // Оновити або додати в "Рейтинг клієнтів"
  addClientRating(row);

  var fieldName = role === "driver" ? "Рейтинг водія" : "Рейтинг менеджера";
  writeLog(raterName, role === "driver" ? "Водій" : "Менеджер", "Рейтинг",
    "Marhrut_crm_v3", routeSheetName, rteId, fieldName,
    (oldRating || "—") + " → " + rating);
}

/**
 * Записує оцінку в аркуш "Рейтинг клієнтів" та оновлює середній в "Клієнти"
 */
function addClientRating(routeRow) {
  // Визначаємо телефон клієнта з маршруту
  var phone = routeRow["Телефон пасажира"] || routeRow["Телефон отримувача"] || "";
  var name = routeRow["Піб пасажира"] || routeRow["Піб відправника"] || "";
  var smartId = "";
  var recordType = routeRow["Тип запису"] || "";

  if (!phone) return;

  // Знаходимо або створюємо клієнта
  var clientResult = findOrCreateClient(smartId, phone, name);

  // Записуємо рядок в "Рейтинг клієнтів"
  var ratingSheet = getSheet(CONFIG.SPREADSHEETS.CLIENTS, CONFIG.SHEETS.CLIENT_RATINGS);
  var ratingEntry = {};
  ratingEntry["RATE_ID"] = generateID(CONFIG.PREFIXES.RAT);
  ratingEntry["Дата оцінки"] = new Date();
  ratingEntry["CLIENT_ID"] = clientResult.clientId;
  ratingEntry["Телефон клієнта"] = phone;
  ratingEntry["Піб клієнта"] = name;
  ratingEntry["RTE_ID"] = routeRow["RTE_ID"] || "";
  ratingEntry["Дата рейсу"] = routeRow["Дата рейсу"] || "";
  ratingEntry["Тип запису"] = recordType;
  ratingEntry["Оцінка водія"] = routeRow["Рейтинг водія"] || "";
  ratingEntry["Коментар водія"] = routeRow["Коментар водія"] || "";
  ratingEntry["Водій"] = routeRow["Водій"] || "";
  ratingEntry["Оцінка менеджера"] = routeRow["Рейтинг менеджера"] || "";
  ratingEntry["Коментар менеджера"] = routeRow["Коментар менеджера"] || "";
  ratingEntry["Менеджер"] = routeRow["Коментар менеджера"] ? (routeRow["Водій"] || "Менеджер") : "Авто";

  appendRow(ratingSheet, CLIENT_RATING_COLUMNS, ratingEntry);

  // Оновлюємо накопичений середній у "Клієнти"
  var driverRating = routeRow["Рейтинг водія"] ? Number(routeRow["Рейтинг водія"]) : null;
  var managerRating = routeRow["Рейтинг менеджера"] ? Number(routeRow["Рейтинг менеджера"]) : null;
  var comment = routeRow["Коментар водія"] || routeRow["Коментар менеджера"] || "";

  updateClientInternalRating(clientResult.sheet, clientResult.row, clientResult.headers,
    driverRating, managerRating, comment);
}

/**
 * Оновлює накопичений внутрішній рейтинг клієнта (без втрати історії)
 */
function updateClientInternalRating(clientSheet, rowNumber, headers, newDriverRating, newManagerRating, comment) {
  var found = { row: rowNumber, data: clientSheet.getRange(rowNumber, 1, 1, headers.length).getValues()[0], headers: headers };
  var clientRow = rowToObject(headers, found.data);

  // ВОДІЙ
  if (newDriverRating) {
    var oldSum = Number(clientRow["Сума балів водія"]) || 0;
    var oldCount = Number(clientRow["Оцінок від водія"]) || 0;
    var newSum = oldSum + newDriverRating;
    var newCount = oldCount + 1;
    clientRow["Сума балів водія"] = newSum;
    clientRow["Оцінок від водія"] = newCount;
    clientRow["Рейт. водія (сер.)"] = Math.round((newSum / newCount) * 10) / 10;
  }

  // МЕНЕДЖЕР
  if (newManagerRating) {
    var oldSumM = Number(clientRow["Сума балів менеджера"]) || 0;
    var oldCountM = Number(clientRow["Оцінок від менеджера"]) || 0;
    var newSumM = oldSumM + newManagerRating;
    var newCountM = oldCountM + 1;
    clientRow["Сума балів менеджера"] = newSumM;
    clientRow["Оцінок від менеджера"] = newCountM;
    clientRow["Рейт. менеджера (сер.)"] = Math.round((newSumM / newCountM) * 10) / 10;
  }

  // ЗАГАЛЬНИЙ ВНУТРІШНІЙ РЕЙТИНГ
  var avgDriver = Number(clientRow["Рейт. водія (сер.)"]) || 0;
  var avgManager = Number(clientRow["Рейт. менеджера (сер.)"]) || 0;
  if (avgDriver && avgManager) {
    clientRow["Внутрішній рейтинг"] = Math.round(((avgDriver + avgManager) / 2) * 10) / 10;
  } else {
    clientRow["Внутрішній рейтинг"] = avgDriver || avgManager || 0;
  }

  // ОСТАННІ 3 КОМЕНТАРІ (тільки якщо є коментар)
  if (comment) {
    var existing = clientRow["Останні 3 коментарі"] ? String(clientRow["Останні 3 коментарі"]) : "";
    var comments = existing ? existing.split(" | ") : [];
    comments.unshift(comment);
    clientRow["Останні 3 коментарі"] = comments.slice(0, 3).join(" | ");
  }

  // Оновлюємо останню активність
  clientRow["Остання активність"] = new Date();

  updateRow(clientSheet, rowNumber, headers, clientRow);
}


// ══════════════════════════════════════════════════════════════
// РЕЙТИНГИ — ЗОВНІШНІЙ (клієнт → нам через бот)
// ══════════════════════════════════════════════════════════════

/**
 * Нормалізує текст оцінки зі Смартсендера
 */
function normalizeRating(value) {
  if (!value) return null;
  var val = value.toString().toLowerCase().trim();
  if (val.indexOf("супер") !== -1) return "Супер";
  if (val.indexOf("добре") !== -1) return "Добре";
  if (val.indexOf("погано") !== -1) return "Погано";
  return null; // все інше — не оцінка, іде в "Коментар"
}

/**
 * Перевірка на дублікат відгуку (24 год)
 */
function isDuplicate(smartId, reviewSheet) {
  var oneDayAgo = new Date(Date.now() - 24 * 60 * 60 * 1000);
  var data = reviewSheet.getDataRange().getValues();
  var headers = data[0];
  var idCol = headers.indexOf("ID_смарт");
  var dateCol = headers.indexOf("Дата відгуку");
  var dupCol = headers.indexOf("Дублікат");

  for (var i = 1; i < data.length; i++) {
    if (data[i][idCol] == smartId &&
        new Date(data[i][dateCol]) > oneDayAgo &&
        data[i][dupCol] !== "Так") {
      return true;
    }
  }
  return false;
}

/**
 * Визначає авто-статус відгуку
 */
function getReviewStatus(managerRating, driverRating) {
  if (!managerRating && !driverRating) return "Сміття";
  if (managerRating === "Погано" || driverRating === "Погано") return "Новий❗";
  return "Опрацьовано"; // Супер або Добре — авто-закриваємо
}

/**
 * Обробляє новий відгук зі Смартсендера
 * @param {string} smartId — ID зі Смартсендера
 * @param {string} phone — телефон клієнта
 * @param {string} rawTextA — сирі дані колонки А (оцінка менеджера)
 * @param {string} rawTextB — сирі дані колонки В (оцінка водія)
 * @param {string} commentText — коментар клієнта
 */
function processExternalReview(smartId, phone, rawTextA, rawTextB, commentText) {
  var reviewSheet = getSheet(CONFIG.SPREADSHEETS.CLIENTS, CONFIG.SHEETS.CLIENT_REVIEWS);

  // Перевірка на дублікат
  var duplicate = isDuplicate(smartId, reviewSheet);

  // Нормалізація оцінок
  var managerRating = normalizeRating(rawTextA);
  var driverRating = normalizeRating(rawTextB);

  // Якщо текст не розпізнаний як оцінка — іде в коментар
  var comment = commentText || "";
  if (!managerRating && rawTextA) {
    comment = comment ? comment + " | " + rawTextA : rawTextA;
  }
  if (!driverRating && rawTextB) {
    comment = comment ? comment + " | " + rawTextB : rawTextB;
  }

  // Визначаємо статус
  var status = duplicate ? "Сміття" : getReviewStatus(managerRating, driverRating);

  // Записуємо відгук
  var reviewEntry = {};
  reviewEntry["ID_смарт"] = smartId;
  reviewEntry["Телефон клієнта"] = phone;
  reviewEntry["Дата відгуку"] = new Date();
  reviewEntry["Оцінка менеджера"] = managerRating || "";
  reviewEntry["Оцінка водія"] = driverRating || "";
  reviewEntry["Коментар"] = comment;
  reviewEntry["Статус"] = status;
  reviewEntry["Опрацював"] = "";
  reviewEntry["Результат"] = "";
  reviewEntry["Дублікат"] = duplicate ? "Так" : "Ні";
  reviewEntry["Оригінальний текст A"] = rawTextA || "";
  reviewEntry["Оригінальний текст B"] = rawTextB || "";

  appendRow(reviewSheet, CLIENT_REVIEW_COLUMNS, reviewEntry);

  // Якщо не дублікат і є хоча б одна оцінка — оновлюємо клієнта
  if (!duplicate && (managerRating || driverRating)) {
    var clientResult = findOrCreateClient(smartId, phone, "");
    updateClientExternalRating(clientResult.sheet, clientResult.row, clientResult.headers,
      managerRating, driverRating, comment);
  }

  writeLog("Смартсендер", "Система", "Рейтинг", "Kliyentu_crm_final",
    CONFIG.SHEETS.CLIENT_REVIEWS, smartId, "Зовнішній відгук",
    "Менеджер: " + (managerRating || "—") + ", Водій: " + (driverRating || "—"));
}

/**
 * Оновлює накопичений зовнішній рейтинг клієнта
 */
function updateClientExternalRating(clientSheet, rowNumber, headers, managerRating, driverRating, comment) {
  var data = clientSheet.getRange(rowNumber, 1, 1, headers.length).getValues()[0];
  var clientRow = rowToObject(headers, data);

  // Рахуємо середній бал по обох оцінках
  var scores = [];
  if (SCORE[managerRating]) scores.push(SCORE[managerRating]);
  if (SCORE[driverRating]) scores.push(SCORE[driverRating]);

  for (var i = 0; i < scores.length; i++) {
    var oldSum = Number(clientRow["Сума балів бот"]) || 0;
    var oldCount = Number(clientRow["Оцінок через бот"]) || 0;
    clientRow["Сума балів бот"] = oldSum + scores[i];
    clientRow["Оцінок через бот"] = oldCount + 1;
    clientRow["Рейт. через бот (сер.)"] = Math.round(
      (Number(clientRow["Сума балів бот"]) / Number(clientRow["Оцінок через бот"])) * 10
    ) / 10;
  }

  // Лічильники Супер/Добре/Погано
  if (managerRating === "Супер" || driverRating === "Супер") {
    clientRow["Супер 😊"] = (Number(clientRow["Супер 😊"]) || 0) + 1;
  }
  if (managerRating === "Добре" || driverRating === "Добре") {
    clientRow["Добре 😐"] = (Number(clientRow["Добре 😐"]) || 0) + 1;
  }
  if (managerRating === "Погано" || driverRating === "Погано") {
    clientRow["Погано 😞"] = (Number(clientRow["Погано 😞"]) || 0) + 1;
  }

  // Останній відгук
  if (comment) {
    clientRow["Останній відгук"] = comment;
    clientRow["Дата останнього відгуку"] = new Date();
  }

  // Оновлюємо останню активність
  clientRow["Остання активність"] = new Date();

  updateRow(clientSheet, rowNumber, headers, clientRow);
}


// ══════════════════════════════════════════════════════════════
// ПОШУК / СТВОРЕННЯ КЛІЄНТА
// ══════════════════════════════════════════════════════════════

/**
 * Знаходить або створює клієнта
 * Повертає { clientId, sheet, row, headers }
 */
function findOrCreateClient(smartId, phone, name) {
  var clientSheet = getSheet(CONFIG.SPREADSHEETS.CLIENTS, CONFIG.SHEETS.CLIENTS);
  var headers = clientSheet.getRange(1, 1, 1, clientSheet.getLastColumn()).getValues()[0];

  // 1. Шукаємо по Ід_смарт/CRM
  var found = null;
  if (smartId) {
    found = findRowByColumn(clientSheet, "Ід_смарт/CRM", smartId);
  }

  // 2. Якщо не знайшли — шукаємо по телефону
  if (!found && phone) {
    found = findRowByColumn(clientSheet, "Телефон", phone);
  }

  // 3. Якщо не знайшли — створюємо нового клієнта
  if (!found) {
    var clientId = generateID(CONFIG.PREFIXES.CLI);
    var newClient = {};
    // Ініціалізуємо всі поля порожніми
    for (var i = 0; i < CLIENT_COLUMNS.length; i++) {
      newClient[CLIENT_COLUMNS[i]] = "";
    }
    newClient["CLIENT_ID"] = clientId;
    newClient["Ід_смарт/CRM"] = smartId || "";
    newClient["Дата реєстрації"] = new Date();
    newClient["Остання активність"] = new Date();
    newClient["Телефон"] = phone || "";
    newClient["Піб"] = name || "";
    newClient["К-сть рейсів"] = 0;
    newClient["К-сть посилок"] = 0;
    newClient["Загальна сума UAH"] = 0;
    newClient["Борг"] = 0;
    newClient["Рейт. водія (сер.)"] = 0;
    newClient["Оцінок від водія"] = 0;
    newClient["Сума балів водія"] = 0;
    newClient["Рейт. менеджера (сер.)"] = 0;
    newClient["Оцінок від менеджера"] = 0;
    newClient["Сума балів менеджера"] = 0;
    newClient["Внутрішній рейтинг"] = 0;
    newClient["Рейт. через бот (сер.)"] = 0;
    newClient["Оцінок через бот"] = 0;
    newClient["Сума балів бот"] = 0;
    newClient["Супер 😊"] = 0;
    newClient["Добре 😐"] = 0;
    newClient["Погано 😞"] = 0;
    newClient["VIP"] = "Ні";
    newClient["Стоп-лист"] = "Ні";

    appendRow(clientSheet, CLIENT_COLUMNS, newClient);

    // Знаходимо щойно створений рядок
    var lastRow = clientSheet.getLastRow();
    return {
      clientId: clientId,
      sheet: clientSheet,
      row: lastRow,
      headers: CLIENT_COLUMNS,
    };
  }

  return {
    clientId: rowToObject(found.headers, found.data)["CLIENT_ID"],
    sheet: clientSheet,
    row: found.row,
    headers: found.headers,
  };
}


// ══════════════════════════════════════════════════════════════
// АРХІВУВАННЯ
// ══════════════════════════════════════════════════════════════

/**
 * Архівує запис з будь-якої таблиці
 * @param {string} sourceType — "passenger", "parcel", "route"
 * @param {string} sourceSheetName — назва аркуша-джерела
 * @param {string} recordId — ID запису (PAX/PKG/RTE)
 * @param {string} archivedBy — хто архівує
 * @param {string} reason — причина
 * @param {string} actionType — "Архівовано" або "Видалено"
 */
function archiveRecord(sourceType, sourceSheetName, recordId, archivedBy, reason, actionType) {
  actionType = actionType || "Архівовано";

  // Визначаємо ID таблиці та ID-колонку
  var spreadsheetId, idColumn, archiveSheetName, originalColumns;

  if (sourceType === "passenger") {
    spreadsheetId = CONFIG.SPREADSHEETS.PASSENGERS;
    idColumn = "PAX_ID";
    archiveSheetName = CONFIG.SHEETS.ARC_PASSENGERS;
    originalColumns = PASSENGER_COLUMNS;
  } else if (sourceType === "parcel") {
    spreadsheetId = CONFIG.SPREADSHEETS.PARCELS;
    idColumn = "PKG_ID";
    archiveSheetName = CONFIG.SHEETS.ARC_PARCELS;
    if (sourceSheetName === CONFIG.SHEETS.PKG_UA_EU) {
      originalColumns = PARCEL_UA_EU_COLUMNS;
    } else {
      originalColumns = PARCEL_EU_UA_COLUMNS;
    }
  } else if (sourceType === "route") {
    spreadsheetId = CONFIG.SPREADSHEETS.ROUTES;
    idColumn = "RTE_ID";
    archiveSheetName = CONFIG.SHEETS.ARC_ROUTES;
    originalColumns = ROUTE_COLUMNS;
  } else {
    throw new Error("Невідомий тип: " + sourceType);
  }

  // Знаходимо оригінальний запис
  var sourceSheet = getSheet(spreadsheetId, sourceSheetName);
  var found = findRowByColumn(sourceSheet, idColumn, recordId);
  if (!found) {
    throw new Error("Запис " + recordId + " не знайдено в " + sourceSheetName);
  }

  var originalRow = rowToObject(found.headers, found.data);

  // Формуємо архівний запис
  var archiveSheet = getSheet(CONFIG.SPREADSHEETS.ARCHIVE, archiveSheetName);
  var archiveId = generateID(CONFIG.PREFIXES.ARC);
  var now = new Date();

  // Мета-дані архіву
  var archiveEntry = [];
  archiveEntry.push(archiveId);                              // ARCHIVE_ID
  archiveEntry.push(actionType);                             // Тип дії
  archiveEntry.push(now);                                    // DATE_ARCHIVE
  archiveEntry.push(archivedBy);                             // ARCHIVED_BY
  archiveEntry.push(reason);                                 // ARCHIVE_REASON
  archiveEntry.push(sourceType === "passenger" ? "Passengers_crm"
    : sourceType === "parcel" ? "Posylki_crm_v2"
    : "Marhrut_crm_v3");                                     // SOURCE_TABLE
  archiveEntry.push(sourceSheetName);                        // SOURCE_SHEET
  archiveEntry.push(recordId);                               // ORIGINAL_ID
  archiveEntry.push(originalRow["Напрям"] || "");            // Напрям
  archiveEntry.push(originalRow["Статус CRM"] || originalRow["Статус"] || ""); // Статус на момент

  // Всі оригінальні поля
  for (var i = 0; i < originalColumns.length; i++) {
    archiveEntry.push(originalRow[originalColumns[i]] || "");
  }

  // Поля відновлення (порожні при архівуванні)
  archiveEntry.push("Ні");   // Відновлено
  archiveEntry.push("");     // Дата відновлення
  archiveEntry.push("");     // Відновив
  archiveEntry.push("");     // Причина відновлення

  archiveSheet.appendRow(archiveEntry);

  // Оновлюємо оригінальний запис — позначаємо як заархівований
  originalRow["DATE_ARCHIVE"] = now;
  originalRow["ARCHIVED_BY"] = archivedBy;
  originalRow["ARCHIVE_REASON"] = reason;
  originalRow["ARCHIVE_ID"] = archiveId;
  originalRow["Статус CRM"] = "Скасовано";

  updateRow(sourceSheet, found.row, found.headers, originalRow);

  // Лог
  writeLog(archivedBy, "Менеджер", actionType === "Видалено" ? "Видалив" : "Архівував",
    sourceType === "passenger" ? "Passengers_crm"
    : sourceType === "parcel" ? "Posylki_crm_v2"
    : "Marhrut_crm_v3",
    sourceSheetName, recordId, "Статус CRM",
    (originalRow["Статус CRM"] || "—") + " → Скасовано (" + actionType + ")");

  return archiveId;
}

/**
 * Відновлює запис з архіву
 */
function restoreFromArchive(archiveId, archiveSheetName, restoredBy, restoreReason) {
  var archiveSheet = getSheet(CONFIG.SPREADSHEETS.ARCHIVE, archiveSheetName);
  var found = findRowByColumn(archiveSheet, "ARCHIVE_ID", archiveId);
  if (!found) {
    throw new Error("Архівний запис " + archiveId + " не знайдено");
  }

  var headers = found.headers;
  var archiveRow = rowToObject(headers, found.data);

  // Позначаємо як відновлений
  archiveRow["Відновлено"] = "Так";
  archiveRow["Дата відновлення"] = new Date();
  archiveRow["Відновив"] = restoredBy;
  archiveRow["Причина відновлення"] = restoreReason || "";

  updateRow(archiveSheet, found.row, headers, archiveRow);

  // Лог
  writeLog(restoredBy, "Менеджер", "Відновив", "archive_crm",
    archiveSheetName, archiveId, "Відновлено", "Ні → Так");

  return archiveId;
}


// ══════════════════════════════════════════════════════════════
// СТВОРЕННЯ ЗАПИСІВ
// ══════════════════════════════════════════════════════════════

/**
 * Створює нового пасажира
 */
function createPassenger(data, sheetName) {
  sheetName = sheetName || CONFIG.SHEETS.PAX_UA_EU;
  var sheet = getSheet(CONFIG.SPREADSHEETS.PASSENGERS, sheetName);

  var paxId = generateID(CONFIG.PREFIXES.PAX);
  var entry = {};
  for (var i = 0; i < PASSENGER_COLUMNS.length; i++) {
    entry[PASSENGER_COLUMNS[i]] = "";
  }

  entry["PAX_ID"] = paxId;
  entry["Ід_смарт/CRM"] = data.smartId || "";
  entry["Напрям"] = sheetName === CONFIG.SHEETS.PAX_UA_EU ? "УК → ЄВ" : "ЄВ → УК";
  entry["SOURCE_SHEET"] = sheetName;
  entry["Дата створення"] = new Date();
  entry["Піб"] = data.name || "";
  entry["Телефон пасажира"] = data.phone || "";
  entry["Телефон реєстратора"] = data.phoneReg || "";
  entry["Кількість місць"] = data.seats || 1;
  entry["Адреса прибуття"] = data.address || "";
  entry["Дата виїзду"] = data.departureDate || "";
  entry["Таймінг"] = data.timing || "";
  entry["Номер авто"] = data.vehicle || "";
  entry["Місце в авто"] = data.seat || "";
  entry["Ціна квитка"] = data.price || "";
  entry["Завдаток"] = data.deposit || "";
  entry["Статус оплати"] = data.payStatus || "Не оплачено";
  entry["Форма оплати"] = data.payForm || "";
  entry["Валюта"] = data.currency || "UAH";
  entry["Вага багажу"] = data.weight || "";
  entry["Ціна багажу"] = data.weightPrice || "";
  entry["Статус ліда"] = "Новий";
  entry["Статус CRM"] = "Активний";

  appendRow(sheet, PASSENGER_COLUMNS, entry);

  // Знаходимо/створюємо клієнта
  if (data.phone) {
    var client = findOrCreateClient(data.smartId || "", data.phone, data.name || "");
    // Збільшуємо лічильник рейсів
    incrementClientCounter(client, "К-сть рейсів");
  }

  writeLog(data.createdBy || "Система", "Менеджер", "Створив",
    "Passengers_crm", sheetName, paxId, "", "");

  return paxId;
}

/**
 * Створює нову посилку
 */
function createParcel(data, sheetName) {
  sheetName = sheetName || CONFIG.SHEETS.PKG_UA_EU;
  var sheet = getSheet(CONFIG.SPREADSHEETS.PARCELS, sheetName);
  var columns = sheetName === CONFIG.SHEETS.PKG_UA_EU ? PARCEL_UA_EU_COLUMNS : PARCEL_EU_UA_COLUMNS;

  var pkgId = generateID(CONFIG.PREFIXES.PKG);
  var entry = {};
  for (var i = 0; i < columns.length; i++) {
    entry[columns[i]] = "";
  }

  entry["PKG_ID"] = pkgId;
  entry["Ід_смарт/CRM"] = data.smartId || "";
  entry["Напрям"] = sheetName === CONFIG.SHEETS.PKG_UA_EU ? "УК → ЄВ" : "ЄВ → УК";
  entry["SOURCE_SHEET"] = sheetName;
  entry["Дата створення"] = new Date();
  entry["Піб відправника"] = data.senderName || "";
  entry["Телефон реєстратора"] = data.phoneReg || "";
  entry["Піб отримувача"] = data.receiverName || "";
  entry["Телефон отримувача"] = data.receiverPhone || "";
  entry["Внутрішній номер№"] = data.internalNumber || "";
  entry["Опис"] = data.description || "";
  entry["Деталі"] = data.details || "";
  entry["Кількість позицій"] = data.qty || "";
  entry["Кг"] = data.weight || "";
  entry["Фото посилки"] = data.photo || "";
  entry["Дата відправки"] = data.sendDate || "";
  entry["Таймінг"] = data.timing || "";
  entry["Номер авто"] = data.vehicle || "";
  entry["Сума"] = data.amount || "";
  entry["Валюта оплати"] = data.currency || "UAH";
  entry["Форма оплати"] = data.payForm || "";
  entry["Статус оплати"] = data.payStatus || "Не оплачено";
  entry["Статус посилки"] = "В дорозі";
  entry["Статус"] = "Новий";
  entry["Статус CRM"] = "Активний";

  // Специфічні поля для аркушів
  if (sheetName === CONFIG.SHEETS.PKG_UA_EU) {
    entry["Адреса в Європі"] = data.addressEU || "";
    entry["Номер ТТН"] = data.ttn || "";
    entry["Оціночна вартість"] = data.estimatedValue || "";
  } else {
    entry["Місто Нова Пошта"] = data.novaPoshtaCity || "";
  }

  appendRow(sheet, columns, entry);

  // Знаходимо/створюємо клієнта по телефону отримувача або відправника
  var clientPhone = data.receiverPhone || data.phoneReg || "";
  var clientName = data.receiverName || data.senderName || "";
  if (clientPhone) {
    var client = findOrCreateClient(data.smartId || "", clientPhone, clientName);
    incrementClientCounter(client, "К-сть посилок");
  }

  writeLog(data.createdBy || "Система", "Менеджер", "Створив",
    "Posylki_crm_v2", sheetName, pkgId, "", "");

  return pkgId;
}

/**
 * Створює запис маршруту
 */
function createRoute(data, sheetName) {
  sheetName = sheetName || CONFIG.SHEETS.RTE_ZURICH;
  var sheet = getSheet(CONFIG.SPREADSHEETS.ROUTES, sheetName);

  var rteId = generateID(CONFIG.PREFIXES.RTE);
  var entry = {};
  for (var i = 0; i < ROUTE_COLUMNS.length; i++) {
    entry[ROUTE_COLUMNS[i]] = "";
  }

  entry["RTE_ID"] = rteId;
  entry["Тип запису"] = data.type || "🧍 Пасажир";
  entry["Напрям"] = data.direction || "";
  entry["SOURCE_SHEET"] = sheetName;
  entry["PAX_ID / PKG_ID"] = data.linkedId || "";
  entry["Дата створення"] = new Date();
  entry["Дата рейсу"] = data.tripDate || "";
  entry["Таймінг"] = data.timing || "";
  entry["Номер авто"] = data.vehicle || "";
  entry["Водій"] = data.driver || "";
  entry["Телефон водія"] = data.driverPhone || "";
  entry["Місце в авто"] = data.seat || "";
  entry["Піб пасажира"] = data.passengerName || "";
  entry["Телефон пасажира"] = data.passengerPhone || "";
  entry["Адреса прибуття"] = data.address || "";
  entry["Кількість місць"] = data.seats || "";
  entry["Вага багажу"] = data.weight || "";
  entry["Піб відправника"] = data.senderName || "";
  entry["Піб отримувача"] = data.receiverName || "";
  entry["Телефон отримувача"] = data.receiverPhone || "";
  entry["Адреса отримувача"] = data.receiverAddress || "";
  entry["Внутрішній номер№"] = data.internalNumber || "";
  entry["Номер ТТН"] = data.ttn || "";
  entry["Опис посилки"] = data.parcelDescription || "";
  entry["Кг посилки"] = data.parcelWeight || "";
  entry["Сума"] = data.amount || "";
  entry["Валюта"] = data.currency || "";
  entry["Форма оплати"] = data.payForm || "";
  entry["Статус оплати"] = data.payStatus || "";
  entry["Статус"] = "Новий";
  entry["Статус CRM"] = "Активний";
  entry["Примітка"] = data.note || "";

  appendRow(sheet, ROUTE_COLUMNS, entry);

  writeLog(data.createdBy || "Система", "Менеджер", "Створив",
    "Marhrut_crm_v3", sheetName, rteId, "", "");

  return rteId;
}


// ══════════════════════════════════════════════════════════════
// ОНОВЛЕННЯ СТАТУСУ
// ══════════════════════════════════════════════════════════════

/**
 * Оновлює статус маршруту і тригерить авто-рейтинг при "Виконано"
 */
function updateRouteStatus(rteId, routeSheetName, newStatus, updatedBy) {
  var routeSheet = getSheet(CONFIG.SPREADSHEETS.ROUTES, routeSheetName);
  var found = findRowByColumn(routeSheet, "RTE_ID", rteId);
  if (!found) {
    throw new Error("Маршрут " + rteId + " не знайдено");
  }

  var headers = found.headers;
  var row = rowToObject(headers, found.data);
  var oldStatus = row["Статус"];

  row["Статус"] = newStatus;
  if (newStatus === "Виконано" || newStatus === "Скасовано") {
    row["Статус CRM"] = newStatus;
  }

  updateRow(routeSheet, found.row, headers, row);

  writeLog(updatedBy || "Система", "Менеджер", "Змінив",
    "Marhrut_crm_v3", routeSheetName, rteId, "Статус",
    oldStatus + " → " + newStatus);

  // Якщо статус "Виконано" — тригеримо авто-рейтинг
  if (newStatus === "Виконано") {
    onRouteComplete(rteId, routeSheetName);
  }
}


// ══════════════════════════════════════════════════════════════
// ДОПОМІЖНІ ФУНКЦІЇ КЛІЄНТІВ
// ══════════════════════════════════════════════════════════════

/**
 * Збільшує лічильник клієнта (рейси або посилки)
 */
function incrementClientCounter(clientResult, counterField) {
  var clientSheet = clientResult.sheet;
  var headers = clientResult.headers;
  var data = clientSheet.getRange(clientResult.row, 1, 1, headers.length).getValues()[0];
  var clientRow = rowToObject(headers, data);

  clientRow[counterField] = (Number(clientRow[counterField]) || 0) + 1;
  clientRow["Остання активність"] = new Date();

  updateRow(clientSheet, clientResult.row, headers, clientRow);
}


// ══════════════════════════════════════════════════════════════
// WEB APP — doGet / doPost
// ══════════════════════════════════════════════════════════════

function doGet(e) {
  var action = e.parameter.action;
  var result = { success: false, error: "Unknown action" };

  try {
    if (action === "getPassengers") {
      var sheetName = e.parameter.sheet || CONFIG.SHEETS.PAX_UA_EU;
      result = getAllRecords(CONFIG.SPREADSHEETS.PASSENGERS, sheetName);
    }
    else if (action === "getParcels") {
      var sheetName = e.parameter.sheet || CONFIG.SHEETS.PKG_UA_EU;
      result = getAllRecords(CONFIG.SPREADSHEETS.PARCELS, sheetName);
    }
    else if (action === "getRoutes") {
      var sheetName = e.parameter.sheet || CONFIG.SHEETS.RTE_ZURICH;
      result = getAllRecords(CONFIG.SPREADSHEETS.ROUTES, sheetName);
    }
    else if (action === "getClients") {
      result = getAllRecords(CONFIG.SPREADSHEETS.CLIENTS, CONFIG.SHEETS.CLIENTS);
    }
    else if (action === "getClientRatings") {
      var clientId = e.parameter.clientId;
      result = getClientRatings(clientId);
    }
    else if (action === "getClientReviews") {
      result = getAllRecords(CONFIG.SPREADSHEETS.CLIENTS, CONFIG.SHEETS.CLIENT_REVIEWS);
    }
    else if (action === "getArchive") {
      var sheetName = e.parameter.sheet || CONFIG.SHEETS.ARC_PASSENGERS;
      result = getAllRecords(CONFIG.SPREADSHEETS.ARCHIVE, sheetName);
    }
    else if (action === "getLogs") {
      result = getAllRecords(CONFIG.SPREADSHEETS.ARCHIVE, CONFIG.SHEETS.ARC_LOGS);
    }
  } catch (err) {
    result = { success: false, error: err.message };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var body = JSON.parse(e.postData.contents);
  var action = body.action;
  var result = { success: false, error: "Unknown action" };

  try {
    if (action === "createPassenger") {
      var id = createPassenger(body.data, body.sheet);
      result = { success: true, id: id };
    }
    else if (action === "createParcel") {
      var id = createParcel(body.data, body.sheet);
      result = { success: true, id: id };
    }
    else if (action === "createRoute") {
      var id = createRoute(body.data, body.sheet);
      result = { success: true, id: id };
    }
    else if (action === "updateRouteStatus") {
      updateRouteStatus(body.rteId, body.sheet, body.status, body.updatedBy);
      result = { success: true };
    }
    else if (action === "setRating") {
      setManualRating(body.rteId, body.sheet, body.role, body.rating, body.comment, body.raterName);
      result = { success: true };
    }
    else if (action === "processReview") {
      processExternalReview(body.smartId, body.phone, body.rawTextA, body.rawTextB, body.comment);
      result = { success: true };
    }
    else if (action === "archive") {
      var archiveId = archiveRecord(body.sourceType, body.sheet, body.recordId,
        body.archivedBy, body.reason, body.actionType);
      result = { success: true, archiveId: archiveId };
    }
    else if (action === "restore") {
      restoreFromArchive(body.archiveId, body.sheet, body.restoredBy, body.reason);
      result = { success: true };
    }
    else if (action === "updateField") {
      updateSingleField(body.spreadsheetId, body.sheet, body.idColumn, body.recordId,
        body.field, body.value, body.updatedBy);
      result = { success: true };
    }
  } catch (err) {
    result = { success: false, error: err.message };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}


// ══════════════════════════════════════════════════════════════
// ДОПОМІЖНІ ФУНКЦІЇ API
// ══════════════════════════════════════════════════════════════

/**
 * Отримує всі записи з аркуша як масив об'єктів
 */
function getAllRecords(spreadsheetId, sheetName) {
  var sheet = getSheet(spreadsheetId, sheetName);
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { success: true, data: [] };

  var headers = data[0];
  var records = [];
  for (var i = 1; i < data.length; i++) {
    records.push(rowToObject(headers, data[i]));
  }
  return { success: true, data: records };
}

/**
 * Отримує рейтинги конкретного клієнта
 */
function getClientRatings(clientId) {
  var sheet = getSheet(CONFIG.SPREADSHEETS.CLIENTS, CONFIG.SHEETS.CLIENT_RATINGS);
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { success: true, data: [] };

  var headers = data[0];
  var clientIdCol = headers.indexOf("CLIENT_ID");
  var records = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][clientIdCol] == clientId) {
      records.push(rowToObject(headers, data[i]));
    }
  }
  return { success: true, data: records };
}

/**
 * Оновлює одне поле запису з логуванням
 */
function updateSingleField(spreadsheetId, sheetName, idColumn, recordId, field, value, updatedBy) {
  var sheet = getSheet(spreadsheetId, sheetName);
  var found = findRowByColumn(sheet, idColumn, recordId);
  if (!found) {
    throw new Error("Запис " + recordId + " не знайдено");
  }

  var headers = found.headers;
  var row = rowToObject(headers, found.data);
  var oldValue = row[field];
  row[field] = value;

  updateRow(sheet, found.row, headers, row);

  // Визначаємо назву таблиці для логу
  var tableName = "";
  if (spreadsheetId === CONFIG.SPREADSHEETS.PASSENGERS) tableName = "Passengers_crm";
  else if (spreadsheetId === CONFIG.SPREADSHEETS.PARCELS) tableName = "Posylki_crm_v2";
  else if (spreadsheetId === CONFIG.SPREADSHEETS.ROUTES) tableName = "Marhrut_crm_v3";
  else if (spreadsheetId === CONFIG.SPREADSHEETS.CLIENTS) tableName = "Kliyentu_crm_final";

  writeLog(updatedBy || "Система", "Менеджер", "Змінив",
    tableName, sheetName, recordId, field,
    (oldValue || "—") + " → " + (value || "—"));
}
