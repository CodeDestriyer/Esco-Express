# ПРОМТ: EscoExpress CRM Посилки — Повна документація для розробника

> Цей файл містить 3 окремі промти для відтворення Посилкової CRM з нуля.
> Кожен промт самодостатній — можна використовувати окремо.

---

## ПРОМТ 1: ДИЗАЙН (UI/UX)

```
Створи дизайн CRM-системи для управління посилками.
Технологія: один HTML файл, весь CSS всередині <style>, vanilla JavaScript,
шрифт Google Fonts Montserrat (400, 500, 600, 700, 800).

=== КОЛЬОРОВА СХЕМА ===

:root {
  --primary: #1a3a5e;        /* темно-синій — основний */
  --primary-dark: #0d1f3c;   /* ще темніший — градієнт */
  --accent: #d90d0d;         /* червоний — логотип, акценти */
  --success: #10b981;        /* зелений — оплачено, підтверджено */
  --warning: #f59e0b;        /* жовтий — частково, новий */
  --danger: #ef4444;         /* червоний — відмова, борг */
  --info: #3b82f6;           /* синій — в роботі */
  --purple: #8b5cf6;         /* фіолетовий — маршрут, вага */
  --orange: #f97316;
  --pink: #ec4899;
  --teal: #14b8a6;
  --bg-main: #f5f7fa;        /* фон сторінки */
  --text-primary: #1e293b;   /* основний текст */
  --text-secondary: #64748b; /* другорядний текст */
  --border: #e2e8f0;         /* рамки */
  --shadow: 0 2px 8px rgba(0,0,0,0.08);
}

=== HEADER (sticky, висота 52px) ===

- Фон: gradient(135deg, --primary → --primary-dark)
- Логотип: "EscoExpress" — слово "Express" кольором --accent, font-weight 800, 13px
- Пошук: input з іконкою 🔍 зліва, фон rgba(255,255,255,0.15), білий текст,
  placeholder напівпрозорий, при фокусі фон rgba(255,255,255,0.25)
- Справа: індикатор синхронізації (сірий текст 10px) + аватар (30x30, червоне коло,
  ініціали білим, font-weight 700)
- Бургер-кнопка: тільки на мобільному (display: none на десктопі)

=== БОКОВА ПАНЕЛЬ — ДЕСКТОП (PC Sidebar, 300px) ===

- Position: sticky, top: 52px, height: calc(100vh - 52px)
- Фон: gradient(180deg, #f8fafc → #f1f5f9)
- Box-shadow: 4px 0 20px rgba(0,0,0,0.08)
- Може складатися (collapsed): width 0, overflow hidden

СЕКЦІЇ (складні, з toggle стрілкою):

1. НАПРЯМОК (🧭):
   - Пункти: "УК → Європа", "Європа → УК"
   - Кожен пункт: біла картка, border-radius 8px, shadow 0 1px 2px
   - Hover: border primary, shadow збільшується, translateX(3px)
   - Active УК→ЄВ: gradient(#dbeafe → #bfdbfe), color #1d4ed8
   - Active ЄВ→УК: gradient(#d1fae5 → #a7f3d0), color #047857
   - Лічильники: badge з кількістю (font-size 11px, font-weight 700)

2. ПЕРЕВІРКА (🔍):
   - Пункти: "Всі", "В перевірці", "Готові", "Невідомі"
   - Active: фон --primary, текст білий
   - Лічильники кожного статусу

3. МАРШРУТИ (🗺️):
   - Список маршрутів динамічний (з API)
   - Кожен: назва + "👤 X 📦 Y" лічильники
   - Active: фон --primary, текст білий
   - Кнопка "+ Новий маршрут": dashed border, текст --info

=== БОКОВА ПАНЕЛЬ — МОБІЛЬНА (< 768px) ===

- Position: fixed, зліва, width 200px, slides з left: -200px → 0
- Фон: gradient(180deg, #1e3a5f → #0f2744)
- Overlay: rgba(0,0,0,0.5), z-index 200
- Пункти: білий текст, фон rgba(255,255,255,0.08), hover 0.15
- Active: фон rgba(255,255,255,0.25)

=== ГОЛОВНИЙ КОНТЕНТ ===

- flex: 1, padding 12px, overflow-y auto

ТАБКИ НАПРЯМІВ (dir-tabs):
- 2 кнопки: "🇺🇦 УК → Європа", "🇪🇺 Європа → УК"
- Кожна: border 2px solid --border, border-radius 10px, padding 10px
- Active УК→ЄВ: border #2563eb, фон #eff6ff, текст #1d4ed8
- Active ЄВ→УК: border #059669, фон #ecfdf5, текст #047857

=== КАРТКА ПОСИЛКИ ===

Контейнер:
- Фон білий, border-radius 8px, shadow --shadow
- border-left: 4px solid (колір за статусом):
  • Новий → --warning (жовтий)
  • В роботі → --info (синій)
  • Підтверджено → --success (зелений)
  • Відмова → --danger (червоний)

Структура картки (3 рядки):

Рядок 1 (card-top): горизонтальний flex
- ☐ Чекбокс (ліворуч, фон #f8fafc)
- Badge напряму: "УК→ЄВ" (синій) або "ЄВ→УК" (зелений)
- Телефон (font-weight 700, font-size 14px)
- Вага: badge фіолетовий (#e0e7ff, текст #3730a3), "⚖️ X кг"
- Ціна: справа, великий (font-size 16px), колір --success, font-weight 800

Рядок 2 (card-info): горизонтальний flex
- Іконка: 📦
- ПІБ відправника (font-weight 600)
- PKG_ID (сірий, font-size 10px)
- Badge статусу ліда (Новий/В роботі/Підтверджено/Відмова)
- Badge оплати (Оплачено/Частково/Не оплачено)

Рядок 3 (card-route): якщо є маршрут
- "📍 Адреса звідки → Адреса куди"

Рядок 4 (card-meta): теги
- 🚐 Номер авто, 🧑 Водій, 📍 Адреса (обрізана 40 символів)
- 💰 Борг: X (червоний), 📝 Примітка (обрізана 30 символів)

=== БЕЙДЖІ СТАТУСІВ ===

Загальний стиль: padding 2px 8px, border-radius 12px, font-size 10px, font-weight 700

- badge-new: фон #dbeafe, текст #1d4ed8
- badge-work: фон #fef3c7, текст #92400e
- badge-confirmed: фон #d1fae5, текст #065f46
- badge-refused: фон #fee2e2, текст #991b1b
- badge-paid: фон #d1fae5, текст #065f46
- badge-partial: фон #fef3c7, текст #92400e
- badge-unpaid: фон #fee2e2, текст #991b1b

=== МАРШРУТНА КАРТКА (route-card) ===

Те саме що картка посилки, але з доповненнями:
- Тип: 👤 (пасажир) або 📦 (посилка)
- Для пасажирів: badge "💺 місце", кількість місць
- Розгортається при кліку → details-grid:
  • Grid auto-fill, minmax(200px, 1fr)
  • Кожне поле: label (сірий 10px) + value + кнопка ✏️ (inline edit)
- Кнопки внизу деталей: 📞 Дзвінок, ✉️ Писати, 🗑️ З маршруту

=== МОДАЛЬНІ ВІКНА ===

Overlay: fixed, inset 0, rgba(0,0,0,0.5), z-index 600
Модалка: білий фон, border-radius 16px, max-width 500px, центр екрану

Bottom Sheet (форма створення):
- Знизу екрану, border-radius 16px 16px 0 0
- Ручка зверху: сіра полоска 40x4px, margin auto
- Секції що складаються з анімацією max-height

=== МАСОВІ ДІЇ (Bulk Toolbar) ===

- Fixed, bottom 60px, z-index 150
- Фон --primary з backdrop-filter blur
- Десктоп (> 901px): в сайдбарі, вертикально, width 300px
- Кнопки: маршрут (--purple), архів (сірий), оптимізація (--purple), скасувати

=== НИЖНЯ НАВІГАЦІЯ (Mobile, < 768px) ===

- Fixed bottom 0, білий фон, border-top 1px
- 5 кнопок: Посилки, Пасажири, Оновити, Колонки, Додати
- Іконки 18px, текст 9px
- Active: колір --accent

=== RESPONSIVE BREAKPOINTS ===

> 901px: sidebar 300px + контент, bulk toolbar в sidebar
768-900px: sidebar collapsed, burger menu, bulk toolbar fixed bottom
< 768px: мобільне меню, bottom nav, bottom sheet для форм
< 480px: details-grid 1 колонка, card-top wrap
```

---

## ПРОМТ 2: ФУНКЦІОНАЛ (Напрями, Перевірка, Маршрути)

```
Створи фронтенд CRM-системи для управління посилками.
Технологія: vanilla JS, один HTML файл, бекенд — Google Apps Script.
Всі запити через fetch POST на GAS_URL, формат: { action: 'назва', ...params }
Content-Type: 'text/plain;charset=utf-8' (не application/json — GAS не підтримує CORS з json)

=== ГЛОБАЛЬНИЙ СТАН (State) ===

let allData = [];                // всі посилки з API
let filteredData = [];           // після фільтрації
let currentDirection = 'ue';     // 'ue' або 'eu'
let currentFilter = 'all';      // фільтр статусу посилки
let currentPayFilter = 'all';   // фільтр оплати
let searchQuery = '';            // пошуковий запит
let openCardId = null;           // розгорнута картка
let openActionsId = null;        // відкриті дії
let stats = {};                  // статистика
let routes = [];                 // масив маршрутів
let activeRouteIdx = null;       // індекс активного маршруту
let currentView = 'parcels';     // 'parcels' | 'verification' | 'routes'
let selectedIds = new Set();     // вибрані для bulk операцій
let routeSelectedIds = new Set();// вибрані в маршруті
let hasNpApiKey = false;         // чи є ключ Нової Пошти

=== ІНІЦІАЛІЗАЦІЯ (DOMContentLoaded) ===

Паралельно завантажити:
1. apiPost('getAll', { sheet: 'all', filter: {} }) → allData
2. apiPost('getStats', {}) → stats
3. apiPost('checkNpApiKey', {}) → hasNpApiKey
4. apiPost('getRoutesList', {}) → routes (lazy, rows = null)

Після завантаження: renderCards(), renderRouteSidebar(), updateCounters()

=== МОДУЛЬ 1: НАПРЯМИ (Directions) ===

ДВА НАПРЯМИ:
- 'ue' = УК → Європа (sheet: "Реєстрація ТТН УК-єв")
- 'eu' = Європа → УК (sheet: "Виклик Курєра ЄВ-ук")

setDirection(dir):
  → currentDirection = dir
  → перефільтрувати allData за напрямом (obj['Напрям'])
  → УК→ЄВ: 'УК→ЄВ' | ЄВ→УК: 'ЄВ→УК'
  → оновити активну вкладку в sidebar
  → renderCards()

ФІЛЬТРАЦІЯ (filterData):
  1. За напрямом (currentDirection)
  2. За статусом CRM: тільки 'Активний' (не 'Архів')
  3. За статусом посилки (currentFilter): 'all' або конкретний
  4. За статусом оплати (currentPayFilter): 'all' або конкретний
  5. За пошуком (searchQuery): шукати в ПІБ відправника, ПІБ отримувача,
     Телефон реєстратора, Телефон отримувача, PKG_ID, Номер ТТН

renderCards():
  → filterData() → для кожної посилки renderCard(p)
  → Якщо 0 результатів — показати "Посилок не знайдено"
  → Оновити лічильники в sidebar

КАРТКА ПОСИЛКИ — ДАНІ:
  - pkgId = p['PKG_ID']
  - name = p['Піб відправника'] || '—'
  - phone = p['Телефон реєстратора'] || '—'
  - receiver = p['Піб отримувача']
  - phoneRecv = p['Телефон отримувача']
  - direction = p['Напрям']
  - addressFrom = p['Адреса відправки']
  - addressTo = p['Адреса в Європі'] || p['Місто Нова Пошта']
  - weight = p['Кг']
  - price = p['Сума'], currency = p['Валюта оплати']
  - deposit = p['Завдаток'], debt = p['Борг']
  - payStatus = p['Статус оплати']
  - leadStatus = p['Статус ліда']
  - pkgStatus = p['Статус посилки']
  - ttn = p['Номер ТТН']
  - internalNum = p['Внутрішній №']
  - routeId = p['RTE_ID']
  - auto = p['Номер авто']

ДІЇ НА КАРТЦІ:
  1. Дзвінок → window.open('tel:' + phone)
  2. Написати → messenger popup (Viber, Telegram, WhatsApp)
  3. Відстежити ТТН → apiPost('trackParcel', { pkg_id, ttn })
  4. Маршрут → openRouteModal(pkgId) — вибір маршруту
  5. Видалити → apiPost('deleteParcel', { pkg_id }) з confirm()

INLINE EDITING (розгорнуті деталі):
  - Клік на значення поля → замінити на <input> з поточним value
  - Enter або blur → apiPost('updateField', { pkg_id, col, value })
  - Escape → скасувати, повернути текст
  - Після збереження — оновити allData локально, перерендерити

ФОРМА СТВОРЕННЯ ПОСИЛКИ (Bottom Sheet):
  Поля:
  - Напрям (2 кнопки: УК→ЄВ / ЄВ→УК)
  - Піб відправника* (обов'язкове), Телефон*
  - Адреса відправки*
  - Піб отримувача*, Телефон отримувача*
  - Адреса (для УК→ЄВ: "Адреса в Європі", для ЄВ→УК: "Місто Нова Пошта")
  - Номер ТТН (тільки для УК→ЄВ)
  - Опис, Деталі, Кількість позицій, Кг
  - Оціночна вартість, Сума, Валюта, Завдаток

  При введенні ПІБ/телефону — debounce 500ms → apiPost('checkDuplicates', { pib, phone })
  Якщо є дублікати — показати попередження

  Submit → apiPost('addParcel', { sheet: currentDirection, data: {...} })
  Після створення: додати в allData, renderCards()

=== МОДУЛЬ 2: ПЕРЕВІРКА (Verification) ===

ПЕРЕХІД:
  currentView = 'verification'
  Завантажити: apiPost('getVerificationStats', {})

ВКЛАДКИ:
  - "Всі" — всі посилки в перевірці
  - "В перевірці" — Контроль перевірки = "В перевірці"
  - "Готові" — Контроль перевірки = "Готова до маршруту"
  - "Невідомі" — Статус ліда = "Невідомий"

СКАНУВАННЯ ТТН:
  Поле введення ТТН (або сканер штрих-коду)
  → apiPost('scanTTN', { ttn: value })

  ВІДПОВІДЬ ТИП A (знайдено в базі):
    → Показати повну картку посилки
    → Статус автоматично = "В перевірці"
    → Показати дублікати по отримувачу (якщо є)
    → Можна редагувати поля inline

  ВІДПОВІДЬ ТИП B (не знайдено):
    → Створити нову картку "Невідома посилка"
    → Підставити ТТН зі сканера
    → Всі поля порожні — заповнювати вручну
    → Статус = "Невідомий"

КАРТКА ПЕРЕВІРКИ (verification card):
  Все що в звичайній картці + мета-теги:
  - 📷 Фото: є/немає (зелений/червоний)
  - 📋 Дані: повні/неповні
  - ✏️ Редаговано (жовтий)
  - 🔄 Дублікат
  - 🗺️ Маршрут (якщо призначено)
  - ❓ Невідома (для типу B)

ПЕРЕВІРКА ПОВНОТИ ДАНИХ:
  Поля що мають бути заповнені:
  ☐ Фото посилки
  ☐ Кг (вага)
  ☐ Кількість позицій
  ☐ Опис
  ☐ Внутрішній №
  ☐ Сума
  ☐ Статус оплати

ДІЇ:
  1. ✅ Завершити перевірку:
     → Валідація: Внутрішній № обов'язковий, Сума обов'язкова
     → apiPost('completeVerification', { pkg_id, params })
     → Контроль перевірки = "Готова до маршруту"

  2. ❌ Відхилити:
     → Введення причини
     → apiPost('rejectVerification', { pkg_id, reason })

  3. 📷 Додати фото:
     → Upload на UPLOAD_URL → отримати URL
     → apiPost('addPhoto', { pkg_id, url, type })

  4. 🗺️ Додати в маршрут → openRouteModal(pkgId)

МАСОВІ ДІЇ ПЕРЕВІРКИ:
  - Масове завершення перевірки (всі вибрані)
  - Масовий перенос іншому менеджеру
  - Масове повторне відкриття
  - Масове призначення маршруту

=== МОДУЛЬ 3: МАРШРУТИ (Routes) ===

ПЕРЕХІД:
  currentView = 'routes'

SIDEBAR МАРШРУТІВ:
  routes[] — масив об'єктів:
    { sheetName: 'Маршрут_Цюріх', rowCount: 5, paxCount: 2, parcelCount: 3, rows: null }

  rows = null означає що дані ще не завантажені (lazy loading)
  При кліку на маршрут:
    → activeRouteIdx = idx
    → Якщо rows === null → loadRouteSheetData(idx):
      apiPost('getRouteSheet', { sheetName }) → зберегти в routes[idx].rows
    → renderRoutes()

ВАЖЛИВО: loadRouteSheetData перевіряє if (sheet.rows !== null) return;
  Тому для перезавантаження потрібно спочатку: routes[idx].rows = null;

renderRoutes():
  → Показати назву маршруту, кількість пасажирів/посилок
  → Фільтри: за типом (всі/пасажири/посилки), за статусом, за оплатою
  → Для кожного рядка → renderRouteCard(r, idx, sheetName)

renderRouteCard(r, idx, sheetName):
  Unified картка — і для пасажирів і для посилок:

  Визначення типу:
    type = r['Тип запису'] → isPax = type.includes('Пасажир')

  Читання полів (з fallback для різних назв колонок!):
    name = isPax ? r['Піб пасажира'] : (r['Піб відправника'] || r['Піб пасажира'] || '—')
    phone = r['Телефон'] || r['Телефон пасажира'] || r['Телефон реєстратора'] || '—'
    receiver = r['Піб отримувача'] || ''
    address = r['Адреса'] || r['Адреса відправки'] || ''
    addressTo = r['Адреса прибуття'] || r['Адреса отримувача'] || r['Адреса в Європі'] || ''
    kg = r['Кг'] || r['Кг посилки'] || ''
    price = r['Сума'], currency = r['Валюта']
    debt = r['Борг'], payStatus = r['Статус оплати']
    status = r['Статус'] || r['Статус ліда']

  Маршрут-лінія:
    Для посилок: "📍 ПІБ відправника → ПІБ отримувача"
    Для пасажирів: "📍 Адреса відправки → Адреса прибуття"

  Деталі (розгортаються по кліку):
    Grid з усіма полями + кнопка ✏️ для inline edit
    Inline edit → apiPost('updateRouteField', { sheet, rte_id, col, value })

ДОДАВАННЯ ПОСИЛКИ В МАРШРУТ:

  1. Користувач натискає "Маршрут" на картці посилки
  2. openRouteModal(pkgId) — показати список маршрутів
  3. Користувач обирає маршрут → selectRoute(sheetName)
  4. Формуємо дані: buildParcelRouteData(p):

     КРИТИЧНО — маппінг полів повинен включати ВСІ варіанти назв колонок
     таблиці маршруту, бо назви можуть відрізнятись:

     {
       'RTE_ID': pkgId,
       'PAX_ID/PKG_ID': pkgId,
       'Тип запису': 'Посилка',
       // Відправник (всі варіанти назв)
       'Піб відправника': p['Піб відправника'],
       'Телефон': p['Телефон реєстратора'],
       'Телефон пасажира': p['Телефон реєстратора'],
       'Телефон реєстратора': p['Телефон реєстратора'],
       // Отримувач
       'Піб отримувача': p['Піб отримувача'],
       'Телефон отримувача': p['Телефон отримувача'],
       'Адреса отримувача': p['Адреса в Європі'] || p['Місто Нова Пошта'],
       // Адреси (всі варіанти)
       'Адреса': p['Адреса відправки'],
       'Адреса відправки': p['Адреса відправки'],
       'Адреса прибуття': p['Адреса в Європі'] || p['Місто Нова Пошта'],
       // Посилка (всі варіанти)
       'Кг': p['Кг'],
       'Кг посилки': p['Кг'],
       'Вага багажу': p['Кг'],
       'Опис': p['Опис'],
       'Опис посилки': p['Опис'],
       // ... решта полів: Сума, Валюта, Завдаток, Борг, Статус, тощо
     }

  5. apiPost('addToRoute', { pkg_id, sheet_name, rte_id, lead_data })
  6. Після успіху: routes[idx].rows = null → перезавантажити → renderRoutes()

ВИДАЛЕННЯ З МАРШРУТУ:
  → apiPost('removeFromRoute', { pkg_id })
  → Очищує RTE_ID, Номер авто, Дата відправки в посилці

CRUD МАРШРУТІВ:
  - Створити: prompt('Назва') → apiPost('createRoute', { name })
    → Копіює 3 шаблони: Маршрут_Назва, Відправка_Назва, Витрати_Назва
  - Видалити: confirm() → apiPost('deleteRoute', { name })

МАСОВІ ОПЕРАЦІЇ:
  - Чекбокси на картках → selectedIds / routeSelectedIds
  - Bulk toolbar: масове додавання в маршрут, масовий архів
  - Масовий вибір: "Обрати всі" toggle

ОПТИМІЗАЦІЯ МАРШРУТУ:
  Крок 1: Обрати метод — за адресою відправки або адресою прибуття
  Крок 2: Ввести стартову адресу
  Крок 3: Геокодування адрес → розрахунок оптимального порядку
  Результат:
    - Оптимальний порядок зупинок
    - Загальна відстань та час
    - ETA для кожної зупинки
    - Посилання на Google Maps / 2GIS
  Дії: зберегти порядок в маршрут
```

---

## ПРОМТ 3: GAS СКРИПТ (Backend)

```
Створи Google Apps Script (GAS) бекенд для CRM посилок.
Deploy: Web App, Execute as: Me, Access: Anyone.
Один файл .gs з doPost що обробляє всі запити через action router.

=== АРХІТЕКТУРА ===

Система: 4 окремі GAS скрипти (Passengers, Posylki, Driver, Client).
Цей скрипт — Posylki (для менеджерів посилок).
Кожен має свій doPost, свій URL, свою чергу запитів.

=== БАЗИ ДАНИХ (Google Sheets) ===

var DB = {
  POSYLKI:  'ID_ТАБЛИЦІ_ПОСИЛОК',
  MARHRUT:  'ID_ТАБЛИЦІ_МАРШРУТІВ',
  KLIYENTU: 'ID_ТАБЛИЦІ_КЛІЄНТІВ',
  FINANCE:  'ID_ТАБЛИЦІ_ФІНАНСІВ',
  CONFIG:   'ID_ТАБЛИЦІ_КОНФІГУРАЦІЇ',
  ARCHIVE:  'ID_ТАБЛИЦІ_АРХІВУ'
};

АРКУШІ В POSYLKI:
- "Реєстрація ТТН УК-єв" — посилки УК→ЄВ (52 колонки)
- "Виклик Курєра ЄВ-ук" — посилки ЄВ→УК (51 колонка)
- "Фото посилок" — фото (12 колонок)

АРКУШІ В MARHRUT:
- "Маршрут_Назва" — дані маршруту (пасажири + посилки)
- "Відправка_Назва" — дані відправки
- "Витрати_Назва" — витрати
- "Маршрут_Шаблон" / "Відправка_Шаблон" / "Витрати_Шаблон" — шаблони

=== КОЛОНКИ ПОСИЛОК ===

Спільні для обох напрямів (порядок важливий — це порядок колонок в таблиці):

УК→ЄВ (PKG_UE_COLS, 52 колонки):
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

ЄВ→УК (PKG_EU_COLS, 51 колонка):
Те саме, але замість 'Адреса в Європі' → 'Місто Нова Пошта',
замість 'Номер ТТН' → немає (є 'НП активна' додатково).

ФОТО (PHOTO_COLS, 12 колонок):
'PHOTO_ID','PKG_ID','Номер ТТН','Штрих-код ТТН',
'Тип фото','Фото посилки','Хто завантажив','Роль',
'Коментар','Статус перевірки','Ід реєстратора','Час'

ФІНАНСИ (FINANCE_COLS):
'PAY_ID','Дата створення','Хто вніс','Роль',
'CLI_ID','PAX_ID','PKG_ID','RTE_ID','CAL_ID',
'Ід_смарт','Тип платежу','Сума','Валюта',
'Форма оплати','Статус платежу','Борг сума','Борг валюта',
'Дата погашення','Примітка','DATE_ARCHIVE','ARCHIVED_BY'

=== HELPER ФУНКЦІЇ ===

КРИТИЧНО ВАЖЛИВО — нормалізація заголовків:
При читанні заголовків з БУДЬ-ЯКОГО аркуша маршруту ОБОВ'ЯЗКОВО замінювати
переноси рядків на пробіли (Google Sheets може мати \n в заголовках):

headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) {
  return String(h).replace(/[\r\n]+/g, ' ').replace(/\s+/g, ' ').trim();
});

Інші хелпери:
- genId(prefix) → prefix + '_' + timestamp + '_' + random
- today() → 'YYYY-MM-DD'
- now() → ISO timestamp
- getSheet(name) → SpreadsheetApp.openById(SS_ID).getSheetByName(name)
- getAllData(sheet) → { headers: [...], data: [[...], ...] }
- rowToObj(headers, row) → { header1: val1, header2: val2, ... }
- objToRow(headers, obj) → [val1, val2, ...]
- findRow(sheet, colName, value) → { rowNum, headers, data } або null
- findAllRows(sheet, colName, value) → [{ rowNum, headers, data }, ...]
- calcDebt(obj) → Math.max(0, parseFloat(Сума) - parseFloat(Завдаток))
- pkgObjFromData(headers, data, shName, rowNum) → об'єкт з метою (_sheet, _rowNum, Борг)

=== doPost — UNIVERSAL ROUTER ===

function doPost(e) {
  var body = JSON.parse(e.postData.contents);
  var action = body.action;
  var result;

  switch (action) {
    // READ
    case 'getAll':             result = apiGetAll(body); break;
    case 'getOne':             result = apiGetOne(body); break;
    case 'getStats':           result = apiGetStats(body); break;
    case 'getPayments':        result = apiGetPayments(body); break;
    case 'getPhotos':          result = apiGetPhotos(body); break;
    case 'getOrderInfo':       result = apiGetOrderInfo(body); break;
    case 'getVerificationStats': result = apiGetVerificationStats(body); break;

    // CREATE
    case 'addParcel':          result = apiAddParcel(body); break;
    case 'addPhoto':           result = apiAddPhoto(body); break;

    // UPDATE
    case 'updateField':        result = apiUpdateField(body); break;
    case 'checkDuplicates':    result = apiCheckDuplicates(body); break;

    // DELETE
    case 'deleteParcel':       result = apiDeleteParcel(body); break;

    // VERIFICATION
    case 'scanTTN':            result = apiScanTTN(body); break;
    case 'findDuplicatesByRecipient': result = apiFindDuplicatesByRecipient(body); break;
    case 'assignRouteNumber':  result = apiAssignRouteNumber(body); break;
    case 'completeVerification': result = apiCompleteVerification(body); break;
    case 'rejectVerification': result = apiRejectVerification(body); break;

    // ROUTES
    case 'getRoutesList':      result = apiGetRoutesList(body); break;
    case 'getRouteSheet':      result = apiGetRouteSheet(body); break;
    case 'addToRoute':         result = apiAddToRoute(body); break;
    case 'removeFromRoute':    result = apiRemoveFromRoute(body); break;
    case 'createRoute':        result = apiCreateRoute(body); break;
    case 'deleteRoute':        result = apiDeleteRoute(body); break;
    case 'updateRouteField':   result = apiUpdateRouteField(body); break;

    // NOVA POSHTA
    case 'trackParcel':        result = apiTrackParcel(body); break;
    case 'checkNpApiKey':      result = apiCheckNpApiKey(body); break;

    default: result = { ok: false, error: 'Unknown action: ' + action };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

=== ДЕТАЛЬНИЙ ОПИС КОЖНОГО ENDPOINT ===

--- apiGetAll(params) ---
params: { sheet: 'all'|'ue'|'eu', filter: { statusPkg, statusLid, statusOplata, statusCrm, tag, search } }
Логіка:
  → Завантажити дані з обох аркушів (або одного за sheet)
  → Пропустити рядки де PKG_ID пустий І Піб відправника пустий
  → Пропустити Статус CRM = 'Архів'
  → Застосувати фільтри (кожен опціональний)
  → search: шукати в Піб відправника, Піб отримувача, Телефон реєстратора,
    Телефон отримувача, PKG_ID, Номер ТТН (case-insensitive)
  → Повернути: { ok: true, data: [об'єкти посилок] }

--- apiGetOne(params) ---
params: { pkg_id }
  → Шукати в обох аркушах за PKG_ID
  → Повернути: { ok: true, data: об'єкт_посилки }

--- apiGetStats(params) ---
  → Порахувати: total, ue, eu, byStatus, byPay, byPkgStatus, totalDebt
  → Пропустити Архів
  → Повернути: { ok: true, total, ue, eu, byStatus: {...}, byPay: {...}, totalDebt }

--- apiAddParcel(params) ---
params: { sheet: 'ue'|'eu', data: { sender, phone, addressFrom, receiver, ... } }
  → genId('PKG')
  → Створити об'єкт з усіма колонками (PKG_UE_COLS або PKG_EU_COLS)
  → Заповнити поля з data
  → Напрям = isUE ? 'УК→ЄВ' : 'ЄВ→УК'
  → Статус ліда = 'Новий', Статус CRM = 'Активний'
  → appendRow
  → Повернути: { ok: true, pkg_id: newId }

--- apiUpdateField(params) ---
params: { pkg_id, col, value }
  → findRow в обох аркушах
  → sheet.getRange(rowNum, colIdx + 1).setValue(value)
  → Повернути: { ok: true }

--- apiDeleteParcel(params) ---
params: { pkg_id }
  → Встановити Статус CRM = 'Архів', DATE_ARCHIVE = now()
  → Повернути: { ok: true }

--- apiCheckDuplicates(params) ---
params: { pib, phone }
  → Шукати в обох аркушах за ПІБ відправника (case-insensitive) АБО Телефон
  → Повернути: { ok: true, duplicates: [...] }

--- apiGetPayments(params) ---
params: { pkg_id }
  → Відкрити DB.FINANCE, аркуш "Платежі"
  → Знайти всі рядки де PKG_ID збігається
  → Сортувати за датою (новіші спочатку)
  → Повернути: { ok: true, data: [...], summary: { totalPaid, totalDebt } }

--- apiAddPhoto(params) ---
params: { pkg_id, url, type }
  → genId('PHOTO')
  → appendRow в аркуш "Фото посилок"
  → Повернути: { ok: true }

--- apiGetPhotos(params) ---
params: { pkg_id }
  → Знайти всі рядки в "Фото посилок" за PKG_ID
  → Повернути: { ok: true, data: [...] }

--- apiScanTTN(params) ---
params: { ttn }
Логіка:
  ТИП A — знайдено за Номер ТТН:
    → Оновити 'Контроль перевірки' = 'В перевірці'
    → Оновити 'Дата перевірки' = now()
    → Знайти дублікати по отримувачу
    → return { ok: true, type: 'found', data: об'єкт, duplicates: [...] }

  ТИП B — НЕ знайдено:
    → Створити новий запис в УК→ЄВ (за замовчуванням)
    → PKG_ID = genId('PKG')
    → Номер ТТН = ttn
    → Статус ліда = 'Невідомий'
    → Контроль перевірки = 'В перевірці'
    → return { ok: true, type: 'new', data: об'єкт }

--- apiCompleteVerification(params) ---
params: { pkg_id, skip_validation }
  → Знайти посилку
  → Якщо не skip_validation:
    → Перевірити наявність 'Внутрішній №' → помилка якщо пусто
    → (Сума — опціонально)
  → Контроль перевірки = 'Готова до маршруту'
  → return { ok: true }

--- apiRejectVerification(params) ---
params: { pkg_id, reason }
  → Статус ліда = 'Відмова'
  → return { ok: true }

--- apiFindDuplicatesByRecipient(params) ---
params: { pkg_id }
  → Знайти посилку → отримати Піб отримувача, Телефон отримувача
  → Шукати інші посилки з тим самим отримувачем (по імені або телефону)
  → Виключити поточну посилку
  → return { ok: true, duplicates: [...] }

--- apiAssignRouteNumber(params) ---
params: { pkg_id, route_base }
  → Визначити діапазон номерів для маршруту:
    базовий 200 → 200-299, overflow 900+
    базовий 500 → 500-599, overflow 800+
  → Зібрати всі існуючі внутрішні номери цього маршруту
  → Знайти наступний вільний номер
  → Оновити 'Внутрішній №' = nextNum
  → return { ok: true, number: nextNum }

--- apiGetVerificationStats(params) ---
  → Порахувати по кожному статусу перевірки
  → return { ok: true, counts: { all, checking, ready, unknown, ... } }

=== МАРШРУТИ ===

--- apiGetRoutesList(params) ---
  → Кеш: 'routesList_v2', TTL 300 сек (5 хв)
  → Відкрити DB.MARHRUT
  → Пропустити аркуші: /^(Лог|Конфіг|Config|Log|Шаблон|Template)/i
  → Для кожного аркуша: порахувати paxCount і parcelCount за колонкою 'Тип запису'
  → return { ok: true, data: [{ sheetName, rowCount, paxCount, parcelCount }] }

--- apiGetRouteSheet(params) ---
params: { sheetName }
  → Кеш: 'routeSheet_' + sheetName, TTL 180 сек (3 хв)
  → НОРМАЛІЗУВАТИ ЗАГОЛОВКИ: replace(/[\r\n]+/g, ' ').replace(/\s+/g, ' ').trim()
  → Пропустити повністю порожні рядки
  → Дати конвертувати через Utilities.formatDate(val, 'Europe/Kiev', 'dd.MM.yyyy')
  → return { ok: true, data: { sheetName, headers, rows, rowCount } }

--- apiAddToRoute(params) ---
params: { pkg_id, sheet_name, rte_id, lead_data }

КРИТИЧНО — правильний порядок:

  1. Відкрити аркуш маршруту, прочитати headers (з нормалізацією!)

  2. ПРІОРИТЕТ: використати params.lead_data якщо є:
     → Це повний набір полів з фронтенду (buildParcelRouteData)
     → Включає ВСІ варіанти назв колонок
     → Просто скопіювати: lead[key] = lead_data[key]

  3. FALLBACK: apiGetOne(pkg_id) + розширений маппінг:
     → Маппінг ОБОВ'ЯЗКОВО включає ВСІ варіанти назв колонок:
       'Телефон' І 'Телефон пасажира' І 'Телефон реєстратора' ← одне значення
       'Адреса' І 'Адреса відправки' ← одне значення
       'Адреса отримувача' І 'Адреса прибуття' ← Адреса в Європі або Місто НП
       'Кг' І 'Кг посилки' І 'Вага багажу' ← одне значення
       'Опис' І 'Опис посилки' ← одне значення
     → Заповнювати тільки порожні поля (не перезаписувати lead_data)

  4. Гарантувати: RTE_ID, PAX_ID/PKG_ID, Тип запису = 'Посилка'

  5. var row = headers.map(function(h) { return lead[h] || ''; });
     sheet.appendRow(row);

  6. Оновити RTE_ID в посилці: apiUpdateField({ pkg_id, col: 'RTE_ID', value: rteId })

  7. ІНВАЛІДУВАТИ КЕШ:
     cache.remove('routeSheet_' + sheetName);
     cache.remove('routesList_v2');

--- apiRemoveFromRoute(params) ---
params: { pkg_id }
  → Очистити: RTE_ID = '', Номер авто = '', Дата відправки = ''
  → return { ok: true }

--- apiUpdateRouteField(params) ---
params: { sheet, rte_id, col, value }
  → Знайти рядок за RTE_ID або PAX_ID/PKG_ID
  → Оновити значення
  → Інвалідувати кеш маршруту

--- apiCreateRoute(params) ---
params: { name }
  → Перевірити чи не існує 'Маршрут_' + name
  → Скопіювати шаблони: Маршрут_Шаблон → Маршрут_name
  → Скопіювати: Відправка_Шаблон → Відправка_name
  → Скопіювати: Витрати_Шаблон → Витрати_name
  → return { ok: true, created: ['Маршрут_name', 'Відправка_name', 'Витрати_name'] }

--- apiDeleteRoute(params) ---
params: { name }
  → Видалити аркуш 'Маршрут_' + name (з перевіркою існування)
  → return { ok: true }

=== НОВА ПОШТА ===

--- apiTrackParcel(params) ---
params: { pkg_id, ttn }
  → Запит до API Нової Пошти (tracking)
  → Оновити статус посилки
  → return { ok: true, tracking: {...} }

--- apiCheckNpApiKey(params) ---
  → Перевірити чи є ключ API в конфігурації
  → return { ok: true, hasKey: true/false }

=== КЕШУВАННЯ ===

CacheService.getScriptCache():
- 'routesList_v2' → 300 сек (5 хв) — список маршрутів
- 'routeSheet_' + sheetName → 180 сек (3 хв) — дані одного маршруту

Інвалідація при: addToRoute, removeFromRoute, updateRouteField,
  createRoute, deleteRoute

ВАЖЛИВО: cache.put може впасти якщо дані > 100KB — обгортати в try/catch

=== ТИПОВІ ПОМИЛКИ ЯКІ ТРЕБА УНИКАТИ ===

1. ЗАГОЛОВКИ З ПЕРЕНОСАМИ: Google Sheets може мати \n в заголовках.
   ЗАВЖДИ нормалізувати: .replace(/[\r\n]+/g, ' ').replace(/\s+/g, ' ').trim()

2. РІЗНІ НАЗВИ КОЛОНОК: таблиця маршруту має інші назви ніж таблиця посилок.
   Приклад: "Адреса отримувача" в маршруті vs "Адреса в Європі" в посилках.
   ЗАВЖДИ маппити під ВСІ можливі варіанти назв.

3. ІГНОРУВАННЯ lead_data: фронтенд відправляє повний набір даних в lead_data.
   ЗАВЖДИ використовувати lead_data як пріоритет, apiGetOne як fallback.

4. КЕШ НЕ ІНВАЛІДОВАНИЙ: після будь-якої зміни в маршруті — видаляти кеш.

5. Content-Type: фронтенд відправляє 'text/plain;charset=utf-8' (не 'application/json')
   бо GAS не підтримує CORS preflight з application/json.
```
