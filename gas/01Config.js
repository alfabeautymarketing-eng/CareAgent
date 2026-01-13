/**привет
 * =======================================================================================
 * A_Config.gs — ФУНДАМЕНТ СИСТЕМЫ (ЕДИНЫЙ ФАЙЛ КОНФИГУРАЦИИ)
 * ---------------------------------------------------------------------------------------
 * ВЕРСИЯ: 2.2 (Эталонная, без модуля FILE_PROCESSING)
 * ОПИСАНИЕ:
 *   Этот файл — "мозг" и "приборная панель" всей экосистемы. Он содержит все
 *   настройки, имена, списки заголовков и константы.
 *
 *   ГЛАВНЫЙ ПРИНЦИП: Конфиг является "Единым Источником Правды" (Source of Truth).
 *   Скрипты читают этот файл, чтобы понять, как должны выглядеть таблицы. Мы меняем
 *   конфиг, а скрипты приводят таблицы в соответствие, а не наоборот.
 * =======================================================================================
 */

(function (global) {
  // =======================================================================================
  // БЛОК 1: ОБЩИЕ НАСТРОЙКИ СИСТЕМЫ
  // =======================================================================================
  const SETTINGS = {
    /** Проект, используемый по умолчанию при запуске функций напрямую из редактора Apps Script для целей отладки. */
    DEBUG_PROJECT_KEY: "SS",
    /** Часовой пояс для форматирования дат в логах и документах. */
    TIMEZONE: SpreadsheetApp.getActiveSpreadsheet()
      ? SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone()
      : "Europe/Moscow",
    /** Уровень детализации логов: DEBUG | INFO | WARN | ERROR | NONE */
    LOG_LEVEL: "DEBUG",
    LOG_LEVELS: { DEBUG: 0, INFO: 1, WARN: 2, ERROR: 3, NONE: 4 },
    get CURRENT_LOG_LEVEL() {
      const levelKey = (this.LOG_LEVEL || "ERROR").toUpperCase();
      const level = Object.prototype.hasOwnProperty.call(this.LOG_LEVELS, levelKey)
        ? this.LOG_LEVELS[levelKey]
        : 3;
      return level;
    },
    /** Время ожидания блокировки (в миллисекундах) для защиты от одновременных запусков. */
    LOCK_TIMEOUT_MS: 30000,
    /** Время жизни кэша для номеров строк (в секундах). 3600 = 1 час. */
    CACHE_EXPIRATION_SECONDS: 3600,
    /** Префикс для ключей в кэше для избежания конфликтов. */
    CACHE_KEY_PREFIX: "ECOSYSTEM_CACHE_V2.2_",
    /** Имя главной функции-диспетчера, которая устанавливается на триггер onChange. */
    SYNC_HANDLER_FUNCTION: "handleOnChange",
    /** Базовые листы для автосоздания строк при добавлении нового артикула. */
    BASE_SHEETS_FOR_CREATION: ["Сертификация", "Этикетки", "Заказ", "Динамика цены", "Расчет цены", "ABC-Анализ"],
  };

  // =======================================================================================
  // БЛОК 2: РЕЕСТР ИМЕН ЛИСТОВ
  // =======================================================================================
  const SHEETS = {
    PRIMARY: "Главная",
    CERTIFICATION: "Сертификация",
    LABELS: "Этикетки",
    NEWS: "New sert",
    INVOICE_FULL: "Для таможни",
    ORDER: "Ордер", // Используется только в специфических функциях (кнопки Обработка/Загрузить)
    ORDER_FORM: "Заказ",
    PRICE: "Прайс",
    STOCKS: "-остатки", // Используется только в специфических функциях (кнопки Обработка/Загрузить)
    PRICE_DYNAMICS: "Динамика цены",
    PRICE_CALCULATION: "Расчет цены",
    ABC_ANALYSIS: "ABC-Анализ",
    TZ_BY_STATUS: "ТЗ по статусам",
    ORDER_VERIFICATION: "Сверка заказа",
    FOR_DATABASE: "Для базы",
    LOCAL_LOOKUP: "Вид и код",
    RULES: "Правила синхро",
    LOG: "Журнал синхро",
    LOG_DEBUG: "Журнал логов",
    EXTERNAL_DOCS: "Внешние документы",
  };

  // =======================================================================================
  // БЛОК 3: СТРУКТУРА ТАБЛИЦ ("ИСТОЧНИК ПРАВДЫ")
  // =======================================================================================
  // ВНИМАНИЕ: HEADERS используются как SET имен.
  // Порядок колонок в реальных листах может отличаться. Индексация по HEADERS запрещена.
  const HEADERS = {};
  HEADERS[SHEETS.PRIMARY] = [
    "ID",
    "ID-P",
    "Статус",
    "Примечание/записная книжка",
    "Арт. Рус",
    "Арт. произв.",
    "Код база",
    "Код в 1С",
    "Название  ENG прайс произв",
    "Наименования англ по ДС",
    "Наименования рус по ДС",
    "Объём",
    "Форма выпуска",
    "шт./уп.",
    "Категория товара",
    "Линия",
    "Группа",
    "BAR CODE",
    "ID-G",
    "ID-L",
  ];
  HEADERS[SHEETS.CERTIFICATION] = [
    "ID",
    "Статус",
    "Наименования англ по ДС",
    "Наименования рус по ДС",
    "Наименование ДС",
    "Объём",
    "Наименование для инвойса",
    "Объём англ.",
    "Наименование для инвойса Англ",
    "Примечание/записная книжка",
    "Арт. Рус",
    "Арт. произв.",
    "Код база",
    "Код в 1С",
    "BAR CODE",
    "код категории",
    "категория",
    "Вид продукции",
    "Код ТН ВЭД",
    "ДС",
    "ДС до",
    "номер ДС",
    "Протокол ДС",
    "Протокол доп",
    "Скан спирт",
    "Спирт До",
    "INCI",
    "COA",
    "Этикетка",
  ];
  HEADERS[SHEETS.LABELS] = [
    "ID",
    "Статус",
    "Арт. Рус",
    "Категория товара",
    "Объём",
    "Наименования рус по ДС",
    "Наименования англ по ДС",
    "Назначение Этикетка",
    "Применение Этикетка",
    "Активные ингредиенты без %",
    "описание БУКЛЕТ",
    "применение БУКЛЕТ",
    "состав БУКЛЕТ",
    "История статусов",
  ];
  HEADERS[SHEETS.NEWS] = [
    "ID",
    "Статус",
    "Арт. Рус",
    "Арт. произв.",
    "Наименование ДС",
    "Объём",
    "использовать для дс",
    "Примечание/записная книжка",
    "Код ТН ВЭД",
    "Вид продукции",
    "INCI",
    "COA",
    "Этикетка",
    "Вид документа",
    "Протокол исп.",
    "доп протокол/Дата протокол",
    "ГТД/№ инвойс",
    "№гтд/дата инвойса",
    "Договор (УЛ)",
    "Заявка ДС",
    "Макет",
    "Скан ДС",
    "ДС до",
    "номер ДС",
    "Количество",
    "pH",
    "Норма pH",
  ];
  HEADERS[SHEETS.INVOICE_FULL] = [
    "ID",
    "Арт. Рус",
    "Арт. произв.",
    "Наименование",
    "Наименование Англ",
    "кол-во",
    "Цена ед.",
    "Сумма",
    "ДС",
    "ДС до",
    "номер ДС",
    "Спирт",
    "Спирт до",
    "статус",
    "Примечание",
    "Назначение Этикетка",
    "Применение Этикетка",
    "Активные ингредиенты без %",
  ];
  HEADERS[SHEETS.ORDER] = [
    "Арт. произв.",
    "Название ENG",
    "Кол-во поставка",
    "Цена ед.",
    "Скидка, %",
    "Итого цена",
    "Годен до",
  ];
  HEADERS[SHEETS.ORDER_FORM] = [
    "ID",
    "ID-P",
    "Статус",
    "Арт. Рус",
    "Арт. произв.",
    "Название  ENG прайс произв",
    "Наименования рус по ДС",
    "Объём",
    "шт./уп.",
    "ПРОДАЖИ",
    "Остаток",
    "товар в  ПУТИ",
    "Остаток 1",
    "СГ 1",
    "Остаток 2",
    "СГ 2",
    "Остаток3",
    "СГ 3",
    "СПИСАНО",
    "РЕЗЕРВ",
    "Среднемесячные продажи, шт",
    "Потребность на 6 месяцев, шт",
    "На сколько мес запас",
    "Необходимо заказать, шт",
    "Предварительный заказ",
    "заказ в упаковках",
    "ЗАКАЗ",
    "EXW  ALFASPA  текущая, €",
    "Предварительная сумма заказа 0,00 €",
    "АКЦИИ",
    "условие#",
    "коментарий#",
    "Срок#",
    "Набор",
    "условие",
    "коментарий",
    "Срок",
    "Добавить в прайс",
    "Категория товара",
    "Группа линии",
    "Линия Прайс",
    "Примечание",
    "Продублировать",
  ];
  HEADERS[SHEETS.PRICE] = [
    "ID",
    "ID-L",
    "Статус",
    "Категория товара",
    "Код база",
    "Арт. Рус",
    "Наименование",
    "Примечание",
    "Объём",
    "Для",
    "Картинка",
    "Картинка 2",
    "Цена",
    "RRP",
    "Е-comerc",
    "Группа линии",
    "Линия Прайс",
    "Продублировать",
    "Линия Рус",
  ];
  HEADERS[SHEETS.STOCKS] = [
    "Код товара",
    "Выбранная подгруппа",
    "Артикул",
    "Название товара",
    "Продано за период",
    "Списано за период",
    "Остаток на текущий день",
    "Из них в резерве",
    "Всего",
    "Срок годности",
    "Цена",
  ];
  HEADERS[SHEETS.LOCAL_LOOKUP] = ["Код", "Вид"]; /// НАДО СДЕЛАТЬ СПРАВОЧНИК
  HEADERS[SHEETS.PRICE_DYNAMICS] = [
    "ID",
    "ID-P",
    "Статус",
    "Арт. Рус",
    "Арт. произв.",
    "Код база",
    "Код в 1С",
    "Название  ENG прайс произв",
    "Наименования англ по ДС",
    "Наименования рус по ДС",
    "Объём",
    "ЦЕНА EXW из Б/З",
    "Комментарий",
    "EXW 2025, €",
    "СКИДКА ОТ EXW 2025, %",
    "EXW ALFASPA 2025, €",
    "Закупочная цена 2025, ₽",
    "DDP-МОСКВА 2025, ₽",
    "Прирост EXW, €",
    "Прирост DDP-МОСКВА, ₽",
  ];
  HEADERS[SHEETS.PRICE_CALCULATION] = [
    "ID",
    "ID-P",
    "Статус",
    "Арт. Рус",
    "Арт. произв.",
    "Код база",
    "Код в 1С",
    "Название  ENG прайс произв",
    "Наименования англ по ДС",
    "Наименования рус по ДС",
    "Объём",
    "ЦЕНА EXW из Б/З",
    "Комментарий",
    "EXW  предыдущая, €",
    "EXW  текущая, €",
    "EXW  ALFASPA  текущая, €",
    "Закупочная цена, ₽",
    "DDP  -МОСКВА,  ₽",
  ];
  HEADERS[SHEETS.ABC_ANALYSIS] = [
    "ID",
    "ID-P",
    "Статус",
    "Арт. Рус",
    "Арт. произв.",
    "Код база",
    "Код в 1С",
    "Название  ENG прайс произв",
  ];
  HEADERS[SHEETS.TZ_BY_STATUS] = []; // Структура будет определена по мере разработки
  HEADERS[SHEETS.ORDER_VERIFICATION] = []; // Структура будет определена по мере разработки
  HEADERS[SHEETS.FOR_DATABASE] = []; // Структура будет определена по мере разработки
  HEADERS[SHEETS.RULES] = [];
  HEADERS[SHEETS.LOG] = [];
  HEADERS[SHEETS.EXTERNAL_DOCS] = [];
  // =======================================================================================
  // БЛОК 3.1: МАППИНГ ЛОГИЧЕСКИХ КОЛОНОК → ТЕКСТЫ ЗАГОЛОВКОВ
  // Используют модули, которым нужны стабильные имена колонок из конфигурации.
  // =======================================================================================
  const COLUMNS = {
    INVOICE_FULL: {
      ID: "ID",
      RUS_ART: "Арт. Рус",
      PRODUCER_ART: "Арт. произв.",
      PRODUCT_NAME: "Наименование",
      PRODUCT_NAME_EN: "Наименование Англ",
      QTY: "кол-во",
      PRICE: "Цена ед.",
      TOTAL: "Сумма",
      DS_LINK: "ДС",
      DS_UNTIL: "ДС до",
      DS_NUMBER: "номер ДС",
      SPIRIT_LINK: "Спирт",
      SPIRIT_UNTIL: "Спирт до",
      STATUS: "статус",
      NOTE: "Примечание",
      PURPOSE: "Назначение Этикетка",
      APPLICATION: "Применение Этикетка",
      ACTIVE_INGREDIENTS_NO_PERCENT: "Активные ингредиенты без %",
    },
    CERTIFICATION: {
      ID: "ID",
      STATUS: "Статус",
      RUS_NAME_BY_DS: "Наименования рус по ДС",
      ENG_NAME_BY_DS: "Наименования англ по ДС",
      DS_NAME: "Наименование ДС",
      INVOICE_NAME: "Наименование для инвойса",
      INVOICE_NAME_EN: "Наименование для инвойса Англ",
      RUS_ART: "Арт. Рус",
      PRODUCER_ART: "Арт. произв.",
      DS_SCAN: "ДС",
      DS_UNTIL: "ДС до",
      DS_NUMBER: "номер ДС",
      SPIRIT_SCAN: "Скан спирт",
      SPIRIT_UNTIL: "Спирт До",
    },
    LABELS: {
      RUS_ART: "Арт. Рус",
      PURPOSE: "Назначение Этикетка",
      APPLICATION: "Применение Этикетка",
      ACTIVE_INGREDIENTS_NO_PERCENT: "Активные ингредиенты без %",
    },
  };
  // =======================================================================================
  // БЛОК 3.5: ОБЩИЕ КОНСТАНТЫ БИЗНЕС-ЛОГИКИ
  // =======================================================================================
  const VALUES = {
    NOT_FOUND_TEXT: "!!! НЕ НАЙДЕНО !!!",
    DEFAULT_NUMBER_FORMAT: "#,##0.00",
    STATUSES: { NEW: "New в работу" },
    DOC_TYPES: { DS_OPTIONS: ["353пп", "ДС", "СГР"] },
    PROTOCOLS_353PP_FILTER: "353пп",
    SPIRITS: { MAX_GROUP_SIZE: 5 },
    STATIC_DOCS_IDS: [
      "1uL8y5in7ARckjoP-4swt4Cb21wSESdcu",
      "149u8lnDWSWGHig5GCN6kNLcPD7PW7lxJ",
    ],
    MONTH_NAMES: [
      "январь",
      "февраль",
      "март",
      "апрель",
      "май",
      "июнь",
      "июль",
      "август",
      "сентябрь",
      "октябрь",
      "ноябрь",
      "декабрь",
    ],
  };

  // =======================================================================================
  // БЛОК 3.6: РЕЕСТР МЕНЮ
  // ---------------------------------------------------------------------------------------
  // ВАЖНО: Все управление меню централизовано в этом файле (01Config.js).
  // Порядок групп меню (слева направо):
  //   1. 🧾 Заказ (динамически формируется для каждого проекта)
  //   2. 📊 Стадии по заказ (включает кнопки сортировки и стадии заказа)
  //   3. 📦 Выгрузка
  //   4. 🚚 Поставка
  //   5. 🔬 Сертификация
  //   6. ⚙️ Синхронизация
  // =======================================================================================
  const MENU_REGISTRY_BLUEPRINT = [
    // Группа "🧾 Заказ" формируется динамически из PRIMARY_DATA_MENU_ACTIONS
    // и добавляется первой через _buildPrimaryDataMenuGroup()
    {
      title: "⚙️ ЭКОСИСТЕМА",
      items: [
        {
          submenu: "🔍 Проверки систем",
          items: [
            { label: "📊 Показать статус всех сервисов", fn: "showAllServicesStatus_proxy" },
          ],
        },
        {
          submenu: "🤖 Агент",
          items: [
            { label: "Проверить сервис агента", fn: "menuCheckService" },
            { label: "⚙️ Настройки Gemini", fn: "setupGeminiComplete" },
            { label: "📋 Показать текущие настройки", fn: "showGeminiSettings" },
          ],
        },
        {
          submenu: "📊 Логи",
          items: [
            { label: "📈 Дашборд логов", fn: "openLogDashboard_proxy" },
            { separator: true },
            { label: "📑 Показать журнал синхро", fn: "showSyncJournal_proxy" },
            { label: "📑 Показать журнал логов", fn: "showLogJournal_proxy" },
            { separator: true },
            { label: "📦 Архивировать логи", fn: "manualArchiveLogs_proxy" },
            { label: "📊 Статус архива", fn: "showArchiveStatus_proxy" },
            { separator: true },
            { label: "🔄 Пересоздать журнал синхро", fn: "recreateLogSheet" },
            { label: "🔄 Пересоздать журнал логов", fn: "recreateDebugLogSheet" },
            { label: "🧹 Очистить журнал (быстро)", fn: "quickCleanLogSheet" },
          ],
        },
        { separator: true },
        { label: "Добавить артикул", fn: "addArticleManually" },
        { label: "Удалить выбранные строки (с синхронизацией)", fn: "deleteSelectedRowsWithSync" },
        { label: "Синхронизировать выбранную строку", fn: "syncSelectedRow" },
        { label: "Проверить и синхронизировать ВСЮ таблицу", fn: "runFullSync" },
      ],
    },
    {
      title: "📦 Выгрузка",
      items: [
        { label: "Выгрузить Акции", fn: "exportPromotions" },
        { label: "Выгрузить Наборы", fn: "exportSets" },
      ],
    },
    {
      title: "🚚 Поставка",
      items: [
        { label: "Форматировать лист 'Ордер'", fn: "formatOrderSheet" },
        { separator: true },
        { label: "1. Создать лист 'Для инвойса'", fn: "createFullInvoice" },
        { label: "2. Собрать документы", fn: "collectAndCopyDocuments_proxy" },
      ],
    },
    {
      title: "🔬 Сертификация",
      items: [
        { label: "Лист новинки", fn: "createNewsSheetFromCertification" },
        { separator: true },
        {
          label: "Создать заявку протоколы (353пп)",
          fn: "generateProtocols_353pp",
        },
        { label: "Создать заявку ДС (353пп)", fn: "generateDsLayouts_353pp" },
        {
          label: "Собрать документы для заявки (353пп)",
          fn: "structureDocuments_353pp",
        },
        { separator: true },
        { label: "Посчитать спирты", fn: "calculateAndAssignSpiritNumbers" },
        { label: "Создать Макеты спирты", fn: "generateSpiritProtocols" },
        { separator: true },
        {
          label: "Пересчитать каскады (Сертификация)",
          fn: "runManualCascadeOnCertification",
        },
      ],
    },
    {
      title: "⚙️ СИНХРОНИЗАЦИЯ",
      items: [
        { label: "Настроить правила", fn: "showSyncRulesManagerDialog" },
        { label: "🧹 Очистить уведомления", fn: "clearAllToasts" },
        { label: "🚀 Миграция правил со старого листа", fn: "migrateLegacyRules" },
        { label: "🗑️ Удалить лист 'Правила синхро'", fn: "deleteLegacyRulesSheet" },
        { separator: true },
        {
          submenu: "Операции с артикулами",
          items: [
            { label: "Добавить артикул", fn: "addArticleManually" },
            { label: "Удалить артикул", fn: "deleteSelectedRowsWithSync" },
            { label: "Синхронизировать строку", fn: "syncSelectedRow" },
            { label: "Синхронизировать ВСЮ таблицу", fn: "runFullSync" },
          ],
        },
        { separator: true },
        { label: "Установить/Переустановить триггеры", fn: "setupTriggers" },
      ],
    },
    {
      title: "🤖 АГЕНТ",
      items: [
        { label: "🎯 Smart Match для строки", fn: "menuSmartMatch" },
        { label: "🤖 Анализировать выбранную строку", fn: "menuAnalyzeSelected" },
        { label: "📊 Анализировать пустые строки", fn: "menuAnalyzeEmpty" },
        { separator: true },
        { label: "📦 Показать категории ТН ВЭД", fn: "menuShowCategories" },
      ],
    },
  ];

  const PRIMARY_DATA_MENU_ORDER = ["MAIN", "TESTER", "SAMPLES", "PROBES", "STOCKS", "NEW_PRICE_YEAR"];
  const ORDER_STAGES_MENU_ORDER = ["SORT_MANUFACTURER", "SORT_PRICE", "STAGE_ALL", "STAGE_ORDER", "STAGE_PROMOTIONS", "STAGE_SET", "STAGE_PRICE"];

  const PRIMARY_DATA_MENU_ACTIONS = {
    MAIN: { fnByProject: { SK: "processSkPriceSheet", SS: "processSsPriceSheet", MT: "processMtMainPrice" } },
    TESTER: { fnByProject: { MT: "processMtTesterPrice" } },
    SAMPLES: { fnByProject: { MT: "processMtSamplesPrice" } },
    PROBES: { fnByProject: { SK: "processSkPriceProbes" } },
    STOCKS: { fnByProject: { SK: "loadSkStockData", SS: "loadSsStockData", MT: "loadMtStockData" } },
    NEW_PRICE_YEAR: { fn: "addNewYearColumnsToPriceDynamics" },
    SORT_MANUFACTURER: { fnByProject: { SK: "sortSkOrderByManufacturer", SS: "sortSsOrderByManufacturer", MT: "sortMtOrderByManufacturer" } },
    SORT_PRICE: { fnByProject: { SK: "sortSkOrderByPrice", SS: "sortSsOrderByPrice", MT: "sortMtOrderByPrice" } },
    STAGE_ALL: { fn: "showAllOrderData" },
    STAGE_ORDER: { fn: "showOrderStage" },
    STAGE_PROMOTIONS: { fn: "showPromotionsStage" },
    STAGE_SET: { fn: "showSetStage" },
    STAGE_PRICE: { fn: "showPriceStage" },
  };

  // =======================================================================================
  // БЛОК 3.8: ОБЩИЕ ПАТТЕРНЫ ИМЁН ФАЙЛОВ/ПАПОК В DRIVE
  // =======================================================================================
  const DRIVE_STRUCTURE_PATTERNS = {
    PROTOCOLS_353PP: {
      FOLDER_NAME_SUFFIX: "353пп",
      DOCS_SUBFOLDER_NAME: "Макеты",
      PDFS_SUBFOLDER_NAME: "Протоколы",
    },
    DS_LAYOUTS: { TARGET_SUBFOLDER_NAME: "Заявка" },
    DOC_STRUCTURING: { ROOT_FOLDER_NAME: "заявка 353пп" },
    INVOICE_DESCRIPTION: { FILE_NAME_SUFFIX: "описание" },
  };

  // =======================================================================================
  // БЛОК 3.9: ПРАВИЛА КАСКАДНЫХ ОБНОВЛЕНИЙ
  // =======================================================================================
  const CASCADE_RULES = {
    CERTIFICATION: {
      SOURCE_HEADERS: [
        "Наименования рус по ДС",
        "Наименования англ по ДС",
        "Объём",
        "Код ТН ВЭД",
      ],
      VOLUME_EN_REPLACEMENTS: [
        { from: "мл", to: "ml" },
        { from: "гр", to: "g" },
        { from: "Тестер", to: "Tester" },
        { from: "шт. х", to: "*" },
      ],
    },
  };

  // =======================================================================================
  // БЛОК 4: УНИКАЛЬНЫЕ НАСТРОЙКИ ДЛЯ ПРОЕКТОВ (БРЕНДОВ)
  // =======================================================================================
  const PROJECTS = {
    // ============================ ПРОЕКТ "SS" ============================
    SS: {
      /** Префикс для артикулов, папок и файлов проекта. */
      PREFIX: "SS-",
      /** Настройки, связанные с Google Drive для проекта SS. */
      DRIVE: {
        /** 🔬 Сертификация → Посчитать спирты / Создать Макеты спирты */
        SPIRITS: {
          /** 📁 Папка-родитель, куда будут сохраняться папки с протоколами спиртов. */
          PARENT_FOLDER_ID: "1TOcKN8DxZr1a3Vr5jwWGwUxU0D83pldI",
          /** 📝 Шаблон Google Doc, на основе которого создаются протоколы спиртов. */
          TEMPLATE_DOC_ID: "1cOSsXdgEkkB9bSvb_S9h0w6IPwrZuufyPV97DVaepDM",
        },
        /** 🔬 Сертификация → Создать заявку протоколы (353пп) / Собрать документы (353пп) */
        PROTOCOLS_353PP: {
          /** 📁 Корневая папка, внутри которой будет создана папка проекта для заявок 353пп. */
          PARENT_FOLDER_ID: "1xajuDu91Epgh5iz8tG1F0xvoTAkgE7Du",
          /** 📝 Шаблон Google Doc для генерации протоколов 353пп. */
          TEMPLATE_DOC_ID: "1O7TRu7iF47KWNmAaqtcvCEfpKHa7IswWLFPcvFTnlqc",
        },
        /** 🔬 Сертификация → Создать заявку ДС (353пп) */
        DS_LAYOUTS: {
          /** 📝 Шаблон Google Doc для создания макетов Заявки ДС. */
          TEMPLATE_DOC_ID: "1CTP6fLTZEmQf-PYHUio417MsFg_0-tLAUvNh845enXQ",
        },
        /** 📄 Инвойсы → Собрать документы */
        INVOICES: {
          /** 🗂️ Папка-сборник, куда будут копироваться документы для инвойсов по месяцам. */
          COLLECTION_PARENT_FOLDER_ID: "1DjrtDKLUMCypFqlC38kjgOESupTjjUpK",
        },
      },
      /** Уникальные текстовые значения для подстановки в документы проекта SS. */
      VALUES: { LABORATORY_NAME: "COSMETIC RESEARCH GROUP" },
      /** Конфигурация для работы с прайс-листами и остатками проекта SS. */
      PRIMARY_DATA: {
        SOURCE: {
          DOC_ID: "1J8Yzfz9621gqJkPh5ZKBa0v34nv3v9_7OL4JIROlHj0",
        },
        SHEETS: {
          SOURCE: {
            MAIN: "-Б/З поставщик",
            STOCKS: "-остатки",
          },
          TARGET: {
            MAIN: "-Б/З поставщик",
            STOCKS: "-остатки",
          },
        },
        OUTPUT_HEADERS: [
          "ID-P",
          "Арт. произв.",
          "Название ENG прайс произв",
          "Объём",
          "Форма выпуска",
          "BAR CODE",
          "шт./уп.",
          "Цена",
          "Группа",
        ],
        MENU: {
          TITLE: "🧾 Заказ",
          ITEMS: {
            MAIN: "Обработка Б/З поставщик",
            STOCKS: "Загрузить остатки",
            NEW_PRICE_YEAR: "New год для динамика",
          },
        },
        ORDER_STAGES_MENU: {
          TITLE: "📊 Стадии по заказ",
          ITEMS: {
            SORT_MANUFACTURER: "Сортировать по производителю",
            SORT_PRICE: "Сортировать по прайсу",
            STAGE_ALL: "1. Все данные",
            STAGE_ORDER: "2. Заказ",
            STAGE_PROMOTIONS: "3. Акции",
            STAGE_SET: "4. Набор",
            STAGE_PRICE: "5. Прайс",
          },
        },
      },
      /** "Точка расширения" для уникальной логики проекта SS. */
      OVERRIDES: {},
    },

    // ============================ ПРОЕКТ "SK" ============================
    SK: {
      /** Префикс для артикулов, папок и файлов проекта. */
      PREFIX: "SK-",
      /** Настройки, связанные с Google Drive для проекта SK. */
      DRIVE: {
        /** 🔬 Сертификация → Посчитать спирты / Создать Макеты спирты */
        SPIRITS: {
          /** 📁 Папка-родитель для протоколов спиртов. */
          PARENT_FOLDER_ID: "1NSJpEkTlEJ7JGJwK4UmDpefsCFAyuWOu",
          /** 📝 Шаблон Google Doc для протоколов спиртов. */
          TEMPLATE_DOC_ID: "11ZCgJqovASwhMgQXX1PwkYyIcGFIJ-5UAmA3LOqYbhY",
        },
        /** 🔬 Сертификация → Создать заявку протоколы (353пп) / Собрать документы (353пп) */
        PROTOCOLS_353PP: {
          /** 📁 Корневая папка для заявок 353пп. */
          PARENT_FOLDER_ID: "1xajuDu91Epgh5iz8tG1F0xvoTAkgE7Du",
          /** 📝 Шаблон Google Doc для протоколов 353пп. */
          TEMPLATE_DOC_ID: "159qAGqM_E6uY0vVE_HMbg8ByIv6zeK4kFCxaeLeUqeI",
        },
        /** 🔬 Сертификация → Создать заявку ДС (353пп) */
        DS_LAYOUTS: {
          /** 📝 Шаблон Google Doc для макетов Заявки ДС. */
          TEMPLATE_DOC_ID: "1ugPWwjsMqt2DtUBJziafxRAxsYHtLJRhdDMs8ZVaf7U",
        },
        /** 📄 Инвойсы → Собрать документы */
        INVOICES: {
          /** 🗂️ Папка-сборник по месяцам для инвойсов. */
          COLLECTION_PARENT_FOLDER_ID: "1DjrtDKLUMCypFqlC38kjgOESupTjjUpK",
        },
      },
      /** Уникальные текстовые значения для подстановки в документы проекта SK. */
      VALUES: { LABORATORY_NAME: "CARMADO S.L." },
      PRIMARY_DATA: {
        SOURCE: {
          DOC_ID: "1zSu0PzKKa5wvwMZCicwLN8N7Rwhs8XlJVrTrt2LMzQs",
        },
        SHEETS: {
          SOURCE: {
            MAIN: "-Б/З поставщик",
            PROBES: "-пробники",
            STOCKS: "-остатки",
          },
          TARGET: {
            MAIN: "-Б/З поставщик",
            PROBES: "-пробники",
            STOCKS: "-остатки",
          },
        },
        OUTPUT_HEADERS: [
          "ID-P",
          "Арт. произв.",
          "Название ENG прайс произв",
          "Линия",
          "кол-во",
          "шт./уп.",
          "цена",
          "всего",
          "RRP",
          "Группа",
        ],
        GROUP_HEADER_COLOR: "#5b0058",
        PROBES_IDL_HEADER: "ID-G",
        MENU: {
          TITLE: "🧾 Заказ",
          ITEMS: {
            MAIN: "Обработка Б/З поставщик",
            PROBES: "Обработка пробники",
            STOCKS: "Загрузить остатки",
            NEW_PRICE_YEAR: "New год для динамика",
          },
        },
        ORDER_STAGES_MENU: {
          TITLE: "📊 Стадии по заказ",
          ITEMS: {
            SORT_MANUFACTURER: "Сортировать по производителю",
            SORT_PRICE: "Сортировать по прайсу",
            STAGE_ALL: "1. Все данные",
            STAGE_ORDER: "2. Заказ",
            STAGE_PROMOTIONS: "3. Акции",
            STAGE_SET: "4. Набор",
            STAGE_PRICE: "5. Прайс",
          },
        },
      },
      /** "Точка расширения" для уникальной логики проекта SK. */
      OVERRIDES: {},
    },

    // ============================ ПРОЕКТ "MT" ============================
    MT: {
      /** Префикс для артикулов, папок и файлов проекта. */
      PREFIX: "MT-",
      /** Настройки, связанные с Google Drive для проекта MT. */
      DRIVE: {
        /** 🔬 Сертификация → Посчитать спирты / Создать Макеты спирты */
        SPIRITS: {
          /** 📁 Папка-родитель для протоколов спиртов. */
          PARENT_FOLDER_ID: "1fsJeFfAgWPADM8Vf4oiKfKQDZAMvQDBP",
          /** 📝 Шаблон Google Doc для протоколов спиртов. */
          TEMPLATE_DOC_ID: "1nP8gndWnssr1EUXUwYWCajtW70p0wKx-s4Qkwe8Ml_8",
        },
        /** 🔬 Сертификация → Создать заявку протоколы (353пп) / Собрать документы (353пп) */
        PROTOCOLS_353PP: {
          /** 📁 Корневая папка для заявок 353пп. */
          PARENT_FOLDER_ID: "1xajuDu91Epgh5iz8tG1F0xvoTAkgE7Du",
          /** 📝 Шаблон Google Doc для протоколов 353пп. */
          TEMPLATE_DOC_ID: "1trf--Ump09D1IXhUz-c6xr5nJwepN3Q_5l2Cbl6e_78",
        },
        /** 🔬 Сертификация → Создать заявку ДС (353пп) */
        DS_LAYOUTS: {
          /** 📝 Шаблон Google Doc для макетов Заявки ДС. */
          TEMPLATE_DOC_ID: "1mfHSLdNeBF8jl54G1Gf-9k2nsQKL-7MKokdPAbtmL0c",
        },
        /** 📄 Инвойсы → Собрать документы */
        INVOICES: {
          /** 🗂️ Папка-сборник по месяцам для инвойсов. */
          COLLECTION_PARENT_FOLDER_ID: "1DjrtDKLUMCypFqlC38kjgOESupTjjUpK",
        },
      },
      /** Уникальные текстовые значения для подстановки в документы проекта MT. */
      VALUES: { LABORATORY_NAME: "COSMETICA COSBAR S.L." },
      /** Конфигурация для работы с прайс-листами и остатками проекта MT. */
      PRIMARY_DATA: {
        SOURCE: {
          DOC_ID: "1BW8Gk5_X2EZVjbnaa2yDm-bPzzlggwQrHepeNCcPCc0",
        },
        SHEETS: {
          SOURCE: {
            MAIN: "-Б/З поставщик",
            TESTER: "-Тестер",
            SAMPLES: "-Пробники",
            STOCKS: "-остатки",
          },
          TARGET: {
            MAIN: "-Б/З поставщик",
            TESTER: "-Тестер",
            SAMPLES: "-Пробники",
            STOCKS: "-остатки",
          },
        },
        OUTPUT_HEADERS: [
          "ID-P",
          "Арт. произв.",
          "Название ENG прайс произв",
          "Объём",
          "BAR CODE",
          "шт./уп.",
          "Цена",
          "Группа",
        ],
        MENU: {
          TITLE: "🧾 Заказ",
          ITEMS: {
            MAIN: "Обработка Б/З поставщик",
            TESTER: "Обработка Тестер",
            SAMPLES: "Обработка Пробники",
            STOCKS: "Загрузить остатки",
            NEW_PRICE_YEAR: "New год для динамика",
          },
        },
        ORDER_STAGES_MENU: {
          TITLE: "📊 Стадии по заказ",
          ITEMS: {
            SORT_MANUFACTURER: "Сортировать по производителю",
            SORT_PRICE: "Сортировать по прайсу",
            STAGE_ALL: "1. Все данные",
            STAGE_ORDER: "2. Заказ",
            STAGE_PROMOTIONS: "3. Акции",
            STAGE_SET: "4. Набор",
            STAGE_PRICE: "5. Прайс",
          },
        },
      },
      /** "Точка расширения" для уникальной логики проекта MT. */
      OVERRIDES: {},
    },

    // ============================ ПРОЕКТ "HD" ============================
    HD: {
      /** Префикс для артикулов, папок и файлов проекта. */
      PREFIX: "HD-",
      /** Настройки, связанные с Google Drive для проекта HD. */
      DRIVE: {
        /** 🔬 Сертификация → Посчитать спирты / Создать Макеты спирты */
        SPIRITS: {
        /** WARNING: Замените на реальный ID папки для спиртов проекта HD */ PARENT_FOLDER_ID:
            "1_ID_ПАПКИ_СПИРТОВ_ДЛЯ_HD",
          /** WARNING: Замените на реальный ID шаблона для спиртов проекта HD */ TEMPLATE_DOC_ID:
            "1_ID_ШАБЛОНА_СПИРТОВ_ДЛЯ_HD",
        },
        /** 🔬 Сертификация → Создать заявку протоколы (353пп) / Собрать документы (353пп) */
        PROTOCOLS_353PP: {
          /** WARNING: Замените на реальный ID папки для 353пп проекта HD */ PARENT_FOLDER_ID:
            "1_ID_ПАПКИ_353ПП_ДЛЯ_HD",
          /** WARNING: Замените на реальный ID шаблона для 353пп проекта HD */ TEMPLATE_DOC_ID:
            "1_ID_ШАБЛОНА_353ПП_ДЛЯ_HD",
        },
        /** 🔬 Сертификация → Создать заявку ДС (353пп) */
        DS_LAYOUTS: {
          /** WARNING: Замените на реальный ID шаблона заявки ДС для проекта HD */ TEMPLATE_DOC_ID:
            "1_ID_ШАБЛОНА_ЗАЯВКИ_ДС_ДЛЯ_HD",
        },
        /** 📄 Инвойсы → Собрать документы */
        INVOICES: {
          /** WARNING: Замените на реальный ID папки-сборника инвойсов для проекта HD */ COLLECTION_PARENT_FOLDER_ID:
            "1_ID_ПАПКИ_ИНВОЙСОВ_ДЛЯ_HD",
        },
      },
      /** Уникальные текстовые значения для подстановки в документы проекта HD. */
      VALUES: {
        /** WARNING: Замените на реальное название лаборатории для проекта HD */ LABORATORY_NAME:
          "НАЗВАНИЕ ЛАБОРАТОРИИ ДЛЯ HD",
      },
      /** "Точка расширения" для уникальной логики проекта HD. */
      OVERRIDES: {},
    },
  };

  // =======================================================================================
  // БЛОК 5: КАРТА ДОКУМЕНТОВ И АВТООПРЕДЕЛЕНИЕ ПРОЕКТА
  // =======================================================================================
  const DOC_TO_PROJECT = {
  '13kB77R67GJOZQ3vsLcwR1nUaRsupR8ZnEaTdDd66CTQ': 'MT',
    "12yIL1CuESZxeUUd-oKK2brtN1FnXE9q95N7SqzNc7vk": "SS",
    "1CpYYLvRYslsyCkuLzL9EbbjsvbNpWCEZcmhKqMoX5zw": "SK",
    "1zSu0PzKKa5wvwMZCicwLN8N7Rwhs8XlJVrTrt2LMzQs": "SK",
    "1fMOjUE7oZV96fCY5j5rPxnhWGJkDqg-GfwPZ8jUVgPw": "MT",
    "1BW8Gk5_X2EZVjbnaa2yDm-bPzzlggwQrHepeNCcPCc0": "MT",
    "1w8Xddj3uSgXE_AmshWIzJoRKDQpZuqDkQFiCMX99DPk": "HD",
  };

  // =======================================================================================
  // БЛОК 6: ФИНАЛЬНАЯ СБОРКА ОБЪЕКТА CONFIG
  // =======================================================================================
  function _detectProjectKey() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      Logger.log(
        'ВНИМАНИЕ: Запуск из редактора. Используется проект по умолчанию "' +
          SETTINGS.DEBUG_PROJECT_KEY +
          '" для отладки.'
      );
      return SETTINGS.DEBUG_PROJECT_KEY;
    }
    const currentDocumentId = ss.getId();
    const projectKey = DOC_TO_PROJECT[currentDocumentId];
    if (!projectKey) {
      const errorMsg = 'CRITICAL: Document ID ' + currentDocumentId + ' not found in DOC_TO_PROJECT. Using fallback "SS".';
      console.error(errorMsg);
      Logger.log(errorMsg);
      // Fallback to SS to prevent crash and allow menu to load
      return "SS"; 
    }
    return projectKey;
  }

  const activeProjectKey = _detectProjectKey();
  const activeProjectConfig = PROJECTS[activeProjectKey];
  if (!activeProjectConfig) {
      const errorMsg = 'CRITICAL: Configuration for project "' + activeProjectKey + '" not found.';
      console.error(errorMsg);
      // This should theoretically not happen if fallback is valid, but just in case:
      throw new Error(errorMsg);
  }

  function _extend(target) {
    for (var i = 1; i < arguments.length; i++) {
      var source = arguments[i];
      if (!source) continue;
      for (var key in source) {
        if (Object.prototype.hasOwnProperty.call(source, key)) {
          target[key] = source[key];
        }
      }
    }
    return target;
  }

  function _clone(value) {
    if (Array.isArray(value)) {
      return value.map(function (item) {
        return _clone(item);
      });
    }
    if (value && typeof value === 'object') {
      var result = {};
      for (var key in value) {
        if (Object.prototype.hasOwnProperty.call(value, key)) {
          result[key] = _clone(value[key]);
        }
      }
      return result;
    }
    return value;
  }

  function _deepFreeze(obj) {
    if (!obj || typeof obj !== 'object' || Object.isFrozen(obj)) {
      return;
    }
    Object.freeze(obj);
    Object.keys(obj).forEach(function (key) {
      var value = obj[key];
      if (value && typeof value === 'object' && !Object.isFrozen(value)) {
        _deepFreeze(value);
      }
    });
  }

  function _buildMenuRegistry(projectKey, projectConfig) {
    var registry = _clone(MENU_REGISTRY_BLUEPRINT);
    if (!Array.isArray(registry)) {
      return [];
    }

    var primaryDataGroup = _buildPrimaryDataMenuGroup(projectKey, projectConfig);
    if (primaryDataGroup) {
      // Добавляем группу "Заказ" в НАЧАЛО меню (первой)
      registry.unshift(primaryDataGroup);
    }

    // Добавляем группу "Стадии по заказ" для проекта SK
    var orderStagesGroup = _buildOrderStagesMenuGroup(projectKey, projectConfig);
    if (orderStagesGroup) {
      // Добавляем после группы "Заказ"
      registry.splice(1, 0, orderStagesGroup);
    }

    return registry;
  }

  function _buildPrimaryDataMenuGroup(projectKey, projectConfig) {
    if (!projectConfig || !projectConfig.PRIMARY_DATA) {
      return null;
    }

    var cfg = projectConfig.PRIMARY_DATA;
    var menuCfg = cfg.MENU || {};
    var title = menuCfg.TITLE || "🧾 Заказ";
    var itemsCfg = menuCfg.ITEMS || {};
    var items = [];

    for (var i = 0; i < PRIMARY_DATA_MENU_ORDER.length; i++) {
      var key = PRIMARY_DATA_MENU_ORDER[i];
      var label = itemsCfg[key];
      if (!label) {
        continue;
      }

      var actionDef = PRIMARY_DATA_MENU_ACTIONS[key];
      if (!actionDef) {
        continue;
      }

      var item = { label: label };

      if (actionDef.fnByProject) {
        if (!actionDef.fnByProject[projectKey]) {
          continue;
        }
        item.fnByProject = _clone(actionDef.fnByProject);
      } else if (actionDef.fn) {
        item.fn = actionDef.fn;
      } else {
        continue;
      }

      if (actionDef.onlyForProjects) {
        item.onlyForProjects = Array.isArray(actionDef.onlyForProjects)
          ? actionDef.onlyForProjects.slice()
          : _clone(actionDef.onlyForProjects);
      }

      items.push(item);
    }

    if (!items.length) {
      return null;
    }

    return { title: title, items: items };
  }

  function _buildOrderStagesMenuGroup(projectKey, projectConfig) {
    // Эта группа меню доступна для всех проектов, у которых есть PRIMARY_DATA
    if (!projectConfig || !projectConfig.PRIMARY_DATA) {
      return null;
    }

    var cfg = projectConfig.PRIMARY_DATA;
    var stagesCfg = cfg.ORDER_STAGES_MENU || {};
    var title = stagesCfg.TITLE || "📊 Стадии по заказ";
    var itemsCfg = stagesCfg.ITEMS || {};
    var items = [];

    // Используем ORDER_STAGES_MENU_ORDER для формирования меню
    // Порядок: сначала кнопки сортировки (без номеров), затем стадии (с номерами)
    for (var i = 0; i < ORDER_STAGES_MENU_ORDER.length; i++) {
      var key = ORDER_STAGES_MENU_ORDER[i];
      var label = itemsCfg[key];
      if (!label) {
        continue;
      }

      var actionDef = PRIMARY_DATA_MENU_ACTIONS[key];
      if (!actionDef) {
        continue;
      }

      var item = { label: label };

      if (actionDef.fnByProject) {
        if (!actionDef.fnByProject[projectKey]) {
          continue;
        }
        item.fnByProject = _clone(actionDef.fnByProject);
      } else if (actionDef.fn) {
        item.fn = actionDef.fn;
      } else {
        continue;
      }

      if (actionDef.onlyForProjects) {
        item.onlyForProjects = Array.isArray(actionDef.onlyForProjects)
          ? actionDef.onlyForProjects.slice()
          : _clone(actionDef.onlyForProjects);
      }

      items.push(item);
    }

    if (!items.length) {
      return null;
    }

    return { title: title, items: items };
  }

  // AUTOMATION config removed - now using standalone library
  const CONFIG = {
    VERSION: "2.2",
    ACTIVE_PROJECT_KEY: activeProjectKey,
    SETTINGS: _extend({}, SETTINGS, { BRAND_PREFIX: activeProjectConfig.PREFIX }),
    SHEETS: _extend({}, SHEETS),
    HEADERS: _extend({}, HEADERS),
    MENU_REGISTRY: _buildMenuRegistry(activeProjectKey, activeProjectConfig),
    DRIVE: {
      SPIRITS: _extend({}, activeProjectConfig.DRIVE.SPIRITS),
      PROTOCOLS_353PP: _extend(
        {},
        DRIVE_STRUCTURE_PATTERNS.PROTOCOLS_353PP,
        activeProjectConfig.DRIVE.PROTOCOLS_353PP
      ),
      DS_LAYOUTS: _extend(
        {},
        DRIVE_STRUCTURE_PATTERNS.DS_LAYOUTS,
        activeProjectConfig.DRIVE.DS_LAYOUTS
      ),
      DOC_STRUCTURING: _extend({}, DRIVE_STRUCTURE_PATTERNS.DOC_STRUCTURING),
      INVOICE_DESCRIPTION: _extend({}, DRIVE_STRUCTURE_PATTERNS.INVOICE_DESCRIPTION),
      INVOICES: _extend({}, activeProjectConfig.DRIVE.INVOICES),
    },
    VALUES: _extend({}, VALUES, activeProjectConfig.VALUES),
    PRIMARY_DATA: activeProjectConfig.PRIMARY_DATA ? _clone(activeProjectConfig.PRIMARY_DATA) : null,
    OVERRIDES: _extend({ onUpdate: 'onUpdateOverride' }, activeProjectConfig.OVERRIDES),
    CASCADE_RULES: _extend({}, CASCADE_RULES, activeProjectConfig.CASCADE_RULES),
    COLUMNS: _extend({}, COLUMNS),
    COLUMNS: _extend({}, COLUMNS),
    getHeaders: function (sheetName) {
      var headers = this.HEADERS[sheetName] || [];
      return headers.slice();
    },
  };

  // "Замораживаем" конфигурацию для защиты от случайных изменений
  Object.freeze(CONFIG);
  Object.freeze(CONFIG.SETTINGS);
  Object.freeze(CONFIG.SHEETS);
  Object.keys(CONFIG.HEADERS).forEach((key) =>
    Object.freeze(CONFIG.HEADERS[key])
  );
  Object.freeze(CONFIG.HEADERS);
  _deepFreeze(CONFIG.MENU_REGISTRY);
  Object.freeze(CONFIG.MENU_REGISTRY);
  Object.freeze(CONFIG.DRIVE);
  Object.keys(CONFIG.DRIVE).forEach((key) => {
    if (typeof CONFIG.DRIVE[key] === "object") Object.freeze(CONFIG.DRIVE[key]);
  });
  if (CONFIG.PRIMARY_DATA) {
    _deepFreeze(CONFIG.PRIMARY_DATA);
  }
  Object.freeze(CONFIG.VALUES);
  Object.freeze(CONFIG.OVERRIDES);
  Object.freeze(CONFIG.CASCADE_RULES);
  Object.keys(CONFIG.CASCADE_RULES).forEach((key) =>
    Object.freeze(CONFIG.CASCADE_RULES[key])
  );
  Object.freeze(CONFIG.COLUMNS);
  Object.keys(CONFIG.COLUMNS).forEach((key) =>
    Object.freeze(CONFIG.COLUMNS[key])
  );
  Object.freeze(CONFIG.COLUMNS);
  Object.keys(CONFIG.COLUMNS).forEach((key) =>
    Object.freeze(CONFIG.COLUMNS[key])
  );

  // =======================================================================================
  // БЛОК 7: КОНСТРУКТОР МЕНЮ
  // ---------------------------------------------------------------------------------------
  // Перенесено из C_ApiMenu.gs (02Menu.js) для централизации управления меню
  // =======================================================================================

  /**
   * Функция onOpen - создает пользовательское меню при открытии документа
   */
  function _onOpen() {
    const registry = CONFIG.MENU_REGISTRY;
    _createMenusFromRegistry(registry);
  }

  function _createMenusFromRegistry(registry) {
    try {
      const ui = SpreadsheetApp.getUi();
      const activeProjectKey = CONFIG.ACTIVE_PROJECT_KEY || null;
      registry.forEach((group) => {
        if (
          group.onlyForProjects &&
          (!activeProjectKey || group.onlyForProjects.indexOf(activeProjectKey) === -1)
        ) {
          return;
        }
        let menu = ui.createMenu(group.title);
        _buildMenuItems(ui, menu, group.items);
        menu.addToUi();
      });
    } catch (e) {
      console.error(`Критическая ошибка при построении меню: ${e.stack}`);
    }
  }

  function _buildMenuItems(ui, menu, items) {
    const activeProjectKey = CONFIG.ACTIVE_PROJECT_KEY || null;
    items.forEach((item) => {
      if (item.separator) {
        menu.addSeparator();
      } else if (item.submenu) {
        let subMenu = ui.createMenu(item.submenu);
        _buildMenuItems(ui, subMenu, item.items);
        menu.addSubMenu(subMenu);
      } else if (item.label && item.fn) {
        if (
          item.onlyForProjects &&
          item.onlyForProjects.indexOf(CONFIG.ACTIVE_PROJECT_KEY || null) === -1
        ) {
          return;
        }
        var targetFn = item.fn;
        const fnName =
          typeof targetFn === "string" && targetFn.endsWith("_proxy")
            ? targetFn
            : `${targetFn}_proxy`;
        menu.addItem(item.label, fnName);
      } else if (item.label && item.fnByProject) {
        const targetFn = item.fnByProject[activeProjectKey];
        if (!targetFn) {
          return;
        }
        const fnName =
          typeof targetFn === "string" && targetFn.endsWith("_proxy")
            ? targetFn
            : `${targetFn}_proxy`;
        menu.addItem(item.label, fnName);
      }
    });
  }

  // =======================================================================================
  // БЛОК 8: СТРУКТУРИРОВАННОЕ ЛОГИРОВАНИЕ С ЭМОДЗИ
  // =======================================================================================

  /**
   * Универсальная функция логирования с эмодзи и записью в "Журнал синхро"
   *
   * @param {string} message - Сообщение для логирования
   * @param {string} level - Уровень: DEBUG | INFO | WARN | ERROR
   * @param {string} emoji - Эмодзи (опционально, автоматически определяется)
   * @param {string} functionName - Имя функции (опционально)
   * @param {string} details - Дополнительные детали (опционально)
   */
  function logWithEmoji(message, level, emoji, functionName, details) {
    level = level || 'INFO';
    emoji = emoji || _getEmojiForMessage(message, level);
    functionName = functionName || '';
    details = details || '';

    // Проверка уровня логирования
    const currentLevel = SETTINGS.CURRENT_LOG_LEVEL;
    const messageLevelValue = SETTINGS.LOG_LEVELS[level] || 1;

    if (messageLevelValue < currentLevel) {
      return; // Пропускаем логи ниже текущего уровня
    }

    // Формат: timestamp, level, emoji + message, function, details
    const timestamp = Utilities.formatDate(
      new Date(),
      SETTINGS.TIMEZONE,
      'yyyy-MM-dd HH:mm:ss'
    );

    const fullMessage = emoji + ' ' + message;

    // 1. Лог в консоль
    Logger.log(`[${timestamp}] ${level}: ${fullMessage}`);

      // 2. Лог в Google Sheets
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) return;

      // ЛОГИКА РАЗДЕЛЕНИЯ ЖУРНАЛОВ:
      // "Журнал синхро" (SHEETS.LOG) - заполняется ТОЛЬКО через Lib.logSynchronization
      // General logs (logWithEmoji) туда больше НЕ пишут.

      // Б) "Журнал логов" (SHEETS.LOG_DEBUG) - Полный технический лог (ВСЁ: DEBUG + INFO + ...)
      // Пишем сюда ВСЕ сообщения.
      if (SHEETS.LOG_DEBUG) {
        const debugSheet = ss.getSheetByName(SHEETS.LOG_DEBUG);
        if (debugSheet) {
          debugSheet.appendRow([timestamp, level, fullMessage, functionName, details]);
        }
      }

    } catch (e) {
      // Если не удается записать в лист, просто игнорируем ошибку
      console.warn('Не удалось записать в лист логов:', e.message);
    }
  }

  /**
   * Автоматическое определение эмодзи на основе сообщения или уровня
   */
  function _getEmojiForMessage(message, level) {
    const lowerMessage = message.toLowerCase();

    // Проверяем ключевые слова в сообщении
    const emojiMap = {
      'запущен': '🚀',
      'начало': '🚀',
      'start': '🚀',
      'обработка': '⏳',
      'processing': '⏳',
      'получение': '📥',
      'загрузка': '📥',
      'download': '📥',
      'успешно': '✅',
      'success': '✅',
      'готово': '✅',
      'done': '✅',
      'запись': '✏️',
      'write': '✏️',
      'сохранение': '✏️',
      'save': '✏️',
      'предупреждение': '⚠️',
      'warning': '⚠️',
      'внимание': '⚠️',
      'ошибка': '❌',
      'error': '❌',
      'fail': '❌',
      'завершен': '🏁',
      'finish': '🏁',
      'complete': '🏁',
      'синхронизация': '🔄',
      'sync': '🔄',
      'выгрузка': '📤',
      'upload': '📤',
      'export': '📤',
      'удаление': '🗑️',
      'delete': '🗑️',
      'создание': '➕',
      'create': '➕',
      'add': '➕',
      'debug': '🐛',
      'отладка': '🐛'
    };

    for (const [keyword, emoji] of Object.entries(emojiMap)) {
      if (lowerMessage.includes(keyword)) {
        return emoji;
      }
    }

    // Fallback по уровню
    const levelEmojis = {
      'DEBUG': '🐛',
      'INFO': 'ℹ️',
      'WARN': '⚠️',
      'ERROR': '❌'
    };

    return levelEmojis[level] || 'ℹ️';
  }

  // Делаем CONFIG доступным глобально
  global.Lib = global.Lib || {};
  global.Lib.CONFIG = CONFIG;
  global.Lib.onOpen = _onOpen;
  global.Lib.logWithEmoji = logWithEmoji; // Экспортируем функцию логирования
  global.CONFIG = CONFIG;
  global.logWithEmoji = logWithEmoji; // Глобальный доступ для удобства
})(this);
