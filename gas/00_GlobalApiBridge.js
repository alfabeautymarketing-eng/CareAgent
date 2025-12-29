/**
 * =======================================================================================
 * 00_GlobalApiBridge.js — ГЛОБАЛЬНЫЙ МОСТ ДЛЯ APPS SCRIPT API
 * ---------------------------------------------------------------------------------------
 * ОПИСАНИЕ:
 *   Apps Script Execution API требует чтобы функции были ГЛОБАЛЬНЫМИ (не в модуле).
 *   Этот файл создаёт глобальные функции-обёртки которые вызывают функции из модуля Lib.
 *
 *   Файл начинается с "00_" чтобы загружаться ПЕРВЫМ и все функции были доступны.
 *
 * УСТАНОВКА:
 *   1. Скопируйте этот файл в /Users/aleksandr/Desktop/MyGoogleScripts/EcosystemLib/
 *   2. Выполните: cd /Users/aleksandr/Desktop/MyGoogleScripts/EcosystemLib && clasp push
 *   3. Обновите развертывание в Apps Script
 * =======================================================================================
 */

/**
 * Simple Trigger: Запускается при открытии документа.
 * Ограничен в правах (не может делать UrlFetchApp).
 * Служит только для построения меню (часто из кэша).
 */
function onOpen(e) {
  // 1. Попытка построить меню (если сервер недоступен, будет взят кэш)
  if (typeof createAgentMenu === 'function') {
    // Simple trigger: запрещаем сетевые вызовы (UrlFetchApp) чтобы избежать ошибок прав.
    createAgentMenu({ allowNetwork: false });
  }
}

/**
 * Installable Trigger: Запускается при открытии документа (нужно установить вручную).
 * Имеет полные права (UrlFetchApp разрешен).
 * Инициализирует логи, упорядочивает листы и обновляет меню.
 */
function handleOnOpen(e) {
  console.log("🚀 Running Installable handleOnOpen...");

  // 0. Гарантируем, что лог-листы первыми в порядке вкладок
  try {
    if (typeof Lib !== 'undefined' && typeof Lib.ensureLogSheetsFirst === 'function') {
      Lib.ensureLogSheetsFirst();
    }
  } catch (err) {
    console.error("Ошибка при переносе лог-листов: " + err);
  }

  // 1. Инициализация сессии логов на сервере (Логи)
  try {
    if (typeof Lib !== 'undefined' && typeof Lib.initSessionLogs === 'function') {
      Lib.initSessionLogs();
      if (typeof Lib.ensureLogSheetsFirst === 'function') {
        Lib.ensureLogSheetsFirst();
      }
    }
    if (typeof Lib !== 'undefined' && typeof Lib.logStep === 'function') {
      Lib.logStep("Startup", "Инициализация логов сессии завершена");
    }
  } catch (err) {
    console.error("Ошибка при инициализации логов: " + err);
    if (typeof Lib !== 'undefined' && typeof Lib.logWarn === 'function') {
      Lib.logWarn("Startup: не удалось инициализировать логи", err);
    }
  }

  // 2. Загрузка меню (требует полного доступа, поэтому делаем после логов)
  if (typeof createAgentMenu === 'function') {
    try {
      if (typeof Lib !== 'undefined' && typeof Lib.logStep === 'function') {
        Lib.logStep("Startup", "Загрузка динамического меню");
      }
      createAgentMenu();
    } catch (err) {
      console.error("Ошибка при загрузке меню: " + err);
      if (typeof Lib !== 'undefined' && typeof Lib.logWarn === 'function') {
        Lib.logWarn("Startup: ошибка загрузки меню", err);
      }
    }
  }

  // 3. Автоматическое упорядочивание листов через Python сервер
  if (typeof reorderSheetsSilent === 'function') {
    try {
      if (typeof Lib !== 'undefined' && typeof Lib.logStep === 'function') {
        Lib.logStep("Startup", "Выстраиваем листы по порядку (сервер)");
      }
      reorderSheetsSilent();
      if (typeof Lib !== 'undefined' && typeof Lib.ensureLogSheetsFirst === 'function') {
        Lib.ensureLogSheetsFirst();
      }
    } catch (err) {
      console.error("Ошибка при упорядочивании листов: " + err);
      if (typeof Lib !== 'undefined' && typeof Lib.logWarn === 'function') {
        Lib.logWarn("Startup: ошибка упорядочивания листов", err);
      }
    }
  }

  // 3.5. Инициализация Gemini API из Script Properties
  if (typeof initGeminiFromStorage === 'function') {
    try {
      if (typeof Lib !== 'undefined' && typeof Lib.logStep === 'function') {
        Lib.logStep("Startup", "Инициализация Gemini API");
      }
      const geminiOk = initGeminiFromStorage();
      if (geminiOk && typeof Lib !== 'undefined' && typeof Lib.logStep === 'function') {
        Lib.logStep("Startup", "Gemini API инициализирован из хранилища");
      }
    } catch (err) {
      console.error("Ошибка при инициализации Gemini: " + err);
    }
  }

  // 4. Обновляем формулы на ключевых листах
  try {
    if (typeof Lib !== 'undefined' && typeof Lib.recalculatePriceDynamicsFormulas === 'function') {
      if (typeof Lib.logStep === 'function') {
        Lib.logStep("Startup", "Обновляем формулы листа \"Динамика цены\"");
      }
      Lib.recalculatePriceDynamicsFormulas();
    }
  } catch (err) {
    console.error("Ошибка при обновлении формул Динамика цены: " + err);
    if (typeof Lib !== 'undefined' && typeof Lib.logWarn === 'function') {
      Lib.logWarn('Startup: ошибка формул "Динамика цены"', err);
    }
  }

  try {
    if (typeof Lib !== 'undefined' && typeof Lib.updatePriceCalculationFormulas === 'function') {
      if (typeof Lib.logStep === 'function') {
        Lib.logStep("Startup", "Обновляем формулы листа \"Расчет цены\"");
      }
      Lib.updatePriceCalculationFormulas(true); // silent
    }
  } catch (err) {
    console.error("Ошибка при обновлении формул Расчет цены: " + err);
    if (typeof Lib !== 'undefined' && typeof Lib.logWarn === 'function') {
      Lib.logWarn('Startup: ошибка формул "Расчет цены"', err);
    }
  }

  // 5. Активация листа "Главная"
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var mainSheet = ss.getSheetByName("Главная");
    if (mainSheet) {
      if (typeof Lib !== 'undefined' && typeof Lib.logStep === 'function') {
        Lib.logStep("Startup", "Переходим на лист \"Главная\"");
      }
      ss.setActiveSheet(mainSheet);
    }
  } catch (err) {
    console.error("Ошибка при активации листа Главная: " + err);
    if (typeof Lib !== 'undefined' && typeof Lib.logWarn === 'function') {
      Lib.logWarn("Startup: не удалось активировать лист Главная", err);
    }
  }
}

/**
 * Simple Trigger: Автоматическая синхронизация при редактировании ячеек.
 * Срабатывает при любом изменении в таблице.
 */
function onEdit(e) {
  // Simple Trigger 'onEdit' cannot call UrlFetchApp (Sync requires it).
  // We rely on 'handleOnEdit' (Installable Trigger) for synchronization.
  // Leaving this empty to prevent "Permission denied" errors in logs.
  /*
  if (typeof Lib !== 'undefined' && typeof Lib.onEdit_internal_ === 'function') {
    Lib.onEdit_internal_(e);
  }
  */
}

/**
 * Installable Trigger: Обработка структурных изменений (добавление/удаление листов).
 * Сбрасывает кэш для корректной работы синхронизации.
 */
function handleOnChange(e) {
  if (typeof Lib !== 'undefined' && typeof Lib.handleOnChange === 'function') {
    Lib.handleOnChange(e);
  }
}

/**
 * Installable Trigger: Обработка редактирования (более надежный чем simple onEdit).
 * Логирует все изменения в лист "Логи".
 */
function handleOnEdit(e) {
  // Логируем событие редактирования
  if (typeof Lib !== 'undefined' && typeof Lib.logEditEvent === 'function') {
    try {
      Lib.logEditEvent(e);
    } catch (logErr) {
      console.error("Ошибка логирования редактирования:", logErr);
    }
  }

  // Выполняем основную логику синхронизации
  if (typeof Lib !== 'undefined' && typeof Lib.onEdit_internal_ === 'function') {
    Lib.onEdit_internal_(e);
  }
}

// ============ ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ЛОГИРОВАНИЯ ============

/**
 * Универсальная функция записи в лист "Логи"
 * @param {string} category - Категория (FUNCTION, MENU, SYSTEM и т.д.)
 * @param {string} action - Действие
 * @param {string} details - Детали
 * @param {string} status - Статус (✅ OK, ❌ ОШИБКА и т.д.)
 */
function _writeToLogSheet_(category, action, details, status) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return;

    var logSheetName = "Логи";
    var sh = ss.getSheetByName(logSheetName);

    if (!sh) {
      // Создаём лист если его нет
      sh = ss.insertSheet(logSheetName);
      var headers = ["🕒 Время", "🏷️ Категория", "💬 Действие", "📝 Детали", "🔘 Статус"];
      sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, headers.length).setBackground("#e8eaf6");
    }

    var timestamp = Utilities.formatDate(new Date(), "Europe/Moscow", "dd.MM.yyyy HH:mm:ss");
    sh.appendRow([timestamp, category, action, details || "", status || "✅ OK"]);
  } catch (e) {
    console.error("_writeToLogSheet_ error:", e);
  }
}

/**
 * Универсальная обёртка для логирования вызовов функций из меню.
 * @param {string} functionName - Имя функции
 * @param {Function} fn - Функция для выполнения
 * @param {string} [source="MENU"] - Источник вызова
 * @returns {*} Результат функции
 */
function _loggedCall_(functionName, fn, source) {
  source = source || "FUNCTION";
  var startTime = Date.now();

  // Логируем начало
  _writeToLogSheet_(source, "Вызов: " + functionName, "Запуск функции", "🔄 В ПРОЦЕССЕ");

  try {
    var result = fn();
    var duration = Date.now() - startTime;

    // Логируем успешное завершение
    _writeToLogSheet_(source, "Завершено: " + functionName, "Время: " + duration + "ms", "✅ OK");

    return result;
  } catch (e) {
    var duration = Date.now() - startTime;
    // Логируем ошибку
    _writeToLogSheet_(source, "ОШИБКА: " + functionName, e.message, "❌ ОШИБКА");
    throw e;
  }
}

// ============ СИНХРОНИЗАЦИЯ ============

/**
 * Служебная функция для принудительного закрытия любых висящих toast-уведомлений.
 * Вызовите эту функцию из меню, если toast "застрял".
 */
function clearAllToasts() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    // Закрываем любые висящие toast
    ss.toast('', '', 1);
    SpreadsheetApp.getUi().alert('Toast-уведомления очищены', 'Все активные уведомления были закрыты.', SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Ошибка', 'Не удалось закрыть toast: ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function addArticleManually() {
  return _loggedCall_("addArticleManually", function() {
    if (typeof Lib !== 'undefined' && Lib.addArticleManually) {
      return Lib.addArticleManually();
    }
    throw new Error('Lib.addArticleManually не определена');
  });
}

function deleteSelectedRowsWithSync() {
  return _loggedCall_("deleteSelectedRowsWithSync", function() {
    if (typeof Lib !== 'undefined' && Lib.deleteSelectedRowsWithSync) {
      return Lib.deleteSelectedRowsWithSync();
    }
    throw new Error('Lib.deleteSelectedRowsWithSync не определена');
  });
}

function syncSelectedRow() {
  return _loggedCall_("syncSelectedRow", function() {
    if (typeof Lib !== 'undefined' && Lib.syncSelectedRow) {
      return Lib.syncSelectedRow();
    }
    throw new Error('Lib.syncSelectedRow не определена');
  });
}

function runFullSync() {
  return _loggedCall_("runFullSync", function() {
    if (typeof Lib !== 'undefined' && Lib.runFullSync) {
      return Lib.runFullSync();
    }
    throw new Error('Lib.runFullSync не определена');
  });
}

function setupTriggers() {
  if (typeof Lib !== 'undefined' && Lib.setupTriggers) {
    return Lib.setupTriggers();
  }
  throw new Error('Lib.setupTriggers не определена');
}

function showSyncConfigDialog() {
  if (typeof Lib !== 'undefined' && Lib.showSyncConfigDialog) {
    return Lib.showSyncConfigDialog();
  }
  throw new Error('Lib.showSyncConfigDialog не определена');
}

function showExternalDocManagerDialog() {
  if (typeof Lib !== 'undefined' && Lib.showExternalDocManagerDialog) {
    return Lib.showExternalDocManagerDialog();
  }
  throw new Error('Lib.showExternalDocManagerDialog не определена');
}

// ============ ТЕСТЕР ============

function runAllTests() {
  if (typeof Lib !== 'undefined' && Lib.runAllTests) {
    return Lib.runAllTests();
  }
  throw new Error('Lib.runAllTests не определена');
}

function clearTestResults() {
  if (typeof Lib !== 'undefined' && Lib.clearTestResults) {
    return Lib.clearTestResults();
  }
  throw new Error('Lib.clearTestResults не определена');
}

// ============ ЛОГИ ============

function refreshLogs() {
  if (typeof Lib !== 'undefined' && Lib.refreshLogs) {
    return Lib.refreshLogs();
  }
  throw new Error('Lib.refreshLogs не определена');
}

function quickCleanLogSheet() {
  if (typeof Lib !== 'undefined' && Lib.quickCleanLogSheet) {
    return Lib.quickCleanLogSheet();
  }
  throw new Error('Lib.quickCleanLogSheet не определена');
}

function recreateLogSheet() {
  if (typeof Lib !== 'undefined' && Lib.recreateLogSheet) {
    return Lib.recreateLogSheet();
  }
  throw new Error('Lib.recreateLogSheet не определена');
}

function recreateDebugLogSheet() {
  if (typeof Lib !== 'undefined' && Lib.recreateDebugLogSheet) {
    return Lib.recreateDebugLogSheet();
  }
  throw new Error('Lib.recreateDebugLogSheet не определена');
}

// ============ ОБРАБОТКА ПРАЙСОВ (SK) ============

function processSkPriceSheet() {
  return _loggedCall_("processSkPriceSheet", function() {
    if (typeof Lib !== 'undefined' && Lib.processSkPriceSheet) {
      return Lib.processSkPriceSheet();
    }
    throw new Error('Lib.processSkPriceSheet не определена');
  });
}

function loadSkStockData() {
  return _loggedCall_("loadSkStockData", function() {
    if (typeof Lib !== 'undefined' && Lib.loadSkStockData) {
      return Lib.loadSkStockData();
    }
    throw new Error('Lib.loadSkStockData не определена');
  });
}

// ============ ОБРАБОТКА ПРАЙСОВ (MT) ============

function processMtMainPrice() {
  return _loggedCall_("processMtMainPrice", function() {
    if (typeof Lib !== 'undefined' && Lib.processMtMainPrice) {
      return Lib.processMtMainPrice();
    }
    throw new Error('Lib.processMtMainPrice не определена');
  });
}

function processMtTesterPrice() {
  return _loggedCall_("processMtTesterPrice", function() {
    if (typeof Lib !== 'undefined' && Lib.processMtTesterPrice) {
      return Lib.processMtTesterPrice();
    }
    throw new Error('Lib.processMtTesterPrice не определена');
  });
}

function processMtSamplesPrice() {
  return _loggedCall_("processMtSamplesPrice", function() {
    if (typeof Lib !== 'undefined' && Lib.processMtSamplesPrice) {
      return Lib.processMtSamplesPrice();
    }
    throw new Error('Lib.processMtSamplesPrice не определена');
  });
}

function loadMtStockData() {
  return _loggedCall_("loadMtStockData", function() {
    if (typeof Lib !== 'undefined' && Lib.loadMtStockData) {
      return Lib.loadMtStockData();
    }
    throw new Error('Lib.loadMtStockData не определена');
  });
}

// ============ ОБРАБОТКА ПРАЙСОВ (SS) ============

function processSsPriceSheet() {
  return _loggedCall_("processSsPriceSheet", function() {
    if (typeof Lib !== 'undefined' && Lib.processSsPriceSheet) {
      return Lib.processSsPriceSheet();
    }
    throw new Error('Lib.processSsPriceSheet не определена');
  });
}

function loadSsStockData() {
  return _loggedCall_("loadSsStockData", function() {
    if (typeof Lib !== 'undefined' && Lib.loadSsStockData) {
      return Lib.loadSsStockData();
    }
    throw new Error('Lib.loadSsStockData не определена');
  });
}

// ============ ОБЩИЕ ФУНКЦИИ ЗАКАЗА ============

function addNewYearColumnsToPriceDynamics() {
  if (typeof Lib !== 'undefined' && Lib.addNewYearColumnsToPriceDynamics) {
    return Lib.addNewYearColumnsToPriceDynamics();
  }
  throw new Error('Lib.addNewYearColumnsToPriceDynamics не определена');
}

// ============ СОРТИРОВКА (PYTHON SERVER) ============
// Эти функции перенаправляют вызовы на Python сервер для быстрой сортировки

function sortSkOrderByManufacturer() {
  if (typeof callServerStructureSort === 'function') {
    return callServerStructureSort('byManufacturer');
  }
  throw new Error('callServerStructureSort не определена');
}

function sortSkOrderByPrice() {
  if (typeof callServerStructureSort === 'function') {
    return callServerStructureSort('byPrice');
  }
  throw new Error('callServerStructureSort не определена');
}

function sortMtOrderByManufacturer() {
  if (typeof callServerStructureSort === 'function') {
    return callServerStructureSort('byManufacturer');
  }
  throw new Error('callServerStructureSort не определена');
}

function sortMtOrderByPrice() {
  if (typeof callServerStructureSort === 'function') {
    return callServerStructureSort('byPrice');
  }
  throw new Error('callServerStructureSort не определена');
}

function sortSsOrderByManufacturer() {
  if (typeof callServerStructureSort === 'function') {
    return callServerStructureSort('byManufacturer');
  }
  throw new Error('callServerStructureSort не определена');
}

function sortSsOrderByPrice() {
  if (typeof callServerStructureSort === 'function') {
    return callServerStructureSort('byPrice');
  }
  throw new Error('callServerStructureSort не определена');
}

function showAllOrderData() {
  if (typeof Lib !== 'undefined' && Lib.showAllOrderData) {
    return Lib.showAllOrderData();
  }
  throw new Error('Lib.showAllOrderData не определена');
}

function showOrderStage() {
  if (typeof Lib !== 'undefined' && Lib.showOrderStage) {
    return Lib.showOrderStage();
  }
  throw new Error('Lib.showOrderStage не определена');
}

function showPromotionsStage() {
  if (typeof Lib !== 'undefined' && Lib.showPromotionsStage) {
    return Lib.showPromotionsStage();
  }
  throw new Error('Lib.showPromotionsStage не определена');
}

function showSetStage() {
  if (typeof Lib !== 'undefined' && Lib.showSetStage) {
    return Lib.showSetStage();
  }
  throw new Error('Lib.showSetStage не определена');
}

function showPriceStage() {
  if (typeof Lib !== 'undefined' && Lib.showPriceStage) {
    return Lib.showPriceStage();
  }
  throw new Error('Lib.showPriceStage не определена');
}

// ============ ВЫГРУЗКА ============

function exportPromotions() {
  if (typeof Lib !== 'undefined' && Lib.exportPromotions) {
    return Lib.exportPromotions();
  }
  throw new Error('Lib.exportPromotions не определена');
}

function exportSets() {
  if (typeof Lib !== 'undefined' && Lib.exportSets) {
    return Lib.exportSets();
  }
  throw new Error('Lib.exportSets не определена');
}

// ============ ПОСТАВКА ============

function formatOrderSheet() {
  return _loggedCall_("formatOrderSheet", function() {
    if (typeof Lib !== 'undefined' && Lib.formatOrderSheet) {
      return Lib.formatOrderSheet();
    }
    throw new Error('Lib.formatOrderSheet не определена');
  });
}

function createFullInvoice() {
  return _loggedCall_("createFullInvoice", function() {
    if (typeof Lib !== 'undefined' && Lib.createFullInvoice) {
      return Lib.createFullInvoice();
    }
    throw new Error('Lib.createFullInvoice не определена');
  });
}

function collectAndCopyDocuments() {
  return _loggedCall_("collectAndCopyDocuments", function() {
    if (typeof Lib !== 'undefined' && Lib.collectAndCopyDocuments) {
      return Lib.collectAndCopyDocuments();
    }
    throw new Error('Lib.collectAndCopyDocuments не определена');
  });
}

// ============ СЕРТИФИКАЦИЯ ============

function createNewsSheetFromCertification() {
  return _loggedCall_("createNewsSheetFromCertification", function() {
    if (typeof Lib !== 'undefined' && Lib.createNewsSheetFromCertification) {
      return Lib.createNewsSheetFromCertification();
    }
    throw new Error('Lib.createNewsSheetFromCertification не определена');
  });
}

function generateProtocols_353pp() {
  return _loggedCall_("generateProtocols_353pp", function() {
    if (typeof Lib !== 'undefined' && Lib.generateProtocols_353pp) {
      return Lib.generateProtocols_353pp();
    }
    throw new Error('Lib.generateProtocols_353pp не определена');
  });
}

function generateDsLayouts_353pp() {
  return _loggedCall_("generateDsLayouts_353pp", function() {
    if (typeof Lib !== 'undefined' && Lib.generateDsLayouts_353pp) {
      return Lib.generateDsLayouts_353pp();
    }
    throw new Error('Lib.generateDsLayouts_353pp не определена');
  });
}

function structureDocuments_353pp() {
  return _loggedCall_("structureDocuments_353pp", function() {
    if (typeof Lib !== 'undefined' && Lib.structureDocuments_353pp) {
      return Lib.structureDocuments_353pp();
    }
    throw new Error('Lib.structureDocuments_353pp не определена');
  });
}

function calculateAndAssignSpiritNumbers() {
  return _loggedCall_("calculateAndAssignSpiritNumbers", function() {
    if (typeof Lib !== 'undefined' && Lib.calculateAndAssignSpiritNumbers) {
      return Lib.calculateAndAssignSpiritNumbers();
    }
    throw new Error('Lib.calculateAndAssignSpiritNumbers не определена');
  });
}

function generateSpiritProtocols() {
  return _loggedCall_("generateSpiritProtocols", function() {
    if (typeof Lib !== 'undefined' && Lib.generateSpiritProtocols) {
      return Lib.generateSpiritProtocols();
    }
    throw new Error('Lib.generateSpiritProtocols не определена');
  });
}

function runManualCascadeOnCertification() {
  return _loggedCall_("runManualCascadeOnCertification", function() {
    if (typeof Lib !== 'undefined' && Lib.runManualCascadeOnCertification) {
      return Lib.runManualCascadeOnCertification();
    }
    throw new Error('Lib.runManualCascadeOnCertification не определена');
  });
}

// ============ DRIVE ============

function uploadFilesToFolder() {
  if (typeof Lib !== 'undefined' && Lib.uploadFilesToFolder) {
    return Lib.uploadFilesToFolder();
  }
  throw new Error('Lib.uploadFilesToFolder не определена');
}

function createFolderStructure() {
  if (typeof Lib !== 'undefined' && Lib.createFolderStructure) {
    return Lib.createFolderStructure();
  }
  throw new Error('Lib.createFolderStructure не определена');
}
// ============ УПОРЯДОЧИВАНИЕ ЛИСТОВ (PYTHON SERVER) ============

function reorderAuxiliarySheets() {
  if (typeof reorderSheets === 'function') {
    return reorderSheets();
  }
  throw new Error('reorderSheets не определена');
}

// ============ СЕРВЕРНАЯ ОБРАБОТКА ПРАЙСОВ (PYTHON SERVER) ============
// Эти функции перенаправляют обработку прайс-листов на Python сервер

/**
 * Обработка основного прайса MT через сервер
 */
function serverProcessMtMain() {
  return _loggedCall_("serverProcessMtMain", function() {
    if (typeof callServerProcessPrice === 'function') {
      return callServerProcessPrice('mt', 'main');
    }
    throw new Error('callServerProcessPrice не определена');
  });
}

/**
 * Обработка тестеров MT через сервер
 */
function serverProcessMtTester() {
  return _loggedCall_("serverProcessMtTester", function() {
    if (typeof callServerProcessPrice === 'function') {
      return callServerProcessPrice('mt', 'tester');
    }
    throw new Error('callServerProcessPrice не определена');
  });
}

/**
 * Обработка пробников MT через сервер
 */
function serverProcessMtSamples() {
  return _loggedCall_("serverProcessMtSamples", function() {
    if (typeof callServerProcessPrice === 'function') {
      return callServerProcessPrice('mt', 'samples');
    }
    throw new Error('callServerProcessPrice не определена');
  });
}

/**
 * Обработка прайса SK через сервер
 */
function serverProcessSkMain() {
  return _loggedCall_("serverProcessSkMain", function() {
    if (typeof callServerProcessPrice === 'function') {
      return callServerProcessPrice('sk', 'main');
    }
    throw new Error('callServerProcessPrice не определена');
  });
}

/**
 * Обработка пробников SK через сервер
 */
function serverProcessSkProbes() {
  return _loggedCall_("serverProcessSkProbes", function() {
    if (typeof callServerProcessPrice === 'function') {
      return callServerProcessPrice('sk', 'probes');
    }
    throw new Error('callServerProcessPrice не определена');
  });
}

/**
 * Обработка прайса SS через сервер
 */
function serverProcessSsMain() {
  return _loggedCall_("serverProcessSsMain", function() {
    if (typeof callServerProcessPrice === 'function') {
      return callServerProcessPrice('ss', 'main');
    }
    throw new Error('callServerProcessPrice не определена');
  });
}

/**
 * Preview обработки прайса (dry run)
 * @param {string} project - Код проекта (mt, sk, ss)
 * @param {string} mode - Режим (main, tester, samples, probes)
 */
function serverProcessPricePreview(project, mode) {
  return _loggedCall_("serverProcessPricePreview", function() {
    if (typeof callServerProcessPrice === 'function') {
      return callServerProcessPrice(project, mode, { dryRun: true });
    }
    throw new Error('callServerProcessPrice не определена');
  });
}
