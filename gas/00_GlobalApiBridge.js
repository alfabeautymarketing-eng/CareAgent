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

function serverSmartMatch(productName) {
  return _loggedCall_("serverSmartMatch", function() {
    if (typeof callServerSmartMatch === 'function') {
      return callServerSmartMatch(productName);
    }
    throw new Error('callServerSmartMatch не определена');
  });
}

// ============ КАСКАДНЫЕ ПРАВИЛА (PYTHON SERVER) ============
// Эти функции перенаправляют обработку каскадов на Python сервер

/**
 * Пересчёт каскадов для всего листа Сертификация через сервер
 * Эквивалент GAS runManualCascadeOnCertification()
 */
function serverRecalculateCascades() {
  return _loggedCall_("serverRecalculateCascades", function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();

    const resp = ui.alert(
      "Пересчёт каскадов",
      'Пересчитать все строки на листе "Сертификация"?',
      ui.ButtonSet.YES_NO
    );
    if (resp !== ui.Button.YES) return { status: 'cancelled' };

    ss.toast("Пересчёт каскадов...", "Выполнение", 30);

    try {
      const result = callServerCascade({
        spreadsheet_id: ss.getId(),
        sheet_name: 'Сертификация',
        dry_run: false
      });

      if (result && result.status === 'success') {
        ss.toast(
          'Обработано: ' + (result.processed || 0) + ', изменено: ' + (result.changed || 0),
          '✅ Готово',
          5
        );
        ui.alert('Каскады пересчитаны.');
      } else {
        ui.alert('Ошибка: ' + (result.message || 'Неизвестная ошибка'));
      }

      return result;
    } catch (err) {
      ss.toast('Ошибка пересчёта', 'Ошибка', 3);
      ui.alert('Ошибка: ' + err.message);
      throw err;
    }
  });
}

/**
 * Обработка каскада для одной строки
 * @param {number} row - Номер строки
 * @param {string} column - Изменённая колонка
 * @param {string} [value] - Новое значение (опционально)
 */
function serverProcessCascadeRow(row, column, value) {
  return _loggedCall_("serverProcessCascadeRow", function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const result = callServerCascade({
      spreadsheet_id: ss.getId(),
      sheet_name: 'Сертификация',
      row: row,
      changed_column: column,
      new_value: value,
      dry_run: false
    });

    return result;
  });
}

/**
 * Вызов сервера для обработки каскадов
 * @param {Object} params - Параметры запроса
 */
function callServerCascade(params) {
  const BASE_URL = PropertiesService.getScriptProperties().getProperty('SERVER_URL') || 'http://localhost:8000';
  const endpoint = params.row ? '/api/v1/cascade/process' : '/api/v1/cascade/recalculate-all';

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      spreadsheet_id: params.spreadsheet_id,
      sheet_name: params.sheet_name || 'Сертификация',
      row: params.row || null,
      changed_column: params.changed_column || null,
      new_value: params.new_value || null,
      dry_run: params.dry_run || false
    }),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(BASE_URL + endpoint, options);
    const status = response.getResponseCode();
    const text = response.getContentText();

    if (status >= 200 && status < 300) {
      return JSON.parse(text);
    } else {
      console.error('Cascade server error: ' + status + ' - ' + text);
      return { status: 'error', message: 'Server returned ' + status };
    }
  } catch (err) {
    console.error('Cascade request failed: ' + err.message);
    return { status: 'error', message: err.message };
  }
}

// ============ СТАДИИ ЗАКАЗА (PYTHON SERVER) ============
// Эти функции перенаправляют фильтрацию стадий на Python сервер

/**
 * Показать все данные на листе Заказ (снять все фильтры)
 * Эквивалент GAS showAllOrderData()
 */
function serverShowAllOrderData() {
  return _loggedCall_("serverShowAllOrderData", function() {
    return _callServerOrderFilter('all');
  });
}

/**
 * Показать стадию "Заказ"
 * Эквивалент GAS showOrderStage()
 */
function serverShowOrderStage() {
  return _loggedCall_("serverShowOrderStage", function() {
    return _callServerOrderFilter('order');
  });
}

/**
 * Показать стадию "Акции"
 * Эквивалент GAS showPromotionsStage()
 */
function serverShowPromotionsStage() {
  return _loggedCall_("serverShowPromotionsStage", function() {
    return _callServerOrderFilter('promotions');
  });
}

/**
 * Показать стадию "Набор"
 * Эквивалент GAS showSetStage()
 */
function serverShowSetStage() {
  return _loggedCall_("serverShowSetStage", function() {
    return _callServerOrderFilter('set');
  });
}

/**
 * Показать стадию "Прайс"
 * Эквивалент GAS showPriceStage()
 */
function serverShowPriceStage() {
  return _loggedCall_("serverShowPriceStage", function() {
    return _callServerOrderFilter('price');
  });
}

/**
 * Вызов сервера для фильтрации стадий заказа
 * @param {string} stage - Стадия (all, order, promotions, set, price)
 * @private
 */
function _callServerOrderFilter(stage) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const menuTitle = "Стадии по заказ";

  ss.toast("Применяю фильтр '" + stage + "'...", menuTitle, 30);

  const BASE_URL = PropertiesService.getScriptProperties().getProperty('SERVER_URL') || 'http://localhost:8000';

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      spreadsheet_id: ss.getId(),
      stage: stage,
      sheet_name: 'Заказ',
      dry_run: false
    }),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(BASE_URL + '/api/v1/order/filter', options);
    const status = response.getResponseCode();
    const text = response.getContentText();

    if (status >= 200 && status < 300) {
      const result = JSON.parse(text);

      if (result && result.status === 'success') {
        const stageNames = {
          'all': 'Все данные',
          'order': 'Заказ',
          'promotions': 'Акции',
          'set': 'Набор',
          'price': 'Прайс'
        };
        ss.toast(
          'Фильтр "' + (stageNames[stage] || stage) + '" применён. Скрыто строк: ' + result.hidden_rows,
          '✅ Готово',
          5
        );
      } else {
        ui.alert('Ошибка', result.message || 'Неизвестная ошибка', ui.ButtonSet.OK);
      }

      return result;
    } else {
      console.error('Order filter server error: ' + status + ' - ' + text);
      ss.toast('Ошибка сервера: ' + status, 'Ошибка', 3);
      return { status: 'error', message: 'Server returned ' + status };
    }
  } catch (err) {
    console.error('Order filter request failed: ' + err.message);
    ss.toast('Ошибка: ' + err.message, 'Ошибка', 3);
    return { status: 'error', message: err.message };
  }
}

// ============ ВЫГРУЗКА ДАННЫХ (PYTHON SERVER) ============
// Эти функции перенаправляют выгрузку акций и наборов на Python сервер

/**
 * Выгрузка акций через сервер
 * Эквивалент GAS exportPromotions()
 */
function serverExportPromotions() {
  return _loggedCall_("serverExportPromotions", function() {
    return _callServerExport('promotions');
  });
}

/**
 * Выгрузка наборов через сервер
 * Эквивалент GAS exportSets()
 */
function serverExportSets() {
  return _loggedCall_("serverExportSets", function() {
    return _callServerExport('sets');
  });
}

/**
 * Вызов сервера для выгрузки данных
 * @param {string} exportType - Тип выгрузки (promotions, sets)
 * @private
 */
function _callServerExport(exportType) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const menuTitle = "Выгрузка";

  // Определяем проект
  const projectKey = (typeof CONFIG !== 'undefined' && CONFIG.ACTIVE_PROJECT_KEY)
    ? CONFIG.ACTIVE_PROJECT_KEY.toLowerCase()
    : null;

  if (!projectKey) {
    ui.alert('Ошибка', 'Проект не определен.', ui.ButtonSet.OK);
    return { status: 'error', message: 'Project not defined' };
  }

  const typeName = exportType === 'promotions' ? 'акций' : 'наборов';
  ss.toast('Выгрузка ' + typeName + '...', menuTitle, 30);

  const BASE_URL = PropertiesService.getScriptProperties().getProperty('SERVER_URL') || 'http://localhost:8000';
  const endpoint = '/api/v1/export/' + exportType;

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      spreadsheet_id: ss.getId(),
      project: projectKey,
      dry_run: false
    }),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(BASE_URL + endpoint, options);
    const status = response.getResponseCode();
    const text = response.getContentText();

    if (status >= 200 && status < 300) {
      const result = JSON.parse(text);

      if (result && result.status === 'success') {
        ss.toast(
          'Выгружено строк: ' + result.exported_rows + '\nЛист: ' + result.target_sheet_name,
          '✅ Готово',
          5
        );
        ui.alert(
          'Успех',
          'Выгрузка ' + typeName + ' завершена.\n' +
          'Записано строк: ' + result.exported_rows + '\n' +
          'Документ: ' + result.target_url,
          ui.ButtonSet.OK
        );
      } else {
        ui.alert('Ошибка', result.message || 'Неизвестная ошибка', ui.ButtonSet.OK);
      }

      return result;
    } else {
      console.error('Export server error: ' + status + ' - ' + text);
      ss.toast('Ошибка сервера: ' + status, 'Ошибка', 3);
      return { status: 'error', message: 'Server returned ' + status };
    }
  } catch (err) {
    console.error('Export request failed: ' + err.message);
    ss.toast('Ошибка: ' + err.message, 'Ошибка', 3);
    ui.alert('Ошибка', err.message, ui.ButtonSet.OK);
    return { status: 'error', message: err.message };
  }
}

// ============ ИНВОЙСЫ (PYTHON SERVER) ============
// Эти функции перенаправляют обработку инвойсов на Python сервер

/**
 * Форматирование листа "Ордер" через сервер
 * Нормализует числовые колонки (кол-во, Цена ед., Сумма)
 * Эквивалент GAS formatOrderSheet()
 */
function serverFormatOrderSheet() {
  return _loggedCall_("serverFormatOrderSheet", function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();

    ss.toast('Форматирование листа "Ордер"...', 'Поставка', 30);

    const BASE_URL = PropertiesService.getScriptProperties().getProperty('SERVER_URL') || 'http://localhost:8000';

    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        spreadsheet_id: ss.getId(),
        sheet_name: 'Ордер',
        dry_run: false
      }),
      muteHttpExceptions: true
    };

    try {
      const response = UrlFetchApp.fetch(BASE_URL + '/api/v1/invoice/format-order', options);
      const status = response.getResponseCode();
      const text = response.getContentText();

      if (status >= 200 && status < 300) {
        const result = JSON.parse(text);

        if (result && result.status === 'success') {
          ss.toast(
            'Отформатировано строк: ' + result.rows_processed,
            '✅ Готово',
            5
          );
          ui.alert('Форматирование завершено', 'Обработано строк: ' + result.rows_processed, ui.ButtonSet.OK);
        } else {
          ui.alert('Ошибка', result.message || 'Неизвестная ошибка', ui.ButtonSet.OK);
        }

        return result;
      } else {
        console.error('Invoice format server error: ' + status + ' - ' + text);
        ss.toast('Ошибка сервера: ' + status, 'Ошибка', 3);
        return { status: 'error', message: 'Server returned ' + status };
      }
    } catch (err) {
      console.error('Invoice format request failed: ' + err.message);
      ss.toast('Ошибка: ' + err.message, 'Ошибка', 3);
      ui.alert('Ошибка', err.message, ui.ButtonSet.OK);
      return { status: 'error', message: err.message };
    }
  });
}

/**
 * Создание листа "Для инвойса" через сервер
 * Объединяет данные из Ордер, Сертификация, Этикетки
 * Эквивалент GAS createFullInvoice()
 */
function serverCreateFullInvoice() {
  return _loggedCall_("serverCreateFullInvoice", function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();

    ss.toast('Создание листа "Для инвойса"...', 'Поставка', 60);

    const BASE_URL = PropertiesService.getScriptProperties().getProperty('SERVER_URL') || 'http://localhost:8000';

    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        spreadsheet_id: ss.getId(),
        order_sheet: 'Ордер',
        certification_sheet: 'Сертификация',
        labels_sheet: 'Этикетки',
        target_sheet: 'Для инвойса',
        dry_run: false
      }),
      muteHttpExceptions: true
    };

    try {
      const response = UrlFetchApp.fetch(BASE_URL + '/api/v1/invoice/create-full', options);
      const status = response.getResponseCode();
      const text = response.getContentText();

      if (status >= 200 && status < 300) {
        const result = JSON.parse(text);

        if (result && result.status === 'success') {
          ss.toast(
            'Создан лист "' + result.target_sheet + '" с ' + result.rows_processed + ' строками',
            '✅ Готово',
            5
          );
          ui.alert(
            'Успех',
            'Лист "Для инвойса" создан.\nЗаписано строк: ' + result.rows_processed,
            ui.ButtonSet.OK
          );
        } else {
          ui.alert('Ошибка', result.message || 'Неизвестная ошибка', ui.ButtonSet.OK);
        }

        return result;
      } else {
        console.error('Invoice create server error: ' + status + ' - ' + text);
        ss.toast('Ошибка сервера: ' + status, 'Ошибка', 3);
        return { status: 'error', message: 'Server returned ' + status };
      }
    } catch (err) {
      console.error('Invoice create request failed: ' + err.message);
      ss.toast('Ошибка: ' + err.message, 'Ошибка', 3);
      ui.alert('Ошибка', err.message, ui.ButtonSet.OK);
      return { status: 'error', message: err.message };
    }
  });
}

// ============ СЕРТИФИКАЦИЯ (PYTHON SERVER) ============
// Эти функции перенаправляют операции сертификации на Python сервер

/**
 * Создание листа новинок из Сертификации через сервер
 * Эквивалент GAS createNewsSheetFromCertification()
 */
function serverCreateNewsSheet() {
  return _loggedCall_("serverCreateNewsSheet", function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();

    ss.toast('Создание листа "New sert"...', 'Сертификация', 30);

    const BASE_URL = PropertiesService.getScriptProperties().getProperty('SERVER_URL') || 'http://localhost:8000';

    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        spreadsheet_id: ss.getId(),
        source_sheet: 'Сертификация',
        target_sheet: 'New sert',
        dry_run: false
      }),
      muteHttpExceptions: true
    };

    try {
      const response = UrlFetchApp.fetch(BASE_URL + '/api/v1/certification/news-sheet', options);
      const status = response.getResponseCode();
      const text = response.getContentText();

      if (status >= 200 && status < 300) {
        const result = JSON.parse(text);

        if (result && result.status === 'success') {
          ss.toast(
            'Создан лист "' + result.sheet_name + '" с ' + result.rows_affected + ' новинками',
            '✅ Готово',
            5
          );
          ui.alert(
            'Успех',
            'Лист новинок создан.\nНайдено новинок: ' + result.rows_affected,
            ui.ButtonSet.OK
          );
        } else {
          ui.alert('Ошибка', result.message || 'Неизвестная ошибка', ui.ButtonSet.OK);
        }

        return result;
      } else {
        console.error('News sheet server error: ' + status + ' - ' + text);
        ss.toast('Ошибка сервера: ' + status, 'Ошибка', 3);
        return { status: 'error', message: 'Server returned ' + status };
      }
    } catch (err) {
      console.error('News sheet request failed: ' + err.message);
      ss.toast('Ошибка: ' + err.message, 'Ошибка', 3);
      ui.alert('Ошибка', err.message, ui.ButtonSet.OK);
      return { status: 'error', message: err.message };
    }
  });
}

/**
 * Расчёт номеров спиртов через сервер
 * Эквивалент GAS calculateAndAssignSpiritNumbers()
 */
function serverCalculateSpiritNumbers() {
  return _loggedCall_("serverCalculateSpiritNumbers", function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();

    ss.toast('Расчёт номеров спиртов...', 'Сертификация', 30);

    const BASE_URL = PropertiesService.getScriptProperties().getProperty('SERVER_URL') || 'http://localhost:8000';

    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        spreadsheet_id: ss.getId(),
        sheet_name: 'Сертификация',
        dry_run: false
      }),
      muteHttpExceptions: true
    };

    try {
      const response = UrlFetchApp.fetch(BASE_URL + '/api/v1/certification/spirits/calculate', options);
      const status = response.getResponseCode();
      const text = response.getContentText();

      if (status >= 200 && status < 300) {
        const result = JSON.parse(text);

        if (result && result.status === 'success') {
          ss.toast(
            'Рассчитано строк: ' + result.rows_affected,
            '✅ Готово',
            5
          );
          ui.alert('Расчёт завершён', result.message, ui.ButtonSet.OK);
        } else if (result && result.status === 'not_implemented') {
          ui.alert('В разработке', result.message, ui.ButtonSet.OK);
        } else {
          ui.alert('Ошибка', result.message || 'Неизвестная ошибка', ui.ButtonSet.OK);
        }

        return result;
      } else {
        console.error('Spirit calc server error: ' + status + ' - ' + text);
        ss.toast('Ошибка сервера: ' + status, 'Ошибка', 3);
        return { status: 'error', message: 'Server returned ' + status };
      }
    } catch (err) {
      console.error('Spirit calc request failed: ' + err.message);
      ss.toast('Ошибка: ' + err.message, 'Ошибка', 3);
      ui.alert('Ошибка', err.message, ui.ButtonSet.OK);
      return { status: 'error', message: err.message };
    }
  });
}

/**
 * Генерация протоколов 353пп через сервер
 * Эквивалент GAS generateProtocols_353pp()
 */
function serverGenerateProtocols353pp() {
  return _loggedCall_("serverGenerateProtocols353pp", function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();

    ss.toast('Генерация протоколов 353пп...', 'Сертификация', 60);

    const BASE_URL = PropertiesService.getScriptProperties().getProperty('SERVER_URL') || 'http://localhost:8000';

    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        spreadsheet_id: ss.getId(),
        protocol_type: '353pp',
        dry_run: false
      }),
      muteHttpExceptions: true
    };

    try {
      const response = UrlFetchApp.fetch(BASE_URL + '/api/v1/certification/protocols-353pp', options);
      const status = response.getResponseCode();
      const text = response.getContentText();

      if (status >= 200 && status < 300) {
        const result = JSON.parse(text);

        if (result && result.status === 'not_implemented') {
          ui.alert(
            'В разработке',
            result.message + '\n\n' + (result.hint || ''),
            ui.ButtonSet.OK
          );
        } else if (result && result.status === 'success') {
          ss.toast('Протоколы сгенерированы', '✅ Готово', 5);
          ui.alert('Успех', result.message, ui.ButtonSet.OK);
        } else {
          ui.alert('Ошибка', result.message || 'Неизвестная ошибка', ui.ButtonSet.OK);
        }

        return result;
      } else {
        console.error('Protocols server error: ' + status + ' - ' + text);
        ss.toast('Ошибка сервера: ' + status, 'Ошибка', 3);
        return { status: 'error', message: 'Server returned ' + status };
      }
    } catch (err) {
      console.error('Protocols request failed: ' + err.message);
      ss.toast('Ошибка: ' + err.message, 'Ошибка', 3);
      ui.alert('Ошибка', err.message, ui.ButtonSet.OK);
      return { status: 'error', message: err.message };
    }
  });
}

/**
 * Генерация макетов ДС через сервер
 * Эквивалент GAS generateDsLayouts_353pp()
 */
function serverGenerateDsLayouts() {
  return _loggedCall_("serverGenerateDsLayouts", function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();

    ss.toast('Генерация макетов ДС...', 'Сертификация', 60);

    const BASE_URL = PropertiesService.getScriptProperties().getProperty('SERVER_URL') || 'http://localhost:8000';

    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        spreadsheet_id: ss.getId(),
        dry_run: false
      }),
      muteHttpExceptions: true
    };

    try {
      const response = UrlFetchApp.fetch(BASE_URL + '/api/v1/certification/ds-layouts', options);
      const status = response.getResponseCode();
      const text = response.getContentText();

      if (status >= 200 && status < 300) {
        const result = JSON.parse(text);

        if (result && result.status === 'not_implemented') {
          ui.alert(
            'В разработке',
            result.message + '\n\n' + (result.hint || ''),
            ui.ButtonSet.OK
          );
        } else if (result && result.status === 'success') {
          ss.toast('Макеты ДС сгенерированы', '✅ Готово', 5);
          ui.alert('Успех', result.message, ui.ButtonSet.OK);
        } else {
          ui.alert('Ошибка', result.message || 'Неизвестная ошибка', ui.ButtonSet.OK);
        }

        return result;
      } else {
        console.error('DS layouts server error: ' + status + ' - ' + text);
        ss.toast('Ошибка сервера: ' + status, 'Ошибка', 3);
        return { status: 'error', message: 'Server returned ' + status };
      }
    } catch (err) {
      console.error('DS layouts request failed: ' + err.message);
      ss.toast('Ошибка: ' + err.message, 'Ошибка', 3);
      ui.alert('Ошибка', err.message, ui.ButtonSet.OK);
      return { status: 'error', message: err.message };
    }
  });
}


// =======================================================================================
// FORMULA OPERATIONS (Python Server)
// =======================================================================================

/**
 * Recalculate price dynamics formulas via Python server.
 * Calculates EXW ALFASPA, Purchase price, DDP, and Growth for all year blocks.
 *
 * Equivalent to local Lib.recalculatePriceDynamicsFormulas().
 *
 * @param {string} sheetName - Optional sheet name (default: "Динамика цены")
 * @param {boolean} dryRun - If true, only preview changes
 * @returns {Object} Result with blocks_processed, rows_updated, message
 */
function serverRecalculatePriceDynamicsFormulas(sheetName, dryRun) {
  return _withLock('formulaPriceDynamics', function() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ui = SpreadsheetApp.getUi();
    var spreadsheetId = ss.getId();

    ss.toast('Пересчитываю формулы Динамика цены...', 'Python сервер', -1);

    try {
      var url = SERVER_URL + '/formulas/price-dynamics';
      var payload = {
        spreadsheet_id: spreadsheetId,
        sheet_name: sheetName || 'Динамика цены',
        dry_run: !!dryRun
      };

      var options = {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true,
        headers: _getAuthHeaders()
      };

      var response = UrlFetchApp.fetch(url, options);
      var status = response.getResponseCode();
      var text = response.getContentText();

      if (status === 200) {
        var result = JSON.parse(text);

        if (result && result.status === 'success') {
          ss.toast(
            'Обработано ' + (result.blocks_processed || 0) + ' блоков, ' +
            (result.rows_updated || 0) + ' строк',
            '✅ Готово', 5
          );
        } else {
          ss.toast(result.message || 'Ошибка', 'Ошибка', 3);
        }

        return result;
      } else {
        console.error('Price dynamics formulas server error: ' + status + ' - ' + text);
        ss.toast('Ошибка сервера: ' + status, 'Ошибка', 3);
        return { status: 'error', message: 'Server returned ' + status };
      }
    } catch (err) {
      console.error('Price dynamics formulas request failed: ' + err.message);
      ss.toast('Ошибка: ' + err.message, 'Ошибка', 3);
      return { status: 'error', message: err.message };
    }
  });
}


/**
 * Update price calculation formulas via Python server.
 * Pulls data from Price Dynamics sheet using INDEX/MATCH logic.
 *
 * Equivalent to local Lib.updatePriceCalculationFormulas().
 *
 * @param {boolean} silent - If true, suppress notifications
 * @param {boolean} dryRun - If true, only preview changes
 * @returns {Object} Result with rows_updated, message
 */
function serverUpdatePriceCalculationFormulas(silent, dryRun) {
  return _withLock('formulaPriceCalc', function() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ui = SpreadsheetApp.getUi();
    var spreadsheetId = ss.getId();

    if (!silent) {
      ss.toast('Обновляю формулы Расчет цены...', 'Python сервер', -1);
    }

    try {
      var url = SERVER_URL + '/formulas/price-calculation';
      var payload = {
        spreadsheet_id: spreadsheetId,
        price_calc_sheet: 'Расчет цены',
        price_dynamics_sheet: 'Динамика цены',
        silent: !!silent,
        dry_run: !!dryRun
      };

      var options = {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true,
        headers: _getAuthHeaders()
      };

      var response = UrlFetchApp.fetch(url, options);
      var status = response.getResponseCode();
      var text = response.getContentText();

      if (status === 200) {
        var result = JSON.parse(text);

        if (result && result.status === 'success') {
          if (!silent) {
            ss.toast(
              'Обновлено ' + (result.rows_updated || 0) + ' строк',
              '✅ Готово', 5
            );
          }
        } else if (!silent) {
          ss.toast(result.message || 'Ошибка', 'Ошибка', 3);
        }

        return result;
      } else {
        console.error('Price calc formulas server error: ' + status + ' - ' + text);
        if (!silent) {
          ss.toast('Ошибка сервера: ' + status, 'Ошибка', 3);
        }
        return { status: 'error', message: 'Server returned ' + status };
      }
    } catch (err) {
      console.error('Price calc formulas request failed: ' + err.message);
      if (!silent) {
        ss.toast('Ошибка: ' + err.message, 'Ошибка', 3);
      }
      return { status: 'error', message: err.message };
    }
  });
}


/**
 * Add new year columns to price dynamics sheet via Python server.
 * Inserts 7 columns after "Комментарий" with proper headers and formatting.
 *
 * Equivalent to local Lib.addNewYearColumnsToPriceDynamics().
 *
 * @param {number} year - Optional year (default: current year)
 * @param {boolean} dryRun - If true, only preview changes
 * @returns {Object} Result with columns_added, year, message
 */
function serverAddNewYearColumns(year, dryRun) {
  return _withLock('formulaAddYear', function() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ui = SpreadsheetApp.getUi();
    var spreadsheetId = ss.getId();

    var targetYear = year || new Date().getFullYear();
    ss.toast('Добавляю столбцы за ' + targetYear + ' год...', 'Python сервер', -1);

    try {
      var url = SERVER_URL + '/formulas/add-year-columns';
      var payload = {
        spreadsheet_id: spreadsheetId,
        sheet_name: 'Динамика цены',
        year: targetYear,
        dry_run: !!dryRun
      };

      var options = {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true,
        headers: _getAuthHeaders()
      };

      var response = UrlFetchApp.fetch(url, options);
      var status = response.getResponseCode();
      var text = response.getContentText();

      if (status === 200) {
        var result = JSON.parse(text);

        if (result && result.status === 'success') {
          ss.toast(
            'Добавлено ' + (result.columns_added || 0) + ' столбцов за ' +
            (result.year || targetYear) + ' год',
            '✅ Готово', 5
          );
          ui.alert(
            'Успех',
            'Добавлены столбцы за ' + (result.year || targetYear) + ' год.\n' +
            result.message,
            ui.ButtonSet.OK
          );
        } else if (result && result.status === 'exists') {
          ss.toast('Блок за ' + targetYear + ' год уже существует', 'Инфо', 3);
          ui.alert('Информация', result.message, ui.ButtonSet.OK);
        } else {
          ss.toast(result.message || 'Ошибка', 'Ошибка', 3);
          ui.alert('Ошибка', result.message || 'Неизвестная ошибка', ui.ButtonSet.OK);
        }

        return result;
      } else {
        console.error('Add year columns server error: ' + status + ' - ' + text);
        ss.toast('Ошибка сервера: ' + status, 'Ошибка', 3);
        return { status: 'error', message: 'Server returned ' + status };
      }
    } catch (err) {
      console.error('Add year columns request failed: ' + err.message);
      ss.toast('Ошибка: ' + err.message, 'Ошибка', 3);
      ui.alert('Ошибка', err.message, ui.ButtonSet.OK);
      return { status: 'error', message: err.message };
    }
  });
}


// =======================================================================================
// LOG ARCHIVING (Python Server)
// =======================================================================================

/**
 * Archive logs to monthly spreadsheet via Python server.
 * Copies data from log sheets to Drive archive folder.
 *
 * Equivalent to local Lib.archiveLogsDaily().
 *
 * @param {string} archiveFolderId - ID of Drive folder for archives
 * @param {string} projectCode - Project code (mt, sk, ss)
 * @param {boolean} dryRun - If true, only preview changes
 * @returns {Object} Result with total_rows, sheets_archived, archive_name
 */
function serverArchiveLogsDaily(archiveFolderId, projectCode, dryRun) {
  return _withLock('logArchive', function() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ui = SpreadsheetApp.getUi();
    var spreadsheetId = ss.getId();

    ss.toast('Архивирую логи...', 'Python сервер', -1);

    try {
      var url = SERVER_URL + '/logs/archive';
      var payload = {
        spreadsheet_id: spreadsheetId,
        archive_folder_id: archiveFolderId || '',
        project_code: projectCode || 'project',
        dry_run: !!dryRun
      };

      var options = {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true,
        headers: _getAuthHeaders()
      };

      var response = UrlFetchApp.fetch(url, options);
      var status = response.getResponseCode();
      var text = response.getContentText();

      if (status === 200) {
        var result = JSON.parse(text);

        if (result && result.status === 'success') {
          ss.toast(
            'Архивировано ' + (result.total_rows || 0) + ' строк',
            '✅ Готово', 5
          );
          ui.alert(
            'Архивирование завершено',
            'Архивировано ' + (result.total_rows || 0) + ' строк.\n\n' +
            'Архив: ' + (result.archive_name || 'unknown'),
            ui.ButtonSet.OK
          );
        } else {
          ss.toast(result.message || 'Ошибка', 'Ошибка', 3);
        }

        return result;
      } else {
        console.error('Log archive server error: ' + status + ' - ' + text);
        ss.toast('Ошибка сервера: ' + status, 'Ошибка', 3);
        return { status: 'error', message: 'Server returned ' + status };
      }
    } catch (err) {
      console.error('Log archive request failed: ' + err.message);
      ss.toast('Ошибка: ' + err.message, 'Ошибка', 3);
      return { status: 'error', message: err.message };
    }
  });
}


/**
 * Reset log sheet via Python server.
 * Clears all data rows, preserving headers.
 *
 * Equivalent to local Lib.resetDailyLogSheet().
 *
 * @param {string} sheetName - Name of log sheet to reset (default: Логи)
 * @param {boolean} dryRun - If true, only preview changes
 * @returns {Object} Result with rows cleared
 */
function serverResetLogSheet(sheetName, dryRun) {
  return _withLock('logReset', function() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var spreadsheetId = ss.getId();

    ss.toast('Очищаю лист логов...', 'Python сервер', -1);

    try {
      var url = SERVER_URL + '/logs/reset';
      var payload = {
        spreadsheet_id: spreadsheetId,
        sheet_name: sheetName || 'Логи',
        dry_run: !!dryRun
      };

      var options = {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true,
        headers: _getAuthHeaders()
      };

      var response = UrlFetchApp.fetch(url, options);
      var status = response.getResponseCode();
      var text = response.getContentText();

      if (status === 200) {
        var result = JSON.parse(text);

        if (result && result.status === 'success') {
          ss.toast(
            'Очищено ' + (result.total_rows || 0) + ' строк',
            '✅ Готово', 5
          );
        } else {
          ss.toast(result.message || 'Ошибка', 'Ошибка', 3);
        }

        return result;
      } else {
        console.error('Log reset server error: ' + status + ' - ' + text);
        ss.toast('Ошибка сервера: ' + status, 'Ошибка', 3);
        return { status: 'error', message: 'Server returned ' + status };
      }
    } catch (err) {
      console.error('Log reset request failed: ' + err.message);
      ss.toast('Ошибка: ' + err.message, 'Ошибка', 3);
      return { status: 'error', message: err.message };
    }
  });
}


/**
 * Perform midnight log rotation via Python server.
 * Archives logs and resets all log sheets.
 *
 * Equivalent to local Lib.midnightLogRotation().
 *
 * @param {string} archiveFolderId - ID of Drive folder for archives
 * @param {string} projectCode - Project code (mt, sk, ss)
 * @param {boolean} force - Force rotation even if already done today
 * @param {boolean} dryRun - If true, only preview changes
 * @returns {Object} Result with rotation details
 */
function serverMidnightLogRotation(archiveFolderId, projectCode, force, dryRun) {
  return _withLock('logRotation', function() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ui = SpreadsheetApp.getUi();
    var spreadsheetId = ss.getId();

    ss.toast('Выполняю ротацию логов...', 'Python сервер', -1);

    try {
      var url = SERVER_URL + '/logs/rotation';
      var payload = {
        spreadsheet_id: spreadsheetId,
        archive_folder_id: archiveFolderId || '',
        project_code: projectCode || 'project',
        force: !!force,
        dry_run: !!dryRun
      };

      var options = {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true,
        headers: _getAuthHeaders()
      };

      var response = UrlFetchApp.fetch(url, options);
      var status = response.getResponseCode();
      var text = response.getContentText();

      if (status === 200) {
        var result = JSON.parse(text);

        if (result && result.status === 'success') {
          ss.toast(
            'Ротация завершена: ' + (result.total_rows || 0) + ' строк',
            '✅ Готово', 5
          );
          ui.alert(
            'Ротация логов завершена',
            result.message || 'Логи архивированы и листы очищены',
            ui.ButtonSet.OK
          );
        } else if (result && result.status === 'skipped') {
          ss.toast('Ротация пропущена: уже выполнена сегодня', 'Инфо', 3);
        } else {
          ss.toast(result.message || 'Ошибка', 'Ошибка', 3);
        }

        return result;
      } else {
        console.error('Log rotation server error: ' + status + ' - ' + text);
        ss.toast('Ошибка сервера: ' + status, 'Ошибка', 3);
        return { status: 'error', message: 'Server returned ' + status };
      }
    } catch (err) {
      console.error('Log rotation request failed: ' + err.message);
      ss.toast('Ошибка: ' + err.message, 'Ошибка', 3);
      return { status: 'error', message: err.message };
    }
  });
}


/**
 * Get log archive status via Python server.
 *
 * Equivalent to local Lib.showArchiveStatus().
 *
 * @param {string} projectCode - Project code (mt, sk, ss)
 * @returns {Object} Status information
 */
function serverGetLogStatus(projectCode) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var spreadsheetId = ss.getId();

  try {
    var url = SERVER_URL + '/logs/status?spreadsheet_id=' +
              encodeURIComponent(spreadsheetId) +
              '&project_code=' + encodeURIComponent(projectCode || 'project');

    var options = {
      method: 'get',
      muteHttpExceptions: true,
      headers: _getAuthHeaders()
    };

    var response = UrlFetchApp.fetch(url, options);
    var status = response.getResponseCode();
    var text = response.getContentText();

    if (status === 200) {
      var result = JSON.parse(text);

      var msg = '📊 СТАТУС АРХИВИРОВАНИЯ ЛОГОВ\n\n';
      msg += '📅 Последнее архивирование: ' + (result.last_archive_date || 'Никогда') + '\n';
      msg += '📁 Текущий архив: ' + (result.current_archive_name || '-') + '\n';
      msg += '📝 Ожидает архивации: ' + (result.total_pending_rows || 0) + ' строк\n\n';

      if (result.current_row_counts) {
        msg += 'По листам:\n';
        for (var sheet in result.current_row_counts) {
          msg += '  • ' + sheet + ': ' + result.current_row_counts[sheet] + ' строк\n';
        }
      }

      ui.alert('Статус архивирования', msg, ui.ButtonSet.OK);

      return result;
    } else {
      console.error('Log status server error: ' + status + ' - ' + text);
      ui.alert('Ошибка', 'Ошибка сервера: ' + status, ui.ButtonSet.OK);
      return { status: 'error', message: 'Server returned ' + status };
    }
  } catch (err) {
    console.error('Log status request failed: ' + err.message);
    ui.alert('Ошибка', err.message, ui.ButtonSet.OK);
    return { status: 'error', message: err.message };
  }
}
