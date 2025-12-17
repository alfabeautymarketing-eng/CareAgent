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

// ============ ТРИГГЕРЫ (SIMPLE TRIGGERS) ============

/**
 * Simple Trigger: Запускается при открытии документа.
 * Инициализирует меню (старое и новое).
 */
function onOpen(e) {
  // 1. Инициализация Legacy меню (из Config)
  if (typeof Lib !== 'undefined' && typeof Lib.onOpen === 'function') {
    try {
      Lib.onOpen(e);
    } catch (err) {
      console.error("Ошибка при создании Legacy меню: " + err);
    }
  }

  // 2. Инициализация меню Агента
  if (typeof createAgentMenu === 'function') {
    createAgentMenu();
  } else {
    console.error("createAgentMenu не найдена!");
  }
}

/**
 * Simple Trigger: Автоматическая синхронизация при редактировании ячеек.
 * Срабатывает при любом изменении в таблице.
 */
function onEdit(e) {
  if (typeof Lib !== 'undefined' && typeof Lib.onEdit_internal_ === 'function') {
    Lib.onEdit_internal_(e);
  }
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
 */
function handleOnEdit(e) {
  if (typeof Lib !== 'undefined' && typeof Lib.onEdit_internal_ === 'function') {
    Lib.onEdit_internal_(e);
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
  if (typeof Lib !== 'undefined' && Lib.addArticleManually) {
    return Lib.addArticleManually();
  }
  throw new Error('Lib.addArticleManually не определена');
}

function deleteSelectedRowsWithSync() {
  if (typeof Lib !== 'undefined' && Lib.deleteSelectedRowsWithSync) {
    return Lib.deleteSelectedRowsWithSync();
  }
  throw new Error('Lib.deleteSelectedRowsWithSync не определена');
}

function syncSelectedRow() {
  if (typeof Lib !== 'undefined' && Lib.syncSelectedRow) {
    return Lib.syncSelectedRow();
  }
  throw new Error('Lib.syncSelectedRow не определена');
}

function runFullSync() {
  if (typeof Lib !== 'undefined' && Lib.runFullSync) {
    return Lib.runFullSync();
  }
  throw new Error('Lib.runFullSync не определена');
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
  if (typeof Lib !== 'undefined' && Lib.processSkPriceSheet) {
    return Lib.processSkPriceSheet();
  }
  throw new Error('Lib.processSkPriceSheet не определена');
}

function loadSkStockData() {
  if (typeof Lib !== 'undefined' && Lib.loadSkStockData) {
    return Lib.loadSkStockData();
  }
  throw new Error('Lib.loadSkStockData не определена');
}

// ============ ОБРАБОТКА ПРАЙСОВ (MT) ============

function processMtMainPrice() {
  if (typeof Lib !== 'undefined' && Lib.processMtMainPrice) {
    return Lib.processMtMainPrice();
  }
  throw new Error('Lib.processMtMainPrice не определена');
}

function processMtTesterPrice() {
  if (typeof Lib !== 'undefined' && Lib.processMtTesterPrice) {
    return Lib.processMtTesterPrice();
  }
  throw new Error('Lib.processMtTesterPrice не определена');
}

function processMtSamplesPrice() {
  if (typeof Lib !== 'undefined' && Lib.processMtSamplesPrice) {
    return Lib.processMtSamplesPrice();
  }
  throw new Error('Lib.processMtSamplesPrice не определена');
}

function loadMtStockData() {
  if (typeof Lib !== 'undefined' && Lib.loadMtStockData) {
    return Lib.loadMtStockData();
  }
  throw new Error('Lib.loadMtStockData не определена');
}

// ============ ОБРАБОТКА ПРАЙСОВ (SS) ============

function processSsPriceSheet() {
  if (typeof Lib !== 'undefined' && Lib.processSsPriceSheet) {
    return Lib.processSsPriceSheet();
  }
  throw new Error('Lib.processSsPriceSheet не определена');
}

function loadSsStockData() {
  if (typeof Lib !== 'undefined' && Lib.loadSsStockData) {
    return Lib.loadSsStockData();
  }
  throw new Error('Lib.loadSsStockData не определена');
}

// ============ ОБЩИЕ ФУНКЦИИ ЗАКАЗА ============

function addNewYearColumnsToPriceDynamics() {
  if (typeof Lib !== 'undefined' && Lib.addNewYearColumnsToPriceDynamics) {
    return Lib.addNewYearColumnsToPriceDynamics();
  }
  throw new Error('Lib.addNewYearColumnsToPriceDynamics не определена');
}

function sortSkOrderByManufacturer() {
  if (typeof Lib !== 'undefined' && Lib.sortSkOrderByManufacturer) {
    return Lib.sortSkOrderByManufacturer();
  }
  throw new Error('Lib.sortSkOrderByManufacturer не определена');
}

function sortSkOrderByPrice() {
  if (typeof Lib !== 'undefined' && Lib.sortSkOrderByPrice) {
    return Lib.sortSkOrderByPrice();
  }
  throw new Error('Lib.sortSkOrderByPrice не определена');
}

function sortMtOrderByManufacturer() {
  if (typeof Lib !== 'undefined' && Lib.sortMtOrderByManufacturer) {
    return Lib.sortMtOrderByManufacturer();
  }
  throw new Error('Lib.sortMtOrderByManufacturer не определена');
}

function sortMtOrderByPrice() {
  if (typeof Lib !== 'undefined' && Lib.sortMtOrderByPrice) {
    return Lib.sortMtOrderByPrice();
  }
  throw new Error('Lib.sortMtOrderByPrice не определена');
}

function sortSsOrderByManufacturer() {
  if (typeof Lib !== 'undefined' && Lib.sortSsOrderByManufacturer) {
    return Lib.sortSsOrderByManufacturer();
  }
  throw new Error('Lib.sortSsOrderByManufacturer не определена');
}

function sortSsOrderByPrice() {
  if (typeof Lib !== 'undefined' && Lib.sortSsOrderByPrice) {
    return Lib.sortSsOrderByPrice();
  }
  throw new Error('Lib.sortSsOrderByPrice не определена');
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
  if (typeof Lib !== 'undefined' && Lib.formatOrderSheet) {
    return Lib.formatOrderSheet();
  }
  throw new Error('Lib.formatOrderSheet не определена');
}

function createFullInvoice() {
  if (typeof Lib !== 'undefined' && Lib.createFullInvoice) {
    return Lib.createFullInvoice();
  }
  throw new Error('Lib.createFullInvoice не определена');
}

function collectAndCopyDocuments() {
  if (typeof Lib !== 'undefined' && Lib.collectAndCopyDocuments) {
    return Lib.collectAndCopyDocuments();
  }
  throw new Error('Lib.collectAndCopyDocuments не определена');
}

// ============ СЕРТИФИКАЦИЯ ============

function createNewsSheetFromCertification() {
  if (typeof Lib !== 'undefined' && Lib.createNewsSheetFromCertification) {
    return Lib.createNewsSheetFromCertification();
  }
  throw new Error('Lib.createNewsSheetFromCertification не определена');
}

function generateProtocols_353pp() {
  if (typeof Lib !== 'undefined' && Lib.generateProtocols_353pp) {
    return Lib.generateProtocols_353pp();
  }
  throw new Error('Lib.generateProtocols_353pp не определена');
}

function generateDsLayouts_353pp() {
  if (typeof Lib !== 'undefined' && Lib.generateDsLayouts_353pp) {
    return Lib.generateDsLayouts_353pp();
  }
  throw new Error('Lib.generateDsLayouts_353pp не определена');
}

function structureDocuments_353pp() {
  if (typeof Lib !== 'undefined' && Lib.structureDocuments_353pp) {
    return Lib.structureDocuments_353pp();
  }
  throw new Error('Lib.structureDocuments_353pp не определена');
}

function calculateAndAssignSpiritNumbers() {
  if (typeof Lib !== 'undefined' && Lib.calculateAndAssignSpiritNumbers) {
    return Lib.calculateAndAssignSpiritNumbers();
  }
  throw new Error('Lib.calculateAndAssignSpiritNumbers не определена');
}

function generateSpiritProtocols() {
  if (typeof Lib !== 'undefined' && Lib.generateSpiritProtocols) {
    return Lib.generateSpiritProtocols();
  }
  throw new Error('Lib.generateSpiritProtocols не определена');
}

function runManualCascadeOnCertification() {
  if (typeof Lib !== 'undefined' && Lib.runManualCascadeOnCertification) {
    return Lib.runManualCascadeOnCertification();
  }
  throw new Error('Lib.runManualCascadeOnCertification не определена');
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
function reorderAuxiliarySheets() {
  if (typeof Lib !== 'undefined' && Lib.reorderAuxiliarySheets) {
    return Lib.reorderAuxiliarySheets();
  }
  throw new Error('Lib.reorderAuxiliarySheets не определена');
}
