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
 * VERIFICATION MARKER: 199Np7xsBiBRQih5_tlUdpt6EmkfRGjZAhTvKm4Ua0Q6XEaMtvAmQUn0g
 * =======================================================================================
 */

/**
 * Simple Trigger: Запускается при открытии документа.
 * Ограничен в правах (не может делать UrlFetchApp).
 * Служит только для построения меню (часто из кэша).
 */
function onOpen(e) {
  // Ecosystem Menu (Simple trigger mode: no network)
  // This ensures a fallback menu is visible immediately.
  if (typeof createAgentMenu === 'function') {
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
  
  // -1. Проверяем и устанавливаем URL сервера если пусто
  try {
    const scriptProps = PropertiesService.getScriptProperties();
    if (!scriptProps.getProperty('SERVER_URL')) {
      console.log("SERVER_URL property missing, initializing...");
      initProductionServerUrl();
    }
  } catch (err) {
    console.error("Ошибка при проверке SERVER_URL: " + err);
  }

  // 1. Загрузка меню (требует полного доступа)
  if (typeof createAgentMenu === 'function') {
    try {
      createAgentMenu();
    } catch (err) {
      console.error("Ошибка при загрузке меню: " + err);
    }
  }

  // 2. Автоматическое упорядочивание листов через Python сервер
  if (typeof reorderSheetsSilent === 'function') {
    try {
      reorderSheetsSilent();
    } catch (err) {
      console.error("Ошибка при упорядочивании листов: " + err);
    }
  }

  // 3.5. Инициализация Gemini API из Script Properties
  if (typeof initGeminiFromStorage === 'function') {
    try {
      const geminiOk = initGeminiFromStorage();
    } catch (err) {
      console.error("Ошибка при инициализации Gemini: " + err);
    }
  }

  // 4. Обновляем формулы на ключевых листах
  try {
    if (typeof Lib !== 'undefined' && typeof Lib.recalculatePriceDynamicsFormulas === 'function') {
      Lib.recalculatePriceDynamicsFormulas();
    }
  } catch (err) {
    console.error("Ошибка при обновлении формул Динамика цены: " + err);
  }

  try {
    if (typeof Lib !== 'undefined' && typeof Lib.updatePriceCalculationFormulas === 'function') {
      Lib.updatePriceCalculationFormulas(true); // silent
    }
  } catch (err) {
    console.error("Ошибка при обновлении формул Расчет цены: " + err);
  }

  // 5. Активация листа "Главная"
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var mainSheet = ss.getSheetByName("Главная");
    if (mainSheet) {
      ss.setActiveSheet(mainSheet);
    }
  } catch (err) {
    console.error("Ошибка при активации листа Главная: " + err);
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
  // Редактирование обрабатывается сервером через вебхуки

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
  // ОТКЛЮЧЕНО: Лист 'Логи' больше не используется.
  // console.log('[LOG]', category, action, details, status);
}

/**
 * Обертка для логирования вызовов функций.
 * Теперь пишет только в console.log для отладки в GAS.
 */
function _loggedCall_(functionName, fn, source) {
  source = source || "FUNCTION";
  var startTime = Date.now();

  try {
    var result = fn();
    var duration = Date.now() - startTime;
    console.log('[OK] ' + functionName + ' (' + duration + 'ms)');
    return result;
  } catch (e) {
    var duration = Date.now() - startTime;
    console.error('[ERROR] ' + functionName + ' after ' + duration + 'ms: ' + e.message);
    throw e;
  }
}

// ============ СИНХРОНИЗАЦИЯ ============

/**
 * Служебная функция для принудительного закрытия любых висящих toast-уведомлений.
 * Вызовите эту функцию из меню, если toast "застрял".
 */
function clearAllToasts() {
  if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
    Lib.logWithEmoji("Запуск принудительной очистки уведомлений", "INFO", "", "clearAllToasts", "Очистка висящих toast-уведомлений", "System", "START", {}, null);
  }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    // Закрываем любые висящие toast
    ss.toast('', '', 1);
    SpreadsheetApp.getUi().alert('Toast-уведомления очищены', 'Все активные уведомления были закрыты.', SpreadsheetApp.getUi().ButtonSet.OK);
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Уведомления успешно очищены", "INFO", "", "clearAllToasts", "Toast-уведомления и алерты закрыты", "System", "SUCCESS", {}, { status: "success" });
    }
  } catch (e) {
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Ошибка при очистке уведомлений", "ERROR", "", "clearAllToasts", e.message, "System", "ERROR", {}, { error: e.toString() });
    }
    SpreadsheetApp.getUi().alert('Ошибка', 'Не удалось закрыть toast: ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function addArticleManually() {
  return _loggedCall_("addArticleManually", function() {
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Запуск ручного добавления артикула", "INFO", "", "addArticleManually", "Создание нового товара через GAS", "Sync", "START", {}, null);
    }
    try {
      if (typeof Lib !== 'undefined' && Lib.addArticleManually) {
        var result = Lib.addArticleManually();
        if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
          Lib.logWithEmoji("Артикул добавлен вручную", "INFO", "", "addArticleManually", "Товар успешно создан", "Sync", "SUCCESS", {}, { result: result });
        }
        return result;
      }
      throw new Error('Lib.addArticleManually не определена');
    } catch (e) {
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Ошибка при ручном добавлении", "ERROR", "", "addArticleManually", e.message, "Sync", "ERROR", {}, { error: e.toString() });
      }
      throw e;
    }
  });
}

function deleteSelectedRowsWithSync() {
  return _loggedCall_("deleteSelectedRowsWithSync", function() {
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Запуск удаления строк с синхронизацией", "INFO", "", "deleteSelectedRowsWithSync", "Удаление выделенных товаров и десинхронизация", "Sync", "START", {}, null);
    }
    try {
      if (typeof Lib !== 'undefined' && Lib.deleteSelectedRowsWithSync) {
        var result = Lib.deleteSelectedRowsWithSync();
        if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
          Lib.logWithEmoji("Удаление строк завершено", "INFO", "", "deleteSelectedRowsWithSync", "Строки удалены и синхронизированы", "Sync", "SUCCESS", {}, { result: result });
        }
        return result;
      }
      throw new Error('Lib.deleteSelectedRowsWithSync не определена');
    } catch (e) {
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Ошибка при удалении строк", "ERROR", "", "deleteSelectedRowsWithSync", e.message, "Sync", "ERROR", {}, { error: e.toString() });
      }
      throw e;
    }
  });
}

function syncSelectedRow() {
  return _loggedCall_("syncSelectedRow", function() {
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Запуск синхронизации выделенной строки", "INFO", "", "syncSelectedRow", "Обновление данных артикула на сервере", "Sync", "START", {}, null);
    }
    try {
      if (typeof Lib !== 'undefined' && Lib.syncSelectedRow) {
        var result = Lib.syncSelectedRow();
        if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
          Lib.logWithEmoji("Синхронизация строки завершена", "INFO", "", "syncSelectedRow", "Данные успешно переданы на сервер", "Sync", "SUCCESS", {}, { result: result });
        }
        return result;
      }
      throw new Error('Lib.syncSelectedRow не определена');
    } catch (e) {
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Ошибка синхронизации строки", "ERROR", "", "syncSelectedRow", e.message, "Sync", "ERROR", {}, { error: e.toString() });
      }
      throw e;
    }
  });
}

function runFullSync() {
  return _loggedCall_("runFullSync", function() {
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Запуск полной синхронизации прайса", "INFO", "", "runFullSync", "Синхронизация всех строк листа", "Sync", "START", {}, null);
    }
    try {
      if (typeof Lib !== 'undefined' && Lib.runFullSync) {
        var result = Lib.runFullSync();
        if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
          Lib.logWithEmoji("Полная синхронизация завершена", "INFO", "", "runFullSync", "Весь прайс синхронизирован", "Sync", "SUCCESS", {}, { result: result });
        }
        return result;
      }
      throw new Error('Lib.runFullSync не определена');
    } catch (e) {
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Ошибка полной синхронизации", "ERROR", "", "runFullSync", e.message, "Sync", "ERROR", {}, { error: e.toString() });
      }
      throw e;
    }
  });
}

function setupTriggers() {
  if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
    Lib.logWithEmoji("Запуск настройки триггеров", "INFO", "", "setupTriggers", "Установка устанавливаемых триггеров", "Settings", "START", {}, null);
  }
  try {
    if (typeof Lib !== 'undefined' && Lib.setupTriggers) {
      var result = Lib.setupTriggers();
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Триггеры успешно настроены", "INFO", "", "setupTriggers", "Все необходимые триггеры установлены", "Settings", "SUCCESS", {}, { result: result });
      }
      return result;
    }
    throw new Error('Lib.setupTriggers не определена');
  } catch (e) {
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Ошибка при настройке триггеров", "ERROR", "", "setupTriggers", e.message, "Settings", "ERROR", {}, { error: e.toString() });
    }
    throw e;
  }
}

/** @deprecated Используйте showSyncRulesManagerDialog */
function showSyncConfigDialog() {
  if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
    Lib.logWithEmoji("Открытие конфигурации синхронизации (legacy)", "INFO", "", "showSyncConfigDialog", "Перенаправление на новый менеджер правил", "Menu", "START", {}, null);
  }
  return showSyncRulesManagerDialog();
}

// Новый UI управления правилами (модалка с CRUD и категориями)
function showSyncRulesManagerDialog() {
  if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
    Lib.logWithEmoji("Открытие менеджера правил синхронизации", "INFO", "", "showSyncRulesManagerDialog", "Запуск UI менеджера правил", "Menu", "START", {}, null);
  }
  try {
    if (typeof Lib !== 'undefined' && typeof Lib.showSyncRulesManagerDialog === 'function') {
      var result = Lib.showSyncRulesManagerDialog();
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Менеджер правил открыт", "INFO", "", "showSyncRulesManagerDialog", "Диалоговое окно успешно отображено", "Menu", "SUCCESS", {}, { status: "opened" });
      }
      return result;
    }
    throw new Error('Lib.showSyncRulesManagerDialog не определена');
  } catch (e) {
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Ошибка при открытии менеджера правил", "ERROR", "", "showSyncRulesManagerDialog", e.message, "Menu", "ERROR", {}, { error: e.toString() });
    }
    throw e;
  }
}

function showExternalDocManagerDialog() {
  if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
    Lib.logWithEmoji("Открытие менеджера внешних документов", "INFO", "", "showExternalDocManagerDialog", "Запуск UI управления внешними связями", "Menu", "START", {}, null);
  }
  try {
    if (typeof Lib !== 'undefined' && Lib.showExternalDocManagerDialog) {
      var result = Lib.showExternalDocManagerDialog();
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Менеджер документов открыт", "INFO", "", "showExternalDocManagerDialog", "Диалоговое окно успешно отображено", "Menu", "SUCCESS", {}, { status: "opened" });
      }
      return result;
    }
    throw new Error('Lib.showExternalDocManagerDialog не определена');
  } catch (e) {
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Ошибка при открытии менеджера документов", "ERROR", "", "showExternalDocManagerDialog", e.message, "Menu", "ERROR", {}, { error: e.toString() });
    }
    throw e;
  }
}

// ============ ТЕСТЕР ============

function runAllTests() {
  if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
    Lib.logWithEmoji("Запуск всех тестов системы", "INFO", "", "runAllTests", "Выполнение набора автоматических тестов", "System", "START", {}, null);
  }
  try {
    if (typeof Lib !== 'undefined' && Lib.runAllTests) {
      var result = Lib.runAllTests();
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Тестирование завершено", "INFO", "", "runAllTests", "Все тесты выполнены", "System", "SUCCESS", {}, { result: result });
      }
      return result;
    }
    throw new Error('Lib.runAllTests не определена');
  } catch (e) {
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Ошибка при запуске тестов", "ERROR", "", "runAllTests", e.message, "System", "ERROR", {}, { error: e.toString() });
    }
    throw e;
  }
}

function clearTestResults() {
  if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
    Lib.logWithEmoji("Очистка результатов тестирования", "INFO", "", "clearTestResults", "Удаление данных последних тестов", "System", "START", {}, null);
  }
  try {
    if (typeof Lib !== 'undefined' && Lib.clearTestResults) {
      var result = Lib.clearTestResults();
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Результаты тестов очищены", "INFO", "", "clearTestResults", "Данные успешно удалены", "System", "SUCCESS", {}, { status: "cleared" });
      }
      return result;
    }
    throw new Error('Lib.clearTestResults не определена');
  } catch (e) {
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Ошибка при очистке результатов тестов", "ERROR", "", "clearTestResults", e.message, "System", "ERROR", {}, { error: e.toString() });
    }
    throw e;
  }
}

// ============ ЛОГИ ============

function refreshLogs() {
  if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
    Lib.logWithEmoji("Запуск обновления структуры логов", "INFO", "", "refreshLogs", "Обновление листа логирования", "System", "START", {}, null);
  }
  try {
    if (typeof Lib !== 'undefined' && Lib.refreshLogs) {
      var result = Lib.refreshLogs();
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Структура логов обновлена", "INFO", "", "refreshLogs", "Лист логов синхронизирован с сервером", "System", "SUCCESS", {}, { result: result });
      }
      return result;
    }
    throw new Error('Lib.refreshLogs не определена');
  } catch (e) {
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Ошибка при обновлении логов", "ERROR", "", "refreshLogs", e.message, "System", "ERROR", {}, { error: e.toString() });
    }
    throw e;
  }
}

// ОТКЛЮЧЕНО: Функции управления логами через лист больше не нужны.
// Все логи теперь ведутся на сервере и доступны через "📜 Открыть Журнал (UI)".

function quickCleanLogSheet() { console.warn('quickCleanLogSheet отключена'); }
function recreateLogSheet() { console.warn('recreateLogSheet отключена'); }
function recreateDebugLogSheet() { console.warn('recreateDebugLogSheet отключена'); }

function serverManualArchiveLogs() { console.warn('serverManualArchiveLogs отключена'); }
function serverResetLogSheet() { console.warn('serverResetLogSheet отключена'); }
function serverMidnightLogRotation() { console.warn('serverMidnightLogRotation отключена'); }
function serverGetLogStatus() { console.warn('serverGetLogStatus отключена'); }

// ============ ОБРАБОТКА ПРАЙСОВ (SK) ============

function processSkPriceSheet() {
  return _loggedCall_("processSkPriceSheet", function() {
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Запуск обработки прайса SK", "INFO", "", "processSkPrice", "Пересчет и обновление листа SK", "Calculations", "START", { project: "SK" }, null);
    }
    try {
      if (typeof Lib !== 'undefined' && Lib.processSkPriceSheet) {
        var result = Lib.processSkPriceSheet();
        if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
          Lib.logWithEmoji("Обработка SK завершена", "INFO", "", "processSkPrice", "Номенклатура SK обновлена", "Calculations", "SUCCESS", { project: "SK" }, { result: result });
        }
        return result;
      }
      throw new Error('Lib.processSkPriceSheet не определена');
    } catch (e) {
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Ошибка при обработке SK", "ERROR", "", "processSkPrice", e.message, "Calculations", "ERROR", { project: "SK" }, { error: e.toString() });
      }
      throw e;
    }
  });
}

function loadSkStockData() {
  return _loggedCall_("loadSkStockData", function() {
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Загрузка остатков SK", "INFO", "", "loadSkStock", "Обновление складских запасов SK", "Calculations", "START", { project: "SK" }, null);
    }
    try {
      if (typeof Lib !== 'undefined' && Lib.loadSkStockData) {
        var result = Lib.loadSkStockData();
        if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
          Lib.logWithEmoji("Остатки SK загружены", "INFO", "", "loadSkStock", "Данные склада SK обновлены", "Calculations", "SUCCESS", { project: "SK" }, { result: result });
        }
        return result;
      }
      throw new Error('Lib.loadSkStockData не определена');
    } catch (e) {
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Ошибка при загрузке остатков SK", "ERROR", "", "loadSkStock", e.message, "Calculations", "ERROR", { project: "SK" }, { error: e.toString() });
      }
      throw e;
    }
  });
}

// ============ ОБРАБОТКА ПРАЙСОВ (MT) ============

function processMtMainPrice() {
  return _loggedCall_("processMtMainPrice", function() {
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Запуск обработки основного прайса MT", "INFO", "", "processMtMain", "Пересчет и обновление листа MT Main", "Calculations", "START", { project: "MT" }, null);
    }
    try {
      if (typeof Lib !== 'undefined' && Lib.processMtMainPrice) {
        var result = Lib.processMtMainPrice();
        if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
          Lib.logWithEmoji("Обработка MT Main завершена", "INFO", "", "processMtMain", "Номенклатура MT Main обновлена", "Calculations", "SUCCESS", { project: "MT" }, { result: result });
        }
        return result;
      }
      throw new Error('Lib.processMtMainPrice не определена');
    } catch (e) {
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Ошибка при обработке MT Main", "ERROR", "", "processMtMain", e.message, "Calculations", "ERROR", { project: "MT" }, { error: e.toString() });
      }
      throw e;
    }
  });
}

function processMtTesterPrice() {
  return _loggedCall_("processMtTesterPrice", function() {
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Запуск обработки тестеров MT", "INFO", "", "processMtTester", "Пересчет и обновление листа MT Tester", "Calculations", "START", { project: "MT" }, null);
    }
    try {
      if (typeof Lib !== 'undefined' && Lib.processMtTesterPrice) {
        var result = Lib.processMtTesterPrice();
        if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
          Lib.logWithEmoji("Обработка MT Tester завершена", "INFO", "", "processMtTester", "Номенклатура MT Tester обновлена", "Calculations", "SUCCESS", { project: "MT" }, { result: result });
        }
        return result;
      }
      throw new Error('Lib.processMtTesterPrice не определена');
    } catch (e) {
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Ошибка при обработке MT Tester", "ERROR", "", "processMtTester", e.message, "Calculations", "ERROR", { project: "MT" }, { error: e.toString() });
      }
      throw e;
    }
  });
}

function processMtSamplesPrice() {
  return _loggedCall_("processMtSamplesPrice", function() {
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Запуск обработки пробников MT", "INFO", "", "processMtSamples", "Пересчет и обновление листа MT Samples", "Calculations", "START", { project: "MT" }, null);
    }
    try {
      if (typeof Lib !== 'undefined' && Lib.processMtSamplesPrice) {
        var result = Lib.processMtSamplesPrice();
        if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
          Lib.logWithEmoji("Обработка MT Samples завершена", "INFO", "", "processMtSamples", "Номенклатура MT Samples обновлена", "Calculations", "SUCCESS", { project: "MT" }, { result: result });
        }
        return result;
      }
      throw new Error('Lib.processMtSamplesPrice не определена');
    } catch (e) {
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Ошибка при обработке MT Samples", "ERROR", "", "processMtSamples", e.message, "Calculations", "ERROR", { project: "MT" }, { error: e.toString() });
      }
      throw e;
    }
  });
}

function loadMtStockData() {
  return _loggedCall_("loadMtStockData", function() {
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Загрузка остатков MT", "INFO", "", "loadMtStock", "Обновление складских запасов MT", "Calculations", "START", { project: "MT" }, null);
    }
    try {
      if (typeof Lib !== 'undefined' && Lib.loadMtStockData) {
        var result = Lib.loadMtStockData();
        if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
          Lib.logWithEmoji("Остатки MT загружены", "INFO", "", "loadMtStock", "Данные склада MT обновлены", "Calculations", "SUCCESS", { project: "MT" }, { result: result });
        }
        return result;
      }
      throw new Error('Lib.loadMtStockData не определена');
    } catch (e) {
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Ошибка при загрузке остатков MT", "ERROR", "", "loadMtStock", e.message, "Calculations", "ERROR", { project: "MT" }, { error: e.toString() });
      }
      throw e;
    }
  });
}

// ============ ОБРАБОТКА ПРАЙСОВ (SS) ============

function processSsPriceSheet() {
  return _loggedCall_("processSsPriceSheet", function() {
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Запуск обработки прайса SS", "INFO", "", "processSsPrice", "Пересчет и обновление листа SS", "Calculations", "START", { project: "SS" }, null);
    }
    try {
      if (typeof Lib !== 'undefined' && Lib.processSsPriceSheet) {
        var result = Lib.processSsPriceSheet();
        if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
          Lib.logWithEmoji("Обработка SS завершена", "INFO", "", "processSsPrice", "Номенклатура SS обновлена", "Calculations", "SUCCESS", { project: "SS" }, { result: result });
        }
        return result;
      }
      throw new Error('Lib.processSsPriceSheet не определена');
    } catch (e) {
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Ошибка при обработке SS", "ERROR", "", "processSsPrice", e.message, "Calculations", "ERROR", { project: "SS" }, { error: e.toString() });
      }
      throw e;
    }
  });
}

function loadSsStockData() {
  return _loggedCall_("loadSsStockData", function() {
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Загрузка остатков SS", "INFO", "", "loadSsStock", "Обновление складских запасов SS", "Calculations", "START", { project: "SS" }, null);
    }
    try {
      if (typeof Lib !== 'undefined' && Lib.loadSsStockData) {
        var result = Lib.loadSsStockData();
        if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
          Lib.logWithEmoji("Остатки SS загружены", "INFO", "", "loadSsStock", "Данные склада SS обновлены", "Calculations", "SUCCESS", { project: "SS" }, { result: result });
        }
        return result;
      }
      throw new Error('Lib.loadSsStockData не определена');
    } catch (e) {
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Ошибка при загрузке остатков SS", "ERROR", "", "loadSsStock", e.message, "Calculations", "ERROR", { project: "SS" }, { error: e.toString() });
      }
      throw e;
    }
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
  if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
    Lib.logWithEmoji("Сортировка заказа SK по производителю", "INFO", "", "sortSkOrder", "Запуск серверной сортировки SK", "Dashboard", "START", { project: "SK", sortBy: "manufacturer" }, null);
  }
  try {
    if (typeof callServerStructureSort === 'function') {
      var result = callServerStructureSort('byManufacturer');
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Сортировка SK завершена", "INFO", "", "sortSkOrder", "Заказ SK отсортирован", "Dashboard", "SUCCESS", { project: "SK" }, { result: result });
      }
      return result;
    }
    throw new Error('callServerStructureSort не определена');
  } catch (e) {
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Ошибка при сортировке SK", "ERROR", "", "sortSkOrder", e.message, "Dashboard", "ERROR", { project: "SK" }, { error: e.toString() });
    }
    throw e;
  }
}

function sortSkOrderByPrice() {
  if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
    Lib.logWithEmoji("Сортировка заказа SK по цене", "INFO", "", "sortSkOrder", "Запуск серверной сортировки SK", "Dashboard", "START", { project: "SK", sortBy: "price" }, null);
  }
  try {
    if (typeof callServerStructureSort === 'function') {
      var result = callServerStructureSort('byPrice');
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Сортировка SK завершена", "INFO", "", "sortSkOrder", "Заказ SK отсортирован", "Dashboard", "SUCCESS", { project: "SK" }, { result: result });
      }
      return result;
    }
    throw new Error('callServerStructureSort не определена');
  } catch (e) {
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Ошибка при сортировке SK", "ERROR", "", "sortSkOrder", e.message, "Dashboard", "ERROR", { project: "SK" }, { error: e.toString() });
    }
    throw e;
  }
}

function sortMtOrderByManufacturer() {
  if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
    Lib.logWithEmoji("Сортировка заказа MT по производителю", "INFO", "", "sortMtOrder", "Запуск серверной сортировки MT", "Dashboard", "START", { project: "MT", sortBy: "manufacturer" }, null);
  }
  try {
    if (typeof callServerStructureSort === 'function') {
      var result = callServerStructureSort('byManufacturer');
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Сортировка MT завершена", "INFO", "", "sortMtOrder", "Заказ MT отсортирован", "Dashboard", "SUCCESS", { project: "MT" }, { result: result });
      }
      return result;
    }
    throw new Error('callServerStructureSort не определена');
  } catch (e) {
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Ошибка при сортировке MT", "ERROR", "", "sortMtOrder", e.message, "Dashboard", "ERROR", { project: "MT" }, { error: e.toString() });
    }
    throw e;
  }
}

function sortMtOrderByPrice() {
  if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
    Lib.logWithEmoji("Сортировка заказа MT по цене", "INFO", "", "sortMtOrder", "Запуск серверной сортировки MT", "Dashboard", "START", { project: "MT", sortBy: "price" }, null);
  }
  try {
    if (typeof callServerStructureSort === 'function') {
      var result = callServerStructureSort('byPrice');
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Сортировка MT завершена", "INFO", "", "sortMtOrder", "Заказ MT отсортирован", "Dashboard", "SUCCESS", { project: "MT" }, { result: result });
      }
      return result;
    }
    throw new Error('callServerStructureSort не определена');
  } catch (e) {
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Ошибка при сортировке MT", "ERROR", "", "sortMtOrder", e.message, "Dashboard", "ERROR", { project: "MT" }, { error: e.toString() });
    }
    throw e;
  }
}

function sortSsOrderByManufacturer() {
  if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
    Lib.logWithEmoji("Сортировка заказа SS по производителю", "INFO", "", "sortSsOrder", "Запуск серверной сортировки SS", "Dashboard", "START", { project: "SS", sortBy: "manufacturer" }, null);
  }
  try {
    if (typeof callServerStructureSort === 'function') {
      var result = callServerStructureSort('byManufacturer');
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Сортировка SS завершена", "INFO", "", "sortSsOrder", "Заказ SS отсортирован", "Dashboard", "SUCCESS", { project: "SS" }, { result: result });
      }
      return result;
    }
    throw new Error('callServerStructureSort не определена');
  } catch (e) {
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Ошибка при сортировке SS", "ERROR", "", "sortSsOrder", e.message, "Dashboard", "ERROR", { project: "SS" }, { error: e.toString() });
    }
    throw e;
  }
}

function sortSsOrderByPrice() {
  if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
    Lib.logWithEmoji("Сортировка заказа SS по цене", "INFO", "", "sortSsOrder", "Запуск серверной сортировки SS", "Dashboard", "START", { project: "SS", sortBy: "price" }, null);
  }
  try {
    if (typeof callServerStructureSort === 'function') {
      var result = callServerStructureSort('byPrice');
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Сортировка SS завершена", "INFO", "", "sortSsOrder", "Заказ SS отсортирован", "Dashboard", "SUCCESS", { project: "SS" }, { result: result });
      }
      return result;
    }
    throw new Error('callServerStructureSort не определена');
  } catch (e) {
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Ошибка при сортировке SS", "ERROR", "", "sortSsOrder", e.message, "Dashboard", "ERROR", { project: "SS" }, { error: e.toString() });
    }
    throw e;
  }
}

function showAllOrderData() {
  if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
    Lib.logWithEmoji("Запуск отображения всех данных заказа", "INFO", "", "showAllOrderData", "Загрузка всех строк листа Заказ", "Dashboard", "START", {}, null);
  }
  if (typeof Lib !== 'undefined' && Lib.showAllOrderData) {
    var result = Lib.showAllOrderData();
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Все данные заказа отображены", "INFO", "", "showAllOrderData", "Фильтры сброшены", "Dashboard", "SUCCESS", {}, { status: "displayed" });
    }
    return result;
  }
  throw new Error('Lib.showAllOrderData не определена');
}

function showOrderStage() {
  if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
    Lib.logWithEmoji("Запрос отображения стадии 'Заказ'", "INFO", "", "showOrderStage", "Фильтрация строк по стадии Заказ", "Dashboard", "START", { stage: "order" }, null);
  }
  if (typeof Lib !== 'undefined' && Lib.showOrderStage) {
    var result = Lib.showOrderStage();
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Стадия 'Заказ' отображена", "INFO", "", "showOrderStage", "Фильтр 'Заказ' применен", "Dashboard", "SUCCESS", { stage: "order" }, { status: "displayed" });
    }
    return result;
  }
  throw new Error('Lib.showOrderStage не определена');
}

function showPromotionsStage() {
  if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
    Lib.logWithEmoji("Запрос отображения стадии 'Акции'", "INFO", "", "showPromotionsStage", "Фильтрация строк по стадии Акции", "Dashboard", "START", { stage: "promotions" }, null);
  }
  if (typeof Lib !== 'undefined' && Lib.showPromotionsStage) {
    var result = Lib.showPromotionsStage();
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Стадия 'Акции' отображена", "INFO", "", "showPromotionsStage", "Фильтр 'Акции' применен", "Dashboard", "SUCCESS", { stage: "promotions" }, { status: "displayed" });
    }
    return result;
  }
  throw new Error('Lib.showPromotionsStage не определена');
}

function showSetStage() {
  if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
    Lib.logWithEmoji("Запрос отображения стадии 'Набор'", "INFO", "", "showSetStage", "Фильтрация строк по стадии Набор", "Dashboard", "START", { stage: "set" }, null);
  }
  if (typeof Lib !== 'undefined' && Lib.showSetStage) {
    var result = Lib.showSetStage();
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Стадия 'Набор' отображена", "INFO", "", "showSetStage", "Фильтр 'Набор' применен", "Dashboard", "SUCCESS", { stage: "set" }, { status: "displayed" });
    }
    return result;
  }
  throw new Error('Lib.showSetStage не определена');
}

function showPriceStage() {
  if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
    Lib.logWithEmoji("Запрос отображения стадии 'Прайс'", "INFO", "", "showPriceStage", "Фильтрация строк по стадии Прайс", "Dashboard", "START", { stage: "price" }, null);
  }
  if (typeof Lib !== 'undefined' && Lib.showPriceStage) {
    var result = Lib.showPriceStage();
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Стадия 'Прайс' отображена", "INFO", "", "showPriceStage", "Фильтр 'Прайс' применен", "Dashboard", "SUCCESS", { stage: "price" }, { status: "displayed" });
    }
    return result;
  }
  throw new Error('Lib.showPriceStage не определена');
}

// ============ ВЫГРУЗКА ============

function exportPromotions() {
  if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
    Lib.logWithEmoji("Запуск экспорта акций", "INFO", "", "exportPromotions", "Выгрузка данных по акциям", "Dashboard", "START", {}, null);
  }
  try {
    if (typeof Lib !== 'undefined' && Lib.exportPromotions) {
      var result = Lib.exportPromotions();
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Экспорт акций завершен", "INFO", "", "exportPromotions", "Данные выгружены", "Dashboard", "SUCCESS", {}, { result: result });
      }
      return result;
    }
    throw new Error('Lib.exportPromotions не определена');
  } catch (e) {
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Ошибка при экспорте акций", "ERROR", "", "exportPromotions", e.message, "Dashboard", "ERROR", {}, { error: e.toString() });
    }
    throw e;
  }
}

function exportSets() {
  if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
    Lib.logWithEmoji("Запуск экспорта наборов", "INFO", "", "exportSets", "Выгрузка данных по наборам", "Dashboard", "START", {}, null);
  }
  try {
    if (typeof Lib !== 'undefined' && Lib.exportSets) {
      var result = Lib.exportSets();
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Экспорт наборов завершен", "INFO", "", "exportSets", "Данные выгружены", "Dashboard", "SUCCESS", {}, { result: result });
      }
      return result;
    }
    throw new Error('Lib.exportSets не определена');
  } catch (e) {
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Ошибка при экспорте наборов", "ERROR", "", "exportSets", e.message, "Dashboard", "ERROR", {}, { error: e.toString() });
    }
    throw e;
  }
}

// ============ ПОСТАВКА ============

function formatOrderSheet() {
  return _loggedCall_("formatOrderSheet", function() {
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Форматирование листа заказа", "INFO", "", "formatOrderSheet", "Применение стилей и разметки", "Dashboard", "START", {}, null);
    }
    try {
      if (typeof Lib !== 'undefined' && Lib.formatOrderSheet) {
        var result = Lib.formatOrderSheet();
        if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
          Lib.logWithEmoji("Форматирование завершено", "INFO", "", "formatOrderSheet", "Лист заказа оформлен", "Dashboard", "SUCCESS", {}, { result: result });
        }
        return result;
      }
      throw new Error('Lib.formatOrderSheet не определена');
    } catch (e) {
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Ошибка форматирования", "ERROR", "", "formatOrderSheet", e.message, "Dashboard", "ERROR", {}, { error: e.toString() });
      }
      throw e;
    }
  });
}

function createFullInvoice() {
  return _loggedCall_("createFullInvoice", function() {
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Генерация полного инвойса", "INFO", "", "createFullInvoice", "Сбор данных для инвойса", "Dashboard", "START", {}, null);
    }
    try {
      if (typeof Lib !== 'undefined' && Lib.createFullInvoice) {
        var result = Lib.createFullInvoice();
        if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
          Lib.logWithEmoji("Инвойс сгенерирован", "INFO", "", "createFullInvoice", "Файл инвойса создан", "Dashboard", "SUCCESS", {}, { result: result });
        }
        return result;
      }
      throw new Error('Lib.createFullInvoice не определена');
    } catch (e) {
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Ошибка при генерации инвойса", "ERROR", "", "createFullInvoice", e.message, "Dashboard", "ERROR", {}, { error: e.toString() });
      }
      throw e;
    }
  });
}

function collectAndCopyDocuments() {
  return _loggedCall_("collectAndCopyDocuments", function() {
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Сбор и копирование документов", "INFO", "", "collectDocuments", "Загрузка документов через сервер", "Dashboard", "START", {}, null);
    }
    try {
      if (typeof callServerCollectDocuments === 'function') {
        var result = callServerCollectDocuments();
        if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
          Lib.logWithEmoji("Документы собраны", "INFO", "", "collectDocuments", "Копирование завершено", "Dashboard", "SUCCESS", {}, { result: result });
        }
        return result;
      }
      throw new Error('callServerCollectDocuments не определена');
    } catch (e) {
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Ошибка при сборе документов", "ERROR", "", "collectDocuments", e.message, "Dashboard", "ERROR", {}, { error: e.toString() });
      }
      throw e;
    }
  });
}

// ============ СЕРТИФИКАЦИЯ ============

function createNewsSheetFromCertification() {
  return _loggedCall_("createNewsSheetFromCertification", function() {
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Создание листа новинок из сертификации", "INFO", "", "createNewsSheet", "Перенос данных новинок в отдельный лист", "Certification", "START", {}, null);
    }
    try {
      if (typeof Lib !== 'undefined' && Lib.createNewsSheetFromCertification) {
        var result = Lib.createNewsSheetFromCertification();
        if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
          Lib.logWithEmoji("Лист новинок создан", "INFO", "", "createNewsSheet", "Данные успешно перенесены", "Certification", "SUCCESS", {}, { result: result });
        }
        return result;
      }
      throw new Error('Lib.createNewsSheetFromCertification не определена');
    } catch (e) {
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Ошибка при создании листа новинок", "ERROR", "", "createNewsSheet", e.message, "Certification", "ERROR", {}, { error: e.toString() });
      }
      throw e;
    }
  });
}

function generateProtocols_353pp() {
  return _loggedCall_("generateProtocols_353pp", function() {
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Генерация протоколов 353пп", "INFO", "", "generateProtocols", "Создание документов протоколов", "Certification", "START", {}, null);
    }
    try {
      if (typeof Lib !== 'undefined' && Lib.generateProtocols_353pp) {
        var result = Lib.generateProtocols_353pp();
        if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
          Lib.logWithEmoji("Протоколы 353пп сгенерированы", "INFO", "", "generateProtocols", "Файлы протоколов созданы", "Certification", "SUCCESS", {}, { result: result });
        }
        return result;
      }
      throw new Error('Lib.generateProtocols_353pp не определена');
    } catch (e) {
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Ошибка при генерации протоколов", "ERROR", "", "generateProtocols", e.message, "Certification", "ERROR", {}, { error: e.toString() });
      }
      throw e;
    }
  });
}

function generateDsLayouts_353pp() {
  return _loggedCall_("generateDsLayouts_353pp", function() {
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Генерация макетов ДС 353пп", "INFO", "", "generateDsLayouts", "Создание файлов макетов", "Certification", "START", {}, null);
    }
    try {
      if (typeof Lib !== 'undefined' && Lib.generateDsLayouts_353pp) {
        var result = Lib.generateDsLayouts_353pp();
        if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
          Lib.logWithEmoji("Макеты ДС 353пп сгенерированы", "INFO", "", "generateDsLayouts", "Файлы макетов созданы", "Certification", "SUCCESS", {}, { result: result });
        }
        return result;
      }
      throw new Error('Lib.generateDsLayouts_353pp не определена');
    } catch (e) {
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Ошибка при генерации макетов", "ERROR", "", "generateDsLayouts", e.message, "Certification", "ERROR", {}, { error: e.toString() });
      }
      throw e;
    }
  });
}

function structureDocuments_353pp() {
  return _loggedCall_("structureDocuments_353pp", function() {
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Структурирование документов 353пп", "INFO", "", "structureDocs", "Организация файлов в Drive", "Certification", "START", {}, null);
    }
    try {
      if (typeof callServerStructureDocuments353pp === 'function') {
        var result = callServerStructureDocuments353pp();
        if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
          Lib.logWithEmoji("Документы 353пп структурированы", "INFO", "", "structureDocs", "Файлы перемещены на сервере", "Certification", "SUCCESS", {}, { result: result });
        }
        return result;
      }
      if (typeof Lib !== 'undefined' && Lib.structureDocuments_353pp) {
        var resultLegacy = Lib.structureDocuments_353pp();
        if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
          Lib.logWithEmoji("Документы 353пп структурированы (Legacy)", "INFO", "", "structureDocs", "Файлы перемещены через GAS", "Certification", "SUCCESS", {}, { result: resultLegacy });
        }
        return resultLegacy;
      }
      throw new Error('callServerStructureDocuments353pp не определена');
    } catch (e) {
       if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Ошибка при структурировании 353пп", "ERROR", "", "structureDocs", e.message, "Certification", "ERROR", {}, { error: e.toString() });
      }
      throw e;
    }
  });
}

function calculateAndAssignSpiritNumbers() {
  return _loggedCall_("calculateAndAssignSpiritNumbers", function() {
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Расчет номеров спирта", "INFO", "", "calcSpirit", "Присвоение номеров спиртосодержащей продукции", "Certification", "START", {}, null);
    }
    try {
      if (typeof Lib !== 'undefined' && Lib.calculateAndAssignSpiritNumbers) {
        var result = Lib.calculateAndAssignSpiritNumbers();
        if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
          Lib.logWithEmoji("Расчет спирта завершен", "INFO", "", "calcSpirit", "Номера успешно присвоены", "Certification", "SUCCESS", {}, { result: result });
        }
        return result;
      }
      throw new Error('Lib.calculateAndAssignSpiritNumbers не определена');
    } catch (e) {
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Ошибка при расчете спирта", "ERROR", "", "calcSpirit", e.message, "Certification", "ERROR", {}, { error: e.toString() });
      }
      throw e;
    }
  });
}

function generateSpiritProtocols() {
  return _loggedCall_("generateSpiritProtocols", function() {
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Генерация спиртовых протоколов", "INFO", "", "genSpiritProtocols", "Создание документов для спирта", "Certification", "START", {}, null);
    }
    try {
      if (typeof Lib !== 'undefined' && Lib.generateSpiritProtocols) {
        var result = Lib.generateSpiritProtocols();
        if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
          Lib.logWithEmoji("Спиртовые протоколы сгенерированы", "INFO", "", "genSpiritProtocols", "Файлы созданы", "Certification", "SUCCESS", {}, { result: result });
        }
        return result;
      }
      throw new Error('Lib.generateSpiritProtocols не определена');
    } catch (e) {
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Ошибка генерации спиртовых протоколов", "ERROR", "", "genSpiritProtocols", e.message, "Certification", "ERROR", {}, { error: e.toString() });
      }
      throw e;
    }
  });
}

function runManualCascadeOnCertification() {
  return _loggedCall_("runManualCascadeOnCertification", function() {
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Запуск каскадного обновления сертификации", "INFO", "", "runCascade", "Принудительное обновление зависимых полей", "Certification", "START", {}, null);
    }
    try {
      if (typeof Lib !== 'undefined' && Lib.runManualCascadeOnCertification) {
        var result = Lib.runManualCascadeOnCertification();
        if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
          Lib.logWithEmoji("Каскадное обновление завершено", "INFO", "", "runCascade", "Сертификация обновлена", "Certification", "SUCCESS", {}, { result: result });
        }
        return result;
      }
      throw new Error('Lib.runManualCascadeOnCertification не определена');
    } catch (e) {
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Ошибка при каскадном обновлении", "ERROR", "", "runCascade", e.message, "Certification", "ERROR", {}, { error: e.toString() });
      }
      throw e;
    }
  });
}

// ============ DRIVE ============

function uploadFilesToFolder() {
  if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
    Lib.logWithEmoji("Запуск загрузки файлов в Google Drive", "INFO", "", "uploadFilesToFolder", "Выбор и загрузка файлов в папку проекта", "Drive", "START", {}, null);
  }
  try {
    if (typeof Lib !== 'undefined' && Lib.uploadFilesToFolder) {
      var result = Lib.uploadFilesToFolder();
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Загрузка файлов завершена", "INFO", "", "uploadFilesToFolder", "Файлы успешно обработаны", "Drive", "SUCCESS", {}, { result: result });
      }
      return result;
    }
    throw new Error('Lib.uploadFilesToFolder не определена');
  } catch (e) {
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Ошибка при загрузке файлов", "ERROR", "", "uploadFilesToFolder", e.message, "Drive", "ERROR", {}, { error: e.toString() });
    }
    throw e;
  }
}

function createFolderStructure() {
  if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
    Lib.logWithEmoji("Создание структуры папок Drive", "INFO", "", "createFolderStructure", "Инициализация иерархии папок проекта", "Drive", "START", {}, null);
  }
  try {
    if (typeof Lib !== 'undefined' && Lib.createFolderStructure) {
      var result = Lib.createFolderStructure();
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Структура папок создана", "INFO", "", "createFolderStructure", "Папки проекта успешно инициализированы", "Drive", "SUCCESS", {}, { result: result });
      }
      return result;
    }
    throw new Error('Lib.createFolderStructure не определена');
  } catch (e) {
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Ошибка при создании структуры папок", "ERROR", "", "createFolderStructure", e.message, "Drive", "ERROR", {}, { error: e.toString() });
    }
    throw e;
  }
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
/**
 * Загрузить остатки (вызывает серверную логику)
 */
function serverLoadStockData() {
  return _loggedCall_("serverLoadStockData", function() {
    if (typeof callServerLoadStocks === 'function') {
      return callServerLoadStocks();
    }
    // Фолбэк на старую логику, если серверная не готова
    var projectKey = (typeof CONFIG !== 'undefined' && CONFIG.ACTIVE_PROJECT_KEY) ? CONFIG.ACTIVE_PROJECT_KEY : 'MT';
    if (Lib.loadStockData) {
      return Lib.loadStockData(projectKey);
    }
    throw new Error('Логика загрузки остатков не найдена');
  });
}

/**
 * Загрузить и обработать прайс (вызывает серверную логику)
 */
function serverProcessPrimaryData() {
  return _loggedCall_("serverProcessPrimaryData", function() {
    var projectKey = (typeof CONFIG !== 'undefined' && CONFIG.ACTIVE_PROJECT_KEY) ? CONFIG.ACTIVE_PROJECT_KEY : 'MT';
    var mode = 'all'; // Запускает ВСЕ циклы обработки последовательно (main → tester → samples и т.д.)

    if (typeof callServerProcessPrice === 'function') {
      return callServerProcessPrice(projectKey.toLowerCase(), mode);
    }
    throw new Error('Функция callServerProcessPrice не найдена');
  });
}

function serverProcessMtMain() {
  return _loggedCall_("serverProcessMtMain", function() {
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Запуск серверной обработки MT Main", "INFO", "", "serverProcessMtMain", "Вызов Python API для MT Main", "Calculations", "START", { project: "MT", mode: "main" }, null);
    }
    try {
      if (typeof callServerProcessPrice === 'function') {
        var result = callServerProcessPrice('mt', 'main');
        if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
          Lib.logWithEmoji("Серверная обработка MT Main завершена", "INFO", "", "serverProcessMtMain", "API выполнено успешно", "Calculations", "SUCCESS", {}, { result: result });
        }
        return result;
      }
      throw new Error('callServerProcessPrice не определена');
    } catch (e) {
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Ошибка API MT Main", "ERROR", "", "serverProcessMtMain", e.message, "Calculations", "ERROR", {}, { error: e.toString() });
      }
      throw e;
    }
  });
}

/**
 * Обработка тестеров MT через сервер
 */
function serverProcessMtTester() {
  return _loggedCall_("serverProcessMtTester", function() {
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Запуск серверной обработки MT Tester", "INFO", "", "serverProcessMtTester", "Вызов Python API для MT Tester", "Calculations", "START", { project: "MT", mode: "tester" }, null);
    }
    try {
      if (typeof callServerProcessPrice === 'function') {
        var result = callServerProcessPrice('mt', 'tester');
        if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
          Lib.logWithEmoji("Серверная обработка MT Tester завершена", "INFO", "", "serverProcessMtTester", "API выполнено успешно", "Calculations", "SUCCESS", {}, { result: result });
        }
        return result;
      }
      throw new Error('callServerProcessPrice не определена');
    } catch (e) {
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Ошибка API MT Tester", "ERROR", "", "serverProcessMtTester", e.message, "Calculations", "ERROR", {}, { error: e.toString() });
      }
      throw e;
    }
  });
}

/**
 * Обработка пробников MT через сервер
 */
function serverProcessMtSamples() {
  return _loggedCall_("serverProcessMtSamples", function() {
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Запуск серверной обработки MT Samples", "INFO", "", "serverProcessMtSamples", "Вызов Python API для MT Samples", "Calculations", "START", { project: "MT", mode: "samples" }, null);
    }
    try {
      if (typeof callServerProcessPrice === 'function') {
        var result = callServerProcessPrice('mt', 'samples');
        if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
          Lib.logWithEmoji("Серверная обработка MT Samples завершена", "INFO", "", "serverProcessMtSamples", "API выполнено успешно", "Calculations", "SUCCESS", {}, { result: result });
        }
        return result;
      }
      throw new Error('callServerProcessPrice не определена');
    } catch (e) {
       if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Ошибка API MT Samples", "ERROR", "", "serverProcessMtSamples", e.message, "Calculations", "ERROR", {}, { error: e.toString() });
      }
      throw e;
    }
  });
}

/**
 * Обработка прайса SK через сервер
 */
function serverProcessSkMain() {
  return _loggedCall_("serverProcessSkMain", function() {
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Запуск серверной обработки SK Main", "INFO", "", "serverProcessSkMain", "Вызов Python API для SK Main", "Calculations", "START", { project: "SK", mode: "main" }, null);
    }
    try {
      if (typeof callServerProcessPrice === 'function') {
        var result = callServerProcessPrice('sk', 'main');
        if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
          Lib.logWithEmoji("Серверная обработка SK Main завершена", "INFO", "", "serverProcessSkMain", "API выполнено успешно", "Calculations", "SUCCESS", {}, { result: result });
        }
        return result;
      }
      throw new Error('callServerProcessPrice не определена');
    } catch (e) {
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Ошибка API SK Main", "ERROR", "", "serverProcessSkMain", e.message, "Calculations", "ERROR", {}, { error: e.toString() });
      }
      throw e;
    }
  });
}

/**
 * Обработка пробников SK через сервер
 */
function serverProcessSkProbes() {
  return _loggedCall_("serverProcessSkProbes", function() {
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Запуск серверной обработки SK Probes", "INFO", "", "serverProcessSkProbes", "Вызов Python API для SK Probes", "Calculations", "START", { project: "SK", mode: "probes" }, null);
    }
    try {
      if (typeof callServerProcessPrice === 'function') {
        var result = callServerProcessPrice('sk', 'probes');
        if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
          Lib.logWithEmoji("Серверная обработка SK Probes завершена", "INFO", "", "serverProcessSkProbes", "API выполнено успешно", "Calculations", "SUCCESS", {}, { result: result });
        }
        return result;
      }
      throw new Error('callServerProcessPrice не определена');
    } catch (e) {
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Ошибка API SK Probes", "ERROR", "", "serverProcessSkProbes", e.message, "Calculations", "ERROR", {}, { error: e.toString() });
      }
      throw e;
    }
  });
}

/**
 * Обработка прайса SS через сервер
 */
function serverProcessSsMain() {
  return _loggedCall_("serverProcessSsMain", function() {
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Запуск серверной обработки SS Main", "INFO", "", "serverProcessSsMain", "Вызов Python API для SS Main", "Calculations", "START", { project: "SS", mode: "main" }, null);
    }
    try {
      if (typeof callServerProcessPrice === 'function') {
        var result = callServerProcessPrice('ss', 'main');
        if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
          Lib.logWithEmoji("Серверная обработка SS Main завершена", "INFO", "", "serverProcessSsMain", "API выполнено успешно", "Calculations", "SUCCESS", {}, { result: result });
        }
        return result;
      }
      throw new Error('callServerProcessPrice не определена');
    } catch (e) {
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Ошибка API SS Main", "ERROR", "", "serverProcessSsMain", e.message, "Calculations", "ERROR", {}, { error: e.toString() });
      }
      throw e;
    }
  });
}

/**
 * Preview обработки прайса (dry run)
 * @param {string} project - Код проекта (mt, sk, ss)
 * @param {string} mode - Режим (main, tester, samples, probes)
 */
function serverProcessPricePreview(project, mode) {
  return _loggedCall_("serverProcessPricePreview", function() {
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Запуск предосмотра обработки прайса", "INFO", "", "serverProcessPreview", "Вызов Python API для Dry Run", "Calculations", "START", { project: project, mode: mode }, null);
    }
    try {
      if (typeof callServerProcessPrice === 'function') {
        var result = callServerProcessPrice(project, mode, { dryRun: true });
        if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
          Lib.logWithEmoji("Предосмотр завершен", "INFO", "", "serverProcessPreview", "API выполнено успешно", "Calculations", "SUCCESS", {}, { result: result });
        }
        return result;
      }
      throw new Error('callServerProcessPrice не определена');
    } catch (e) {
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Ошибка API Preview", "ERROR", "", "serverProcessPreview", e.message, "Calculations", "ERROR", {}, { error: e.toString() });
      }
      throw e;
    }
  });
}

function serverSmartMatch(productName) {
  return _loggedCall_("serverSmartMatch", function() {
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Запуск Smart Match", "INFO", "", "smartMatch", "Поиск совпадений для: " + productName, "Calculations", "START", { name: productName }, null);
    }
    try {
      if (typeof callServerSmartMatch === 'function') {
        var result = callServerSmartMatch(productName);
        if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
          Lib.logWithEmoji("Smart Match завершен", "INFO", "", "smartMatch", "Поиск окончен", "Calculations", "SUCCESS", {}, { result: result });
        }
        return result;
      }
      throw new Error('callServerSmartMatch не определена');
    } catch (e) {
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Ошибка Smart Match", "ERROR", "", "smartMatch", e.message, "Calculations", "ERROR", {}, { error: e.toString() });
      }
      throw e;
    }
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

    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Запуск полного пересчета каскадов", "INFO", "", "recalculateCascades", "Вызов серверной обработки сертификации", "Certification", "START", {}, null);
    }
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
        if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
          Lib.logWithEmoji("Пересчет каскадов завершен", "INFO", "", "recalculateCascades", "Сертифкация успешно обновлена", "Certification", "SUCCESS", {}, { processed: result.processed, changed: result.changed });
        }
      } else {
        ui.alert('Ошибка: ' + (result.message || 'Неизвестная ошибка'));
        if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
          Lib.logWithEmoji("Ошибка при пересчете каскадов", "WARNING", "", "recalculateCascades", result.message, "Certification", "WARNING", {}, { result: result });
        }
      }

      return result;
    } catch (err) {
      ss.toast('Ошибка пересчёта', 'Ошибка', 3);
      ui.alert('Ошибка: ' + err.message);
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Исключение при пересчете каскадов", "ERROR", "", "recalculateCascades", err.message, "Certification", "ERROR", {}, { error: err.toString() });
      }
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
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Запуск каскада для строки", "INFO", "", "processCascadeRow", "Обработка строки " + row, "Certification", "START", { row: row, column: column }, null);
    }
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    try {
      const result = callServerCascade({
        spreadsheet_id: ss.getId(),
        sheet_name: 'Сертификация',
        row: row,
        changed_column: column,
        new_value: value,
        dry_run: false
      });
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Каскад строки завершен", "INFO", "", "processCascadeRow", "Строка " + row + " обработана", "Certification", "SUCCESS", { row: row }, { result: result });
      }
      return result;
    } catch (e) {
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Ошибка каскада строки", "ERROR", "", "processCascadeRow", e.message, "Certification", "ERROR", { row: row }, { error: e.toString() });
      }
      throw e;
    }
  });
}

/**
 * Вызов сервера для обработки каскадов
 * @param {Object} params - Параметры запроса
 */
function callServerCascade(params) {
  if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
    Lib.logWithEmoji("Вызов API каскадов", "INFO", "", "callServerCascade", "HTTP запрос к серверу каскадов", "Internal", "START", { endpoint: params.row ? '/api/v1/cascade/process' : '/api/v1/cascade/recalculate-all' }, null);
  }
  const BASE_URL = PropertiesService.getScriptProperties().getProperty('SERVER_URL') || 'http://46.226.167.153:8000';
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
      var result = JSON.parse(text);
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("API каскадов: успех", "INFO", "", "callServerCascade", "Сервер вернул " + status, "Internal", "SUCCESS", {}, { status_code: status });
      }
      return result;
    } else {
      console.error('Cascade server error: ' + status + ' - ' + text);
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("API каскадов: ошибка сервера", "ERROR", "", "callServerCascade", text, "Internal", "ERROR", {}, { status_code: status });
      }
      return { status: 'error', message: 'Server returned ' + status };
    }
  } catch (err) {
    console.error('Cascade request failed: ' + err.message);
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("API каскадов: исключение", "ERROR", "", "callServerCascade", err.message, "Internal", "ERROR", {}, { error: err.toString() });
    }
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
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Запуск серверного отображения всех данных", "INFO", "", "serverShowAllOrderData", "Удаление фильтров на сервере", "Dashboard", "START", {}, null);
    }
    var result = _callServerOrderFilter('all');
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Все данные (сервер) отображены", "INFO", "", "serverShowAllOrderData", "Фильтры на сервере сброшены", "Dashboard", "SUCCESS", {}, { result: result });
    }
    return result;
  });
}

/**
 * Показать стадию "Заказ"
 * Эквивалент GAS showOrderStage()
 */
function serverShowOrderStage() {
  return _loggedCall_("serverShowOrderStage", function() {
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Запуск серверной стадии 'Заказ'", "INFO", "", "serverShowOrderStage", "Фильтрация по стадии Заказ на сервере", "Dashboard", "START", { stage: "order" }, null);
    }
    var result = _callServerOrderFilter('order');
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Стадия 'Заказ' (сервер) отображена", "INFO", "", "serverShowOrderStage", "Фильтр Заказ успешно применен", "Dashboard", "SUCCESS", { stage: "order" }, { result: result });
    }
    return result;
  });
}

/**
 * Показать стадию "Акции"
 * Эквивалент GAS showPromotionsStage()
 */
function serverShowPromotionsStage() {
  return _loggedCall_("serverShowPromotionsStage", function() {
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Запуск серверной стадии 'Акции'", "INFO", "", "serverShowPromotionsStage", "Фильтрация по стадии Акции на сервере", "Dashboard", "START", { stage: "promotions" }, null);
    }
    var result = _callServerOrderFilter('promotions');
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Стадия 'Акции' (сервер) отображена", "INFO", "", "serverShowPromotionsStage", "Фильтр Акции успешно применен", "Dashboard", "SUCCESS", { stage: "promotions" }, { result: result });
    }
    return result;
  });
}

/**
 * Показать стадию "Набор"
 * Эквивалент GAS showSetStage()
 */
function serverShowSetStage() {
  return _loggedCall_("serverShowSetStage", function() {
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Запуск серверной стадии 'Набор'", "INFO", "", "serverShowSetStage", "Фильтрация по стадии Набор на сервере", "Dashboard", "START", { stage: "set" }, null);
    }
    var result = _callServerOrderFilter('set');
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Стадия 'Набор' (сервер) отображена", "INFO", "", "serverShowSetStage", "Фильтр Набор успешно применен", "Dashboard", "SUCCESS", { stage: "set" }, { result: result });
    }
    return result;
  });
}

/**
 * Показать стадию "Прайс"
 * Эквивалент GAS showPriceStage()
 */
function serverShowPriceStage() {
  return _loggedCall_("serverShowPriceStage", function() {
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Запуск серверной стадии 'Прайс'", "INFO", "", "serverShowPriceStage", "Фильтрация по стадии Прайс на сервере", "Dashboard", "START", { stage: "price" }, null);
    }
    var result = _callServerOrderFilter('price');
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Стадия 'Прайс' (сервер) отображена", "INFO", "", "serverShowPriceStage", "Фильтр Прайс успешно применен", "Dashboard", "SUCCESS", { stage: "price" }, { result: result });
    }
    return result;
  });
}

/**
 * Вызов сервера для фильтрации стадий заказа
 * @param {string} stage - Стадия (all, order, promotions, set, price)
 * @private
 */
function _callServerOrderFilter(stage) {
  if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
    Lib.logWithEmoji("Вызов API фильтрации заказов", "INFO", "", "_callServerOrderFilter", "Отправка запроса на сервер (стадия: " + stage + ")", "Dashboard", "PROGRESS", { stage: stage }, null);
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const menuTitle = "Сортировка";

  ss.toast("Применяю фильтр '" + stage + "'...", menuTitle, 30);

  const BASE_URL = PropertiesService.getScriptProperties().getProperty('SERVER_URL') || 'http://46.226.167.153:8000';

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
        if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
          Lib.logWithEmoji("Ответ API фильтрации: успех", "INFO", "", "_callServerOrderFilter", "Фильтр применен на сервере", "Dashboard", "SUCCESS", { stage: stage }, { hidden_rows: result.hidden_rows });
        }
      } else {
        if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
          Lib.logWithEmoji("Ошибка API фильтрации: ошибка в ответе", "WARNING", "", "_callServerOrderFilter", result.message || "Неизвестная ошибка", "Dashboard", "WARNING", { stage: stage }, { result: result });
        }
        ui.alert('Ошибка', result.message || 'Неизвестная ошибка', ui.ButtonSet.OK);
      }

      return result;
    } else {
      console.error('Order filter server error: ' + status + ' - ' + text);
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Ошибка API фильтрации: сервер вернул " + status, "ERROR", "", "_callServerOrderFilter", text, "Dashboard", "ERROR", { stage: stage }, { status_code: status });
      }
      ss.toast('Ошибка сервера: ' + status, 'Ошибка', 3);
      return { status: 'error', message: 'Server returned ' + status };
    }
  } catch (err) {
    console.error('Order filter request failed: ' + err.message);
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Ошибка API фильтрации: исключение", "ERROR", "", "_callServerOrderFilter", err.message, "Dashboard", "ERROR", { stage: stage }, { error: err.toString() });
    }
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
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Запуск серверной выгрузки акций", "INFO", "", "serverExportPromotions", "Перенаправление на _callServerExport", "Dashboard", "START", { exportType: "promotions" }, null);
    }
    return _callServerExport('promotions');
  });
}

/**
 * Выгрузка наборов через сервер
 * Эквивалент GAS exportSets()
 */
function serverExportSets() {
  return _loggedCall_("serverExportSets", function() {
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Запуск серверной выгрузки наборов", "INFO", "", "serverExportSets", "Перенаправление на _callServerExport", "Dashboard", "START", { exportType: "sets" }, null);
    }
    return _callServerExport('sets');
  });
}

/**
 * Вызов сервера для выгрузки данных
 * @param {string} exportType - Тип выгрузки (promotions, sets)
 * @private
 */
function _callServerExport(exportType) {
  if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
    Lib.logWithEmoji("Вызов API выгрузки", "INFO", "", "_callServerExport", "Подготовка запроса (тип: " + exportType + ")", "Dashboard", "PROGRESS", { exportType: exportType }, null);
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const menuTitle = "Выгрузка";

  // Определяем проект
  const projectKey = (typeof CONFIG !== 'undefined' && CONFIG.ACTIVE_PROJECT_KEY)
    ? CONFIG.ACTIVE_PROJECT_KEY.toLowerCase()
    : null;

  if (!projectKey) {
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Ошибка выгрузки: проект не определен", "ERROR", "", "_callServerExport", "CONFIG.ACTIVE_PROJECT_KEY отсутствует", "Dashboard", "ERROR", {}, null);
    }
    ui.alert('Ошибка', 'Проект не определен.', ui.ButtonSet.OK);
    return { status: 'error', message: 'Project not defined' };
  }

  const typeName = exportType === 'promotions' ? 'акций' : 'наборов';
  ss.toast('Выгрузка ' + typeName + '...', menuTitle, 30);

  const BASE_URL = PropertiesService.getScriptProperties().getProperty('SERVER_URL') || 'http://46.226.167.153:8000';
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
        if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
          Lib.logWithEmoji("Ответ API выгрузки: успех", "INFO", "", "_callServerExport", "Данные выгружены", "Dashboard", "SUCCESS", { exportType: exportType }, { exported_rows: result.exported_rows, url: result.target_url });
        }
      } else {
        if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
          Lib.logWithEmoji("Ошибка API выгрузки: ошибка в ответе", "WARNING", "", "_callServerExport", result.message || "Неизвестная ошибка", "Dashboard", "WARNING", { exportType: exportType }, { result: result });
        }
        ui.alert('Ошибка', result.message || 'Неизвестная ошибка', ui.ButtonSet.OK);
      }

      return result;
    } else {
      console.error('Export server error: ' + status + ' - ' + text);
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Ошибка API выгрузки: сервер вернул " + status, "ERROR", "", "_callServerExport", text, "Dashboard", "ERROR", { exportType: exportType }, { status_code: status });
      }
      ss.toast('Ошибка сервера: ' + status, 'Ошибка', 3);
      return { status: 'error', message: 'Server returned ' + status };
    }
  } catch (err) {
    console.error('Export request failed: ' + err.message);
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Ошибка API выгрузки: исключение", "ERROR", "", "_callServerExport", err.message, "Dashboard", "ERROR", { exportType: exportType }, { error: err.toString() });
    }
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
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Запуск серверного форматирования заказа", "INFO", "", "serverFormatOrder", "Вызов API для нормализации колонок", "Dashboard", "START", {}, null);
    }
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();

    ss.toast('Форматирование листа "Ордер"...', 'Поставка', 30);

    const BASE_URL = PropertiesService.getScriptProperties().getProperty('SERVER_URL') || 'http://46.226.167.153:8000';

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
          if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
            Lib.logWithEmoji("Серверное форматирование завершено", "INFO", "", "serverFormatOrder", "Лист нормализован", "Dashboard", "SUCCESS", {}, { rows: result.rows_processed });
          }
        } else {
          ui.alert('Ошибка', result.message || 'Неизвестная ошибка', ui.ButtonSet.OK);
          if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
            Lib.logWithEmoji("Ошибка серверного форматирования", "WARNING", "", "serverFormatOrder", result.message, "Dashboard", "WARNING", {}, { result: result });
          }
        }

        return result;
      } else {
        console.error('Invoice format server error: ' + status + ' - ' + text);
        if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
          Lib.logWithEmoji("Ошибка API форматирования: сервер вернул " + status, "ERROR", "", "serverFormatOrder", text, "Dashboard", "ERROR", {}, { status_code: status });
        }
        ss.toast('Ошибка сервера: ' + status, 'Ошибка', 3);
        return { status: 'error', message: 'Server returned ' + status };
      }
    } catch (err) {
      console.error('Invoice format request failed: ' + err.message);
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Ошибка API форматирования: исключение", "ERROR", "", "serverFormatOrder", err.message, "Dashboard", "ERROR", {}, { error: err.toString() });
      }
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
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Запуск серверной генерации инвойса", "INFO", "", "serverCreateInvoice", "Сбор данных для 'Для инвойса'", "Dashboard", "START", {}, null);
    }
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();

    ss.toast('Создание листа "Для инвойса"...', 'Поставка', 60);

    const BASE_URL = PropertiesService.getScriptProperties().getProperty('SERVER_URL') || 'http://46.226.167.153:8000';

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
          if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
            Lib.logWithEmoji("Серверная генерация инвойса завершена", "INFO", "", "serverCreateInvoice", "Лист успешно создан", "Dashboard", "SUCCESS", {}, { rows: result.rows_processed });
          }
        } else {
          ui.alert('Ошибка', result.message || 'Неизвестная ошибка', ui.ButtonSet.OK);
          if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
            Lib.logWithEmoji("Ошибка генерации инвойса", "WARNING", "", "serverCreateInvoice", result.message, "Dashboard", "WARNING", {}, { result: result });
          }
        }

        return result;
      } else {
        console.error('Invoice create server error: ' + status + ' - ' + text);
        if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
          Lib.logWithEmoji("Ошибка API инвойса: сервер вернул " + status, "ERROR", "", "serverCreateInvoice", text, "Dashboard", "ERROR", {}, { status_code: status });
        }
        ss.toast('Ошибка сервера: ' + status, 'Ошибка', 3);
        return { status: 'error', message: 'Server returned ' + status };
      }
    } catch (err) {
      console.error('Invoice create request failed: ' + err.message);
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Ошибка API инвойса: исключение", "ERROR", "", "serverCreateInvoice", err.message, "Dashboard", "ERROR", {}, { error: err.toString() });
      }
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
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Запуск серверного создания листа новинок", "INFO", "", "serverCreateNews", "Перенос новинок из сертификации", "Certification", "START", {}, null);
    }
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();

    ss.toast('Создание листа "New sert"...', 'Сертификация', 30);

    const BASE_URL = PropertiesService.getScriptProperties().getProperty('SERVER_URL') || 'http://46.226.167.153:8000';

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
           if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
            Lib.logWithEmoji("Серверный лист новинок создан", "INFO", "", "serverCreateNews", "Лист " + result.sheet_name + " готов", "Certification", "SUCCESS", {}, { rows: result.rows_affected });
          }
        } else {
          ui.alert('Ошибка', result.message || 'Неизвестная ошибка', ui.ButtonSet.OK);
          if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
            Lib.logWithEmoji("Ошибка создания листа новинок", "WARNING", "", "serverCreateNews", result.message, "Certification", "WARNING", {}, { result: result });
          }
        }

        return result;
      } else {
        console.error('News sheet server error: ' + status + ' - ' + text);
        if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
          Lib.logWithEmoji("Ошибка API новинок: сервер вернул " + status, "ERROR", "", "serverCreateNews", text, "Certification", "ERROR", {}, { status_code: status });
        }
        ss.toast('Ошибка сервера: ' + status, 'Ошибка', 3);
        return { status: 'error', message: 'Server returned ' + status };
      }
    } catch (err) {
      console.error('News sheet request failed: ' + err.message);
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Ошибка API новинок: исключение", "ERROR", "", "serverCreateNews", err.message, "Certification", "ERROR", {}, { error: err.toString() });
      }
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
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Запуск серверного расчета спирта", "INFO", "", "serverCalcSpirit", "Вызов API для расчета номеров", "Certification", "START", {}, null);
    }
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();

    ss.toast('Расчёт номеров спиртов...', 'Сертификация', 30);

    const BASE_URL = PropertiesService.getScriptProperties().getProperty('SERVER_URL') || 'http://46.226.167.153:8000';

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
          if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
            Lib.logWithEmoji("Серверный расчет спирта завершен", "INFO", "", "serverCalcSpirit", result.message, "Certification", "SUCCESS", {}, { rows: result.rows_affected });
          }
        } else if (result && result.status === 'not_implemented') {
          ui.alert('В разработке', result.message, ui.ButtonSet.OK);
          if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
            Lib.logWithEmoji("API расчета спирта: в разработке", "INFO", "", "serverCalcSpirit", result.message, "Certification", "SUCCESS", {}, { result: result });
          }
        } else {
          ui.alert('Ошибка', result.message || 'Неизвестная ошибка', ui.ButtonSet.OK);
          if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
            Lib.logWithEmoji("Ошибка API расчета спирта", "WARNING", "", "serverCalcSpirit", result.message, "Certification", "WARNING", {}, { result: result });
          }
        }

        return result;
      } else {
        console.error('Spirit calc server error: ' + status + ' - ' + text);
        if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
          Lib.logWithEmoji("Ошибка API расчета спирта: сервер вернул " + status, "ERROR", "", "serverCalcSpirit", text, "Certification", "ERROR", {}, { status_code: status });
        }
        ss.toast('Ошибка сервера: ' + status, 'Ошибка', 3);
        return { status: 'error', message: 'Server returned ' + status };
      }
    } catch (err) {
      console.error('Spirit calc request failed: ' + err.message);
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Ошибка API расчета спирта: исключение", "ERROR", "", "serverCalcSpirit", err.message, "Certification", "ERROR", {}, { error: err.toString() });
      }
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
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Запуск серверной генерации протоколов 353пп", "INFO", "", "serverGenProtocols", "Вызов API для создания документов", "Certification", "START", {}, null);
    }
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();

    ss.toast('Генерация протоколов 353пп...', 'Сертификация', 60);

    const BASE_URL = PropertiesService.getScriptProperties().getProperty('SERVER_URL') || 'http://46.226.167.153:8000';

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
          if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
            Lib.logWithEmoji("API протоколов 353пп: в разработке", "INFO", "", "serverGenProtocols", result.message, "Certification", "SUCCESS", {}, { result: result });
          }
        } else if (result && result.status === 'success') {
          ss.toast('Протоколы сгенерированы', '✅ Готово', 5);
          ui.alert('Успех', result.message, ui.ButtonSet.OK);
          if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
            Lib.logWithEmoji("Серверная генерация протоколов завершена", "INFO", "", "serverGenProtocols", result.message, "Certification", "SUCCESS", {}, { result: result });
          }
        } else {
          ui.alert('Ошибка', result.message || 'Неизвестная ошибка', ui.ButtonSet.OK);
          if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
            Lib.logWithEmoji("Ошибка API протоколов 353пп", "WARNING", "", "serverGenProtocols", result.message, "Certification", "WARNING", {}, { result: result });
          }
        }

        return result;
      } else {
        console.error('Protocols server error: ' + status + ' - ' + text);
        if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
          Lib.logWithEmoji("Ошибка API протоколов 353пп: сервер вернул " + status, "ERROR", "", "serverGenProtocols", text, "Certification", "ERROR", {}, { status_code: status });
        }
        ss.toast('Ошибка сервера: ' + status, 'Ошибка', 3);
        return { status: 'error', message: 'Server returned ' + status };
      }
    } catch (err) {
      console.error('Protocols request failed: ' + err.message);
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Ошибка API протоколов 353пп: исключение", "ERROR", "", "serverGenProtocols", err.message, "Certification", "ERROR", {}, { error: err.toString() });
      }
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
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Запуск серверной генерации макетов ДС", "INFO", "", "serverGenDsLayouts", "Вызов API для создания макетов", "Certification", "START", {}, null);
    }
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();

    ss.toast('Генерация макетов ДС...', 'Сертификация', 60);

    const BASE_URL = PropertiesService.getScriptProperties().getProperty('SERVER_URL') || 'http://46.226.167.153:8000';

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
          if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
            Lib.logWithEmoji("API макетов ДС: в разработке", "INFO", "", "serverGenDsLayouts", result.message, "Certification", "SUCCESS", {}, { result: result });
          }
        } else if (result && result.status === 'success') {
          ss.toast('Макеты ДС сгенерированы', '✅ Готово', 5);
          ui.alert('Успех', result.message, ui.ButtonSet.OK);
          if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
            Lib.logWithEmoji("Серверная генерация макетов ДС завершена", "INFO", "", "serverGenDsLayouts", result.message, "Certification", "SUCCESS", {}, { result: result });
          }
        } else {
          ui.alert('Ошибка', result.message || 'Неизвестная ошибка', ui.ButtonSet.OK);
          if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
            Lib.logWithEmoji("Ошибка API макетов ДС", "WARNING", "", "serverGenDsLayouts", result.message, "Certification", "WARNING", {}, { result: result });
          }
        }

        return result;
      } else {
        console.error('DS layouts server error: ' + status + ' - ' + text);
        if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
          Lib.logWithEmoji("Ошибка API макетов ДС: сервер вернул " + status, "ERROR", "", "serverGenDsLayouts", text, "Certification", "ERROR", {}, { status_code: status });
        }
        ss.toast('Ошибка сервера: ' + status, 'Ошибка', 3);
        return { status: 'error', message: 'Server returned ' + status };
      }
    } catch (err) {
      console.error('DS layouts request failed: ' + err.message);
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Ошибка API макетов ДС: исключение", "ERROR", "", "serverGenDsLayouts", err.message, "Certification", "ERROR", {}, { error: err.toString() });
      }
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
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Запуск пересчета динамики цен", "INFO", "", "serverRecalculatePriceDynamics", "Вызов серверного API для обновления формул", "Calculations", "START", { sheet: sheetName || "Динамика цены", dryRun: dryRun }, null);
    }
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
          if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
            Lib.logWithEmoji("Пересчет динамики завершен", "INFO", "", "serverRecalculatePriceDynamics", "Формулы успешно обновлены", "Calculations", "SUCCESS", { sheet: sheetName }, { blocks: result.blocks_processed, rows: result.rows_updated });
          }
        } else {
          ss.toast(result.message || 'Ошибка', 'Ошибка', 3);
          if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
            Lib.logWithEmoji("Ошибка при пересчете динамики", "WARNING", "", "serverRecalculatePriceDynamics", result.message, "Calculations", "WARNING", {}, { result: result });
          }
        }

        return result;
      } else {
        console.error('Price dynamics formulas server error: ' + status + ' - ' + text);
        ss.toast('Ошибка сервера: ' + status, 'Ошибка', 3);
        if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
          Lib.logWithEmoji("Ошибка API динамики: сервер вернул " + status, "ERROR", "", "serverRecalculatePriceDynamics", text, "Calculations", "ERROR", {}, { status_code: status });
        }
        return { status: 'error', message: 'Server returned ' + status };
      }
    } catch (err) {
      console.error('Price dynamics formulas request failed: ' + err.message);
      ss.toast('Ошибка: ' + err.message, 'Ошибка', 3);
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Ошибка API динамики: исключение", "ERROR", "", "serverRecalculatePriceDynamics", err.message, "Calculations", "ERROR", {}, { error: err.toString() });
      }
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
    if (!silent && typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Запуск обновления расчета цены", "INFO", "", "serverUpdatePriceCalc", "Вызов серверного API для обновления INDEX/MATCH формул", "Calculations", "START", { silent: silent, dryRun: dryRun }, null);
    }
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
            if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
              Lib.logWithEmoji("Обновление расчета завершено", "INFO", "", "serverUpdatePriceCalc", "Формулы успешно притянуты", "Calculations", "SUCCESS", {}, { rows: result.rows_updated });
            }
          }
        } else if (!silent) {
          ss.toast(result.message || 'Ошибка', 'Ошибка', 3);
          if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
            Lib.logWithEmoji("Ошибка при обновлении расчета", "WARNING", "", "serverUpdatePriceCalc", result.message, "Calculations", "WARNING", {}, { result: result });
          }
        }

        return result;
      } else {
        console.error('Price calc formulas server error: ' + status + ' - ' + text);
        if (!silent) {
          ss.toast('Ошибка сервера: ' + status, 'Ошибка', 3);
          if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
            Lib.logWithEmoji("Ошибка API расчета: сервер вернул " + status, "ERROR", "", "serverUpdatePriceCalc", text, "Calculations", "ERROR", {}, { status_code: status });
          }
        }
        return { status: 'error', message: 'Server returned ' + status };
      }
    } catch (err) {
      console.error('Price calc formulas request failed: ' + err.message);
      if (!silent) {
        ss.toast('Ошибка: ' + err.message, 'Ошибка', 3);
        if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
          Lib.logWithEmoji("Ошибка API расчета: исключение", "ERROR", "", "serverUpdatePriceCalc", err.message, "Calculations", "ERROR", {}, { error: err.toString() });
        }
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
    var targetYear = year || new Date().getFullYear();
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Запуск добавления года в Динамику", "INFO", "", "serverAddNewYearColumns", "Добавление колонок за " + targetYear + " год", "Calculations", "START", { year: targetYear, dryRun: dryRun }, null);
    }
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ui = SpreadsheetApp.getUi();
    var spreadsheetId = ss.getId();

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
          if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
            Lib.logWithEmoji("Год успешно добавлен", "INFO", "", "serverAddNewYearColumns", "Столбцы добавлены", "Calculations", "SUCCESS", { year: targetYear }, { columns: result.columns_added });
          }
        } else if (result && result.status === 'exists') {
          ss.toast('Блок за ' + targetYear + ' год уже существует', 'Инфо', 3);
          ui.alert('Информация', result.message, ui.ButtonSet.OK);
          if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
            Lib.logWithEmoji("Добавление года пропущено", "INFO", "", "serverAddNewYearColumns", "Год уже существует", "Calculations", "SKIP", { year: targetYear }, { message: result.message });
          }
        } else {
          ss.toast(result.message || 'Ошибка', 'Ошибка', 3);
          ui.alert('Ошибка', result.message || 'Неизвестная ошибка', ui.ButtonSet.OK);
          if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
            Lib.logWithEmoji("Ошибка при добавлении года", "WARNING", "", "serverAddNewYearColumns", result.message, "Calculations", "WARNING", { year: targetYear }, { result: result });
          }
        }

        return result;
      } else {
        console.error('Add year columns server error: ' + status + ' - ' + text);
        ss.toast('Ошибка сервера: ' + status, 'Ошибка', 3);
        if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
          Lib.logWithEmoji("Ошибка API года: сервер вернул " + status, "ERROR", "", "serverAddNewYearColumns", text, "Calculations", "ERROR", { year: targetYear }, { status_code: status });
        }
        return { status: 'error', message: 'Server returned ' + status };
      }
    } catch (err) {
      console.error('Add year columns request failed: ' + err.message);
      ss.toast('Ошибка: ' + err.message, 'Ошибка', 3);
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Ошибка API года: исключение", "ERROR", "", "serverAddNewYearColumns", err.message, "Calculations", "ERROR", { year: targetYear }, { error: err.toString() });
      }
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
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Запуск серверной архивации логов", "INFO", "", "serverArchiveLogs", "Вынос логов в архивную таблицу", "Admin", "START", { project: projectCode, dryRun: dryRun }, null);
    }
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
          if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
            Lib.logWithEmoji("Архивация завершена", "INFO", "", "serverArchiveLogs", "Логи успешно перемещены", "Admin", "SUCCESS", { project: projectCode }, { rows: result.total_rows, archive: result.archive_name });
          }
        } else {
          ss.toast(result.message || 'Ошибка', 'Ошибка', 3);
          if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
            Lib.logWithEmoji("Ошибка при архивации логов", "WARNING", "", "serverArchiveLogs", result.message, "Admin", "WARNING", { project: projectCode }, { result: result });
          }
        }

        return result;
      } else {
        console.error('Log archive server error: ' + status + ' - ' + text);
        ss.toast('Ошибка сервера: ' + status, 'Ошибка', 3);
        if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
          Lib.logWithEmoji("Ошибка API архивации: сервер вернул " + status, "ERROR", "", "serverArchiveLogs", text, "Admin", "ERROR", { project: projectCode }, { status_code: status });
        }
        return { status: 'error', message: 'Server returned ' + status };
      }
    } catch (err) {
      console.error('Log archive request failed: ' + err.message);
      ss.toast('Ошибка: ' + err.message, 'Ошибка', 3);
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Ошибка API архивации: исключение", "ERROR", "", "serverArchiveLogs", err.message, "Admin", "ERROR", { project: projectCode }, { error: err.toString() });
      }
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
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Запуск серверной очистки логов", "INFO", "", "serverResetLogSheet", "Очистка листа " + (sheetName || "Логи"), "Admin", "START", { sheet: sheetName, dryRun: dryRun }, null);
    }
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
          if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
            Lib.logWithEmoji("Лист логов очищен", "INFO", "", "serverResetLogSheet", "Данные успешно удалены", "Admin", "SUCCESS", { sheet: sheetName }, { rows: result.total_rows });
          }
        } else {
          ss.toast(result.message || 'Ошибка', 'Ошибка', 3);
          if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
            Lib.logWithEmoji("Ошибка при очистке логов", "WARNING", "", "serverResetLogSheet", result.message, "Admin", "WARNING", { sheet: sheetName }, { result: result });
          }
        }

        return result;
      } else {
        console.error('Log reset server error: ' + status + ' - ' + text);
        ss.toast('Ошибка сервера: ' + status, 'Ошибка', 3);
        if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
          Lib.logWithEmoji("Ошибка API очистки: сервер вернул " + status, "ERROR", "", "serverResetLogSheet", text, "Admin", "ERROR", { sheet: sheetName }, { status_code: status });
        }
        return { status: 'error', message: 'Server returned ' + status };
      }
    } catch (err) {
      console.error('Log reset request failed: ' + err.message);
      ss.toast('Ошибка: ' + err.message, 'Ошибка', 3);
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Ошибка API очистки: исключение", "ERROR", "", "serverResetLogSheet", err.message, "Admin", "ERROR", { sheet: sheetName }, { error: err.toString() });
      }
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
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Запуск ночной ротации логов", "INFO", "", "serverMidnightRotation", "Автоматическая архивация и очистка", "Admin", "START", { project: projectCode, force: force, dryRun: dryRun }, null);
    }
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
          if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
            Lib.logWithEmoji("Ротация завершена успешно", "INFO", "", "serverMidnightRotation", "Логи обработаны", "Admin", "SUCCESS", { project: projectCode }, { rows: result.total_rows, message: result.message });
          }
        } else if (result && result.status === 'skipped') {
          ss.toast('Ротация пропущена: уже выполнена сегодня', 'Инфо', 3);
          if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
            Lib.logWithEmoji("Ротация пропущена", "INFO", "", "serverMidnightRotation", "Уже выполнено сегодня", "Admin", "SKIP", { project: projectCode }, { message: result.message });
          }
        } else {
          ss.toast(result.message || 'Ошибка', 'Ошибка', 3);
          if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
            Lib.logWithEmoji("Ошибка при ротации логов", "WARNING", "", "serverMidnightRotation", result.message, "Admin", "WARNING", { project: projectCode }, { result: result });
          }
        }

        return result;
      } else {
        console.error('Log rotation server error: ' + status + ' - ' + text);
        ss.toast('Ошибка сервера: ' + status, 'Ошибка', 3);
        if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
          Lib.logWithEmoji("Ошибка API ротации: сервер вернул " + status, "ERROR", "", "serverMidnightRotation", text, "Admin", "ERROR", { project: projectCode }, { status_code: status });
        }
        return { status: 'error', message: 'Server returned ' + status };
      }
    } catch (err) {
      console.error('Log rotation request failed: ' + err.message);
      ss.toast('Ошибка: ' + err.message, 'Ошибка', 3);
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Ошибка API ротации: исключение", "ERROR", "", "serverMidnightRotation", err.message, "Admin", "ERROR", { project: projectCode }, { error: err.toString() });
      }
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
  if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
    Lib.logWithEmoji("Запрос статуса логов", "INFO", "", "serverGetLogStatus", "Обращение к серверу за метаданными архива", "Calculations", "START", { project: projectCode }, null);
  }
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
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Статус логов получен", "INFO", "", "serverGetLogStatus", "Данные успешно отображены", "Calculations", "SUCCESS", { project: projectCode }, { result: result });
      }
      return result;
    } else {
      console.error('Log status server error: ' + status + ' - ' + text);
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji("Ошибка статуса логов", "ERROR", "", "serverGetLogStatus", text, "Calculations", "ERROR", { project: projectCode }, { status_code: status });
      }
      ui.alert('Ошибка', 'Ошибка сервера: ' + status, ui.ButtonSet.OK);
      return { status: 'error', message: 'Server returned ' + status };
    }
  } catch (err) {
    console.error('Log status request failed: ' + err.message);
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji("Ошибка статуса логов: исключение", "ERROR", "", "serverGetLogStatus", err.message, "Calculations", "ERROR", { project: projectCode }, { error: err.toString() });
    }
    ui.alert('Ошибка', err.message, ui.ButtonSet.OK);
    return { status: 'error', message: err.message };
  }
}

// =======================================================================================
// УПРАВЛЕНИЕ НАСТРОЙКАМИ СЕРВЕРА
// =======================================================================================

/**
 * Инициализирует URL сервера в свойствах скрипта.
 * Можно вызвать вручную из редактора или через меню, если сервер не отвечает.
 */
function initProductionServerUrl() {
  const prodUrl = 'http://46.226.167.153:8000';
  PropertiesService.getScriptProperties().setProperty('SERVER_URL', prodUrl);
  console.log('SERVER_URL set to: ' + prodUrl);
  if (typeof SpreadsheetApp !== 'undefined') {
    SpreadsheetApp.getActiveSpreadsheet().toast('URL сервера установлен: ' + prodUrl, '⚙️ Настройки');
  }
}

// Вызываем при установке или обновлении, если нужно гарантировать наличие URL
function forceInitServerUrl() {
  initProductionServerUrl();
}
