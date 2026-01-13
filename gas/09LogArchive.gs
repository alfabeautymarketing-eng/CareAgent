// =======================================================================================
// 09LogArchive.js — АРХИВИРОВАНИЕ И РОТАЦИЯ ЛОГОВ
// =======================================================================================
// DESCRIPTION:
//   Модуль для ежедневного архивирования логов в полночь.
//   Создаёт месячные архивные таблицы в указанной папке Google Drive.
//   Обеспечивает полное логирование всех функций проекта.
// =======================================================================================

(function (global) {
  "use strict";

  global.Lib = global.Lib || {};

  // =======================================================================================
  // КОНФИГУРАЦИЯ АРХИВИРОВАНИЯ
  // =======================================================================================

  const LOG_ARCHIVE_CONFIG = {
    // ID папки для архивов логов
    ARCHIVE_FOLDER_ID: "1oWoxwuOMlriwgHuua6MpLheROEIT4SIP",

    // Листы которые нужно архивировать
    SHEETS_TO_ARCHIVE: ["Логи", "Журнал синхро", "Журнал логов"],

    // Названия месяцев на русском
    MONTHS_RU: [
      "январь", "февраль", "март", "апрель", "май", "июнь",
      "июль", "август", "сентябрь", "октябрь", "ноябрь", "декабрь"
    ],

    // Ключ для хранения даты последнего архивирования
    LAST_ARCHIVE_KEY: "ECOSYS_LAST_LOG_ARCHIVE_DATE",

    // Ключ для хранения ID текущей месячной таблицы
    MONTHLY_SHEET_ID_KEY: "ECOSYS_MONTHLY_ARCHIVE_SHEET_ID"
  };

  // Экспортируем конфигурацию
  Lib.LOG_ARCHIVE_CONFIG = LOG_ARCHIVE_CONFIG;

  // =======================================================================================
  // ОСНОВНЫЕ ФУНКЦИИ АРХИВИРОВАНИЯ
  // =======================================================================================

  /**
   * Получает код проекта из конфигурации (mt, sk, ss)
   * @returns {string} Код проекта в нижнем регистре
   */
  function _getProjectCode_() {
    try {
      // Пытаемся получить из меню конфига
      if (typeof getMenuConfig === "function") {
        const config = getMenuConfig({ skipFetch: true });
        if (config && config.project) {
          return config.project.toLowerCase();
        }
      }

      // Fallback: пытаемся определить по названию таблицы
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const name = ss.getName().toLowerCase();
      if (name.includes("mt") || name.includes("mettler")) return "mt";
      if (name.includes("sk") || name.includes("skinakas")) return "sk";
      if (name.includes("ss") || name.includes("skinscience")) return "ss";

      return "project"; // Fallback
    } catch (e) {
      return "project";
    }
  }

  /**
   * Получает или создаёт месячную архивную таблицу.
   * Название формата: "mt-январь" / "sk-февраль" и т.д.
   * @param {Date} date - Дата для определения месяца
   * @returns {Spreadsheet} Архивная таблица
   */
  function _getOrCreateMonthlyArchiveSheet_(date) {
    const projectCode = _getProjectCode_();
    const monthIndex = date.getMonth();
    const year = date.getFullYear();
    const monthName = LOG_ARCHIVE_CONFIG.MONTHS_RU[monthIndex];
    const archiveName = `${projectCode}-${monthName}-${year}`;

    const props = PropertiesService.getScriptProperties();
    const monthKey = `${LOG_ARCHIVE_CONFIG.MONTHLY_SHEET_ID_KEY}_${year}_${monthIndex}`;

    // Проверяем есть ли уже созданная таблица для этого месяца
    const existingId = props.getProperty(monthKey);
    if (existingId) {
      try {
        const existingSheet = SpreadsheetApp.openById(existingId);
        if (existingSheet) {
          return existingSheet;
        }
      } catch (e) {
        // Таблица удалена или недоступна, создадим новую
        props.deleteProperty(monthKey);
      }
    }

    // Ищем таблицу в папке архивов
    try {
      const folder = DriveApp.getFolderById(LOG_ARCHIVE_CONFIG.ARCHIVE_FOLDER_ID);
      const files = folder.getFilesByName(archiveName);

      if (files.hasNext()) {
        const file = files.next();
        if (file.getMimeType() === MimeType.GOOGLE_SHEETS) {
          const sheet = SpreadsheetApp.openById(file.getId());
          props.setProperty(monthKey, file.getId());
          return sheet;
        }
      }

      // Создаём новую архивную таблицу
      const newSheet = SpreadsheetApp.create(archiveName);
      const newFile = DriveApp.getFileById(newSheet.getId());

      // Перемещаем в папку архивов
      newFile.moveTo(folder);

      // Сохраняем ID
      props.setProperty(monthKey, newSheet.getId());

      // Инициализируем структуру архивной таблицы
      _initArchiveSheetStructure_(newSheet);

      _logToSheet_("ARCHIVE", `Создана новая архивная таблица: ${archiveName}`, "INFO");

      return newSheet;
    } catch (e) {
      _logToSheet_("ARCHIVE", `Ошибка создания архивной таблицы: ${e.message}`, "ERROR");
      throw e;
    }
  }

  /**
   * Инициализирует структуру архивной таблицы с листами для каждого типа логов
   * @param {Spreadsheet} archiveSheet - Архивная таблица
   */
  function _initArchiveSheetStructure_(archiveSheet) {
    const sheetsToCreate = LOG_ARCHIVE_CONFIG.SHEETS_TO_ARCHIVE;

    sheetsToCreate.forEach(function (sheetName, index) {
      let sh;
      if (index === 0) {
        // Используем первый лист (Sheet1) и переименовываем
        sh = archiveSheet.getSheets()[0];
        sh.setName(sheetName);
      } else {
        sh = archiveSheet.insertSheet(sheetName);
      }

      // Устанавливаем заголовки в зависимости от типа листа
      let headers;
      if (sheetName === "Логи") {
        headers = ["🕒 Время", "🏷️ Категория", "💬 Действие", "📝 Детали", "🔘 Статус", "📅 Дата архивации"];
      } else if (sheetName === "Журнал синхро") {
        headers = ["Дата/время", "ID", "Источник", "Цель", "Старое значение", "Новое значение", "Категория", "Хэштеги", "Событие", "📅 Дата архивации"];
      } else {
        headers = ["Timestamp", "Level", "Message", "Function", "Details", "📅 Дата архивации"];
      }

      sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
      sh.setFrozenRows(1);

      // Форматирование
      sh.getRange(1, 1, 1, headers.length).setBackground("#e8eaf6");
    });
  }

  /**
   * Проверяет, нужно ли выполнять архивирование (раз в сутки)
   * @returns {boolean} true если нужно архивировать
   */
  function _shouldArchiveToday_() {
    const props = PropertiesService.getScriptProperties();
    const lastArchiveStr = props.getProperty(LOG_ARCHIVE_CONFIG.LAST_ARCHIVE_KEY);

    if (!lastArchiveStr) {
      return true; // Никогда не архивировали
    }

    const today = new Date();
    const todayStr = Utilities.formatDate(today, "Europe/Moscow", "yyyy-MM-dd");

    return lastArchiveStr !== todayStr;
  }

  /**
   * Копирует данные из листа логов в архивную таблицу
   * @param {Sheet} sourceSheet - Исходный лист
   * @param {Spreadsheet} archiveSpreadsheet - Архивная таблица
   * @param {string} sheetName - Имя листа
   */
  function _copyLogsToArchive_(sourceSheet, archiveSpreadsheet, sheetName) {
    if (!sourceSheet) return 0;

    const lastRow = sourceSheet.getLastRow();
    if (lastRow <= 1) return 0; // Только заголовки или пусто

    const lastCol = sourceSheet.getLastColumn();
    if (lastCol <= 0) return 0;

    // Получаем данные (без заголовков)
    const data = sourceSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

    // Добавляем колонку с датой архивации
    const archiveDate = Utilities.formatDate(new Date(), "Europe/Moscow", "dd.MM.yyyy HH:mm:ss");
    const dataWithDate = data.map(function (row) {
      return row.concat([archiveDate]);
    });

    // Получаем или создаём лист в архивной таблице
    let archiveSheet = archiveSpreadsheet.getSheetByName(sheetName);
    if (!archiveSheet) {
      archiveSheet = archiveSpreadsheet.insertSheet(sheetName);
      // Копируем заголовки
      const headers = sourceSheet.getRange(1, 1, 1, lastCol).getValues()[0];
      headers.push("📅 Дата архивации");
      archiveSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
      archiveSheet.setFrozenRows(1);
    }

    // Добавляем данные в архив
    const archiveLastRow = archiveSheet.getLastRow();
    if (dataWithDate.length > 0) {
      archiveSheet.getRange(archiveLastRow + 1, 1, dataWithDate.length, dataWithDate[0].length)
        .setValues(dataWithDate);
    }

    return dataWithDate.length;
  }

  /**
   * ПУБЛИЧНАЯ: Выполняет полное архивирование всех логов.
   * Вызывается ежедневно в полночь по триггеру.
   */
  Lib.archiveLogsDaily = function () {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      console.error("archiveLogsDaily: No active spreadsheet");
      return { success: false, error: "No active spreadsheet" };
    }

    try {
      const now = new Date();
      const archiveSpreadsheet = _getOrCreateMonthlyArchiveSheet_(now);
      let totalRows = 0;
      const results = {};

      LOG_ARCHIVE_CONFIG.SHEETS_TO_ARCHIVE.forEach(function (sheetName) {
        const sourceSheet = ss.getSheetByName(sheetName);
        if (sourceSheet) {
          const rowsCopied = _copyLogsToArchive_(sourceSheet, archiveSpreadsheet, sheetName);
          results[sheetName] = rowsCopied;
          totalRows += rowsCopied;
        } else {
          results[sheetName] = 0;
        }
      });

      // Записываем дату последнего архивирования
      const props = PropertiesService.getScriptProperties();
      const todayStr = Utilities.formatDate(now, "Europe/Moscow", "yyyy-MM-dd");
      props.setProperty(LOG_ARCHIVE_CONFIG.LAST_ARCHIVE_KEY, todayStr);

      _logToSheet_("ARCHIVE", `Архивирование завершено: ${totalRows} записей`, "INFO");

      return {
        success: true,
        totalRows: totalRows,
        details: results,
        archiveSheet: archiveSpreadsheet.getName(),
        archiveUrl: archiveSpreadsheet.getUrl()
      };
    } catch (e) {
      _logToSheet_("ARCHIVE", `Ошибка архивирования: ${e.message}`, "ERROR");
      return { success: false, error: e.message };
    }
  };

  /**
   * ПУБЛИЧНАЯ: Очищает LOG_DEBUG лист после архивирования.
   * ОБНОВЛЕНО: Использует новую систему LOG_DEBUG вместо старого листа "Логи"
   * Вызывается после archiveLogsDaily.
   */
  Lib.resetDailyLogSheet = function () {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss || !Lib.CONFIG || !Lib.CONFIG.SHEETS || !Lib.CONFIG.SHEETS.LOG_DEBUG) {
      return false;
    }

    try {
      const logSheetName = Lib.CONFIG.SHEETS.LOG_DEBUG;
      let sh = ss.getSheetByName(logSheetName);

      if (sh) {
        // Очищаем данные лист, сохраняя заголовки
        const lastRow = sh.getLastRow();
        if (lastRow > 1) {
          sh.deleteRows(2, lastRow - 1);
        }
      }

      _logToSheet_("SYSTEM", "Лист LOG_DEBUG очищен после архивации", "INFO");
      return true;
    } catch (e) {
      console.error("resetDailyLogSheet error:", e);
      return false;
    }
  };

  /**
   * ПУБЛИЧНАЯ: Комбинированная функция для полночного запуска.
   * 1. Архивирует логи в месячную таблицу
   * 2. Очищает лист "Логи"
   */
  Lib.midnightLogRotation = function () {
    // Проверяем нужно ли архивировать
    if (!_shouldArchiveToday_()) {
      console.log("midnightLogRotation: Already archived today, skipping");
      return { skipped: true, reason: "Already archived today" };
    }

    const archiveResult = Lib.archiveLogsDaily();

    if (archiveResult.success) {
      Lib.resetDailyLogSheet();
      return {
        success: true,
        message: "Логи архивированы и лист очищен",
        details: archiveResult
      };
    }

    return archiveResult;
  };

  // =======================================================================================
  // ТРИГГЕР ДЛЯ ПОЛНОЧНОГО ЗАПУСКА
  // =======================================================================================

  /**
   * ПУБЛИЧНАЯ: Устанавливает триггер для запуска архивирования в полночь.
   * Вызывается один раз при настройке проекта.
   */
  Lib.setupMidnightLogTrigger = function () {
    const ui = SpreadsheetApp.getUi();

    // Удаляем существующие триггеры архивирования
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(function (trigger) {
      if (trigger.getHandlerFunction() === "midnightLogRotation_trigger") {
        ScriptApp.deleteTrigger(trigger);
      }
    });

    // Создаём новый триггер на полночь
    ScriptApp.newTrigger("midnightLogRotation_trigger")
      .timeBased()
      .atHour(0)
      .everyDays(1)
      .inTimezone("Europe/Moscow")
      .create();

    ui.alert(
      "✅ Триггер установлен",
      "Архивирование логов будет выполняться каждый день в полночь (МСК).\n\n" +
      "Логи будут сохраняться в месячные таблицы в папке архивов.",
      ui.ButtonSet.OK
    );
  };

  /**
   * ПУБЛИЧНАЯ: Удаляет триггер полночного архивирования.
   */
  Lib.removeMidnightLogTrigger = function () {
    const triggers = ScriptApp.getProjectTriggers();
    let removed = 0;

    triggers.forEach(function (trigger) {
      if (trigger.getHandlerFunction() === "midnightLogRotation_trigger") {
        ScriptApp.deleteTrigger(trigger);
        removed++;
      }
    });

    SpreadsheetApp.getUi().alert(
      "Триггеры удалены",
      `Удалено ${removed} триггеров архивирования.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  };

  // =======================================================================================
  // ФУНКЦИИ ДЛЯ ВСЕОБЪЕМЛЮЩЕГО ЛОГИРОВАНИЯ
  // =======================================================================================

  /**
   * Универсальная функция логирования
   * ОБНОВЛЕНО: Использует новую систему LOG_DEBUG вместо старого листа "Логи"
   * @param {string} category - Категория действия
   * @param {string} action - Описание действия
   * @param {string} level - Уровень (INFO, WARN, ERROR, DEBUG)
   * @param {string} [details] - Дополнительные детали
   */
  function _logToSheet_(category, action, level, details) {
    try {
      // Используем новую систему логирования через Lib.logWithEmoji
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        const message = action + (details ? ": " + details : "");
        Lib.logWithEmoji(message, level, "", category, details || "");
      } else {
        // Fallback: просто логируем в консоль если Lib недоступен
        console.log(`[${category}] ${action}: ${details || ""}`);
      }
    } catch (e) {
      console.error("_logToSheet_ error:", e);
    }
  }

  /**
   * ПУБЛИЧНАЯ: Логирует выполнение функции с полным контекстом.
   * Используйте в начале каждой важной функции.
   * @param {string} functionName - Имя функции
   * @param {string} [context] - Контекст/аргументы
   * @param {string} [source] - Источник вызова (menu, trigger, api)
   */
  Lib.logFunctionCall = function (functionName, context, source) {
    const category = source || "FUNCTION";
    const action = `Вызов: ${functionName}`;
    const details = context ? `Контекст: ${context}` : "";
    _logToSheet_(category, action, "INFO", details);
  };

  /**
   * ПУБЛИЧНАЯ: Логирует успешное завершение функции.
   * @param {string} functionName - Имя функции
   * @param {string} [result] - Результат выполнения
   * @param {number} [durationMs] - Время выполнения в мс
   */
  Lib.logFunctionSuccess = function (functionName, result, durationMs) {
    const action = `Завершено: ${functionName}`;
    const details = [];
    if (result) details.push(`Результат: ${result}`);
    if (durationMs) details.push(`Время: ${durationMs}ms`);
    _logToSheet_("FUNCTION", action, "INFO", details.join(" | "));
  };

  /**
   * ПУБЛИЧНАЯ: Логирует ошибку функции.
   * @param {string} functionName - Имя функции
   * @param {Error|string} error - Ошибка
   */
  Lib.logFunctionError = function (functionName, error) {
    const action = `Ошибка: ${functionName}`;
    const details = error instanceof Error ? `${error.message}\n${error.stack}` : String(error);
    _logToSheet_("FUNCTION", action, "ERROR", details);
  };

  /**
   * ПУБЛИЧНАЯ: Обёртка для автоматического логирования функции.
   * @param {string} name - Имя функции
   * @param {Function} fn - Функция для выполнения
   * @param {string} [source] - Источник вызова
   * @returns {*} Результат функции
   */
  Lib.withLogging = function (name, fn, source) {
    const startTime = Date.now();
    Lib.logFunctionCall(name, null, source);

    try {
      const result = fn();
      const duration = Date.now() - startTime;
      Lib.logFunctionSuccess(name, typeof result === "object" ? "Object" : String(result).substring(0, 50), duration);
      return result;
    } catch (e) {
      Lib.logFunctionError(name, e);
      throw e;
    }
  };

  /**
   * ПУБЛИЧНАЯ: Логирует событие редактирования
   * @param {Object} e - Событие onEdit
   */
  Lib.logEditEvent = function (e) {
    if (!e || !e.range) return;

    const sheet = e.range.getSheet();
    const action = `Редактирование: ${sheet.getName()} [${e.range.getA1Notation()}]`;
    const details = `Старое: "${e.oldValue || ""}" → Новое: "${e.value || ""}"`;
    _logToSheet_("EDIT", action, "INFO", details);
  };

  /**
   * ПУБЛИЧНАЯ: Логирует системное событие
   * @param {string} event - Тип события
   * @param {string} details - Детали
   */
  Lib.logSystemEvent = function (event, details) {
    _logToSheet_("SYSTEM", event, "INFO", details);
  };

  /**
   * ПУБЛИЧНАЯ: Логирует событие синхронизации
   * @param {string} action - Действие
   * @param {string} details - Детали
   */
  Lib.logSyncEvent = function (action, details) {
    _logToSheet_("SYNC", action, "INFO", details);
  };

  /**
   * ПУБЛИЧНАЯ: Логирует событие API
   * @param {string} endpoint - Эндпоинт
   * @param {string} method - Метод (GET, POST)
   * @param {number} statusCode - Код ответа
   * @param {string} [details] - Детали
   */
  Lib.logApiEvent = function (endpoint, method, statusCode, details) {
    const action = `API ${method}: ${endpoint}`;
    const detailsStr = `Статус: ${statusCode}${details ? " | " + details : ""}`;
    const level = statusCode >= 400 ? "ERROR" : "INFO";
    _logToSheet_("API", action, level, detailsStr);
  };

  // =======================================================================================
  // МЕНЮ ФУНКЦИИ
  // =======================================================================================

  /**
   * ПУБЛИЧНАЯ: Ручной запуск архивирования (из меню)
   */
  Lib.manualArchiveLogs = function () {
    const ui = SpreadsheetApp.getUi();

    const confirm = ui.alert(
      "Архивирование логов",
      "Скопировать все текущие логи в месячный архив?\n\n" +
      "Логи будут скопированы в папку архивов, лист 'Журнал логов' будет очищен.",
      ui.ButtonSet.YES_NO
    );

    if (confirm !== ui.Button.YES) return;

    SpreadsheetApp.getActiveSpreadsheet().toast("⏳ Архивирование...", "Логи", 30);

    const result = Lib.midnightLogRotation();

    if (result.success || result.skipped) {
      ui.alert(
        "✅ Архивирование завершено",
        result.skipped
          ? "Логи уже были архивированы сегодня."
          : `Скопировано ${result.details.totalRows} записей.\n\nАрхив: ${result.details.archiveSheet}`,
        ui.ButtonSet.OK
      );
    } else {
      ui.alert("❌ Ошибка", result.error, ui.ButtonSet.OK);
    }
  };

  /**
   * ПУБЛИЧНАЯ: Показывает статус архивирования
   */
  Lib.showArchiveStatus = function () {
    const ui = SpreadsheetApp.getUi();
    const props = PropertiesService.getScriptProperties();

    const lastArchive = props.getProperty(LOG_ARCHIVE_CONFIG.LAST_ARCHIVE_KEY) || "Никогда";
    const projectCode = _getProjectCode_();
    const now = new Date();
    const monthName = LOG_ARCHIVE_CONFIG.MONTHS_RU[now.getMonth()];
    const currentArchiveName = `${projectCode}-${monthName}-${now.getFullYear()}`;

    // Проверяем триггеры
    const triggers = ScriptApp.getProjectTriggers();
    const hasMiddnightTrigger = triggers.some(function (t) {
      return t.getHandlerFunction() === "midnightLogRotation_trigger";
    });

    let msg = "📊 СТАТУС АРХИВИРОВАНИЯ ЛОГОВ\n\n";
    msg += `📅 Последнее архивирование: ${lastArchive}\n`;
    msg += `📁 Текущий архив: ${currentArchiveName}\n`;
    msg += `⏰ Автоматический триггер: ${hasMiddnightTrigger ? "✅ Установлен (полночь МСК)" : "❌ Не установлен"}\n`;
    msg += `📂 Папка архивов: ${LOG_ARCHIVE_CONFIG.ARCHIVE_FOLDER_ID}\n`;

    ui.alert("Статус архивирования", msg, ui.ButtonSet.OK);
  };

})(this);

// =======================================================================================
// ГЛОБАЛЬНЫЕ ПРОКСИ-ФУНКЦИИ (для триггеров и меню)
// =======================================================================================

/**
 * Прокси для триггера полночного архивирования.
 * НЕ ПЕРЕИМЕНОВЫВАТЬ - привязана к триггеру!
 */
function midnightLogRotation_trigger() {
  if (typeof Lib !== "undefined" && Lib.midnightLogRotation) {
    return Lib.midnightLogRotation();
  }
}

/**
 * Прокси для ручного архивирования из меню
 */
function manualArchiveLogs_proxy() {
  if (typeof callServerArchiveLogs === 'function') {
    return callServerArchiveLogs();
  }
  // Fallback to local if server auth missing or undefined, but for migration we prefer server
  if (typeof Lib !== "undefined" && Lib.manualArchiveLogs) {
    return Lib.manualArchiveLogs();
  }
}

/**
 * Прокси для установки триггера из меню
 */
function setupMidnightLogTrigger_proxy() {
  if (typeof Lib !== "undefined" && Lib.setupMidnightLogTrigger) {
    return Lib.setupMidnightLogTrigger();
  }
}

/**
 * Прокси для удаления триггера из меню
 */
function removeMidnightLogTrigger_proxy() {
  if (typeof Lib !== "undefined" && Lib.removeMidnightLogTrigger) {
    return Lib.removeMidnightLogTrigger();
  }
}

/**
 * Прокси для показа статуса архивирования
 */
function showArchiveStatus_proxy() {
  if (typeof callServerShowArchiveStatus === 'function') {
    return callServerShowArchiveStatus();
  }
  if (typeof Lib !== "undefined" && Lib.showArchiveStatus) {
    return Lib.showArchiveStatus();
  }
}
