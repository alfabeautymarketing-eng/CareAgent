/* ОПИСАНИЕ:
 *   Этот файл содержит всю основную логику работы библиотеки. Он разделен на модули
 *   для обеспечения читаемости и простоты поддержки.
 * =======================================================================================
 */

// Создаем глобальный объект-контейнер для всей логики библиотеки,
// если он еще не был создан в A_Config.gs.
var Lib = Lib || {};

// =======================================================================================
// ПАТЧ СОВМЕСТИМОСТИ И ЗАЩИТНЫЕ ДЕФОЛТЫ
// ---------------------------------------------------------------------------------------
// Описание: Этот блок выполняется до основной логики, чтобы обеспечить совместимость
//           и предотвратить ошибки, если CONFIG неполный или имеет старую структуру.
// =======================================================================================
(function (Lib, global) {
  // 1. Мост для совместимости со старым экспортом `global.CONFIG`
  if (!Lib.CONFIG && typeof global.CONFIG !== "undefined") {
    Lib.CONFIG = global.CONFIG;
  }
  // 2. Создаем пустые заглушки, чтобы код не падал, если эти разделы отсутствуют в конфиге.
  Lib.CONFIG.OVERRIDES = Lib.CONFIG.OVERRIDES || {};
  Lib.CONFIG.CASCADE_RULES = Lib.CONFIG.CASCADE_RULES || {};
})(Lib, this);

// =======================================================================================
// ОСНОВНОЙ КОД БИБЛИОТЕКИ
// ---------------------------------------------------------------------------------------
// Описание: Вся логика инкапсулирована в самовызывающуюся функцию,
//           чтобы избежать загрязнения глобального пространства имен.
// =======================================================================================
(function (Lib, global) {
  // =======================================================================================
  // МОДУЛЬ 1: ГЛОБАЛЬНЫЕ УТИЛИТЫ
  // ---------------------------------------------------------------------------------------
  // Описание: Набор надежных вспомогательных функций, которые используются
  //           во всей библиотеке. Они обеспечивают логирование, кэширование
  //           и эффективную работу с данными таблиц.
  // =======================================================================================

  // ----------------------------------
  // Секция 1.1: Система логирования
  // ----------------------------------

  /**
   * @private Универсальная функция для записи логов с учетом уровня из CONFIG.
   * @param {string} levelString Уровень лога ('DEBUG', 'INFO', 'WARN', 'ERROR').
   * @param {string} message Сообщение для записи.
   * @param {Error|Object} [errorObject=null] Необязательный объект ошибки для вывода стека.
   */
  function _customLog(levelString, message, errorObject = null) {
    // Получаем числовой уровень для текущего сообщения.
    const level = Lib.CONFIG.SETTINGS.LOG_LEVELS[levelString.toUpperCase()];
    // Сравниваем с уровнем, установленным в конфиге. Если уровень сообщения ниже, не логируем.
    if (level === undefined || level < Lib.CONFIG.SETTINGS.CURRENT_LOG_LEVEL) {
      return;
    }
    // Форматируем временную метку для лога.
    const timestamp = Utilities.formatDate(
      new Date(),
      Lib.CONFIG.SETTINGS.TIMEZONE,
      "yyyy-MM-dd HH:mm:ss"
    );
    // Собираем основное сообщение лога.
    let logMessage = `${message}`;
    let details = '';
    // Если передан объект ошибки, добавляем его детали.
    if (errorObject) {
      details = errorObject.stack || errorObject.message || JSON.stringify(errorObject);
    }
    // Используем новую функцию логирования с эмодзи
    if (typeof global.logWithEmoji === 'function') {
      global.logWithEmoji(logMessage, levelString, null, 'Синхронизация', details);
    } else {
      // Fallback на стандартный логгер
      Logger.log(`[${levelString}] ${timestamp} :: ${logMessage}${details ? '\n  >>> Детали ошибки: ' + details : ''}`);
    }
  }

  /** Записывает сообщение уровня DEBUG. Видно только при LOG_LEVEL = 'DEBUG'. */
  Lib.logDebug = function (message, errorObject = null) {
    _customLog("DEBUG", message, errorObject);
  };
  /** Записывает информационное сообщение уровня INFO. */
  Lib.logInfo = function (message, errorObject = null) {
    _customLog("INFO", message, errorObject);
  };
  /** Записывает предупреждение уровня WARN. */
  Lib.logWarn = function (message, errorObject = null) {
    _customLog("WARN", message, errorObject);
  };
  /** Записывает сообщение об ошибке уровня ERROR. */
  Lib.logError = function (message, errorObject = null) {
    _customLog("ERROR", message, errorObject);
  };

  // ----------------------------------
  // Секция 1.2: Работа с данными
  // ----------------------------------

  /**
   * Умное сравнение двух значений. Корректно обрабатывает null, undefined,
   * пустые строки, числа, текст и даты.
   * @private
   * @param {*} v1 Первое значение.
   * @param {*} v2 Второе значение.
   * @returns {boolean} True, если значения эквивалентны.
   */
  // ----------------------------------
  // Секция 1.2.1: Хелперы "по именам заголовков"
  // ----------------------------------

  /**
   * Читает указанную строку как объект { "Имя столбца": значение } по ФАКТИЧЕСКИМ заголовкам листа.
   * Не зависит от порядка колонок и CONFIG.HEADERS.
   */
  Lib.readRowAsObject_ = function (sheet, row) {
    const lastCol = sheet.getLastColumn();
    if (lastCol < 1) return {};
    const headers = sheet
      .getRange(1, 1, 1, lastCol)
      .getValues()[0]
      .map((h) => String(h || "").trim());
    const values = sheet.getRange(row, 1, 1, lastCol).getValues()[0];
    const obj = {};
    headers.forEach((h, i) => {
      if (h) obj[h] = values[i];
    });
    return obj;
  };

  /**
   * Безопасно записывает значение в ячейку по ИМЕНИ столбца.
   * Записывает только если значение действительно изменилось.
   * @returns {boolean} true, если запись была выполнена.
   */
  Lib.setCellIfChangedByHeader_ = function (sheet, row, headerName, newValue) {
    const col = Lib.findColumnIndexByHeader_(
      sheet,
      String(headerName || "").trim()
    );
    if (col <= 0) return false;
    const prev = sheet.getRange(row, col).getValue();
    if (Lib.areValuesEqual_(prev, newValue)) return false;
    sheet.getRange(row, col).setValue(newValue);
    return true;
  };

  /**
   * Гарантирует наличие требуемых заголовков на листе, добавляя недостающие В КОНЕЦ.
   * Существующий порядок НЕ меняется.
   */
  Lib.ensureHeadersPresent_ = function (sheet, requiredHeaders) {
    if (!requiredHeaders || requiredHeaders.length === 0) return;
    const lastCol = sheet.getLastColumn();
    const current =
      lastCol > 0
        ? sheet
            .getRange(1, 1, 1, lastCol)
            .getValues()[0]
            .map((h) => String(h || "").trim())
        : [];
    const missing = requiredHeaders.filter((h) => h && !current.includes(h));
    if (missing.length === 0) return;
    sheet.insertColumnsAfter(Math.max(1, lastCol), missing.length);
    sheet
      .getRange(1, lastCol + 1, 1, missing.length)
      .setValues([missing])
      .setFontWeight("bold");
  };

  Lib.areValuesEqual_ = function (v1, v2) {
    // Внутренняя функция для приведения любого значения к единому формату.
    const normalize = (val) => {
      // null, undefined и строки из пробелов считаем эквивалентными пустой строке.
      if (
        val === null ||
        val === undefined ||
        (typeof val === "string" && val.trim() === "")
      )
        return "";
      // Даты сравниваем по их числовому представлению (миллисекунды).
      if (val instanceof Date) return isNaN(val.getTime()) ? "" : val.getTime();
      return val;
    };
    // Сравниваем нормализованные значения как строки, чтобы избежать проблем с типами (например, 123 и "123").
    return String(normalize(v1)) === String(normalize(v2));
  };

  /** Кэш в оперативной памяти для хранения карты "Имя заголовка -> номер столбца". Живет только во время одного запуска скрипта. */
  const _headerCache = {};

  /**
   * Находит номер столбца (начиная с 1) по его текстовому заголовку.
   * Использует кэш в памяти для максимальной скорости при повторных вызовах.
   * @private
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Лист, на котором ищем столбец.
   * @param {string} headerName Имя искомого заголовка.
   * @returns {number} Номер столбца или -1, если не найден.
   */
  Lib.findColumnIndexByHeader_ = function (sheet, headerName) {
    const sheetName = sheet.getName();
    const trimmedHeaderName = headerName.trim();
    // 1. Пытаемся быстро вернуть результат из кэша.
    if (
      _headerCache[sheetName] &&
      _headerCache[sheetName][trimmedHeaderName] !== undefined
    ) {
      return _headerCache[sheetName][trimmedHeaderName];
    }
    // 2. Если в кэше нет, читаем первую строку (заголовки) с листа.
    const headerRow = sheet
      .getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0];
    const headerMap = {};
    headerRow.forEach((header, index) => {
      const trimmedHeader = String(header || "").trim();
      if (trimmedHeader) {
        // Игнорируем пустые заголовки
        headerMap[trimmedHeader] = index + 1; // Запоминаем 1-based индекс
      }
    });
    // 3. Сохраняем всю карту заголовков для этого листа в кэш.
    _headerCache[sheetName] = headerMap;
    // 4. Возвращаем результат из свежесозданной карты.
    return _headerCache[sheetName][trimmedHeaderName] || -1;
  };

  // ----------------------------------------------------
  // Секция 1.3: Работа с долговременным кэшем (CacheService)
  // ----------------------------------------------------

  /**
   * Получает или создает карту "ключ (ID) -> номер строки" для листа.
   * Использует CacheService для хранения данных между запусками скрипта.
   * @private
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Лист для анализа.
   * @param {boolean} [forceRead=false] Если true, кэш будет проигнорирован и данные будут прочитаны с листа.
   * @returns {Map<string, number>} Карта, где ключ - это ID из столбца А, а значение - номер строки.
   */
  function _getKeyMapForSheet(sheet, forceRead = false) {
    const sheetName = sheet.getName();
    const cache = CacheService.getScriptCache();
    const cacheKey = Lib.CONFIG.SETTINGS.CACHE_KEY_PREFIX + sheetName;
    // 1. Пытаемся достать карту из кэша, если это не принудительное чтение.
    if (!forceRead) {
      const cachedData = cache.get(cacheKey);
      if (cachedData) {
        try {
          // Превращаем строку из кэша обратно в объект Map.
          return new Map(JSON.parse(cachedData));
        } catch (e) {
          Lib.logWarn(
            `Кэш для листа "${sheetName}" поврежден. Будет произведено перечитывание.`
          );
        }
      }
    }
    // 2. Если в кэше нет (или он поврежден/игнорируется), читаем данные с листа.
    Lib.logInfo(`Читаем ключи с листа "${sheetName}" для построения кэша.`);
    const headerRows = 1; // Мы договорились, что заголовки всегда на первой строке.
    const lastRow = sheet.getLastRow();
    const newKeyMap = new Map();
    // Если на листе есть данные кроме заголовков...
    if (lastRow > headerRows) {
      // ...читаем все значения из первого столбца за один вызов API.
      const keys = sheet
        .getRange(headerRows + 1, 1, lastRow - headerRows, 1)
        .getValues();
      for (let i = 0; i < keys.length; i++) {
        const currentKey = String(keys[i][0]).trim();
        if (currentKey) {
          // Игнорируем пустые строки
          // Записываем в карту: ключ и его реальный номер строки.
          newKeyMap.set(currentKey, headerRows + 1 + i);
        }
      }
    }
    // 3. Сохраняем новую карту в кэш на будущее.
    try {
      cache.put(
        cacheKey,
        JSON.stringify(Array.from(newKeyMap.entries())),
        Lib.CONFIG.SETTINGS.CACHE_EXPIRATION_SECONDS
      );
      Lib.logDebug(
        `Кэш для "${sheetName}" обновлен. Записей: ${newKeyMap.size}.`
      );
    } catch (e) {
      Lib.logError(`Не удалось сохранить кэш для "${sheetName}"`, e);
    }
    return newKeyMap;
  }

  /**
   * Находит номер строки по значению ключа в столбце 'A'. Это основная
   * функция для поиска строк по ID.
   * @private
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Лист, на котором ищем.
   * @param {string|number} key Искомый ключ (ID).
   * @param {boolean} [forceRead=false] Если true, кэш будет проигнорирован.
   * @returns {number} Номер строки (начиная с 1) или -1, если не найдено.
   */
  Lib.findRowByKey_ = function (sheet, key, forceRead = false) {
    if (!sheet || key === null || key === undefined) return -1;
    const stringKey = String(key).trim();
    if (!stringKey) return -1;
    // Получаем карту "ключ -> номер строки" и ищем в ней.
    const keyMap = _getKeyMapForSheet(sheet, forceRead);
    return keyMap.get(stringKey) || -1;
  };

  /**
   * Полностью удаляет кэш ключей для указанного листа.
   * Обязательно вызывать после массового добавления или удаления строк.
   */
  Lib.deleteKeyCacheForSheet = function (sheetName) {
    const cache = CacheService.getScriptCache();
    const cacheKey = Lib.CONFIG.SETTINGS.CACHE_KEY_PREFIX + sheetName;
    cache.remove(cacheKey);
    Lib.logDebug(`Кэш для листа "${sheetName}" был сброшен.`);
  };

  // =======================================================================================
  // МОДУЛЬ 2: ЯДРО СИНХРОНИЗАЦИИ (onEdit)
  // ---------------------------------------------------------------------------------------
  // Обрабатывает изменения, ищет правила, копирует значения, дергает каскады и пишет журнал.
  // =======================================================================================

  /** Кэш правил на время одного выполнения */
  let _cachedSyncRules = null;

  /**
   * Главный обработчик onEdit (вызов из обёртки: EcosystemLib.Lib.onEdit_internal_(e))
   */
  Lib.onEdit_internal_ = function (e) {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(Lib.CONFIG.SETTINGS.LOCK_TIMEOUT_MS)) {
      Lib.logWarn("onEdit: занято другим запуском, пропускаем.");
      return;
    }
    try {
      if (!e || !e.range) return;

      const range = e.range;
      const sheet = range.getSheet();
      const sheetName = sheet.getName();
      const notation = range.getA1Notation();
      const rowsChanged = range.getNumRows();
      const colsChanged = range.getNumColumns();
      const valuePreview =
        typeof e.value !== "undefined" ? JSON.stringify(e.value) : "<multi>";
      const oldValuePreview =
        typeof e.oldValue !== "undefined"
          ? JSON.stringify(e.oldValue)
          : "<multi>";
      Lib.logDebug(
        '[onEdit] sheet="' +
          sheetName +
          '", range=' +
          notation +
          ", rowStart=" +
          range.getRow() +
          ", colStart=" +
          range.getColumn() +
          ", rows=" +
          rowsChanged +
          ", cols=" +
          colsChanged +
          ", value=" +
          valuePreview +
          ", oldValue=" +
          oldValuePreview
      );

      // Обрабатываем изменения на листе "ТЗ по статусам" (чекбокс "Отработано")
      if (typeof Lib.handleTaskSheetEdit === 'function') {
        Lib.handleTaskSheetEdit(e);
      }

      // Игнорируем служебные листы; при правке правил сбрасываем кэш
      const serviceSheetNames = [
        Lib.CONFIG.SHEETS.RULES,
        Lib.CONFIG.SHEETS.LOG,
        Lib.CONFIG.SHEETS.EXTERNAL_DOCS,
      ];
      if (serviceSheetNames.includes(sheetName)) {
        if (sheetName === Lib.CONFIG.SHEETS.RULES) {
          _cachedSyncRules = null;
          Lib.logInfo("onEdit: правила изменились — кэш очищен.");
        }
        return;
      }

      _processEditEvent(e);
    } catch (err) {
      Lib.logError("onEdit_internal_: критическая ошибка", err);
    } finally {
      lock.releaseLock();
    }
  };

  /**
   * Установка/сброс триггеров проекта.
   * Удаляет старые триггеры и создает новые (handleOnChange + handleOnEdit).
   */
  Lib.setupTriggers = function () {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const triggers = ScriptApp.getProjectTriggers();
      let count = 0;
      
      // Удаляем старые триггеры, связанные с нашими функциями
      triggers.forEach((t) => {
        const handler = t.getHandlerFunction();
        if (handler === "handleOnChange" || handler === "handleOnEdit") {
          ScriptApp.deleteTrigger(t);
          count++;
        }
      });
      
      // 1. Installable onEdit (для надежной обработки, вставок и т.д.)
      ScriptApp.newTrigger("handleOnEdit")
        .forSpreadsheet(ss)
        .onEdit()
        .create();

      // 2. Installable onChange (для структурных изменений)
      ScriptApp.newTrigger("handleOnChange")
        .forSpreadsheet(ss)
        .onChange()
        .create();
        
      Lib.logInfo(`[SetupTriggers] Удалено старых: ${count}. Созданы новые: handleOnEdit, handleOnChange.`);
      return `Триггеры переустановлены. (Удалено: ${count}, Создано: 2)`;
    } catch (e) {
      Lib.logError("setupTriggers: ошибка", e);
      throw e;
    }
  };

  /**
   * Обработчик onChange (installable trigger)
   */
  Lib.handleOnChange = function (e) {
    try {
      const changeType = e && e.changeType ? String(e.changeType) : "UNKNOWN";
Lib.logDebug(`[onChange] ${changeType}`);

if (!Lib.CONFIG || !Lib.CONFIG.SHEETS) {
  Lib.logWarn("handleOnChange: CONFIG не загружен");
  return;
}

const ss = (e && e.source) || SpreadsheetApp.getActiveSpreadsheet();
const touchedSheets = [];

if (ss) {
  if (changeType === "REMOVE_SHEET" || changeType === "INSERT_SHEET") {
    Object.values(Lib.CONFIG.SHEETS || {}).forEach((name) => {
      if (name) touchedSheets.push(name);
    });
  } else {
    const activeSheet = ss.getActiveSheet && ss.getActiveSheet();
    if (activeSheet) touchedSheets.push(activeSheet.getName());
  }
}

      if (touchedSheets.length === 0) {
        Object.values(Lib.CONFIG.SHEETS || {}).forEach((name) => {
          if (name) touchedSheets.push(name);
        });
      }

      touchedSheets.filter(Boolean).forEach((name) => {
        try {
          Lib.deleteKeyCacheForSheet(name);
        } catch (err) {
          Lib.logWarn(
            `handleOnChange: не удалось сбросить кэш для "${name}"`,
            err
          );
        }
      });
    } catch (err) {
      Lib.logError("handleOnChange: критическая ошибка", err);
    }
  };

  // =======================================================================================
  // СПЕЦИАЛЬНЫЕ ФУНКЦИИ: Автозаполнение "Срок#" из "СГ 1-3"
  // ---------------------------------------------------------------------------------------
  // Эти функции должны быть объявлены ДО _processEditEvent, так как вызываются из неё
  // =======================================================================================

  /**
   * Автоматически заполняет столбец "Срок#" значениями из столбцов "СГ 1", "СГ 2", "СГ 3"
   * @param {Sheet} sheet - Лист "Заказ"
   * @param {number} row - Номер строки, в которой произошло изменение
   * @private
   */
  function _autoFillDeadlineFromExpiry(sheet, row) {
    try {
      // Получаем заголовки из строки 1 (на листе "Заказ" заголовки всегда в строке 1)
      var lastColumn = sheet.getLastColumn();
      var HEADER_ROW = 1;
      var headers = sheet.getRange(HEADER_ROW, 1, 1, lastColumn).getValues()[0];

      // Ищем индексы нужных столбцов
      var sg1Index = -1;
      var sg2Index = -1;
      var sg3Index = -1;
      var deadlineIndex = -1;

      // Функция нормализации заголовка (убирает лишние пробелы и приводит к нижнему регистру)
      function normalizeHeader(h) {
        return String(h || "")
          .trim()
          .replace(/\s+/g, " ")
          .toLowerCase();
      }

      for (var i = 0; i < headers.length; i++) {
        var header = String(headers[i] || "").trim();
        var normalized = normalizeHeader(header);

        if (normalized === normalizeHeader("СГ 1")) sg1Index = i;
        else if (normalized === normalizeHeader("СГ 2")) sg2Index = i;
        else if (normalized === normalizeHeader("СГ 3")) sg3Index = i;
        else if (normalized === normalizeHeader("Срок#")) deadlineIndex = i;
      }

      // Проверяем, что нашли все необходимые столбцы
      if (deadlineIndex === -1) {
        // Логируем все заголовки для отладки
        var headersDebug = [];
        for (var j = 0; j < Math.min(headers.length, 40); j++) {
          headersDebug.push(
            'Col' + (j + 1) + ':"' + String(headers[j] || "").trim() + '"'
          );
        }
        Lib.logWarn(
          '[_autoFillDeadlineFromExpiry] Не найден столбец "Срок#" на листе "Заказ". Заголовки: ' +
            headersDebug.join(", ")
        );
        return;
      }

      if (sg1Index === -1 && sg2Index === -1 && sg3Index === -1) {
        Lib.logWarn(
          '[_autoFillDeadlineFromExpiry] Не найдены столбцы "СГ 1", "СГ 2", "СГ 3" на листе "Заказ"'
        );
        return;
      }

      Lib.logDebug(
        '[_autoFillDeadlineFromExpiry] Индексы столбцов: СГ 1=' +
          sg1Index +
          ", СГ 2=" +
          sg2Index +
          ", СГ 3=" +
          sg3Index +
          ", Срок#=" +
          deadlineIndex
      );

      // Читаем значения из столбцов СГ 1-3 для текущей строки
      var rowData = sheet.getRange(row, 1, 1, lastColumn).getValues()[0];

      var expiryDates = [];

      // Собираем непустые значения из СГ 1-3
      if (sg1Index !== -1) {
        var sg1Value = rowData[sg1Index];
        if (sg1Value && String(sg1Value).trim() !== "") {
          expiryDates.push(_formatExpiryDate(sg1Value));
        }
      }

      if (sg2Index !== -1) {
        var sg2Value = rowData[sg2Index];
        if (sg2Value && String(sg2Value).trim() !== "") {
          expiryDates.push(_formatExpiryDate(sg2Value));
        }
      }

      if (sg3Index !== -1) {
        var sg3Value = rowData[sg3Index];
        if (sg3Value && String(sg3Value).trim() !== "") {
          expiryDates.push(_formatExpiryDate(sg3Value));
        }
      }

      // Формируем итоговое значение для "Срок#" (каждое значение с новой строки)
      var deadlineValue = expiryDates.join("\n");

      Lib.logInfo(
        '[_autoFillDeadlineFromExpiry] Заполняю "Срок#" в строке ' +
          row +
          ' значением: "' +
          deadlineValue +
          '"'
      );

      // Записываем значение в ячейку "Срок#"
      if (deadlineValue) {
        sheet.getRange(row, deadlineIndex + 1).setValue(deadlineValue);
      } else {
        // Если нет значений в СГ 1-3, очищаем ячейку "Срок#"
        sheet.getRange(row, deadlineIndex + 1).clearContent();
        Lib.logDebug(
          '[_autoFillDeadlineFromExpiry] Нет значений в СГ 1-3, очищаем "Срок#"'
        );
      }
    } catch (error) {
      Lib.logError(
        "[_autoFillDeadlineFromExpiry] Ошибка при заполнении срока",
        error
      );
      throw error;
    }
  }

  /**
   * Форматирует дату срока годности в формат MM.YYYY
   * @param {Date|string} value - Значение даты (Date объект или строка)
   * @returns {string} Отформатированная дата в формате MM.YYYY
   * @private
   */
  function _formatExpiryDate(value) {
    try {
      if (value instanceof Date) {
        var month = String(value.getMonth() + 1).padStart(2, "0");
        var year = value.getFullYear();
        return month + "." + year;
      }

      var str = String(value).trim();

      // Если уже в формате MM.YYYY или MM/YYYY
      var match = str.match(/^(\d{1,2})[\./](\d{4})$/);
      if (match) {
        var m = String(match[1]).padStart(2, "0");
        return m + "." + match[2];
      }

      // Пробуем распарсить как дату
      var dateObj = new Date(value);
      if (!isNaN(dateObj.getTime())) {
        var month2 = String(dateObj.getMonth() + 1).padStart(2, "0");
        var year2 = dateObj.getFullYear();
        return month2 + "." + year2;
      }

      // Если не удалось распознать, возвращаем как есть
      return str;
    } catch (e) {
      // В случае ошибки возвращаем строковое представление
      return String(value);
    }
  }

  /**
   * Автоматически заполняет столбец "Срок" значениями из столбцов "СГ 1", "СГ 2", "СГ 3"
   * Вызывается при изменении столбца "Набор"
   * @param {Sheet} sheet - Лист "Заказ"
   * @param {number} row - Номер строки, в которой произошло изменение
   * @private
   */
  function _autoFillSetDeadlineFromExpiry(sheet, row) {
    try {
      // Получаем заголовки из строки 1 (на листе "Заказ" заголовки всегда в строке 1)
      var lastColumn = sheet.getLastColumn();
      var HEADER_ROW = 1;
      var headers = sheet.getRange(HEADER_ROW, 1, 1, lastColumn).getValues()[0];

      // Ищем индексы нужных столбцов
      var sg1Index = -1;
      var sg2Index = -1;
      var sg3Index = -1;
      var deadlineIndex = -1;

      // Функция нормализации заголовка (убирает лишние пробелы и приводит к нижнему регистру)
      function normalizeHeader(h) {
        return String(h || "")
          .trim()
          .replace(/\s+/g, " ")
          .toLowerCase();
      }

      for (var i = 0; i < headers.length; i++) {
        var header = String(headers[i] || "").trim();
        var normalized = normalizeHeader(header);

        if (normalized === normalizeHeader("СГ 1")) sg1Index = i;
        else if (normalized === normalizeHeader("СГ 2")) sg2Index = i;
        else if (normalized === normalizeHeader("СГ 3")) sg3Index = i;
        else if (normalized === normalizeHeader("Срок")) deadlineIndex = i;
      }

      // Проверяем, что нашли все необходимые столбцы
      if (deadlineIndex === -1) {
        // Логируем все заголовки для отладки
        var headersDebug = [];
        for (var j = 0; j < Math.min(headers.length, 40); j++) {
          headersDebug.push(
            'Col' + (j + 1) + ':"' + String(headers[j] || "").trim() + '"'
          );
        }
        Lib.logWarn(
          '[_autoFillSetDeadlineFromExpiry] Не найден столбец "Срок" на листе "Заказ". Заголовки: ' +
            headersDebug.join(", ")
        );
        return;
      }

      if (sg1Index === -1 && sg2Index === -1 && sg3Index === -1) {
        Lib.logWarn(
          '[_autoFillSetDeadlineFromExpiry] Не найдены столбцы "СГ 1", "СГ 2", "СГ 3" на листе "Заказ"'
        );
        return;
      }

      Lib.logDebug(
        '[_autoFillSetDeadlineFromExpiry] Индексы столбцов: СГ 1=' +
          sg1Index +
          ", СГ 2=" +
          sg2Index +
          ", СГ 3=" +
          sg3Index +
          ", Срок=" +
          deadlineIndex
      );

      // Читаем значения из столбцов СГ 1-3 для текущей строки
      var rowData = sheet.getRange(row, 1, 1, lastColumn).getValues()[0];

      var expiryDates = [];

      // Собираем непустые значения из СГ 1-3
      if (sg1Index !== -1) {
        var sg1Value = rowData[sg1Index];
        if (sg1Value && String(sg1Value).trim() !== "") {
          expiryDates.push(_formatExpiryDate(sg1Value));
        }
      }

      if (sg2Index !== -1) {
        var sg2Value = rowData[sg2Index];
        if (sg2Value && String(sg2Value).trim() !== "") {
          expiryDates.push(_formatExpiryDate(sg2Value));
        }
      }

      if (sg3Index !== -1) {
        var sg3Value = rowData[sg3Index];
        if (sg3Value && String(sg3Value).trim() !== "") {
          expiryDates.push(_formatExpiryDate(sg3Value));
        }
      }

      // Формируем итоговое значение для "Срок" (каждое значение с новой строки)
      var deadlineValue = expiryDates.join("\n");

      Lib.logInfo(
        '[_autoFillSetDeadlineFromExpiry] Заполняю "Срок" в строке ' +
          row +
          ' значением: "' +
          deadlineValue +
          '"'
      );

      // Записываем значение в ячейку "Срок"
      if (deadlineValue) {
        sheet.getRange(row, deadlineIndex + 1).setValue(deadlineValue);
      } else {
        // Если нет значений в СГ 1-3, очищаем ячейку "Срок"
        sheet.getRange(row, deadlineIndex + 1).clearContent();
        Lib.logDebug(
          '[_autoFillSetDeadlineFromExpiry] Нет значений в СГ 1-3, очищаем "Срок"'
        );
      }
    } catch (error) {
      Lib.logError(
        "[_autoFillSetDeadlineFromExpiry] Ошибка при заполнении срока",
        error
      );
      throw error;
    }
  }

  /**
   * Разбор изменённого диапазона: применение правил + автосоздание строк при новом ID
   * @private
   */
  function _processEditEvent(e) {
    const range = e.range;
    const sheet = range.getSheet();
    const sheetName = sheet.getName();
    const isOrderForm = sheetName === Lib.CONFIG.SHEETS.ORDER_FORM;
    const isPriceSheet = sheetName === Lib.CONFIG.SHEETS.PRICE;

    // читаем новые значения "оптом"
    const values = range.getValues();
    const richTextValues = range.getRichTextValues();
    const headerRowIndex = 1; // единственная строка заголовков

    for (let r = 0; r < range.getNumRows(); r++) {
      const row = range.getRow() + r;
      if (row <= headerRowIndex && !isOrderForm) continue; // шапку не трогаем

      if (isOrderForm && row <= headerRowIndex) {
        Lib.logDebug(
          "[Order:onEdit] анализ строки заголовков " +
            row +
            " (colStart=" +
            range.getColumn() +
            ")"
        );
      }

      // Новая строка по ID (колонка A была пустой → стала непустой)
      if (range.getColumn() === 1 && String(e.oldValue || "").trim() === "") {
        const newKey = String(e.value || "").trim();
        if (newKey) _ensureRowExistsOnBaseSheets(newKey);
      }

      // Основная синхронизация по правилам для каждой затронутой колонки
      for (let c = 0; c < range.getNumColumns(); c++) {
        const col = range.getColumn() + c;
        // Чтение верхней подписи: сначала value, если пусто — displayValue, учесть merged ranges
        var header = String(sheet.getRange(1, col).getValue() || "").trim();
        if (!header) {
          try {
            header = String(
              sheet.getRange(1, col).getDisplayValue() || ""
            ).trim();
          } catch (e) {
            /* ignore */
          }
        }
        var dataHeader = header;
        if (isOrderForm && row > headerRowIndex) {
          try {
            dataHeader = String(
              sheet.getRange(headerRowIndex, col).getValue() || ""
            ).trim();
            if (!dataHeader) {
              // fallback to displayValue
              try {
                dataHeader = String(
                  sheet.getRange(headerRowIndex, col).getDisplayValue() || ""
                ).trim();
              } catch (e) {
                /* ignore */
              }
            }
          } catch (e) {
            dataHeader = header;
          }
        }
        const headerForRules = header || dataHeader;

        if (isOrderForm && row <= headerRowIndex) {
          Lib.logDebug(
            "[Order:onEdit] R" + row + "C" + col + ' header="' + header + '"'
          );
        } else if (isOrderForm && row > headerRowIndex) {
          Lib.logDebug(
            "[Order:onEdit] R" +
              row +
              "C" +
              col +
              ' header1="' +
              header +
              '" dataHeader="' +
              dataHeader +
              '"'
          );
        }
        // Если всё ещё нет заголовка — попробуем узнать, может ячейка в merged range
        if (!headerForRules) {
          try {
            var mrList = sheet.getMergedRanges();
            for (var mi = 0; mi < mrList.length; mi++) {
              var mr = mrList[mi];
              var r1 = mr.getRow();
              var r2 = mr.getLastRow();
              var c1 = mr.getColumn();
              var c2 = mr.getLastColumn();
              if (1 >= r1 && 1 <= r2 && col >= c1 && col <= c2) {
                // верхняя строка 1 в merged range
                try {
                  headerForRules = String(mr.getValues()[0][0] || "").trim();
                } catch (e) {
                  headerForRules = header;
                }
                break;
              }
              if (
                headerRowIndex >= r1 &&
                headerRowIndex <= r2 &&
                col >= c1 &&
                col <= c2
              ) {
                try {
                  headerForRules = String(mr.getValues()[0][0] || "").trim();
                } catch (e) {
                  headerForRules = header;
                }
                break;
              }
            }
          } catch (e) {
            /* ignore */
          }
        }
        if (!headerForRules) continue;

        const normalizedHeaderForPrice = String(headerForRules || "")
          .trim()
          .replace(/\s+/g, " ")
          .toLowerCase();

        // ============================================================================
        // СПЕЦИАЛЬНАЯ ОБРАБОТКА: Автозаполнение "Срок#" из "СГ 1-3" при выборе "АКЦИИ"
        // ============================================================================
        if (
          isOrderForm &&
          row > headerRowIndex &&
          (headerForRules === "АКЦИИ" || dataHeader === "АКЦИИ")
        ) {
          try {
            var promotionValue = values[r][c];
            if (promotionValue && String(promotionValue).trim() !== "") {
              Lib.logInfo(
                '[Order:onEdit] Обнаружено изменение в столбце "АКЦИИ" (строка ' +
                  row +
                  '), запускаем автозаполнение "Срок#"'
              );
              _autoFillDeadlineFromExpiry(sheet, row);
            }
          } catch (autoFillError) {
            Lib.logError(
              '[Order:onEdit] Ошибка автозаполнения "Срок#" из "СГ 1-3"',
              autoFillError
            );
          }
        }
        // ============================================================================

        // ============================================================================
        // СПЕЦИАЛЬНАЯ ОБРАБОТКА: Автозаполнение "Срок" из "СГ 1-3" при выборе "Набор"
        // ============================================================================
        if (
          isOrderForm &&
          row > headerRowIndex &&
          (headerForRules === "Набор" || dataHeader === "Набор")
        ) {
          try {
            var setValue = values[r][c];
            if (setValue && String(setValue).trim() !== "") {
              Lib.logInfo(
                '[Order:onEdit] Обнаружено изменение в столбце "Набор" (строка ' +
                  row +
                  '), запускаем автозаполнение "Срок"'
              );
              _autoFillSetDeadlineFromExpiry(sheet, row);
            }
          } catch (autoFillError2) {
            Lib.logError(
              '[Order:onEdit] Ошибка автозаполнения "Срок" из "СГ 1-3"',
              autoFillError2
            );
          }
        }
        // ============================================================================

        // В новой модели строка 1 - единственная строка заголовков,
        // строки 2-3 больше не имеют специального значения
        // определяем, нужно ли напрямую вызывать override для Order
        let shouldInvokeOrderOverrideDirectly = false;
        if (isOrderForm) {
          if (row === 1) {
            // Правка заголовка в строке 1
            var topHeader = String(headerForRules || header || "").trim();
            var normalizedTopHeader = topHeader.toLowerCase();
            if (
              normalizedTopHeader === "кол-во месяцев" ||
              /^\s*потребность\s+на\s*\d+\s*месяц/i.test(topHeader)
            ) {
              shouldInvokeOrderOverrideDirectly = true;
            }
          }
          // В новой модели строки 2 и 3 больше не имеют специального значения,
          // их обработка проходит через обычные правила синхронизации
        }

        const allRules = _loadSyncRules();
        const applicable = allRules.filter(
          (rule) =>
            rule &&
            typeof rule === "object" &&
            rule.sourceSheet === sheetName &&
            rule.sourceHeader === headerForRules &&
            rule.targetSheet &&
            rule.targetHeader
        );

        // Диагностическое логирование для всех листов
        if (isOrderForm || applicable.length > 0 || allRules.length > 0) {
          Lib.logDebug(
            '[Sync:onEdit] sheet="' +
              sheetName +
              '", row=' +
              row +
              ', col=' +
              col +
              ', header="' +
              header +
              '", dataHeader="' +
              dataHeader +
              '", headerForRules="' +
              headerForRules +
              '", totalRules=' +
              allRules.length +
              ', applicable=' +
              applicable.length
          );
        }
        if (applicable.length === 0) {
          if (shouldInvokeOrderOverrideDirectly) {
            Lib.logDebug(
              '[Order:onEdit] прямой вызов override для "' +
                header +
                '" (row ' +
                row +
                ", col " +
                col +
                ")"
            );
            _invokeOnUpdateOverride(sheet, row, col, "Order:onEdit:noRules");
          } else if (isOrderForm) {
            const headerForTrigger = dataHeader || header;
            const normalizedHeader = String(headerForTrigger || "")
              .trim()
              .replace(/\s+/g, " ")
              .toLowerCase();
            const normalizedTopHeader = String(header || "")
              .trim()
              .replace(/\s+/g, " ")
              .toLowerCase();
            let normalizedHeaderKey = normalizedHeader;
            if (
              normalizedHeaderKey === "да" &&
              normalizedTopHeader === "добавить в прайс"
            ) {
              normalizedHeaderKey = normalizedTopHeader;
            }
            let triggerLabel = null;
            if (row > headerRowIndex) {
              switch (normalizedHeaderKey) {
                case "предварительный заказ":
                  triggerLabel = "preOrder";
                  break;
                case "заказ в упаковках":
                  triggerLabel = "packs";
                  break;
                case "заказ":
                  triggerLabel = "order";
                  break;
                case "шт./уп.":
                  triggerLabel = "unitsPerPack";
                  break;
                case "exw alfaspa текущая, €":
                  triggerLabel = "price";
                  break;
                case "предварительная сумма заказа":
                  triggerLabel = "sum";
                  break;
                case "добавить в прайс":
                  triggerLabel = "addToPrice";
                  break;
                case "акции":
                  triggerLabel = "promotions";
                  break;
                case "набор":
                  triggerLabel = "sets";
                  break;
                default:
                  break;
              }
            }
            if (triggerLabel) {
              Lib.logDebug(
                '[Order:onEdit] прямой вызов override для "' +
                  header +
                  '" (row ' +
                  row +
                  ", col " +
                  col +
                  ", trigger=" +
                  triggerLabel +
                  ")"
              );
              _invokeOnUpdateOverride(
                sheet,
                row,
                col,
                "Order:onEdit:" + triggerLabel
              );
            } else if (row > headerRowIndex) {
              // Доп. случай: изменения в колонках Остаток/Резерв/Товар в пути/Продажи/Остаток 1-3/СГ 1-3/СПИСАНО → пересчёт производных
              const compact = normalizedHeader.replace(/\s+/g, "");
              const shouldRecalcDerived =
                normalizedHeader === "остаток" ||
                normalizedHeader.indexOf("остаток1") === 0 ||
                normalizedHeader.indexOf("остаток 1") === 0 ||
                normalizedHeader.indexOf("остаток2") === 0 ||
                normalizedHeader.indexOf("остаток 2") === 0 ||
                normalizedHeader.indexOf("остаток3") === 0 ||
                normalizedHeader.indexOf("остаток 3") === 0 ||
                normalizedHeader.indexOf("сг 1") === 0 ||
                normalizedHeader.indexOf("сг 2") === 0 ||
                normalizedHeader.indexOf("сг 3") === 0 ||
                normalizedHeader === "списано" ||
                normalizedHeader === "резерв" ||
                compact.indexOf("товарвпути") === 0 ||
                normalizedHeader.indexOf("продажи") === 0 ||
                normalizedHeader.indexOf("среднемесячные продажи") === 0;
              if (shouldRecalcDerived) {
                Lib.logDebug(
                  '[Order:onEdit] пересчёт производных по колонке "' +
                    header +
                    '" (row ' +
                    row +
                    ", col " +
                    col +
                    ")"
                );
                _invokeOnUpdateOverride(sheet, row, col, "Order:onEdit:derived");
              }
            }
          }
          if (isPriceSheet && row > 1) {
            if (
              normalizedHeaderForPrice === "группа линии" ||
              normalizedHeaderForPrice === "линия прайс"
            ) {
              _autoAssignPriceLineId(sheet, row);
            } else if (normalizedHeaderForPrice === "продублировать") {
              const duplicateValue = values[r][c];
              if (_isPriceDuplicateFlagSet(duplicateValue)) {
                _duplicatePriceRow(sheet, row);
              }
            }
          }
          continue;
        }
        // Если есть правила, но мы редактируем строку 3 с заголовком 'Потребность на N месяцев' — принудим override
        if (isOrderForm && row === 3) {
          try {
            if (/^\s*Потребность\s+на\s*\d+\s*месяц/i.test(row3Label)) {
              Lib.logDebug(
                "[Order:onEdit] row3Label matches Потребность pattern and applicable.length=" +
                  applicable.length +
                  ", forcing override"
              );
              _invokeOnUpdateOverride(sheet, row, col, "Order:onEdit:force");
              continue;
            }
          } catch (e) {
            /* ignore */
          }
        }

        const key = sheet.getRange(row, 1).getValue();
        if (!key || !String(key).trim()) continue;

        const newValue = values[r][c];
        const newRich = richTextValues[r][c];
        const srcInfo = `${sheetName}!(${header} R${row})`;

        applicable.forEach((rule) =>
          _syncSingleValue(rule, key, newValue, newRich, srcInfo)
        );

        // ВАЖНО: Завершаем цикл обработки каскадом для текущего листа.
        // Это необходимо, чтобы локальные вычисления (например на листе Сертификация)
        // срабатывали даже если нет правил синхронизации "наружу".
        if (sheetName === Lib.CONFIG.SHEETS.CERTIFICATION) {
           _triggerCascadeUpdates(sheet, row, col);
        }

        // Дополнительно: для листа Заказ всегда пересчитываем зависимые поля
        // даже если есть правила синхронизации (post-sync триггер).
        if (isOrderForm && row > headerRowIndex) {
          try {
            const headerForTrigger = dataHeader || header;
            const normalizedHeader = String(headerForTrigger || "")
              .trim()
              .replace(/\s+/g, " ")
              .toLowerCase();
            const normalizedTopHeader = String(header || "")
              .trim()
              .replace(/\s+/g, " ")
              .toLowerCase();
            let normalizedHeaderKey = normalizedHeader;
            if (
              normalizedHeaderKey === "да" &&
              normalizedTopHeader === "добавить в прайс"
            ) {
              normalizedHeaderKey = normalizedTopHeader;
            }
            let triggerLabel = null;
            switch (normalizedHeaderKey) {
              case "предварительный заказ":
                triggerLabel = "preOrder";
                break;
              case "заказ в упаковках":
                triggerLabel = "packs";
                break;
              case "заказ":
                triggerLabel = "order";
                break;
              case "шт./уп.":
                triggerLabel = "unitsPerPack";
                break;
              case "exw alfaspa текущая, €":
                triggerLabel = "price";
                break;
              case "предварительная сумма заказа":
                triggerLabel = "sum";
                break;
              case "добавить в прайс":
                triggerLabel = "addToPrice";
                break;
              default: {
                const compact = normalizedHeader.replace(/\s+/g, "");
                const derived =
                  normalizedHeader === "остаток" ||
                  normalizedHeader === "резерв" ||
                  compact.indexOf("товарвпути") === 0 ||
                  normalizedHeader.indexOf("продажи") === 0 ||
                  normalizedHeader.indexOf("среднемесячные продажи") === 0;
                if (derived) triggerLabel = "derived";
              }
            }
            if (triggerLabel) {
              _invokeOnUpdateOverride(
                sheet,
                row,
                col,
                "Order:onEdit:postSync:" + triggerLabel
              );
            }
          } catch (_) {
            // soft-fail
          }
        }
        if (isPriceSheet && row > 1) {
          if (
            applicable.length > 0 &&
            (normalizedHeaderForPrice === "группа линии" ||
              normalizedHeaderForPrice === "линия прайс")
          ) {
            _autoAssignPriceLineId(sheet, row);
          } else if (
            normalizedHeaderForPrice === "продублировать" &&
            _isPriceDuplicateFlagSet(values[r][c])
          ) {
            _duplicatePriceRow(sheet, row);
          }
        }
      }
    }
  }

  /**
   * Загрузка/кэш правил из листа «Правила синхро»
   * @private
   */
  function _loadSyncRules(forceReload = false) {
    if (_cachedSyncRules && !forceReload) return _cachedSyncRules;

    const rulesSheetName = Lib.CONFIG.SHEETS.RULES;
    const sheet =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName(rulesSheetName);
    if (!sheet || sheet.getLastRow() <= 1) {
      _cachedSyncRules = [];
      return _cachedSyncRules;
    }

    // Берём ширину по факту, но минимум 10 колонок (как у тебя по структуре)
    const width = Math.max(10, sheet.getLastColumn());
    const data = sheet
      .getRange(2, 1, sheet.getLastRow() - 1, width)
      .getValues();

    const result = [];
    for (let i = 0; i < data.length; i++) {
      const row = data[i];

      // Булевы чекбоксы/«Да»
      const enabled =
        row[1] === true ||
        String(row[1] || "")
          .trim()
          .toLowerCase() === "да";
      const isExternal =
        row[8] === true ||
        String(row[8] || "")
          .trim()
          .toLowerCase() === "да";
      if (!enabled) continue;

      const rule = {
        id: String(row[0] || "").trim(),
        enabled,
        category: String(row[2] || "").trim(),
        hashtags: String(row[3] || "").trim(),
        sourceSheet: String(row[4] || "").trim(),
        sourceHeader: String(row[5] || "").trim(),
        targetSheet: String(row[6] || "").trim(),
        targetHeader: String(row[7] || "").trim(),
        isExternal,
        targetDocId: String(row[9] || "").trim(),
      };

      // Жёсткая валидация, иначе пропускаем строку
      if (
        !rule.sourceSheet ||
        !rule.sourceHeader ||
        !rule.targetSheet ||
        !rule.targetHeader
      )
        continue;
      if (rule.isExternal && !rule.targetDocId) continue;

      result.push(rule);
    }

    _cachedSyncRules = result;
    Lib.logDebug(`_loadSyncRules: активных правил = ${result.length}`);
    return _cachedSyncRules;
  }

  /**
   * Единая точка переноса значения по правилу (устойчивая к битым правилам)
   * @param {Object} rule              Правило синхронизации
   * @param {string} key               ID (значение из колонки A)
   * @param {*} newValue               Новое значение из источника
   * @param {GoogleAppsScript.Spreadsheet.RichTextValue} newRichTextValue  (опционально)
   * @param {string} sourceInfoForLog  Текст-указатель на исходную ячейку для журнала
   */
  function _syncSingleValue(
    rule,
    key,
    newValue,
    newRichTextValue,
    sourceInfoForLog
  ) {
    try {
      // ---- ГАРДЫ НА ВХОДЕ ----------------------------------------------------
      if (!rule || typeof rule !== "object") {
        Lib.logSyncError(
          String(key || ""),
          rule,
          sourceInfoForLog,
          "Некорректное правило (undefined/не object)"
        );
        return;
      }
      if (!key || String(key).trim() === "") {
        Lib.logSyncError(
          String(key || ""),
          rule,
          sourceInfoForLog,
          "Пустой ключ (ID)"
        );
        return;
      }
      if (!rule.targetSheet || !rule.targetHeader) {
        Lib.logSyncError(
          String(key || ""),
          rule,
          sourceInfoForLog,
          "Правило без targetSheet/targetHeader"
        );
        return;
      }
      if (rule.isExternal && !rule.targetDocId) {
        Lib.logSyncError(
          String(key || ""),
          rule,
          sourceInfoForLog,
          "Внешняя цель без targetDocId"
        );
        return;
      }

      // ---- Определяем целевую таблицу/лист ----------------------------------
      const ss = rule.isExternal
        ? SpreadsheetApp.openById(rule.targetDocId)
        : SpreadsheetApp.getActiveSpreadsheet();

      const targetSheet = ss.getSheetByName(rule.targetSheet);
      if (!targetSheet) throw new Error(`Лист "${rule.targetSheet}" не найден`);

      // ---- Находим целевую ячейку по ИМЕНИ колонки и по ID -------------------
      const targetCol = Lib.findColumnIndexByHeader_(
        targetSheet,
        rule.targetHeader
      );
      if (targetCol === -1)
        throw new Error(`Столбец "${rule.targetHeader}" не найден`);

      const targetRow = Lib.findRowByKey_(targetSheet, key); // быстрый поиск по кэшу
      if (targetRow === -1) {
        // Нет строки с таким ID — молча выходим (ничего создавать здесь не нужно)
        return;
      }

      const cell = targetSheet.getRange(targetRow, targetCol);
      const oldValue = cell.getValue();
      if (Lib.areValuesEqual_(oldValue, newValue)) {
        // Значение не изменилось — ничего не пишем и не логируем
        return;
      }

      // ---- Запись значения (с поддержкой RichText) --------------------------
      if (newRichTextValue && typeof cell.setRichTextValue === "function") {
        cell.setRichTextValue(newRichTextValue);
      } else {
        cell.setValue(newValue);
      }
      if (rule.isExternal) SpreadsheetApp.flush(); // чтобы каскад/лог видел актуализацию

      // ---- Каскады (если предусмотрены) -------------------------------------
      _triggerCascadeUpdates(targetSheet, targetRow, targetCol);

      // ---- Журнал ------------------------------------------------------------
      const targetInfo = `${targetSheet.getName()}!(${
        rule.targetHeader
      } R${targetRow})`;
      Lib.logSynchronization(
        key,
        sourceInfoForLog,
        targetInfo,
        oldValue,
        newValue,
        rule.category,
        rule.hashtags,
        "SYNC"
      );
    } catch (err) {
      // Любая ошибка здесь НЕ должна падать дальше — логируем аккуратно
      Lib.logSyncError(
        String(key || ""),
        rule,
        sourceInfoForLog,
        err && err.message ? err.message : String(err)
      );
      Lib.logError(
        `Ошибка в _syncSingleValue (rule: ${rule && rule.id ? rule.id : "?"})`,
        err
      );
    }
  }

  // =======================================================================================
  // МОДУЛЬ 3: СТРУКТУРНЫЕ ОПЕРАЦИИ (создание «пустышек», пакетное удаление)
  // ---------------------------------------------------------------------------------------
  // =======================================================================================

  /**
   * Создаёт связанные строки на базовых листах при появлении нового ID (колонка A)
   * @private
   */
  function _ensureRowExistsOnBaseSheets(key) {
    const list = Lib.CONFIG.SETTINGS.BASE_SHEETS_FOR_CREATION || [];
    if (!list || list.length === 0) return;

    Lib.logInfo(`Проверяем автосоздание строк для "${key}"...`);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    list.forEach((name) => {
      const sh = ss.getSheetByName(name);
      if (!sh) {
        Lib.logWarn(`BASE_SHEETS_FOR_CREATION: лист "${name}" не найден`);
        return;
      }

      const rowNum = Lib.findRowByKey_(sh, key, true);
      if (rowNum === -1) {
        sh.appendRow([key]);
        Lib.deleteKeyCacheForSheet(name);
        Lib.logInfo(`Создана строка на "${name}" для ID ${key}`);
      }
    });
  }

  /**
   * Публичная обёртка для создания строк на базовых листах
   * Используется при добавлении новых артикулов через обработку прайс-листов
   */
  Lib.ensureRowExistsOnBaseSheets = function (key) {
    _ensureRowExistsOnBaseSheets(key);
  };

  /**
   * Пакетное удаление записей по массиву ID на всех листах проекта
   */
  Lib.processBatchDeletion_ = function (keysToDelete) {
    Lib.logInfo(`Пакетное удаление: ключей ${keysToDelete.length}`);
    let total = 0;
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const all = Object.values(Lib.CONFIG.SHEETS);

      all.forEach((name) => {
        const sh = ss.getSheetByName(name);
        if (!sh) return;

        const keyMap = (function (sheet) {
          return (function () {
            return _getKeyMapForSheet(sheet, true);
          })();
        })(sh); // свежая карта
        const rows = [];
        keysToDelete.forEach((k) => {
          const r = keyMap.get(String(k));
          if (r) rows.push(r);
        });
        if (rows.length > 0) {
          rows.sort((a, b) => b - a).forEach((r) => sh.deleteRow(r));
          Lib.deleteKeyCacheForSheet(name);
          total += rows.length;
          Lib.logInfo(`Удалено ${rows.length} строк на "${name}"`);
        }
      });

      return { message: `Готово: удалено ${total} строк(и) во всех листах.` };
    } catch (err) {
      Lib.logError("processBatchDeletion_: ошибка", err);
      return { message: `Ошибка: ${err.message}` };
    }
  };

  // =======================================================================================
  // МОДУЛЬ 4: КАСКАДЫ (бизнес-логика после записи)
  // ---------------------------------------------------------------------------------------
  // =======================================================================================

  /**
   * Диспетчер каскадных обновлений
   * @private
   */
  function _triggerCascadeUpdates(sheet, row, col) {
    const sheetName = sheet.getName();
    Lib.logDebug(`[CASCADE] Диспетчер: ${sheetName}, R${row}, C${col}`);

    if (sheetName === Lib.CONFIG.SHEETS.CERTIFICATION) {
      const updatedHeader = String(
        sheet.getRange(1, col).getValue() || ""
      ).trim();
      // Мы убрали строгую проверку trigger.includes(updatedHeader) отсюда,
      // так как теперь _runCertificationCascade имеет внутри "умную" проверку
      // (нечувствительную к регистру и е/ё), чтобы не дублировать логику.
      _runCertificationCascade(sheet, row, updatedHeader);
    }

    _invokeOnUpdateOverride(sheet, row, col, "CASCADE");
  }

  /**
   * Унифицированный запуск override-обработчика Lib.CONFIG.OVERRIDES.onUpdate
   * @private
   */
  function _invokeOnUpdateOverride(sheet, row, col, debugSource) {
    const overrideFunc = Lib.CONFIG.OVERRIDES.onUpdate;
    if (!overrideFunc) {
      Lib.logDebug(
        `[OVERRIDE] обработчик onUpdate не задан (source=${debugSource})`
      );
      return;
    }
    if (typeof Lib[overrideFunc] !== "function") {
      Lib.logWarn(
        `[OVERRIDE] функция "${overrideFunc}" не найдена (source=${debugSource})`
      );
      return;
    }
    const sheetName = sheet.getName();
    Lib.logDebug(
      '[OVERRIDE] запуск "' +
        overrideFunc +
        '" (source=' +
        debugSource +
        ") для " +
        sheetName +
        "!R" +
        row +
        "C" +
        col
    );
    try {
      // Передаём debugSource в override как 4-й необязательный аргумент,
      // чтобы обработчик мог точнее определить тип триггера.
      Lib[overrideFunc](sheet, row, col, debugSource);
    } catch (e) {
      Lib.logError(
        `[OVERRIDE] Ошибка override "${overrideFunc}" (source=${debugSource})`,
        e
      );
    }
  }

  function _autoAssignPriceLineId(sheet, row) {
    try {
      const priceSheetName = Lib.CONFIG.SHEETS.PRICE;
      if (!sheet || sheet.getName() !== priceSheetName) return;
      if (!row || row <= 1) return;

      const idlCol = Lib.findColumnIndexByHeader_(sheet, "ID-L");
      const groupCol = Lib.findColumnIndexByHeader_(sheet, "Группа линии");
      const lineCol = Lib.findColumnIndexByHeader_(sheet, "Линия Прайс");

      if (idlCol === -1 || groupCol === -1 || lineCol === -1) return;

      const groupValue = sheet.getRange(row, groupCol).getValue();
      const lineValue = sheet.getRange(row, lineCol).getValue();
      const resolver =
        typeof Lib.resolvePriceLineId_ === "function"
          ? Lib.resolvePriceLineId_
          : null;
      if (!resolver) return;

      const resolved = resolver(groupValue, lineValue);
      if (resolved === null || resolved === undefined || resolved === "") {
        return;
      }

      const targetCell = sheet.getRange(row, idlCol);
      const current = targetCell.getValue();
      if (String(current || "").trim() === String(resolved)) {
        return;
      }

      targetCell.setValue(resolved);
      Lib.logDebug(
        `[PriceAutoID] R${row}: ID-L=${resolved} (${String(groupValue || "").trim()} / ${String(lineValue || "").trim()})`
      );
    } catch (error) {
      Lib.logError("[PriceAutoID] автоназначение ID-L: ошибка", error);
    }
  }

  /**
   * Публичная функция для автоматического присвоения ID-L для строки на листе Прайс
   * @param {Sheet} sheet - лист Прайс
   * @param {number} row - номер строки (1-based)
   */
  Lib.autoAssignPriceLineIdForRow = function(sheet, row) {
    _autoAssignPriceLineId(sheet, row);
  };

  function _duplicatePriceRow(sheet, row) {
    try {
      const priceSheetName = Lib.CONFIG.SHEETS.PRICE;
      if (!sheet || sheet.getName() !== priceSheetName) return;
      if (!row || row <= 1) return;

      const totalCols = sheet.getLastColumn();
      const sourceRow = sheet.getRange(row, 1, 1, totalCols).getValues()[0];

      const headersToCopy = [
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
      ];

      const newRowValues = new Array(totalCols).fill("");
      headersToCopy.forEach(function (header) {
        const colIndex = Lib.findColumnIndexByHeader_(sheet, header);
        if (colIndex !== -1) {
          newRowValues[colIndex - 1] = sourceRow[colIndex - 1];
        }
      });

      sheet.insertRowAfter(row);
      sheet.getRange(row + 1, 1, 1, totalCols).setValues([newRowValues]);

      Lib.logInfo(
        `[PriceDuplicate] Продублирована строка ${row} → ${row + 1} (ID=${String(
          sourceRow[0] || ""
        ).trim()})`
      );
    } catch (error) {
      Lib.logError("[PriceDuplicate] ошибка дублирования строки", error);
    }
  }

  function _isPriceDuplicateFlagSet(value) {
    if (value === true) return true;
    if (value === false || value === null || value === undefined) return false;
    const normalized = String(value).trim().toLowerCase();
    return (
      normalized === "true" ||
      normalized === "yes" ||
      normalized === "да" ||
      normalized === "1"
    );
  }

  /**
   * Каскад для листа «Сертификация» (header-driven, без привязки к порядку колонок)
   * - Индексация только по ФАКТИЧЕСКИМ заголовкам листа.
   * - Триггерится только от изменений в SOURCE_HEADERS.
   * - УЛУЧШЕНИЕ: Нечувствительность к регистру и замене е/ё в заголовках.
   */
  function _runCertificationCascade(sheet, row, updatedHeader) {
    if (!updatedHeader) return;

    // --- ХЕЛПЕРЫ НОРМАЛИЗАЦИИ ---
    // Приводим к нижнему регистру и заменяем ё -> е для надежного сравнения
    const normalize = (str) => String(str || "").trim().toLowerCase().replace(/ё/g, "е");

    const rules = Lib.CONFIG.CASCADE_RULES?.CERTIFICATION || {};
    const rawSourceHeaders = rules.SOURCE_HEADERS || [
      "Наименования рус по ДС",
      "Наименования англ по ДС",
      "Объём",
      "Код ТН ВЭД",
    ];

    // Нормализуем список триггеров
    const sourceHeadersMap = new Set(rawSourceHeaders.map(normalize));
    const normalizedUpdatedHeader = normalize(updatedHeader);

    // Триггер не из списка — выходим
    if (!sourceHeadersMap.has(normalizedUpdatedHeader)) {
      Lib.logDebug(`[CASCADE] Skip: Заголовок "${updatedHeader}" (norm: "${normalizedUpdatedHeader}") не является триггером.`);
      return;
    }

    Lib.logInfo(`[CASCADE] Start: Запуск каскада для R${row}, триггер "${updatedHeader}"`);

    // Имена ключевых полей (словарь "ключ -> нормализованное имя")
    // Мы ищем колонки, чьи заголовки при нормализации совпадают с этими значениями
    const TARGET_FIELDS = {
      RUS: normalize("Наименования рус по ДС"),
      ENG: normalize("Наименования англ по ДС"),
      VOL: normalize("Объём"),
      TNVED: normalize("Код ТН ВЭД"),
      VOL_EN: normalize("Объём англ."),
      DS_NAME: normalize("Наименование ДС"),
      INV_RU: normalize("Наименование для инвойса"),
      INV_EN: normalize("Наименование для инвойса Англ"),
    };

    // Карта заголовков по факту: [ { norm: '...', col: 1 }, ... ]
    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const headerMap = {};

    headers.forEach((h, idx) => {
      const n = normalize(h);
      if (n) headerMap[n] = idx + 1; // 1-based index
    });

    const getCol = (normalizedFieldName) => headerMap[normalizedFieldName] || -1;
    const getValueByNormName = (normName) => {
      const c = getCol(normName);
      return c > 0 ? sheet.getRange(row, c).getValue() : null;
    };

    // Проверяем наличие источников
    const missing = [];
    if (getCol(TARGET_FIELDS.RUS) === -1) missing.push("Наименования рус по ДС");
    if (getCol(TARGET_FIELDS.ENG) === -1) missing.push("Наименования англ по ДС");
    if (getCol(TARGET_FIELDS.VOL) === -1) missing.push("Объём");
    // Код ТН ВЭД не обязательно (может не быть колонки), но желательно

    if (missing.length) {
      Lib.logWarn(`[CASCADE] Отмена: Не найдены обязательные колонки: ${missing.join(", ")}`);
      return;
    }

    // Читаем источники
    const rusName = String(getValueByNormName(TARGET_FIELDS.RUS) || "").trim();
    const engName = String(getValueByNormName(TARGET_FIELDS.ENG) || "").trim();
    const volume = String(getValueByNormName(TARGET_FIELDS.VOL) || "");
    const tnved = String(getValueByNormName(TARGET_FIELDS.TNVED) || "").trim();

    // Если в строке нет исходников (кроме ID) — каскад не нужен
    if (!rusName && !engName && !volume && !tnved) {
      Lib.logDebug(`[CASCADE] R${row}: пустые источники — выходим`);
      return;
    }

    // Текущее «Объём англ.»
    let currentVolEn = "";
    const colVolEn = getCol(TARGET_FIELDS.VOL_EN);
    if (colVolEn > 0) {
      currentVolEn = String(sheet.getRange(row, colVolEn).getValue() || "");
    }

    // Наименование ДС (правило с запятой в рус. имени)
    let newDsName = "";
    if (rusName && engName) {
      const rusEndsWithComma = /,\s*$/.test(rusName);
      newDsName = rusEndsWithComma
        ? `${rusName} ${engName}`
        : `${rusName} / ${engName}`;
    } else {
      newDsName = rusName || engName;
    }

    // Пересчёт объёма (англ.) — только если триггер «Объём»
    let newVolEn = currentVolEn;
    if (normalizedUpdatedHeader === TARGET_FIELDS.VOL) {
      newVolEn = String(volume || "");
      const repl = rules.VOLUME_EN_REPLACEMENTS || [];
      repl.forEach((r) => {
        try {
          newVolEn = newVolEn.replace(new RegExp(r.from, "gi"), r.to);
        } catch (_) {}
      });
      newVolEn = newVolEn.replace(/\s+/g, " ").trim();
    }

    // Инвойсные названия
    let newInvRu = newDsName
      ? `${newDsName} ${volume}`.trim()
      : String(volume || "").trim();
    if (tnved) newInvRu += `\nКод ТН ВЭД: ${tnved}`;

    let newInvEn = engName
      ? `${engName} ${newVolEn || currentVolEn}`.trim()
      : String(newVolEn || currentVolEn || "").trim();
    if (tnved) newInvEn += `\nCode: ${tnved}`;

    // Запись только если изменилось
    const writeSafe = (normField, val) => {
      const c = getCol(normField);
      if (c > 0) {
        const prev = sheet.getRange(row, c).getValue();
        if (!Lib.areValuesEqual_(prev, val)) {
          sheet.getRange(row, c).setValue(val);
          return true; // записали
        }
      }
      return false; // не было изменений или колонки нет
    };

    let changesCount = 0;
    if (rusName || engName) {
      if (writeSafe(TARGET_FIELDS.DS_NAME, newDsName)) changesCount++;
    }
    if (normalizedUpdatedHeader === TARGET_FIELDS.VOL) {
      if (writeSafe(TARGET_FIELDS.VOL_EN, newVolEn)) changesCount++;
    }
    if (writeSafe(TARGET_FIELDS.INV_RU, newInvRu)) changesCount++;
    if (writeSafe(TARGET_FIELDS.INV_EN, newInvEn)) changesCount++;

    Lib.logInfo(`[CASCADE] Сертификация R${row}: обработано, записей изменений: ${changesCount}`);
  }

  // =======================================================================================
  // МОДУЛЬ 5: ОБРАБОТЧИКИ UI (вызовы из меню)
  // ---------------------------------------------------------------------------------------
  // =======================================================================================

  /**
   * Добавить новый артикул на «Главная» и запустить создание/синхронизацию
   */
  Lib.addArticleManually = function () {
    let ui = null;
    try { ui = SpreadsheetApp.getUi(); } catch (e) { console.warn("UI not available"); }
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sh = ss.getSheetByName(Lib.CONFIG.SHEETS.PRIMARY);
      if (!sh) throw new Error(`Лист "${Lib.CONFIG.SHEETS.PRIMARY}" не найден`);

      // найти максимальный номер по префиксу бренда
      const prefix = Lib.CONFIG.SETTINGS.BRAND_PREFIX;
      const lastRow = sh.getLastRow();
      let maxNum = 0;
      if (lastRow > 1) {
        const keys = sh
          .getRange(2, 1, lastRow - 1, 1)
          .getValues()
          .flat()
          .map((v) => String(v || "").trim())
          .filter(Boolean);
        keys.forEach((id) => {
          if (id.startsWith(prefix)) {
            const n = parseInt(id.substring(prefix.length), 10);
            if (!isNaN(n) && n > maxNum) maxNum = n;
          }
        });
      }
      const next = String(maxNum + 1).padStart(3, "0");
      const newId = `${prefix}${next}`;

      sh.appendRow([newId]);
      Lib.deleteKeyCacheForSheet(Lib.CONFIG.SHEETS.PRIMARY);

      // имитируем onEdit для автосоздания и синхронизации
      const e = {
        range: sh.getRange(sh.getLastRow(), 1),
        value: newId,
        oldValue: "",
      };
      // Проверка на наличие onEdit_internal_ (если вызывается из API)
      if (Lib.onEdit_internal_) {
          Lib.onEdit_internal_(e);
      }

      if (ui) ui.alert(`Новый артикул ${newId} создан.`);
      console.log(`Новый артикул ${newId} создан.`);
    } catch (err) {
      console.error("addArticleManually error: " + err.message);
      if (ui) ui.alert(`Ошибка: ${err.message}`);
      Lib.logError("addArticleManually: ошибка", err);
      // Пробрасываем ошибку дальше, чтобы API увидел её
      throw err;
    }
  };

  /**
   * Удалить выделенные строки и синхронизировать удаление в связанных листах
   */
  Lib.deleteSelectedRowsWithSync = function () {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sel = SpreadsheetApp.getActiveRangeList();
    if (!sel) {
      ui.alert("Выберите строки для удаления");
      return;
    }

    const sheet = ss.getActiveSheet();
    const sheetName = sheet.getName();
    const sheetsConfig = (Lib.CONFIG && Lib.CONFIG.SHEETS) || {};
    if (!sheetsConfig || Object.keys(sheetsConfig).length === 0) {
      ui.alert("Конфигурация проекта не загружена. Обновите библиотеку.");
      return;
    }
    const service = [
      sheetsConfig.RULES,
      sheetsConfig.LOG,
      sheetsConfig.EXTERNAL_DOCS,
    ].filter(Boolean);
    if (service.includes(sheetName)) {
      ui.alert("Нельзя удалять строки на служебных листах.");
      return;
    }

    const rows = new Set();
    sel.getRanges().forEach((r) => {
      for (let i = r.getRow(); i <= r.getLastRow(); i++) if (i > 1) rows.add(i);
    });
    const rowsSorted = Array.from(rows).sort((a, b) => b - a);
    if (rowsSorted.length === 0) {
      ui.alert("Нет строк данных для удаления");
      return;
    }

    const confirm = ui.alert(
      "Подтверждение",
      `Удалить ${rowsSorted.length} строк(и) на листе "${sheetName}" и синхронно в связанных?`,
      ui.ButtonSet.YES_NO
    );
    if (confirm !== ui.Button.YES) return;

    // собираем ID и удаляем локально
    const ids = [];
    rowsSorted.forEach((r) => {
      const id = String(sheet.getRange(r, 1).getValue() || "").trim();
      if (id) ids.push(id);
      sheet.deleteRow(r);
    });
    Lib.deleteKeyCacheForSheet(sheetName);

    if (ids.length > 0) {
      const res = Lib.processBatchDeletion_(ids);
      ui.alert(res.message);
    } else {
      ui.alert(`Удалено ${rowsSorted.length} строк(и).`);
    }
  };

  /**
   * Ручная синхронизация данных для выделенных строк на активном листе
   */
  Lib.syncSelectedRow = function () {
    const ui = SpreadsheetApp.getUi();
    const range = SpreadsheetApp.getActiveRange();
    if (!range) {
      ui.alert("Выделите строки");
      return;
    }

    const sheet = range.getSheet();
    if (range.getRow() <= 1 && range.getLastRow() <= 1) {
      ui.alert("Выберите строки с данными");
      return;
    }

    const resp = ui.alert(
      "Синхронизация",
      `Синхронизировать ${range.getNumRows()} строк(и) на "${sheet.getName()}"?`,
      ui.ButtonSet.YES_NO
    );
    if (resp !== ui.Button.YES) return;

    try {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `Синхронизация...`,
        "Выполнение",
        2
      );
      const rules = _loadSyncRules(true);
      if (rules.length === 0) {
        ui.alert("Нет активных правил");
        return;
      }

      const lastCol = sheet.getLastColumn();
      const headers = sheet
        .getRange(1, 1, 1, lastCol)
        .getValues()[0]
        .map((h) => String(h || "").trim());

      // соберём уникальные номера строк (без заголовка)
      const rows = [];
      for (let r = range.getRow(); r <= range.getLastRow(); r++)
        if (r > 1) rows.push(r);

      let counter = 0;
      rows.forEach((row) => {
        const key = String(sheet.getRange(row, 1).getValue() || "").trim();
        if (!key) return;

        const rowVals = sheet.getRange(row, 1, 1, lastCol).getValues()[0];
        const rowRich = sheet
          .getRange(row, 1, 1, lastCol)
          .getRichTextValues()[0];

        headers.forEach((h, idx) => {
          if (!h) return;
          const applicable = rules.filter(
            (r) =>
              r &&
              typeof r === "object" &&
              r.sourceSheet === sheet.getName() &&
              r.sourceHeader === h &&
              r.targetSheet &&
              r.targetHeader
          );
          if (applicable.length === 0) return;

          const value = rowVals[idx];
          const rich = rowRich[idx];
          const src = `${sheet.getName()}!(${h} R${row})`;
          applicable.forEach((rule) =>
            _syncSingleValue(rule, key, value, rich, src)
          );
          counter += applicable.length;
        });
      });

      SpreadsheetApp.getActiveSpreadsheet().toast(
        `Готово. Обработано ${counter} полей.`,
        "Готово",
        5
      );
      ui.alert(`Синхронизация завершена. Полей: ${counter}`);
    } catch (err) {
      Lib.logError("syncSelectedRow: ошибка", err);
      ui.alert(`Ошибка: ${err.message}`);
      SpreadsheetApp.getActiveSpreadsheet().toast('Ошибка синхронизации', 'Ошибка', 3);
    } finally {
      // Гарантированно закрываем любой висящий toast
      SpreadsheetApp.getActiveSpreadsheet().toast('', '', 1);
    }
  };

  /**
   * Автоматическая синхронизация данных для указанных строк БЕЗ диалогов
   * @param {Array<number>} rowNumbers - массив номеров строк для синхронизации
   * @param {string} sheetName - имя листа (опционально, по умолчанию "Главная")
   */
  Lib.syncMultipleRows = function (rowNumbers, sheetName) {
    if (!rowNumbers || rowNumbers.length === 0) {
      Lib.logInfo("[syncMultipleRows] Нет строк для синхронизации");
      return;
    }

    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const targetSheetName = sheetName || (global.CONFIG && global.CONFIG.SHEETS && global.CONFIG.SHEETS.PRIMARY) || "Главная";
      const sheet = ss.getSheetByName(targetSheetName);

      if (!sheet) {
        Lib.logError("[syncMultipleRows] Лист не найден: " + targetSheetName);
        return;
      }

      const rules = _loadSyncRules(true);
      if (rules.length === 0) {
        Lib.logInfo("[syncMultipleRows] Нет активных правил синхронизации");
        return;
      }

      const lastCol = sheet.getLastColumn();
      const headers = sheet
        .getRange(1, 1, 1, lastCol)
        .getValues()[0]
        .map((h) => String(h || "").trim());

      let counter = 0;
      rowNumbers.forEach((row) => {
        if (row <= 1) return; // Пропускаем заголовок

        const key = String(sheet.getRange(row, 1).getValue() || "").trim();
        if (!key) return;

        const rowVals = sheet.getRange(row, 1, 1, lastCol).getValues()[0];
        const rowRich = sheet
          .getRange(row, 1, 1, lastCol)
          .getRichTextValues()[0];

        headers.forEach((h, idx) => {
          if (!h) return;
          const applicable = rules.filter(
            (r) =>
              r &&
              typeof r === "object" &&
              r.sourceSheet === sheet.getName() &&
              r.sourceHeader === h &&
              r.targetSheet &&
              r.targetHeader
          );
          if (applicable.length === 0) return;

          const value = rowVals[idx];
          const rich = rowRich[idx];
          const src = `${sheet.getName()}!(${h} R${row})`;
          applicable.forEach((rule) =>
            _syncSingleValue(rule, key, value, rich, src)
          );
          counter += applicable.length;
        });
      });

      Lib.logInfo(
        `[syncMultipleRows] Синхронизировано ${rowNumbers.length} строк(и), обработано ${counter} полей`
      );

      // Обновляем формулы на листе "Расчет цены" после добавления новых строк
      // Это необходимо для подтягивания данных из листа "Динамика цены" по новым строкам
      if (typeof Lib.updatePriceCalculationFormulas === 'function') {
        Lib.logInfo('[syncMultipleRows] Обновление формул на листе "Расчет цены"');
        Lib.updatePriceCalculationFormulas();
      }

      // Обновляем вертикальные границы на листе "Расчет цены" после добавления новых строк
      if (typeof Lib.updatePriceCalculationBorders === 'function') {
        Lib.updatePriceCalculationBorders();
      }

      // Обновляем вертикальные границы на листе "Динамика цены" после добавления новых строк
      if (typeof Lib.updatePriceDynamicsBorders === 'function') {
        Lib.updatePriceDynamicsBorders();
      }

      // Обновляем вертикальные границы на листе "Заказ" после добавления новых строк
      if (typeof Lib.updateOrderFormBorders === 'function') {
        Lib.updateOrderFormBorders();
      }
    } catch (err) {
      Lib.logError("syncMultipleRows: ошибка", err);
    }
  };

  /**
   * Ручной запуск каскадов для всего листа «Сертификация»
   */
  Lib.runManualCascadeOnCertification = function () {
    const ui = SpreadsheetApp.getUi();
    const name = Lib.CONFIG.SHEETS.CERTIFICATION;
    const resp = ui.alert(
      "Пересчёт каскадов",
      `Пересчитать все строки на листе "${name}"?`,
      ui.ButtonSet.YES_NO
    );
    if (resp !== ui.Button.YES) return;

    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
    if (!sh) {
      ui.alert(`Лист "${name}" не найден`);
      return;
    }

    const last = sh.getLastRow();
    if (last <= 1) {
      ui.alert("Нет данных для пересчёта");
      return;
    }

    try {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        "Пересчёт каскадов...",
        "Выполнение",
        2
      );
      const fakeHeader = "Объём"; // триггерим полный пересчёт
      for (let r = 2; r <= last; r++)
        _runCertificationCascade(sh, r, fakeHeader);
      SpreadsheetApp.getActiveSpreadsheet().toast("Готово!", "OK", 5);
      ui.alert("Каскады пересчитаны.");
    } catch (err) {
      Lib.logError("runManualCascadeOnCertification: ошибка", err);
      ui.alert(`Ошибка: ${err.message}`);
      SpreadsheetApp.getActiveSpreadsheet().toast('Ошибка пересчёта', 'Ошибка', 3);
    } finally {
      // Гарантированно закрываем любой висящий toast
      SpreadsheetApp.getActiveSpreadsheet().toast('', '', 1);
    }
  };

  // =======================================================================================
  // МОДУЛЬ 5.5: УПРАВЛЕНИЕ ID-P (очистка и заполнение)
  // ---------------------------------------------------------------------------------------
  // =======================================================================================

  /**
   * ПУБЛИЧНАЯ: Заполнить столбец ID-P на всех листах по данным из листа Главная
   * Используется после обработки Б/З поставщик для синхронизации ID-P на все листы
   * @param {Array<string>} targetSheetNames - Массив имен листов для заполнения ID-P.
   *   Если не указан, заполняется ID-P на листах: Заказ, Динамика цены, Расчет цены, ABC-Анализ
   */
  Lib.fillIdpOnSheetsByIdFromPrimary = function(targetSheetNames) {
    try {
      Lib.logInfo('[fillIdpOnSheetsByIdFromPrimary] === НАЧАЛО ЗАПОЛНЕНИЯ ID-P НА ВСЕХ ЛИСТАХ ===');

      const ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) {
        Lib.logError('[fillIdpOnSheetsByIdFromPrimary] Не удалось получить активную таблицу');
        return;
      }

      // Получаем лист Главная
      const primarySheetName = Lib.CONFIG.SHEETS.PRIMARY;
      Lib.logInfo('[fillIdpOnSheetsByIdFromPrimary] Ищем лист Главная: "' + primarySheetName + '"');

      const primarySheet = ss.getSheetByName(primarySheetName);
      if (!primarySheet) {
        Lib.logError('[fillIdpOnSheetsByIdFromPrimary] Лист "' + primarySheetName + '" не найден');
        return;
      }

      const primaryLastRow = primarySheet.getLastRow();
      if (primaryLastRow <= 1) {
        Lib.logDebug('[fillIdpOnSheetsByIdFromPrimary] Лист Главная пуст');
        return;
      }

      const primaryLastColumn = primarySheet.getLastColumn();
      if (primaryLastColumn <= 0) {
        Lib.logDebug('[fillIdpOnSheetsByIdFromPrimary] Лист Главная не имеет столбцов');
        return;
      }

      // Читаем заголовки и данные с листа Главная
      const primaryHeaders = primarySheet
        .getRange(1, 1, 1, primaryLastColumn)
        .getValues()[0]
        .map(function(value) {
          return String(value || '').trim();
        });

      // Логируем все заголовки для отладки
      Lib.logInfo('[fillIdpOnSheetsByIdFromPrimary] Заголовки на листе Главная: ' + primaryHeaders.join(' | '));

      const primaryIdIndex = primaryHeaders.indexOf('ID');
      const primaryIdpIndex = primaryHeaders.indexOf('ID-P');

      Lib.logInfo('[fillIdpOnSheetsByIdFromPrimary] Индекс столбца ID: ' + primaryIdIndex);
      Lib.logInfo('[fillIdpOnSheetsByIdFromPrimary] Индекс столбца ID-P: ' + primaryIdpIndex);

      if (primaryIdIndex === -1 || primaryIdpIndex === -1) {
        Lib.logError('[fillIdpOnSheetsByIdFromPrimary] На листе Главная не найдены столбцы ID или ID-P');
        Lib.logError('[fillIdpOnSheetsByIdFromPrimary] Доступные заголовки: ' + primaryHeaders.join(', '));
        return;
      }

      // Читаем все данные с листа Главная (ID и ID-P)
      const primaryData = primarySheet.getRange(2, 1, primaryLastRow - 1, primaryLastColumn).getValues();

      // Создаем карту ID -> ID-P
      const idToIdpMap = {};
      for (let i = 0; i < primaryData.length; i++) {
        const row = primaryData[i];
        const id = String(row[primaryIdIndex] || '').trim();
        const idp = row[primaryIdpIndex];

        if (id && idp !== '' && idp !== null && idp !== undefined) {
          idToIdpMap[id] = idp;
        }
      }

      const mapSize = Object.keys(idToIdpMap).length;
      Lib.logInfo('[fillIdpOnSheetsByIdFromPrimary] Создана карта ID -> ID-P, записей: ' + mapSize);

      if (mapSize === 0) {
        Lib.logWarn('[fillIdpOnSheetsByIdFromPrimary] Нет данных для заполнения ID-P');
        return;
      }

      // Список листов по умолчанию (без Главной, т.к. ID-P там уже заполнен)
      const defaultSheets = [
        Lib.CONFIG.SHEETS.ORDER_FORM,        // Заказ
        Lib.CONFIG.SHEETS.PRICE_DYNAMICS,    // Динамика цены
        Lib.CONFIG.SHEETS.PRICE_CALCULATION, // Расчет цены
        Lib.CONFIG.SHEETS.ABC_ANALYSIS,      // ABC-Анализ
      ];

      const targetSheets = targetSheetNames || defaultSheets;
      let totalUpdated = 0;

      // Заполняем ID-P на каждом целевом листе
      targetSheets.forEach(function(sheetName) {
        if (!sheetName) return;

        const sheet = ss.getSheetByName(sheetName);
        if (!sheet) {
          Lib.logWarn('[fillIdpOnSheetsByIdFromPrimary] Лист "' + sheetName + '" не найден');
          return;
        }

        const lastRow = sheet.getLastRow();
        if (lastRow <= 1) {
          Lib.logDebug('[fillIdpOnSheetsByIdFromPrimary] Лист "' + sheetName + '" пуст, пропускаем');
          return;
        }

        const lastColumn = sheet.getLastColumn();
        if (lastColumn <= 0) {
          Lib.logDebug('[fillIdpOnSheetsByIdFromPrimary] Лист "' + sheetName + '" не имеет столбцов');
          return;
        }

        // Ищем столбцы ID и ID-P
        const headers = sheet
          .getRange(1, 1, 1, lastColumn)
          .getValues()[0]
          .map(function(value) {
            return String(value || '').trim();
          });

        const idIndex = headers.indexOf('ID');
        const idpIndex = headers.indexOf('ID-P');

        if (idIndex === -1) {
          Lib.logWarn('[fillIdpOnSheetsByIdFromPrimary] Столбец ID не найден на листе "' + sheetName + '"');
          return;
        }

        if (idpIndex === -1) {
          Lib.logDebug('[fillIdpOnSheetsByIdFromPrimary] Столбец ID-P не найден на листе "' + sheetName + '", пропускаем');
          return;
        }

        // Читаем данные (только столбец ID)
        const data = sheet.getRange(2, idIndex + 1, lastRow - 1, 1).getValues();

        // Подготавливаем массив значений для записи
        const idpValues = [];
        let updatedCount = 0;

        for (let i = 0; i < data.length; i++) {
          const id = String(data[i][0] || '').trim();

          if (id && idToIdpMap.hasOwnProperty(id)) {
            idpValues.push([idToIdpMap[id]]);
            updatedCount++;
          } else {
            idpValues.push(['']); // Оставляем пустым, если ID не найден
          }
        }

        // Записываем все значения ID-P одним вызовом
        if (idpValues.length > 0) {
          sheet.getRange(2, idpIndex + 1, idpValues.length, 1).setValues(idpValues);
          totalUpdated += updatedCount;
          Lib.logInfo('[fillIdpOnSheetsByIdFromPrimary] Заполнено ID-P на листе "' + sheetName + '": ' + updatedCount + ' строк');
        }
      });

      if (totalUpdated > 0) {
        Lib.logInfo('[fillIdpOnSheetsByIdFromPrimary] Всего заполнено строк: ' + totalUpdated);
      } else {
        Lib.logInfo('[fillIdpOnSheetsByIdFromPrimary] Не найдено строк для заполнения');
      }
    } catch (err) {
      Lib.logError('[fillIdpOnSheetsByIdFromPrimary] Ошибка', err);
    }
  };

  /**
   * ПУБЛИЧНАЯ: Копировать цену из обработанных данных на листы Динамика цены и Расчет цены
   * Используется в проекте MT для заполнения столбца "ЦЕНА EXW из Б/З"
   * @param {Object} processed - Объект с обработанными данными (содержит headers и rows)
   * @param {Array<string>} targetSheetNames - Массив имен листов для копирования цены.
   *   Если не указан, копируется на листы: Динамика цены, Расчет цены
   * @param {number} priceMultiplier - Множитель для цены (по умолчанию 1, для пробников MT = 10)
   */
  Lib.copyPriceFromPrimaryToSheets = function(processed, targetSheetNames, priceMultiplier) {
    try {
      if (!processed || !processed.headers || !processed.rows) {
        Lib.logError('[copyPriceFromPrimaryToSheets] Не переданы обработанные данные');
        return;
      }

      // По умолчанию множитель = 1 (без изменений)
      const multiplier = priceMultiplier || 1;
      if (multiplier !== 1) {
        Lib.logInfo('[copyPriceFromPrimaryToSheets] Применяется множитель цены: ' + multiplier);
      }

      // Находим индексы столбцов в обработанных данных
      const articleIndex = processed.headers.indexOf('Арт. произв.');
      const priceIndex = processed.headers.indexOf('Цена');

      if (articleIndex === -1) {
        Lib.logError('[copyPriceFromPrimaryToSheets] В обработанных данных не найден столбец "Арт. произв."');
        return;
      }

      if (priceIndex === -1) {
        Lib.logWarn('[copyPriceFromPrimaryToSheets] В обработанных данных не найден столбец "Цена"');
        return;
      }

      // Создаем карту Арт. произв. -> Цена из обработанных данных
      const articleToPriceMap = {};
      for (let i = 0; i < processed.rows.length; i++) {
        const row = processed.rows[i];
        const article = String(row[articleIndex] || '').trim();
        const price = row[priceIndex];

        if (article && price !== '' && price !== null && price !== undefined) {
          let priceNum = null;
          if (typeof price === 'number') {
            priceNum = price;
          } else if (typeof price === 'string') {
            const cleaned = price.replace(/\s+/g, '').replace(',', '.');
            priceNum = parseFloat(cleaned);
          } else {
            priceNum = parseFloat(price);
          }

          if (!isNaN(priceNum)) {
            // Применяем множитель к цене
            articleToPriceMap[article] = priceNum * multiplier;
          }
        }
      }

      const mapSize = Object.keys(articleToPriceMap).length;
      Lib.logInfo('[copyPriceFromPrimaryToSheets] Создана карта Арт. произв. -> Цена, записей: ' + mapSize);

      if (mapSize === 0) {
        Lib.logWarn('[copyPriceFromPrimaryToSheets] Нет данных для копирования цены');
        return;
      }

      const ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) {
        Lib.logError('[copyPriceFromPrimaryToSheets] Не удалось получить активную таблицу');
        return;
      }

      // Список листов по умолчанию
      const defaultSheets = [
        Lib.CONFIG.SHEETS.PRICE_DYNAMICS,    // Динамика цены
        Lib.CONFIG.SHEETS.PRICE_CALCULATION, // Расчет цены
      ];

      const targetSheets = targetSheetNames || defaultSheets;
      let totalUpdated = 0;

      // Копируем цену на каждый целевой лист
      targetSheets.forEach(function(sheetName) {
        if (!sheetName) return;

        const sheet = ss.getSheetByName(sheetName);
        if (!sheet) {
          Lib.logWarn('[copyPriceFromPrimaryToSheets] Лист "' + sheetName + '" не найден');
          return;
        }

        const lastRow = sheet.getLastRow();
        if (lastRow <= 1) {
          Lib.logDebug('[copyPriceFromPrimaryToSheets] Лист "' + sheetName + '" пуст, пропускаем');
          return;
        }

        const lastColumn = sheet.getLastColumn();
        if (lastColumn <= 0) {
          Lib.logDebug('[copyPriceFromPrimaryToSheets] Лист "' + sheetName + '" не имеет столбцов');
          return;
        }

        // Ищем столбцы "Арт. произв." и "ЦЕНА EXW из Б/З"
        const headers = sheet
          .getRange(1, 1, 1, lastColumn)
          .getValues()[0]
          .map(function(value) {
            return String(value || '').trim();
          });

        const targetArticleIndex = headers.indexOf('Арт. произв.');
        let priceColumnIndex = headers.indexOf('ЦЕНА EXW из Б/З');
        if (priceColumnIndex === -1) {
          priceColumnIndex = headers.indexOf('ЦЕНА EXW из Б/З, €');
        }

        if (targetArticleIndex === -1) {
          Lib.logWarn('[copyPriceFromPrimaryToSheets] Столбец "Арт. произв." не найден на листе "' + sheetName + '"');
          return;
        }

        if (priceColumnIndex === -1) {
          Lib.logDebug('[copyPriceFromPrimaryToSheets] Столбец "ЦЕНА EXW из Б/З" или "ЦЕНА EXW из Б/З, €" не найден на листе "' + sheetName + '", пропускаем');
          return;
        }

        // Читаем все данные листа (чтобы получить текущие значения цен)
        const data = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();

        // Подготавливаем массив значений для записи
        const priceValues = [];
        let updatedCount = 0;

        for (let i = 0; i < data.length; i++) {
          const row = data[i];
          const article = String(row[targetArticleIndex] || '').trim();
          let currentPrice = row[priceColumnIndex]; // текущее значение

          if (article && articleToPriceMap.hasOwnProperty(article)) {
            priceValues.push([articleToPriceMap[article]]);
            updatedCount++;
          } else {
            // Оставляем существующее значение, но преобразуем в число
            if (currentPrice !== null && currentPrice !== undefined && currentPrice !== '') {
              // Если это строка, убираем пробелы, €, запятые заменяем на точки
              if (typeof currentPrice === 'string') {
                currentPrice = currentPrice.replace(/\s+/g, '').replace('€', '').replace(',', '.');
                currentPrice = parseFloat(currentPrice);
                if (isNaN(currentPrice)) {
                  currentPrice = '';
                }
              }
            }
            priceValues.push([currentPrice]);
          }
        }

        // Записываем все значения цены одним вызовом
        if (priceValues.length > 0) {
          const targetRange = sheet.getRange(2, priceColumnIndex + 1, priceValues.length, 1);
          targetRange.setValues(priceValues);

          // Устанавливаем формат числа с 2 знаками после запятой и символом € для всего диапазона
          targetRange.setNumberFormat('0.00 "€"');

          totalUpdated += updatedCount;
          Lib.logInfo('[copyPriceFromPrimaryToSheets] Скопировано цен на лист "' + sheetName + '": ' + updatedCount + ' строк');
        }
      });

      if (totalUpdated > 0) {
        Lib.logInfo('[copyPriceFromPrimaryToSheets] Всего скопировано строк: ' + totalUpdated);
      } else {
        Lib.logInfo('[copyPriceFromPrimaryToSheets] Не найдено строк для копирования');
      }
    } catch (err) {
      Lib.logError('[copyPriceFromPrimaryToSheets] Ошибка', err);
    }
  };

  /**
   * ПУБЛИЧНАЯ: Очистить столбец "ЦЕНА EXW из Б/З" на листах Динамика цены и Расчет цены
   * Используется при обработке Б/З поставщик перед копированием новых цен
   * @param {Array<string>} sheetNames - Массив имен листов для очистки.
   *   Если не указан, очищается на листах: Динамика цены, Расчет цены
   */
  Lib.clearPriceExwColumnOnSheets = function(sheetNames) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) {
        Lib.logError('[clearPriceExwColumnOnSheets] Не удалось получить активную таблицу');
        return;
      }

      // Список листов по умолчанию
      const defaultSheets = [
        Lib.CONFIG.SHEETS.PRICE_DYNAMICS,    // Динамика цены
        Lib.CONFIG.SHEETS.PRICE_CALCULATION, // Расчет цены
      ];

      const targetSheets = sheetNames || defaultSheets;
      let clearedCount = 0;

      targetSheets.forEach(function(sheetName) {
        if (!sheetName) return;

        const sheet = ss.getSheetByName(sheetName);
        if (!sheet) {
          Lib.logWarn('[clearPriceExwColumnOnSheets] Лист "' + sheetName + '" не найден');
          return;
        }

        const lastRow = sheet.getLastRow();
        if (lastRow <= 1) {
          Lib.logDebug('[clearPriceExwColumnOnSheets] Лист "' + sheetName + '" пуст, пропускаем');
          return;
        }

        const lastColumn = sheet.getLastColumn();
        if (lastColumn <= 0) {
          Lib.logDebug('[clearPriceExwColumnOnSheets] Лист "' + sheetName + '" не имеет столбцов');
          return;
        }

        // Ищем столбец "ЦЕНА EXW из Б/З" или "ЦЕНА EXW из Б/З, €"
        const headers = sheet
          .getRange(1, 1, 1, lastColumn)
          .getValues()[0]
          .map(function(value) {
            return String(value || '').trim();
          });

        let priceIndex = headers.indexOf('ЦЕНА EXW из Б/З');
        if (priceIndex === -1) {
          priceIndex = headers.indexOf('ЦЕНА EXW из Б/З, €');
        }
        if (priceIndex === -1) {
          Lib.logDebug('[clearPriceExwColumnOnSheets] Столбец "ЦЕНА EXW из Б/З" или "ЦЕНА EXW из Б/З, €" не найден на листе "' + sheetName + '"');
          return;
        }

        // Очищаем столбец "ЦЕНА EXW из Б/З" (со 2-й строки)
        sheet.getRange(2, priceIndex + 1, lastRow - 1, 1).clearContent();
        clearedCount++;
        Lib.logInfo('[clearPriceExwColumnOnSheets] Очищен столбец "ЦЕНА EXW из Б/З" на листе "' + sheetName + '"');
      });

      if (clearedCount > 0) {
        Lib.logInfo('[clearPriceExwColumnOnSheets] Всего очищено листов: ' + clearedCount);
      }
    } catch (err) {
      Lib.logError('[clearPriceExwColumnOnSheets] Ошибка', err);
    }
  };

  /**
   * ПУБЛИЧНАЯ: Очистить столбец ID-P на всех указанных листах
   * Используется при обработке Б/З поставщик перед назначением новых ID-P
   * @param {Array<string>} sheetNames - Массив имен листов для очистки ID-P.
   *   Если не указан, очищается ID-P на листах: Главная, Заказ, Динамика цены, Расчет цены, ABC-Анализ
   */
  Lib.clearIdpColumnOnSheets = function(sheetNames) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) {
        Lib.logError('[clearIdpColumnOnSheets] Не удалось получить активную таблицу');
        return;
      }

      // Список листов по умолчанию
      const defaultSheets = [
        Lib.CONFIG.SHEETS.PRIMARY,           // Главная
        Lib.CONFIG.SHEETS.ORDER_FORM,        // Заказ
        Lib.CONFIG.SHEETS.PRICE_DYNAMICS,    // Динамика цены
        Lib.CONFIG.SHEETS.PRICE_CALCULATION, // Расчет цены
        Lib.CONFIG.SHEETS.ABC_ANALYSIS,      // ABC-Анализ
      ];

      const targetSheets = sheetNames || defaultSheets;
      let clearedCount = 0;

      targetSheets.forEach(function(sheetName) {
        if (!sheetName) return;

        const sheet = ss.getSheetByName(sheetName);
        if (!sheet) {
          Lib.logWarn('[clearIdpColumnOnSheets] Лист "' + sheetName + '" не найден');
          return;
        }

        const lastRow = sheet.getLastRow();
        if (lastRow <= 1) {
          Lib.logDebug('[clearIdpColumnOnSheets] Лист "' + sheetName + '" пуст, пропускаем');
          return;
        }

        const lastColumn = sheet.getLastColumn();
        if (lastColumn <= 0) {
          Lib.logDebug('[clearIdpColumnOnSheets] Лист "' + sheetName + '" не имеет столбцов');
          return;
        }

        // Ищем столбец ID-P
        const headers = sheet
          .getRange(1, 1, 1, lastColumn)
          .getValues()[0]
          .map(function(value) {
            return String(value || '').trim();
          });

        const idpIndex = headers.indexOf('ID-P');
        if (idpIndex === -1) {
          Lib.logDebug('[clearIdpColumnOnSheets] Столбец ID-P не найден на листе "' + sheetName + '"');
          return;
        }

        // Очищаем столбец ID-P (со 2-й строки)
        sheet.getRange(2, idpIndex + 1, lastRow - 1, 1).clearContent();
        clearedCount++;
        Lib.logInfo('[clearIdpColumnOnSheets] Очищен столбец ID-P на листе "' + sheetName + '"');
      });

      if (clearedCount > 0) {
        Lib.logInfo('[clearIdpColumnOnSheets] Всего очищено листов: ' + clearedCount);
      }
    } catch (err) {
      Lib.logError('[clearIdpColumnOnSheets] Ошибка', err);
    }
  };

  /**
   * Очистка столбца ID-L на указанных листах
   * @param {Array<string>} sheetNames - Массив названий листов (по умолчанию: Прайс)
   */
  Lib.clearIdlColumnOnSheets = function(sheetNames) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) {
        Lib.logError('[clearIdlColumnOnSheets] Не удалось получить активную таблицу');
        return;
      }

      // Список листов по умолчанию (ID-L используется только на листе "Прайс")
      const defaultSheets = [
        Lib.CONFIG.SHEETS.PRICE,  // Прайс
      ];

      const targetSheets = sheetNames || defaultSheets;
      let clearedCount = 0;

      targetSheets.forEach(function(sheetName) {
        if (!sheetName) return;

        const sheet = ss.getSheetByName(sheetName);
        if (!sheet) {
          Lib.logWarn('[clearIdlColumnOnSheets] Лист "' + sheetName + '" не найден');
          return;
        }

        const lastRow = sheet.getLastRow();
        if (lastRow <= 1) {
          Lib.logDebug('[clearIdlColumnOnSheets] Лист "' + sheetName + '" пуст, пропускаем');
          return;
        }

        const lastColumn = sheet.getLastColumn();
        if (lastColumn <= 0) {
          Lib.logDebug('[clearIdlColumnOnSheets] Лист "' + sheetName + '" не имеет столбцов');
          return;
        }

        // Ищем столбец ID-L
        const headers = sheet
          .getRange(1, 1, 1, lastColumn)
          .getValues()[0]
          .map(function(value) {
            return String(value || '').trim();
          });

        const idlIndex = headers.indexOf('ID-L');
        if (idlIndex === -1) {
          Lib.logDebug('[clearIdlColumnOnSheets] Столбец ID-L не найден на листе "' + sheetName + '"');
          return;
        }

        // Очищаем столбец ID-L (со 2-й строки)
        sheet.getRange(2, idlIndex + 1, lastRow - 1, 1).clearContent();
        clearedCount++;
        Lib.logInfo('[clearIdlColumnOnSheets] Очищен столбец ID-L на листе "' + sheetName + '"');
      });

      if (clearedCount > 0) {
        Lib.logInfo('[clearIdlColumnOnSheets] Всего очищено листов: ' + clearedCount);
      }
    } catch (err) {
      Lib.logError('[clearIdlColumnOnSheets] Ошибка', err);
    }
  };

  // =======================================================================================
  // МОДУЛЬ 6: ЖУРНАЛ СИНХРОНИЗАЦИИ
  // ---------------------------------------------------------------------------------------
  // =======================================================================================

  /**
   * Запись строки в «Журнал синхро»
   */
  Lib.logSynchronization = function (
    key,
    source,
    target,
    oldValue,
    newValue,
    category,
    hashtag,
    event = "SYNC"
  ) {
    try {
      const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
        Lib.CONFIG.SHEETS.LOG
      );
      if (!sh) {
        Lib.logError("Журнал не найден");
        return;
      }

      const ts = new Date();
      const toStr = (v) => {
        const empty = "[пусто]";
        if (v === null || v === undefined) return empty;
        if (v instanceof Date)
          return Utilities.formatDate(
            v,
            Lib.CONFIG.SETTINGS.TIMEZONE,
            "dd.MM.yyyy HH:mm:ss"
          );
        const s = String(v);
        return s.trim() === "" ? empty : s.substring(0, 500);
      };

      // HYPERLINK по ключу → на исходную ячейку (если удалось вычислить строку)
      let keyCell = toStr(key);
      const match = String(source || "").match(/R(\d+)/);
      if (match && match[1]) {
        const srcSheetName = String(source).split("!")[0];
        const srcSheet =
          SpreadsheetApp.getActiveSpreadsheet().getSheetByName(srcSheetName);
        if (srcSheet) {
          const url = `${SpreadsheetApp.getActiveSpreadsheet().getUrl()}#gid=${srcSheet.getSheetId()}&range=A${
            match[1]
          }`;
          keyCell = `=HYPERLINK("${url}"; "${toStr(key)}")`;
        }
      }

      const rowData = [
        ts,
        keyCell,
        toStr(source),
        toStr(target),
        toStr(oldValue),
        toStr(newValue),
        toStr(category),
        toStr(hashtag),
        toStr(event),
      ];

      // Попытка добавить строку с обработкой ошибок валидации
      try {
        sh.appendRow(rowData);
      } catch (appendError) {
        // Если ошибка связана с валидацией данных, пытаемся очистить валидацию и повторить
        const errMsg = String(appendError.message || appendError);
        if (
          errMsg.indexOf("правила проверки данных") !== -1 ||
          errMsg.indexOf("data validation") !== -1 ||
          errMsg.indexOf("validation rule") !== -1
        ) {
          Lib.logWarn(
            `[LOG] Обнаружена ошибка валидации при добавлении строки. Очищаем валидацию...`
          );

          try {
            // Принудительная очистка валидации со всего листа
            const maxRows = sh.getMaxRows();
            const maxCols = sh.getMaxColumns();
            if (maxRows > 0 && maxCols > 0) {
              sh.getRange(1, 1, maxRows, maxCols).clearDataValidations();
            }

            // Повторная попытка добавления строки
            sh.appendRow(rowData);
            Lib.logWarn(
              `[LOG] Валидация очищена, строка успешно добавлена в журнал`
            );
          } catch (retryError) {
            // Если и после очистки не получилось - логируем критическую ошибку
            Lib.logError(
              `[LOG] КРИТИЧНО: Не удалось добавить строку в журнал даже после очистки валидации`,
              retryError
            );

            // В крайнем случае пытаемся добавить упрощённую версию без формул
            try {
              const simpleRowData = [
                ts,
                toStr(key), // без HYPERLINK
                toStr(source),
                toStr(target),
                toStr(oldValue),
                toStr(newValue),
                toStr(category),
                toStr(hashtag),
                toStr(event),
              ];
              sh.appendRow(simpleRowData);
              Lib.logWarn(
                `[LOG] Добавлена упрощённая версия строки (без гиперссылки)`
              );
            } catch (finalError) {
              // Если совсем ничего не помогло - тихо игнорируем, чтобы не ломать синхронизацию
              Lib.logError(
                `[LOG] ФАТАЛЬНО: Невозможно записать в журнал. Пропускаем запись.`,
                finalError
              );
            }
          }
        } else {
          // Если ошибка не связана с валидацией - пробрасываем дальше
          throw appendError;
        }
      }
    } catch (err) {
      Lib.logError("logSynchronization: ошибка", err);
    }
  };

  /**
   * Удобная запись ошибки синхронизации в «Журнал синхро»
   * — безопасна к rule === undefined/null/битому объекту
   */
  Lib.logSyncError = function (
    key,
    rule,
    sourceInfo,
    errorMessage,
    eventType = "ERROR"
  ) {
    const r = rule && typeof rule === "object" ? rule : {};
    const category = r.category || "[Ошибка]";
    const hashtags = r.hashtags || "[Ошибка]";

    // Собираем targetInfo максимально безопасно
    let targetInfo;
    if (r?.isExternal) {
      const doc = r?.targetDocId || "?";
      const sh = r?.targetSheet || "?";
      const col = r?.targetHeader || "?";
      targetInfo = `[Ext:${doc}] ${sh}!(${col})`;
    } else {
      const sh = r?.targetSheet || "?";
      const col = r?.targetHeader || "?";
      targetInfo = `[${sh}]!(${col})`;
    }

    Lib.logSynchronization(
      key,
      sourceInfo || "[?]",
      targetInfo,
      "[Ошибка]",
      `[${errorMessage}]`,
      category,
      hashtags,
      eventType
    );
  };

  // =======================================================================================
  // МОДУЛЬ 6.1: Сервисные операции с листом «Журнал синхро»
  // ---------------------------------------------------------------------------------------
  // Экспорт: Lib.quickCleanLogSheet, Lib.recreateLogSheet
  // =======================================================================================

  /** Единые заголовки журнала — строго в том же порядке, что и appendRow в logSynchronization */
  Lib.__LOG_HEADERS__ = [
    "Дата/время",
    "ID",
    "Источник",
    "Цель",
    "Старое значение",
    "Новое значение",
    "Категория",
    "Хэштеги",
    "Событие",
  ];

  /** Возвращает лист журнала, при необходимости создаёт и инициализирует структуру */
  function _getOrCreateLogSheet_() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logName = Lib.CONFIG.SHEETS.LOG;
    let sh = ss.getSheetByName(logName);

    if (!sh) {
      sh = ss.insertSheet(logName);
      _initLogSheetStructure_(sh);
      Lib.logInfo(`Создан лист журнала "${logName}"`);
      return sh;
    }

    // Базовая проверка шапки: если что-то не сходится — донастроим
    const lastCol = sh.getLastColumn();
    const headers =
      lastCol > 0
        ? sh
            .getRange(1, 1, 1, lastCol)
            .getValues()[0]
            .map((v) => String(v || "").trim())
        : [];
    let needFix = headers.length < Lib.__LOG_HEADERS__.length;
    for (
      let i = 0;
      i < Math.min(headers.length, Lib.__LOG_HEADERS__.length);
      i++
    ) {
      if (headers[i] !== Lib.__LOG_HEADERS__[i]) {
        needFix = true;
        break;
      }
    }
    if (needFix) _initLogSheetStructure_(sh);

    return sh;
  }

  /** Инициализирует/чинит структуру листа журнала: шапка, форматирование, фриз */
  function _initLogSheetStructure_(sh) {
    // КРИТИЧНО: Принудительно удаляем ВСЮ валидацию данных с листа
    // Это предотвращает ошибки "правила проверки данных" при добавлении строк
    try {
      Lib.logDebug("[LOG] Очистка валидации данных с листа журнала...");
      const dataRange = sh.getDataRange();
      if (dataRange) {
        dataRange.clearDataValidations();
      }

      // Дополнительно очищаем весь лист (на случай пустых ячеек с валидацией)
      const maxRows = sh.getMaxRows();
      const maxCols = sh.getMaxColumns();
      if (maxRows > 0 && maxCols > 0) {
        sh.getRange(1, 1, maxRows, maxCols).clearDataValidations();
      }

      Lib.logDebug("[LOG] Валидация данных успешно очищена");
    } catch (e) {
      Lib.logWarn(
        `[LOG] Не удалось очистить валидацию данных: ${e.message}. Продолжаем...`
      );
    }

    // Удаляем защищённые диапазоны (если есть)
    try {
      const protections = sh.getProtections(
        SpreadsheetApp.ProtectionType.RANGE
      );
      if (protections.length > 0) {
        Lib.logDebug(
          `[LOG] Удаление ${protections.length} защищённых диапазонов...`
        );
        protections.forEach((p) => {
          try {
            p.remove();
          } catch (_) {
            /* ignore */
          }
        });
      }
    } catch (e) {
      Lib.logWarn(
        `[LOG] Не удалось очистить защищённые диапазоны: ${e.message}`
      );
    }

    // Очистим только первую строку и нужное число столбцов под шапку
    if (sh.getMaxColumns() < Lib.__LOG_HEADERS__.length) {
      sh.insertColumnsAfter(
        Math.max(1, sh.getMaxColumns()),
        Lib.__LOG_HEADERS__.length - sh.getMaxColumns()
      );
    }
    sh.getRange(1, 1, 1, Lib.__LOG_HEADERS__.length)
      .setValues([Lib.__LOG_HEADERS__])
      .setFontWeight("bold")
      .setWrap(true);

    // Заморозим шапку
    sh.setFrozenRows(1);

    // Формат времени для первого столбца на всю колонку (кроме шапки)
    const maxRows2 = sh.getMaxRows();
    if (maxRows2 > 1) {
      sh.getRange(2, 1, maxRows2 - 1, 1).setNumberFormat("dd.MM.yyyy HH:mm:ss");
    }

    // Немного удобных ширин (не обязательно)
    try {
      sh.setColumnWidth(1, 160); // Дата/время
      sh.setColumnWidth(2, 180); // ID (там HYPERLINK)
      sh.setColumnWidth(3, 220); // Источник
      sh.setColumnWidth(4, 220); // Цель
    } catch (_) {
      /* не критично */
    }
  }

  /**
   * ПУБЛИЧНАЯ: Очистить журнал, оставив 100 последних записей (по дате/порядку добавления).
   * Удаляет самые старые строки сверху (строки 2..N-100).
   */
  Lib.quickCleanLogSheet = function () {
    const ui = SpreadsheetApp.getUi();
    try {
      const sh = _getOrCreateLogSheet_();
      const lastRow = sh.getLastRow(); // включая шапку
      const dataCount = Math.max(0, lastRow - 1);

      if (dataCount <= 100) {
        ui.alert(`В журнале ${dataCount} записей — очистка не требуется.`);
        return;
      }

      const toDelete = dataCount - 100;
      const confirm = ui.alert(
        "Очистка журнала",
        `Удалить ${toDelete} самых старых записей и оставить последние 100?`,
        ui.ButtonSet.YES_NO
      );
      if (confirm !== ui.Button.YES) return;

      // Удаляем одним куском: со 2-й строки (после шапки) нужное количество
      sh.deleteRows(2, toDelete);
      ui.alert(`Готово. Удалено ${toDelete}, оставлено 100 последних записей.`);
      Lib.logInfo(
        `[LOG] quickCleanLogSheet: удалено ${toDelete}, осталось 100`
      );
    } catch (e) {
      Lib.logError("quickCleanLogSheet: ошибка", e);
      SpreadsheetApp.getUi().alert(`Ошибка очистки журнала: ${e.message}`);
    }
  };

  /**
   * ПУБЛИЧНАЯ: Пересоздать журнал — удалить лист целиком и создать заново с шапкой.
   * ВНИМАНИЕ: удаляет ВСЁ содержимое.
   */
  Lib.recreateLogSheet = function () {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const name = Lib.CONFIG.SHEETS.LOG;

    const confirm = ui.alert(
      "Пересоздать журнал",
      "Удалить ВСЕ записи журнала и пересоздать структуру? Это действие необратимо.",
      ui.ButtonSet.YES_NO
    );
    if (confirm !== ui.Button.YES) return;

    try {
      const existing = ss.getSheetByName(name);
      if (existing) ss.deleteSheet(existing);

      const sh = ss.insertSheet(name);
      _initLogSheetStructure_(sh);

      ui.alert("Журнал пересоздан.");
      Lib.logDebug("[LOG] recreateLogSheet: журнал пересоздан");
    } catch (e) {
      Lib.logError("recreateLogSheet: ошибка", e);
      ui.alert(`Ошибка при пересоздании журнала: ${e.message}`);
    }
  };

  /**
   * ПУБЛИЧНАЯ: Пересоздать ТОЛЬКО журнал логов (DEBUG).
   */
  Lib.recreateDebugLogSheet = function () {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const name = Lib.CONFIG.SHEETS.LOG_DEBUG;

    if (!name) {
      ui.alert("Имя листа 'Журнал логов' не задано в конфигурации.");
      return;
    }

    const confirm = ui.alert(
      "Пересоздать журнал логов",
      "Очистить технический журнал (DEBUG)? Основной журнал не будет затронут.",
      ui.ButtonSet.YES_NO
    );
    if (confirm !== ui.Button.YES) return;

    try {
      const existing = ss.getSheetByName(name);
      if (existing) ss.deleteSheet(existing);

      const sh = ss.insertSheet(name);
      // Используем ту же структуру, что и для основного журнала
      _initLogSheetStructure_(sh);

      ui.alert("Журнал логов пересоздан.");
      Lib.logDebug("[LOG] recreateDebugLogSheet: журнал логов пересоздан");
    } catch (e) {
      Lib.logError("recreateDebugLogSheet: ошибка", e);
      ui.alert(`Ошибка: ${e.message}`);
    }
  };

  // =======================================================================================
  // МОДУЛЬ 7: ПОЛНАЯ СИНХРОНИЗАЦИЯ (header-driven, без привязки к порядку колонок)
  // ---------------------------------------------------------------------------------------
  // Запускается из меню "Проверить и синхронизировать ВСЮ таблицу".
  // Работает батчами, продолжение через time-based trigger `continueFullSync_proxy`.
  // =======================================================================================

  (function () {
    // Ключ для состояния полной синхронизации
    const STATE_KEY = "ECOSYS_FULLSYNC_STATE_V1";
    // Имя прокси-функции в файле-обёртке (нужна для таймерного продолжения)
    const CONTINUE_HANDLER = "continueFullSync_proxy";

    /**
     * ПУБЛИЧНАЯ: старт полной синхронизации (через меню)
     * - валидирует и группирует правила;
     * - сбрасывает прошлое состояние/триггеры;
     * - сохраняет state в ScriptProperties;
     * - запускает первую итерацию Lib._continueFullSync().
     */
    Lib.runFullSync = function () {
      const ui = SpreadsheetApp.getUi();
      const ss = SpreadsheetApp.getActiveSpreadsheet();

      const resp = ui.alert(
        "Запуск полной синхронизации",
        "Операция проверит и, при необходимости, исправит все несоответствия по активным правилам. Продолжить?",
        ui.ButtonSet.YES_NO
      );
      if (resp !== ui.Button.YES) return;

      const lock = LockService.getScriptLock();
      if (!lock.tryLock(15000)) {
        ui.alert("Процесс уже выполняется.");
        return;
      }

      try {
        // 1) Сбросить прошлое состояние + удалить таймеры продолжения
        PropertiesService.getScriptProperties().deleteProperty(STATE_KEY);
        ScriptApp.getProjectTriggers().forEach((t) => {
          try {
            if (
              typeof t.getHandlerFunction === "function" &&
              t.getHandlerFunction() === CONTINUE_HANDLER
            ) {
              ScriptApp.deleteTrigger(t);
            }
          } catch (_) {
            /* no-op */
          }
        });

        // 2) Загрузить и ОТФИЛЬТРОВАТЬ правила
        const allRulesRaw = _loadSyncRules(true);
        const allRules = (allRulesRaw || []).filter(
          (r) =>
            r &&
            typeof r === "object" &&
            typeof r.sourceSheet === "string" &&
            r.sourceSheet &&
            typeof r.sourceHeader === "string" &&
            r.sourceHeader &&
            typeof r.targetSheet === "string" &&
            r.targetSheet &&
            typeof r.targetHeader === "string" &&
            r.targetHeader &&
            (!r.isExternal ||
              (typeof r.targetDocId === "string" && r.targetDocId))
        );
        if (allRules.length === 0) {
          ui.alert("Активные валидные правила не найдены.");
          return;
        }

        // 3) Сгруппировать по исходному листу
        const rulesBySourceSheet = allRules.reduce((acc, r) => {
          (acc[r.sourceSheet] = acc[r.sourceSheet] || []).push(r);
          return acc;
        }, {});
        const sheetsToProcess = Object.keys(rulesBySourceSheet);
        if (sheetsToProcess.length === 0) {
          ui.alert("Нет листов-источников для проверки.");
          return;
        }

        // 4) Сохранить состояние
        const state = {
          version: 1,
          sheets: sheetsToProcess, // список листов-источников
          rulesBySheet: rulesBySourceSheet, // карта: источник -> массив правил
          sheetIndex: 0, // текущий индекс листа
          lastProcessedRow: 0, // от какой строки продолжать
          totalCorrections: 0, // счётчик исправлений
          startedAtUtc: new Date().toISOString(),
        };
        PropertiesService.getScriptProperties().setProperty(
          STATE_KEY,
          JSON.stringify(state)
        );

        ss.toast(
          `Старт полной синхронизации… (${sheetsToProcess.length} листов)`,
          "Full Sync",
          8
        );

        // 5) Первая итерация (дальше _continueFullSync сама поставит таймер, если надо)
        Lib._continueFullSync();
      } catch (e) {
        Lib.logError("[FullSync] Ошибка запуска runFullSync", e);
        ui.alert(`Ошибка запуска: ${e && e.message ? e.message : e}`);
      } finally {
        lock.releaseLock();
      }
    };

    // Экспортируем ключи — удобно использовать из Lib._continueFullSync()
    Lib.__FULLSYNC_STATE_KEY__ = STATE_KEY;
    Lib.__FULLSYNC_CONTINUE_HANDLER__ = CONTINUE_HANDLER;

    /**
     * ПУБЛИЧНАЯ: автоматический запуск полной синхронизации БЕЗ диалогов
     * Используется для автоматического запуска после добавления новых артикулов
     */
    Lib.runFullSyncSilent = function () {
      const lock = LockService.getScriptLock();
      if (!lock.tryLock(15000)) {
        Lib.logWarn("[FullSync Silent] Процесс уже выполняется.");
        return;
      }

      try {
        // 1) Сбросить прошлое состояние + удалить таймеры продолжения
        PropertiesService.getScriptProperties().deleteProperty(STATE_KEY);
        ScriptApp.getProjectTriggers().forEach((t) => {
          try {
            if (
              typeof t.getHandlerFunction === "function" &&
              t.getHandlerFunction() === CONTINUE_HANDLER
            ) {
              ScriptApp.deleteTrigger(t);
            }
          } catch (_) {
            /* no-op */
          }
        });

        // 2) Загрузить и ОТФИЛЬТРОВАТЬ правила
        const allRulesRaw = _loadSyncRules(true);
        const allRules = (allRulesRaw || []).filter(
          (r) =>
            r &&
            typeof r === "object" &&
            typeof r.sourceSheet === "string" &&
            r.sourceSheet &&
            typeof r.sourceHeader === "string" &&
            r.sourceHeader &&
            typeof r.targetSheet === "string" &&
            r.targetSheet &&
            typeof r.targetHeader === "string" &&
            r.targetHeader &&
            (!r.isExternal ||
              (typeof r.targetDocId === "string" && r.targetDocId))
        );
        if (allRules.length === 0) {
          Lib.logInfo("[FullSync Silent] Активные валидные правила не найдены.");
          return;
        }

        // 3) Сгруппировать по исходному листу
        const rulesBySourceSheet = allRules.reduce((acc, r) => {
          (acc[r.sourceSheet] = acc[r.sourceSheet] || []).push(r);
          return acc;
        }, {});
        const sheetsToProcess = Object.keys(rulesBySourceSheet);
        if (sheetsToProcess.length === 0) {
          Lib.logInfo("[FullSync Silent] Нет листов-источников для проверки.");
          return;
        }

        // 4) Сохранить состояние
        const state = {
          version: 1,
          sheets: sheetsToProcess,
          rulesBySheet: rulesBySourceSheet,
          sheetIndex: 0,
          lastProcessedRow: 0,
          totalCorrections: 0,
          startedAtUtc: new Date().toISOString(),
        };
        PropertiesService.getScriptProperties().setProperty(
          STATE_KEY,
          JSON.stringify(state)
        );

        Lib.logInfo(
          `[FullSync Silent] Старт автоматической синхронизации (${sheetsToProcess.length} листов)`
        );

        // 5) Первая итерация
        Lib._continueFullSync();
      } catch (e) {
        Lib.logError("[FullSync Silent] Ошибка запуска", e);
      } finally {
        lock.releaseLock();
      }
    };
  })();

  // --- Санитизация массива правил: выбросить всё битое/пустое, выровнять типы
  function _sanitizeRulesArray_(rulesRaw) {
    const arr = Array.isArray(rulesRaw) ? rulesRaw : [];
    const out = [];
    let dropped = 0;

    for (const r of arr) {
      if (!r || typeof r !== "object") {
        dropped++;
        continue;
      }

      const sourceSheet = String(r.sourceSheet || "").trim();
      const sourceHeader = String(r.sourceHeader || "").trim();
      const targetSheet = String(r.targetSheet || "").trim();
      const targetHeader = String(r.targetHeader || "").trim();
      const isExternal =
        r.isExternal === true ||
        String(r.isExternal || "").toLowerCase() === "да";
      const targetDocId = String(r.targetDocId || "").trim();

      if (!sourceSheet || !sourceHeader || !targetSheet || !targetHeader) {
        dropped++;
        continue;
      }
      if (isExternal && !targetDocId) {
        dropped++;
        continue;
      }

      out.push({
        id: r.id || "",
        category: r.category || "",
        hashtags: r.hashtags || "",
        sourceSheet,
        sourceHeader,
        targetSheet,
        targetHeader,
        isExternal,
        targetDocId,
      });
    }

    if (dropped)
      Lib.logWarn(`[FullSync] Отфильтровано некорректных правил: ${dropped}`);
    return out;
  }

  // ==== CONTINUE: батчевое продолжение полной синхронизации по таймеру ====
  Lib._continueFullSync = function () {
    const STATE_KEY = Lib.__FULLSYNC_STATE_KEY__ || "ECOSYS_FULLSYNC_STATE_V1";
    const CONTINUE_HANDLER =
      Lib.__FULLSYNC_CONTINUE_HANDLER__ || "continueFullSync_proxy";

    const lock = LockService.getScriptLock();
    if (!lock.tryLock(15000)) {
      Lib.logWarn("[FullSync] Занято другим запуском");
      return;
    }

    try {
      const props = PropertiesService.getScriptProperties();
      let state = {};
      try {
        state = JSON.parse(props.getProperty(STATE_KEY) || "{}");
      } catch (_) {
        state = {};
      }

      if (!state || !Array.isArray(state.sheets) || state.sheets.length === 0) {
        Lib.logWarn("[FullSync] Состояние пустое — выходим");
        props.deleteProperty(STATE_KEY);
        return;
      }

      const MAX_MS =
        Lib.CONFIG?.SETTINGS?.FULLSYNC_MAX_RUNTIME_MS || 4 * 60 * 1000;
      const MAX_ROWS = Lib.CONFIG?.SETTINGS?.FULLSYNC_MAX_ROWS_PER_TICK || 300;
      const MAX_FIX =
        Lib.CONFIG?.SETTINGS?.FULLSYNC_MAX_CORRECTIONS_PER_TICK || 1000;
      const deadline = Date.now() + MAX_MS;

      while (state.sheetIndex < state.sheets.length) {
        const sourceName = state.sheets[state.sheetIndex];

        // санитарная фильтрация правил из state
        const rawRules =
          (state.rulesBySheet && state.rulesBySheet[sourceName]) || [];
        const rules = (Array.isArray(rawRules) ? rawRules : []).filter(
          (r) =>
            r &&
            typeof r === "object" &&
            typeof r.sourceSheet === "string" &&
            r.sourceSheet &&
            typeof r.sourceHeader === "string" &&
            r.sourceHeader &&
            typeof r.targetSheet === "string" &&
            r.targetSheet &&
            typeof r.targetHeader === "string" &&
            r.targetHeader
        );

        const res = _processSheetForFullSync_(
          sourceName,
          rules,
          Number(state.lastProcessedRow || 0),
          {
            maxRows: MAX_ROWS,
            maxCorrections: MAX_FIX,
            endAt: deadline,
          }
        );

        state.totalCorrections =
          (state.totalCorrections || 0) + (res.corrections || 0);
        state.lastProcessedRow = res.nextRowIndex0 || 0;

        // сохранить прогресс
        props.setProperty(STATE_KEY, JSON.stringify(state));

        // если ещё не дошли до конца листа или упёрлись во время/лимит — поставить таймер и выйти
        if (!res.reachedEnd || Date.now() >= deadline) {
          ScriptApp.newTrigger(CONTINUE_HANDLER)
            .timeBased()
            .after(1000)
            .create();
          SpreadsheetApp.getActiveSpreadsheet().toast(
            `Full Sync… ${sourceName}: обработано до строки ${state.lastProcessedRow}, исправлений: ${state.totalCorrections}`,
            "Продолжаю",
            5
          );
          return;
        }

        // лист закончен — перейти к следующему
        state.sheetIndex += 1;
        state.lastProcessedRow = 0;
      }

      // всё готово
      props.deleteProperty(STATE_KEY);
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `Полная синхронизация завершена. Исправлений: ${
          state.totalCorrections || 0
        }`,
        "Готово",
        8
      );
    } catch (e) {
      Lib.logError("[FullSync] _continueFullSync: ошибка", e);
      try {
        SpreadsheetApp.getUi().alert(
          `FullSync: ${e && e.message ? e.message : e}`
        );
      } catch (_) {}
    } finally {
      lock.releaseLock();
    }
  };

  /**
   * ПРИВАТНАЯ: обработка одного исходного листа, начиная с заданной строки.
   * Возвращает { corrections, lastRow, isFinished }.
   * — Чтение и запись строго по ИМЕНАМ заголовков (header-driven).
   * — Работает с локальными и внешними таблицами.
   */
  /**
   * Обработка одного исходного листа батчем.
   * Возвращает { corrections, reachedEnd, nextRowIndex0 }.
   * - Чтение/запись строго по ИМЕНАМ заголовков.
   * - Защита от битых правил.
   * - Поддержка ограничений по времени/строкам/исправлениям (opts).
   */
  function _processSheetForFullSync_(
    sourceSheetName,
    rulesRaw,
    startRowIndex0,
    opts
  ) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // ---- Опции по умолчанию
    opts = opts || {};
    const DEADLINE_MS =
      typeof opts.endAt === "number" ? opts.endAt : Date.now() + 4 * 60 * 1000;
    const MAX_ROWS = Number.isFinite(opts.maxRows) ? opts.maxRows : 300;
    const MAX_CORRECTIONS = Number.isFinite(opts.maxCorrections)
      ? opts.maxCorrections
      : 1000;

    let correctionsCount = 0;

    // ---- Санитизируем правила
    const rules = _sanitizeRulesArray_(rulesRaw);
    if (rules.length === 0) {
      Lib.logInfo(`[FullSync] Для "${sourceSheetName}" нет валидных правил.`);
      return { corrections: 0, reachedEnd: true, nextRowIndex0: 0 };
    }

    // ---- Источник
    const sourceSheet = ss.getSheetByName(sourceSheetName);
    if (!sourceSheet) {
      Lib.logWarn(`[FullSync] Исходный лист не найден: ${sourceSheetName}`);
      return { corrections: 0, reachedEnd: true, nextRowIndex0: 0 };
    }

    const sourceValues = sourceSheet.getDataRange().getValues();
    if (!sourceValues || sourceValues.length <= 1) {
      return { corrections: 0, reachedEnd: true, nextRowIndex0: 0 };
    }

    const sourceHeaders = (sourceValues[0] || []).map((h) =>
      String(h || "").trim()
    );
    const sourceHeaderMap = {};
    sourceHeaders.forEach((h, i) => {
      if (h) sourceHeaderMap[h] = i;
    });

    const sourceData = sourceValues.slice(1); // без заголовков
    const totalRows = sourceData.length;
    let i = Math.max(0, Number(startRowIndex0 || 0));

    // ---- Подготовим кэш целей: { key -> { sheet, data, headers, headerMap, keyMap } }
    const targets = {};
    for (const rule of rules) {
      const tKey = rule.isExternal
        ? `${rule.targetDocId}::${rule.targetSheet}`
        : rule.targetSheet;
      if (targets[tKey]) continue;

      try {
        const tSS = rule.isExternal
          ? SpreadsheetApp.openById(rule.targetDocId)
          : ss;
        const tSheet = tSS.getSheetByName(rule.targetSheet);
        if (!tSheet) {
          Lib.logWarn(`[FullSync] Целевой лист отсутствует: ${tKey}`);
          continue;
        }

        const tValues = tSheet.getDataRange().getValues();
        if (!tValues || tValues.length === 0) {
          Lib.logWarn(`[FullSync] Пустой целевой лист: ${tKey}`);
          continue;
        }

        const tHeaders = (tValues[0] || []).map((h) => String(h || "").trim());
        const tHeaderMap = {};
        tHeaders.forEach((h, idx) => {
          if (h) tHeaderMap[h] = idx;
        });

        const tData = tValues.slice(1);
        const tKeyMap = new Map(); // ключ (A) -> индекс строки (0-based) в tData
        tData.forEach((row, idx) => {
          const key = String(row[0] || "").trim();
          if (key) tKeyMap.set(key, idx);
        });

        targets[tKey] = {
          sheet: tSheet,
          data: tData,
          headers: tHeaders,
          headerMap: tHeaderMap,
          keyMap: tKeyMap,
        };
      } catch (e) {
        Lib.logWarn(`[FullSync] Ошибка доступа к цели ${tKey}: ${e.message}`);
      }
    }

    // ---- Накопитель изменений по целям
    const updates = {}; // tKey -> [{row, col, value, log:{...}}]

    // ---- Основной цикл по строкам источника
    const hardDeadline = DEADLINE_MS; // абсолютный timestamp
    let rowsProcessed = 0;

    while (i < totalRows) {
      // лимиты на тик
      if (Date.now() >= hardDeadline) break;
      if (rowsProcessed >= MAX_ROWS) break;
      if (correctionsCount >= MAX_CORRECTIONS) break;

      const srcRow = sourceData[i];
      const key = String(srcRow[0] || "").trim();
      if (!key) {
        i++;
        rowsProcessed++;
        continue;
      }

      const sourceRowNum = i + 2; // 1-based + заголовок

      for (const rule of rules) {
        // быстрые проверки наличия требуемых колонок
        const srcColIndex0 = sourceHeaderMap[rule.sourceHeader];
        if (srcColIndex0 === undefined) continue;

        const tKey = rule.isExternal
          ? `${rule.targetDocId}::${rule.targetSheet}`
          : rule.targetSheet;
        const T = targets[tKey];
        if (!T) continue;

        const tgtColIndex0 = T.headerMap[rule.targetHeader];
        const tgtRowIndex0 = T.keyMap.get(key);
        if (tgtColIndex0 === undefined || tgtRowIndex0 === undefined) continue;

        const sourceValue = srcRow[srcColIndex0];
        const targetValue = T.data[tgtRowIndex0][tgtColIndex0];

        if (Lib.areValuesEqual_(sourceValue, targetValue)) continue;

        (updates[tKey] = updates[tKey] || []).push({
          row: tgtRowIndex0 + 2,
          col: tgtColIndex0 + 1,
          value: sourceValue,
          log: {
            key,
            sourceInfo: `${sourceSheetName}!(${rule.sourceHeader} R${sourceRowNum})`,
            targetInfo: `${rule.targetSheet}!(${rule.targetHeader} R${
              tgtRowIndex0 + 2
            })`,
            category: rule.category,
            hashtags: rule.hashtags,
          },
        });
        correctionsCount++;

        if (correctionsCount >= MAX_CORRECTIONS) break;
        if (Date.now() >= hardDeadline) break;
      }

      i++;
      rowsProcessed++;
    }

    // ---- Запись накопленных изменений и логирование
    for (const tKey in updates) {
      const ops = updates[tKey];
      const T = targets[tKey];
      if (!T || !ops || ops.length === 0) continue;

      for (const u of ops) {
        try {
          T.sheet.getRange(u.row, u.col).setValue(u.value);
          Lib.logSynchronization(
            u.log.key,
            u.log.sourceInfo,
            u.log.targetInfo,
            "[batched old]",
            u.value,
            u.log.category,
            u.log.hashtags,
            "FORCE_SYNC"
          );
        } catch (e) {
          Lib.logError(
            `[FullSync] Ошибка записи ${tKey} R${u.row}C${u.col}: ${e.message}`,
            e
          );
        }
      }
    }

    const reachedEnd = i >= totalRows;
    return {
      corrections: correctionsCount,
      reachedEnd,
      nextRowIndex0: reachedEnd ? 0 : i,
    };
  }

  // =======================================================================================
  // МОДУЛЬ 8: Диалог «Настройка правил» (server-side для HTML)
  // ---------------------------------------------------------------------------------------
  // Требуемые серверные функции, на которые ссылается твой HTML:
  // - Lib.openRulesDialog()                    — открыть модальный диалог
  // - getSheetsList()                          — список локальных листов
  // - getSheetColumns(sheetName)               — список заголовков конкретного листа
  // - getExternalDocsList()                    — список внешних документов из служебного листа
  // - getExternalSheetsList(docId)             — список листов во внешнем документе
  // - getExternalSheetColumns(docId, sheet)    — список заголовков во внешнем листе
  // - saveSyncRule(ruleData)                   — создать/обновить правило в «Правила синхро»
  // =======================================================================================

  // ==========================
  // RULES DIALOG: SERVER SIDE
  // ==========================
  Lib.showSyncConfigDialog = function () {
    const HTML_FILE_NAME = "RuleEditor"; // имя твоего HTML-файла
    const html = HtmlService.createHtmlOutputFromFile(HTML_FILE_NAME)
      .setWidth(720)
      .setHeight(680);
    SpreadsheetApp.getUi().showModalDialog(
      html,
      "Настройка правила синхронизации"
    );
  };

  // внутренний помощник: создать/починить лист «Правила синхро»
  function _ensureRulesSheet_() {
    const name = Lib.CONFIG.SHEETS.RULES;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
    }
    // Заголовки по нашей текущей схеме чтения/записи
    const headers = [
      "ID правила", // A: row[0]
      "Активно", // B: row[1] (checkbox/Да)
      "Категория", // C: row[2]
      "Хэштеги", // D: row[3]
      "Источник: Лист", // E: row[4]
      "Источник: Колонка", // F: row[5]
      "Цель: Лист", // G: row[6]
      "Цель: Колонка", // H: row[7]
      "Внешний документ", // I: row[8] (checkbox/Да)
      "Target Doc ID", // J: row[9]
    ];

    const lastCol = sh.getLastColumn();
    const have =
      lastCol > 0
        ? sh
            .getRange(1, 1, 1, Math.max(lastCol, headers.length))
            .getValues()[0]
            .map((x) => String(x || "").trim())
        : [];
    let changed = false;
    headers.forEach((h, i) => {
      if (have[i] !== h) {
        sh.getRange(1, i + 1)
          .setValue(h)
          .setFontWeight("bold");
        changed = true;
      }
    });
    if (changed) {
      sh.setFrozenRows(1);
    }
    return sh;
  }

  // Список внешних документов:
  // 1) из глобальной карты DOC_TO_PROJECT, если она есть;
  // 2) плюс уникальные targetDocId уже встречавшиеся в правилах;
  // 3) статус ✅/⛔ в зависимости от доступа.
  Lib.getExternalDocsList = function () {
    const out = [];
    const seen = new Set();

    const pushDoc = (docId, labelHint) => {
      if (!docId || seen.has(docId)) return;
      seen.add(docId);
      let name = labelHint || "";
      let status = "⛔ нет доступа";
      try {
        const ss = SpreadsheetApp.openById(docId);
        name = name || ss.getName();
        status = "✅ доступ";
      } catch (_) {
        // нет доступа или не таблица
      }
      out.push({ docId, name: name || docId, status });
    };

    // 1) карта, если есть
    try {
      if (typeof DOC_TO_PROJECT !== "undefined" && DOC_TO_PROJECT) {
        Object.keys(DOC_TO_PROJECT).forEach((id) =>
          pushDoc(id, DOC_TO_PROJECT[id])
        );
      }
    } catch (_) {}

    // 2) из существующих правил
    try {
      const rules = (function () {
        return (function () {
          return typeof _loadSyncRules === "function"
            ? _loadSyncRules(true)
            : [];
        })();
      })();
      rules
        .filter((r) => r.isExternal && r.targetDocId)
        .forEach((r) => pushDoc(r.targetDocId, ""));
    } catch (_) {}

    // Упорядочим по имени
    out.sort((a, b) => a.name.localeCompare(b.name, "ru"));
    return out;
  };

  Lib.getExternalSheetsList = function (docId) {
    if (!docId) return [];
    try {
      const ss = SpreadsheetApp.openById(docId);
      return ss.getSheets().map((s) => s.getName());
    } catch (e) {
      throw new Error("Нет доступа к документу: " + e.message);
    }
  };

  Lib.getExternalSheetColumns = function (docId, sheetName) {
    if (!docId || !sheetName) return [];
    try {
      const ss = SpreadsheetApp.openById(docId);
      const sh = ss.getSheetByName(sheetName);
      if (!sh) return [];
      const lastCol = sh.getLastColumn();
      if (lastCol < 1) return [];
      return sh
        .getRange(1, 1, 1, lastCol)
        .getValues()[0]
        .map((h) => String(h || "").trim())
        .filter(Boolean);
    } catch (e) {
      throw new Error("Нет доступа/лист не найден: " + e.message);
    }
  };

  // Диалог "Управление внешними документами"
  Lib.showExternalDocManagerDialog = function () {
    try {
      if (typeof Lib.ensureInfra === "function") {
        Lib.ensureInfra(); // создаст лист "Внешние документы" и т.п., если нужно
      }
      var html = HtmlService.createHtmlOutputFromFile("ExternalDocsManager") // имя HTML-файла см. п.2
        .setWidth(720)
        .setHeight(640);
      SpreadsheetApp.getUi().showModalDialog(
        html,
        "Управление внешними документами"
      );
    } catch (e) {
      SpreadsheetApp.getUi().alert("Ошибка открытия диалога: " + e.message);
      Lib.logError && Lib.logError("showExternalDocManagerDialog: ошибка", e);
    }
  };

  // =======================================================================================
  // ИНФРАСТРУКТУРА: служебные листы + API для RuleEditor (диалог правил)
  // ---------------------------------------------------------------------------------------

  (function () {
    const S = (Lib.CONFIG && Lib.CONFIG.SHEETS) || {};
    const NAMES = {
      RULES: S.RULES || "Правила синхро",
      LOG: S.LOG || "Журнал синхро",
      EXTS: S.EXTERNAL_DOCS || "Внешние документы",
    };

    // --- Эталонные заголовки ---
    const LOG_HEADERS = [
      "Дата/время",
      "ID",
      "Источник",
      "Цель",
      "Старое значение",
      "Новое значение",
      "Категория",
      "Хэштеги",
      "Событие",
    ];
    const EXTS_HEADERS = [
      "Название",
      "Doc ID",
      "Статус",
      "Последняя проверка",
      "Комментарий",
    ];

    // ПУБЛИЧНО: единая автоконфигурация инфраструктуры
    Lib.ensureInfra = function () {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      // 1) Правила
      if (typeof Lib.ensureRulesSheetStructure === "function") {
        Lib.ensureRulesSheetStructure();
      }
      // 2) Журнал
      _ensureLogSheet_();
      // 3) Внешние документы
      _ensureExternalDocsSheet_();
      // 4) Лёгкая проверка «Внешних документов» (не обязательно)
      try {
        Lib.updateExternalDocsStatus && Lib.updateExternalDocsStatus();
      } catch (_) {}
      ss.toast("Инфраструктура проверена/создана", "OK", 3);
    };

    Lib.ensureRulesSheetStructure = function () {
      const name =
        (Lib.CONFIG && Lib.CONFIG.SHEETS && Lib.CONFIG.SHEETS.RULES) ||
        "Правила синхро";
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let sh = ss.getSheetByName(name);
      if (!sh) sh = ss.insertSheet(name);

      const HEADERS = [
        "ID правила",
        "Активно",
        "Категория",
        "Хэштеги",
        "Источник: Лист",
        "Источник: Колонка",
        "Цель: Лист",
        "Цель: Колонка",
        "Внешний документ",
        "Target Doc ID",
        "Дата создания",
        "Последнее изм.",
      ];

      // Шапка
      if (sh.getMaxColumns() < HEADERS.length) {
        sh.insertColumnsAfter(
          Math.max(1, sh.getMaxColumns()),
          HEADERS.length - sh.getMaxColumns()
        );
      }
      sh.getRange(1, 1, 1, HEADERS.length)
        .setValues([HEADERS])
        .setFontWeight("bold");
      sh.setFrozenRows(1);

      // «Да/Нет» на B и I
      const yn = SpreadsheetApp.newDataValidation()
        .requireValueInList(["Да", "Нет"], true)
        .build();
      sh.getRange(2, 2, Math.max(1, sh.getMaxRows() - 1), 1).setDataValidation(
        yn
      );
      sh.getRange(2, 9, Math.max(1, sh.getMaxRows() - 1), 1).setDataValidation(
        yn
      );

      // Формат времени для K/L
      sh.getRange(2, 11, Math.max(1, sh.getMaxRows() - 1), 2).setNumberFormat(
        "dd.MM.yyyy HH:mm"
      );
    };

    // ---------------- ЖУРНАЛ ----------------
    function _ensureLogSheet_() {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let sh = ss.getSheetByName(NAMES.LOG);
      if (!sh) {
        sh = ss.insertSheet(NAMES.LOG);
        sh.getRange(1, 1, 1, LOG_HEADERS.length)
          .setValues([LOG_HEADERS])
          .setFontWeight("bold");
        sh.setFrozenRows(1);
        sh.setColumnWidth(1, 165); // время
        sh.getRange(2, 1, sh.getMaxRows() - 1, 1).setNumberFormat(
          "dd.MM.yyyy HH:mm:ss"
        );
        try {
          sh.getRange(1, 1, 1, LOG_HEADERS.length).createFilter();
        } catch (_) {}
        return;
      }
      // добить шапку, если «поплыла»
      const hdr = sh
        .getRange(1, 1, 1, Math.max(LOG_HEADERS.length, sh.getLastColumn()))
        .getValues()[0]
        .map((v) => String(v || "").trim());
      const missing = LOG_HEADERS.filter((h) => !hdr.includes(h));
      if (missing.length) {
        sh.insertColumnsAfter(Math.max(1, sh.getLastColumn()), missing.length);
        sh.getRange(1, hdr.length + 1, 1, missing.length)
          .setValues([missing])
          .setFontWeight("bold");
      }
      // формат времени на первой колонке
      sh.getRange(2, 1, Math.max(0, sh.getMaxRows() - 1), 1).setNumberFormat(
        "dd.MM.yyyy HH:mm:ss"
      );
      try {
        sh.getFilter() ||
          sh.getRange(1, 1, 1, LOG_HEADERS.length).createFilter();
      } catch (_) {}
    }

    // ---------------- ВНЕШНИЕ ДОКУМЕНТЫ ----------------
    function _ensureExternalDocsSheet_() {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let sh = ss.getSheetByName(NAMES.EXTS);
      if (!sh) {
        sh = ss.insertSheet(NAMES.EXTS);
        sh.getRange(1, 1, 1, EXTS_HEADERS.length)
          .setValues([EXTS_HEADERS])
          .setFontWeight("bold");
        sh.setFrozenRows(1);
        sh.setColumnWidths(1, EXTS_HEADERS.length, 180);
        sh.setColumnWidth(2, 280); // Doc ID
        sh.getRange(2, 4, sh.getMaxRows() - 1, 1).setNumberFormat(
          "dd.MM.yyyy HH:mm"
        );
        try {
          sh.getRange(1, 1, 1, EXTS_HEADERS.length).createFilter();
        } catch (_) {}
        return;
      }
      // добить шапку
      const hdr = sh
        .getRange(1, 1, 1, Math.max(EXTS_HEADERS.length, sh.getLastColumn()))
        .getValues()[0]
        .map((v) => String(v || "").trim());
      const missing = EXTS_HEADERS.filter((h) => !hdr.includes(h));
      if (missing.length) {
        sh.insertColumnsAfter(Math.max(1, sh.getLastColumn()), missing.length);
        sh.getRange(1, hdr.length + 1, 1, missing.length)
          .setValues([missing])
          .setFontWeight("bold");
      }
      sh.getRange(2, 4, Math.max(0, sh.getMaxRows() - 1), 1).setNumberFormat(
        "dd.MM.yyyy HH:mm"
      );
      try {
        sh.getFilter() ||
          sh.getRange(1, 1, 1, EXTS_HEADERS.length).createFilter();
      } catch (_) {}
    }

    // ПУБЛИЧНО: проверить доступность внешних документов и заполнить «Статус/Последняя проверка»
    Lib.updateExternalDocsStatus = function () {
      const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
        NAMES.EXTS
      );
      if (!sh || sh.getLastRow() <= 1) return;
      const values = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();
      for (let i = 0; i < values.length; i++) {
        const name = String(values[i][0] || "").trim();
        const docId = String(values[i][1] || "").trim();
        if (!docId) continue;
        let status = "❌ Нет доступа";
        try {
          const ext = SpreadsheetApp.openById(docId);
          const title = ext.getName();
          if (title) status = "✅ Доступен";
          // если имя пусто — оставим как есть
        } catch (_) {}
        sh.getRange(i + 2, 3).setValue(status);
        sh.getRange(i + 2, 4).setValue(new Date());
      }
    };

    // ---------------- API ДЛЯ ДИАЛОГА ПРАВИЛ ----------------

    // список внутренних листов (исключая служебные)
    Lib.getSheetsList = function () {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const service = new Set([NAMES.RULES, NAMES.LOG, NAMES.EXTS]);
      return ss
        .getSheets()
        .map((s) => s.getName())
        .filter((n) => !service.has(n));
    };

    // заголовки выбранного внутреннего листа
    Lib.getSheetColumns = function (sheetName) {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sh = ss.getSheetByName(sheetName);
      if (!sh) return [];
      const header =
        sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] || [];
      return header.map((v) => String(v || "").trim()).filter(Boolean);
    };

    // список внешних документов для селекта
    Lib.getExternalDocsList = function () {
      _ensureExternalDocsSheet_();
      const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
        NAMES.EXTS
      );
      if (!sh || sh.getLastRow() <= 1) return [];
      const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();
      return rows
        .map((r) => ({
          name: String(r[0] || "").trim(),
          docId: String(r[1] || "").trim(),
          status: String(r[2] || "").trim(),
        }))
        .filter((o) => o.docId);
    };

    // листы внутри внешней таблицы
    Lib.getExternalSheetsList = function (docId) {
      if (!docId) return [];
      try {
        const ext = SpreadsheetApp.openById(docId);
        return ext.getSheets().map((s) => s.getName());
      } catch (e) {
        Lib.logWarn(`getExternalSheetsList: ${e.message}`);
        return [];
      }
    };

    // заголовки листа во внешней таблице
    Lib.getExternalSheetColumns = function (docId, sheetName) {
      if (!docId || !sheetName) return [];
      try {
        const ext = SpreadsheetApp.openById(docId);
        const sh = ext.getSheetByName(sheetName);
        if (!sh) return [];
        const header =
          sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] || [];
        return header.map((v) => String(v || "").trim()).filter(Boolean);
      } catch (e) {
        Lib.logWarn(`getExternalSheetColumns: ${e.message}`);
        return [];
      }
    };
  })();
  // Публичная: сохранить/обновить правило синхронизации из RuleEditor.html
  Lib.saveSyncRule = function (ruleData) {
    const lock = LockService.getUserLock();
    lock.waitLock(10000); // ждём до 10 секунд именно пользовательский лок

    try {
      // Подстрахуемся: инфраструктура и лист с правилами должны существовать
      if (typeof Lib.ensureInfra === "function") Lib.ensureInfra();
      const rulesSheetName =
        (Lib.CONFIG && Lib.CONFIG.SHEETS && Lib.CONFIG.SHEETS.RULES) ||
        "Правила синхро";
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sh = ss.getSheetByName(rulesSheetName);
      if (!sh) throw new Error(`Лист "${rulesSheetName}" не найден`);

      // Нормализаторы
      const toYesNo = (v) => {
        const s = String(v).trim().toLowerCase();
        return v === true ||
          s === "true" ||
          s === "1" ||
          s === "да" ||
          s === "yes" ||
          s === "y"
          ? "Да"
          : "Нет";
      };
      const asStr = (v) => String(v || "").trim();

      // Поля из формы
      const id = asStr(ruleData.id);
      const enabledYN = toYesNo(ruleData.enabled);
      const category = asStr(ruleData.category);
      const hashtags = asStr(ruleData.hashtags);
      const sourceSheet = asStr(ruleData.sourceSheet);
      const sourceHeader = asStr(ruleData.sourceHeader);
      const targetSheet = asStr(ruleData.targetSheet);
      const targetHeader = asStr(ruleData.targetHeader);
      const isExternalYN = toYesNo(ruleData.isExternal);
      const targetDocId =
        isExternalYN === "Да" ? asStr(ruleData.targetDocId) : "";

      // Валидация
      if (!sourceSheet || !sourceHeader || !targetSheet || !targetHeader) {
        throw new Error(
          "Заполните: Исходный лист/столбец и Целевой лист/столбец."
        );
      }
      if (isExternalYN === "Да" && !targetDocId) {
        throw new Error(
          "Включён внешний документ, но не указан Target Doc ID."
        );
      }

      // Эталонная шапка (A…L)
      const HEADERS = [
        "ID правила",
        "Активно",
        "Категория",
        "Хэштеги",
        "Источник: Лист",
        "Источник: Колонка",
        "Цель: Лист",
        "Цель: Колонка",
        "Внешний документ",
        "Target Doc ID",
        "Дата создания",
        "Последнее изм.",
      ];
      // Добьем шапку при необходимости (в конец, порядок не ломаем)
      if (typeof Lib.ensureHeadersPresent_ === "function")
        Lib.ensureHeadersPresent_(sh, HEADERS);

      // Проставим валидацию «Да/Нет» на B и I (мягко, каждый раз безопасно)
      try {
        const ynRule = SpreadsheetApp.newDataValidation()
          .requireValueInList(["Да", "Нет"], true)
          .build();
        sh.getRange(
          2,
          2,
          Math.max(1, sh.getMaxRows() - 1),
          1
        ).setDataValidation(ynRule); // B
        sh.getRange(
          2,
          9,
          Math.max(1, sh.getMaxRows() - 1),
          1
        ).setDataValidation(ynRule); // I
      } catch (_) {}

      const now = new Date();
      const tz =
        (Lib.CONFIG && Lib.CONFIG.SETTINGS && Lib.CONFIG.SETTINGS.TIMEZONE) ||
        "UTC";

      // Попытка найти строку по ID
      let row = -1;
      if (id) {
        // используем быстрый поиск по колонке A
        row =
          typeof Lib.findRowByKey_ === "function"
            ? Lib.findRowByKey_(sh, id, true)
            : -1;
      }

      // Если ID нет или не найден — проверим, не существует ли уже дубль по ключевым полям
      if (row === -1 && sh.getLastRow() > 1) {
        const data = sh
          .getRange(2, 1, sh.getLastRow() - 1, HEADERS.length)
          .getValues();
        for (let i = 0; i < data.length; i++) {
          const r = data[i];
          const same =
            asStr(r[4]) === sourceSheet &&
            asStr(r[5]) === sourceHeader &&
            asStr(r[6]) === targetSheet &&
            asStr(r[7]) === targetHeader &&
            asStr(r[8]) === isExternalYN &&
            asStr(r[9]) === targetDocId;
          if (same) {
            row = i + 2;
            break;
          }
        }
      }

      // Если строка не найдена — создаём новую
      if (row === -1) {
        const newId =
          id ||
          "SR-" +
            Utilities.formatDate(now, tz, "yyyyMMdd-HHmmss") +
            "-" +
            (100 + Math.floor(Math.random() * 900));
        sh.appendRow([
          newId,
          enabledYN,
          category,
          hashtags,
          sourceSheet,
          sourceHeader,
          targetSheet,
          targetHeader,
          isExternalYN,
          targetDocId,
          now,
          now,
        ]);
        // Сброс кэша индексов по ID для этого листа
        if (typeof Lib.deleteKeyCacheForSheet === "function")
          Lib.deleteKeyCacheForSheet(sh.getName());
        // Обнулить кэш правил
        try {
          _cachedSyncRules = null;
        } catch (_) {
          /* no-op */
        }
        return { success: true, id: newId, action: "created" };
      }

      // Обновление существующей строки (сохраняем «Дата создания» из листа)
      const createdCell = sh.getRange(row, 11).getValue(); // K
      const created = createdCell ? createdCell : now;
      const keepId = sh.getRange(row, 1).getValue() || id;

      sh.getRange(row, 1, 1, HEADERS.length).setValues([
        [
          keepId,
          enabledYN,
          category,
          hashtags,
          sourceSheet,
          sourceHeader,
          targetSheet,
          targetHeader,
          isExternalYN,
          targetDocId,
          created,
          now,
        ],
      ]);

      try {
        _cachedSyncRules = null;
      } catch (_) {}
      return { success: true, id: keepId, action: "updated" };
    } catch (e) {
      return { success: false, error: e && e.message ? e.message : String(e) };
    } finally {
      try {
        lock.releaseLock();
      } catch (_) {}
    }
  };
})(Lib, this);

/** =============================================================================
 * D_ExternalDocsLogic.gs — Реестр внешних документов (для HTML-диалога)
 * =============================================================================
 */
var Lib = Lib || {};

(function (Lib, global) {
  const S = (Lib.CONFIG && Lib.CONFIG.SHEETS) || {};
  const SHEET_NAME = S.EXTERNAL_DOCS || "Внешние документы";

  // Эталонная шапка реестра
  const HEADERS = [
    "regId", // внутренний идентификатор записи (UUID)
    "name", // удобное имя
    "docId", // чистый ID таблицы
    "addedAt", // ISO-строка когда добавлено/обновлено
    "lastStatus", // человекочитаемый статус (для справки)
  ];

  /** ПУБЛИЧНО: вернуть список зарегистрированных документов для рендера в HTML */
  Lib.getExternalDocsList = function () {
    const sh = _ensureSheet_();
    const rng = sh.getDataRange();
    const values = rng.getValues();
    if (values.length <= 1) return [];

    const hdr = values[0].map((x) => String(x || "").trim());
    const rows = values.slice(1);

    const i = _indexOfHeaders_(hdr, HEADERS);
    const out = [];

    rows.forEach((r) => {
      const regId = String(r[i.regId] || "").trim();
      if (!regId) return;

      const name = r[i.name] || "";
      const docId = r[i.docId] || "";
      let status = "Проверка…";

      // Пробуем проверить доступ
      try {
        const ss = SpreadsheetApp.openById(String(docId));
        // если открылась — доступ есть
        status = "✅ доступ";
      } catch (e) {
        status = "❌ нет доступа";
      }

      out.push({ regId, name, docId, status });
    });

    // можно отсортировать по имени, чтобы было аккуратнее
    out.sort((a, b) =>
      String(a.name || "").localeCompare(String(b.name || ""), "ru")
    );

    return out;
  };

  /**
   * ПУБЛИЧНО: сохранить запись (create/update).
   * @param {{regId?: string|null, name: string, docIdOrUrl: string}} data
   * @returns {{success: boolean, message: string}}
   */
  Lib.saveExternalDoc = function (data) {
    try {
      if (
        !data ||
        !String(data.name || "").trim() ||
        !String(data.docIdOrUrl || "").trim()
      ) {
        return {
          success: false,
          message: "Не все обязательные поля заполнены.",
        };
      }

      const name = String(data.name).trim();
      const docId = _extractSpreadsheetId_(String(data.docIdOrUrl).trim());
      if (!docId) {
        return {
          success: false,
          message: "Не удалось извлечь ID из указанного значения.",
        };
      }

      // Проверяем доступ сразу — чтобы отобразить корректный статус
      let statusText = "❌ нет доступа";
      try {
        SpreadsheetApp.openById(docId);
        statusText = "✅ доступ";
      } catch (e) {
        // доступ отсутствует — не блокируем сохранение, но помечаем
        statusText = "❌ нет доступа";
      }

      const sh = _ensureSheet_();
      const lastRow = sh.getLastRow();
      const hdr = sh
        .getRange(1, 1, 1, sh.getLastColumn())
        .getValues()[0]
        .map((x) => String(x || "").trim());
      const idx = _indexOfHeaders_(hdr, HEADERS);

      const nowIso = new Date().toISOString();

      if (data.regId) {
        // UPDATE по regId
        const rowIdx = _findRowByValue_(sh, idx.regId + 1, String(data.regId));
        if (rowIdx > 1) {
          sh.getRange(rowIdx, idx.name + 1).setValue(name);
          sh.getRange(rowIdx, idx.docId + 1).setValue(docId);
          sh.getRange(rowIdx, idx.addedAt + 1).setValue(nowIso);
          sh.getRange(rowIdx, idx.lastStatus + 1).setValue(statusText);
          return { success: true, message: "Запись обновлена." };
        } else {
          // если не нашли по regId — считаем как новую
          const regId = Utilities.getUuid();
          sh.appendRow([regId, name, docId, nowIso, statusText]);
          return {
            success: true,
            message: "Запись создана (regId не найден, создана новая).",
          };
        }
      } else {
        // CREATE
        const regId = Utilities.getUuid();
        sh.appendRow([regId, name, docId, nowIso, statusText]);
        return { success: true, message: "Запись успешно сохранена." };
      }
    } catch (e) {
      Lib.logError && Lib.logError("saveExternalDoc: критическая ошибка", e);
      return { success: false, message: `Критическая ошибка: ${e.message}` };
    }
  };

  /**
   * ПУБЛИЧНО: удалить запись по regId
   * @param {string} regId
   * @returns {{success: boolean, message: string}}
   */
  Lib.deleteExternalDoc = function (regId) {
    try {
      if (!regId) return { success: false, message: "regId не передан." };

      const sh = _ensureSheet_();
      const hdr = sh
        .getRange(1, 1, 1, sh.getLastColumn())
        .getValues()[0]
        .map((x) => String(x || "").trim());
      const idx = _indexOfHeaders_(hdr, HEADERS);
      const rowIdx = _findRowByValue_(sh, idx.regId + 1, String(regId));

      if (rowIdx > 1) {
        sh.deleteRow(rowIdx);
        return { success: true, message: "Запись удалена." };
      } else {
        return { success: false, message: "Запись не найдена." };
      }
    } catch (e) {
      Lib.logError && Lib.logError("deleteExternalDoc: критическая ошибка", e);
      return { success: false, message: `Критическая ошибка: ${e.message}` };
    }
  };

  // ------------------------- helpers -------------------------

  /** Создать лист и шапку при необходимости */
  function _ensureSheet_() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) {
      sh = ss.insertSheet(SHEET_NAME);
    }
    // проверяем шапку
    const lastCol = sh.getLastColumn();
    const hasHeader = sh.getLastRow() >= 1 && lastCol >= HEADERS.length;
    if (!hasHeader) {
      sh.clear();
      sh.getRange(1, 1, 1, HEADERS.length)
        .setValues([HEADERS])
        .setFontWeight("bold");
      sh.setFrozenRows(1);
      sh.autoResizeColumns(1, HEADERS.length);
    } else {
      const hdr = sh
        .getRange(1, 1, 1, lastCol)
        .getValues()[0]
        .map((x) => String(x || "").trim());
      // если шапка отличается — перезапишем эталонной
      if (HEADERS.some((h, i) => hdr[i] !== h)) {
        sh.getRange(1, 1, 1, HEADERS.length)
          .setValues([HEADERS])
          .setFontWeight("bold");
        sh.setFrozenRows(1);
      }
    }
    return sh;
  }

  /** Быстрый индекс имён колонок -> индексы */
  function _indexOfHeaders_(hdr, names) {
    const m = {};
    names.forEach((h, i) => {
      const idx = hdr.indexOf(h);
      m[h] = idx;
      // также доступ по "ключу"
      if (h === "regId") m.regId = idx;
      if (h === "name") m.name = idx;
      if (h === "docId") m.docId = idx;
      if (h === "addedAt") m.addedAt = idx;
      if (h === "lastStatus") m.lastStatus = idx;
    });
    return m;
  }

  /** Найти строку по значению в колонке (возвращает индекс строки, 1-based; 0/1 — не найдено/шапка) */
  function _findRowByValue_(sh, colIndex1, value) {
    const lastRow = sh.getLastRow();
    if (lastRow <= 1) return -1;
    const rng = sh.getRange(2, colIndex1, lastRow - 1, 1).getValues();
    for (let i = 0; i < rng.length; i++) {
      if (String(rng[i][0]) === String(value)) return i + 2;
    }
    return -1;
  }

  /** Извлечь чистый spreadsheetId из URL или «голого» ID */
  function _extractSpreadsheetId_(input) {
    let s = String(input || "").trim();
    // Если это чистый UUID-подобный ID — просто вернуть
    if (/^[a-zA-Z0-9-_]{20,}$/.test(s)) return s;

    // Пытаемся вытащить из URL Google Sheets
    // Форматы: /spreadsheets/d/{ID}/..., ?id={ID}, /d/{ID}
    const m1 = s.match(/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    if (m1 && m1[1]) return m1[1];

    const m2 = s.match(/[?&]id=([a-zA-Z0-9-_]+)/);
    if (m2 && m2[1]) return m2[1];

    // на всякий случай — последний сегмент после /d/
    const m3 = s.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (m3 && m3[1]) return m3[1];

    return "";
  }

  // =======================================================================================
  // МОДУЛЬ: УПОРЯДОЧИВАНИЕ ЛИСТОВ
  // ---------------------------------------------------------------------------------------
  // Описание: Функция для автоматического упорядочивания листов в документе
  //           согласно заданному порядку из конфигурации.
  // =======================================================================================

  /**
   * Упорядочивает листы в документе согласно заданному порядку.
   * Листы, не указанные в порядке, остаются в конце без изменения позиций.
   * Вызывается автоматически при открытии документа (onOpen).
   */
  Lib.reorderSheets = function () {
    try {
      Lib.logDebug("Начало автоматического упорядочивания листов...");

      const ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) {
        Lib.logWarn("Не удалось получить активную таблицу для упорядочивания листов");
        return;
      }

      // Определяем порядок листов из конфигурации
      const orderedSheetNames = [
        Lib.CONFIG.SHEETS.PRIMARY,           // Главная
        Lib.CONFIG.SHEETS.ORDER_FORM,        // Заказ
        Lib.CONFIG.SHEETS.PRICE_DYNAMICS,    // Динамика цены
        Lib.CONFIG.SHEETS.PRICE_CALCULATION, // Расчёт цены
        Lib.CONFIG.SHEETS.PRICE,             // Прайс
        Lib.CONFIG.SHEETS.LABELS,            // Этикетки
        Lib.CONFIG.SHEETS.CERTIFICATION,     // Сертификация
        Lib.CONFIG.SHEETS.ABC_ANALYSIS,      // ABC-Анализ
        Lib.CONFIG.SHEETS.TZ_BY_STATUS,      // ТЗ по статусам
        Lib.CONFIG.SHEETS.ORDER_VERIFICATION,// Сверка заказа
        Lib.CONFIG.SHEETS.INVOICE_FULL,      // Для таможни
        Lib.CONFIG.SHEETS.NEWS,              // New sert
        Lib.CONFIG.SHEETS.FOR_DATABASE,      // Для базы
      ];

      const allSheets = ss.getSheets();
      Lib.logDebug(`Всего листов в документе: ${allSheets.length}`);

      // Создаем карту для быстрого поиска листов по имени
      const sheetMap = new Map();
      allSheets.forEach(sheet => {
        sheetMap.set(sheet.getName(), sheet);
      });

      // Перемещаем листы в заданном порядке
      let position = 0;
      for (const sheetName of orderedSheetNames) {
        const sheet = sheetMap.get(sheetName);
        if (sheet) {
          // Перемещаем лист на нужную позицию (позиции начинаются с 1 в Google Sheets API)
          ss.setActiveSheet(sheet);
          ss.moveActiveSheet(position + 1);
          Lib.logDebug(`Лист "${sheetName}" перемещен на позицию ${position + 1}`);
          position++;
        } else {
          Lib.logDebug(`Лист "${sheetName}" не найден в документе, пропускаем`);
        }
      }

      Lib.logDebug(`✅ Листы упорядочены успешно. Упорядочено листов: ${position}`);

      // Активируем лист "Главная" после упорядочивания
      // Активируем лист "Главная" после упорядочивания
      let mainSheetName = Lib.CONFIG.SHEETS.PRIMARY;
      let mainSheet = ss.getSheetByName(mainSheetName);

      if (!mainSheet && Lib.CONFIG.PRIMARY_DATA && Lib.CONFIG.PRIMARY_DATA.SHEETS && Lib.CONFIG.PRIMARY_DATA.SHEETS.TARGET) {
        // Fallback: пробуем найти лист по конфигурации проекта (например, "-Б/З поставщик" для MT)
        const projectMain = Lib.CONFIG.PRIMARY_DATA.SHEETS.TARGET.MAIN;
        if (projectMain) {
           mainSheet = ss.getSheetByName(projectMain);
           if (mainSheet) {
             mainSheetName = projectMain;
             Lib.logDebug(`Лист "Главная" не найден, используем проектный: "${mainSheetName}"`);
           }
        }
      }

      if (mainSheet) {
        ss.setActiveSheet(mainSheet);
        Lib.logDebug(`Лист "${mainSheetName}" активирован после упорядочивания`);
      } else {
        Lib.logWarn(`Лист "${Lib.CONFIG.SHEETS.PRIMARY}" (или аналог) не найден для активации`);
      }

    } catch (error) {
      Lib.logError("Ошибка при упорядочивании листов", error);
      // Не показываем alert при автоматическом упорядочивании, только логируем
    }
  };


})(Lib, this);
