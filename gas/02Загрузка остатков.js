var Lib = Lib || {};

/**
 * =======================================================================================
 * МОДУЛЬ: ЗАГРУЗКА ОСТАТКОВ (StocksLoader)
 * ---------------------------------------------------------------------------------------
 * Назначение: Универсальная загрузка остатков из листа "-остатки" для всех проектов
 *
 * Функции:
 * 1. Загрузка данных из листа "-остатки"
 * 2. Обновление столбцов на листе "Заказ":
 *    - ПРОДАЖИ, СПИСАНО, Остаток, товар в ПУТИ, РЕЗЕРВ
 *    - Остаток 1-3, СГ 1-3
 * 3. Автоматический пересчёт всех производных столбцов после загрузки
 * =======================================================================================
 */

(function (Lib, global) {
  /**
   * ПУБЛИЧНАЯ: Универсальная загрузка остатков для всех проектов
   * Вызывается из проектных файлов (SK, SS, MT и т.д.)
   */
  Lib.loadStockData = function (projectKey) {
    var ui = SpreadsheetApp.getUi();
    var config = global.CONFIG && global.CONFIG.PRIMARY_DATA;
    var menuTitle = (config && config.MENU && config.MENU.TITLE) || "🧾 Заказ";

    try {
      Lib.logInfo("[" + projectKey + "] Загрузка остатков: старт");

      var ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) {
        throw new Error("Не удалось получить активную таблицу.");
      }

      var orderSheetName =
        (global.CONFIG.SHEETS && global.CONFIG.SHEETS.ORDER_FORM) || "Заказ";
      var stocksSheetName =
        (global.CONFIG.SHEETS && global.CONFIG.SHEETS.STOCKS) || "-остатки";

      var orderSheet = ss.getSheetByName(orderSheetName);
      if (!orderSheet) {
        throw new Error('Не найден лист "' + orderSheetName + '".');
      }

      // Получаем лист остатков (может быть во внешнем документе)
      var stocksSheet = _getStocksSheet(ss, stocksSheetName, config, projectKey);

      // Проверяем наличие данных на листе Заказ
      var orderLastRow = orderSheet.getLastRow();
      if (orderLastRow <= 1) {
        ui.alert(
          menuTitle,
          'На листе "' + orderSheetName + '" нет строк для обновления.',
          ui.ButtonSet.OK
        );
        return;
      }

      // Получаем заголовки и данные листа Заказ
      var orderData = _getOrderData(orderSheet, orderSheetName, menuTitle, ui);
      if (!orderData) return;

      // ВАЖНО: Полностью очищаем столбцы перед загрузкой новых данных
      _clearStockColumns(orderSheet, orderData, projectKey);

      // Получаем данные листа остатков
      var stocksData = _getStocksData(stocksSheet, stocksSheetName, menuTitle, ui);
      if (!stocksData) return;

      // Обрабатываем данные остатков и создаём карту по артикулам
      var stocksMap = _buildStocksMap(stocksData);

      // Формируем значения для обновления листа Заказ
      var columnValues = _buildColumnValues(orderData, stocksMap);

      // Записываем значения на лист Заказ
      _writeValuesToOrderSheet(orderSheet, orderData, columnValues);

      // После загрузки остатков пересчитаем производные столбцы на листе Заказ
      try {
        if (typeof Lib.recalculateOrderSheet === 'function') {
          Lib.recalculateOrderSheet();
          Lib.logInfo('[' + projectKey + '] Производные столбцы пересчитаны после загрузки остатков');
        }
      } catch (e) {
        Lib.logWarn('[' + projectKey + '] Не удалось пересчитать производные столбцы заказа после загрузки остатков', e);
      }

      // Получаем обновлённые данные листа Заказ для проверки статусов
      var updatedOrderData = _getOrderData(orderSheet, orderSheetName, menuTitle, ui);
      if (!updatedOrderData) {
        Lib.logWarn('[' + projectKey + '] Не удалось получить обновлённые данные для проверки статусов');
      }

      // Проверяем статусы и обновляем "Снят в работу" -> "Снят архив" при необходимости
      try {
        if (updatedOrderData) {
          _checkAndUpdateArchivedStatus(ss, orderSheet, updatedOrderData, projectKey);
        }
      } catch (e) {
        Lib.logWarn('[' + projectKey + '] Не удалось обновить статусы после загрузки остатков', e);
      }

      Lib.logInfo("[" + projectKey + "] Загрузка остатков: обновлено строк " + orderData.numRows);
      ui.alert(
        menuTitle,
        "Загрузка остатков завершена. Обновлено строк: " + orderData.numRows + ".",
        ui.ButtonSet.OK
      );
    } catch (error) {
      Lib.logError("[" + projectKey + "] loadStockData: ошибка", error);
      ui.alert(
        "Ошибка загрузки остатков",
        error.message || String(error),
        ui.ButtonSet.OK
      );
    }
  };

  // =======================================================================================
  // ВНУТРЕННИЕ ФУНКЦИИ
  // =======================================================================================

  /**
   * Получение листа остатков (из текущего документа или внешнего)
   */
  function _getStocksSheet(ss, stocksSheetName, config, projectKey) {
    var stocksSheet = ss.getSheetByName(stocksSheetName);
    if (!stocksSheet) {
      var sourceSheetName = _resolveSheetName(config, "SOURCE", "STOCKS") || stocksSheetName;
      var sourceDocId = config && config.SOURCE && config.SOURCE.DOC_ID;

      if (!sourceDocId) {
        throw new Error(
          'Не найден лист "' +
            stocksSheetName +
            '" в активной таблице, и в конфигурации не указан документ источника.'
        );
      }

      var sourceSpreadsheet;
      if (ss.getId && ss.getId() === sourceDocId) {
        sourceSpreadsheet = ss;
      } else {
        sourceSpreadsheet = SpreadsheetApp.openById(sourceDocId);
      }

      stocksSheet = sourceSpreadsheet.getSheetByName(sourceSheetName);
      if (!stocksSheet) {
        throw new Error(
          'Не найден лист "' +
            sourceSheetName +
            '" в документе источнике остатков.'
        );
      }

      Lib.logInfo(
        '[' + projectKey + '] Загрузка остатков: данные читаются из внешнего документа ' +
          sourceSheetName
      );
    }
    return stocksSheet;
  }

  /**
   * Очистка столбцов остатков перед загрузкой новых данных
   * Очищает: Остаток, товар в ПУТИ, Остаток 1-3, СГ 1-3, СПИСАНО, РЕЗЕРВ, ПРОДАЖИ
   */
  function _clearStockColumns(orderSheet, orderData, projectKey) {
    Lib.logInfo('[' + projectKey + '] Начало очистки столбцов остатков');

    var startRow = orderData.startRow;
    var numRows = orderData.numRows;
    var idx = orderData.idx;

    // Создаём массив пустых значений для очистки
    var emptyValues = [];
    for (var i = 0; i < numRows; i++) {
      emptyValues.push(['']);
    }

    // Очищаем каждый столбец, если он существует
    var columnsToClear = [
      { name: 'ПРОДАЖИ', index: idx.sales },
      { name: 'СПИСАНО', index: idx.writtenOff },
      { name: 'Остаток', index: idx.stock },
      { name: 'товар в ПУТИ', index: idx.inTransit },
      { name: 'РЕЗЕРВ', index: idx.reserve },
      { name: 'Остаток 1', index: idx.qty1 },
      { name: 'СГ 1', index: idx.exp1 },
      { name: 'Остаток 2', index: idx.qty2 },
      { name: 'СГ 2', index: idx.exp2 },
      { name: 'Остаток3', index: idx.qty3 },
      { name: 'СГ 3', index: idx.exp3 }
    ];

    var clearedCount = 0;
    for (var i = 0; i < columnsToClear.length; i++) {
      var col = columnsToClear[i];
      if (col.index !== -1) {
        try {
          orderSheet.getRange(startRow, col.index + 1, numRows, 1).setValues(emptyValues);
          clearedCount++;
        } catch (e) {
          Lib.logWarn('[' + projectKey + '] Не удалось очистить столбец "' + col.name + '"', e);
        }
      }
    }

    Lib.logInfo('[' + projectKey + '] Очищено столбцов: ' + clearedCount + ' из ' + columnsToClear.length);
  }

  /**
   * Получение данных листа Заказ
   */
  function _getOrderData(orderSheet, orderSheetName, menuTitle, ui) {
    var orderLastRow = orderSheet.getLastRow();
    var orderLastColumn = orderSheet.getLastColumn();

    var orderHeaderInfo = _findHeaderRow(
      orderSheet,
      orderLastColumn,
      ["Арт. Рус"]
    );
    if (!orderHeaderInfo) {
      throw new Error(
        'Не удалось определить строку заголовков на листе "' +
          orderSheetName +
          '" — не найден столбец "Арт. Рус".'
      );
    }

    var orderHeaderRow = orderHeaderInfo.row;
    var orderHeaders = orderHeaderInfo.headers;

    if (orderLastRow <= orderHeaderRow) {
      ui.alert(
        menuTitle,
        'На листе "' + orderSheetName + '" нет строк для обновления.',
        ui.ButtonSet.OK
      );
      return null;
    }

    var orderIdx = {
      article: orderHeaders.indexOf("Арт. Рус"),
      sales: orderHeaders.indexOf("ПРОДАЖИ"),
      writtenOff: orderHeaders.indexOf("СПИСАНО"),
      stock: orderHeaders.indexOf("Остаток"),
      inTransit: orderHeaders.indexOf("товар в  ПУТИ"),
      reserve: orderHeaders.indexOf("РЕЗЕРВ"),
      qty1: orderHeaders.indexOf("Остаток 1"),
      exp1: orderHeaders.indexOf("СГ 1"),
      qty2: orderHeaders.indexOf("Остаток 2"),
      exp2: orderHeaders.indexOf("СГ 2"),
      qty3: orderHeaders.indexOf("Остаток3"),
      exp3: orderHeaders.indexOf("СГ 3"),
    };
    if (orderIdx.qty3 === -1) {
      orderIdx.qty3 = orderHeaders.indexOf("Остаток 3");
    }

    if (orderIdx.article === -1) {
      throw new Error(
        'На листе "' + orderSheetName + '" отсутствует столбец "Арт. Рус".'
      );
    }

    var orderDataStartRow = orderHeaderRow + 1;
    var orderDataRowCount = orderLastRow - orderHeaderRow;
    if (orderDataRowCount <= 0) {
      ui.alert(
        menuTitle,
        'На листе "' + orderSheetName + '" нет строк для обновления.',
        ui.ButtonSet.OK
      );
      return null;
    }

    var orderDataRange = orderSheet.getRange(
      orderDataStartRow,
      1,
      orderDataRowCount,
      orderLastColumn
    );
    var orderValues = orderDataRange.getValues();

    return {
      startRow: orderDataStartRow,
      numRows: orderValues.length,
      values: orderValues,
      idx: orderIdx,
    };
  }

  /**
   * Получение данных листа остатков
   * ВАЖНО: Заголовки на листе "-остатки" всегда находятся в строке 3, столбец "Артикул" всегда в колонке C
   */
  function _getStocksData(stocksSheet, stocksSheetName, menuTitle, ui) {
    var stocksLastRow = stocksSheet.getLastRow();
    if (stocksLastRow <= 1) {
      ui.alert(
        menuTitle,
        'На листе "' + stocksSheetName + '" нет данных.',
        ui.ButtonSet.OK
      );
      return null;
    }

    var stocksLastColumn = stocksSheet.getLastColumn();

    // ИСПРАВЛЕНО: Для листа остатков заголовки ВСЕГДА в строке 3
    var STOCKS_HEADER_ROW = 3;

    // Логируем информацию о листе для отладки
    var sheetInfo = 'имя: "' + stocksSheet.getName() + '", строк: ' + stocksLastRow + ', колонок: ' + stocksLastColumn;
    try {
      var parentSpreadsheet = stocksSheet.getParent();
      if (parentSpreadsheet) {
        sheetInfo += ', документ: "' + parentSpreadsheet.getName() + '" (ID: ' + parentSpreadsheet.getId() + ')';
      }
    } catch (e) {
      sheetInfo += ', не удалось получить информацию о документе';
    }
    Lib.logInfo('[StocksLoader] Информация о листе остатков: ' + sheetInfo);

    Lib.logInfo('[StocksLoader] Проверка заголовков листа остатков в строке ' + STOCKS_HEADER_ROW);

    // Проверяем, что строка 3 существует
    if (stocksLastRow < STOCKS_HEADER_ROW) {
      ui.alert(
        menuTitle,
        'На листе "' + stocksSheetName + '" недостаточно строк. Ожидается строка заголовков в строке ' + STOCKS_HEADER_ROW + '.',
        ui.ButtonSet.OK
      );
      return null;
    }

    // Для отладки: читаем первые 5 строк листа
    try {
      var debugRowCount = Math.min(5, stocksLastRow);
      var debugColCount = Math.min(10, stocksLastColumn);
      Lib.logDebug('[StocksLoader] Чтение первых ' + debugRowCount + ' строк для анализа структуры листа:');
      for (var dr = 1; dr <= debugRowCount; dr++) {
        var debugRow = stocksSheet.getRange(dr, 1, 1, debugColCount).getValues()[0];
        var debugRowPreview = [];
        for (var dc = 0; dc < debugRow.length; dc++) {
          var val = debugRow[dc];
          if (val instanceof Date) {
            debugRowPreview.push('Date(' + val.toISOString().split('T')[0] + ')');
          } else if (typeof val === 'string' && val.length > 30) {
            debugRowPreview.push('"' + val.substring(0, 27) + '..."');
          } else {
            debugRowPreview.push(JSON.stringify(val));
          }
        }
        Lib.logDebug('[StocksLoader]   Строка ' + dr + ': [' + debugRowPreview.join(', ') + ']');
      }
    } catch (debugError) {
      Lib.logWarn('[StocksLoader] Не удалось прочитать первые строки для отладки: ' + debugError);
    }

    // Читаем строку 3 как заголовки
    var stocksHeaders = stocksSheet.getRange(STOCKS_HEADER_ROW, 1, 1, stocksLastColumn).getValues()[0];

    // Логируем первые 10 заголовков для отладки
    var headersPreview = [];
    for (var h = 0; h < Math.min(10, stocksHeaders.length); h++) {
      headersPreview.push(stocksHeaders[h]);
    }
    Lib.logInfo('[StocksLoader] Заголовки в строке ' + STOCKS_HEADER_ROW + ': ' + JSON.stringify(headersPreview));

    // Проверяем наличие столбца "Артикул" в колонке C (индекс 2)
    var articleColumnC = String(stocksHeaders[2] || '').trim();
    Lib.logInfo('[StocksLoader] Столбец C (индекс 2): "' + articleColumnC + '"');

    // Определяем индексы столбцов
    var articleIdx = _findColumnIndex(stocksHeaders, "Артикул");
    if (articleIdx === -1) {
      articleIdx = _findColumnIndex(stocksHeaders, "Арт. Рус");
    }

    // Если не нашли через поиск, но в колонке C есть "Артикул", используем индекс 2
    if (articleIdx === -1 && articleColumnC.toLowerCase().indexOf('артикул') !== -1) {
      articleIdx = 2;
      Lib.logInfo('[StocksLoader] Столбец "Артикул" найден в колонке C (индекс 2) через прямую проверку');
    }

    var stocksIdx = {
      article: articleIdx,
      sold: _findColumnIndex(stocksHeaders, "Продано за период"),
      writtenOff: _findColumnIndex(stocksHeaders, "Списано за период"),
      stock: _findColumnIndex(stocksHeaders, "Остаток на текущий день"),
      reserve: _findColumnIndex(stocksHeaders, "Из них в резерве"),
      total: _findColumnIndex(stocksHeaders, "Всего"),
      expiry: _findColumnIndex(stocksHeaders, "Срок годности"),
    };

    Lib.logInfo('[StocksLoader] Индексы столбцов остатков: ' + JSON.stringify(stocksIdx));

    if (stocksIdx.article === -1) {
      Lib.logError('[StocksLoader] Все заголовки листа: ' + JSON.stringify(stocksHeaders));
      throw new Error('На листе "' + stocksSheetName + '" отсутствует столбец "Артикул" в строке ' + STOCKS_HEADER_ROW + '. Проверьте структуру листа.');
    }

    // Проверяем наличие данных после заголовков
    if (stocksLastRow <= STOCKS_HEADER_ROW) {
      ui.alert(
        menuTitle,
        'На листе "' + stocksSheetName + '" нет данных после строки заголовков (строка ' + STOCKS_HEADER_ROW + ').',
        ui.ButtonSet.OK
      );
      return null;
    }

    // Читаем данные начиная со строки 4 (после заголовков в строке 3)
    var stocksDataStartRow = STOCKS_HEADER_ROW + 1;
    var stocksDataRowCount = stocksLastRow - STOCKS_HEADER_ROW;

    Lib.logInfo('[StocksLoader] Чтение данных: строки ' + stocksDataStartRow + '-' + stocksLastRow + ' (' + stocksDataRowCount + ' строк данных)');

    var stocksDataRange = stocksSheet.getRange(
      stocksDataStartRow,
      1,
      stocksDataRowCount,
      stocksLastColumn
    );
    var stocksValues = stocksDataRange.getValues();

    // Логируем первые 3 строки данных для проверки
    var dataPreview = Math.min(3, stocksValues.length);
    for (var d = 0; d < dataPreview; d++) {
      var rowPreview = [];
      for (var c = 0; c < Math.min(10, stocksValues[d].length); c++) {
        rowPreview.push(stocksValues[d][c]);
      }
      Lib.logDebug('[StocksLoader] Данные строка ' + (stocksDataStartRow + d) + ': ' + JSON.stringify(rowPreview));
    }

    return {
      values: stocksValues,
      idx: stocksIdx,
    };
  }

  /**
   * Построение карты остатков по артикулам
   */
  function _buildStocksMap(stocksData) {
    var stocksMap = Object.create(null);
    var currentArticle = "";
    var currentSold = null;
    var currentWrittenOff = null;
    var currentStock = null;

    stocksData.values.forEach(function (row) {
      var articleCell =
        stocksData.idx.article !== -1
          ? _asTrimmedString(row[stocksData.idx.article])
          : "";
      if (articleCell) {
        currentArticle = articleCell;
        currentSold = null;
        currentWrittenOff = null;
        currentStock = null;
      }

      if (!currentArticle) {
        return;
      }

      if (stocksData.idx.sold !== -1) {
        var soldCell = row[stocksData.idx.sold];
        if (soldCell !== "" && soldCell !== null) {
          currentSold = _parseNumber(soldCell);
        }
      }
      if (stocksData.idx.writtenOff !== -1) {
        var writtenCell = row[stocksData.idx.writtenOff];
        if (writtenCell !== "" && writtenCell !== null) {
          currentWrittenOff = _parseNumber(writtenCell);
        }
      }
      if (stocksData.idx.stock !== -1) {
        var stockCell = row[stocksData.idx.stock];
        if (stockCell !== "" && stockCell !== null) {
          currentStock = _parseNumber(stockCell);
        }
      }

      var entry = stocksMap[currentArticle];
      if (!entry) {
        entry = {
          sales: currentSold,
          writtenOff: currentWrittenOff,
          stock: currentStock,
          reserve: 0,
          batches: Object.create(null),
          batchesMeta: Object.create(null),
        };
        stocksMap[currentArticle] = entry;
      } else {
        if (currentSold !== null && !isNaN(currentSold))
          entry.sales = currentSold;
        if (currentWrittenOff !== null && !isNaN(currentWrittenOff))
          entry.writtenOff = currentWrittenOff;
        if (currentStock !== null && !isNaN(currentStock))
          entry.stock = currentStock;
      }

      if (stocksData.idx.reserve !== -1) {
        var reserveCell = row[stocksData.idx.reserve];
        var reserveValue = _parseNumber(reserveCell);
        if (reserveValue !== null && !isNaN(reserveValue)) {
          entry.reserve += reserveValue;
        }
      }

      if (stocksData.idx.total !== -1) {
        var totalCell = row[stocksData.idx.total];
        var totalValue = _parseNumber(totalCell);
        if (totalValue !== null && !isNaN(totalValue) && totalValue !== 0) {
          var expiryCell =
            stocksData.idx.expiry !== -1 ? row[stocksData.idx.expiry] : "";
          var expiryInfo = _normalizeExpiry(expiryCell);
          var key = expiryInfo.label;
          if (!entry.batches[key]) {
            entry.batches[key] = 0;
            entry.batchesMeta[key] = expiryInfo.date;
          }
          entry.batches[key] += totalValue;
        }
      }
    });

    return stocksMap;
  }

  /**
   * Формирование значений для обновления листа Заказ
   */
  function _buildColumnValues(orderData, stocksMap) {
    var salesColumnValues = [];
    var writtenOffColumnValues = [];
    var stockColumnValues = [];
    var inTransitColumnValues = [];
    var reserveColumnValues = [];
    var qty1ColumnValues = [];
    var exp1ColumnValues = [];
    var qty2ColumnValues = [];
    var exp2ColumnValues = [];
    var qty3ColumnValues = [];
    var exp3ColumnValues = [];

    for (var i = 0; i < orderData.numRows; i++) {
      var orderRow = orderData.values[i];
      var article = _asTrimmedString(orderRow[orderData.idx.article]);
      var stats = article ? stocksMap[article] : null;

      if (stats) {
        salesColumnValues.push([
          stats.sales !== null && !isNaN(stats.sales) ? stats.sales : "",
        ]);
        writtenOffColumnValues.push([
          stats.writtenOff !== null && !isNaN(stats.writtenOff)
            ? stats.writtenOff
            : "",
        ]);
        stockColumnValues.push([
          stats.stock !== null && !isNaN(stats.stock) ? stats.stock : "",
        ]);
        inTransitColumnValues.push([""]); // товар в пути - пустое значение (загружается отдельно)
        reserveColumnValues.push([
          stats.reserve !== null && !isNaN(stats.reserve)
            ? stats.reserve
            : "",
        ]);

        var batchEntries = [];
        for (var batchKey in stats.batches) {
          if (!Object.prototype.hasOwnProperty.call(stats.batches, batchKey))
            continue;
          var qty = stats.batches[batchKey];
          if (qty === null || isNaN(qty) || qty === 0) continue;
          var dateObj = stats.batchesMeta[batchKey];
          batchEntries.push({
            label: batchKey,
            qty: qty,
            dateObj: dateObj,
            sortValue: dateObj ? dateObj.getTime() : null,
          });
        }

        batchEntries.sort(function (a, b) {
          var aSort = a.sortValue;
          var bSort = b.sortValue;
          if (aSort !== null && bSort !== null) {
            return bSort - aSort;
          }
          if (aSort !== null) return -1;
          if (bSort !== null) return 1;
          return String(b.label).localeCompare(String(a.label));
        });

        var batchesForWrite = [
          batchEntries[0] || null,
          batchEntries[1] || null,
          batchEntries[2] || null,
        ];

        qty1ColumnValues.push([
          batchesForWrite[0] ? batchesForWrite[0].qty : "",
        ]);
        exp1ColumnValues.push([
          batchesForWrite[0]
            ? batchesForWrite[0].dateObj || batchesForWrite[0].label
            : "",
        ]);
        qty2ColumnValues.push([
          batchesForWrite[1] ? batchesForWrite[1].qty : "",
        ]);
        exp2ColumnValues.push([
          batchesForWrite[1]
            ? batchesForWrite[1].dateObj || batchesForWrite[1].label
            : "",
        ]);
        qty3ColumnValues.push([
          batchesForWrite[2] ? batchesForWrite[2].qty : "",
        ]);
        exp3ColumnValues.push([
          batchesForWrite[2]
            ? batchesForWrite[2].dateObj || batchesForWrite[2].label
            : "",
        ]);
      } else {
        salesColumnValues.push([""]);
        writtenOffColumnValues.push([""]);
        stockColumnValues.push([""]);
        inTransitColumnValues.push([""]);
        reserveColumnValues.push([""]);
        qty1ColumnValues.push([""]);
        exp1ColumnValues.push([""]);
        qty2ColumnValues.push([""]);
        exp2ColumnValues.push([""]);
        qty3ColumnValues.push([""]);
        exp3ColumnValues.push([""]);
      }
    }

    return {
      sales: salesColumnValues,
      writtenOff: writtenOffColumnValues,
      stock: stockColumnValues,
      inTransit: inTransitColumnValues,
      reserve: reserveColumnValues,
      qty1: qty1ColumnValues,
      exp1: exp1ColumnValues,
      qty2: qty2ColumnValues,
      exp2: exp2ColumnValues,
      qty3: qty3ColumnValues,
      exp3: exp3ColumnValues,
    };
  }

  /**
   * Запись значений на лист Заказ
   */
  function _writeValuesToOrderSheet(orderSheet, orderData, columnValues) {
    var startRow = orderData.startRow;
    var numRows = orderData.numRows;
    var idx = orderData.idx;

    if (idx.sales !== -1) {
      orderSheet
        .getRange(startRow, idx.sales + 1, numRows, 1)
        .setValues(columnValues.sales);
    }
    if (idx.writtenOff !== -1) {
      orderSheet
        .getRange(startRow, idx.writtenOff + 1, numRows, 1)
        .setValues(columnValues.writtenOff);
    }
    if (idx.stock !== -1) {
      orderSheet
        .getRange(startRow, idx.stock + 1, numRows, 1)
        .setValues(columnValues.stock);
    }
    if (idx.inTransit !== -1) {
      orderSheet
        .getRange(startRow, idx.inTransit + 1, numRows, 1)
        .setValues(columnValues.inTransit);
    }
    if (idx.reserve !== -1) {
      orderSheet
        .getRange(startRow, idx.reserve + 1, numRows, 1)
        .setValues(columnValues.reserve);
    }
    if (idx.qty1 !== -1) {
      orderSheet
        .getRange(startRow, idx.qty1 + 1, numRows, 1)
        .setValues(columnValues.qty1);
    }
    if (idx.exp1 !== -1) {
      orderSheet
        .getRange(startRow, idx.exp1 + 1, numRows, 1)
        .setValues(columnValues.exp1);
    }
    if (idx.qty2 !== -1) {
      orderSheet
        .getRange(startRow, idx.qty2 + 1, numRows, 1)
        .setValues(columnValues.qty2);
    }
    if (idx.exp2 !== -1) {
      orderSheet
        .getRange(startRow, idx.exp2 + 1, numRows, 1)
        .setValues(columnValues.exp2);
    }
    if (idx.qty3 !== -1) {
      orderSheet
        .getRange(startRow, idx.qty3 + 1, numRows, 1)
        .setValues(columnValues.qty3);
    }
    if (idx.exp3 !== -1) {
      orderSheet
        .getRange(startRow, idx.exp3 + 1, numRows, 1)
        .setValues(columnValues.exp3);
    }
  }

  // =======================================================================================
  // ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
  // =======================================================================================

  /**
   * Поиск индекса столбца с нормализацией (игнорирует регистр и лишние пробелы)
   */
  function _findColumnIndex(headers, columnName) {
    var normalizedTarget = String(columnName || "").trim().toLowerCase().replace(/\s+/g, " ");

    for (var i = 0; i < headers.length; i++) {
      var normalizedHeader = String(headers[i] || "").trim().toLowerCase().replace(/\s+/g, " ");
      if (normalizedHeader === normalizedTarget) {
        return i;
      }
    }

    return -1;
  }

  /**
   * Поиск строки заголовков на листе (используется только для листа "Заказ")
   * Для листа "-остатки" заголовки всегда в строке 3, поэтому эта функция там не используется
   */
  function _findHeaderRow(sheet, lastColumn, requiredHeaders) {
    if (lastColumn <= 0) return null;

    // Нормализуем required headers для сравнения
    var normalizedRequired = requiredHeaders.map(function(h) {
      return String(h || "").trim().toLowerCase().replace(/\s+/g, " ");
    });

    var maxRowsToScan = Math.min(20, sheet.getLastRow());
    Lib.logDebug('[StocksLoader] _findHeaderRow: сканирование строк 1-' + maxRowsToScan + ', ищем: ' + JSON.stringify(requiredHeaders));

    for (var row = 1; row <= maxRowsToScan; row++) {
      var rowValues = sheet.getRange(row, 1, 1, lastColumn).getValues()[0];

      // Проверяем, похожа ли строка на заголовок (содержит текст, а не только числа)
      var numericCells = 0;
      var textCells = 0;

      for (var k = 0; k < Math.min(5, rowValues.length); k++) {
        var val = rowValues[k];
        if (val !== null && val !== undefined && String(val).trim() !== '') {
          if (typeof val === 'number' || !isNaN(val)) {
            numericCells++;
          } else {
            textCells++;
          }
        }
      }

      // Если первые ячейки в основном числовые, вероятно это данные, а не заголовки
      if (numericCells > 0 && textCells === 0) {
        Lib.logDebug('[StocksLoader] Строка ' + row + ' пропущена (похоже на данные, не заголовок): ' + JSON.stringify(rowValues.slice(0, 5)));
        continue;
      }

      var found = true;

      for (var i = 0; i < normalizedRequired.length; i++) {
        var required = normalizedRequired[i];
        var foundHeader = false;

        for (var j = 0; j < rowValues.length; j++) {
          var cellValue = String(rowValues[j] || "").trim();
          var normalizedCell = cellValue.toLowerCase().replace(/\s+/g, " ");

          if (normalizedCell === required) {
            foundHeader = true;
            break;
          }
        }

        if (!foundHeader) {
          found = false;
          break;
        }
      }

      if (found) {
        Lib.logInfo('[StocksLoader] Найдена строка заголовков: строка ' + row + ', первые 5 ячеек: ' + JSON.stringify(rowValues.slice(0, 5)));
        return {
          row: row,
          headers: rowValues.map(function (h) {
            return String(h || "").trim();
          }),
        };
      }
    }

    Lib.logWarn('[StocksLoader] Не найдена строка заголовков в первых ' + maxRowsToScan + ' строках');
    return null;
  }

  function _resolveSheetName(config, section, key) {
    if (!config || !config.SHEETS || !config.SHEETS[section]) {
      return null;
    }
    return config.SHEETS[section][key] || null;
  }

  function _asTrimmedString(value) {
    if (value === null || value === undefined) {
      return "";
    }
    return String(value).trim();
  }

  function _parseNumber(value) {
    if (value === null || value === undefined) {
      return null;
    }
    if (typeof value === "number") {
      return isNaN(value) ? null : value;
    }
    var str = String(value).trim();
    if (!str) return null;
    str = str.replace(/\s+/g, "").replace(",", ".");
    var num = parseFloat(str);
    return isNaN(num) ? null : num;
  }

  function _normalizeExpiry(value) {
    if (!value || value === "") {
      return { label: "Без даты", date: null };
    }

    if (value instanceof Date) {
      var year = value.getFullYear();
      var month = String(value.getMonth() + 1).padStart(2, "0");
      return {
        label: month + "/" + year,
        date: value,
      };
    }

    var str = String(value).trim();
    if (!str) {
      return { label: "Без даты", date: null };
    }

    var dateMatch = str.match(/(\d{1,2})[\/\.\-](\d{4})/);
    if (dateMatch) {
      var m = parseInt(dateMatch[1], 10);
      var y = parseInt(dateMatch[2], 10);
      if (m >= 1 && m <= 12 && y >= 1900 && y <= 2100) {
        var dateObj = new Date(y, m - 1, 1);
        return {
          label: String(m).padStart(2, "0") + "/" + y,
          date: dateObj,
        };
      }
    }

    return { label: str, date: null };
  }

  /**
   * Проверка и обновление статусов после загрузки остатков
   * Меняет статус "Снят в работу" на "Снят архив", если Остаток и ПРОДАЖИ пусты или 0
   */
  function _checkAndUpdateArchivedStatus(ss, orderSheet, orderData, projectKey) {
    Lib.logInfo('[' + projectKey + '] Начало проверки статусов для архивации');

    // Получаем индексы необходимых столбцов
    var statusIndex = orderData.idx.article; // Пока используем article как reference

    // Нам нужно получить индекс столбца "Статус"
    var orderLastColumn = orderSheet.getLastColumn();
    var orderHeaderInfo = _findHeaderRow(orderSheet, orderLastColumn, ["Арт. Рус"]);

    if (!orderHeaderInfo) {
      Lib.logError('[' + projectKey + '] Не удалось определить строку заголовков');
      return;
    }

    var orderHeaders = orderHeaderInfo.headers;
    statusIndex = _findColumnIndex(orderHeaders, "Статус");

    if (statusIndex === -1) {
      Lib.logInfo('[' + projectKey + '] Столбец "Статус" не найден, пропускаем проверку');
      return;
    }

    var salesIndex = orderData.idx.sales;
    var stockIndex = orderData.idx.stock;

    if (salesIndex === -1 || stockIndex === -1) {
      Lib.logInfo('[' + projectKey + '] Столбцы "ПРОДАЖИ" или "Остаток" не найдены, пропускаем проверку');
      return;
    }

    var archivedCount = 0;
    var statusUpdates = [];
    var rowsToSync = []; // Индексы строк для синхронизации с ТЗ

    // Проходим по всем строкам
    for (var i = 0; i < orderData.numRows; i++) {
      var orderRow = orderData.values[i];
      var currentStatus = _asTrimmedString(orderRow[statusIndex]);

      // Проверяем, что статус = "Снят в работу"
      if (currentStatus === 'Снят в работу') {
        var salesValue = orderRow[salesIndex];
        var stockValue = orderRow[stockIndex];

        // Проверяем, что продажи и остатки пусты или равны 0
        var isSalesEmpty = salesValue === '' || salesValue === null || salesValue === undefined || salesValue === 0;
        var isStockEmpty = stockValue === '' || stockValue === null || stockValue === undefined || stockValue === 0;

        if (isSalesEmpty && isStockEmpty) {
          // Меняем статус на "Снят архив"
          orderRow[statusIndex] = 'Снят архив';
          statusUpdates.push({
            rowNumber: orderData.startRow + i,
            columnNumber: statusIndex + 1
          });
          rowsToSync.push(i); // Запоминаем индекс строки для синхронизации
          archivedCount++;
        }
      }
    }

    // Записываем обновлённые статусы на лист
    if (statusUpdates.length > 0) {
      for (var j = 0; j < statusUpdates.length; j++) {
        var update = statusUpdates[j];
        orderSheet.getRange(update.rowNumber, update.columnNumber).setValue('Снят архив');
      }
      Lib.logInfo('[' + projectKey + '] Обновлено статусов на "Снят архив": ' + archivedCount);

      // ВАЖНО: Прямая синхронизация статусов через правила
      // Этап 1: Заказ → все целевые листы (например, Главная)
      // Этап 2: Главная → все целевые листы (например, Сертификация, Этикетки)
      try {
        _syncStatusDirectly(ss, orderSheet, statusUpdates, 'Снят архив', projectKey);
      } catch (e) {
        Lib.logWarn('[' + projectKey + '] Не удалось выполнить прямую синхронизацию статусов', e);
      }

      // Синхронизируем с листом "ТЗ по статусам"
      try {
        _syncArchivedToTaskSheet(ss, orderData, orderHeaders, rowsToSync, projectKey);
      } catch (e) {
        Lib.logWarn('[' + projectKey + '] Не удалось синхронизировать архивные статусы с ТЗ', e);
      }
    } else {
      Lib.logInfo('[' + projectKey + '] Нет строк для архивации');
    }
  }

  /**
   * Прямая синхронизация статусов через правила синхронизации
   * Ищет все правила с targetHeader = "Статус" и применяет их напрямую
   */
  function _syncStatusDirectly(ss, orderSheet, statusUpdates, newStatusValue, projectKey) {
    if (!statusUpdates || statusUpdates.length === 0) return;

    Lib.logInfo('[' + projectKey + '] Начало прямой синхронизации статусов');

    try {
      // Получаем правила синхронизации с сервера (лист "Правила синхро" не используется)
      if (typeof Lib.callServer !== 'function') {
        Lib.logWarn('[' + projectKey + '] Lib.callServer недоступен, правила не загружены');
        return;
      }

      var ssId = ss.getId();
      var rulesRes = Lib.callServer('/rules/' + ssId, { force_reload: true, include_disabled: false }, 'GET');
      var rulesData = (rulesRes && rulesRes.rules) || [];

      var orderSheetName = orderSheet.getName();
      var applicableRules = [];

      // Находим правила для синхронизации статуса
      for (var i = 0; i < rulesData.length; i++) {
        var rule = rulesData[i] || {};
        var enabled = rule.enabled === true || String(rule.enabled || '').trim().toLowerCase() === 'да';
        if (!enabled) continue;

        var mode = String(rule.mode || 'unidirectional').trim();

        if (mode === 'bidirectional') {
          var sheetA = String(rule.sheet_a || rule.sheetA || '').trim();
          var headerA = String(rule.header_a || rule.headerA || '').trim();
          var sheetB = String(rule.sheet_b || rule.sheetB || '').trim();
          var headerB = String(rule.header_b || rule.headerB || '').trim();

          if (sheetA === orderSheetName && headerA === 'Статус') {
            applicableRules.push({ targetSheet: sheetB, targetHeader: headerB, isExternal: false, targetDocId: '' });
          } else if (sheetB === orderSheetName && headerB === 'Статус') {
            applicableRules.push({ targetSheet: sheetA, targetHeader: headerA, isExternal: false, targetDocId: '' });
          }
          continue;
        }

        var sourceSheet = String(rule.source_sheet || rule.sourceSheet || '').trim();
        var sourceHeader = String(rule.source_header || rule.sourceHeader || '').trim();
        var targetSheet = String(rule.target_sheet || rule.targetSheet || '').trim();
        var targetHeader = String(rule.target_header || rule.targetHeader || '').trim();
        var isExternal = !!(rule.is_external ?? rule.isExternal);
        var targetDocId = String(rule.target_doc_id || rule.targetDocId || '').trim();

        // Фильтруем правила: активные, источник = наш лист, колонка = Статус
        if (sourceSheet === orderSheetName && sourceHeader === 'Статус' && targetSheet && targetHeader === 'Статус') {
          applicableRules.push({
            targetSheet: targetSheet,
            targetHeader: targetHeader,
            isExternal: isExternal,
            targetDocId: targetDocId
          });
        }
      }

      if (applicableRules.length === 0) {
        Lib.logInfo('[' + projectKey + '] Нет применимых правил для синхронизации статуса');
        return;
      }

      Lib.logInfo('[' + projectKey + '] Найдено применимых правил: ' + applicableRules.length);

      // Собираем информацию о обновлённых листах для каскадной синхронизации
      var updatedSheets = {}; // { "Главная": [список ID] }

      // ЭТАП 1: Для каждого правила применяем изменения
      for (var i = 0; i < applicableRules.length; i++) {
        var rule = applicableRules[i];

        try {
          var updatedIds = _applyStatusToTargetSheet(ss, orderSheet, statusUpdates, newStatusValue, rule, projectKey);

          // Запоминаем обновлённые ID для каждого целевого листа
          if (updatedIds && updatedIds.length > 0) {
            if (!updatedSheets[rule.targetSheet]) {
              updatedSheets[rule.targetSheet] = [];
            }
            for (var j = 0; j < updatedIds.length; j++) {
              if (updatedSheets[rule.targetSheet].indexOf(updatedIds[j]) === -1) {
                updatedSheets[rule.targetSheet].push(updatedIds[j]);
              }
            }
          }
        } catch (e) {
          Lib.logWarn('[' + projectKey + '] Ошибка применения правила для листа "' + rule.targetSheet + '"', e);
        }
      }

      // ЭТАП 2: Каскадная синхронизация - обновляем листы, которые зависят от только что обновлённых
      try {
        _cascadeStatusSync(ss, updatedSheets, newStatusValue, rulesData, projectKey);
      } catch (e) {
        Lib.logWarn('[' + projectKey + '] Ошибка каскадной синхронизации', e);
      }

      Lib.logInfo('[' + projectKey + '] Прямая синхронизация статусов завершена');
    } catch (e) {
      Lib.logError('[' + projectKey + '] _syncStatusDirectly: критическая ошибка', e);
    }
  }

  /**
   * Применяет новый статус на целевой лист по правилу
   * Возвращает массив обновлённых ID
   */
  function _applyStatusToTargetSheet(ss, orderSheet, statusUpdates, newStatusValue, rule, projectKey) {
    var updatedIds = [];

    // Получаем целевую таблицу
    var targetSS = rule.isExternal && rule.targetDocId
      ? SpreadsheetApp.openById(rule.targetDocId)
      : ss;

    var targetSheet = targetSS.getSheetByName(rule.targetSheet);
    if (!targetSheet) {
      Lib.logWarn('[' + projectKey + '] Целевой лист "' + rule.targetSheet + '" не найден');
      return updatedIds;
    }

    // Находим столбец ID на целевом листе
    var targetLastColumn = targetSheet.getLastColumn();
    var targetHeaders = targetSheet.getRange(1, 1, 1, targetLastColumn).getValues()[0];
    var targetIdIndex = -1;
    var targetStatusIndex = -1;

    for (var i = 0; i < targetHeaders.length; i++) {
      var header = String(targetHeaders[i] || '').trim();
      if (header === 'ID') targetIdIndex = i;
      if (header === rule.targetHeader) targetStatusIndex = i;
    }

    if (targetIdIndex === -1 || targetStatusIndex === -1) {
      Lib.logWarn('[' + projectKey + '] На листе "' + rule.targetSheet + '" не найдены столбцы ID или ' + rule.targetHeader);
      return updatedIds;
    }

    // Получаем все данные целевого листа
    var targetLastRow = targetSheet.getLastRow();
    if (targetLastRow <= 1) return updatedIds;

    var targetData = targetSheet.getRange(2, targetIdIndex + 1, targetLastRow - 1, 1).getValues();
    var targetIdMap = {};
    for (var i = 0; i < targetData.length; i++) {
      var id = String(targetData[i][0] || '').trim();
      if (id) {
        targetIdMap[id] = i + 2; // +2 потому что строка 1 = заголовок, и индекс с 0
      }
    }

    // Для каждой обновлённой строки находим соответствие по ID
    var updatedCount = 0;
    for (var i = 0; i < statusUpdates.length; i++) {
      var update = statusUpdates[i];
      var rowNumber = update.rowNumber;

      // Получаем ID из листа Заказ
      var id = String(orderSheet.getRange(rowNumber, 1).getValue() || '').trim();
      if (!id) continue;

      // Находим строку на целевом листе
      var targetRow = targetIdMap[id];
      if (targetRow) {
        // Обновляем статус
        targetSheet.getRange(targetRow, targetStatusIndex + 1).setValue(newStatusValue);
        updatedCount++;
        updatedIds.push(id); // Запоминаем ID
      }
    }

    if (updatedCount > 0) {
      Lib.logInfo('[' + projectKey + '] Обновлено статусов на листе "' + rule.targetSheet + '": ' + updatedCount);
    }

    return updatedIds;
  }

  /**
   * Синхронизация архивных записей с листом "ТЗ по статусам"
   * Данные берутся из листа Главная (основная информация) и листа Заказ (остатки и продажи)
   */
  function _syncArchivedToTaskSheet(ss, orderData, orderHeaders, rowsToSync, projectKey) {
    if (!rowsToSync || rowsToSync.length === 0) {
      return;
    }

    Lib.logInfo('[' + projectKey + '] Начало синхронизации архивных записей с ТЗ');

    var TASK_SHEET_NAME = 'ТЗ по статусам';
    var taskSheet = ss.getSheetByName(TASK_SHEET_NAME);

    if (!taskSheet) {
      taskSheet = ss.insertSheet(TASK_SHEET_NAME);
      Lib.logInfo('[' + projectKey + '] Создан новый лист: ' + TASK_SHEET_NAME);
    }

    // Проверяем/создаём заголовки
    _ensureTaskSheetHeaders(taskSheet, projectKey);

    // Получаем текущие данные из ТЗ
    var taskData = taskSheet.getDataRange().getValues();
    var taskHeaders = taskData.length > 0 ? taskData[0].map(function(h) { return String(h || '').trim(); }) : [];
    var taskHeaderIndex = _makeHeaderIndex(taskHeaders);

    // Создаём карту ID из ТЗ для быстрой проверки
    var existingIds = {};
    for (var i = 1; i < taskData.length; i++) {
      var idIndex = taskHeaderIndex['ID'];
      if (idIndex !== undefined) {
        var id = taskData[i][idIndex];
        if (id) existingIds[String(id)] = true;
      }
    }

    // Получаем лист Главная
    var primarySheetName = (global.CONFIG && global.CONFIG.SHEETS && global.CONFIG.SHEETS.PRIMARY) || 'Главная';
    var primarySheet = ss.getSheetByName(primarySheetName);

    if (!primarySheet) {
      Lib.logError('[' + projectKey + '] Лист "' + primarySheetName + '" не найден для синхронизации с ТЗ');
      return;
    }

    // Получаем данные листа Главная
    var primaryData = primarySheet.getDataRange().getValues();
    if (primaryData.length <= 1) {
      Lib.logError('[' + projectKey + '] Нет данных на листе "' + primarySheetName + '"');
      return;
    }

    var primaryHeaders = primaryData[0].map(function(h) { return String(h || '').trim(); });
    var primaryHeaderIndex = _makeHeaderIndex(primaryHeaders);
    var primaryIdIndex = primaryHeaderIndex['ID'];

    if (primaryIdIndex === undefined) {
      Lib.logError('[' + projectKey + '] Не найден столбец ID на листе "' + primarySheetName + '"');
      return;
    }

    // Создаём карту данных из листа Главная по ID
    var primaryDataMap = {};
    for (var i = 1; i < primaryData.length; i++) {
      var id = primaryData[i][primaryIdIndex];
      if (id) {
        primaryDataMap[String(id)] = primaryData[i];
      }
    }

    // Создаём индекс заголовков листа Заказ
    var orderHeaderIndex = _makeHeaderIndex(orderHeaders);
    var orderIdIndex = orderHeaderIndex['ID'];
    var orderStockIndex = orderHeaderIndex['Остаток'];
    var orderSalesIndex = orderHeaderIndex['ПРОДАЖИ'];

    // Собираем строки для добавления в ТЗ
    var rowsToAdd = [];
    for (var i = 0; i < rowsToSync.length; i++) {
      var rowIndex = rowsToSync[i];
      var orderRow = orderData.values[rowIndex];
      var id = orderRow[orderIdIndex];

      // Проверяем, что строка ещё не добавлена в ТЗ
      if (id && !existingIds[String(id)]) {
        var primaryRow = primaryDataMap[String(id)];

        if (primaryRow) {
          // Извлекаем остаток и продажи из листа Заказ
          var stock = (orderStockIndex !== undefined && orderStockIndex !== -1) ? orderRow[orderStockIndex] : '';
          var sales = (orderSalesIndex !== undefined && orderSalesIndex !== -1) ? orderRow[orderSalesIndex] : '';

          var newRow = _extractRowForTaskSheet(primaryRow, primaryHeaderIndex, stock, sales);
          if (newRow) {
            rowsToAdd.push(newRow);
            existingIds[String(id)] = true;
          }
        } else {
          Lib.logWarn('[' + projectKey + '] Не найдена строка с ID=' + id + ' на листе "' + primarySheetName + '"');
        }
      }
    }

    // Добавляем новые строки в ТЗ
    if (rowsToAdd.length > 0) {
      var lastRow = taskSheet.getLastRow();
      var targetRange = taskSheet.getRange(lastRow + 1, 1, rowsToAdd.length, rowsToAdd[0].length);
      targetRange.setValues(rowsToAdd);

      // Добавляем чекбоксы в столбец "Отработано"
      var completedColIndex = taskHeaderIndex['Отработано'];
      if (completedColIndex !== undefined) {
        var checkboxRange = taskSheet.getRange(lastRow + 1, completedColIndex + 1, rowsToAdd.length, 1);
        checkboxRange.insertCheckboxes();
      }

      Lib.logInfo('[' + projectKey + '] Добавлено в ТЗ архивных записей: ' + rowsToAdd.length);
    }
  }

  /**
   * Создаёт заголовки на листе ТЗ по статусам
   */
  function _ensureTaskSheetHeaders(sheet, projectKey) {
    var TASK_SHEET_HEADERS = [
      'ID',
      'Статус',
      'Арт. Рус',
      'Арт. произв.',
      'Код база',
      'Код в 1С',
      'Название  ENG прайс произв',
      'Наименования англ по ДС',
      'Наименования рус по ДС',
      'Объём',
      'Категория товара',
      'Остаток',
      'Продажи',
      'Комментарий',
      'Отработано',
    ];

    var lastColumn = sheet.getLastColumn();

    if (lastColumn === 0 || sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, TASK_SHEET_HEADERS.length).setValues([TASK_SHEET_HEADERS]);

      // Форматируем заголовки
      sheet.getRange(1, 1, 1, TASK_SHEET_HEADERS.length)
        .setFontWeight('bold')
        .setBackground('#4a86e8')
        .setFontColor('#ffffff');

      // Замораживаем первую строку
      sheet.setFrozenRows(1);

      Lib.logInfo('[' + projectKey + '] Созданы заголовки для ТЗ по статусам');
    }
  }

  /**
   * Извлекает данные строки для листа ТЗ
   * Данные берутся из листа Главная, остатки и продажи передаются отдельно
   */
  function _extractRowForTaskSheet(primaryRow, primaryHeaderIndex, stock, sales) {
    var COLUMNS_TO_COPY_FROM_PRIMARY = [
      'ID',
      'Статус',
      'Арт. Рус',
      'Арт. произв.',
      'Код база',
      'Код в 1С',
      'Название  ENG прайс произв',
      'Наименования англ по ДС',
      'Наименования рус по ДС',
      'Объём',
      'Категория товара',
    ];

    var newRow = [];

    // Копируем столбцы из листа Главная
    for (var i = 0; i < COLUMNS_TO_COPY_FROM_PRIMARY.length; i++) {
      var columnName = COLUMNS_TO_COPY_FROM_PRIMARY[i];
      var colIndex = primaryHeaderIndex[columnName];

      if (colIndex !== undefined) {
        newRow.push(primaryRow[colIndex]);
      } else {
        newRow.push('');
      }
    }

    // Добавляем Остаток и Продажи из листа Заказ
    newRow.push(stock || '');
    newRow.push(sales || '');

    // Добавляем пустое значение для комментария
    newRow.push('');

    // Добавляем пустое значение для чекбокса "Отработано"
    newRow.push(false);

    return newRow;
  }

  /**
   * Создаёт карту индексов заголовков
   */
  function _makeHeaderIndex(headers) {
    var map = {};
    for (var i = 0; i < headers.length; i++) {
      var key = String(headers[i] || '').trim();
      if (key && !(key in map)) {
        map[key] = i;
      }
    }
    return map;
  }
})(Lib, this);
