var Lib = Lib || {};

(function (Lib, global) {
  var TARGET_PROJECT_KEY = "SS";
  var OUTPUT_HEADERS = [
    "ID-P",
    "Арт. произв.",
    "Название ENG прайс произв",
    "Объём",
    "Форма выпуска",
    "BAR CODE",
    "шт./уп.",
    "Цена",
    "Группа",
  ];

  Lib.processSsPriceSheet = function () {
    var ui = SpreadsheetApp.getUi();
    var config = _getPrimaryDataConfig_();
    var menuTitle = _getMenuTitle_(config);

    if (!_isActiveProject_()) {
      ui.alert(
        menuTitle,
        "Эта функция доступна только в проекте SS.",
        ui.ButtonSet.OK
      );
      return;
    }

    try {
      Lib.logInfo("[SS] Обработка основной: старт");
      var source = _getSourceData_(config, "MAIN");
      if (!source.values || !source.values.length) {
        throw new Error("В исходном документе нет данных для обработки.");
      }

      var processed = _buildProcessedMainData_(source.values, config);
      if (!processed.rows.length) {
        ui.alert(
          menuTitle,
          "Не найдено данных для обработки. Проверьте исходный прайс-лист.",
          ui.ButtonSet.OK
        );
        return;
      }

      // Очищаем ID-P на всех листах (Главная, Заказ, Динамика цены, Расчет цены, ABC-Анализ)
      if (typeof Lib.clearIdpColumnOnSheets === 'function') {
        Lib.clearIdpColumnOnSheets();
      } else {
        // Fallback на старую функцию, если новая недоступна
        _clearIdpColumnOnMain();
      }

      // Очищаем ID-L на листе Прайс
      if (typeof Lib.clearIdlColumnOnSheets === 'function') {
        Lib.clearIdlColumnOnSheets();
      }

      // Очищаем ЦЕНА EXW из Б/З на листах Динамика цены и Расчет цены
      if (typeof Lib.clearPriceExwColumnOnSheets === 'function') {
        Lib.clearPriceExwColumnOnSheets();
      }

      _clearIdgColumnOnMain();

      var syncResult = _syncIdWithMain(
        processed.rows,
        processed.headers,
        processed.articles,
        processed.groups
      );

      _applyAssignedIdpToProcessed_(
        processed,
        syncResult && syncResult.assignedIdp
      );
      _storeProcessedSnapshot_(config, processed, "MAIN");

      // Обработка изменений групп (показываем диалоги)
      if (syncResult && syncResult.groupChanges && syncResult.groupChanges.length > 0) {
        _handleGroupChanges_(syncResult.groupChanges);
      }

      // Проставляем ID-G для строк без ID-P по совпадению Группа
      _fillIdgForRowsWithoutIdp();

      // Заполняем ID-P на остальных листах (Заказ, Динамика цены, Расчет цены, ABC-Анализ)
      if (typeof Lib.fillIdpOnSheetsByIdFromPrimary === 'function') {
        Lib.logInfo("[SS] Заполнение ID-P на всех листах");
        Lib.fillIdpOnSheetsByIdFromPrimary();
      } else {
        Lib.logWarn("[SS] fillIdpOnSheetsByIdFromPrimary не найдена");
      }

      // Копируем цену из Б/З в столбец "ЦЕНА EXW из Б/З" на листах Динамика цены и Расчет цены
      if (typeof Lib.copyPriceFromPrimaryToSheets === 'function') {
        Lib.logInfo("[SS] Копирование цены из Б/З на листы Динамика цены и Расчет цены");
        Lib.copyPriceFromPrimaryToSheets(processed);
      } else {
        Lib.logWarn("[SS] copyPriceFromPrimaryToSheets не найдена");
      }

      // Применяем формулы на листе "Динамика цены" (EXW ALFASPA, Закупочная цена, DDP-МОСКВА, Прирост)
      if (typeof Lib.recalculatePriceDynamicsFormulas === 'function') {
        Lib.logInfo("[SS] Применение формул на листе Динамика цены");
        Lib.recalculatePriceDynamicsFormulas();
      } else {
        Lib.logWarn("[SS] recalculatePriceDynamicsFormulas не найдена");
      }

      if (
        syncResult &&
        syncResult.createdRows &&
        syncResult.createdRows.length
      ) {
        var newArticleRows = _handleNewArticles_(
          syncResult.createdRows,
          processed.rows,
          processed.headers
        );

        // Запускаем построчную синхронизацию для новых артикулов со статусом "New завод"
        if (newArticleRows && newArticleRows.length > 0) {
          Lib.logInfo(
            "[SS] Запуск построчной синхронизации для " +
              newArticleRows.length +
              " новых артикулов"
          );
          SpreadsheetApp.flush();

          if (typeof Lib.syncMultipleRows === "function") {
            Lib.syncMultipleRows(newArticleRows);
          } else {
            Lib.logWarn("[SS] syncMultipleRows не найдена");
          }

          // Устанавливаем цены из processed для новых артикулов
          if (
            newArticleRows.articleCodes &&
            newArticleRows.articleCodes.length &&
            typeof _setPriceFromProcessedByArticle_ === "function"
          ) {
            newArticleRows.articleCodes.forEach(function(articleCode) {
              if (articleCode) {
                _setPriceFromProcessedByArticle_(
                  articleCode,
                  processed,
                  newArticleRows.articlePriceMap
                );
              }
            });
          }

          // Заполняем ID-P для новых строк после синхронизации
          if (typeof Lib.fillIdpOnSheetsByIdFromPrimary === 'function') {
            Lib.logInfo("[SS] Заполнение ID-P для новых артикулов на всех листах");
            Lib.fillIdpOnSheetsByIdFromPrimary();
          }

          // Заполняем ЦЕНА EXW из Б/З для новых артикулов
          if (typeof Lib.copyPriceFromPrimaryToSheets === 'function') {
            Lib.logInfo("[SS] Копирование цен для новых артикулов на Динамика цены и Расчет цены");
            Lib.copyPriceFromPrimaryToSheets(processed);
          }
        }
      }

      Lib.logInfo(
        "[SS] Обработка основной: завершено, строк " + processed.rows.length
      );

      // Обновляем формулы на листе "Расчет цены"
      // 1. Сначала применяем INDEX/MATCH формулы для подтягивания данных из "Динамика цены"
      if (typeof Lib.updatePriceCalculationFormulas === 'function') {
        Lib.logInfo("[SS] Обновление INDEX/MATCH формул на листе Расчет цены");
        Lib.updatePriceCalculationFormulas(true); // silent=true
      } else {
        Lib.logWarn("[SS] updatePriceCalculationFormulas не найдена");
      }

      // 2. Затем применяем расчетные формулы (К-т, Расчетная цена Опт, РРЦ и т.д.)
      if (typeof Lib.applyCalculationFormulas === 'function') {
        Lib.logInfo("[SS] Применение расчетных формул на листе Расчет цены");
        Lib.applyCalculationFormulas(true); // silent=true
      } else {
        Lib.logWarn("[SS] applyCalculationFormulas не найдена");
      }

      // Обновляем статусы и синхронизируем с ТЗ по статусам
      if (typeof Lib.updateStatusesAfterProcessing === 'function') {
        Lib.updateStatusesAfterProcessing();
      }

      var message = "Обработка завершена. Найдено товаров: " + processed.rows.length + ".";

      // Добавляем информацию о несовпадениях BAR CODE
      if (syncResult && syncResult.barcodeMismatches && syncResult.barcodeMismatches.length > 0) {
        message += "\n\n⚠️ Обнаружены несовпадения BAR CODE (" + syncResult.barcodeMismatches.length + "):\n";
        for (var bm = 0; bm < Math.min(syncResult.barcodeMismatches.length, 10); bm++) {
          var mismatch = syncResult.barcodeMismatches[bm];
          message += "\n• " + mismatch.article + ":\n  База: " + mismatch.existing + "\n  Прайс: " + mismatch.newValue;
        }
        if (syncResult.barcodeMismatches.length > 10) {
          message += "\n\n... и ещё " + (syncResult.barcodeMismatches.length - 10) + " несовпадений.";
        }
        message += "\n\nПроверьте журнал для подробностей.";
      }

      ui.alert(menuTitle, message, ui.ButtonSet.OK);
    } catch (error) {
      Lib.logError("processSsPriceSheet: ошибка", error);
      ui.alert(
        "Ошибка обработки прайс",
        error.message || String(error),
        ui.ButtonSet.OK
      );
    }
  };

  Lib.loadSsStockData = function () {
    var ui = SpreadsheetApp.getUi();
    var config = _getPrimaryDataConfig_();
    var menuTitle = _getMenuTitle_(config);

    if (!_isActiveProject_()) {
      ui.alert(
        menuTitle,
        "Эта функция доступна только в проекте SS.",
        ui.ButtonSet.OK
      );
      return;
    }

    // Используем универсальную функцию загрузки остатков
    if (typeof Lib.loadStockData === 'function') {
      Lib.loadStockData('SS');
      return;
    }

    // DEPRECATED: Старая реализация (будет удалена)
    try {
      Lib.logInfo("[SS] Загрузка остатков: старт (DEPRECATED - используйте Lib.loadStockData)");

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

      var stocksSheet = ss.getSheetByName(stocksSheetName);
      if (!stocksSheet) {
        var sourceSheetName =
          _resolveSheetName_(config, "SOURCE", "STOCKS") || stocksSheetName;
        var sourceDocId = config.SOURCE && config.SOURCE.DOC_ID;

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
          '[SS] Загрузка остатков: данные читаются из внешнего документа ' +
            sourceSheetName
        );
      }

      var orderLastRow = orderSheet.getLastRow();
      if (orderLastRow <= 1) {
        ui.alert(
          menuTitle,
          'На листе "' + orderSheetName + '" нет строк для обновления.',
          ui.ButtonSet.OK
        );
        return;
      }

      var orderLastColumn = orderSheet.getLastColumn();
      var orderHeaderInfo = _findHeaderRow_(
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
        return;
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
        return;
      }

      var orderDataRange = orderSheet.getRange(
        orderDataStartRow,
        1,
        orderDataRowCount,
        orderLastColumn
      );
      var orderValues = orderDataRange.getValues();
      var numRows = orderValues.length;

      var stocksLastRow = stocksSheet.getLastRow();
      if (stocksLastRow <= 1) {
        ui.alert(
          menuTitle,
          'На листе "' + stocksSheetName + '" нет данных.',
          ui.ButtonSet.OK
        );
        return;
      }

      var stocksLastColumn = stocksSheet.getLastColumn();
      var stocksHeaderInfo = _findHeaderRow_(
        stocksSheet,
        stocksLastColumn,
        ["Артикул"]
      );
      if (!stocksHeaderInfo) {
        throw new Error(
          'Не удалось определить строку заголовков на листе "' +
            stocksSheetName +
            '" — не найден столбец "Артикул".'
        );
      }

      var stocksHeaderRow = stocksHeaderInfo.row;
      var stocksHeaders = stocksHeaderInfo.headers;

      if (stocksLastRow <= stocksHeaderRow) {
        ui.alert(
          menuTitle,
          'На листе "' + stocksSheetName + '" нет данных.',
          ui.ButtonSet.OK
        );
        return;
      }

      var stocksIdx = {
        article: stocksHeaders.indexOf("Артикул"),
        sold: stocksHeaders.indexOf("Продано за период"),
        writtenOff: stocksHeaders.indexOf("Списано за период"),
        stock: stocksHeaders.indexOf("Остаток на текущий день"),
        reserve: stocksHeaders.indexOf("Из них в резерве"),
        total: stocksHeaders.indexOf("Всего"),
        expiry: stocksHeaders.indexOf("Срок годности"),
      };

      if (stocksIdx.article === -1) {
        throw new Error(
          'На листе "' + stocksSheetName + '" отсутствует столбец "Артикул".'
        );
      }

      var stocksDataStartRow = stocksHeaderRow + 1;
      var stocksDataRowCount = stocksLastRow - stocksHeaderRow;
      if (stocksDataRowCount <= 0) {
        ui.alert(
          menuTitle,
          'На листе "' + stocksSheetName + '" нет данных.',
          ui.ButtonSet.OK
        );
        return;
      }

      var stocksDataRange = stocksSheet.getRange(
        stocksDataStartRow,
        1,
        stocksDataRowCount,
        stocksLastColumn
      );
      var stocksValues = stocksDataRange.getValues();

      var stocksMap = Object.create(null);

      var currentArticle = "";
      var currentSold = null;
      var currentWrittenOff = null;
      var currentStock = null;

      stocksValues.forEach(function (row) {
        var articleCell =
          stocksIdx.article !== -1
            ? _asTrimmedString(row[stocksIdx.article])
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

        if (stocksIdx.sold !== -1) {
          var soldCell = row[stocksIdx.sold];
          if (soldCell !== "" && soldCell !== null) {
            currentSold = _parseNumber(soldCell);
          }
        }
        if (stocksIdx.writtenOff !== -1) {
          var writtenCell = row[stocksIdx.writtenOff];
          if (writtenCell !== "" && writtenCell !== null) {
            currentWrittenOff = _parseNumber(writtenCell);
          }
        }
        if (stocksIdx.stock !== -1) {
          var stockCell = row[stocksIdx.stock];
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

        if (stocksIdx.reserve !== -1) {
          var reserveCell = row[stocksIdx.reserve];
          var reserveValue = _parseNumber(reserveCell);
          if (reserveValue !== null && !isNaN(reserveValue)) {
            entry.reserve += reserveValue;
          }
        }

        if (stocksIdx.total !== -1) {
          var totalCell = row[stocksIdx.total];
          var totalValue = _parseNumber(totalCell);
          if (totalValue !== null && !isNaN(totalValue) && totalValue !== 0) {
            var expiryCell =
              stocksIdx.expiry !== -1 ? row[stocksIdx.expiry] : "";
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

      for (var i = 0; i < numRows; i++) {
        var orderRow = orderValues[i];
        var article = _asTrimmedString(orderRow[orderIdx.article]);
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

      if (orderIdx.sales !== -1) {
        orderSheet
          .getRange(orderDataStartRow, orderIdx.sales + 1, numRows, 1)
          .setValues(salesColumnValues);
      }
      if (orderIdx.writtenOff !== -1) {
        orderSheet
          .getRange(orderDataStartRow, orderIdx.writtenOff + 1, numRows, 1)
          .setValues(writtenOffColumnValues);
      }
      if (orderIdx.stock !== -1) {
        orderSheet
          .getRange(orderDataStartRow, orderIdx.stock + 1, numRows, 1)
          .setValues(stockColumnValues);
      }
      if (orderIdx.inTransit !== -1) {
        orderSheet
          .getRange(orderDataStartRow, orderIdx.inTransit + 1, numRows, 1)
          .setValues(inTransitColumnValues);
      }
      if (orderIdx.reserve !== -1) {
        orderSheet
          .getRange(orderDataStartRow, orderIdx.reserve + 1, numRows, 1)
          .setValues(reserveColumnValues);
      }
      if (orderIdx.qty1 !== -1) {
        orderSheet
          .getRange(orderDataStartRow, orderIdx.qty1 + 1, numRows, 1)
          .setValues(qty1ColumnValues);
      }
      if (orderIdx.exp1 !== -1) {
        orderSheet
          .getRange(orderDataStartRow, orderIdx.exp1 + 1, numRows, 1)
          .setValues(exp1ColumnValues);
      }
      if (orderIdx.qty2 !== -1) {
        orderSheet
          .getRange(orderDataStartRow, orderIdx.qty2 + 1, numRows, 1)
          .setValues(qty2ColumnValues);
      }
      if (orderIdx.exp2 !== -1) {
        orderSheet
          .getRange(orderDataStartRow, orderIdx.exp2 + 1, numRows, 1)
          .setValues(exp2ColumnValues);
      }
      if (orderIdx.qty3 !== -1) {
        orderSheet
          .getRange(orderDataStartRow, orderIdx.qty3 + 1, numRows, 1)
          .setValues(qty3ColumnValues);
      }
      if (orderIdx.exp3 !== -1) {
        orderSheet
          .getRange(orderDataStartRow, orderIdx.exp3 + 1, numRows, 1)
          .setValues(exp3ColumnValues);
      }

      // После загрузки остатков пересчитаем производные столбцы на листе Заказ
      try {
        if (typeof Lib.recalculateOrderSheet === 'function') {
          Lib.recalculateOrderSheet();
          Lib.logInfo('[SS] Производные столбцы пересчитаны после загрузки остатков');
        }
      } catch (e) {
        Lib.logWarn('[SS] Не удалось пересчитать производные столбцы заказа после загрузки остатков', e);
      }

      Lib.logInfo("[SS] Загрузка остатков: обновлено строк " + numRows);
      ui.alert(
        menuTitle,
        "Загрузка остатков завершена. Обновлено строк: " + numRows + ".",
        ui.ButtonSet.OK
      );
    } catch (error) {
      Lib.logError("loadSsStockData: ошибка", error);
      ui.alert(
        "Ошибка загрузки остатков",
        error.message || String(error),
        ui.ButtonSet.OK
      );
    }
  };

  /**
   * УСТАРЕВШАЯ: Сортировка заказа (использует группировку по умолчанию)
   * @deprecated Используйте sortSsOrderByManufacturer или sortSsOrderByPrice
   */
  Lib.sortSsOrderData = function () {
    Lib.sortSsOrderByManufacturer();
  };

  /**
   * ПУБЛИЧНАЯ: Сортировка заказа по Производителю (Группа + ID-G)
   * Теперь сортирует одновременно три листа: Заказ, Динамика цены, Расчет цены
   */
  Lib.sortSsOrderByManufacturer = function () {
    var ui = SpreadsheetApp.getUi();
    var config = _getPrimaryDataConfig_();
    var menuTitle = _getMenuTitle_(config);

    if (!_isActiveProject_()) {
      ui.alert(
        menuTitle,
        "Эта функция доступна только в проекте SS.",
        ui.ButtonSet.OK
      );
      return;
    }

    try {
      if (typeof Lib.structureMultipleSheets !== "function") {
        throw new Error(
          "Функция Lib.structureMultipleSheets недоступна. Обновите библиотеку."
        );
      }
      Lib.structureMultipleSheets('byManufacturer');
    } catch (error) {
      Lib.logError("[SS] Сортировка по Производителю: ошибка", error);
      ui.alert(
        menuTitle,
        error.message || String(error),
        ui.ButtonSet.OK
      );
    }
  };

  /**
   * ПУБЛИЧНАЯ: Сортировка заказа по Прайсу (Линия + ID-L)
   * Теперь сортирует одновременно три листа: Заказ, Динамика цены, Расчет цены
   */
  Lib.sortSsOrderByPrice = function () {
    var ui = SpreadsheetApp.getUi();
    var config = _getPrimaryDataConfig_();
    var menuTitle = _getMenuTitle_(config);

    if (!_isActiveProject_()) {
      ui.alert(
        menuTitle,
        "Эта функция доступна только в проекте SS.",
        ui.ButtonSet.OK
      );
      return;
    }

    try {
      if (typeof Lib.structureMultipleSheets !== "function") {
        throw new Error(
          "Функция Lib.structureMultipleSheets недоступна. Обновите библиотеку."
        );
      }
      Lib.structureMultipleSheets('byPrice');
    } catch (error) {
      Lib.logError("[SS] Сортировка по Прайсу: ошибка", error);
      ui.alert(
        menuTitle,
        error.message || String(error),
        ui.ButtonSet.OK
      );
    }
  };

  function _findHeaderRow_(sheet, lastColumn, requiredHeaders) {
    if (!sheet || lastColumn <= 0) return null;
    var maxRowsToScan = Math.min(sheet.getLastRow(), 10);
    for (var row = 1; row <= maxRowsToScan; row++) {
      var headersRaw = sheet.getRange(row, 1, 1, lastColumn).getValues();
      if (!headersRaw || !headersRaw.length) continue;
      var headers = headersRaw[0].map(function (value) {
        return String(value || "").trim();
      });
      var hasAnyContent = headers.some(function (value) {
        return value !== "";
      });
      if (!hasAnyContent) {
        continue;
      }
      if (
        requiredHeaders &&
        requiredHeaders.some(function (required) {
          return headers.indexOf(required) === -1;
        })
      ) {
        continue;
      }
      return { row: row, headers: headers };
    }
    return null;
  }

  function _getPrimaryDataConfig_() {
    var cfg = global.CONFIG && global.CONFIG.PRIMARY_DATA;
    if (!cfg) {
      throw new Error("Конфигурация PRIMARY_DATA не настроена для проекта SS.");
    }
    return cfg;
  }

  function _isActiveProject_() {
    return (
      global.CONFIG && global.CONFIG.ACTIVE_PROJECT_KEY === TARGET_PROJECT_KEY
    );
  }

  function _getMenuTitle_(config) {
    var defaultTitle = "🧾 Заказ";
    if (!config || !config.MENU || !config.MENU.TITLE) {
      return defaultTitle;
    }
    return config.MENU.TITLE;
  }

  function recreateSheet(ss, sheetName) {
    var existingSheet = ss.getSheetByName(sheetName);
    if (existingSheet) {
      ss.deleteSheet(existingSheet);
    }
    var newSheet = ss.insertSheet(sheetName);
    return newSheet;
  }

  function _getSourceData_(config, key) {
    var sourceInfo = config && config.SOURCE;
    if (!sourceInfo || !sourceInfo.DOC_ID) {
      throw new Error(
        "В конфигурации не указан исходный документ для обработки."
      );
    }
    var baseSheetName = _resolveSheetName_(config, "SOURCE", key);
    if (!baseSheetName) {
      throw new Error(
        "В конфигурации не указан лист источника для ключа " + key + "."
      );
    }

    var sourceSpreadsheet;
    var active = SpreadsheetApp.getActiveSpreadsheet();
    if (active && active.getId() === sourceInfo.DOC_ID) {
      sourceSpreadsheet = active;
    } else {
      sourceSpreadsheet = SpreadsheetApp.openById(sourceInfo.DOC_ID);
    }

    var originalSheet = sourceSpreadsheet.getSheetByName(baseSheetName);
    if (!originalSheet) {
      throw new Error(
        'Не найден лист "' + baseSheetName + '" в документе источнике.'
      );
    }

    var originalRange = originalSheet.getDataRange();
    var originalValues = originalRange.getValues();
    var originalBackgrounds = originalRange.getBackgrounds();

    var timezone =
      (Lib.CONFIG && Lib.CONFIG.SETTINGS && Lib.CONFIG.SETTINGS.TIMEZONE) ||
      "Europe/Moscow";
    var dateString = Utilities.formatDate(new Date(), timezone, "dd.MM.yy");
    var fullSheetName = baseSheetName + " " + dateString;

    try {
      var allSheets = sourceSpreadsheet.getSheets();
      for (var si = 0; si < allSheets.length; si++) {
        var sh = allSheets[si];
        var n = sh.getName();
        if (n.indexOf(baseSheetName + " ") === 0 && n !== fullSheetName) {
          try {
            sourceSpreadsheet.deleteSheet(sh);
          } catch (e) {
            /* пропустить */
          }
        }
      }
    } catch (e) {
      // ignore
    }

    var sheet = recreateSheet(sourceSpreadsheet, fullSheetName);
    if (!sheet) {
      throw new Error(
        'Не удалось создать лист "' + fullSheetName + '" в документе источнике.'
      );
    }

    if (originalValues && originalValues.length > 0) {
      var targetRange = sheet.getRange(
        1,
        1,
        originalValues.length,
        originalValues[0].length
      );
      targetRange.setValues(originalValues);
      try {
        targetRange.setBackgrounds(originalBackgrounds);
      } catch (e) {
        /* ignore */
      }

      try {
        var fullRange = sheet.getDataRange();
        fullRange.clearDataValidations();
      } catch (e) {
        /* ignore */
      }
    }

    return {
      values: originalValues,
      backgrounds: originalBackgrounds,
    };
  }

  function _resolveSheetName_(config, section, key) {
    if (!config || !config.SHEETS || !config.SHEETS[section]) {
      return null;
    }
    return config.SHEETS[section][key] || null;
  }

  function _buildProcessedMainData_(values, config) {
    var rows = [];
    var articles = [];
    var groups = [];
    var currentGroup = "";
    var isProfessionalMode = false;

    // Ищем строку заголовков (строка 2, индекс 1)
    var headerRowIndex = 1;
    if (values.length <= headerRowIndex) {
      throw new Error("Недостаточно строк в исходном документе.");
    }

    var headers = values[headerRowIndex];
    var codeIdx = -1;
    var productNameIdx = -1;
    var sizeIdx = -1;
    var packIdx = -1;
    var barCodeIdx = -1;
    var qtyBoxIdx = -1;
    var priceIdx = -1;

    // Определяем индексы столбцов по заголовкам
    for (var h = 0; h < headers.length; h++) {
      var headerText = String(headers[h] || "").trim().toUpperCase();
      if (headerText === "CODE") codeIdx = h;
      else if (headerText === "PRODUCT NAME") productNameIdx = h;
      else if (headerText === "SIZE") sizeIdx = h;
      else if (headerText === "PACK") packIdx = h;
      else if (headerText === "BAR CODE/ACL") barCodeIdx = h;
      else if (headerText === "QTY/BOX") qtyBoxIdx = h;
      else if (headerText === "EX WORKS CARROS (06)") priceIdx = h;
    }

    if (codeIdx === -1 || productNameIdx === -1) {
      throw new Error(
        "Не найдены обязательные столбцы Code или Product name в строке заголовков."
      );
    }

    // Обрабатываем данные начиная со строки после заголовков
    for (var rowIndex = headerRowIndex + 1; rowIndex < values.length; rowIndex++) {
      var row = values[rowIndex];

      var codeValue = _asTrimmedString(row[codeIdx]);
      var productNameValue = _asTrimmedString(row[productNameIdx]);
      var sizeValue = sizeIdx !== -1 ? _getValue(row, sizeIdx) : "";

      // Проверка на маркер окончания обработки
      if (productNameValue.toUpperCase() === "PROMOTIONAL MATERIALS") {
        break;
      }

      // Проверка на режим PROFESSIONAL PRODUCTS
      if (productNameValue.toUpperCase() === "PROFESSIONAL PRODUCTS") {
        isProfessionalMode = true;
        continue;
      }

      // Определение группы: если Product name заполнен, а Code и Size пусты
      if (productNameValue && !codeValue && !sizeValue) {
        currentGroup = productNameValue;
        // Если есть SAMPLES, не добавляем "-ПРОФ"
        if (isProfessionalMode && productNameValue.toUpperCase().indexOf("SAMPLES") === -1) {
          currentGroup = productNameValue + "-ПРОФ";
        }
        continue;
      }

      // Если нет артикула, пропускаем строку
      if (!codeValue) {
        continue;
      }

      // Извлекаем данные продукта
      var packValue = packIdx !== -1 ? _getValue(row, packIdx) : "";
      var barCodeValue = barCodeIdx !== -1 ? _getValue(row, barCodeIdx) : "";
      var qtyBoxValue = qtyBoxIdx !== -1 ? _getValue(row, qtyBoxIdx) : "";
      var priceValue = priceIdx !== -1 ? _getValue(row, priceIdx) : "";

      rows.push([
        "", // ID-P (будет заполнен позже)
        codeValue, // Арт. произв.
        productNameValue, // Название ENG прайс произв
        sizeValue, // Объём
        packValue, // Форма выпуска
        barCodeValue, // BAR CODE
        qtyBoxValue, // шт./уп.
        priceValue, // Цена
        currentGroup, // Группа
      ]);
      articles.push(codeValue);
      groups.push(currentGroup);
    }

    return {
      headers: OUTPUT_HEADERS.slice(),
      rows: rows,
      articles: articles,
      groups: groups,
    };
  }

  function _storeProcessedSnapshot_(config, processed, sheetKey) {
    if (!processed || !processed.rows) {
      return;
    }

    var sourceInfo = config && config.SOURCE;
    if (!sourceInfo || !sourceInfo.DOC_ID) {
      return;
    }

    var resolvedKey = sheetKey || "MAIN";
    var baseSheetName = _resolveSheetName_(config, "SOURCE", resolvedKey);
    if (!baseSheetName) {
      baseSheetName = "-Б/З поставщик";
    }

    var timezone =
      (Lib.CONFIG && Lib.CONFIG.SETTINGS && Lib.CONFIG.SETTINGS.TIMEZONE) ||
      "Europe/Moscow";
    var dateLabel = Utilities.formatDate(new Date(), timezone, "dd.MM.yy");
    var snapshotName = baseSheetName + " " + dateLabel;

    var sourceSpreadsheet;
    var active = SpreadsheetApp.getActiveSpreadsheet();
    if (active && active.getId() === sourceInfo.DOC_ID) {
      sourceSpreadsheet = active;
    } else {
      sourceSpreadsheet = SpreadsheetApp.openById(sourceInfo.DOC_ID);
    }

    var snapshotSheet = sourceSpreadsheet.getSheetByName(snapshotName);
    if (!snapshotSheet) {
      snapshotSheet = sourceSpreadsheet.insertSheet(snapshotName);
    }

    var headers = processed.headers && processed.headers.length
        ? processed.headers.slice()
        : OUTPUT_HEADERS.slice();
    var rowsCount = processed.rows.length;
    var requiredRows = Math.max(rowsCount + 1, 1);
    var requiredColumns = Math.max(headers.length, 1);

    var currentRows = snapshotSheet.getMaxRows();
    if (currentRows < requiredRows) {
      snapshotSheet.insertRowsAfter(currentRows, requiredRows - currentRows);
    } else if (currentRows > requiredRows) {
      snapshotSheet.deleteRows(requiredRows + 1, currentRows - requiredRows);
    }

    var currentColumns = snapshotSheet.getMaxColumns();
    if (currentColumns < requiredColumns) {
      snapshotSheet.insertColumnsAfter(
        currentColumns,
        requiredColumns - currentColumns
      );
    } else if (currentColumns > requiredColumns) {
      snapshotSheet.deleteColumns(
        requiredColumns + 1,
        currentColumns - requiredColumns
      );
    }

    snapshotSheet.clear();

    if (headers.length) {
      snapshotSheet
        .getRange(1, 1, 1, headers.length)
        .setValues([headers])
        .setFontWeight("bold");
    }

    if (rowsCount && headers.length) {
      snapshotSheet
        .getRange(2, 1, rowsCount, headers.length)
        .setValues(processed.rows);
    }

    snapshotSheet.autoResizeColumns(1, Math.max(headers.length, 1));
  }

  function _syncIdWithMain(sourceRows, sourceHeaders, articles, groups) {
    var emptyResult = { createdRows: [], assignedIdp: {} };
    if (!global.CONFIG || !global.CONFIG.SHEETS) {
      return emptyResult;
    }
    try {
      Lib.logDebug(
        "[SS] syncIdWithMain: старт, строк " + (articles ? articles.length : 0)
      );
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var mainSheetName = global.CONFIG.SHEETS.PRIMARY;
      if (!mainSheetName) {
        return emptyResult;
      }
      var mainSheet = ss.getSheetByName(mainSheetName);
      if (!mainSheet) {
        return emptyResult;
      }
      var headers =
        typeof global.CONFIG.getHeaders === "function"
          ? global.CONFIG.getHeaders(mainSheetName)
          : [];
      if (!headers || headers.length === 0) {
        return emptyResult;
      }

      headers = headers.map(function (name) {
        return String(name || "").trim();
      });

      var headerToIndex = {};
      for (var h = 0; h < headers.length; h++) {
        var normalizedHeaderKey = _normalizeHeaderKey(headers[h]);
        if (
          normalizedHeaderKey &&
          headerToIndex[normalizedHeaderKey] === undefined
        ) {
          headerToIndex[normalizedHeaderKey] = h;
        }
      }

      var targetColumnCount = mainSheet.getLastColumn();
      if (targetColumnCount < headers.length) {
        targetColumnCount = headers.length;
      }

      var sourceHeaderToIndex = {};
      if (sourceHeaders && sourceHeaders.length) {
        for (var sh = 0; sh < sourceHeaders.length; sh++) {
          var normalizedSourceKey = _normalizeHeaderKey(sourceHeaders[sh]);
          if (
            normalizedSourceKey &&
            sourceHeaderToIndex[normalizedSourceKey] === undefined
          ) {
            sourceHeaderToIndex[normalizedSourceKey] = sh;
          }
        }
      }
      var sourceIdpIdx = sourceHeaderToIndex[_normalizeHeaderKey("ID-P")];

      var articleColumnIndex =
        headerToIndex[_normalizeHeaderKey("Арт. произв.")];
      if (articleColumnIndex === undefined) {
        return emptyResult;
      }

      var idColumnIndex = headerToIndex[_normalizeHeaderKey("ID")];
      var idpColumnIndex = headerToIndex[_normalizeHeaderKey("ID-P")];
      var idgColumnIndex = headerToIndex[_normalizeHeaderKey("ID-G")];
      var nameColumnIndex =
        headerToIndex[_normalizeHeaderKey("Название ENG прайс произв")];
      var groupColumnIndex = headerToIndex[_normalizeHeaderKey("Группа")];
      var unitsColumnIndex = headerToIndex[_normalizeHeaderKey("шт./уп.")];

      var brandPrefix =
        (global.CONFIG &&
          global.CONFIG.SETTINGS &&
          global.CONFIG.SETTINGS.BRAND_PREFIX) ||
        "";

      var mainRange = mainSheet.getDataRange();
      var mainValues = mainRange.getValues();
      if (mainValues.length <= 1) {
        mainValues = [headers.slice()];
      }

      var articleEntries = {};
      var groupIdMap = Object.create(null);
      var maxExistingIdP = 0;
      var maxExistingIdL = 0;
      var maxPrimaryNumeric = 0;

      for (var i = 1; i < mainValues.length; i++) {
        var rowArray = mainValues[i].slice();
        while (rowArray.length < targetColumnCount) {
          rowArray.push("");
        }
        if (rowArray.length > targetColumnCount) {
          rowArray = rowArray.slice(0, targetColumnCount);
        }
        var article = _asTrimmedString(mainValues[i][articleColumnIndex]);
        if (article) {
          articleEntries[article] = {
            rowArray: rowArray,
            rowNumber: i + 1,
          };
        }

        if (idpColumnIndex !== undefined) {
          var existingIdP = parseInt(mainValues[i][idpColumnIndex], 10);
          if (!isNaN(existingIdP) && existingIdP > maxExistingIdP) {
            maxExistingIdP = existingIdP;
          }
        }

        if (idgColumnIndex !== undefined) {
          var existingIdL = parseInt(mainValues[i][idgColumnIndex], 10);
          if (!isNaN(existingIdL) && existingIdL > maxExistingIdL) {
            maxExistingIdL = existingIdL;
          }

          var groupExisting =
            groupColumnIndex !== undefined
              ? _asTrimmedString(rowArray[groupColumnIndex])
              : "";
          if (groupExisting) {
            groupIdMap[groupExisting] = existingIdL;
          }
        }

        if (idColumnIndex !== undefined) {
          var primaryIdValue = _asTrimmedString(rowArray[idColumnIndex]);
          if (primaryIdValue) {
            if (brandPrefix && primaryIdValue.indexOf(brandPrefix) === 0) {
              var numeric = parseInt(
                primaryIdValue.substring(brandPrefix.length),
                10
              );
              if (!isNaN(numeric) && numeric > maxPrimaryNumeric) {
                maxPrimaryNumeric = numeric;
              }
            } else if (!brandPrefix) {
              var numericPlain = parseInt(primaryIdValue, 10);
              if (!isNaN(numericPlain) && numericPlain > maxPrimaryNumeric) {
                maxPrimaryNumeric = numericPlain;
              }
            }
          }
        }
      }

      var rowsToUpdate = {};
      var newRows = [];
      var createdEntries = [];
      var idpCounter = maxExistingIdP;
      var idlCounter = maxExistingIdL;
      var primaryCounter = maxPrimaryNumeric;
      var assignedIdpMap = {};

      for (var j = 0; j < articles.length; j++) {
        var articleCode = _asTrimmedString(articles[j]);
        if (!articleCode) {
          continue;
        }

        var groupValueInput =
          groups && groups.length > j ? _asTrimmedString(groups[j]) : "";

        var entry = articleEntries[articleCode];
        var rowValues;
        var entryMeta = null;
        if (!entry) {
          rowValues = new Array(targetColumnCount).fill("");

          var newIdL = null;
          if (idgColumnIndex !== undefined) {
            if (groupIdMap[groupValueInput]) {
              newIdL = groupIdMap[groupValueInput];
            } else {
              newIdL = ++idlCounter;
              groupIdMap[groupValueInput] = newIdL;
            }
          }

          var newPrimaryId = null;
          if (idColumnIndex !== undefined) {
            primaryCounter += 1;
            if (brandPrefix) {
              newPrimaryId =
                brandPrefix + String(primaryCounter).padStart(3, "0");
            } else {
              newPrimaryId = String(primaryCounter);
            }
          }

          var srcRowForMeta = sourceRows && sourceRows[j];
          var nameIdxForMeta =
            sourceHeaderToIndex[
              _normalizeHeaderKey("Название ENG прайс произв")
            ];
          var productNameMeta = "";
          if (
            srcRowForMeta &&
            nameIdxForMeta !== undefined &&
            nameIdxForMeta < srcRowForMeta.length
          ) {
            productNameMeta = _asTrimmedString(srcRowForMeta[nameIdxForMeta]);
          }

          entryMeta = {
            rowNumber: null,
            articleCode: articleCode,
            productName: productNameMeta || articleCode,
            idp: null,
            idl: newIdL,
            primaryId: newPrimaryId,
            group: groupValueInput,
            sourceRowIndex: j, // Индекс строки в sourceRows для получения данных
          };

          entry = {
            rowArray: rowValues,
            rowNumber: null,
            __meta: entryMeta,
          };
          articleEntries[articleCode] = entry;
          newRows.push(entry);
          createdEntries.push(entryMeta);
        } else {
          rowValues = entry.rowArray;
          while (rowValues.length < targetColumnCount) {
            rowValues.push("");
          }
          if (entry.rowNumber) {
            rowsToUpdate[entry.rowNumber] = rowValues;
          }
        }

        var assignedIdP = null;
        if (idpColumnIndex !== undefined) {
          var currentIdpValue = _parseNumber(rowValues[idpColumnIndex]);
          if (currentIdpValue !== null && !isNaN(currentIdpValue)) {
            assignedIdP = currentIdpValue;
            if (currentIdpValue > idpCounter) {
              idpCounter = currentIdpValue;
            }
          }
        }
        if (assignedIdP === null && idpColumnIndex !== undefined) {
          assignedIdP = ++idpCounter;
          Lib.logDebug(
            "[SS] syncIdWithMain: назначен новый ID-P " +
              assignedIdP +
              " для артикула " +
              articleCode
          );
        } else if (assignedIdP !== null) {
          Lib.logDebug(
            "[SS] syncIdWithMain: переиспользован ID-P " +
              assignedIdP +
              " для артикула " +
              articleCode
          );
        }
        if (idpColumnIndex !== undefined) {
          rowValues[idpColumnIndex] = assignedIdP;
        }
        if (entry.__meta) {
          entry.__meta.idp = assignedIdP;
        }

        if (assignedIdP !== null && sourceRows && sourceIdpIdx !== undefined) {
          var srcRowRefEarly = sourceRows[j];
          if (srcRowRefEarly && sourceIdpIdx < srcRowRefEarly.length) {
            srcRowRefEarly[sourceIdpIdx] = assignedIdP;
          }
        }

        // Копируем данные из прайс-листа
        var isNewArticle = entry.__meta !== null && entry.__meta !== undefined;
        var srcRow = sourceRows && sourceRows[j];

        if (isNewArticle && srcRow) {
          // Для новых артикулов: сначала заполняем обязательные поля
          if (entry.__meta) {
            if (idColumnIndex !== undefined && entry.__meta.primaryId) {
              rowValues[idColumnIndex] = entry.__meta.primaryId;
            }
            if (idgColumnIndex !== undefined && entry.__meta.idl !== null) {
              rowValues[idgColumnIndex] = entry.__meta.idl;
            }
          }

          if (articleColumnIndex !== undefined) {
            rowValues[articleColumnIndex] = articleCode;
          }

          // Затем копируем данные из прайс-листа
          if (nameColumnIndex !== undefined && srcRow) {
            var nameIdx =
              sourceHeaderToIndex[
                _normalizeHeaderKey("Название ENG прайс произв")
              ];
            if (nameIdx !== undefined && nameIdx < srcRow.length) {
              rowValues[nameColumnIndex] = srcRow[nameIdx] || "";
            }
          }

          // Группа - из groupValueInput (построена в _buildProcessedMainData_)
          if (groupColumnIndex !== undefined && groupValueInput) {
            rowValues[groupColumnIndex] = groupValueInput;
          }

          if (unitsColumnIndex !== undefined && srcRow) {
            var unitsIdx = sourceHeaderToIndex[_normalizeHeaderKey("шт./уп.")];
            if (unitsIdx !== undefined && unitsIdx < srcRow.length) {
              rowValues[unitsColumnIndex] = srcRow[unitsIdx] || "";
            }
          }

          // Копируем остальные поля из прайс-листа (Объём, Форма выпуска, BAR CODE, Цена)
          var volumeSrcIdx = sourceHeaderToIndex[_normalizeHeaderKey("Объём")];
          if (volumeSrcIdx !== undefined && volumeSrcIdx < srcRow.length) {
            var volumeTargetIdx = headerToIndex[_normalizeHeaderKey("Объём")];
            if (volumeTargetIdx !== undefined) {
              rowValues[volumeTargetIdx] = srcRow[volumeSrcIdx] || "";
            }
          }

          var formSrcIdx = sourceHeaderToIndex[_normalizeHeaderKey("Форма выпуска")];
          if (formSrcIdx !== undefined && formSrcIdx < srcRow.length) {
            var formTargetIdx = headerToIndex[_normalizeHeaderKey("Форма выпуска")];
            if (formTargetIdx !== undefined) {
              rowValues[formTargetIdx] = srcRow[formSrcIdx] || "";
            }
          }

          var barcodeSrcIdx = sourceHeaderToIndex[_normalizeHeaderKey("BAR CODE")];
          if (barcodeSrcIdx !== undefined && barcodeSrcIdx < srcRow.length) {
            var barcodeTargetIdx = headerToIndex[_normalizeHeaderKey("BAR CODE")];
            if (barcodeTargetIdx !== undefined) {
              rowValues[barcodeTargetIdx] = srcRow[barcodeSrcIdx] || "";
            }
          }

          var priceSrcIdx = sourceHeaderToIndex[_normalizeHeaderKey("Цена")];
          if (priceSrcIdx !== undefined && priceSrcIdx < srcRow.length) {
            var priceTargetIdx = headerToIndex[_normalizeHeaderKey("Цена")];
            if (priceTargetIdx !== undefined) {
              rowValues[priceTargetIdx] = srcRow[priceSrcIdx] || "";
            }
          }
        } else if (!isNewArticle && srcRow) {
          // Для существующих артикулов: проверяем изменения в столбце "Группа"
          if (groupColumnIndex !== undefined && groupValueInput) {
            var currentGroupValue = _asTrimmedString(rowValues[groupColumnIndex]);

            // Если столбец Группа пустой - заполняем без вопросов
            if (!currentGroupValue) {
              rowValues[groupColumnIndex] = groupValueInput;
              Lib.logInfo(
                "[SS] syncIdWithMain: заполнена пустая Группа '" +
                  groupValueInput +
                  "' для артикула " +
                  articleCode
              );
            } else if (currentGroupValue !== groupValueInput) {
              // Если значение отличается - запоминаем для последующего диалога
              if (!entry.__groupChange) {
                entry.__groupChange = {
                  articleCode: articleCode,
                  oldValue: currentGroupValue,
                  newValue: groupValueInput,
                  rowNumber: entry.rowNumber,
                };
              }
            }
          }
        }

        var resolvedGroup =
          groupColumnIndex !== undefined
            ? _asTrimmedString(rowValues[groupColumnIndex])
            : groupValueInput;
        if (entry.__meta) {
          entry.__meta.group = resolvedGroup;
        }

        if (idgColumnIndex !== undefined) {
          var currentIdlValue = rowValues[idgColumnIndex];
          if (
            currentIdlValue === undefined ||
            currentIdlValue === null ||
            currentIdlValue === ""
          ) {
            var mappedIdl = groupIdMap[resolvedGroup];
            if (!mappedIdl) {
              mappedIdl = ++idlCounter;
              groupIdMap[resolvedGroup] = mappedIdl;
            }
            currentIdlValue = mappedIdl;
            rowValues[idgColumnIndex] = currentIdlValue;
          }
          if (!groupIdMap[resolvedGroup]) {
            groupIdMap[resolvedGroup] = currentIdlValue;
          }

          if (entry.__meta) {
            entry.__meta.idl = currentIdlValue;
          }
        }

        var assignedIdpValue =
          idpColumnIndex !== undefined
            ? rowValues[idpColumnIndex]
            : assignedIdP;
        if (assignedIdpValue === undefined || assignedIdpValue === null) {
          assignedIdpValue = "";
        }
        assignedIdpMap[articleCode] = assignedIdpValue;
      }

      for (var rowNumber in rowsToUpdate) {
        if (!rowsToUpdate.hasOwnProperty(rowNumber)) continue;
        var rowNumInt = parseInt(rowNumber, 10);
        if (!isNaN(rowNumInt)) {
          var rowVals = rowsToUpdate[rowNumber].slice();
          while (rowVals.length < targetColumnCount) rowVals.push("");
          if (rowVals.length > targetColumnCount)
            rowVals = rowVals.slice(0, targetColumnCount);
          Lib.logDebug(
            "[SS] syncIdWithMain: обновление строки " +
              rowNumInt +
              " значениями " +
              JSON.stringify(rowVals.slice(0, 10))
          );
          mainSheet
            .getRange(rowNumInt, 1, 1, targetColumnCount)
            .setValues([rowVals]);
        }
      }

      if (newRows.length > 0) {
        var insertStart = mainSheet.getLastRow() + 1;
        mainSheet.insertRowsAfter(mainSheet.getLastRow(), newRows.length);
        var rowsValues = [];
        for (var nr = 0; nr < newRows.length; nr++) {
          var arr = newRows[nr].rowArray.slice();
          while (arr.length < targetColumnCount) arr.push("");
          if (arr.length > targetColumnCount)
            arr = arr.slice(0, targetColumnCount);
          rowsValues.push(arr);
          if (newRows[nr].rowNumber === null) {
            newRows[nr].rowNumber = insertStart + nr;
          }
          if (newRows[nr].__meta) {
            newRows[nr].__meta.rowNumber = newRows[nr].rowNumber;
          }
        }
        mainSheet
          .getRange(insertStart, 1, rowsValues.length, targetColumnCount)
          .setValues(rowsValues);
        Lib.logDebug(
          "[SS] syncIdWithMain: вставлено " +
            rowsValues.length +
            " новых строк, пример " +
            JSON.stringify(rowsValues[0] ? rowsValues[0].slice(0, 10) : [])
        );
      }

      if (idpColumnIndex !== undefined) {
        for (var articleKey in articleEntries) {
          if (
            !Object.prototype.hasOwnProperty.call(articleEntries, articleKey)
          ) {
            continue;
          }
          var entryRef = articleEntries[articleKey];
          if (!entryRef || !entryRef.rowNumber || !entryRef.rowArray) {
            continue;
          }
          var idpValueForRow = entryRef.rowArray[idpColumnIndex];
          if (
            idpValueForRow === undefined ||
            idpValueForRow === null ||
            idpValueForRow === ""
          ) {
            continue;
          }
          Lib.logDebug(
            "[SS] syncIdWithMain: запись ID-P " +
              idpValueForRow +
              " в строку " +
              entryRef.rowNumber +
              " (артикул " +
              articleKey +
              ")"
          );
          mainSheet
            .getRange(entryRef.rowNumber, idpColumnIndex + 1)
            .setValue(idpValueForRow);
        }
      }

      Lib.logDebug(
        "[SS] syncIdWithMain: итоговый счётчик ID-P = " + idpCounter
      );

      // Собираем все изменения групп для диалога
      var groupChanges = [];
      for (var articleKey in articleEntries) {
        if (!Object.prototype.hasOwnProperty.call(articleEntries, articleKey))
          continue;
        var entryRef = articleEntries[articleKey];
        if (entryRef && entryRef.__groupChange) {
          groupChanges.push(entryRef.__groupChange);
        }
      }

      Lib.logDebug(
        "[SS] syncIdWithMain: завершено, создано новых строк: " +
          createdEntries.length +
          ", изменений групп: " +
          groupChanges.length
      );
      return {
        createdRows: createdEntries,
        assignedIdp: assignedIdpMap,
        groupChanges: groupChanges,
      };
    } catch (syncError) {
      Lib.logError("syncIdWithMain: ошибка", syncError);
      return emptyResult;
    }
  }

  function _handleGroupChanges_(groupChanges) {
    if (!groupChanges || groupChanges.length === 0) {
      return;
    }

    Lib.logInfo(
      "[SS] _handleGroupChanges_: обнаружено " +
        groupChanges.length +
        " изменений в столбце Группа"
    );

    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var mainSheetName =
        global.CONFIG && global.CONFIG.SHEETS && global.CONFIG.SHEETS.PRIMARY;
      var mainSheet = mainSheetName ? ss.getSheetByName(mainSheetName) : null;
      if (!mainSheet) {
        Lib.logError("[SS] Не найден лист Главная");
        return;
      }

      var lastColumn = mainSheet.getLastColumn();
      if (lastColumn <= 0) {
        Lib.logError("[SS] Нет столбцов на листе Главная");
        return;
      }

      var headersRow = mainSheet
        .getRange(1, 1, 1, lastColumn)
        .getValues()[0]
        .map(function (value) {
          return String(value || "").trim();
        });

      var headerToIndex = {};
      headersRow.forEach(function (name, idx) {
        var normalized = _normalizeHeaderKey(name);
        if (normalized) {
          headerToIndex[normalized] = idx;
        }
      });

      var groupIdx = headerToIndex[_normalizeHeaderKey("Группа")];
      if (groupIdx === undefined || groupIdx === -1) {
        Lib.logError("[SS] Не найден столбец Группа");
        return;
      }

      var idgIdx = headerToIndex[_normalizeHeaderKey("ID-G")];

      var ui = SpreadsheetApp.getUi();

      // Показываем диалог для каждого изменения
      groupChanges.forEach(function (change) {
        if (!change || !change.rowNumber) {
          return;
        }

        var message =
          'Артикул: "' +
          change.articleCode +
          '"\n\n' +
          "Текущая Группа: " +
          change.oldValue +
          "\n" +
          "Новая Группа в прайс-листе: " +
          change.newValue +
          "\n\n" +
          'Нажмите "Да" - заменить на новое значение\n' +
          'Нажмите "Нет" - оставить текущее значение';

        var result = ui.alert(
          "Изменение группы",
          message,
          ui.ButtonSet.YES_NO
        );

        if (result === ui.Button.YES) {
          // Заменяем значение Группы
          mainSheet
            .getRange(change.rowNumber, groupIdx + 1)
            .setValue(change.newValue);

          Lib.logInfo(
            "[SS] Группа заменена для артикула " +
              change.articleCode +
              ": '" +
              change.oldValue +
              "' → '" +
              change.newValue +
              "'"
          );

          // Если есть столбец ID-G, очищаем его для пересчёта
          if (idgIdx !== undefined && idgIdx !== -1) {
            mainSheet.getRange(change.rowNumber, idgIdx + 1).clearContent();
            Lib.logDebug(
              "[SS] Очищен ID-G для строки " +
                change.rowNumber +
                " (группа изменена)"
            );
          }
        } else {
          Lib.logInfo(
            "[SS] Группа оставлена без изменений для артикула " +
              change.articleCode
          );
        }
      });

      SpreadsheetApp.flush();
    } catch (err) {
      Lib.logError("_handleGroupChanges_: ошибка", err);
    }
  }

  function _handleNewArticles_(createdEntries, sourceRows, sourceHeaders) {
    if (!createdEntries || !createdEntries.length) {
      return [];
    }

    Lib.logInfo("[SS] _handleNewArticles_: обработка " + createdEntries.length + " новых артикулов");
    Lib.logDebug("[SS] sourceRows: " + (sourceRows ? sourceRows.length : 0) + ", sourceHeaders: " + (sourceHeaders ? sourceHeaders.length : 0));

    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var mainSheetName =
        global.CONFIG && global.CONFIG.SHEETS && global.CONFIG.SHEETS.PRIMARY;
      var mainSheet = mainSheetName ? ss.getSheetByName(mainSheetName) : null;
      if (!mainSheet) {
        Lib.logError("[SS] Не найден лист Главная");
        return [];
      }

      var lastColumn = mainSheet.getLastColumn();
      if (lastColumn <= 0) {
        Lib.logError("[SS] Нет столбцов на листе Главная");
        return [];
      }

      var headersRow = mainSheet
        .getRange(1, 1, 1, lastColumn)
        .getValues()[0]
        .map(function (value) {
          return String(value || "").trim();
        });

      // Строим индекс по НОРМАЛИЗОВАННЫМ ключам
      var headerToIndex = {};
      headersRow.forEach(function (name, idx) {
        var normalized = _normalizeHeaderKey(name);
        if (normalized) {
          headerToIndex[normalized] = idx;
        }
      });

      var statusIdx = headerToIndex[_normalizeHeaderKey("Статус")];
      if (statusIdx === undefined) statusIdx = -1;
      var volumeIdx = headerToIndex[_normalizeHeaderKey("Объём")];
      if (volumeIdx === undefined) volumeIdx = -1;
      var groupIdx = headerToIndex[_normalizeHeaderKey("Группа")];
      if (groupIdx === undefined) groupIdx = -1;
      var unitsIdx = headerToIndex[_normalizeHeaderKey("шт./уп.")];
      if (unitsIdx === undefined) unitsIdx = -1;
      var formIdx = headerToIndex[_normalizeHeaderKey("Форма выпуска")];
      if (formIdx === undefined) formIdx = -1;
      var idIdx = headerToIndex[_normalizeHeaderKey("ID")];
      if (idIdx === undefined) idIdx = -1;

      Lib.logDebug("[SS] _handleNewArticles_: statusIdx=" + statusIdx + ", volumeIdx=" + volumeIdx + ", groupIdx=" + groupIdx + ", unitsIdx=" + unitsIdx + ", formIdx=" + formIdx);

      var processedRows = [];

      createdEntries.forEach(function (entry) {
        if (!entry || !entry.rowNumber) {
          return;
        }

        var productName =
          entry.productName || entry.articleCode || "Новый продукт";

        Lib.logInfo("[SS] Обработка нового артикула: " + productName + ", строка=" + entry.rowNumber);

        // Получаем группу из entry или из таблицы
        var groupValue = entry.group || "";
        if (!groupValue && groupIdx !== -1) {
          groupValue = _asTrimmedString(
            mainSheet.getRange(entry.rowNumber, groupIdx + 1).getValue()
          );
        }

        // НОВАЯ ЛОГИКА: автоматически ставим статус "New завод"
        if (statusIdx !== undefined && statusIdx !== -1) {
          Lib.logDebug("[SS] Установка статуса 'New завод' для строки " + entry.rowNumber);
          mainSheet
            .getRange(entry.rowNumber, statusIdx + 1)
            .setValue("New завод");
        } else {
          Lib.logWarn("[SS] Не удалось найти столбец 'Статус' для строки " + entry.rowNumber);
        }

        // Запрашиваем только Объём
        var volumeValue = _promptVolumeInput_(productName);
        if (
          volumeValue !== null &&
          volumeIdx !== undefined &&
          volumeIdx !== -1
        ) {
          Lib.logDebug("[SS] Установка объёма '" + volumeValue + "' для строки " + entry.rowNumber);
          mainSheet
            .getRange(entry.rowNumber, volumeIdx + 1)
            .setValue(volumeValue);
        }

        // Логика для SAMPLES: извлекаем число из "Форма выпуска" и записываем в "шт./уп."
        if (groupValue && groupValue.toUpperCase().indexOf("SAMPLES") !== -1) {
          if (formIdx !== -1 && unitsIdx !== -1) {
            var formValue = _asTrimmedString(
              mainSheet.getRange(entry.rowNumber, formIdx + 1).getValue()
            );
            if (formValue) {
              // Извлекаем первое число из строки (например, "24 sachets" -> 24)
              var numberMatch = formValue.match(/\d+/);
              if (numberMatch) {
                var extractedNumber = parseInt(numberMatch[0], 10);
                Lib.logDebug("[SS] SAMPLES: извлечено число " + extractedNumber + " из '" + formValue + "' для строки " + entry.rowNumber);
                mainSheet
                  .getRange(entry.rowNumber, unitsIdx + 1)
                  .setValue(extractedNumber);
              }
            }
          }
        }

        // Получаем ID для создания строк на базовых листах
        var primaryId = "";
        if (idIdx !== -1) {
          primaryId = _asTrimmedString(
            mainSheet.getRange(entry.rowNumber, idIdx + 1).getValue()
          );
        }

        // Применяем изменения
        SpreadsheetApp.flush();

        // Создаём строки на базовых листах
        if (primaryId && typeof Lib.ensureRowExistsOnBaseSheets === "function") {
          Lib.logInfo("[SS] Создание строк на базовых листах для " + primaryId);
          Lib.ensureRowExistsOnBaseSheets(primaryId);
        }

        processedRows.push(entry.rowNumber);
      });

      // Остальные поля (Название ENG, шт./уп., Группа, Арт. произв.) уже заполнены из прайс-листа

      SpreadsheetApp.flush();
      return processedRows;
    } catch (err) {
      Lib.logError("_handleNewArticles_: ошибка", err);
      return [];
    }
  }

  function _promptVolumeInput_(productName) {
    var ui = SpreadsheetApp.getUi();
    var message = 'Объём "' + productName + '"';

    while (true) {
      var prompt = ui.prompt(
        "Объём продукта",
        message,
        ui.ButtonSet.OK_CANCEL
      );
      var button = prompt.getSelectedButton();
      if (button !== ui.Button.OK) {
        return null;
      }
      var text = _asTrimmedString(prompt.getResponseText());
      if (text) {
        return text;
      }
      ui.alert(
        "Укажите значение объёма или нажмите «Отмена», чтобы пропустить."
      );
    }
  }

  /**
   * Устанавливает цену из обработанных данных (processed) в листы Динамика цены и Расчет цены
   * для конкретного артикула
   * Ищет артикул по "Арт. произв." в processed и берёт цену оттуда
   * @param {string} article - Артикул производителя
   * @param {Object} processed - Обработанные данные из исходного документа
   * @param {Object<string, number>} [priceMap] - Предрассчитанная карта цен (артикул -> цена)
   */
  function _setPriceFromProcessedByArticle_(article, processed, priceMap) {
    if (!article || !processed || !processed.headers || !processed.rows) {
      return;
    }

    try {
      // Находим индексы столбцов в processed
      var articleIndex = processed.headers.indexOf('Арт. произв.');
      var priceIndex = processed.headers.indexOf('Цена');

      if (articleIndex === -1 || priceIndex === -1) {
        Lib.logWarn('[SS] _setPriceFromProcessedByArticle_: не найдены необходимые столбцы в processed');
        return;
      }

      // Ищем строку с нужным артикулом
      var price = null;
      if (priceMap && Object.prototype.hasOwnProperty.call(priceMap, article)) {
        price = priceMap[article];
      }

      if (price === null || price === undefined) {
        for (var i = 0; i < processed.rows.length; i++) {
          var row = processed.rows[i];
          var rowArticle = String(row[articleIndex] || "").trim();
          if (rowArticle === article) {
            var priceValue = row[priceIndex];
            if (priceValue !== "" && priceValue !== null && priceValue !== undefined) {
              var priceNum;
              if (typeof priceValue === "number") {
                priceNum = priceValue;
              } else {
                var cleaned = String(priceValue).replace(/\s+/g, "").replace(",", ".");
                priceNum = parseFloat(cleaned);
              }
              if (!isNaN(priceNum)) {
                price = priceNum;
                break;
              }
            }
          }
        }
      }

      if (price === null) {
        Lib.logDebug('[SS] _setPriceFromProcessedByArticle_: цена не найдена для артикула ' + article);
        return;
      }

      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var targets = [
        Lib.CONFIG.SHEETS.PRICE_DYNAMICS,    // Динамика цены
        Lib.CONFIG.SHEETS.PRICE_CALCULATION, // Расчет цены
      ];

      targets.forEach(function(sheetName) {
        var sheet = ss.getSheetByName(sheetName);
        if (!sheet) {
          return;
        }

        var lastRow = sheet.getLastRow();
        if (lastRow <= 1) {
          return;
        }

        var lastColumn = sheet.getLastColumn();
        if (lastColumn <= 0) {
          return;
        }

        // Читаем заголовки
        var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0]
          .map(function(h) { return String(h || '').trim(); });

        var articleColIndex = headers.indexOf('Арт. произв.');
        var priceColIndex = headers.indexOf('ЦЕНА EXW из Б/З');

        if (articleColIndex === -1 || priceColIndex === -1) {
          return;
        }

        // Ищем строку с нужным артикулом
        var articleData = sheet.getRange(2, articleColIndex + 1, lastRow - 1, 1).getValues();
        var targetRow = -1;
        for (var i = 0; i < articleData.length; i++) {
          if (String(articleData[i][0] || '').trim() === article) {
            targetRow = i + 2;
            break;
          }
        }

        if (targetRow === -1) {
          return;
        }

        // Записываем цену с форматом "0.00 €"
        var range = sheet.getRange(targetRow, priceColIndex + 1);
        range.setValue(price);
        range.setNumberFormat('0.00 "€"');
        Lib.logInfo('[SS] Установлена цена ' + price + ' € из исходного документа в "ЦЕНА EXW из Б/З" на листе "' + sheetName + '" для артикула ' + article);
      });
    } catch (err) {
      Lib.logError('[SS] _setPriceFromProcessedByArticle_: ошибка', err);
    }
  }

  function _applyAssignedIdpToProcessed_(processed, assignedMap) {
    if (!processed || !processed.rows || !assignedMap) {
      return;
    }
    var headers = processed.headers || [];
    var idpIndex = headers.indexOf("ID-P");
    var articleIndex = headers.indexOf("Арт. произв.");

    var fallback = false;
    if (idpIndex === -1) {
      idpIndex = 0;
      fallback = true;
    }
    if (articleIndex === -1) {
      articleIndex = 1;
    }

    for (var i = 0; i < processed.rows.length; i++) {
      var row = processed.rows[i];
      if (!row || row.length === 0) {
        continue;
      }
      var article = row[articleIndex];
      var key = _asTrimmedString(article);
      if (!key) {
        continue;
      }
      if (Object.prototype.hasOwnProperty.call(assignedMap, key)) {
        row[idpIndex] = assignedMap[key];
      }
    }
  }

  function _clearIdpColumnOnMain() {
    try {
      var mainSheetName =
        global.CONFIG && global.CONFIG.SHEETS && global.CONFIG.SHEETS.PRIMARY;
      if (!mainSheetName) {
        return;
      }
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) {
        return;
      }
      var sheet = ss.getSheetByName(mainSheetName);
      if (!sheet) {
        return;
      }

      var lastRow = sheet.getLastRow();
      if (lastRow <= 1) {
        return;
      }

      var lastColumn = sheet.getLastColumn();
      if (lastColumn <= 0) {
        return;
      }

      var headers = sheet
        .getRange(1, 1, 1, lastColumn)
        .getValues()[0]
        .map(function (value) {
          return String(value || "").trim();
        });
      var idpIndex = headers.indexOf("ID-P");
      if (idpIndex === -1) {
        return;
      }

      sheet.getRange(2, idpIndex + 1, lastRow - 1, 1).clearContent();
    } catch (err) {
      Lib.logError("_clearIdpColumnOnMain: ошибка", err);
    }
  }

  function _clearIdgColumnOnMain() {
    try {
      var mainSheetName =
        global.CONFIG && global.CONFIG.SHEETS && global.CONFIG.SHEETS.PRIMARY;
      if (!mainSheetName) {
        return;
      }
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) {
        return;
      }
      var sheet = ss.getSheetByName(mainSheetName);
      if (!sheet) {
        return;
      }

      var lastRow = sheet.getLastRow();
      if (lastRow <= 1) {
        return;
      }

      var lastColumn = sheet.getLastColumn();
      if (lastColumn <= 0) {
        return;
      }

      var headers = sheet
        .getRange(1, 1, 1, lastColumn)
        .getValues()[0]
        .map(function (value) {
          return String(value || "").trim();
        });
      var idgIndex = headers.indexOf("ID-G");
      if (idgIndex === -1) {
        return;
      }

      sheet.getRange(2, idgIndex + 1, lastRow - 1, 1).clearContent();
      Lib.logDebug("[SS] _clearIdgColumnOnMain: очищен столбец ID-G");
    } catch (err) {
      Lib.logError("_clearIdgColumnOnMain: ошибка", err);
    }
  }

  function _fillIdgForRowsWithoutIdp() {
    try {
      var mainSheetName =
        global.CONFIG && global.CONFIG.SHEETS && global.CONFIG.SHEETS.PRIMARY;
      if (!mainSheetName) {
        return;
      }
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) {
        return;
      }
      var sheet = ss.getSheetByName(mainSheetName);
      if (!sheet) {
        return;
      }

      var lastRow = sheet.getLastRow();
      if (lastRow <= 1) {
        return;
      }

      var lastColumn = sheet.getLastColumn();
      if (lastColumn <= 0) {
        return;
      }

      var headers = sheet
        .getRange(1, 1, 1, lastColumn)
        .getValues()[0]
        .map(function (value) {
          return String(value || "").trim();
        });

      var idpIndex = headers.indexOf("ID-P");
      var idgIndex = headers.indexOf("ID-G");
      var groupIndex = headers.indexOf("Группа");

      if (idpIndex === -1 || idgIndex === -1 || groupIndex === -1) {
        Lib.logDebug(
          "[SS] _fillIdgForRowsWithoutIdp: не найдены необходимые столбцы (ID-P=" +
            idpIndex +
            ", ID-G=" +
            idgIndex +
            ", Группа=" +
            groupIndex +
            ")"
        );
        return;
      }

      var dataRange = sheet.getRange(2, 1, lastRow - 1, lastColumn);
      var values = dataRange.getValues();

      // Сначала собираем карту Группа -> ID-G из строк с ID-P
      var groupToIdg = Object.create(null);
      for (var i = 0; i < values.length; i++) {
        var row = values[i];
        var idpValue = row[idpIndex];
        var idgValue = row[idgIndex];
        var groupValue = _asTrimmedString(row[groupIndex]);

        if (idpValue && idgValue && groupValue) {
          if (!groupToIdg[groupValue]) {
            groupToIdg[groupValue] = idgValue;
          }
        }
      }

      // Теперь проставляем ID-G для строк без ID-P по совпадению Группы
      var updatesCount = 0;
      for (var j = 0; j < values.length; j++) {
        var rowData = values[j];
        var idpVal = rowData[idpIndex];
        var groupVal = _asTrimmedString(rowData[groupIndex]);

        if (!idpVal && groupVal && groupToIdg[groupVal]) {
          sheet.getRange(j + 2, idgIndex + 1).setValue(groupToIdg[groupVal]);
          updatesCount++;
        }
      }

      if (updatesCount > 0) {
        Lib.logInfo(
          "[SS] _fillIdgForRowsWithoutIdp: проставлено " +
            updatesCount +
            " значений ID-G"
        );
      }
    } catch (err) {
      Lib.logError("_fillIdgForRowsWithoutIdp: ошибка", err);
    }
  }

  function _normalizeHeaderKey(name) {
    if (name === null || name === undefined) {
      return "";
    }
    return String(name)
      .trim()
      .toLowerCase()
      .replace(/[^0-9a-zа-яё]/gi, "");
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
    if (value instanceof Date && !isNaN(value.getTime())) {
      var date = value;
      var label = Utilities.formatDate(
        date,
        Lib.CONFIG.SETTINGS.TIMEZONE || "Europe/Moscow",
        "dd.MM.yyyy"
      );
      return { label: label, date: date };
    }
    if (value === null || value === undefined) {
      return { label: "", date: null };
    }
    var str = String(value).trim();
    if (!str) {
      return { label: "", date: null };
    }
    var parsed = new Date(str);
    if (!isNaN(parsed.getTime())) {
      var labelParsed = Utilities.formatDate(
        parsed,
        Lib.CONFIG.SETTINGS.TIMEZONE || "Europe/Moscow",
        "dd.MM.yyyy"
      );
      return { label: labelParsed, date: parsed };
    }
    return { label: str, date: null };
  }

  function _getValue(row, index) {
    if (!row || index >= row.length) {
      return "";
    }
    var value = row[index];
    if (value === null || value === undefined) {
      return "";
    }
    if (typeof value === "string") {
      return value.trim();
    }
    return value;
  }

  function _asTrimmedString(value) {
    if (value === null || value === undefined) {
      return "";
    }
    return String(value).trim();
  }

  function _isEmpty(value) {
    if (value === null || value === undefined) {
      return true;
    }
    if (typeof value === "string") {
      return value.trim() === "";
    }
    return false;
  }
})(Lib, this);
