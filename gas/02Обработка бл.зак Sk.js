var Lib = Lib || {};

(function (Lib, global) {
  var TARGET_PROJECT_KEY = "SK";

  // Заголовки для основного прайса (Б/З поставщик)
  var MAIN_OUTPUT_HEADERS = [
    "ID-P",
    "Арт. произв.",
    "Название ENG прайс произв",
    "шт./уп.",
    "Цена",
    "всего",
    "RRP",
    "Группа",
  ];

  // Заголовки для пробников
  var PROBES_OUTPUT_HEADERS = [
    "ID-P",
    "Арт. произв.",
    "Название ENG прайс произв",
    "шт./уп.",
    "кол-во",
    "Цена",
    "Всего",
    "Группа",
  ];

  Lib.processSkPriceSheet = function () {
    var ui = SpreadsheetApp.getUi();
    var config = _getPrimaryDataConfig_();
    var menuTitle = _getMenuTitle_(config);

    if (!_isActiveProject_()) {
      ui.alert(
        menuTitle,
        "Эта функция доступна только в проекте SK.",
        ui.ButtonSet.OK
      );
      return;
    }

    try {
      Lib.logInfo("[SK] Обработка основной: старт");
      var source = _getSourceData_(config, "MAIN");
      if (!source.values || !source.values.length) {
        throw new Error("В исходном документе нет данных для обработки.");
      }

      var processed = _buildProcessedMainData_(
        source.values,
        source.backgrounds,
        config
      );
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

      // Проставляем ID-G для строк без ID-P по совпадению Группа
      _fillIdgForRowsWithoutIdp();

      // Заполняем ID-P на остальных листах (Заказ, Динамика цены, Расчет цены, ABC-Анализ)
      if (typeof Lib.fillIdpOnSheetsByIdFromPrimary === 'function') {
        Lib.logInfo("[SK] Заполнение ID-P на всех листах");
        Lib.fillIdpOnSheetsByIdFromPrimary();
      } else {
        Lib.logWarn("[SK] fillIdpOnSheetsByIdFromPrimary не найдена");
      }

      // Копируем цену из Б/З в столбец "ЦЕНА EXW из Б/З" на листах Динамика цены и Расчет цены
      if (typeof Lib.copyPriceFromPrimaryToSheets === 'function') {
        Lib.logInfo("[SK] Копирование цены из Б/З на листы Динамика цены и Расчет цены");
        Lib.copyPriceFromPrimaryToSheets(processed);

        // Заполняем скидку 20% на листе "Динамика цены" где есть цена
        Lib.logInfo("[SK] Заполнение скидки 20% на листе Динамика цены");
        _fillDiscountForPriceDynamics();
      } else {
        Lib.logWarn("[SK] copyPriceFromPrimaryToSheets не найдена");
      }

      // Применяем формулы на листе "Динамика цены" (EXW ALFASPA, Закупочная цена, DDP-МОСКВА, Прирост)
      if (typeof Lib.recalculatePriceDynamicsFormulas === 'function') {
        Lib.logInfo("[SK] Применение формул на листе Динамика цены");
        Lib.recalculatePriceDynamicsFormulas();
      } else {
        Lib.logWarn("[SK] recalculatePriceDynamicsFormulas не найдена");
      }

      if (
        syncResult &&
        syncResult.createdRows &&
        syncResult.createdRows.length
      ) {
        var newArticleRows = _handleNewArticles_(syncResult.createdRows);

        // Запускаем построчную синхронизацию для новых артикулов со статусом "New завод"
        if (newArticleRows && newArticleRows.length > 0) {
          Lib.logInfo(
            "[SK] Запуск построчной синхронизации для " +
              newArticleRows.length +
              " новых артикулов"
          );
          SpreadsheetApp.flush();

          if (typeof Lib.syncMultipleRows === "function") {
            Lib.syncMultipleRows(newArticleRows);
          } else {
            Lib.logWarn("[SK] syncMultipleRows не найдена");
          }

          // Заполняем ID-P для новых строк после синхронизации
          if (typeof Lib.fillIdpOnSheetsByIdFromPrimary === 'function') {
            Lib.logInfo("[SK] Заполнение ID-P для новых артикулов на всех листах");
            Lib.fillIdpOnSheetsByIdFromPrimary();
          }

          // Заполняем ЦЕНА EXW из Б/З для новых артикулов
          if (typeof Lib.copyPriceFromPrimaryToSheets === 'function') {
            Lib.logInfo("[SK] Копирование цен для новых артикулов на Динамика цены и Расчет цены");
            Lib.copyPriceFromPrimaryToSheets(processed);

            // Заполняем скидку 20% для новых артикулов
            Lib.logInfo("[SK] Заполнение скидки 20% для новых артикулов");
            _fillDiscountForPriceDynamics();
          }
        }
      }

      Lib.logInfo(
        "[SK] Обработка основной: завершено, строк " + processed.rows.length
      );

      // Обновляем формулы на листе "Расчет цены"
      // 1. Сначала применяем INDEX/MATCH формулы для подтягивания данных из "Динамика цены"
      if (typeof Lib.updatePriceCalculationFormulas === 'function') {
        Lib.logInfo("[SK] Обновление INDEX/MATCH формул на листе Расчет цены");
        Lib.updatePriceCalculationFormulas(true); // silent=true
      } else {
        Lib.logWarn("[SK] updatePriceCalculationFormulas не найдена");
      }

      // 2. Затем применяем расчетные формулы (К-т, Расчетная цена Опт, РРЦ и т.д.)
      if (typeof Lib.applyCalculationFormulas === 'function') {
        Lib.logInfo("[SK] Применение расчетных формул на листе Расчет цены");
        Lib.applyCalculationFormulas(true); // silent=true
      } else {
        Lib.logWarn("[SK] applyCalculationFormulas не найдена");
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
      Lib.logError("processSkPriceSheet: ошибка", error);
      ui.alert(
        "Ошибка обработки прайс",
        error.message || String(error),
        ui.ButtonSet.OK
      );
    }
  };

  Lib.processSkPriceProbes = function () {
    var ui = SpreadsheetApp.getUi();
    var config = _getPrimaryDataConfig_();
    var menuTitle = _getMenuTitle_(config);

    if (!_isActiveProject_()) {
      ui.alert(
        menuTitle,
        "Эта функция доступна только в проекте SK.",
        ui.ButtonSet.OK
      );
      return;
    }

    try {
      Lib.logInfo("[SK] Обработка пробники: старт");
      var source = _getSourceData_(config, "PROBES");
      if (!source.values || source.values.length <= 1) {
        ui.alert(
          menuTitle,
          "На листе нет данных для обработки.",
          ui.ButtonSet.OK
        );
        return;
      }

      var processed = _buildProcessedProbesData_(source.values, config);
      if (!processed.rows.length) {
        ui.alert(
          menuTitle,
          "Не найдено данных для обработки.",
          ui.ButtonSet.OK
        );
        return;
      }

      // НЕ очищаем ID-P и цены - пробники добавляются к основным данным
      // ID-P будет продолжать нумерацию от максимального значения

      _assignSequentialIdpForProbes_(processed);

      var syncResult = _syncIdWithMain(
        processed.rows,
        processed.headers,
        processed.articles,
        processed.groups,
        {
          overrideIdpFromSource: true,
          samplesGroupKey: "Пробники",
        }
      );

      _applyAssignedIdpToProcessed_(
        processed,
        syncResult && syncResult.assignedIdp
      );
      _storeProcessedSnapshot_(config, processed, "PROBES");

      // Проставляем ID-G для строк без ID-P по совпадению Группа
      _fillIdgForRowsWithoutIdp();

      // Заполняем ID-P на остальных листах (Заказ, Динамика цены, Расчет цены, ABC-Анализ)
      if (typeof Lib.fillIdpOnSheetsByIdFromPrimary === 'function') {
        Lib.logInfo("[SK] (Пробники) Заполнение ID-P на всех листах");
        Lib.fillIdpOnSheetsByIdFromPrimary();
      } else {
        Lib.logWarn("[SK] (Пробники) fillIdpOnSheetsByIdFromPrimary не найдена");
      }

      // Копируем цену из Б/З в столбец "ЦЕНА EXW из Б/З" на листах Динамика цены и Расчет цены
      if (typeof Lib.copyPriceFromPrimaryToSheets === 'function') {
        Lib.logInfo("[SK] (Пробники) Копирование цены из Б/З на листы Динамика цены и Расчет цены");
        Lib.copyPriceFromPrimaryToSheets(processed);
      } else {
        Lib.logWarn("[SK] (Пробники) copyPriceFromPrimaryToSheets не найдена");
      }

      // Применяем формулы на листе "Динамика цены" (EXW ALFASPA, Закупочная цена, DDP-МОСКВА, Прирост)
      if (typeof Lib.recalculatePriceDynamicsFormulas === 'function') {
        Lib.logInfo("[SK] (Пробники) Применение формул на листе Динамика цены");
        Lib.recalculatePriceDynamicsFormulas();
      } else {
        Lib.logWarn("[SK] (Пробники) recalculatePriceDynamicsFormulas не найдена");
      }

      if (
        syncResult &&
        syncResult.createdRows &&
        syncResult.createdRows.length
      ) {
        var newArticleRows2 = _handleNewArticles_(syncResult.createdRows);

        // Запускаем построчную синхронизацию для новых артикулов со статусом "New завод"
        if (newArticleRows2 && newArticleRows2.length > 0) {
          Lib.logInfo(
            "[SK] (Пробники) Постаточно синхро для новых артикулов: " +
              newArticleRows2.length
          );
          SpreadsheetApp.flush();
          if (typeof Lib.syncMultipleRows === "function") {
            Lib.syncMultipleRows(newArticleRows2);
          } else {
            Lib.logWarn("[SK] syncMultipleRows не найдена");
          }

          // Заполняем ID-P для новых строк после синхронизации
          if (typeof Lib.fillIdpOnSheetsByIdFromPrimary === 'function') {
            Lib.logInfo("[SK] (Пробники) Заполнение ID-P для новых артикулов на всех листах");
            Lib.fillIdpOnSheetsByIdFromPrimary();
          }
        }
      }

      Lib.logInfo(
        "[SK] Обработка пробники: завершено, строк " + processed.rows.length
      );

      // Обновляем формулы на листе "Расчет цены"
      // 1. Сначала применяем INDEX/MATCH формулы для подтягивания данных из "Динамика цены"
      if (typeof Lib.updatePriceCalculationFormulas === 'function') {
        Lib.logInfo("[SK] (Пробники) Обновление INDEX/MATCH формул на листе Расчет цены");
        Lib.updatePriceCalculationFormulas(true); // silent=true
      } else {
        Lib.logWarn("[SK] (Пробники) updatePriceCalculationFormulas не найдена");
      }

      // 2. Затем применяем расчетные формулы (К-т, Расчетная цена Опт, РРЦ и т.д.)
      if (typeof Lib.applyCalculationFormulas === 'function') {
        Lib.logInfo("[SK] (Пробники) Применение расчетных формул на листе Расчет цены");
        Lib.applyCalculationFormulas(true); // silent=true
      } else {
        Lib.logWarn("[SK] (Пробники) applyCalculationFormulas не найдена");
      }

      // Обновляем статусы и синхронизируем с ТЗ по статусам
      if (typeof Lib.updateStatusesAfterProcessing === 'function') {
        Lib.updateStatusesAfterProcessing();
      }

      ui.alert(
        menuTitle,
        "Обработка прайс завершена. Обновлено строк: " +
          processed.rows.length +
          ".",
        ui.ButtonSet.OK
      );
    } catch (error) {
      Lib.logError("processSkPriceProbes: ошибка", error);
      ui.alert(
        "Ошибка обработки прайс",
        error.message || String(error),
        ui.ButtonSet.OK
      );
    }
  };

  Lib.loadSkStockData = function () {
    var ui = SpreadsheetApp.getUi();
    var config = _getPrimaryDataConfig_();
    var menuTitle = _getMenuTitle_(config);

    if (!_isActiveProject_()) {
      ui.alert(
        menuTitle,
        "Эта функция доступна только в проекте SK.",
        ui.ButtonSet.OK
      );
      return;
    }

    // Используем универсальную функцию загрузки остатков
    if (typeof Lib.loadStockData === 'function') {
      Lib.loadStockData('SK');
      return;
    }

    // DEPRECATED: Старая реализация (будет удалена)
    try {
      Lib.logInfo("[SK] Загрузка остатков: старт (DEPRECATED - используйте Lib.loadStockData)");

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
          '[SK] Загрузка остатков: данные читаются из внешнего документа ' +
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
          Lib.logInfo('[SK] Производные столбцы пересчитаны после загрузки остатков');
        }
      } catch (e) {
        Lib.logWarn('[SK] Не удалось пересчитать производные столбцы заказа после загрузки остатков', e);
      }

      Lib.logInfo("[SK] Загрузка остатков: обновлено строк " + numRows);
      ui.alert(
        menuTitle,
        "Загрузка остатков завершена. Обновлено строк: " + numRows + ".",
        ui.ButtonSet.OK
      );
    } catch (error) {
      Lib.logError("loadSkStockData: ошибка", error);
      ui.alert(
        "Ошибка загрузки остатков",
        error.message || String(error),
        ui.ButtonSet.OK
      );
    }
  };

  /**
   * УСТАРЕВШАЯ: Сортировка заказа (использует группировку по умолчанию)
   * @deprecated Используйте sortSkOrderByManufacturer или sortSkOrderByPrice
   */
  Lib.sortSkOrderData = function () {
    Lib.sortSkOrderByManufacturer();
  };

  /**
   * ПУБЛИЧНАЯ: Сортировка заказа по Производителю (Группа + ID-G)
   * Теперь сортирует одновременно три листа: Заказ, Динамика цены, Расчет цены
   */
  Lib.sortSkOrderByManufacturer = function () {
    var ui = SpreadsheetApp.getUi();
    var config = _getPrimaryDataConfig_();
    var menuTitle = _getMenuTitle_(config);

    if (!_isActiveProject_()) {
      ui.alert(
        menuTitle,
        "Эта функция доступна только в проекте SK.",
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
      Lib.logError("[SK] Сортировка по Производителю: ошибка", error);
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
  Lib.sortSkOrderByPrice = function () {
    var ui = SpreadsheetApp.getUi();
    var config = _getPrimaryDataConfig_();
    var menuTitle = _getMenuTitle_(config);

    if (!_isActiveProject_()) {
      ui.alert(
        menuTitle,
        "Эта функция доступна только в проекте SK.",
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
      Lib.logError("[SK] Сортировка по Прайсу: ошибка", error);
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
      throw new Error("Конфигурация PRIMARY_DATA не настроена для проекта SK.");
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
    // Создает новый лист с именем sheetName. Не удаляет другие листы.
    // Предполагается, что удаление старых датированных листов выполняется отдельно.
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

    // Сначала получаем данные с исходного листа
    var originalSheet = sourceSpreadsheet.getSheetByName(baseSheetName);
    if (!originalSheet) {
      throw new Error(
        'Не найден лист "' + baseSheetName + '" в документе источнике.'
      );
    }

    // Получаем данные с исходного листа
    var originalRange = originalSheet.getDataRange();
    var originalValues = originalRange.getValues();
    var originalBackgrounds = originalRange.getBackgrounds();

    // Формируем имя листа с датой
    var timezone =
      (Lib.CONFIG && Lib.CONFIG.SETTINGS && Lib.CONFIG.SETTINGS.TIMEZONE) ||
      "Europe/Moscow";
    var dateString = Utilities.formatDate(new Date(), timezone, "dd.MM.yy");
    var fullSheetName = baseSheetName + " " + dateString;

    // Удаляем старые датированные листы с тем же префиксом (например "-пробники ")
    try {
      var allSheets = sourceSpreadsheet.getSheets();
      for (var si = 0; si < allSheets.length; si++) {
        var sh = allSheets[si];
        var n = sh.getName();
        if (n.indexOf(baseSheetName + " ") === 0 && n !== fullSheetName) {
          // удалить старую датированную версию
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

    // Пересоздаем лист с текущей датой
    var sheet = recreateSheet(sourceSpreadsheet, fullSheetName);
    if (!sheet) {
      throw new Error(
        'Не удалось создать лист "' + fullSheetName + '" в документе источнике.'
      );
    }

    // Копируем данные на новый лист
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

      // Удаляем все валидации данных (выпадающие списки) на новом листе
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

  function _buildProcessedMainData_(values, backgrounds, config) {
    var rows = [];
    var articles = [];
    var lines = [];
    var groups = [];
    var groupColor = (config.GROUP_HEADER_COLOR || "").toLowerCase();
    var currentGroup = "";
    var isDataStarted = false;
    var lastLineValue = "";

    for (var rowIndex = 0; rowIndex < values.length; rowIndex++) {
      var row = values[rowIndex];
      var bgRow = backgrounds[rowIndex] || [];
      var groupName = _detectGroupRow(row, bgRow, groupColor);
      if (groupName) {
        currentGroup = groupName;
        isDataStarted = true;
        lastLineValue = "";
        continue;
      }

      if (!isDataStarted) {
        continue;
      }

      var articleIndex = 0;
      var productCode = _asTrimmedString(row[articleIndex]);
      if (!productCode) {
        articleIndex = 1;
        productCode = _asTrimmedString(row[articleIndex]);
      }

      if (!productCode) {
        continue;
      }

      var base = articleIndex;
      var nameValue = _getValue(row, base + 1);
      if (!nameValue) {
        continue;
      }
      var lineValue = _getValue(row, base + 2);
      if (!lineValue) {
        lineValue = lastLineValue;
      } else {
        lastLineValue = lineValue;
      }
      var unitsValue = _getValue(row, base + 3);
      var masterBoxValue = _getValue(row, base + 4);
      var priceValue = _getValue(row, base + 5);
      var totalValue = _getValue(row, base + 6);
      var rrpValue = _getValue(row, base + 7);

      // НОВАЯ ЛОГИКА: объединяем Линия + Группа в одно значение для столбца "Группа"
      var combinedGroup = "";
      if (lineValue && currentGroup) {
        combinedGroup = lineValue + " - " + currentGroup;
      } else if (lineValue) {
        combinedGroup = lineValue;
      } else if (currentGroup) {
        combinedGroup = currentGroup;
      }

      // Структура для основного прайса: ID-P | Арт. произв. | Название ENG | шт./уп. | Цена | всего | RRP | Группа
      rows.push([
        "", // ID-P (заполнится позже)
        productCode, // Арт. произв.
        nameValue, // Название ENG прайс произв
        masterBoxValue, // шт./уп.
        priceValue, // Цена
        totalValue, // всего
        rrpValue, // RRP
        combinedGroup, // Группа (объединенное значение "Линия - Группа")
      ]);
      articles.push(productCode);
      lines.push(""); // Линия теперь пустая
      groups.push(combinedGroup); // Группа содержит объединенное значение
    }

    return {
      headers: MAIN_OUTPUT_HEADERS.slice(),
      rows: rows,
      articles: articles,
      lines: lines,
      groups: groups,
    };
  }

  function _buildProcessedProbesData_(values, config) {
    var expectedHeaders = [
      "REFERENC.",
      "PRODUCT",
      "PRODUCTO",
      "TYPE",
      "AREA",
      "UNITS",
      "PRICE",
      "TOTAL",
    ];
    var headerRowIndex = -1;
    var normalizedHeaderRow = [];

    for (var i = 0; i < values.length; i++) {
      var row = values[i] || [];
      var normalized = row.map(function (value) {
        return String(value || "")
          .trim()
          .toUpperCase();
      });
      if (!normalized.length) {
        continue;
      }

      var headerFound = false;
      for (
        var start = 0;
        start <= normalized.length - expectedHeaders.length;
        start++
      ) {
        var match = true;
        for (var h = 0; h < expectedHeaders.length; h++) {
          if (normalized[start + h] !== expectedHeaders[h]) {
            match = false;
            break;
          }
        }
        if (match) {
          headerFound = true;
          break;
        }
      }
      if (headerFound) {
        headerRowIndex = i;
        normalizedHeaderRow = normalized;
        break;
      }
    }

    if (headerRowIndex === -1) {
      throw new Error(
        'На листе "-пробники" не найден блок с заголовками REFERENC./PRODUCT/...'
      );
    }

    var headerRow = values[headerRowIndex] || [];
    var headerToIndex = {};
    for (var c = 0; c < headerRow.length; c++) {
      var key = normalizedHeaderRow[c];
      if (key && headerToIndex[key] === undefined) {
        headerToIndex[key] = c;
      }
    }

    var referenceIndex = headerToIndex["REFERENC."];
    if (referenceIndex === undefined) {
      throw new Error("Не найден столбец REFERENC. в заголовке пробников.");
    }
    var productIndex = headerToIndex["PRODUCT"];
    var typeIndex = headerToIndex["TYPE"];
    var areaIndex = headerToIndex["AREA"];
    var unitsIndex = headerToIndex["UNITS"];
    var priceIndex = headerToIndex["PRICE"];
    var totalIndex = headerToIndex["TOTAL"];
    var idpIndex = headerToIndex["ID-P"];

    var possibleGroupHeaders = ["GROUP", "GRUPO", "GRUPA", "GRUPPO"];
    var groupColumnIndex = -1;
    for (var g = 0; g < possibleGroupHeaders.length; g++) {
      var idx = headerToIndex[possibleGroupHeaders[g]];
      if (idx !== undefined) {
        groupColumnIndex = idx;
        break;
      }
    }

    var rows = [];
    var articles = [];
    var lines = [];
    var groups = [];
    var currentGroup = "Пробники";
    var stopMarker =
      "CATÁLOGOS Y MATERIALES IMPRESOS / CATALOGUES & PRINTED MATERIALS";
    var stopMarkerUpper = stopMarker.toUpperCase();

    for (var r = headerRowIndex + 1; r < values.length; r++) {
      var rawRow = values[r] || [];
      var rowText = rawRow
        .map(function (value) {
          return String(value || "").trim();
        })
        .join(" ")
        .replace(/\s+/g, " ")
        .trim()
        .toUpperCase();
      if (rowText.indexOf(stopMarkerUpper) !== -1) {
        break;
      }

      var referenceValue = _asTrimmedString(
        referenceIndex !== undefined && referenceIndex < rawRow.length
          ? rawRow[referenceIndex]
          : ""
      );
      if (!referenceValue) {
        continue;
      }

      if (!/^0{2}/.test(referenceValue)) {
        currentGroup = referenceValue;
        continue;
      }

      var productValue = _asTrimmedString(
        productIndex !== undefined && productIndex < rawRow.length
          ? rawRow[productIndex]
          : ""
      );
      if (!productValue) {
        continue;
      }

      var typeValue = _asTrimmedString(
        typeIndex !== undefined && typeIndex < rawRow.length
          ? rawRow[typeIndex]
          : ""
      );
      var areaValue = _asTrimmedString(
        areaIndex !== undefined && areaIndex < rawRow.length
          ? rawRow[areaIndex]
          : ""
      );
      var unitsValue =
        unitsIndex !== undefined && unitsIndex < rawRow.length
          ? rawRow[unitsIndex]
          : rawRow[5];
      var priceValue =
        priceIndex !== undefined && priceIndex < rawRow.length
          ? rawRow[priceIndex]
          : rawRow[6];
      var totalValue =
        totalIndex !== undefined && totalIndex < rawRow.length
          ? rawRow[totalIndex]
          : rawRow[7];
      var idpValue = _asTrimmedString(
        idpIndex !== undefined && idpIndex < rawRow.length
          ? rawRow[idpIndex]
          : ""
      );
      if (idpValue === "-") {
        idpValue = "";
      }

      var groupFromSheet = _asTrimmedString(
        groupColumnIndex !== -1 && groupColumnIndex < rawRow.length
          ? rawRow[groupColumnIndex]
          : ""
      );
      var baseGroupValue =
        groupFromSheet || currentGroup || "Пробники";

      var combinedGroup = "";
      if (areaValue && baseGroupValue) {
        combinedGroup = areaValue + " - " + baseGroupValue;
      } else if (areaValue) {
        combinedGroup = areaValue;
      } else if (baseGroupValue) {
        combinedGroup = baseGroupValue;
      }
      if (!combinedGroup) {
        combinedGroup = "Пробники";
      }

      // Структура для пробников: ID-P | Арт. произв. | Название ENG | шт./уп. | кол-во | Цена | Всего | Группа
      rows.push([
        idpValue, // ID-P
        referenceValue, // Арт. произв.
        productValue, // Название ENG прайс произв
        typeValue, // шт./уп.
        unitsValue, // кол-во
        priceValue, // Цена
        totalValue, // Всего
        combinedGroup, // Группа
      ]);
      articles.push(referenceValue);
      lines.push("");
      groups.push(combinedGroup);
    }

    return {
      headers: PROBES_OUTPUT_HEADERS.slice(),
      rows: rows,
      articles: articles,
      lines: lines,
      groups: groups,
    };
  }

  function _assignSequentialIdpForProbes_(processed) {
    if (!processed || !processed.rows || !processed.rows.length) {
      return;
    }
    var headers = processed.headers || [];
    var idpIndex = headers.indexOf("ID-P");
    if (idpIndex === -1) {
      idpIndex = 0;
    }
    var startId = _getMaxIdpOnMain_();
    if (startId === null || startId === undefined || isNaN(startId)) {
      startId = 0;
    }
    var currentId = startId;
    for (var i = 0; i < processed.rows.length; i++) {
      var row = processed.rows[i];
      if (!row) {
        continue;
      }
      while (row.length <= idpIndex) {
        row.push("");
      }
      currentId += 1;
      row[idpIndex] = currentId;
    }
  }

  function _getMaxIdpOnMain_() {
    try {
      var mainSheetName =
        global.CONFIG && global.CONFIG.SHEETS && global.CONFIG.SHEETS.PRIMARY;
      if (!mainSheetName) {
        return 0;
      }
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) {
        return 0;
      }
      var sheet = ss.getSheetByName(mainSheetName);
      if (!sheet) {
        return 0;
      }
      var lastRow = sheet.getLastRow();
      if (lastRow <= 1) {
        return 0;
      }
      var lastColumn = sheet.getLastColumn();
      if (lastColumn <= 0) {
        return 0;
      }
      var headers = sheet
        .getRange(1, 1, 1, lastColumn)
        .getValues()[0]
        .map(function (value) {
          return String(value || "").trim();
        });
      var idpIndex = headers.indexOf("ID-P");
      if (idpIndex === -1) {
        return 0;
      }
      var idpValues = sheet
        .getRange(2, idpIndex + 1, lastRow - 1, 1)
        .getValues();
      var maxId = 0;
      for (var i = 0; i < idpValues.length; i++) {
        var val = _parseNumber(idpValues[i][0]);
        if (val !== null && !isNaN(val) && val > maxId) {
          maxId = val;
        }
      }
      return maxId;
    } catch (err) {
      Lib.logError("_getMaxIdpOnMain_: ошибка", err);
      return 0;
    }
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
      baseSheetName = resolvedKey === "PROBES" ? "-пробники" : "-Б/З поставщик";
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

    var headers =
      processed.headers && processed.headers.length
        ? processed.headers.slice()
        : config.OUTPUT_HEADERS
        ? config.OUTPUT_HEADERS.slice()
        : [];
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

  function _detectGroupRow(row, bgRow, loweredGroupColor) {
    if (!row || row.length < 3) {
      return "";
    }
    var cellB = _asTrimmedString(row[1]);
    if (!cellB || cellB.indexOf("-") === -1) {
      return "";
    }
    var cellCEmpty = _isEmpty(row[2]);
    if (!cellCEmpty) {
      return "";
    }
    var background = bgRow && bgRow.length > 1 ? bgRow[1] : "";
    var colorMatch = (background || "").toLowerCase() === loweredGroupColor;
    if (!colorMatch) {
      return "";
    }
    var dashPos = cellB.indexOf("-");
    return cellB.substring(dashPos + 1).trim();
  }

  function _syncIdWithMain(
    sourceRows,
    sourceHeaders,
    articles,
    groups,
    options
  ) {
    var emptyResult = { createdRows: [], assignedIdp: {} };
    if (!global.CONFIG || !global.CONFIG.SHEETS) {
      return emptyResult;
    }
    options = options || {};
    var overrideIdpFromSource = !!options.overrideIdpFromSource;
    var forceSamplesIdg = !!options.forceSamplesIdg;
    var samplesGroupKey = options.samplesGroupKey || "Пробники";
    try {
      Lib.logDebug(
        "[SK] syncIdWithMain: старт, строк " + (articles ? articles.length : 0)
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
      // Опираемся на ФАКТИЧЕСКУЮ шапку листа, а не на конфиг
      var targetColumnCount = mainSheet.getLastColumn();
      var headersRange = targetColumnCount > 0 ? mainSheet.getRange(1, 1, 1, targetColumnCount) : null;
      var headers = headersRange
        ? headersRange
            .getValues()[0]
            .map(function (name) {
              return String(name || "").trim();
            })
        : [];
      if (!headers || headers.length === 0) {
        return emptyResult;
      }

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
      var lineColumnIndex = headerToIndex[_normalizeHeaderKey("Линия")];
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
      var maxExistingIdG = 0;
      var maxPrimaryNumeric = 0;
      var samplesGroupIdG = null; // Для хранения единого ID-G группы Пробники

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
          var existingIdG = parseInt(mainValues[i][idgColumnIndex], 10);
          if (!isNaN(existingIdG) && existingIdG > maxExistingIdG) {
            maxExistingIdG = existingIdG;
          }

          var groupExisting =
            groupColumnIndex !== undefined
              ? _asTrimmedString(rowArray[groupColumnIndex])
              : "";
          // НОВАЯ ЛОГИКА: используем только "Группа" для создания карты ID-G
          if (groupExisting) {
            groupIdMap[groupExisting] = existingIdG;
          }
          // Если в main уже есть ID-G для группы 'Пробники', запомним его
          if (
            String(groupExisting || "").trim() === samplesGroupKey &&
            !isNaN(existingIdG)
          ) {
            if (!samplesGroupIdG || existingIdG > samplesGroupIdG) {
              samplesGroupIdG = existingIdG;
            }
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
      var idgCounter = maxExistingIdG;
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

          var newIdG = null;
          if (idgColumnIndex !== undefined) {
            if (forceSamplesIdg) {
              if (!samplesGroupIdG) {
                if (groupIdMap[samplesGroupKey]) {
                  samplesGroupIdG = groupIdMap[samplesGroupKey];
                } else if (!isNaN(maxExistingIdG) && maxExistingIdG > 0) {
                  samplesGroupIdG = maxExistingIdG;
                }
              }
              if (!samplesGroupIdG) {
                samplesGroupIdG = ++idgCounter;
              } else if (samplesGroupIdG > idgCounter) {
                idgCounter = samplesGroupIdG;
              }
              newIdG = samplesGroupIdG;
              groupIdMap[samplesGroupKey] = newIdG;
            } else if (groupValueInput === samplesGroupKey) {
              if (!samplesGroupIdG) {
                if (groupIdMap[samplesGroupKey]) {
                  samplesGroupIdG = groupIdMap[samplesGroupKey];
                } else if (!isNaN(maxExistingIdG) && maxExistingIdG > 0) {
                  samplesGroupIdG = maxExistingIdG;
                } else {
                  samplesGroupIdG = ++idgCounter;
                }
              }
              newIdG = samplesGroupIdG;
              groupIdMap[samplesGroupKey] = newIdG;
            } else {
              // Для остальных групп используем стандартную логику на основе "Группа"
              if (groupIdMap[groupValueInput]) {
                newIdG = groupIdMap[groupValueInput];
              } else {
                newIdG = ++idgCounter;
                groupIdMap[groupValueInput] = newIdG;
              }
            }
            if (
              groupValueInput &&
              newIdG !== null &&
              groupIdMap[groupValueInput] !== newIdG
            ) {
              groupIdMap[groupValueInput] = newIdG;
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
            idg: newIdG,
            primaryId: newPrimaryId,
            line: "",
            group: groupValueInput,
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

        var srcRow = sourceRows && sourceRows[j];
        var assignedIdP = null;
        var forcedIdp = false;
        if (
          overrideIdpFromSource &&
          srcRow &&
          sourceIdpIdx !== undefined &&
          sourceIdpIdx < srcRow.length
        ) {
          var forcedIdpValue = _parseNumber(srcRow[sourceIdpIdx]);
          if (forcedIdpValue !== null && !isNaN(forcedIdpValue)) {
            assignedIdP = forcedIdpValue;
            forcedIdp = true;
            if (forcedIdpValue > idpCounter) {
              idpCounter = forcedIdpValue;
            }
          }
        }
        if (!forcedIdp && idpColumnIndex !== undefined) {
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
            "[SK] syncIdWithMain: назначен новый ID-P " +
              assignedIdP +
              " для артикула " +
              articleCode
          );
        } else if (!forcedIdp && assignedIdP !== null) {
          Lib.logDebug(
            "[SK] syncIdWithMain: переиспользован ID-P " +
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
          var srcRowRefEarly = srcRow;
          if (srcRowRefEarly && sourceIdpIdx < srcRowRefEarly.length) {
            srcRowRefEarly[sourceIdpIdx] = assignedIdP;
          }
        }

        if (srcRow) {
          for (var name in sourceHeaderToIndex) {
            if (!sourceHeaderToIndex.hasOwnProperty(name)) continue;
            var targetIndex = headerToIndex[name];
            if (targetIndex === undefined || targetIndex >= targetColumnCount)
              continue;
            // Пропускаем столбец "Линия", он не должен заполняться
            if (targetIndex === lineColumnIndex) continue;
            var srcIdx = sourceHeaderToIndex[name];
            rowValues[targetIndex] = srcRow[srcIdx] || "";
          }
        }

        if (entry.__meta) {
          if (idColumnIndex !== undefined && entry.__meta.primaryId) {
            rowValues[idColumnIndex] = entry.__meta.primaryId;
          }
          if (idgColumnIndex !== undefined && entry.__meta.idg !== null) {
            rowValues[idgColumnIndex] = entry.__meta.idg;
          }
        }

        if (articleColumnIndex !== undefined) {
          rowValues[articleColumnIndex] = articleCode;
        }
        if (nameColumnIndex !== undefined && srcRow) {
          var nameIdx =
            sourceHeaderToIndex[
              _normalizeHeaderKey("Название ENG прайс произв")
            ];
          if (nameIdx !== undefined && nameIdx < srcRow.length) {
            rowValues[nameColumnIndex] =
              srcRow[nameIdx] || rowValues[nameColumnIndex];
          }
        }
        // Столбец "Линия" теперь остается пустым, не заполняем
        // Для столбца "Группа" проверяем: если артикул существующий, сравниваем значения
        if (groupColumnIndex !== undefined && groupValueInput) {
          var existingGroupValue = _asTrimmedString(rowValues[groupColumnIndex]);

          // Проверяем, является ли артикул существующим (не в списке новых)
          var isExistingArticle = !newRows.some(function(newRow) {
            return newRow.rowArray === rowValues;
          });

          if (isExistingArticle && existingGroupValue && existingGroupValue !== groupValueInput) {
            // Для существующего артикула значение изменилось - выводим уведомление
            Lib.logInfo(
              "[SK] Группа для артикула " + articleCode +
              ": было '" + existingGroupValue + "' => стало '" + groupValueInput + "'"
            );
          }

          // Записываем новое значение (для новых артикулов или для обновления существующих)
          rowValues[groupColumnIndex] = groupValueInput;
        }
        if (unitsColumnIndex !== undefined && srcRow) {
          var unitsIdx = sourceHeaderToIndex[_normalizeHeaderKey("шт./уп.")];
          if (unitsIdx !== undefined && unitsIdx < srcRow.length) {
            rowValues[unitsColumnIndex] =
              srcRow[unitsIdx] || rowValues[unitsColumnIndex];
          }
        }

        // НОВАЯ ЛОГИКА: Линия больше не используется, всегда пустая
        var resolvedLine = "";
        var resolvedGroup =
          groupColumnIndex !== undefined
            ? _asTrimmedString(rowValues[groupColumnIndex])
            : groupValueInput;
        if (entry.__meta) {
          entry.__meta.line = resolvedLine;
          entry.__meta.group = resolvedGroup;
        }

        if (idgColumnIndex !== undefined) {
          var currentIdgValue = rowValues[idgColumnIndex];
          var normalizedResolvedGroup = _asTrimmedString(resolvedGroup);
          var mapKey = forceSamplesIdg ? samplesGroupKey : normalizedResolvedGroup;

          if (forceSamplesIdg) {
            var parsedCurrentIdg = _parseNumber(currentIdgValue);
            if (!samplesGroupIdG) {
              if (parsedCurrentIdg !== null && !isNaN(parsedCurrentIdg)) {
                samplesGroupIdG = parsedCurrentIdg;
              } else if (groupIdMap[mapKey]) {
                samplesGroupIdG = groupIdMap[mapKey];
              } else if (!isNaN(maxExistingIdG) && maxExistingIdG > 0) {
                samplesGroupIdG = maxExistingIdG;
              }
            }
            if (!samplesGroupIdG) {
              samplesGroupIdG = ++idgCounter;
            } else if (samplesGroupIdG > idgCounter) {
              idgCounter = samplesGroupIdG;
            }
            currentIdgValue = samplesGroupIdG;
            rowValues[idgColumnIndex] = currentIdgValue;
            groupIdMap[mapKey] = currentIdgValue;
            if (
              normalizedResolvedGroup &&
              !groupIdMap[normalizedResolvedGroup]
            ) {
              groupIdMap[normalizedResolvedGroup] = currentIdgValue;
            }
          } else if (normalizedResolvedGroup === samplesGroupKey) {
            if (!samplesGroupIdG) {
              var existing = groupIdMap[mapKey];
              if (existing) {
                samplesGroupIdG = existing;
              } else if (!isNaN(maxExistingIdG) && maxExistingIdG > 0) {
                samplesGroupIdG = maxExistingIdG;
              } else {
                samplesGroupIdG = ++idgCounter;
              }
            }
            currentIdgValue = samplesGroupIdG;
            rowValues[idgColumnIndex] = currentIdgValue;
            groupIdMap[mapKey] = currentIdgValue;
          } else {
            var mappedIdg = groupIdMap[mapKey];
            if (!mappedIdg) {
              mappedIdg = ++idgCounter;
              groupIdMap[mapKey] = mappedIdg;
            }
            if (
              currentIdgValue === undefined ||
              currentIdgValue === null ||
              currentIdgValue === "" ||
              currentIdgValue !== mappedIdg
            ) {
              currentIdgValue = mappedIdg;
              rowValues[idgColumnIndex] = currentIdgValue;
            }
            groupIdMap[mapKey] = mappedIdg;
          }

          if (entry.__meta) {
            entry.__meta.idg = currentIdgValue;
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
            "[SK] syncIdWithMain: обновление строки " +
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
          "[SK] syncIdWithMain: вставлено " +
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
            "[SK] syncIdWithMain: запись ID-P " +
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
        "[SK] syncIdWithMain: итоговый счётчик ID-P = " + idpCounter
      );

      Lib.logDebug(
        "[SK] syncIdWithMain: завершено, создано новых строк: " +
          createdEntries.length
      );
      return { createdRows: createdEntries, assignedIdp: assignedIdpMap };
    } catch (syncError) {
      Lib.logError("syncIdWithMain: ошибка", syncError);
      return emptyResult;
    }
  }

  function _handleNewArticles_(createdEntries) {
    if (!createdEntries || !createdEntries.length) {
      return;
    }
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var mainSheetName =
        global.CONFIG && global.CONFIG.SHEETS && global.CONFIG.SHEETS.PRIMARY;
      var mainSheet = mainSheetName ? ss.getSheetByName(mainSheetName) : null;
      if (!mainSheet) {
        return;
      }

      var lastColumn = mainSheet.getLastColumn();
      if (lastColumn <= 0) {
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
        var key = _normalizeHeaderKey(name);
        if (key && headerToIndex[key] === undefined) headerToIndex[key] = idx;
      });

      var statusIdx = headerToIndex[_normalizeHeaderKey("Статус")];
      if (statusIdx === undefined) statusIdx = -1;
      var volumeIdx = headerToIndex[_normalizeHeaderKey("Объём")];
      if (volumeIdx === undefined) volumeIdx = -1;
      var unitsIdxMain = headerToIndex[_normalizeHeaderKey("шт./уп.")];
      if (unitsIdxMain === undefined) unitsIdxMain = -1;
      var formIdxMain = headerToIndex[_normalizeHeaderKey("Форма выпуска")];
      if (formIdxMain === undefined) formIdxMain = -1;

      Lib.logDebug("[SK] _handleNewArticles_: statusIdx=" + statusIdx + ", volumeIdx=" + volumeIdx);

      var processedRows = [];

      createdEntries.forEach(function (entry) {
        if (!entry || !entry.rowNumber) {
          return;
        }

        // Создаём строки на базовых листах (Сертификация, Этикетки)
        var primaryId = _asTrimmedString(
          mainSheet.getRange(entry.rowNumber, 1).getValue()
        );
        if (primaryId && typeof Lib.ensureRowExistsOnBaseSheets === "function") {
          Lib.ensureRowExistsOnBaseSheets(primaryId);
        }

        var productName =
          entry.productName || entry.articleCode || "Новый продукт";

        // НОВАЯ ЛОГИКА: автоматически ставим статус "New завод"
        if (statusIdx !== undefined && statusIdx !== -1) {
          Lib.logDebug("[SK] Установка статуса 'New завод' для строки " + entry.rowNumber);
          mainSheet
            .getRange(entry.rowNumber, statusIdx + 1)
            .setValue("New завод");
        } else {
          Lib.logWarn("[SK] Не удалось найти столбец 'Статус' для строки " + entry.rowNumber);
        }

        // Запрашиваем только Объём
        var volumeValue = _promptVolumeInput_(productName);
        if (
          volumeValue !== null &&
          volumeIdx !== undefined &&
          volumeIdx !== -1
        ) {
          Lib.logDebug("[SK] Установка объёма '" + volumeValue + "' для строки " + entry.rowNumber);
          mainSheet
            .getRange(entry.rowNumber, volumeIdx + 1)
            .setValue(volumeValue);
        }

        // Доп. логика: если в колонке "шт./уп." указан тип упаковки (Jar/Tarro, Sachet, Tube, Dropper, Bottle, Airless),
        // то ставим 25 в "шт./уп.", а в "Форма выпуска" — соответствующее значение на русском
        if (unitsIdxMain !== -1) {
          var rawUnits = _asTrimmedString(
            mainSheet.getRange(entry.rowNumber, unitsIdxMain + 1).getValue()
          );
          var unitsUpper = rawUnits.toUpperCase();
          var packagingForm = null;
          if (unitsUpper) {
            // Приоритетные синонимы и ключевые слова
            // Прим.: выбираем первое совпадение по приоритету
            if (unitsUpper.indexOf("SACHET") !== -1) packagingForm = "Саше";
            else if (unitsUpper.indexOf("TUBE") !== -1) packagingForm = "Тюбик";
            else if (unitsUpper.indexOf("BOTTLE") !== -1) packagingForm = "Бутылка";
            else if (unitsUpper.indexOf("AIRLESS") !== -1) packagingForm = "Баночка";
            else if (unitsUpper.indexOf("DROPPER") !== -1) packagingForm = "Баночка";
            else if (
              unitsUpper.indexOf("JAR") !== -1 ||
              unitsUpper.indexOf("TARRO") !== -1
            )
              packagingForm = "Баночка";
          }

          if (packagingForm) {
            Lib.logDebug(
              "[SK] Обнаружен тип упаковки '" +
                rawUnits +
                "' → ставим шт./уп.=25 и Форма выпуска='" +
                packagingForm +
                "' в строке " +
                entry.rowNumber
            );
            mainSheet.getRange(entry.rowNumber, unitsIdxMain + 1).setValue(25);
            if (formIdxMain !== -1) {
              mainSheet
                .getRange(entry.rowNumber, formIdxMain + 1)
                .setValue(packagingForm);
            }
          }
        }

        processedRows.push(entry.rowNumber);

        // Остальные поля (Название ENG, шт./уп., Группа, Арт. произв.) уже заполнены при синхронизации
      });

      SpreadsheetApp.flush();
      return processedRows;
    } catch (err) {
      Lib.logError("_handleNewArticles_: ошибка", err);
      return [];
    }
  }

  function _collectColumnUniqueValues_(sheet, columnIndex) {
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return [];
    }
    var values = sheet.getRange(2, columnIndex, lastRow - 1, 1).getValues();
    var seen = Object.create(null);
    var result = [];
    for (var i = 0; i < values.length; i++) {
      var value = _asTrimmedString(values[i][0]);
      if (value && !seen[value]) {
        seen[value] = true;
        result.push(value);
      }
    }
    return result;
  }

  function _promptStatusChoice_(productName) {
    var ui = SpreadsheetApp.getUi();
    var message =
      'Продукт "' +
      productName +
      '" это?\n\n' +
      "Нажмите «Да» → «New в работу»\n" +
      "Нажмите «Нет» → «Новый артикул»\n" +
      "Нажмите «Отмена» → «Нужна проверка»";
    var result = ui.alert("Новый продукт", message, ui.ButtonSet.YES_NO_CANCEL);
    if (result === ui.Button.YES) {
      return "New в работу";
    }
    if (result === ui.Button.NO) {
      return "Новый артикул";
    }
    if (result === ui.Button.CANCEL) {
      return "Нужна проверка";
    }
    return null;
  }

  function _promptVolumeInput_(productName) {
    var ui = SpreadsheetApp.getUi();
    while (true) {
      var prompt = ui.prompt(
        'Объём "' + productName + '"',
        'Введите значение объёма или нажмите "Отмена", чтобы пропустить.',
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

  function _promptFormChoice_(productName, options) {
    if (!options || !options.length) {
      return null;
    }
    var ui = SpreadsheetApp.getUi();
    while (true) {
      var message =
        'Форма выпуска для "' +
        productName +
        '".\nВведите номер варианта и нажмите OK либо нажмите "Отмена", чтобы пропустить.\n\n';
      for (var i = 0; i < options.length; i++) {
        message += i + 1 + ". " + options[i] + "\n";
      }
      var response = ui.prompt(
        "Форма выпуска",
        message,
        ui.ButtonSet.OK_CANCEL
      );
      var button = response.getSelectedButton();
      if (button !== ui.Button.OK) {
        return null;
      }
      var text = _asTrimmedString(response.getResponseText());
      var index = parseInt(text, 10);
      if (!isNaN(index) && index >= 1 && index <= options.length) {
        return options[index - 1];
      }
      ui.alert(
        "Введите число из списка или нажмите «Отмена», чтобы пропустить."
      );
    }
  }

  function _promptCategoryChoice_(productName, lineValue, groupValue, options) {
    if (!options || !options.length) {
      return null;
    }
    var ui = SpreadsheetApp.getUi();
    while (true) {
      var message =
        '"' +
        productName +
        '" линии "' +
        (lineValue || "—") +
        '" группы "' +
        (groupValue || "—") +
        '" — выберите категорию товара.\n' +
        'Введите номер варианта и нажмите OK либо нажмите "Отмена", чтобы пропустить.\n\n';
      for (var i = 0; i < options.length; i++) {
        message += i + 1 + ". " + options[i] + "\n";
      }
      var response = ui.prompt(
        "Категория товара",
        message,
        ui.ButtonSet.OK_CANCEL
      );
      var button = response.getSelectedButton();
      if (button !== ui.Button.OK) {
        return null;
      }
      var text = _asTrimmedString(response.getResponseText());
      var index = parseInt(text, 10);
      if (!isNaN(index) && index >= 1 && index <= options.length) {
        return options[index - 1];
      }
      ui.alert(
        "Введите номер из списка или нажмите «Отмена», чтобы пропустить."
      );
    }
  }

  function _applyAssignedIdpToProcessed_(processed, assignedMap) {
    if (!processed || !processed.rows || !assignedMap) {
      Lib.logDebug("[SK] _applyAssignedIdpToProcessed_: пропуск - нет данных");
      return;
    }
    var headers = processed.headers || [];
    var idpIndex = headers.indexOf("ID-P");
    var articleIndex = headers.indexOf("Арт. произв.");

    Lib.logDebug(
      "[SK] _applyAssignedIdpToProcessed_: idpIndex=" +
        idpIndex +
        ", articleIndex=" +
        articleIndex +
        ", assignedMap keys=" +
        Object.keys(assignedMap).length
    );

    var fallback = false;
    if (idpIndex === -1) {
      idpIndex = 0;
      fallback = true;
    }
    if (articleIndex === -1) {
      articleIndex = fallback ? 1 : headers.indexOf("REFERENC.");
      if (articleIndex === -1) {
        articleIndex = 1;
      }
    }

    var updatedCount = 0;
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
        updatedCount++;
      }
    }
    Lib.logDebug(
      "[SK] _applyAssignedIdpToProcessed_: обновлено строк " + updatedCount
    );
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
      Lib.logDebug("[SK] _clearIdpColumnOnMain: очищен столбец ID-P");
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
      Lib.logDebug("[SK] _clearIdgColumnOnMain: очищен столбец ID-G");
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
          "[SK] _fillIdgForRowsWithoutIdp: не найдены необходимые столбцы"
        );
        return;
      }

      // Получаем все данные листа
      var allData = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();

      // Этап 1: Создаем карту "Группа" -> "ID-G" для строк с ID-P
      var groupToIdgMap = {};
      for (var i = 0; i < allData.length; i++) {
        var row = allData[i];
        var idpValue = row[idpIndex];
        var idgValue = row[idgIndex];
        var groupValue = _asTrimmedString(row[groupIndex]);

        // Если есть ID-P и ID-G и Группа, сохраняем в карту
        if (
          idpValue !== null &&
          idpValue !== undefined &&
          idpValue !== "" &&
          idgValue !== null &&
          idgValue !== undefined &&
          idgValue !== "" &&
          groupValue
        ) {
          if (!groupToIdgMap[groupValue]) {
            groupToIdgMap[groupValue] = idgValue;
          }
        }
      }

      Lib.logDebug(
        "[SK] _fillIdgForRowsWithoutIdp: создана карта Группа->ID-G, записей: " +
          Object.keys(groupToIdgMap).length
      );

      // Этап 2: Проставляем ID-G для строк без ID-P, используя карту
      var updatedCount = 0;
      for (var j = 0; j < allData.length; j++) {
        var currentRow = allData[j];
        var currentIdp = currentRow[idpIndex];
        var currentIdg = currentRow[idgIndex];
        var currentGroup = _asTrimmedString(currentRow[groupIndex]);

        // Если нет ID-P, но есть Группа, и для этой группы есть ID-G в карте
        if (
          (currentIdp === null || currentIdp === undefined || currentIdp === "") &&
          currentGroup &&
          groupToIdgMap[currentGroup]
        ) {
          // Проставляем ID-G
          sheet.getRange(j + 2, idgIndex + 1).setValue(groupToIdgMap[currentGroup]);
          updatedCount++;
          Lib.logDebug(
            "[SK] _fillIdgForRowsWithoutIdp: строка " +
              (j + 2) +
              ", Группа '" +
              currentGroup +
              "' -> ID-G " +
              groupToIdgMap[currentGroup]
          );
        }
      }

      Lib.logInfo(
        "[SK] _fillIdgForRowsWithoutIdp: завершено, обновлено строк: " + updatedCount
      );
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

  /**
   * Функция для автоматического заполнения скидки 20% на листе "Динамика цены"
   * Заполняет столбец "СКИДКА ОТ EXW 2024, %" значением 0.2 (20%), если есть значение в "ЦЕНА EXW из Б/З, €"
   */
  function _fillDiscountForPriceDynamics() {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) {
        Lib.logWarn("[SK] _fillDiscountForPriceDynamics: не удалось получить активную таблицу");
        return;
      }

      var priceDynamicsSheetName =
        global.CONFIG &&
        global.CONFIG.SHEETS &&
        global.CONFIG.SHEETS.PRICE_DYNAMICS
          ? global.CONFIG.SHEETS.PRICE_DYNAMICS
          : "Динамика цены";

      var sheet = ss.getSheetByName(priceDynamicsSheetName);
      if (!sheet) {
        Lib.logWarn("[SK] _fillDiscountForPriceDynamics: лист \"" + priceDynamicsSheetName + "\" не найден");
        return;
      }

      var lastRow = sheet.getLastRow();
      if (lastRow <= 1) {
        Lib.logDebug("[SK] _fillDiscountForPriceDynamics: нет данных на листе");
        return;
      }

      var lastColumn = sheet.getLastColumn();
      if (lastColumn <= 0) {
        Lib.logDebug("[SK] _fillDiscountForPriceDynamics: нет столбцов на листе");
        return;
      }

      var headers = sheet
        .getRange(1, 1, 1, lastColumn)
        .getValues()[0]
        .map(function (value) {
          return String(value || "").trim();
        });

      var priceExwIndex = headers.indexOf("ЦЕНА EXW из Б/З, €");
      if (priceExwIndex === -1) {
        priceExwIndex = headers.indexOf("ЦЕНА EXW из Б/З");
      }

      // Ищем столбец скидки для текущего года
      var currentYear = new Date().getFullYear();
      var discountIndex = headers.indexOf("СКИДКА ОТ EXW " + currentYear + ", %");
      // Если не найден для текущего года, пробуем предыдущий год
      if (discountIndex === -1) {
        discountIndex = headers.indexOf("СКИДКА ОТ EXW " + (currentYear - 1) + ", %");
      }

      if (priceExwIndex === -1) {
        Lib.logWarn("[SK] _fillDiscountForPriceDynamics: столбец \"ЦЕНА EXW из Б/З, €\" не найден");
        return;
      }

      if (discountIndex === -1) {
        Lib.logWarn("[SK] _fillDiscountForPriceDynamics: столбец \"СКИДКА ОТ EXW " + currentYear + ", %\" не найден");
        return;
      }

      // Получаем данные столбца "ЦЕНА EXW из Б/З, €"
      var priceData = sheet.getRange(2, priceExwIndex + 1, lastRow - 1, 1).getValues();
      var discountData = [];
      var filledCount = 0;

      for (var i = 0; i < priceData.length; i++) {
        var priceValue = priceData[i][0];
        // Если есть цена (не пустая и не равна 0), ставим 20% (0.2)
        if (priceValue !== "" && priceValue !== null && priceValue !== undefined && priceValue !== 0) {
          discountData.push([0.2]);
          filledCount++;
        } else {
          discountData.push([""]);
        }
      }

      // Записываем данные в столбец скидки
      if (filledCount > 0) {
        sheet.getRange(2, discountIndex + 1, discountData.length, 1).setValues(discountData);
        // Применяем форматирование процентов
        sheet.getRange(2, discountIndex + 1, discountData.length, 1).setNumberFormat("0%");
        Lib.logInfo("[SK] _fillDiscountForPriceDynamics: заполнено " + filledCount + " строк со скидкой 20%");
      } else {
        Lib.logDebug("[SK] _fillDiscountForPriceDynamics: нет строк с ценой для заполнения скидки");
      }
    } catch (err) {
      Lib.logError("[SK] _fillDiscountForPriceDynamics: ошибка", err);
    }
  }
})(Lib, this);
