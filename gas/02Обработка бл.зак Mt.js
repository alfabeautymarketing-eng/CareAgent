var Lib = Lib || {};

(function (Lib, global) {
  var TARGET_PROJECT_KEY = "MT";
  var MAIN_OUTPUT_HEADERS = [
    "ID-P",
    "Арт. произв.",
    "Название ENG прайс произв",
    "Объём",
    "BAR CODE",
    "шт./уп.",
    "Цена",
    "Группа",
  ];

  var TESTER_OUTPUT_HEADERS = [
    "ID-P",
    "Арт. произв.",
    "Название ENG прайс произв",
    "Объём",
    "BAR CODE",
    "шт./уп.",
    "Цена",
    "Группа",
  ];

  var SAMPLES_OUTPUT_HEADERS = [
    "ID-P",
    "Арт. произв.",
    "Название ENG прайс произв",
    "Объём",
    "BAR CODE",
    "шт./уп.",
    "Цена",
    "Группа",
  ];

  // ========== ФУНКЦИЯ 1: Б/З поставщик ==========
  Lib.processMtMainPrice = function () {
    var ui = SpreadsheetApp.getUi();
    var config = _getPrimaryDataConfig_();
    var menuTitle = _getMenuTitle_(config);

    if (!_isActiveProject_()) {
      ui.alert(
        menuTitle,
        "Эта функция доступна только в проекте MT.",
        ui.ButtonSet.OK
      );
      return;
    }

    try {
      // START: Function entry
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji(
          "Обработка основного прайс-листа MT начата",
          "INFO",
          "",
          "processMtMainPrice",
          "Пользователь инициировал обработку основного прайс-листа для проекта MT",
          "Price",
          "START"
        );
      }

      Lib.logInfo("[MT] Обработка Б/З поставщик: старт");
      var source = _getSourceData_(config, "MAIN");
      if (!source.values || !source.values.length) {
        throw new Error("В исходном документе нет данных для обработки.");
      }

      var processed = _buildProcessedMainData_(
        source.values,
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

      _fillIdgForRowsWithoutIdp();

      // Заполняем ID-P на остальных листах (Заказ, Динамика цены, Расчет цены, ABC-Анализ)
      if (typeof Lib.fillIdpOnSheetsByIdFromPrimary === 'function') {
        Lib.logInfo("[MT] Заполнение ID-P на всех листах");
        Lib.fillIdpOnSheetsByIdFromPrimary();
      } else {
        Lib.logWarn("[MT] fillIdpOnSheetsByIdFromPrimary не найдена");
      }

      // Копируем цены из обработанных данных на листы Динамика цены и Расчет цены
      if (typeof Lib.copyPriceFromPrimaryToSheets === 'function') {
        Lib.logInfo("[MT] Копирование цен на Динамика цены и Расчет цены");
        Lib.copyPriceFromPrimaryToSheets(processed);
      } else {
        Lib.logWarn("[MT] copyPriceFromPrimaryToSheets не найдена");
      }

      // Применяем формулы на листе "Динамика цены" (EXW ALFASPA, Закупочная цена, DDP-МОСКВА, Прирост)
      if (typeof Lib.recalculatePriceDynamicsFormulas === 'function') {
        Lib.logInfo("[MT] Применение формул на листе Динамика цены");
        Lib.recalculatePriceDynamicsFormulas();
      } else {
        Lib.logWarn("[MT] recalculatePriceDynamicsFormulas не найдена");
      }

      if (
        syncResult &&
        syncResult.createdRows &&
        syncResult.createdRows.length
      ) {
        var newArticleRows = _handleNewArticles_(syncResult.createdRows, "MAIN", processed);

        if (newArticleRows && newArticleRows.length > 0) {
          Lib.logInfo(
            "[MT] Запуск построчной синхронизации для " +
              newArticleRows.length +
              " новых артикулов"
          );
          SpreadsheetApp.flush();

          if (typeof Lib.syncMultipleRows === "function") {
            Lib.syncMultipleRows(newArticleRows);
          } else {
            Lib.logWarn("[MT] syncMultipleRows не найдена");
          }

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
            Lib.logInfo("[MT] Заполнение ID-P для новых артикулов на всех листах");
            Lib.fillIdpOnSheetsByIdFromPrimary();
          }

          // Заполняем ЦЕНА EXW из Б/З для новых артикулов
          if (typeof Lib.copyPriceFromPrimaryToSheets === 'function') {
            Lib.logInfo("[MT] Копирование цен для новых артикулов на Динамика цены и Расчет цены");
            Lib.copyPriceFromPrimaryToSheets(processed);
          }
        }
      }

      Lib.logInfo(
        "[MT] Обработка Б/З поставщик: завершено, строк " + processed.rows.length
      );

      // Обновляем формулы на листе "Расчет цены"
      // 1. Сначала применяем INDEX/MATCH формулы для подтягивания данных из "Динамика цены"
      if (typeof Lib.updatePriceCalculationFormulas === 'function') {
        Lib.logInfo("[MT] Обновление INDEX/MATCH формул на листе Расчет цены");
        Lib.updatePriceCalculationFormulas(true); // silent=true
      } else {
        Lib.logWarn("[MT] updatePriceCalculationFormulas не найдена");
      }

      // 2. Затем применяем расчетные формулы (К-т, Расчетная цена Опт, РРЦ и т.д.)
      if (typeof Lib.applyCalculationFormulas === 'function') {
        Lib.logInfo("[MT] Применение расчетных формул на листе Расчет цены");
        Lib.applyCalculationFormulas(true); // silent=true
      } else {
        Lib.logWarn("[MT] applyCalculationFormulas не найдена");
      }

      // ВАЖНО: updateStatusesAfterProcessing НЕ вызываем здесь!
      // Статусы должны обновляться ТОЛЬКО на последнем шаге (processMtSamplesPrice)

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

      // SUCCESS: Function completed successfully
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji(
          "Обработка основного прайс-листа MT завершена успешно",
          "INFO",
          "",
          "processMtMainPrice",
          `Обработано товаров: ${processed.rows.length}, ${syncResult && syncResult.createdRows ? syncResult.createdRows.length + ' новых артикулов' : 'нет новых артикулов'}`,
          "Price",
          "SUCCESS",
          null,
          { rowsProcessed: processed.rows.length, newArticles: syncResult && syncResult.createdRows ? syncResult.createdRows.length : 0, barcodeMismatches: syncResult && syncResult.barcodeMismatches ? syncResult.barcodeMismatches.length : 0 }
        );
      }

      ui.alert(menuTitle, message, ui.ButtonSet.OK);
    } catch (error) {
      Lib.logError("processMtMainPrice: ошибка", error);

      // ERROR: Enhanced error logging
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji(
          "Ошибка при обработке основного прайс-листа MT",
          "ERROR",
          "",
          "processMtMainPrice",
          error && error.message ? error.message : String(error),
          "Price",
          "ERROR",
          null,
          { error: error ? error.toString() : "Unknown error", stack: error && error.stack ? error.stack : null }
        );
      }

      ui.alert(
        "Ошибка обработки прайс",
        error.message || String(error),
        ui.ButtonSet.OK
      );
    }
  };

  // ========== ФУНКЦИЯ 2: Тестер ==========
  Lib.processMtTesterPrice = function () {
    var ui = SpreadsheetApp.getUi();
    var config = _getPrimaryDataConfig_();
    var menuTitle = _getMenuTitle_(config);

    if (!_isActiveProject_()) {
      ui.alert(
        menuTitle,
        "Эта функция доступна только в проекте MT.",
        ui.ButtonSet.OK
      );
      return;
    }

    try {
      // START: Function entry
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji(
          "Обработка тестеров MT начата",
          "INFO",
          "",
          "processMtTesterPrice",
          "Пользователь инициировал обработку тестеров для проекта MT",
          "Price",
          "START"
        );
      }

      Lib.logInfo("[MT] Обработка Тестер: старт");
      var source = _getSourceData_(config, "TESTER");
      if (!source.values || source.values.length <= 2) {
        ui.alert(
          menuTitle,
          "На листе нет данных для обработки.",
          ui.ButtonSet.OK
        );
        return;
      }

      var processed = _buildProcessedTesterData_(source.values, config);
      if (!processed.rows.length) {
        ui.alert(
          menuTitle,
          "Не найдено данных для обработки.",
          ui.ButtonSet.OK
        );
        return;
      }

      // НЕ очищаем ID-P и цены - тестеры добавляются к основным данным
      // ID-P будет продолжать нумерацию от максимального значения

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
      _storeProcessedSnapshot_(config, processed, "TESTER");

      // Обработка изменений групп (показываем диалоги)
      if (syncResult && syncResult.groupChanges && syncResult.groupChanges.length > 0) {
        _handleGroupChanges_(syncResult.groupChanges);
      }

      _fillIdgForRowsWithoutIdp();

      // Заполняем ID-P на остальных листах (Заказ, Динамика цены, Расчет цены, ABC-Анализ)
      if (typeof Lib.fillIdpOnSheetsByIdFromPrimary === 'function') {
        Lib.logInfo("[MT] (Тестер) Заполнение ID-P на всех листах");
        Lib.fillIdpOnSheetsByIdFromPrimary();
      } else {
        Lib.logWarn("[MT] (Тестер) fillIdpOnSheetsByIdFromPrimary не найдена");
      }

      // Копируем цены из обработанных данных на листы Динамика цены и Расчет цены
      if (typeof Lib.copyPriceFromPrimaryToSheets === 'function') {
        Lib.logInfo("[MT] (Тестер) Копирование цен на Динамика цены и Расчет цены");
        Lib.copyPriceFromPrimaryToSheets(processed);
      } else {
        Lib.logWarn("[MT] (Тестер) copyPriceFromPrimaryToSheets не найдена");
      }

      // Применяем формулы на листе "Динамика цены" (EXW ALFASPA, Закупочная цена, DDP-МОСКВА, Прирост)
      if (typeof Lib.recalculatePriceDynamicsFormulas === 'function') {
        Lib.logInfo("[MT] (Тестер) Применение формул на листе Динамика цены");
        Lib.recalculatePriceDynamicsFormulas();
      } else {
        Lib.logWarn("[MT] (Тестер) recalculatePriceDynamicsFormulas не найдена");
      }

      if (
        syncResult &&
        syncResult.createdRows &&
        syncResult.createdRows.length
      ) {
        var newArticleRows2 = _handleNewArticles_(syncResult.createdRows, "TESTER", processed);

        if (newArticleRows2 && newArticleRows2.length > 0) {
          Lib.logInfo(
            "[MT] (Тестер) Построчная синхро для новых артикулов: " +
              newArticleRows2.length
          );
          SpreadsheetApp.flush();
          if (typeof Lib.syncMultipleRows === "function") {
            Lib.syncMultipleRows(newArticleRows2);
          } else {
            Lib.logWarn("[MT] syncMultipleRows не найдена");
          }

          if (
            newArticleRows2.articleCodes &&
            newArticleRows2.articleCodes.length &&
            typeof _setPriceFromProcessedByArticle_ === "function"
          ) {
            newArticleRows2.articleCodes.forEach(function(articleCode) {
              if (articleCode) {
                _setPriceFromProcessedByArticle_(
                  articleCode,
                  processed,
                  newArticleRows2.articlePriceMap
                );
              }
            });
          }

          // Заполняем ID-P для новых строк после синхронизации
          if (typeof Lib.fillIdpOnSheetsByIdFromPrimary === 'function') {
            Lib.logInfo("[MT] (Тестер) Заполнение ID-P для новых артикулов на всех листах");
            Lib.fillIdpOnSheetsByIdFromPrimary();
          }
        }
      }

      Lib.logInfo(
        "[MT] Обработка Тестер: завершено, строк " + processed.rows.length
      );

      var message2 = "Обработка прайс завершена. Обновлено строк: " + processed.rows.length + ".";

      // Добавляем информацию о несовпадениях BAR CODE
      if (syncResult && syncResult.barcodeMismatches && syncResult.barcodeMismatches.length > 0) {
        message2 += "\n\n⚠️ Обнаружены несовпадения BAR CODE (" + syncResult.barcodeMismatches.length + "):\n";
        for (var bm2 = 0; bm2 < Math.min(syncResult.barcodeMismatches.length, 10); bm2++) {
          var mismatch2 = syncResult.barcodeMismatches[bm2];
          message2 += "\n• " + mismatch2.article + ":\n  База: " + mismatch2.existing + "\n  Прайс: " + mismatch2.newValue;
        }
        if (syncResult.barcodeMismatches.length > 10) {
          message2 += "\n\n... и ещё " + (syncResult.barcodeMismatches.length - 10) + " несовпадений.";
        }
        message2 += "\n\nПроверьте журнал для подробностей.";
      }

      // SUCCESS: Function completed successfully
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji(
          "Обработка тестеров MT завершена успешно",
          "INFO",
          "",
          "processMtTesterPrice",
          `Обработано тестеров: ${processed.rows.length}, ${syncResult && syncResult.createdRows ? syncResult.createdRows.length + ' новых артикулов' : 'нет новых артикулов'}`,
          "Price",
          "SUCCESS",
          null,
          { rowsProcessed: processed.rows.length, newArticles: syncResult && syncResult.createdRows ? syncResult.createdRows.length : 0, barcodeMismatches: syncResult && syncResult.barcodeMismatches ? syncResult.barcodeMismatches.length : 0 }
        );
      }

      ui.alert(menuTitle, message2, ui.ButtonSet.OK);
    } catch (error) {
      Lib.logError("processMtTesterPrice: ошибка", error);

      // ERROR: Enhanced error logging
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji(
          "Ошибка при обработке тестеров MT",
          "ERROR",
          "",
          "processMtTesterPrice",
          error && error.message ? error.message : String(error),
          "Price",
          "ERROR",
          null,
          { error: error ? error.toString() : "Unknown error", stack: error && error.stack ? error.stack : null }
        );
      }
      ui.alert(
        "Ошибка обработки прайс",
        error.message || String(error),
        ui.ButtonSet.OK
      );
    }
  };

  // ========== ФУНКЦИЯ 3: Пробники ==========
  Lib.processMtSamplesPrice = function () {
    var ui = SpreadsheetApp.getUi();
    var config = _getPrimaryDataConfig_();
    var menuTitle = _getMenuTitle_(config);

    if (!_isActiveProject_()) {
      ui.alert(
        menuTitle,
        "Эта функция доступна только в проекте MT.",
        ui.ButtonSet.OK
      );
      return;
    }

    try {
      // START: Function entry
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji(
          "Обработка пробников MT начата",
          "INFO",
          "",
          "processMtSamplesPrice",
          "Пользователь инициировал обработку пробников для проекта MT",
          "Price",
          "START"
        );
      }

      Lib.logInfo("[MT] Обработка Пробники: старт");
      var source = _getSourceData_(config, "SAMPLES");
      if (!source.values || source.values.length <= 2) {
        ui.alert(
          menuTitle,
          "На листе нет данных для обработки.",
          ui.ButtonSet.OK
        );
        return;
      }

      var processed = _buildProcessedSamplesData_(source.values, config);
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
      _storeProcessedSnapshot_(config, processed, "SAMPLES");

      // Обработка изменений групп (показываем диалоги)
      if (syncResult && syncResult.groupChanges && syncResult.groupChanges.length > 0) {
        _handleGroupChanges_(syncResult.groupChanges);
      }

      _fillIdgForRowsWithoutIdp();

      // Заполняем ID-P на остальных листах (Заказ, Динамика цены, Расчет цены, ABC-Анализ)
      if (typeof Lib.fillIdpOnSheetsByIdFromPrimary === 'function') {
        Lib.logInfo("[MT] (Пробники) Заполнение ID-P на всех листах");
        Lib.fillIdpOnSheetsByIdFromPrimary();
      } else {
        Lib.logWarn("[MT] (Пробники) fillIdpOnSheetsByIdFromPrimary не найдена");
      }

      // Копируем цены из обработанных данных на листы Динамика цены и Расчет цены
      // Для пробников MT цена умножается на 10 (упаковка из 10 шт)
      if (typeof Lib.copyPriceFromPrimaryToSheets === 'function') {
        Lib.logInfo("[MT] (Пробники) Копирование цен на Динамика цены и Расчет цены (цена × 10 для упаковки)");
        Lib.copyPriceFromPrimaryToSheets(processed, null, 10);
      } else {
        Lib.logWarn("[MT] (Пробники) copyPriceFromPrimaryToSheets не найдена");
      }

      // Применяем формулы на листе "Динамика цены" (EXW ALFASPA, Закупочная цена, DDP-МОСКВА, Прирост)
      if (typeof Lib.recalculatePriceDynamicsFormulas === 'function') {
        Lib.logInfo("[MT] (Пробники) Применение формул на листе Динамика цены");
        Lib.recalculatePriceDynamicsFormulas();
      } else {
        Lib.logWarn("[MT] (Пробники) recalculatePriceDynamicsFormulas не найдена");
      }

      if (
        syncResult &&
        syncResult.createdRows &&
        syncResult.createdRows.length
      ) {
        var newArticleRows3 = _handleNewArticles_(syncResult.createdRows, "SAMPLES", processed);

        if (newArticleRows3 && newArticleRows3.length > 0) {
          Lib.logInfo(
            "[MT] (Пробники) Построчная синхро для новых артикулов: " +
              newArticleRows3.length
          );
          SpreadsheetApp.flush();
          if (typeof Lib.syncMultipleRows === "function") {
            Lib.syncMultipleRows(newArticleRows3);
          } else {
            Lib.logWarn("[MT] syncMultipleRows не найдена");
          }

          if (
            newArticleRows3.articleCodes &&
            newArticleRows3.articleCodes.length &&
            typeof _setPriceFromProcessedByArticle_ === "function"
          ) {
            newArticleRows3.articleCodes.forEach(function(articleCode) {
              if (articleCode) {
                _setPriceFromProcessedByArticle_(
                  articleCode,
                  processed,
                  newArticleRows3.articlePriceMap
                );
              }
            });
          }

          // Заполняем ID-P для новых строк после синхронизации
          if (typeof Lib.fillIdpOnSheetsByIdFromPrimary === 'function') {
            Lib.logInfo("[MT] (Пробники) Заполнение ID-P для новых артикулов на всех листах");
            Lib.fillIdpOnSheetsByIdFromPrimary();
          }
        }
      }

      Lib.logInfo(
        "[MT] Обработка Пробники: завершено, строк " + processed.rows.length
      );

      // Обновляем статусы и синхронизируем с ТЗ по статусам
      if (typeof Lib.updateStatusesAfterProcessing === 'function') {
        Lib.updateStatusesAfterProcessing();
      }

      var message3 = "Обработка прайс завершена. Обновлено строк: " + processed.rows.length + ".";

      // Добавляем информацию о несовпадениях BAR CODE
      if (syncResult && syncResult.barcodeMismatches && syncResult.barcodeMismatches.length > 0) {
        message3 += "\n\n⚠️ Обнаружены несовпадения BAR CODE (" + syncResult.barcodeMismatches.length + "):\n";
        for (var bm3 = 0; bm3 < Math.min(syncResult.barcodeMismatches.length, 10); bm3++) {
          var mismatch3 = syncResult.barcodeMismatches[bm3];
          message3 += "\n• " + mismatch3.article + ":\n  База: " + mismatch3.existing + "\n  Прайс: " + mismatch3.newValue;
        }
        if (syncResult.barcodeMismatches.length > 10) {
          message3 += "\n\n... и ещё " + (syncResult.barcodeMismatches.length - 10) + " несовпадений.";
        }
        message3 += "\n\nПроверьте журнал для подробностей.";
      }

      // SUCCESS: Function completed successfully
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji(
          "Обработка пробников MT завершена успешно",
          "INFO",
          "",
          "processMtSamplesPrice",
          `Обработано пробников: ${processed.rows.length}, ${syncResult && syncResult.createdRows ? syncResult.createdRows.length + ' новых артикулов' : 'нет новых артикулов'}`,
          "Price",
          "SUCCESS",
          null,
          { rowsProcessed: processed.rows.length, newArticles: syncResult && syncResult.createdRows ? syncResult.createdRows.length : 0, barcodeMismatches: syncResult && syncResult.barcodeMismatches ? syncResult.barcodeMismatches.length : 0 }
        );
      }

      ui.alert(menuTitle, message3, ui.ButtonSet.OK);
    } catch (error) {
      Lib.logError("processMtSamplesPrice: ошибка", error);

      // ERROR: Enhanced error logging
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji(
          "Ошибка при обработке пробников MT",
          "ERROR",
          "",
          "processMtSamplesPrice",
          error && error.message ? error.message : String(error),
          "Price",
          "ERROR",
          null,
          { error: error ? error.toString() : "Unknown error", stack: error && error.stack ? error.stack : null }
        );
      }

      ui.alert(
        "Ошибка обработки прайс",
        error.message || String(error),
        ui.ButtonSet.OK
      );
    }
  };

  // ========== ФУНКЦИЯ 4: Загрузка остатков ==========
  Lib.loadMtStockData = function () {
    var ui = SpreadsheetApp.getUi();
    var config = _getPrimaryDataConfig_();
    var menuTitle = _getMenuTitle_(config);

    if (!_isActiveProject_()) {
      ui.alert(
        menuTitle,
        "Эта функция доступна только в проекте MT.",
        ui.ButtonSet.OK
      );
      return;
    }

    // START: Function entry
    if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
      Lib.logWithEmoji(
        "Загрузка остатков MT начата",
        "INFO",
        "",
        "loadMtStockData",
        "Пользователь инициировал загрузку остатков для проекта MT",
        "Price",
        "START"
      );
    }

    // Используем универсальную функцию загрузки остатков
    if (typeof Lib.loadStockData === 'function') {
      try {
        Lib.loadStockData('MT');

        // SUCCESS: Stock data loaded successfully
        if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
          Lib.logWithEmoji(
            "Загрузка остатков MT завершена успешно",
            "INFO",
            "",
            "loadMtStockData",
            "Остатки успешно загружены из источника в лист Заказ",
            "Price",
            "SUCCESS",
            null,
            { source: "MT", status: "loaded" }
          );
        }
      } catch (err) {
        if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
          Lib.logWithEmoji(
            "Ошибка при загрузке остатков MT",
            "ERROR",
            "",
            "loadMtStockData",
            err && err.message ? err.message : String(err),
            "Price",
            "ERROR",
            null,
            { error: err ? err.toString() : "Unknown error", source: "MT" }
          );
        }
      }
    } else {
      if (typeof Lib !== 'undefined' && typeof Lib.logWithEmoji === 'function') {
        Lib.logWithEmoji(
          "Функция загрузки остатков недоступна",
          "WARNING",
          "",
          "loadMtStockData",
          "Lib.loadStockData не найдена в библиотеке",
          "Price",
          "ERROR",
          null,
          { source: "MT", status: "library_not_found" }
        );
      }
      ui.alert(
        menuTitle,
        "Функция загрузки остатков недоступна. Обновите библиотеку.",
        ui.ButtonSet.OK
      );
    }
  };

  /**
   * ПУБЛИЧНАЯ: Сортировка заказа по Производителю (Группа + ID-G)
   * Теперь сортирует одновременно три листа: Заказ, Динамика цены, Расчет цены
   */
  Lib.sortMtOrderByManufacturer = function () {
    var ui = SpreadsheetApp.getUi();
    var config = _getPrimaryDataConfig_();
    var menuTitle = _getMenuTitle_(config);

    if (!_isActiveProject_()) {
      ui.alert(menuTitle, "Эта функция доступна только в проекте MT.", ui.ButtonSet.OK);
      return;
    }

    try {
      if (typeof Lib.structureMultipleSheets !== "function") {
        throw new Error("Функция Lib.structureMultipleSheets недоступна. Обновите библиотеку.");
      }
      Lib.structureMultipleSheets('byManufacturer');
    } catch (error) {
      Lib.logError("[MT] Сортировка по Производителю: ошибка", error);
      ui.alert(menuTitle, error.message || String(error), ui.ButtonSet.OK);
    }
  };

  /**
   * ПУБЛИЧНАЯ: Сортировка заказа по Прайсу (Линия + ID-L)
   * Теперь сортирует одновременно три листа: Заказ, Динамика цены, Расчет цены
   */
  Lib.sortMtOrderByPrice = function () {
    var ui = SpreadsheetApp.getUi();
    var config = _getPrimaryDataConfig_();
    var menuTitle = _getMenuTitle_(config);

    if (!_isActiveProject_()) {
      ui.alert(menuTitle, "Эта функция доступна только в проекте MT.", ui.ButtonSet.OK);
      return;
    }

    try {
      if (typeof Lib.structureMultipleSheets !== "function") {
        throw new Error("Функция Lib.structureMultipleSheets недоступна. Обновите библиотеку.");
      }
      Lib.structureMultipleSheets('byPrice');
    } catch (error) {
      Lib.logError("[MT] Сортировка по Прайсу: ошибка", error);
      ui.alert(menuTitle, error.message || String(error), ui.ButtonSet.OK);
    }
  };

  // ========== ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ==========

  function _getPrimaryDataConfig_() {
    var cfg = global.CONFIG && global.CONFIG.PRIMARY_DATA;
    if (!cfg) {
      throw new Error("Конфигурация PRIMARY_DATA не настроена для проекта MT.");
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
      sheetName: fullSheetName,
    };
  }

  function _resolveSheetName_(config, section, key) {
    if (!config || !config.SHEETS || !config.SHEETS[section]) {
      return null;
    }
    return config.SHEETS[section][key] || null;
  }

  // ========== Б/З поставщик: парсинг ==========
  function _buildProcessedMainData_(values, config) {
    var rows = [];
    var articles = [];
    var groups = [];
    var headerRowIndex = 0;
    var currentGroup = "";

    // Заголовки в первой строке
    var headers = values[headerRowIndex];
    var codeIdx = _findColumnIndex(headers, "CODE");
    var categoryIdx = _findColumnIndex(headers, "CATEGORY");
    var descriptionIdx = _findColumnIndex(headers, "DESCRIPTION");
    var formatIdx = _findColumnIndex(headers, "FORMAT");
    var unitsIdx = _findColumnIndex(headers, "UNITS");
    var priceIdx = _findColumnIndex(headers, "PRICE");
    var eanIdx = _findColumnIndex(headers, "EAN");

    if (codeIdx === -1 || descriptionIdx === -1) {
      throw new Error("Не найдены обязательные столбцы CODE и DESCRIPTION");
    }

    for (var i = headerRowIndex + 1; i < values.length; i++) {
      var row = values[i];
      var codeValue = _asTrimmedString(row[codeIdx]);
      var categoryValue = categoryIdx !== -1 ? _asTrimmedString(row[categoryIdx]) : "";
      var descriptionValue = _asTrimmedString(row[descriptionIdx]);

      // Логика определения группы: CODE есть, CATEGORY нет => группа
      if (codeValue && !categoryValue) {
        currentGroup = codeValue;
        continue;
      }

      // Если есть CODE и DESCRIPTION, то это артикул
      if (codeValue && descriptionValue) {
        var formatValue = formatIdx !== -1 ? _getValue(row, formatIdx) : "";
        var unitsValue = unitsIdx !== -1 ? _getValue(row, unitsIdx) : "";
        var priceValue = priceIdx !== -1 ? _getValue(row, priceIdx) : "";
        var eanValue = eanIdx !== -1 ? _getValue(row, eanIdx) : "";

        rows.push([
          "", // ID-P
          codeValue, // Арт. произв.
          descriptionValue, // Название ENG прайс произв
          formatValue, // Объём
          eanValue, // BAR CODE
          unitsValue, // шт./уп.
          priceValue, // Цена
          currentGroup, // Группа
        ]);
        articles.push(codeValue);
        groups.push(currentGroup);
      }
    }

    return {
      headers: MAIN_OUTPUT_HEADERS.slice(),
      rows: rows,
      articles: articles,
      groups: groups,
    };
  }

  // ========== Тестер: парсинг ==========
  function _buildProcessedTesterData_(values, config) {
    var rows = [];
    var articles = [];
    var groups = [];
    var currentGroup = "";

    // Ищем строку с заголовками (первая строка с CODE и DESCRIPTION)
    var headerRowIndex = -1;
    for (var h = 0; h < Math.min(values.length, 5); h++) {
      var testRow = values[h];
      var codeTest = _findColumnIndex(testRow, "CODE");
      var descTest = _findColumnIndex(testRow, "DESCRIPTION");
      if (codeTest !== -1 && descTest !== -1) {
        headerRowIndex = h;
        break;
      }
    }

    if (headerRowIndex === -1) {
      throw new Error("Не найдена строка заголовков с CODE и DESCRIPTION");
    }

    var headers = values[headerRowIndex];
    var codeIdx = _findColumnIndex(headers, "CODE");
    var descriptionIdx = _findColumnIndex(headers, "DESCRIPTION");
    var formatIdx = _findColumnIndex(headers, "FORMAT");
    var unitsIdx = _findColumnIndex(headers, "UNITS");
    var priceIdx = _findColumnIndex(headers, "PRICE");
    var eanIdx = _findColumnIndex(headers, "EAN");

    if (codeIdx === -1 || descriptionIdx === -1) {
      throw new Error("Не найдены обязательные столбцы CODE и DESCRIPTION");
    }

    for (var i = headerRowIndex + 1; i < values.length; i++) {
      var row = values[i];
      var codeValue = _asTrimmedString(row[codeIdx]);
      var descriptionValue = _asTrimmedString(row[descriptionIdx]);

      // Логика определения группы: CODE есть, DESCRIPTION нет => группа
      if (codeValue && !descriptionValue) {
        currentGroup = codeValue;
        continue;
      }

      // Если есть CODE и DESCRIPTION, то это артикул
      if (codeValue && descriptionValue) {
        var formatValue = formatIdx !== -1 ? _getValue(row, formatIdx) : "";
        var unitsValue = unitsIdx !== -1 ? _getValue(row, unitsIdx) : "";
        var priceValue = priceIdx !== -1 ? _getValue(row, priceIdx) : "";
        var eanValue = eanIdx !== -1 ? _getValue(row, eanIdx) : "";

        // Для Тестера: добавляем префикс "-Тестер " к объёму
        var volumeWithPrefix = formatValue ? "-Тестер " + formatValue : "";

        // Для Тестера: добавляем суффикс " - Тестер" к группе
        var groupWithSuffix = currentGroup ? currentGroup + " - Тестер" : " - Тестер";

        rows.push([
          "", // ID-P
          codeValue, // Арт. произв.
          descriptionValue, // Название ENG прайс произв
          volumeWithPrefix, // Объём с префиксом "-Тестер "
          eanValue, // BAR CODE
          unitsValue, // шт./уп.
          priceValue, // Цена
          groupWithSuffix, // Группа с суффиксом " - Тестер"
        ]);
        articles.push(codeValue);
        groups.push(groupWithSuffix);
      }
    }

    return {
      headers: TESTER_OUTPUT_HEADERS.slice(),
      rows: rows,
      articles: articles,
      groups: groups,
    };
  }

  // ========== Пробники: парсинг ==========
  function _buildProcessedSamplesData_(values, config) {
    var rows = [];
    var articles = [];
    var groups = [];
    var currentGroup = "";

    // Ищем строку с заголовками (первая строка с CODE и DESCRIPTION)
    var headerRowIndex = -1;
    for (var h = 0; h < Math.min(values.length, 5); h++) {
      var testRow = values[h];
      var codeTest = _findColumnIndex(testRow, "CODE");
      var descTest = _findColumnIndex(testRow, "DESCRIPTION");
      if (codeTest !== -1 && descTest !== -1) {
        headerRowIndex = h;
        break;
      }
    }

    if (headerRowIndex === -1) {
      throw new Error("Не найдена строка заголовков с CODE и DESCRIPTION");
    }

    var headers = values[headerRowIndex];
    var codeIdx = _findColumnIndex(headers, "CODE");
    var descriptionIdx = _findColumnIndex(headers, "DESCRIPTION");
    var formatIdx = _findColumnIndex(headers, "FORMAT");
    var unitsIdx = _findColumnIndex(headers, "UNITS");
    var priceIdx = _findColumnIndex(headers, "PRICE");
    var eanIdx = _findColumnIndex(headers, "EAN");

    if (codeIdx === -1 || descriptionIdx === -1) {
      throw new Error("Не найдены обязательные столбцы CODE и DESCRIPTION");
    }

    for (var i = headerRowIndex + 1; i < values.length; i++) {
      var row = values[i];
      var codeValue = _asTrimmedString(row[codeIdx]);
      var descriptionValue = _asTrimmedString(row[descriptionIdx]);

      // Логика определения группы: CODE есть, DESCRIPTION нет => группа
      if (codeValue && !descriptionValue) {
        currentGroup = codeValue;
        continue;
      }

      // Если есть CODE и DESCRIPTION, то это артикул
      if (codeValue && descriptionValue) {
        var formatValue = formatIdx !== -1 ? _getValue(row, formatIdx) : "";
        var unitsValue = unitsIdx !== -1 ? _getValue(row, unitsIdx) : "";
        var priceValue = priceIdx !== -1 ? _getValue(row, priceIdx) : "";
        var eanValue = eanIdx !== -1 ? _getValue(row, eanIdx) : "";

        // Для Пробников: добавляем суффикс " - пробник" к группе
        var groupWithSuffix = currentGroup ? currentGroup + " - пробник" : " - пробник";

        rows.push([
          "", // ID-P
          codeValue, // Арт. произв.
          descriptionValue, // Название ENG прайс произв
          formatValue, // Объём (для пробников будет запрашиваться вручную)
          eanValue, // BAR CODE
          unitsValue, // шт./уп.
          priceValue, // Цена
          groupWithSuffix, // Группа с суффиксом " - пробник"
        ]);
        articles.push(codeValue);
        groups.push(groupWithSuffix);
      }
    }

    return {
      headers: SAMPLES_OUTPUT_HEADERS.slice(),
      rows: rows,
      articles: articles,
      groups: groups,
    };
  }

  function _findColumnIndex(headers, keyword) {
    for (var i = 0; i < headers.length; i++) {
      var header = String(headers[i] || "").trim().toUpperCase();
      if (header.indexOf(keyword.toUpperCase()) !== -1) {
        return i;
      }
    }
    return -1;
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
      baseSheetName = resolvedKey === "TESTER" ? "-Тестер" : (resolvedKey === "SAMPLES" ? "-Пробники" : "-Б/З поставщик");
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

  function _syncIdWithMain(sourceRows, sourceHeaders, articles, groups) {
    var emptyResult = { createdRows: [], assignedIdp: {} };
    if (!global.CONFIG || !global.CONFIG.SHEETS) {
      return emptyResult;
    }
    try {
      Lib.logDebug(
        "[MT] syncIdWithMain: старт, строк " + (articles ? articles.length : 0)
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
      var groupColumnIndex = headerToIndex[_normalizeHeaderKey("Группа")];

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
          if (groupExisting) {
            groupIdMap[groupExisting] = existingIdG;
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
      var barcodeMismatches = []; // Массив для сбора несовпадений BAR CODE
      var groupChanges = []; // Массив для изменений в столбце "Группа"

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
            if (groupIdMap[groupValueInput]) {
              newIdG = groupIdMap[groupValueInput];
            } else {
              newIdG = ++idgCounter;
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
            "[MT] syncIdWithMain: назначен новый ID-P " +
              assignedIdP +
              " для артикула " +
              articleCode
          );
        } else if (assignedIdP !== null) {
          Lib.logDebug(
            "[MT] syncIdWithMain: переиспользован ID-P " +
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

        var srcRow = sourceRows && sourceRows[j];
        var isNewArticle = !entry.rowNumber; // Новый артикул, если нет номера строки

        if (srcRow) {
          // Проверяем BAR CODE для существующих артикулов
          var barcodeIdx = headerToIndex[_normalizeHeaderKey("BAR CODE")];
          if (!isNewArticle && barcodeIdx !== undefined) {
            var existingBarcode = _asTrimmedString(rowValues[barcodeIdx]);
            var sourceBarcodeIdx = sourceHeaderToIndex[_normalizeHeaderKey("BAR CODE")];
            var newBarcode = sourceBarcodeIdx !== undefined ? _asTrimmedString(srcRow[sourceBarcodeIdx]) : "";

            if (existingBarcode && newBarcode && existingBarcode !== newBarcode) {
              Lib.logWarn(
                "[MT] ⚠️ НЕСОВПАДЕНИЕ BAR CODE для артикула " + articleCode +
                ": в базе '" + existingBarcode + "', в прайсе '" + newBarcode + "'"
              );
              barcodeMismatches.push({
                article: articleCode,
                existing: existingBarcode,
                newValue: newBarcode
              });
            }
          }

          // Для существующих артикулов НЕ обновляем данные из прайса
          if (isNewArticle) {
            for (var name in sourceHeaderToIndex) {
              if (!sourceHeaderToIndex.hasOwnProperty(name)) continue;
              var targetIndex = headerToIndex[name];
              if (targetIndex === undefined || targetIndex >= targetColumnCount)
                continue;
              var srcIdx = sourceHeaderToIndex[name];
              rowValues[targetIndex] = srcRow[srcIdx] || "";
            }
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

        if (groupColumnIndex !== undefined && groupValueInput) {
          var existingGroupValue = _asTrimmedString(rowValues[groupColumnIndex]);

          // Для НОВЫХ артикулов - сразу заполняем группу
          if (isNewArticle) {
            rowValues[groupColumnIndex] = groupValueInput;
          } else {
            // Для существующих артикулов:
            // 1. Если столбец "Группа" пустой - заполняем без вопросов
            if (!existingGroupValue) {
              rowValues[groupColumnIndex] = groupValueInput;
              Lib.logInfo(
                "[MT] Заполнена пустая Группа '" + groupValueInput + "' для артикула " + articleCode
              );
            }
            // 2. Если группа заполнена и отличается - добавляем в список изменений
            else if (existingGroupValue !== groupValueInput) {
              groupChanges.push({
                articleCode: articleCode,
                rowNumber: entry.rowNumber,
                oldValue: existingGroupValue,
                newValue: groupValueInput
              });
              Lib.logInfo(
                "[MT] Обнаружено изменение Группы для артикула " + articleCode +
                ": в базе '" + existingGroupValue + "', в прайсе '" + groupValueInput + "'"
              );
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
          var currentIdgValue = rowValues[idgColumnIndex];
          var mapKey = resolvedGroup;

          if (
            currentIdgValue === undefined ||
            currentIdgValue === null ||
            currentIdgValue === ""
          ) {
            var mappedIdg = groupIdMap[mapKey];
            if (!mappedIdg) {
              mappedIdg = ++idgCounter;
              groupIdMap[mapKey] = mappedIdg;
            }
            currentIdgValue = mappedIdg;
            rowValues[idgColumnIndex] = currentIdgValue;
          }
          if (!groupIdMap[mapKey]) {
            groupIdMap[mapKey] = currentIdgValue;
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
            "[MT] syncIdWithMain: обновление строки " +
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
          "[MT] syncIdWithMain: вставлено " +
            rowsValues.length +
            " новых строк"
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
            "[MT] syncIdWithMain: запись ID-P " +
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
        "[MT] syncIdWithMain: завершено, создано новых строк: " +
          createdEntries.length
      );
      return {
        createdRows: createdEntries,
        assignedIdp: assignedIdpMap,
        barcodeMismatches: barcodeMismatches,
        groupChanges: groupChanges
      };
    } catch (syncError) {
      Lib.logError("syncIdWithMain: ошибка", syncError);
      return emptyResult;
    }
  }

  function _handleGroupChanges_(groupChanges) {
    if (!groupChanges || !groupChanges.length) {
      return;
    }

    Lib.logInfo(
      "[MT] _handleGroupChanges_: обнаружено " +
        groupChanges.length +
        " изменений в столбце Группа"
    );

    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var mainSheetName =
        global.CONFIG && global.CONFIG.SHEETS && global.CONFIG.SHEETS.PRIMARY;
      var mainSheet = mainSheetName ? ss.getSheetByName(mainSheetName) : null;
      if (!mainSheet) {
        Lib.logError("[MT] Не найден лист Главная");
        return;
      }

      var lastColumn = mainSheet.getLastColumn();
      if (lastColumn <= 0) {
        Lib.logError("[MT] Нет столбцов на листе Главная");
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
        Lib.logError("[MT] Не найден столбец Группа");
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
            "[MT] Группа заменена для артикула " +
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
              "[MT] Очищен ID-G для строки " +
                change.rowNumber +
                " (группа изменена)"
            );
          }
        } else {
          Lib.logInfo(
            "[MT] Группа оставлена без изменений для артикула " +
              change.articleCode
          );
        }
      });

      SpreadsheetApp.flush();
    } catch (err) {
      Lib.logError("_handleGroupChanges_: ошибка", err);
    }
  }

  function _handleNewArticles_(createdEntries, processType, processed) {
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
      var categoryIdx = headerToIndex[_normalizeHeaderKey("Категория товара")];
      if (categoryIdx === undefined) categoryIdx = -1;
      var volumeIdx = headerToIndex[_normalizeHeaderKey("Объём")];
      if (volumeIdx === undefined) volumeIdx = -1;
      var nameIdx = headerToIndex[_normalizeHeaderKey("Название ENG прайс произв")];
      if (nameIdx === undefined) nameIdx = -1;
      var articleIdx = headerToIndex[_normalizeHeaderKey("Арт. произв.")];
      if (articleIdx === undefined) articleIdx = -1;

      // Заготавливаем карту Артикул -> Цена из обработанных данных
      var processedArticleIndex = -1;
      var processedPriceIndex = -1;
      var processedPriceMap = {};
      var processedPriceDisplayMap = {};
      var hasProcessedPrice = false;

      if (
        processed &&
        processed.headers &&
        processed.rows &&
        processed.rows.length
      ) {
        processedArticleIndex = processed.headers.indexOf("Арт. произв.");
        processedPriceIndex = processed.headers.indexOf("Цена");

        if (processedArticleIndex !== -1 && processedPriceIndex !== -1) {
          for (var pr = 0; pr < processed.rows.length; pr++) {
            var processedRow = processed.rows[pr];
            var processedArticle = String(
              processedRow[processedArticleIndex] || ""
            ).trim();
            if (!processedArticle) {
              continue;
            }
            var rawPrice = processedRow[processedPriceIndex];
            if (rawPrice === "" || rawPrice === null || rawPrice === undefined) {
              continue;
            }

            var numericPrice = null;
            if (typeof rawPrice === "number") {
              numericPrice = rawPrice;
            } else {
              var rawString = String(rawPrice).replace(/\s+/g, "").replace(",", ".");
              var parsed = parseFloat(rawString);
              if (!isNaN(parsed)) {
                numericPrice = parsed;
              }
            }

            if (numericPrice === null || isNaN(numericPrice)) {
              continue;
            }

            processedPriceMap[processedArticle] = numericPrice;
            processedPriceDisplayMap[processedArticle] = String(
              typeof rawPrice === "number" ? numericPrice.toFixed(2) : rawPrice
            ).trim();
            hasProcessedPrice = true;
          }
        }
      }

      var processedRows = [];
      var processedArticles = [];

      createdEntries.forEach(function (entry) {
        if (!entry || !entry.rowNumber) {
          return;
        }

        var primaryId = _asTrimmedString(
          mainSheet.getRange(entry.rowNumber, 1).getValue()
        );
        if (primaryId && typeof Lib.ensureRowExistsOnBaseSheets === "function") {
          Lib.ensureRowExistsOnBaseSheets(primaryId);
        }

        if (statusIdx !== undefined && statusIdx !== -1) {
          Lib.logDebug("[MT] Установка статуса 'New завод' для строки " + entry.rowNumber);
          mainSheet
            .getRange(entry.rowNumber, statusIdx + 1)
            .setValue("New завод");
        }

        // Устанавливаем категорию в зависимости от типа обработки
        if (categoryIdx !== -1) {
          if (processType === "TESTER") {
            mainSheet.getRange(entry.rowNumber, categoryIdx + 1).setValue("Тестер");
            Lib.logDebug("[MT] Установлена категория 'Тестер' для строки " + entry.rowNumber);
          } else if (processType === "SAMPLES") {
            mainSheet.getRange(entry.rowNumber, categoryIdx + 1).setValue("Пробники");
            Lib.logDebug("[MT] Установлена категория 'Пробники' для строки " + entry.rowNumber);
          }
        }

        // Для Пробников запрашиваем объём вручную
        if (processType === "SAMPLES" && volumeIdx !== -1) {
          var productName = nameIdx !== -1 ? _asTrimmedString(mainSheet.getRange(entry.rowNumber, nameIdx + 1).getValue()) : "";
          var currentVolume = volumeIdx !== -1 ? _asTrimmedString(mainSheet.getRange(entry.rowNumber, volumeIdx + 1).getValue()) : "";

          var volumeValue = _promptVolumeInputForSamples_(productName, currentVolume);
          if (volumeValue !== null) {
            mainSheet.getRange(entry.rowNumber, volumeIdx + 1).setValue(volumeValue);
            Lib.logDebug("[MT] Установлен объём для пробника: '" + volumeValue + "' для строки " + entry.rowNumber);
          }
        }

        // Запрашиваем цену EXW для всех новых артикулов
        var productName = entry.productName
          ? _asTrimmedString(entry.productName)
          : nameIdx !== -1
          ? _asTrimmedString(mainSheet.getRange(entry.rowNumber, nameIdx + 1).getValue())
          : "Новый артикул";

        var articleCode = entry.articleCode
          ? _asTrimmedString(entry.articleCode)
          : articleIdx !== -1
          ? _asTrimmedString(mainSheet.getRange(entry.rowNumber, articleIdx + 1).getValue())
          : "";

        if (articleCode && processedArticles.indexOf(articleCode) === -1) {
          processedArticles.push(articleCode);
        }

        var promptPriceSource =
          articleCode && processedPriceDisplayMap.hasOwnProperty(articleCode)
            ? processedPriceDisplayMap[articleCode]
            : "";

        var priceValue = _promptPriceInput_(productName, promptPriceSource);
        if (priceValue !== null && primaryId) {
          // Применяем изменения перед записью цены на другие листы
          SpreadsheetApp.flush();

          // Записываем цену на листы Динамика цены и Расчет цены
          _setPriceOnDynamicsAndCalculationSheets_(primaryId, priceValue);

          Lib.logInfo("[MT] Установлена цена EXW " + priceValue + " € для артикула ID=" + primaryId);
        }

        processedRows.push(entry.rowNumber);
      });

      if (processedRows.length && processedArticles.length) {
        processedRows.articleCodes = processedArticles;
        if (hasProcessedPrice) {
          processedRows.articlePriceMap = processedPriceMap;
        }
      }

      SpreadsheetApp.flush();
      return processedRows;
    } catch (err) {
      Lib.logError("_handleNewArticles_: ошибка", err);
      return [];
    }
  }

  function _promptVolumeInputForSamples_(productName, currentVolume) {
    var ui = SpreadsheetApp.getUi();
    var promptText = 'Наименование: "' + productName + '"\nОбъём: "' + currentVolume + '"\n\nВведите объём для пробника или нажмите "Отмена", чтобы оставить текущее значение.';

    var response = ui.prompt(
      'Объём пробника',
      promptText,
      ui.ButtonSet.OK_CANCEL
    );

    var button = response.getSelectedButton();
    if (button !== ui.Button.OK) {
      return null;
    }

    var text = _asTrimmedString(response.getResponseText());
    return text || null;
  }

  /**
   * Запрашивает ввод цены для нового артикула
   * @param {string} productName - название продукта
   * @param {number|string} currentPrice - текущая цена из Б/З
   * @returns {number|null} - введённая цена или null если отменено
   */
  function _promptPriceInput_(productName, currentPrice) {
    var ui = SpreadsheetApp.getUi();
    var priceDisplay;
    if (typeof currentPrice === 'number' && !isNaN(currentPrice)) {
      priceDisplay = currentPrice.toFixed(2);
    } else if (currentPrice !== '' && currentPrice !== null && currentPrice !== undefined) {
      priceDisplay = String(currentPrice);
    } else {
      priceDisplay = 'не указана';
    }

    var priceLine = priceDisplay === 'не указана'
      ? 'Цена не указана'
      : 'Цена ' + priceDisplay + ' €';

    var promptText = 'Наименование: "' + productName + '"\n' + priceLine + '\n\nВведите цену EXW (в евро) или нажмите "Отмена", чтобы пропустить.';

    var response = ui.prompt(
      'Цена EXW для нового артикула',
      promptText,
      ui.ButtonSet.OK_CANCEL
    );

    var button = response.getSelectedButton();
    if (button !== ui.Button.OK) {
      return null;
    }

    var text = _asTrimmedString(response.getResponseText());
    if (!text) {
      return null;
    }

    // Преобразуем в число
    var priceNum = parseFloat(text.replace(',', '.'));
    if (isNaN(priceNum)) {
      ui.alert('Ошибка', 'Введено некорректное значение цены. Цена не будет установлена.', ui.ButtonSet.OK);
      return null;
    }

    return priceNum;
  }

  /**
   * Записывает цену EXW (введенную пользователем) на листы Динамика цены и Расчет цены
   * Записывается в столбцы: "ЦЕНА EXW из Б/З, €", "EXW YYYY, €" и "EXW текущая, €"
   * @param {string} primaryId - ID артикула
   * @param {number} price - цена EXW (введенная пользователем)
   */
  function _setPriceOnDynamicsAndCalculationSheets_(primaryId, price) {
    if (!primaryId || price === null || price === undefined) {
      return;
    }

    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var currentYear = new Date().getFullYear();

      // Определяем целевые листы и столбцы
      // Цена из диалога записывается в годовые столбцы И в "ЦЕНА EXW из Б/З, €"
      var targets = [
        {
          sheetName: Lib.CONFIG.SHEETS.PRICE_DYNAMICS,
          columns: ['ЦЕНА EXW из Б/З, €', 'EXW ' + currentYear + ', €']
        },
        {
          sheetName: Lib.CONFIG.SHEETS.PRICE_CALCULATION,
          columns: ['ЦЕНА EXW из Б/З, €', 'EXW текущая, €']
        }
      ];

      targets.forEach(function(target) {
        var sheet = ss.getSheetByName(target.sheetName);
        if (!sheet) {
          Lib.logWarn('[MT] _setPriceOnDynamicsAndCalculationSheets_: лист "' + target.sheetName + '" не найден');
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

        var idIndex = headers.indexOf('ID');
        if (idIndex === -1) {
          Lib.logWarn('[MT] _setPriceOnDynamicsAndCalculationSheets_: столбец ID не найден на листе "' + target.sheetName + '"');
          return;
        }

        // Ищем строку с нужным ID
        var idData = sheet.getRange(2, idIndex + 1, lastRow - 1, 1).getValues();
        var targetRow = -1;
        for (var i = 0; i < idData.length; i++) {
          if (String(idData[i][0] || '').trim() === primaryId) {
            targetRow = i + 2; // +2 для учёта заголовка
            break;
          }
        }

        if (targetRow === -1) {
          Lib.logWarn('[MT] _setPriceOnDynamicsAndCalculationSheets_: строка с ID "' + primaryId + '" не найдена на листе "' + target.sheetName + '"');
          return;
        }

        // Записываем цену в указанные столбцы с форматом "0.00 €"
        target.columns.forEach(function(columnName) {
          var colIndex = headers.indexOf(columnName);
          if (colIndex !== -1) {
            var range = sheet.getRange(targetRow, colIndex + 1);
            range.setValue(price);
            range.setNumberFormat('0.00 "€"');
            Lib.logInfo('[MT] Установлена цена ' + price + ' € в столбец "' + columnName + '" на листе "' + target.sheetName + '" для ID ' + primaryId);
          } else {
            Lib.logDebug('[MT] Столбец "' + columnName + '" не найден на листе "' + target.sheetName + '"');
          }
        });
      });
    } catch (err) {
      Lib.logError('[MT] _setPriceOnDynamicsAndCalculationSheets_: ошибка', err);
    }
  }

  /**
   * Записывает цену из исходного документа (processed) в "ЦЕНА EXW из Б/З"
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
        Lib.logWarn('[MT] _setPriceFromProcessedByArticle_: не найдены необходимые столбцы в processed');
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
        Lib.logDebug('[MT] _setPriceFromProcessedByArticle_: цена не найдена для артикула ' + article);
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
        Lib.logInfo('[MT] Установлена цена ' + price + ' € из исходного документа в "ЦЕНА EXW из Б/З" на листе "' + sheetName + '" для артикула ' + article);
      });
    } catch (err) {
      Lib.logError('[MT] _setPriceFromProcessedByArticle_: ошибка', err);
    }
  }

  function _applyAssignedIdpToProcessed_(processed, assignedMap) {
    if (!processed || !processed.rows || !assignedMap) {
      Lib.logDebug("[MT] _applyAssignedIdpToProcessed_: пропуск - нет данных");
      return;
    }
    var headers = processed.headers || [];
    var idpIndex = headers.indexOf("ID-P");
    var articleIndex = headers.indexOf("Арт. произв.");

    Lib.logDebug(
      "[MT] _applyAssignedIdpToProcessed_: idpIndex=" +
        idpIndex +
        ", articleIndex=" +
        articleIndex
    );

    if (idpIndex === -1) {
      idpIndex = 0;
    }
    if (articleIndex === -1) {
      articleIndex = 1;
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
      "[MT] _applyAssignedIdpToProcessed_: обновлено строк " + updatedCount
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
      Lib.logDebug("[MT] _clearIdpColumnOnMain: очищен столбец ID-P");
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
      Lib.logDebug("[MT] _clearIdgColumnOnMain: очищен столбец ID-G");
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
          "[MT] _fillIdgForRowsWithoutIdp: не найдены необходимые столбцы"
        );
        return;
      }

      var allData = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();

      var groupToIdgMap = {};
      for (var i = 0; i < allData.length; i++) {
        var row = allData[i];
        var idpValue = row[idpIndex];
        var idgValue = row[idgIndex];
        var groupValue = _asTrimmedString(row[groupIndex]);

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
        "[MT] _fillIdgForRowsWithoutIdp: создана карта Группа->ID-G, записей: " +
          Object.keys(groupToIdgMap).length
      );

      var updatedCount = 0;
      for (var j = 0; j < allData.length; j++) {
        var currentRow = allData[j];
        var currentIdp = currentRow[idpIndex];
        var currentIdg = currentRow[idgIndex];
        var currentGroup = _asTrimmedString(currentRow[groupIndex]);

        if (
          (currentIdp === null || currentIdp === undefined || currentIdp === "") &&
          currentGroup &&
          groupToIdgMap[currentGroup]
        ) {
          sheet.getRange(j + 2, idgIndex + 1).setValue(groupToIdgMap[currentGroup]);
          updatedCount++;
          Lib.logDebug(
            "[MT] _fillIdgForRowsWithoutIdp: строка " +
              (j + 2) +
              ", Группа '" +
              currentGroup +
              "' -> ID-G " +
              groupToIdgMap[currentGroup]
          );
        }
      }

      Lib.logInfo(
        "[MT] _fillIdgForRowsWithoutIdp: завершено, обновлено строк: " + updatedCount
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
})(Lib, this);
