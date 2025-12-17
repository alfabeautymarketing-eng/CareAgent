/**
 * =======================================================================================
 * 12MTPriceManager.js — УПРАВЛЕНИЕ ЛИСТОМ ПРАЙС ДЛЯ ПРОЕКТА MT
 * ---------------------------------------------------------------------------------------
 * ОПИСАНИЕ:
 *   Этот модуль отвечает за автоматическое заполнение листа Прайс из листа Заказ.
 *
 *   Основная функция: addOrderItemsToPrice
 *   - Находит строки на листе Заказ где колонка "Добавить в прайс" = "Добавить"
 *   - Берет ID из этих строк
 *   - Ищет ID на листе Главная
 *   - Копирует данные на лист Прайс в соответствующие колонки
 *   - Копирует из листа Заказ: Примечание, Группа линии, Линия Прайс, Продублировать
 *   - Если флаг "Продублировать" установлен, создает дополнительную строку
 *     с пустыми полями "Группа линии" и "Линия Прайс"
 *
 *   Основная функция: removeOrderItemsFromPrice
 *   - Находит строки на листе Заказ где колонка "Добавить в прайс" = "Убрать"
 *   - Удаляет соответствующие строки из листа Прайс по ID
 * =======================================================================================
 */

var Lib = Lib || {};

(function (Lib, global) {

  /**
   * Автоматическое заполнение листа Прайс из листа Заказ
   * Находит строки на листе Заказ где колонка "Добавить в прайс" = "Добавить",
   * берет ID, ищет его на листе Главная и копирует данные на лист Прайс.
   * Копирует также: Примечание, Группа линии, Линия Прайс, Продублировать.
   * Если флаг "Продублировать" = true, создает дополнительную строку с пустыми
   * полями "Группа линии" и "Линия Прайс".
   */
  Lib.addOrderItemsToPrice = function(options) {
    options = options || {};
    var silent = !!options.silent;
    var triggerSource = options.triggerSource || "manual";
    var menuTitle = "Добавить в прайс";
    var ui = null;

    try {
      if (!silent) {
        ui = SpreadsheetApp.getUi();
      }
      Lib.logInfo(
        "[AddToPrice] Старт процесса добавления товаров в прайс (source=" +
          triggerSource +
          (silent ? ", silent" : "") +
          ")"
      );

      var ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) {
        throw new Error("Не удалось получить активную таблицу");
      }

      // Получаем имена листов из конфигурации
      var orderSheetName = Lib.CONFIG && Lib.CONFIG.SHEETS && Lib.CONFIG.SHEETS.ORDER_FORM;
      var primarySheetName = Lib.CONFIG && Lib.CONFIG.SHEETS && Lib.CONFIG.SHEETS.PRIMARY;
      var priceSheetName = Lib.CONFIG && Lib.CONFIG.SHEETS && Lib.CONFIG.SHEETS.PRICE;

      if (!orderSheetName || !primarySheetName || !priceSheetName) {
        throw new Error("Не найдена конфигурация листов");
      }

      // Получаем листы
      var orderSheet = ss.getSheetByName(orderSheetName);
      var primarySheet = ss.getSheetByName(primarySheetName);
      var priceSheet = ss.getSheetByName(priceSheetName);

      if (!orderSheet) {
        throw new Error("Не найден лист: " + orderSheetName);
      }
      if (!primarySheet) {
        throw new Error("Не найден лист: " + primarySheetName);
      }
      if (!priceSheet) {
        throw new Error("Не найден лист: " + priceSheetName);
      }

      // Получаем все данные с листа Заказ
      var orderData = orderSheet.getDataRange().getValues();
      if (orderData.length < 2) {
        throw new Error("Лист Заказ пуст");
      }

      var orderHeaders = orderData[0];

      // Находим индексы нужных колонок на листе Заказ
      var orderIdIndex = _findColumnIndex(orderHeaders, "ID");
      var orderAddToPriceIndex = _findColumnIndex(orderHeaders, "Добавить в прайс");
      var orderNoteIndex = _findColumnIndex(orderHeaders, "Примечание");
      var orderGroupLineIndex = _findColumnIndex(orderHeaders, "Группа линии");
      var orderLinePriceIndex = _findColumnIndex(orderHeaders, "Линия Прайс");
      var orderDuplicateIndex = _findColumnIndex(orderHeaders, "Продублировать");

      if (orderIdIndex === -1) {
        throw new Error("На листе Заказ не найдена колонка 'ID'");
      }
      if (orderAddToPriceIndex === -1) {
        throw new Error("На листе Заказ не найдена колонка 'Добавить в прайс'");
      }

      // Получаем данные с листа Главная
      var primaryData = primarySheet.getDataRange().getValues();
      if (primaryData.length < 2) {
        throw new Error("Лист Главная пуст");
      }

      var primaryHeaders = primaryData[0];

      // Находим индексы колонок на листе Главная
      var primaryIdIndex = _findColumnIndex(primaryHeaders, "ID");
      var primaryStatusIndex = _findColumnIndex(primaryHeaders, "Статус");
      var primaryArtRusIndex = _findColumnIndex(primaryHeaders, "Арт. Рус");
      var primaryCodeBaseIndex = _findColumnIndex(primaryHeaders, "Код база");
      var primaryCategoryIndex = _findColumnIndex(primaryHeaders, "Категория товара");
      var primaryVolumeIndex = _findColumnIndex(primaryHeaders, "Объём");
      var primaryNameRusIndex = _findColumnIndex(primaryHeaders, "Наименования рус по ДС");
      var primaryNameEngIndex = _findColumnIndex(primaryHeaders, "Наименования англ по ДС");

      if (primaryIdIndex === -1) {
        throw new Error("На листе Главная не найдена колонка 'ID'");
      }

      // Получаем данные с листа Прайс
      var priceData = priceSheet.getDataRange().getValues();
      var priceHeaders = priceData[0];

      // Находим индексы колонок на листе Прайс
      var priceIdIndex = _findColumnIndex(priceHeaders, "ID");
      var priceStatusIndex = _findColumnIndex(priceHeaders, "Статус");
      var priceArtRusIndex = _findColumnIndex(priceHeaders, "Арт. Рус");
      var priceCodeBaseIndex = _findColumnIndex(priceHeaders, "Код база");
      var priceCategoryIndex = _findColumnIndex(priceHeaders, "Категория товара");
      var priceVolumeIndex = _findColumnIndex(priceHeaders, "Объём");
      var priceNameIndex = _findColumnIndex(priceHeaders, "Наименование");
      var priceNoteIndex = _findColumnIndex(priceHeaders, "Примечание");
      var priceGroupLineIndex = _findColumnIndex(priceHeaders, "Группа линии");
      var priceLinePriceIndex = _findColumnIndex(priceHeaders, "Линия Прайс");
      var priceDuplicateIndex = _findColumnIndex(priceHeaders, "Продублировать");

      // Создаем Map для быстрого поиска ID на листе Главная
      var primaryIdMap = {};
      for (var i = 1; i < primaryData.length; i++) {
        var id = primaryData[i][primaryIdIndex];
        if (id) {
          primaryIdMap[String(id).trim()] = i;
        }
      }

      // Создаем Set для проверки существующих ID на листе Прайс
      var existingPriceIds = {};
      for (var i = 1; i < priceData.length; i++) {
        var id = priceData[i][priceIdIndex];
        if (id) {
          existingPriceIds[String(id).trim()] = true;
        }
      }

      var addedCount = 0;
      var skippedCount = 0;
      var rowsToAdd = [];

      // Если указана конкретная строка, обрабатываем только её
      var startRow = 1;
      var endRow = orderData.length;
      if (options.orderRow && options.orderRow > 0) {
        // orderRow - это номер строки в Google Sheets (1-based), а orderData - это массив (0-based)
        // Строка заголовков - это строка 1, первая строка данных - строка 2 (индекс 1 в массиве)
        // Поэтому: arrayIndex = sheetRow - 1
        startRow = options.orderRow - 1;
        endRow = options.orderRow;
        Lib.logInfo("[AddToPrice] Обработка только строки " + options.orderRow);
      }

      // Проходим по строкам листа Заказ
      for (var i = startRow; i < endRow; i++) {
        if (_isAddToPriceFlagSet(orderData[i][orderAddToPriceIndex])) {
          var orderId = String(orderData[i][orderIdIndex]).trim();

          if (!orderId) {
            Lib.logWarn("[AddToPrice] Пустой ID в строке " + (i + 1) + " листа Заказ");
            continue;
          }

          // Проверяем, есть ли уже этот ID на листе Прайс
          if (existingPriceIds[orderId]) {
            Lib.logInfo("[AddToPrice] ID " + orderId + " уже существует на листе Прайс, пропускаем");
            skippedCount++;
            continue;
          }

          // Ищем этот ID на листе Главная
          var primaryRowIndex = primaryIdMap[orderId];
          if (primaryRowIndex === undefined) {
            Lib.logWarn("[AddToPrice] ID " + orderId + " не найден на листе Главная");
            skippedCount++;
            continue;
          }

          var primaryRow = primaryData[primaryRowIndex];
          var orderRow = orderData[i];

          // Формируем новую строку для листа Прайс
          var newPriceRow = new Array(priceHeaders.length).fill("");

          // ID
          if (priceIdIndex !== -1) {
            newPriceRow[priceIdIndex] = orderId;
          }

          // Статус
          if (priceStatusIndex !== -1 && primaryStatusIndex !== -1) {
            newPriceRow[priceStatusIndex] = primaryRow[primaryStatusIndex];
          }

          // Арт. Рус
          if (priceArtRusIndex !== -1 && primaryArtRusIndex !== -1) {
            newPriceRow[priceArtRusIndex] = primaryRow[primaryArtRusIndex];
          }

          // Код база
          if (priceCodeBaseIndex !== -1 && primaryCodeBaseIndex !== -1) {
            newPriceRow[priceCodeBaseIndex] = primaryRow[primaryCodeBaseIndex];
          }

          // Категория товара
          if (priceCategoryIndex !== -1 && primaryCategoryIndex !== -1) {
            newPriceRow[priceCategoryIndex] = primaryRow[primaryCategoryIndex];
          }

          // Объём
          if (priceVolumeIndex !== -1 && primaryVolumeIndex !== -1) {
            newPriceRow[priceVolumeIndex] = primaryRow[primaryVolumeIndex];
          }

          // Наименование (Наименования рус по ДС + перенос строки + Наименования англ по ДС)
          if (priceNameIndex !== -1) {
            var nameRus = (primaryNameRusIndex !== -1) ? primaryRow[primaryNameRusIndex] : "";
            var nameEng = (primaryNameEngIndex !== -1) ? primaryRow[primaryNameEngIndex] : "";

            if (nameRus && nameEng) {
              newPriceRow[priceNameIndex] = nameRus + "\n" + nameEng;
            } else if (nameRus) {
              newPriceRow[priceNameIndex] = nameRus;
            } else if (nameEng) {
              newPriceRow[priceNameIndex] = nameEng;
            }
          }

          // Примечание - из листа Заказ
          if (priceNoteIndex !== -1 && orderNoteIndex !== -1) {
            newPriceRow[priceNoteIndex] = orderRow[orderNoteIndex];
          }

          // Группа линии - из листа Заказ
          if (priceGroupLineIndex !== -1 && orderGroupLineIndex !== -1) {
            newPriceRow[priceGroupLineIndex] = orderRow[orderGroupLineIndex];
          }

          // Линия Прайс - из листа Заказ
          if (priceLinePriceIndex !== -1 && orderLinePriceIndex !== -1) {
            newPriceRow[priceLinePriceIndex] = orderRow[orderLinePriceIndex];
          }

          // Продублировать - из листа Заказ
          if (priceDuplicateIndex !== -1 && orderDuplicateIndex !== -1) {
            newPriceRow[priceDuplicateIndex] = orderRow[orderDuplicateIndex];
          }

          rowsToAdd.push(newPriceRow);
          addedCount++;

          // Проверяем флаг "Продублировать"
          var shouldDuplicate = false;
          if (orderDuplicateIndex !== -1) {
            var duplicateValue = orderRow[orderDuplicateIndex];
            // Проверяем на значение true (checkbox) или "правда"
            if (duplicateValue === true || String(duplicateValue).trim().toLowerCase() === "правда") {
              shouldDuplicate = true;
            }
          }

          // Если нужно дублировать, создаем дополнительную строку
          if (shouldDuplicate) {
            var duplicatedRow = newPriceRow.slice(); // Копируем массив

            // Очищаем колонки "Группа линии" и "Линия Прайс" в дублированной строке
            if (priceGroupLineIndex !== -1) {
              duplicatedRow[priceGroupLineIndex] = "";
            }
            if (priceLinePriceIndex !== -1) {
              duplicatedRow[priceLinePriceIndex] = "";
            }

            rowsToAdd.push(duplicatedRow);
            addedCount++;
          }

          // Помечаем ID как добавленный
          existingPriceIds[orderId] = true;
        }
      }

      // Добавляем все новые строки на лист Прайс
      if (rowsToAdd.length > 0) {
        var lastRow = priceSheet.getLastRow();
        priceSheet.getRange(lastRow + 1, 1, rowsToAdd.length, priceHeaders.length).setValues(rowsToAdd);
        Lib.logInfo("[AddToPrice] Добавлено " + addedCount + " товаров на лист Прайс");

        // Автоматически присваиваем ID-L для всех добавленных строк
        if (typeof Lib.autoAssignPriceLineIdForRow === 'function') {
          for (var i = 0; i < rowsToAdd.length; i++) {
            var rowNum = lastRow + 1 + i;
            Lib.autoAssignPriceLineIdForRow(priceSheet, rowNum);
          }
          Lib.logInfo("[AddToPrice] Присвоены ID-L для " + rowsToAdd.length + " строк");
        }
      }

      var message = "Обработка завершена.\n";
      message += "Добавлено товаров: " + addedCount + "\n";
      if (skippedCount > 0) {
        message += "Пропущено (уже есть или не найдено): " + skippedCount;
      }

      Lib.logInfo("[AddToPrice] " + message);
      if (!silent && ui) {
        ui.alert(menuTitle, message, ui.ButtonSet.OK);
      }

      return { added: addedCount, skipped: skippedCount };

    } catch (error) {
      Lib.logError("[AddToPrice] Ошибка при добавлении товаров в прайс", error);
      if (!silent && ui) {
        ui.alert(menuTitle, "Ошибка: " + (error.message || String(error)), ui.ButtonSet.OK);
      }
      return { added: 0, skipped: 0, error: error };
    }
  };

  /**
   * Удаление товаров из листа Прайс на основе флага "Убрать" в колонке "Добавить в прайс"
   * Находит строки на листе Заказ где колонка "Добавить в прайс" = "Убрать",
   * берет ID и удаляет соответствующие строки из листа Прайс
   */
  Lib.removeOrderItemsFromPrice = function(options) {
    options = options || {};
    var silent = !!options.silent;
    var triggerSource = options.triggerSource || "manual";
    var menuTitle = "Удалить из прайса";
    var ui = null;

    try {
      if (!silent) {
        ui = SpreadsheetApp.getUi();
      }
      Lib.logInfo(
        "[RemoveFromPrice] Старт процесса удаления товаров из прайса (source=" +
          triggerSource +
          (silent ? ", silent" : "") +
          ")"
      );

      var ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) {
        throw new Error("Не удалось получить активную таблицу");
      }

      // Получаем имена листов из конфигурации
      var orderSheetName = Lib.CONFIG && Lib.CONFIG.SHEETS && Lib.CONFIG.SHEETS.ORDER_FORM;
      var priceSheetName = Lib.CONFIG && Lib.CONFIG.SHEETS && Lib.CONFIG.SHEETS.PRICE;

      if (!orderSheetName || !priceSheetName) {
        throw new Error("Не найдена конфигурация листов");
      }

      // Получаем листы
      var orderSheet = ss.getSheetByName(orderSheetName);
      var priceSheet = ss.getSheetByName(priceSheetName);

      if (!orderSheet) {
        throw new Error("Не найден лист: " + orderSheetName);
      }
      if (!priceSheet) {
        throw new Error("Не найден лист: " + priceSheetName);
      }

      // Получаем все данные с листа Заказ
      var orderData = orderSheet.getDataRange().getValues();
      if (orderData.length < 2) {
        throw new Error("Лист Заказ пуст");
      }

      var orderHeaders = orderData[0];

      // Находим индексы нужных колонок на листе Заказ
      var orderIdIndex = _findColumnIndex(orderHeaders, "ID");
      var orderAddToPriceIndex = _findColumnIndex(orderHeaders, "Добавить в прайс");

      if (orderIdIndex === -1) {
        throw new Error("На листе Заказ не найдена колонка 'ID'");
      }
      if (orderAddToPriceIndex === -1) {
        throw new Error("На листе Заказ не найдена колонка 'Добавить в прайс'");
      }

      // Получаем данные с листа Прайс
      var priceData = priceSheet.getDataRange().getValues();
      if (priceData.length < 2) {
        Lib.logInfo("[RemoveFromPrice] Лист Прайс пуст, нечего удалять");
        return { removed: 0, notFound: 0 };
      }

      var priceHeaders = priceData[0];

      // Находим индекс колонки ID на листе Прайс
      var priceIdIndex = _findColumnIndex(priceHeaders, "ID");
      if (priceIdIndex === -1) {
        throw new Error("На листе Прайс не найдена колонка 'ID'");
      }

      // Если указана конкретная строка, обрабатываем только её
      var startRow = 1;
      var endRow = orderData.length;
      if (options.orderRow && options.orderRow > 0) {
        // orderRow - это номер строки в Google Sheets (1-based), а orderData - это массив (0-based)
        // Строка заголовков - это строка 1, первая строка данных - строка 2 (индекс 1 в массиве)
        // Поэтому: arrayIndex = sheetRow - 1
        startRow = options.orderRow - 1;
        endRow = options.orderRow;
        Lib.logInfo("[RemoveFromPrice] Обработка только строки " + options.orderRow);
      }

      // Собираем ID из листа Заказ, которые нужно удалить (где флаг = "Убрать")
      var idsToRemove = {};
      for (var i = startRow; i < endRow; i++) {
        if (_isRemoveFromPriceFlagSet(orderData[i][orderAddToPriceIndex])) {
          var orderId = String(orderData[i][orderIdIndex]).trim();
          if (orderId) {
            idsToRemove[orderId] = true;
          }
        }
      }

      if (Object.keys(idsToRemove).length === 0) {
        Lib.logInfo("[RemoveFromPrice] Не найдено ID для удаления");
        if (!silent && ui) {
          ui.alert(menuTitle, "Нет товаров для удаления из прайса", ui.ButtonSet.OK);
        }
        return { removed: 0, notFound: 0 };
      }

      // Находим строки для удаления на листе Прайс (идем с конца, чтобы не сбивались индексы)
      var rowsToDelete = [];
      var notFoundCount = 0;
      for (var i = priceData.length - 1; i >= 1; i--) {
        var priceId = String(priceData[i][priceIdIndex]).trim();
        if (priceId && idsToRemove[priceId]) {
          rowsToDelete.push(i + 1); // +1 потому что индексы Google Sheets начинаются с 1
          delete idsToRemove[priceId]; // Удаляем из списка, чтобы отследить не найденные
        }
      }

      // Подсчитываем ID, которые не были найдены на листе Прайс
      notFoundCount = Object.keys(idsToRemove).length;

      // Удаляем строки (идем с конца, чтобы не сбивались индексы)
      var removedCount = 0;
      for (var i = 0; i < rowsToDelete.length; i++) {
        priceSheet.deleteRow(rowsToDelete[i]);
        removedCount++;
      }

      var message = "Обработка завершена.\n";
      message += "Удалено товаров: " + removedCount + "\n";
      if (notFoundCount > 0) {
        message += "Не найдено на листе Прайс: " + notFoundCount;
      }

      Lib.logInfo("[RemoveFromPrice] " + message);
      if (!silent && ui) {
        ui.alert(menuTitle, message, ui.ButtonSet.OK);
      }

      return { removed: removedCount, notFound: notFoundCount };

    } catch (error) {
      Lib.logError("[RemoveFromPrice] Ошибка при удалении товаров из прайса", error);
      if (!silent && ui) {
        ui.alert(menuTitle, "Ошибка: " + (error.message || String(error)), ui.ButtonSet.OK);
      }
      return { removed: 0, notFound: 0, error: error };
    }
  };

  /**
   * Вспомогательная функция для поиска индекса колонки по имени
   * @private
   */
  function _findColumnIndex(headers, columnName) {
    for (var i = 0; i < headers.length; i++) {
      if (String(headers[i]).trim() === columnName) {
        return i;
      }
    }
    return -1;
  }

  function _normalizePriceLookupKey(value) {
    if (value === null || value === undefined) return "";
    return String(value)
      .replace(/\s+/g, " ")
      .trim()
      .toLowerCase();
  }

  var PRICE_LINE_ID_LOOKUP = (function() {
    var entries = [
      // Проф для лица
      ["Проф для лица", "PURE & BOTANICS", 1],
      ["Проф для лица", "HYALU FEEL", 2],
      ["Проф для лица", "OXYGEN", 3],
      ["Проф для лица", "RE·EQUILIBRIUM", 4],
      ["Проф для лица", "NEUROSENS", 5],
      ["Проф для лица", "VITA C PURE", 6],
      ["Проф для лица", "PROLAGENIST", 7],
      ["Проф для лица", "RETIDERMA", 8],
      ["Проф для лица", "D-WHITE", 9],
      ["Проф для лица", "SKIN EXPERT", 10],
      ["Проф для лица", "SKIN EXPERT ADVANCE", 11],
      ["Проф для лица", "ARÛDE", 12],
      // Проф для тела
      ["Проф для тела", "BODY TREAT", 13],
      ["Проф для тела", "BODY SENSES", 14],
      // Дом для лица
      ["Дом для лица", "PURE & BOTANICS", 15],
      ["Дом для лица", "HYALU FEEL", 16],
      ["Дом для лица", "OXYGEN", 17],
      ["Дом для лица", "RE·EQUILIBRIUM", 18],
      ["Дом для лица", "NEUROSENS", 19],
      ["Дом для лица", "VITA C PURE", 20],
      ["Дом для лица", "RETIDERMA", 21],
      ["Дом для лица", "PROLAGENIST", 22],
      ["Дом для лица", "D-WHITE", 23],
      ["Дом для лица", "SKIN EXPERT", 24],
      ["Дом для лица", "ELIXIR COLLECTION", 25],
      ["Дом для лица", "ARÛDE", 26],
      // Дом для тела
      ["Дом для тела", "SUN AGE", 27],
      ["Дом для тела", "BODY TREAT", 28],
      ["Дом для тела", "BODY SENSES", 29]
    ];
    var map = {};
    for (var i = 0; i < entries.length; i++) {
      var entry = entries[i];
      var groupKey = _normalizePriceLookupKey(entry[0]);
      var lineKey = _normalizePriceLookupKey(entry[1]);
      if (!groupKey) continue;
      if (!map[groupKey]) {
        map[groupKey] = {};
      }
      map[groupKey][lineKey] = entry[2];
    }
    return map;
  })();

  Lib.resolvePriceLineId_ = function(groupLine, linePrice) {
    if (!PRICE_LINE_ID_LOOKUP) return null;
    var groupKey = _normalizePriceLookupKey(groupLine);
    if (!groupKey) return null;
    var lineKey = _normalizePriceLookupKey(linePrice);
    var groupMap = PRICE_LINE_ID_LOOKUP[groupKey];
    if (!groupMap) return null;
    if (Object.prototype.hasOwnProperty.call(groupMap, lineKey)) {
      return groupMap[lineKey];
    }
    return null;
  };

  function _isAddToPriceFlagSet(value) {
    if (value === true) return true;
    if (value === false || value === null || value === undefined) return false;
    var normalized = String(value).trim().toLowerCase();
    return normalized === "добавить";
  }

  function _isRemoveFromPriceFlagSet(value) {
    if (value === false) return true;
    if (value === true || value === null || value === undefined) return false;
    var normalized = String(value).trim().toLowerCase();
    return normalized === "убрать";
  }

})(Lib, this);
