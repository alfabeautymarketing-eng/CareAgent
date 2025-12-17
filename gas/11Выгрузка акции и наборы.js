var Lib = Lib || {};

/**
 * =======================================================================================
 * МОДУЛЬ: ВЫГРУЗКА АКЦИЙ И НАБОРОВ (PromotionsExport)
 * ---------------------------------------------------------------------------------------
 * Назначение: Выгрузка данных об акциях и наборах из листа "Заказ" в целевые документы
 *
 * Функции:
 * 1. Выгрузка акций для проектов SK, SS, MT
 * 2. Выгрузка наборов для проектов SK, SS, MT
 * 3. Создание листов с названием следующего месяца
 * 4. Заполнение данных с учетом связей между листами
 * =======================================================================================
 */

(function (Lib, global) {
  /**
   * КОНСТАНТЫ: ID целевых документов для каждого проекта
   */
  var TARGET_DOCS = {
    SK: '1YkGP-1Ipn7qLMKJyxLtm3ATrOhCxO2OuFLMt5WK8tsg',
    SS: '1Q20jk9Cy8gIEJyKQ2-Ph34qqX3Y_oEdKAOK-o_oaFHQ',
    MT: '140vuIAJ1dcuAoc10T5EnIFjx1lUq7e7oroBJlBs1TDA'
  };

  /**
   * КОНСТАНТЫ: Заголовки для листов акций и наборов
   */
  var PROMOTIONS_HEADERS = [
    'ID',
    'Статус',
    'Арт. Рус',
    'Арт. произв.',
    'Наименование ДС',
    'Объём',
    'Категория товара',
    'Закупочная цена, ₽',
    'DDP  -МОСКВА,  ₽',
    'Утвержденная цена ОПТ ₽',
    'Утвержденная РРЦ',
    'Цена дистрибьютора -30%,  ₽',
    'Остаток на текущий день',
    'АКЦИИ',
    'условие',
    'коментарий',
    'Срок'
  ];

  /**
   * ПУБЛИЧНАЯ: Выгрузка акций
   */
  Lib.exportPromotions = function() {
    _exportData('promotions');
  };

  /**
   * ПУБЛИЧНАЯ: Выгрузка наборов
   */
  Lib.exportSets = function() {
    _exportData('sets');
  };

  /**
   * ВНУТРЕННЯЯ: Основная логика выгрузки
   */
  function _exportData(type) {
    var ui = SpreadsheetApp.getUi();
    var projectKey = global.CONFIG && global.CONFIG.ACTIVE_PROJECT_KEY;

    if (!projectKey || !TARGET_DOCS[projectKey]) {
      ui.alert(
        'Ошибка',
        'Проект не определен или не поддерживается для данной операции.',
        ui.ButtonSet.OK
      );
      return;
    }

    try {
      Lib.logInfo('[' + projectKey + '] Выгрузка ' + (type === 'promotions' ? 'акций' : 'наборов') + ': старт');

      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var orderSheet = ss.getSheetByName('Заказ');

      if (!orderSheet) {
        throw new Error('Не найден лист "Заказ"');
      }

      // Получаем данные с листа Заказ
      var orderData = _getOrderSheetData(orderSheet, type);

      if (!orderData || !orderData.rows || orderData.rows.length === 0) {
        ui.alert(
          'Информация',
          'Не найдено данных для выгрузки ' + (type === 'promotions' ? 'акций' : 'наборов') + '.',
          ui.ButtonSet.OK
        );
        return;
      }

      // Получаем данные с листов Главная и Сертификация для заполнения
      var mainData = _getMainSheetData(ss);
      var certData = _getCertificationSheetData(ss);

      // Обогащаем данные
      var enrichedRows = _enrichOrderData(orderData.rows, mainData, certData);

      // Создаем/обновляем лист в целевом документе
      var targetDocId = TARGET_DOCS[projectKey];
      var sheetName = _createTargetSheetName(type);

      _writeToTargetDocument(targetDocId, sheetName, enrichedRows);

      Lib.logInfo('[' + projectKey + '] Выгрузка завершена: записано строк ' + enrichedRows.length);
      ui.alert(
        'Успех',
        'Выгрузка ' + (type === 'promotions' ? 'акций' : 'наборов') + ' завершена.\nЗаписано строк: ' + enrichedRows.length,
        ui.ButtonSet.OK
      );

    } catch (error) {
      Lib.logError('[' + projectKey + '] Ошибка выгрузки', error);
      ui.alert(
        'Ошибка',
        error.message || String(error),
        ui.ButtonSet.OK
      );
    }
  }

  /**
   * ВНУТРЕННЯЯ: Получение данных с листа Заказ
   */
  function _getOrderSheetData(orderSheet, type) {
    var lastRow = orderSheet.getLastRow();
    var lastColumn = orderSheet.getLastColumn();

    if (lastRow <= 1) {
      return null;
    }

    var headerInfo = _findHeaderRow(orderSheet, lastColumn, ['ID']);
    if (!headerInfo) {
      throw new Error('Не удалось определить строку заголовков на листе Заказ');
    }

    var headerRow = headerInfo.row;
    var headers = headerInfo.headers;

    // Определяем индексы нужных столбцов
    var columnName = type === 'promotions' ? 'АКЦИИ' : 'Набор';
    var conditionColumn = type === 'promotions' ? 'условие#' : 'условие';
    var commentColumn = type === 'promotions' ? 'коментарий#' : 'коментарий';
    var termColumn = type === 'promotions' ? 'Срок#' : 'Срок';

    var indices = {
      id: headers.indexOf('ID'),
      status: headers.indexOf('Статус'),
      rusArt: headers.indexOf('Арт. Рус'),
      prodArt: headers.indexOf('Арт. произв.'),
      volume: headers.indexOf('Объём'),
      mainColumn: headers.indexOf(columnName),
      condition: headers.indexOf(conditionColumn),
      comment: headers.indexOf(commentColumn),
      term: headers.indexOf(termColumn),
      stock: headers.indexOf('Остаток')
    };

    if (indices.mainColumn === -1) {
      throw new Error('Не найден столбец "' + columnName + '" на листе Заказ');
    }

    // Читаем данные
    var dataRange = orderSheet.getRange(headerRow + 1, 1, lastRow - headerRow, lastColumn);
    var values = dataRange.getValues();

    var rows = [];
    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      var mainValue = row[indices.mainColumn];

      // Проверяем, есть ли значение в основном столбце
      if (mainValue !== null && mainValue !== undefined && String(mainValue).trim() !== '') {
        rows.push({
          id: row[indices.id] || '',
          status: row[indices.status] || '',
          rusArt: row[indices.rusArt] || '',
          prodArt: row[indices.prodArt] || '',
          volume: row[indices.volume] || '',
          mainValue: mainValue,
          condition: indices.condition !== -1 ? (row[indices.condition] || '') : '',
          comment: indices.comment !== -1 ? (row[indices.comment] || '') : '',
          term: indices.term !== -1 ? (row[indices.term] || '') : '',
          stock: indices.stock !== -1 ? (row[indices.stock] || '') : ''
        });
      }
    }

    return { rows: rows, indices: indices };
  }

  /**
   * ВНУТРЕННЯЯ: Получение данных с листа Главная
   */
  function _getMainSheetData(ss) {
    var mainSheet = ss.getSheetByName('Главная');
    if (!mainSheet) {
      return null;
    }

    var lastRow = mainSheet.getLastRow();
    var lastColumn = mainSheet.getLastColumn();

    if (lastRow <= 1) {
      return null;
    }

    var headerInfo = _findHeaderRow(mainSheet, lastColumn, ['ID']);
    if (!headerInfo) {
      return null;
    }

    var headers = headerInfo.headers;
    var idIndex = headers.indexOf('ID');
    var categoryIndex = headers.indexOf('Категория товара');

    if (idIndex === -1 || categoryIndex === -1) {
      return null;
    }

    var dataRange = mainSheet.getRange(headerInfo.row + 1, 1, lastRow - headerInfo.row, lastColumn);
    var values = dataRange.getValues();

    var map = {};
    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      var id = String(row[idIndex] || '').trim();
      if (id) {
        map[id] = {
          category: row[categoryIndex] || ''
        };
      }
    }

    return map;
  }

  /**
   * ВНУТРЕННЯЯ: Получение данных с листа Сертификация
   */
  function _getCertificationSheetData(ss) {
    var certSheet = ss.getSheetByName('Сертификация');
    if (!certSheet) {
      return null;
    }

    var lastRow = certSheet.getLastRow();
    var lastColumn = certSheet.getLastColumn();

    if (lastRow <= 1) {
      return null;
    }

    var headerInfo = _findHeaderRow(certSheet, lastColumn, ['ID']);
    if (!headerInfo) {
      return null;
    }

    var headers = headerInfo.headers;
    var idIndex = headers.indexOf('ID');
    var dsNameIndex = headers.indexOf('Наименование ДС');

    if (idIndex === -1 || dsNameIndex === -1) {
      return null;
    }

    var dataRange = certSheet.getRange(headerInfo.row + 1, 1, lastRow - headerInfo.row, lastColumn);
    var values = dataRange.getValues();

    var map = {};
    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      var id = String(row[idIndex] || '').trim();
      if (id) {
        map[id] = {
          dsName: row[dsNameIndex] || ''
        };
      }
    }

    return map;
  }

  /**
   * ВНУТРЕННЯЯ: Обогащение данных из листа Заказ
   */
  function _enrichOrderData(rows, mainData, certData) {
    var enrichedRows = [];

    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];
      var id = String(row.id || '').trim();

      var dsName = '';
      var category = '';

      if (id && certData && certData[id]) {
        dsName = certData[id].dsName || '';
      }

      if (id && mainData && mainData[id]) {
        category = mainData[id].category || '';
      }

      enrichedRows.push([
        row.id,
        row.status,
        row.rusArt,
        row.prodArt,
        dsName,
        row.volume,
        category,
        '', // Закупочная цена, ₽
        '', // DDP  -МОСКВА,  ₽
        '', // Утвержденная цена ОПТ ₽
        '', // Утвержденная РРЦ
        '', // Цена дистрибьютора -30%,  ₽
        row.stock,
        row.mainValue,
        row.condition,
        row.comment,
        row.term
      ]);
    }

    return enrichedRows;
  }

  /**
   * ВНУТРЕННЯЯ: Создание имени листа с названием следующего месяца
   */
  function _createTargetSheetName(type) {
    var monthNames = global.CONFIG && global.CONFIG.VALUES && global.CONFIG.VALUES.MONTH_NAMES
      ? global.CONFIG.VALUES.MONTH_NAMES
      : [
          'январь', 'февраль', 'март', 'апрель', 'май', 'июнь',
          'июль', 'август', 'сентябрь', 'октябрь', 'ноябрь', 'декабрь'
        ];

    var now = new Date();
    var nextMonth = new Date(now.getFullYear(), now.getMonth() + 1, 1);
    var monthIndex = nextMonth.getMonth();
    var monthName = monthNames[monthIndex];

    var prefix = type === 'promotions' ? 'Акции ' : 'Наборы ';
    return prefix + monthName;
  }

  /**
   * ВНУТРЕННЯЯ: Запись данных в целевой документ
   */
  function _writeToTargetDocument(docId, sheetName, rows) {
    var targetSs = SpreadsheetApp.openById(docId);
    var targetSheet = targetSs.getSheetByName(sheetName);

    // Если лист уже существует, очищаем его
    if (targetSheet) {
      targetSheet.clear();
    } else {
      // Создаем новый лист
      targetSheet = targetSs.insertSheet(sheetName);
    }

    // Записываем заголовки
    targetSheet.getRange(1, 1, 1, PROMOTIONS_HEADERS.length)
      .setValues([PROMOTIONS_HEADERS])
      .setFontWeight('bold');

    // Записываем данные
    if (rows && rows.length > 0) {
      targetSheet.getRange(2, 1, rows.length, PROMOTIONS_HEADERS.length)
        .setValues(rows);
    }

    // Автоматически подбираем ширину столбцов
    for (var i = 1; i <= PROMOTIONS_HEADERS.length; i++) {
      targetSheet.autoResizeColumn(i);
    }
  }

  /**
   * ВСПОМОГАТЕЛЬНАЯ: Поиск строки заголовков
   */
  function _findHeaderRow(sheet, lastColumn, requiredHeaders) {
    if (lastColumn <= 0) return null;

    for (var row = 1; row <= Math.min(10, sheet.getLastRow()); row++) {
      var rowValues = sheet.getRange(row, 1, 1, lastColumn).getValues()[0];
      var found = true;

      for (var i = 0; i < requiredHeaders.length; i++) {
        var required = requiredHeaders[i];
        var foundHeader = false;

        for (var j = 0; j < rowValues.length; j++) {
          if (String(rowValues[j] || '').trim() === required) {
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
        return {
          row: row,
          headers: rowValues.map(function(h) {
            return String(h || '').trim();
          })
        };
      }
    }

    return null;
  }

})(Lib, this);
