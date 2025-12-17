var Lib = Lib || {};

/**
 * =======================================================================================
 * МОДУЛЬ: АВТОМАТИЧЕСКИЕ ВЫЧИСЛЕНИЯ ДЛЯ ЛИСТА ЗАКАЗ (AvtoZakaz)
 * ---------------------------------------------------------------------------------------
 * Назначение: Автоматический пересчёт производных столбцов на листе Заказ
 * Работает независимо от функций сортировки для всех проектов
 *
 * Вычисляемые столбцы:
 * - Среднемесячные продажи, шт
 * - Потребность на N месяцев, шт
 * - На сколько мес запас
 * - Необходимо заказать, шт
 * - Предварительный заказ
 * - заказ в упаковках
 * - ЗАКАЗ
 * - EXW ALFASPA текущая, €
 * - Предварительная сумма заказа
 * =======================================================================================
 */

(function (Lib, global) {
  const HEADER_ROW_INDEX = 1;

  /**
   * ПУБЛИЧНАЯ: Пересчитать все производные столбцы на листе Заказ
   */
  Lib.recalculateOrderSheet = function () {
    try {
      _populateDerivedColumns();
    } catch (error) {
      Lib.logError('[AvtoZakaz] recalculateOrderSheet: ошибка', error);
    }
  };

  /**
   * ПУБЛИЧНАЯ: Обработка изменений на листе Заказ для автопересчёта
   * Вызывается из onUpdateOverride в 06OrderForm.js
   */
  Lib.handleAvtoZakazUpdate = function (sheet, row, col, debugSource) {
    try {
      Lib.logDebug('[AvtoZakaz] handleAvtoZakazUpdate: вызвана для row=' + row + ', col=' + col + ', source=' + debugSource);

      if (!sheet || !global.CONFIG) {
        Lib.logDebug('[AvtoZakaz] handleAvtoZakazUpdate: sheet/CONFIG отсутствуют');
        return;
      }

      var ORDER_SHEET =
        (global.CONFIG.SHEETS && global.CONFIG.SHEETS.ORDER_FORM) || 'Заказ';
      if (sheet.getName() !== ORDER_SHEET) {
        Lib.logDebug('[AvtoZakaz] handleAvtoZakazUpdate: лист не ORDER_SHEET (' + sheet.getName() + ' !== ' + ORDER_SHEET + ')');
        return;
      }

      Lib.logDebug('[AvtoZakaz] handleAvtoZakazUpdate: обработка начата для листа ' + ORDER_SHEET);

      var TOP_HEADER_ROW = 1;
      var DATA_HEADERS_ROW = 3;
      var FIRST_DATA_ROW = 2;

      function readHeaderAt(r, c) {
        var v = String(sheet.getRange(r, c).getValue() || '').trim();
        if (v) return v;
        try {
          var d = String(sheet.getRange(r, c).getDisplayValue() || '').trim();
          if (d) return d;
        } catch (_) {}
        try {
          var mrs = sheet.getMergedRanges();
          for (var i = 0; i < mrs.length; i++) {
            var mr = mrs[i];
            if (
              r >= mr.getRow() &&
              r <= mr.getLastRow() &&
              c >= mr.getColumn() &&
              c <= mr.getLastColumn()
            ) {
              var mv = mr.getValues()[0][0];
              return String(mv || '').trim();
            }
          }
        } catch (_) {}
        return '';
      }

      // 1) Правка строки 1 (верхний заголовок)
      if (row === TOP_HEADER_ROW) {
        var topHeader = readHeaderAt(TOP_HEADER_ROW, col);
        var m = topHeader.match(/потребность\s+на\s*(\d+)\s*месяц/i);
        if (topHeader && (topHeader.toLowerCase() === 'кол-во месяцев' || m)) {
          Lib.logDebug(
            '[AvtoZakaz] handleAvtoZakazUpdate: строка 1 изменена: ' + topHeader
          );

          if (m) {
            var n = parseInt(m[1], 10);
            Lib.logDebug(
              '[AvtoZakaz] handleAvtoZakazUpdate: найден коэффициент N=' +
                n +
                ', начинаем пересчет'
            );

            var lastRow = sheet.getLastRow();
            var lastCol = sheet.getLastColumn();
            var headersInfo = _collectOrderHeaders_(sheet, lastCol, DATA_HEADERS_ROW);
            var headerMap = _buildOrderHeaderIndexMap(headersInfo);
            var needCol = col;
            var avgCol = headerMap.avg !== -1 ? headerMap.avg + 1 : -1;
            var avgLabel = _resolveHeaderLabel_(headersInfo, headerMap.avg);
            var needOrderCol =
              headerMap.needOrder !== -1 ? headerMap.needOrder + 1 : -1;
            var stockCol = headerMap.stock !== -1 ? headerMap.stock + 1 : -1;
            var reserveCol =
              headerMap.reserve !== -1 ? headerMap.reserve + 1 : -1;
            var inTransitCol =
              headerMap.inTransit !== -1 ? headerMap.inTransit + 1 : -1;

            var rowsToProcess = lastRow >= FIRST_DATA_ROW ? lastRow - FIRST_DATA_ROW + 1 : 0;
            if (avgCol !== -1 && rowsToProcess > 0) {
              Lib.logDebug(
                '[AvtoZakaz] handleAvtoZakazUpdate: найден столбец "' +
                  (avgLabel || 'Среднемесячные продажи, шт') +
                  '" #' +
                  avgCol +
                  ', пересчитываем ' +
                  rowsToProcess +
                  ' строк'
              );
              var avgColumnValues = sheet
                .getRange(FIRST_DATA_ROW, avgCol, rowsToProcess, 1)
                .getValues();
              var stockColumnValues =
                stockCol !== -1 && needOrderCol !== -1
                  ? sheet
                      .getRange(FIRST_DATA_ROW, stockCol, rowsToProcess, 1)
                      .getValues()
                  : null;
              var transitColumnValues =
                inTransitCol !== -1 && needOrderCol !== -1
                  ? sheet
                      .getRange(FIRST_DATA_ROW, inTransitCol, rowsToProcess, 1)
                      .getValues()
                  : null;
              var reserveColumnValues =
                reserveCol !== -1 && needOrderCol !== -1
                  ? sheet
                      .getRange(FIRST_DATA_ROW, reserveCol, rowsToProcess, 1)
                      .getValues()
                  : null;

              var needValues = [];
              var needOrderValues = needOrderCol !== -1 ? [] : null;
              for (var r = 0; r < rowsToProcess; r++) {
                var avgVal = _parseNumber(avgColumnValues[r][0]) || 0;
                var need = avgVal * n;
                needValues.push([need]);

                if (needOrderValues) {
                  var stock = stockColumnValues
                    ? _parseNumber(stockColumnValues[r][0]) || 0
                    : 0;
                  var transit = transitColumnValues
                    ? _parseNumber(transitColumnValues[r][0]) || 0
                    : 0;
                  var reserve = reserveColumnValues
                    ? _parseNumber(reserveColumnValues[r][0]) || 0
                    : 0;
                  needOrderValues.push([need - (stock + transit - reserve)]);
                }
              }

              sheet
                .getRange(FIRST_DATA_ROW, needCol, rowsToProcess, 1)
                .setValues(needValues);
              if (needOrderValues) {
                sheet
                  .getRange(FIRST_DATA_ROW, needOrderCol, rowsToProcess, 1)
                  .setValues(needOrderValues);
              }
              Lib.logDebug('[AvtoZakaz] handleAvtoZakazUpdate: пересчет завершен');
            } else if (rowsToProcess === 0) {
              Lib.logDebug(
                '[AvtoZakaz] handleAvtoZakazUpdate: данных для пересчёта нет'
              );
            } else {
              Lib.logWarn(
                '[AvtoZakaz] handleAvtoZakazUpdate: не найден столбец "Среднемесячные продажи"'
              );
            }
          } else {
            _populateDerivedColumns(col);
          }
        }
        return;
      }

      // 2) Правка строки 2 (коэффициент под «Кол-во месяцев»)
      if (row === 2) {
        var topHeader = readHeaderAt(TOP_HEADER_ROW, col).toLowerCase();
        if (
          topHeader === 'кол-во месяцев' ||
          /\bпотребность\s+на\b/i.test(topHeader)
        ) {
          Lib.logDebug(
            '[AvtoZakaz] handleAvtoZakazUpdate: строка 2 изменена под "Кол-во месяцев" — пересчёт'
          );
          _populateDerivedColumns(col);
        }
        return;
      }

      // 3) Правка строки 3 (заголовки данных): «Потребность на N месяцев, шт»
      if (row === DATA_HEADERS_ROW) {
        var dataHeader = readHeaderAt(DATA_HEADERS_ROW, col);
        var m = dataHeader.match(/потребность\s+на\s*(\d+)\s*месяц/i);
        if (m) {
          Lib.logDebug(
            '[AvtoZakaz] handleAvtoZakazUpdate: строка 3 изменена: ' + dataHeader
          );

          var topHeader = 'Потребность на ' + m[1] + ' месяцев, шт';
          sheet.getRange(TOP_HEADER_ROW, col).setValue(topHeader);

          var row3CellForData = sheet.getRange(DATA_HEADERS_ROW, col);
          var row3CurrentValue = String(row3CellForData.getValue() || '').trim();
          var canResetRow3 =
            !row3CurrentValue ||
            /потребность\s+на\s*\d+\s*месяц/i.test(row3CurrentValue);

          if (canResetRow3) {
            row3CellForData.clearContent();
          }

          var n = parseInt(m[1], 10);
          Lib.logDebug(
            '[AvtoZakaz] handleAvtoZakazUpdate: найден коэффициент N=' +
              n +
              ', начинаем пересчет'
          );

          var lastRow = sheet.getLastRow();
          var lastCol = sheet.getLastColumn();
          var headersInfo = _collectOrderHeaders_(sheet, lastCol, DATA_HEADERS_ROW);
          var headerMap = _buildOrderHeaderIndexMap(headersInfo);

          var needCol = col;
          var avgCol = headerMap.avg !== -1 ? headerMap.avg + 1 : -1;
          var avgLabel = _resolveHeaderLabel_(headersInfo, headerMap.avg);
          var needOrderCol =
            headerMap.needOrder !== -1 ? headerMap.needOrder + 1 : -1;
          var stockCol = headerMap.stock !== -1 ? headerMap.stock + 1 : -1;
          var reserveCol =
            headerMap.reserve !== -1 ? headerMap.reserve + 1 : -1;
          var inTransitCol =
            headerMap.inTransit !== -1 ? headerMap.inTransit + 1 : -1;

          var rowsToProcess = lastRow >= FIRST_DATA_ROW ? lastRow - FIRST_DATA_ROW + 1 : 0;
          if (avgCol !== -1 && rowsToProcess > 0) {
            Lib.logDebug(
              '[AvtoZakaz] handleAvtoZakazUpdate: найден столбец "' +
                (avgLabel || 'Среднемесячные продажи, шт') +
                '" #' +
                avgCol +
                ', пересчитываем ' +
                rowsToProcess +
                ' строк'
            );
            var avgColumnValues = sheet
              .getRange(FIRST_DATA_ROW, avgCol, rowsToProcess, 1)
              .getValues();
            var stockColumnValues =
              stockCol !== -1 && needOrderCol !== -1
                ? sheet
                    .getRange(FIRST_DATA_ROW, stockCol, rowsToProcess, 1)
                    .getValues()
                : null;
            var transitColumnValues =
              inTransitCol !== -1 && needOrderCol !== -1
                ? sheet
                    .getRange(FIRST_DATA_ROW, inTransitCol, rowsToProcess, 1)
                    .getValues()
                : null;
            var reserveColumnValues =
              reserveCol !== -1 && needOrderCol !== -1
                ? sheet
                    .getRange(FIRST_DATA_ROW, reserveCol, rowsToProcess, 1)
                    .getValues()
                : null;

            var needValues = [];
            var needOrderValues = needOrderCol !== -1 ? [] : null;
            for (var r = 0; r < rowsToProcess; r++) {
              var avgVal = _parseNumber(avgColumnValues[r][0]) || 0;
              var need = avgVal * n;
              needValues.push([need]);

              if (needOrderValues) {
                var stock = stockColumnValues
                  ? _parseNumber(stockColumnValues[r][0]) || 0
                  : 0;
                var transit = transitColumnValues
                  ? _parseNumber(transitColumnValues[r][0]) || 0
                  : 0;
                var reserve = reserveColumnValues
                  ? _parseNumber(reserveColumnValues[r][0]) || 0
                  : 0;
                needOrderValues.push([need - (stock + transit - reserve)]);
              }
            }

            sheet
              .getRange(FIRST_DATA_ROW, needCol, rowsToProcess, 1)
              .setValues(needValues);
            if (needOrderValues) {
              sheet
                .getRange(FIRST_DATA_ROW, needOrderCol, rowsToProcess, 1)
                .setValues(needOrderValues);
            }
            Lib.logDebug('[AvtoZakaz] handleAvtoZakazUpdate: пересчет завершен');
          } else if (rowsToProcess === 0) {
            Lib.logDebug(
              '[AvtoZakaz] handleAvtoZakazUpdate: данных для пересчёта нет'
            );
          } else {
            Lib.logWarn(
              '[AvtoZakaz] handleAvtoZakazUpdate: не найден столбец "Среднемесячные продажи"'
            );
          }
        }
        return;
      }

      // 4) Ниже строки 3 — пересчёт по конкретному полю
      if (row > DATA_HEADERS_ROW) {
        Lib.logDebug('[AvtoZakaz] handleAvtoZakazUpdate: обработка строки ' + row + ', колонка ' + col);
        var h = readHeaderAt(DATA_HEADERS_ROW, col);
        if (!h) h = readHeaderAt(TOP_HEADER_ROW, col);
        var n = _normalizeHeaderName_(h);
        var cellValue = sheet.getRange(row, col).getValue();
        Lib.logDebug('[AvtoZakaz] handleAvtoZakazUpdate: заголовок="' + h + '", нормализованный="' + n + '", значение="' + cellValue + '"');

        // Проверяем, является ли это колонкой "Добавить в прайс"
        if (n === 'добавить' || n === 'убрать') {
          var topHeaderNormalized = _normalizeHeaderName_(readHeaderAt(TOP_HEADER_ROW, col));
          Lib.logDebug('[AvtoZakaz] handleAvtoZakazUpdate: верхний заголовок нормализованный="' + topHeaderNormalized + '"');
          if (topHeaderNormalized === 'добавить в прайс') {
            n = 'добавить в прайс';
            Lib.logDebug('[AvtoZakaz] handleAvtoZakazUpdate: определена колонка "добавить в прайс"');
          }
        }

        if (n === 'добавить в прайс') {
          var normalizedFlag =
            cellValue === true
              ? 'добавить'
              : String(cellValue || '')
                  .trim()
                  .toLowerCase();
          if (normalizedFlag === 'добавить') {
            Lib.logDebug(
              '[AvtoZakaz] handleAvtoZakazUpdate: обнаружен флаг добавления в прайс для строки ' +
                row +
                ', запускаю перенос'
            );
            if (Lib.addOrderItemsToPrice && typeof Lib.addOrderItemsToPrice === 'function') {
              try {
                Lib.addOrderItemsToPrice({
                  silent: true,
                  triggerSource: 'onEdit:addToPrice',
                  orderRow: row,
                });
              } catch (priceErr) {
                Lib.logError('[AvtoZakaz] handleAvtoZakazUpdate: addOrderItemsToPrice error', priceErr);
              }
            } else {
              Lib.logWarn('[AvtoZakaz] handleAvtoZakazUpdate: addOrderItemsToPrice недоступна');
            }
          } else if (normalizedFlag === 'убрать') {
            Lib.logDebug(
              '[AvtoZakaz] handleAvtoZakazUpdate: обнаружен флаг удаления из прайса для строки ' +
                row +
                ', запускаю удаление'
            );
            if (Lib.removeOrderItemsFromPrice && typeof Lib.removeOrderItemsFromPrice === 'function') {
              try {
                Lib.removeOrderItemsFromPrice({
                  silent: true,
                  triggerSource: 'onEdit:removeFromPrice',
                  orderRow: row,
                });
              } catch (priceErr) {
                Lib.logError('[AvtoZakaz] handleAvtoZakazUpdate: removeOrderItemsFromPrice error', priceErr);
              }
            } else {
              Lib.logWarn('[AvtoZakaz] handleAvtoZakazUpdate: removeOrderItemsFromPrice недоступна');
            }
          }
          return;
        }

        var trigger = null;
        var compactName = n.replace(/\s+/g, '');

        var shouldRefreshDerived =
          n === 'остаток' ||
          n === 'резерв' ||
          compactName.indexOf('товарвпути') === 0 ||
          compactName.indexOf('продажи') === 0 ||
          compactName.indexOf('среднемесячныепродажи') === 0 ||
          compactName.indexOf('насколькомесзапас') === 0;

        // НОВАЯ ЛОГИКА: автоматическое заполнение столбцов "Срок#" и "Срок"
        // Проверяем значение в ячейке - если оно непустое, заполняем срок
        var columnTypeForExpiry = null;

        if (n === 'акции') {
          columnTypeForExpiry = 'акции';
        } else if (n === 'набор') {
          columnTypeForExpiry = 'набор';
        }

        if (columnTypeForExpiry) {
          var hasValue = cellValue !== null && cellValue !== undefined && String(cellValue).trim() !== '';
          if (hasValue) {
            Lib.logDebug(
              '[AvtoZakaz] handleAvtoZakazUpdate: обнаружено значение "' + cellValue + '" в столбце "' + h + '", запуск автозаполнения срока'
            );
            _autoFillExpiryDates(sheet, row, columnTypeForExpiry);
          } else {
            // Если значение удалено, очищаем поле срока
            Lib.logDebug(
              '[AvtoZakaz] handleAvtoZakazUpdate: значение удалено из столбца "' + h + '", очищаем срок'
            );
            _clearExpiryDate(sheet, row, columnTypeForExpiry);
          }
        }

        if (n.indexOf('предварительный заказ') === 0) trigger = 'preOrder';
        else if (n.indexOf('заказ в упаковках') === 0) trigger = 'packs';
        else if (n === 'заказ') trigger = 'order';
        else if (n.replace(/\s+/g, '') === 'штуп' || n.indexOf('шт уп') !== -1)
          trigger = 'unitsPerPack';
        else if (n.indexOf('exw alfaspa текущая') === 0) trigger = 'price';
        else if (n.indexOf('предварительная сумма заказа') === 0)
          trigger = 'sum';

        if (!trigger && debugSource && typeof debugSource === 'string') {
          try {
            var ds = String(debugSource).toLowerCase();
            if (ds.indexOf(':preorder') !== -1) trigger = 'preOrder';
            else if (ds.indexOf(':packs') !== -1) trigger = 'packs';
            else if (ds.indexOf(':order') !== -1) trigger = 'order';
            else if (ds.indexOf(':unitsperpack') !== -1) trigger = 'unitsPerPack';
            else if (ds.indexOf(':price') !== -1) trigger = 'price';
            else if (ds.indexOf(':sum') !== -1) trigger = 'sum';
            else if (ds.indexOf(':derived') !== -1) {
              shouldRefreshDerived = true;
            }
          } catch (_) {}
        }

        if (!trigger && shouldRefreshDerived) {
          Lib.logDebug(
            '[AvtoZakaz] handleAvtoZakazUpdate: пересчёт производных столбцов для R' +
              row +
              'C' +
              col +
              ' (' +
              h +
              ')'
          );
          _populateDerivedColumns();
          return;
        }

        if (trigger) {
          Lib.logDebug(
            '[AvtoZakaz] handleAvtoZakazUpdate: trigger=' +
              trigger +
              ' for R' +
              row +
              'C' +
              col
          );
          _recalculateOrderRow_(sheet, row, { trigger: trigger });
        }
      }
    } catch (err) {
      Lib.logError('[AvtoZakaz] handleAvtoZakazUpdate: ошибка', err);
    }
  };

  // =======================================================================================
  // ВНУТРЕННИЕ ФУНКЦИИ
  // =======================================================================================

  function _populateDerivedColumns(monthsColumnOverride) {
    try {
      Lib.logDebug(
        '[AvtoZakaz] _populateDerivedColumns: старт (override=' +
          (typeof monthsColumnOverride === 'number'
            ? monthsColumnOverride
            : 'нет') +
          ')'
      );
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) {
        return;
      }

      var orderSheetName =
        (global.CONFIG.SHEETS && global.CONFIG.SHEETS.ORDER_FORM) || 'Заказ';
      var orderSheet = ss.getSheetByName(orderSheetName);
      if (!orderSheet) {
        Lib.logWarn('[AvtoZakaz] _populateDerivedColumns: лист заказа не найден');
        return;
      }

      var headerRowIndex = HEADER_ROW_INDEX;
      var firstDataRow = _resolveFirstDataRow_(orderSheet, headerRowIndex);
      var lastRow = orderSheet.getLastRow();
      if (lastRow < firstDataRow) {
        Lib.logDebug(
          '[AvtoZakaz] _populateDerivedColumns: нет строк для обновления'
        );
        return;
      }

      var lastColumn = orderSheet.getLastColumn();
      var headersInfo = _collectOrderHeaders_(orderSheet, lastColumn, headerRowIndex);
      var headersEffective = headersInfo.effective;

      var idx = _buildOrderHeaderIndexMap(headersInfo);

      Lib.logDebug(
        '[AvtoZakaz] _populateDerivedColumns: header indices: ' +
          JSON.stringify(idx)
      );

      if (idx.need === -1) {
        Lib.logWarn(
          '[AvtoZakaz] _populateDerivedColumns: столбец "Потребность на 6 месяцев, шт" не найден'
        );
        return;
      }

      if (idx.avg === -1) {
        Lib.logWarn(
          '[AvtoZakaz] _populateDerivedColumns: столбец "Среднемесячные продажи, шт" не найден'
        );
        return;
      }

      var monthsColumn = null;
      if (typeof monthsColumnOverride === 'number') {
        monthsColumn = monthsColumnOverride;
      } else {
        var topRow = headersInfo.top.map(function (value) {
          return String(value || '')
            .trim()
            .toLowerCase();
        });
        var found = topRow.indexOf('кол-во месяцев');
        if (found !== -1) {
          monthsColumn = found + 1;
        }
      }

      var monthsFactor = null;
      if (monthsColumn !== null) {
        var headerNeedValue =
          firstDataRow > 2
            ? orderSheet.getRange(2, monthsColumn).getValue()
            : null;
        monthsFactor = _parseNumber(headerNeedValue);
        if (monthsFactor === null || monthsFactor <= 0) {
          var headerFromText = headersInfo.effective[monthsColumn - 1] || '';
          var matchFromHeader = headerFromText.match(/потребность\s+на\s*(\d+)/i);
          if (!matchFromHeader) {
            var topHeaderText = headersInfo.top[monthsColumn - 1] || '';
            matchFromHeader = topHeaderText.match(/потребность\s+на\s*(\d+)/i);
          }
          if (matchFromHeader) {
            monthsFactor = parseInt(matchFromHeader[1], 10);
          }
        }
      }

      if ((monthsFactor === null || monthsFactor <= 0) && idx.need !== -1) {
        var needHeader = _resolveHeaderLabel_(headersInfo, idx.need) || '';
        var matchNeed = needHeader.match(/потребность\s+на\s*(\d+)/i);
        if (matchNeed) {
          monthsFactor = parseInt(matchNeed[1], 10);
        }
      }

      if (monthsFactor === null || monthsFactor <= 0) {
        var fallback =
          (Lib.CONFIG &&
            Lib.CONFIG.SETTINGS &&
            Number(Lib.CONFIG.SETTINGS.DEFAULT_MONTHS)) || 6;
        monthsFactor = fallback;
      }

      Lib.logDebug(
        '[AvtoZakaz] _populateDerivedColumns: коэффициент = ' + monthsFactor
      );

      var numRows = lastRow - firstDataRow + 1;

      var dataRange = orderSheet.getRange(firstDataRow, 1, numRows, lastColumn);
      var values = dataRange.getValues();

      var avgOutputValues = idx.avg !== -1 ? [] : null;
      var needValues = [];
      var needAccumulator = idx.need !== -1 ? 0 : null;
      var needToOrderValues = idx.needOrder !== -1 ? [] : null;
      var needToOrderAccumulator = idx.needOrder !== -1 ? 0 : null;
      var monthsValues = idx.months !== -1 ? [] : null;
      var packValues = idx.packs !== -1 ? [] : null;
      var orderValues = idx.order !== -1 ? [] : null;
      var sumValues = idx.sum !== -1 ? [] : null;
      var preOrderAccumulator = idx.preOrder !== -1 ? 0 : null;
      var packAccumulator = packValues ? 0 : null;
      var orderAccumulator = orderValues ? 0 : null;
      var sumAccumulator = sumValues ? 0 : null;

      for (var i = 0; i < values.length; i++) {
        var row = values[i];
        var totalSalesValue =
          idx.totalSales !== -1 ? _parseNumber(row[idx.totalSales]) : null;
        var stockValue = idx.stock !== -1 ? _parseNumber(row[idx.stock]) : null;
        var reserveValue = idx.reserve !== -1 ? _parseNumber(row[idx.reserve]) : null;
        var inTransitValue =
          idx.inTransit !== -1 ? _parseNumber(row[idx.inTransit]) : null;
        var avgValue = null;
        if (idx.avg !== -1) {
          if (totalSalesValue !== null && !isNaN(totalSalesValue)) {
            avgValue = Math.round(totalSalesValue / 12);
          } else {
            avgValue = _parseNumber(row[idx.avg]);
          }
          row[idx.avg] = avgValue;
          if (avgOutputValues) {
            avgOutputValues.push([_formatNumberOrBlank_(avgValue)]);
          }
        } else {
          avgValue = totalSalesValue !== null && !isNaN(totalSalesValue)
            ? Math.round(totalSalesValue / 12)
            : _parseNumber(row[idx.avg]);
        }

        var need = avgValue && !isNaN(avgValue) ? avgValue * monthsFactor : 0;
        if (!need || isNaN(need)) {
          need = 0;
        }
        needValues.push([need]);
        if (needAccumulator !== null) {
          needAccumulator += need;
        }

        if (stockValue === null || isNaN(stockValue)) {
          stockValue = 0;
        }
        if (reserveValue === null || isNaN(reserveValue)) {
          reserveValue = 0;
        }
        if (inTransitValue === null || isNaN(inTransitValue)) {
          inTransitValue = 0;
        }

        if (needToOrderValues) {
          var needToOrder = need - (stockValue + inTransitValue - reserveValue);
          needToOrderValues.push([needToOrder]);
          if (needToOrderAccumulator !== null) {
            needToOrderAccumulator += needToOrder;
          }
        }

        if (monthsValues) {
          var monthsCover = 0;
          if (avgValue !== null && avgValue !== undefined && !isNaN(avgValue) && avgValue > 0) {
            monthsCover = stockValue / avgValue;
          }
          monthsValues.push([monthsCover ? Number(monthsCover.toFixed(1)) : 0]);
        }

        var preOrderValue =
          idx.preOrder !== -1 ? _parseNumber(row[idx.preOrder]) : null;
        if (
          preOrderAccumulator !== null &&
          preOrderValue !== null &&
          !isNaN(preOrderValue)
        ) {
          preOrderAccumulator += preOrderValue;
        }

        if (packValues || orderValues || sumValues) {
          var unitsPerPackValue =
            idx.unitsPerPack !== -1
              ? _parseNumber(row[idx.unitsPerPack])
              : null;
          var existingPackValue =
            idx.packs !== -1 ? _parseNumber(row[idx.packs]) : null;

          var computedPack = existingPackValue;
          if (packValues) {
            var forcePackFromPreOrder = preOrderValue !== null;
            computedPack = _calculatePackValue_(
              preOrderValue,
              unitsPerPackValue,
              existingPackValue,
              forcePackFromPreOrder
            );
            packValues.push([_formatNumberOrBlank_(computedPack)]);
            if (
              packAccumulator !== null &&
              computedPack !== null &&
              !isNaN(computedPack)
            ) {
              packAccumulator += computedPack;
            }
          }

          var orderUnits = null;
          if (orderValues) {
            orderUnits = _calculateOrderUnits_(computedPack, unitsPerPackValue);
            orderValues.push([_formatNumberOrBlank_(orderUnits)]);
            if (
              orderAccumulator !== null &&
              orderUnits !== null &&
              orderUnits !== undefined &&
              !isNaN(orderUnits)
            ) {
              orderAccumulator += orderUnits;
            }
          } else if (idx.order !== -1) {
            orderUnits = _parseNumber(row[idx.order]);
          }

          if (sumValues) {
            var priceValue =
              idx.price !== -1 ? _parseNumber(row[idx.price]) : null;
            var preliminarySum = _calculatePreliminarySum_(
              orderUnits,
              priceValue
            );
            if (preliminarySum !== null && !isNaN(preliminarySum)) {
              sumAccumulator += preliminarySum;
              sumValues.push([preliminarySum]);
            } else {
              sumValues.push(['']);
            }
          }
        }
      }

      if (avgOutputValues) {
        orderSheet
          .getRange(firstDataRow, idx.avg + 1, numRows, 1)
          .setValues(avgOutputValues)
          .setFontFamily('Century Gothic')
          .setFontSize(9)
          .setNumberFormat('0')
          .setBorder(false, false, false, false, false, false);
      }

      orderSheet
        .getRange(firstDataRow, idx.need + 1, numRows, 1)
        .setValues(needValues)
        .setFontFamily('Century Gothic')
        .setFontSize(9)
        .setNumberFormat('0')
        .setBorder(false, false, false, false, false, false);

      var updatedColumns = [];
      if (avgOutputValues) {
        updatedColumns.push('«Среднемесячные продажи…»');
      }
      updatedColumns.push('«Потребность…»');

      if (needToOrderValues) {
        orderSheet
          .getRange(firstDataRow, idx.needOrder + 1, numRows, 1)
          .setValues(needToOrderValues)
          .setFontFamily('Century Gothic')
          .setFontSize(9)
          .setNumberFormat('0')
          .setBorder(false, false, false, false, false, false);
        updatedColumns.push('«Необходимо заказать…»');
      }

      if (monthsValues) {
        orderSheet
          .getRange(firstDataRow, idx.months + 1, numRows, 1)
          .setValues(monthsValues)
          .setFontFamily('Century Gothic')
          .setFontSize(9)
          .setNumberFormat('0.0 "мес."')
          .setBorder(false, false, false, false, false, false);
        updatedColumns.push('«На сколько мес запас»');
      }

      if (packValues) {
        orderSheet
          .getRange(firstDataRow, idx.packs + 1, numRows, 1)
          .setValues(packValues)
          .setFontFamily('Century Gothic')
          .setFontSize(9)
          .setNumberFormat('0')
          .setBorder(false, false, false, false, false, false);
        updatedColumns.push('«заказ в упаковках»');
      }

      if (orderValues) {
        orderSheet
          .getRange(firstDataRow, idx.order + 1, numRows, 1)
          .setValues(orderValues)
          .setFontFamily('Century Gothic')
          .setFontSize(9)
          .setNumberFormat('0')
          .setBorder(false, false, false, false, false, false);
        updatedColumns.push('«ЗАКАЗ»');
      }

      if (sumValues) {
        orderSheet
          .getRange(firstDataRow, idx.sum + 1, numRows, 1)
          .setValues(sumValues)
          .setFontFamily('Century Gothic')
          .setFontSize(9)
          .setNumberFormat('#,##0.00"€"')
          .setBorder(false, false, false, false, false, false);

        var newTotalForHeader = _recalculateColumnTotal_(
          orderSheet,
          idx.sum + 1,
          Math.max(1, firstDataRow - 1),
          '#,##0.00"€"'
        );
        var sumHeaderLabel =
          'Предварительная сумма заказа ' + _formatEuroLabel_(newTotalForHeader);
        orderSheet.getRange(HEADER_ROW_INDEX, idx.sum + 1).setValue(sumHeaderLabel);
        updatedColumns.push('«предварительная сумма заказа»');
      }

      Lib.logInfo(
        '[AvtoZakaz] _populateDerivedColumns: обновлены столбцы ' +
          updatedColumns.join(', ') +
          ' для ' +
          numRows +
          ' строк'
      );
    } catch (err) {
      Lib.logError('[AvtoZakaz] _populateDerivedColumns: ошибка', err);
    }
  }

  function _recalculateOrderRow_(sheet, row, options) {
    if (!sheet || !row) return;
    options = options || {};
    var detectedFirstDataRow = _resolveFirstDataRow_(sheet, HEADER_ROW_INDEX);
    var headerRowIndex = Math.max(1, detectedFirstDataRow - 1);
    if (row <= headerRowIndex) return;
    var lastColumn = sheet.getLastColumn();
    var headersInfo = _collectOrderHeaders_(sheet, lastColumn, headerRowIndex);
    var idx = _buildOrderHeaderIndexMap(headersInfo);
    var rowValues = sheet.getRange(row, 1, 1, lastColumn).getValues()[0];
    var preOrderValue =
      idx.preOrder !== -1 ? _parseNumber(rowValues[idx.preOrder]) : null;
    var unitsPerPackValue =
      idx.unitsPerPack !== -1
        ? _parseNumber(rowValues[idx.unitsPerPack])
        : null;
    var packCellValue =
      idx.packs !== -1 ? _parseNumber(rowValues[idx.packs]) : null;
    var orderCellValue =
      idx.order !== -1 ? _parseNumber(rowValues[idx.order]) : null;
    var priceValue =
      idx.price !== -1 ? _parseNumber(rowValues[idx.price]) : null;
    var trigger = options.trigger || '';
    var forceFromPreOrder =
      trigger === 'preOrder' ||
      (trigger === 'unitsPerPack' && preOrderValue !== null);
    Lib.logDebug(
      '[AvtoZakaz] recalc row ' +
        row +
        ': trigger=' +
        trigger +
        ', preOrder=' +
        (preOrderValue !== null ? preOrderValue : '∅') +
        ', unitsPerPack=' +
        (unitsPerPackValue !== null ? unitsPerPackValue : '∅') +
        ', packs(cell)=' +
        (packCellValue !== null ? packCellValue : '∅') +
        ', order(cell)=' +
        (orderCellValue !== null ? orderCellValue : '∅') +
        ', price=' +
        (priceValue !== null ? priceValue : '∅')
    );
    var computedPack = packCellValue;
    if (idx.packs !== -1) {
      computedPack = _calculatePackValue_(
        preOrderValue,
        unitsPerPackValue,
        packCellValue,
        forceFromPreOrder
      );
      var packOutput = _formatNumberOrBlank_(computedPack);
      Lib.logDebug(
        '[AvtoZakaz] recalc row ' +
          row +
          ': computedPack=' +
          computedPack +
          ' (forceFromPreOrder=' +
          forceFromPreOrder +
          ')'
      );
      sheet
        .getRange(row, idx.packs + 1)
        .setValue(packOutput === '' ? '' : packOutput)
        .setFontFamily('Century Gothic')
        .setFontSize(9)
        .setNumberFormat('0')
        .setBorder(false, false, false, false, false, false);
      Lib.logDebug(
        '[AvtoZakaz] recalc row ' +
          row +
          ': pack updated -> ' +
          (packOutput === '' ? '∅' : packOutput)
      );
    }
    var effectiveOrder = orderCellValue;
    if (idx.order !== -1) {
      if (trigger === 'order') {
        effectiveOrder = orderCellValue;
      } else {
        var calculatedOrder = _calculateOrderUnits_(
          computedPack,
          unitsPerPackValue
        );
        Lib.logDebug(
          '[AvtoZakaz] recalc row ' +
            row +
            ': calculatedOrder=' +
            calculatedOrder +
            ' from computedPack=' +
            computedPack +
            ', unitsPerPack=' +
            unitsPerPackValue
        );
        if (calculatedOrder !== null && !isNaN(calculatedOrder))
          effectiveOrder = calculatedOrder;
      }
      var orderOutput = _formatNumberOrBlank_(effectiveOrder);
      sheet
        .getRange(row, idx.order + 1)
        .setValue(orderOutput === '' ? '' : orderOutput)
        .setFontFamily('Century Gothic')
        .setFontSize(9)
        .setNumberFormat('0')
        .setBorder(false, false, false, false, false, false);
      Lib.logDebug(
        '[AvtoZakaz] recalc row ' +
          row +
          ': order updated -> ' +
          (orderOutput === '' ? '∅' : orderOutput)
      );
    }
    if (idx.sum !== -1) {
      var sumValue = _calculatePreliminarySum_(effectiveOrder, priceValue);
      Lib.logDebug(
        '[AvtoZakaz] recalc row ' +
          row +
          ': sumValue=' +
          sumValue +
          ' (order=' +
          effectiveOrder +
          ', price=' +
          priceValue +
          ')'
      );
      var sumOutput = _formatNumberOrBlank_(sumValue);
      sheet
        .getRange(row, idx.sum + 1)
        .setValue(sumOutput === '' ? '' : sumValue)
        .setFontFamily('Century Gothic')
        .setFontSize(9)
        .setNumberFormat('#,##0.00"€"')
        .setBorder(false, false, false, false, false, false);
      Lib.logDebug(
        '[AvtoZakaz] recalc row ' +
          row +
          ': sum updated -> ' +
          (sumOutput === '' ? '∅' : sumValue)
      );
    }

    if (idx.sum !== -1) {
      var newTotal = _recalculateColumnTotal_(
        sheet,
        idx.sum + 1,
        headerRowIndex,
        '#,##0.00"€"'
      );
      try {
        var sumHeaderLabel =
          'Предварительная сумма заказа ' + _formatEuroLabel_(newTotal);
        sheet.getRange(HEADER_ROW_INDEX, idx.sum + 1).setValue(sumHeaderLabel);
      } catch (_) {}
    }
  }

  function _collectOrderHeaders_(sheet, lastColumn, dataHeaderRow) {
    var rowsToFetch = Math.max(dataHeaderRow, 3);
    try {
      var maxRows = typeof sheet.getMaxRows === 'function' ? sheet.getMaxRows() : rowsToFetch;
      if (maxRows < rowsToFetch) {
        rowsToFetch = maxRows;
      }
    } catch (e) {}
    var headerValues = sheet.getRange(1, 1, rowsToFetch, lastColumn).getValues();
    var topRow = headerValues[0]
      ? headerValues[0].map(_trimHeaderValue_)
      : [];
    var bannerRow = headerValues[1]
      ? headerValues[1].map(_trimHeaderValue_)
      : [];
    var dataRow = headerValues[dataHeaderRow - 1]
      ? headerValues[dataHeaderRow - 1].map(_trimHeaderValue_)
      : [];
    var effective = [];
    for (var i = 0; i < lastColumn; i++) {
      var value = topRow[i] || '';
      if (!value) {
        var candidate = dataRow[i];
        if (_isHeaderCandidate_(candidate)) {
          value = candidate;
        }
      }
      if (!value) {
        value = bannerRow[i] || '';
      }
      effective.push(_trimHeaderValue_(value));
    }
    return {
      effective: effective,
      top: topRow,
      banner: bannerRow,
      data: dataRow,
    };
  }

  function _trimHeaderValue_(value) {
    return String(value || '').trim();
  }

  function _resolveFirstDataRow_(sheet, headerRowIndex) {
    if (!sheet || headerRowIndex < 1) return 2;
    try {
      var lastColumn = sheet.getLastColumn();
      if (lastColumn <= 0) return 2;
      var rowValues = sheet
        .getRange(headerRowIndex, 1, 1, lastColumn)
        .getValues()[0];
      return _rowContainsHeaderCandidate_(rowValues)
        ? headerRowIndex + 1
        : 2;
    } catch (e) {
      return 2;
    }
  }

  function _isHeaderCandidate_(value) {
    if (!value) return false;
    var normalized = _normalizeHeaderName_(value);
    if (!normalized) return false;
    if (/^-?\d+[\d\s,.]*$/.test(normalized)) return false;
    if (normalized.indexOf('потребность на') === 0) return true;
    if (normalized.indexOf('среднемесячные продажи') === 0) return true;
    if (normalized.indexOf('необходимо заказать') === 0) return true;
    if (normalized === 'остаток') return true;
    if (normalized === 'резерв') return true;
    if (normalized.indexOf('товар в пути') === 0) return true;
    if (normalized.indexOf('предварительный заказ') === 0) return true;
    if (normalized === 'заказ') return true;
    if (normalized.indexOf('заказ в упаковках') === 0) return true;
    if (
      normalized.indexOf('шт уп') !== -1 ||
      normalized.indexOf('штуп') !== -1 ||
      normalized.replace(/\s+/g, '') === 'штуп'
    )
      return true;
    if (normalized.indexOf('exw alfaspa текущая') === 0) return true;
    if (normalized.indexOf('предварительная сумма заказа') === 0) return true;
    if (normalized.indexOf('кол-во месяцев') === 0) return true;
    return false;
  }

  function _rowContainsHeaderCandidate_(row) {
    if (!row || !row.length) return false;
    for (var i = 0; i < row.length; i++) {
      if (_isHeaderCandidate_(row[i])) return true;
    }
    return false;
  }

  function _getHeaderCandidates_(headersInfo, index) {
    var candidates = [];
    var seen = Object.create(null);
    if (!headersInfo) {
      return candidates;
    }
    var sources = [
      headersInfo.top || [],
      headersInfo.data || [],
      headersInfo.banner || [],
      headersInfo.effective || [],
    ];
    for (var s = 0; s < sources.length; s++) {
      var source = sources[s];
      if (!source || index >= source.length) continue;
      var candidate = _trimHeaderValue_(source[index]);
      if (!candidate) continue;
      if (!seen[candidate]) {
        candidates.push(candidate);
        seen[candidate] = true;
      }
    }
    return candidates;
  }

  function _resolveHeaderLabel_(headersInfo, index) {
    if (!headersInfo || index === undefined || index === null || index < 0) {
      return '';
    }
    var candidates = _getHeaderCandidates_(headersInfo, index);
    return candidates.length ? candidates[0] : '';
  }

  function _buildOrderHeaderIndexMap(headersInfo) {
    var map = {
      need: -1,
      avg: -1,
      totalSales: -1,
      stock: -1,
      reserve: -1,
      inTransit: -1,
      needOrder: -1,
      preOrder: -1,
      packs: -1,
      order: -1,
      unitsPerPack: -1,
      price: -1,
      sum: -1,
      months: -1,
    };
    var total = headersInfo && headersInfo.effective ? headersInfo.effective.length : 0;

    function compact(value) {
      return String(value || '')
        .toLowerCase()
        .replace(/[^a-zа-я0-9]/g, '');
    }

    for (var i = 0; i < total; i++) {
      var candidates = _getHeaderCandidates_(headersInfo, i);
      for (var j = 0; j < candidates.length; j++) {
        var raw = candidates[j];
        var normalized = _normalizeHeaderName_(raw);
        var cmp = compact(raw);

        if (map.need === -1 && normalized.indexOf('потребность на') === 0 && normalized.indexOf('месяц') !== -1)
          map.need = i;
        if (map.avg === -1 && normalized.indexOf('среднемесячные продажи') === 0)
          map.avg = i;
        if (
          map.totalSales === -1 &&
          (normalized.indexOf('продажи всего') === 0 ||
            normalized.indexOf('общие продажи') === 0 ||
            normalized === 'продажи' ||
            cmp === 'продажи' ||
            cmp.indexOf('prodazhivsego') === 0 ||
            cmp.indexOf('obshieprodazhi') === 0)
        )
          map.totalSales = i;

        if (map.stock === -1 && (normalized === 'остаток' || cmp === 'ostatok'))
          map.stock = i;
        if (map.reserve === -1 && (normalized === 'резерв' || cmp === 'rezerv'))
          map.reserve = i;

        if (
          map.inTransit === -1 &&
          (normalized.indexOf('товар в пути') === 0 ||
            normalized.indexOf('в пути') === 0 ||
            normalized.indexOf('в дороге') === 0 ||
            cmp.indexOf('vputi') === 0)
        )
          map.inTransit = i;

        if (map.needOrder === -1 && normalized.indexOf('необходимо заказать') === 0)
          map.needOrder = i;
        if (map.preOrder === -1 && normalized.indexOf('предварительный заказ') === 0)
          map.preOrder = i;
        if (map.packs === -1 && normalized.indexOf('заказ в упаковках') === 0)
          map.packs = i;
        if (map.order === -1 && normalized === 'заказ')
          map.order = i;

        if (
          map.unitsPerPack === -1 &&
          (normalized.indexOf('шт уп') !== -1 || cmp.indexOf('штуп') !== -1)
        )
          map.unitsPerPack = i;

        if (
          map.price === -1 &&
          (normalized.indexOf('exw alfaspa текущая') === 0 || cmp.indexOf('exwalfaspa') === 0)
        )
          map.price = i;

        if (map.sum === -1 && normalized.indexOf('предварительная сумма заказа') === 0)
          map.sum = i;

        if (
          map.months === -1 &&
          (normalized.indexOf('на сколько мес запас') === 0 || cmp.indexOf('насколькомесзапас') === 0)
        )
          map.months = i;

        if (
          map.need !== -1 &&
          map.avg !== -1 &&
          map.needOrder !== -1 &&
          map.packs !== -1 &&
          map.order !== -1 &&
          map.sum !== -1 &&
          (map.unitsPerPack !== -1 || map.price !== -1 || map.inTransit !== -1)
        )
          break;
      }
    }
    return map;
  }

  function _normalizeHeaderName_(value) {
    if (value === null || value === undefined) return '';
    return String(value).trim().toLowerCase();
  }

  function _parseNumber(value) {
    if (value === null || value === undefined) return null;
    if (typeof value === 'number') return isNaN(value) ? null : value;
    var str = String(value).trim().replace(/\s+/g, '').replace(',', '.');
    var num = parseFloat(str);
    return isNaN(num) ? null : num;
  }

  function _calculatePackValue_(
    preOrder,
    unitsPerPack,
    existingPack,
    forceFromPreOrder
  ) {
    var hasPreOrder =
      preOrder !== null && preOrder !== undefined && !isNaN(preOrder);
    var currentPack =
      existingPack !== null &&
      existingPack !== undefined &&
      !isNaN(existingPack)
        ? existingPack
        : null;
    if (forceFromPreOrder) {
      if (!hasPreOrder) return preOrder === 0 ? 0 : null;
      if (
        unitsPerPack !== null &&
        unitsPerPack !== undefined &&
        !isNaN(unitsPerPack) &&
        unitsPerPack > 0
      ) {
        var ratio = preOrder / unitsPerPack;
        if (ratio > 0 && ratio < 1) return 1;
        return Math.round(ratio);
      }
      return preOrder;
    }
    return currentPack;
  }

  function _calculateOrderUnits_(pack, unitsPerPack) {
    if (pack === null || pack === undefined || isNaN(pack)) return null;
    if (
      unitsPerPack === null ||
      unitsPerPack === undefined ||
      isNaN(unitsPerPack) ||
      unitsPerPack <= 0
    )
      return pack;
    return pack * unitsPerPack;
  }

  function _calculatePreliminarySum_(orderUnits, price) {
    if (orderUnits === null || orderUnits === undefined || isNaN(orderUnits))
      return null;
    if (price === null || price === undefined || isNaN(price)) return 0;
    return orderUnits * price;
  }

  function _formatNumberOrBlank_(value) {
    if (value === null || value === undefined || isNaN(value)) return '';
    return value;
  }

  function _fmtThousands_(value) {
    var num = Number(value || 0);
    var sign = num < 0 ? '-' : '';
    var abs = Math.abs(num);
    // ИЗМЕНЕНО: Округляем до целого числа для заголовка "Предварительная сумма заказа"
    var rounded = Math.round(abs);
    var formatted = rounded.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ' ');
    return sign + formatted;
  }

  function _formatEuroLabel_(value) {
    return _fmtThousands_(value) + ' €';
  }

  function _recalculateColumnTotal_(
    sheet,
    columnIndex,
    headerRowIndex,
    numberFormat
  ) {
    if (!sheet || !columnIndex) return 0;
    var firstDataRow = (headerRowIndex || 3) + 1;
    var lastRow = sheet.getLastRow();
    if (lastRow < firstDataRow) {
      return 0;
    }
    var numRows = lastRow - firstDataRow + 1;
    var values = sheet
      .getRange(firstDataRow, columnIndex, numRows, 1)
      .getValues();
    var total = 0;
    var hasValue = false;
    for (var i = 0; i < values.length; i++) {
      var parsed = _parseNumber(values[i][0]);
      if (parsed !== null && !isNaN(parsed)) {
        total += parsed;
        hasValue = true;
      }
    }
    return hasValue ? total : 0;
  }

  /**
   * ВНУТРЕННЯЯ: Автоматическое заполнение столбцов "Срок#" и "Срок"
   * на основе данных из "СГ 1", "СГ 2", "СГ 3"
   */
  function _autoFillExpiryDates(sheet, row, columnType) {
    try {
      var lastColumn = sheet.getLastColumn();
      var headersInfo = _collectOrderHeaders_(sheet, lastColumn, 3);
      var headers = headersInfo.effective;

      // Определяем индексы нужных столбцов
      var sg1Index = -1, sg2Index = -1, sg3Index = -1;
      var targetExpiryIndex = -1;

      for (var i = 0; i < headers.length; i++) {
        var h = _normalizeHeaderName_(headers[i]);
        if (h === 'сг 1' || h === 'сг1') sg1Index = i;
        else if (h === 'сг 2' || h === 'сг2') sg2Index = i;
        else if (h === 'сг 3' || h === 'сг3') sg3Index = i;
        else if (columnType === 'акции' && (h === 'срок#' || h === 'срок №')) targetExpiryIndex = i;
        else if (columnType === 'набор' && h === 'срок') {
          // Проверяем, что это не "Срок#"
          var rawHeader = headers[i];
          if (String(rawHeader).indexOf('#') === -1 && String(rawHeader).indexOf('№') === -1) {
            targetExpiryIndex = i;
          }
        }
      }

      Lib.logDebug(
        '[AvtoZakaz] _autoFillExpiryDates: sg1=' + sg1Index +
        ', sg2=' + sg2Index +
        ', sg3=' + sg3Index +
        ', target=' + targetExpiryIndex +
        ' (columnType=' + columnType + ')'
      );

      if (targetExpiryIndex === -1) {
        Lib.logWarn('[AvtoZakaz] _autoFillExpiryDates: целевой столбец не найден');
        return;
      }

      if (sg1Index === -1 && sg2Index === -1 && sg3Index === -1) {
        Lib.logWarn('[AvtoZakaz] _autoFillExpiryDates: столбцы СГ не найдены');
        return;
      }

      // Читаем значения из столбцов СГ
      var rowValues = sheet.getRange(row, 1, 1, lastColumn).getValues()[0];
      var expiryDates = [];

      if (sg1Index !== -1) {
        var sg1Value = rowValues[sg1Index];
        if (sg1Value !== null && sg1Value !== undefined && String(sg1Value).trim() !== '') {
          expiryDates.push(_formatExpiryDate(sg1Value));
        }
      }

      if (sg2Index !== -1) {
        var sg2Value = rowValues[sg2Index];
        if (sg2Value !== null && sg2Value !== undefined && String(sg2Value).trim() !== '') {
          expiryDates.push(_formatExpiryDate(sg2Value));
        }
      }

      if (sg3Index !== -1) {
        var sg3Value = rowValues[sg3Index];
        if (sg3Value !== null && sg3Value !== undefined && String(sg3Value).trim() !== '') {
          expiryDates.push(_formatExpiryDate(sg3Value));
        }
      }

      // Формируем строку с датами через запятую, каждое значение с новой строки
      var resultValue = expiryDates.length > 0 ? expiryDates.join(',\n') : '';

      // Записываем значение в целевой столбец
      sheet.getRange(row, targetExpiryIndex + 1).setValue(resultValue);

      Lib.logDebug(
        '[AvtoZakaz] _autoFillExpiryDates: записано значение в строку ' +
        row + ', столбец ' + (targetExpiryIndex + 1) + ': "' + resultValue + '"'
      );

    } catch (err) {
      Lib.logError('[AvtoZakaz] _autoFillExpiryDates: ошибка', err);
    }
  }

  /**
   * ВНУТРЕННЯЯ: Очистка столбца "Срок#" или "Срок"
   */
  function _clearExpiryDate(sheet, row, columnType) {
    try {
      var lastColumn = sheet.getLastColumn();
      var headersInfo = _collectOrderHeaders_(sheet, lastColumn, 3);
      var headers = headersInfo.effective;

      var targetExpiryIndex = -1;

      for (var i = 0; i < headers.length; i++) {
        var h = _normalizeHeaderName_(headers[i]);
        if (columnType === 'акции' && (h === 'срок#' || h === 'срок №')) {
          targetExpiryIndex = i;
          break;
        } else if (columnType === 'набор' && h === 'срок') {
          var rawHeader = headers[i];
          if (String(rawHeader).indexOf('#') === -1 && String(rawHeader).indexOf('№') === -1) {
            targetExpiryIndex = i;
            break;
          }
        }
      }

      if (targetExpiryIndex !== -1) {
        sheet.getRange(row, targetExpiryIndex + 1).clearContent();
        Lib.logDebug('[AvtoZakaz] _clearExpiryDate: очищен столбец ' + (targetExpiryIndex + 1) + ' в строке ' + row);
      }
    } catch (err) {
      Lib.logError('[AvtoZakaz] _clearExpiryDate: ошибка', err);
    }
  }

  /**
   * ВНУТРЕННЯЯ: Форматирование даты для вывода
   */
  function _formatExpiryDate(value) {
    if (!value) return '';

    // Если это Date объект
    if (value instanceof Date && !isNaN(value.getTime())) {
      var month = String(value.getMonth() + 1).padStart(2, '0');
      var year = value.getFullYear();
      return month + '/' + year;
    }

    // Если это строка
    var str = String(value).trim();
    if (!str) return '';

    // Пытаемся распознать паттерны типа "MM/YYYY" или "DD.MM.YYYY"
    var match = str.match(/(\d{1,2})[\/\.\-](\d{4})/);
    if (match) {
      var m = String(match[1]).padStart(2, '0');
      return m + '/' + match[2];
    }

    // Возвращаем как есть
    return str;
  }

})(Lib, this);
