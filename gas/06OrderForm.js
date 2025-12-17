var Lib = Lib || {};

(function (Lib, global) {
  const GROUP_HEADER_FONT = "#FFFFFF";
  const GROUP_HEADER_BG = "#666666"; // Тепло-серый (3)
  const HEADER_ROW_INDEX = 1;
  const DATA_START_ROW = HEADER_ROW_INDEX + 1;
  const DEFAULT_COLUMNS_TO_COPY = [
    "ID",
    "ID-P",
    "Статус",
    "Арт. Рус",
    "Арт. произв.",
    "Название  ENG прайс произв",
    "Наименования рус по ДС",
    "Объём",
    "шт./уп.",
    "ПРОДАЖИ",
    "Остаток",
    "товар в  ПУТИ",
    "Остаток 1",
    "СГ 1",
    "Остаток 2",
    "СГ 2",
    "Остаток3",
    "СГ 3",
    "СПИСАНО",
    "РЕЗЕРВ",
    "Среднемесячные продажи, шт",
    "Потребность на 6 месяцев, шт",
    "На сколько мес запас",
    "Необходимо заказать, шт",
    "Предварительный заказ",
    "заказ в упаковках",
    "ЗАКАЗ",
    "EXW  ALFASPA  текущая, €",
    "Предварительная сумма заказа 0,00 €",
    "АКЦИИ",
    "условие#",
    "коментарий#",
    "Срок#",
    "Набор",
    "условие",
    "коментарий",
    "Срок",
  ];
  const FIELDS = {
    LINE: "Линия",
    GROUP: "Группа",
    IDP: "ID-P",
    IDG: "ID-G",
    TITLE: "Название  ENG прайс произв",
  };

  const RU_COLLATOR =
    typeof Intl !== "undefined" && typeof Intl.Collator === "function"
      ? new Intl.Collator("ru", { sensitivity: "base" })
      : null;

  /**
   * УСТАРЕВШАЯ: Пересборка заказа из листа Главная с группировкой по ID-G
   * Сейчас используется только для обратной совместимости
   * @deprecated Используйте Lib.structureOrderSheet вместо этой функции
   */
  Lib.rebuildOrderFromPrimary = function () {
    try {
      Lib.structureOrderSheet('byManufacturer');
    } catch (error) {
      Lib.logError('rebuildOrderFromPrimary: ошибка', error);
      _toastOrder_(error.message || String(error));
    }
  };

  /**
   * ПУБЛИЧНАЯ: Универсальная функция структурирования нескольких листов одновременно
   * @param {string} mode - Режим группировки: 'byManufacturer' (по Производителю/Группе) или 'byPrice' (по Прайсу/Линии)
   */
  Lib.structureMultipleSheets = function (mode) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) {
        throw new Error('Spreadsheet недоступен');
      }

      const sheetsToSort = [
        Lib.CONFIG.SHEETS.ORDER_FORM,
        Lib.CONFIG.SHEETS.PRICE_DYNAMICS,
        Lib.CONFIG.SHEETS.PRICE_CALCULATION
      ];

      const modeLabel = mode === 'byPrice' ? 'по Прайсу' : 'по Производителю';
      let successCount = 0;
      let errorCount = 0;
      const errors = [];

      sheetsToSort.forEach(function(sheetName) {
        try {
          Lib.logInfo('[MULTI-SORT] Начинаем сортировку листа: ' + sheetName + ' ' + modeLabel);
          Lib.structureOrderSheetInternal(mode, sheetName);
          successCount++;
          Lib.logInfo('[MULTI-SORT] Лист ' + sheetName + ' успешно отсортирован');
        } catch (error) {
          errorCount++;
          const errorMsg = 'Ошибка сортировки листа "' + sheetName + '": ' + (error.message || String(error));
          errors.push(errorMsg);
          Lib.logError('[MULTI-SORT] ' + errorMsg, error);
        }
      });

      if (errorCount > 0) {
        _toastOrder_(
          'Сортировка ' + modeLabel + ' завершена с ошибками.\n' +
          'Успешно: ' + successCount + ' из ' + sheetsToSort.length + ' листов.\n' +
          'Ошибки:\n' + errors.join('\n')
        );
      } else {
        _toastOrder_(
          'Сортировка ' + modeLabel + ' успешно завершена для всех ' + successCount + ' листов:\n' +
          sheetsToSort.join(', ')
        );
      }

      Lib.logInfo('[MULTI-SORT] Завершено. Успешно: ' + successCount + ', ошибок: ' + errorCount);
    } catch (error) {
      Lib.logError('[MULTI-SORT] Общая ошибка', error);
      _toastOrder_('Ошибка при сортировке: ' + (error.message || String(error)));
    }
  };

  /**
   * ПУБЛИЧНАЯ: Универсальная функция структурирования листа Заказ (обратная совместимость)
   * @param {string} mode - Режим группировки: 'byManufacturer' (по Производителю/Группе) или 'byPrice' (по Прайсу/Линии)
   */
  Lib.structureOrderSheet = function (mode) {
    try {
      Lib.structureOrderSheetInternal(mode, Lib.CONFIG.SHEETS.ORDER_FORM);
    } catch (error) {
      Lib.logError('structureOrderSheet: ошибка', error);
      _toastOrder_(error.message || String(error));
    }
  };

  /**
   * ВНУТРЕННЯЯ: Универсальная функция структурирования указанного листа
   * @param {string} mode - Режим группировки: 'byManufacturer' (по Производителю/Группе) или 'byPrice' (по Прайсу/Линии)
   * @param {string} targetSheetName - Название листа для сортировки
   */
  Lib.structureOrderSheetInternal = function (mode, targetSheetName) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      throw new Error('Spreadsheet недоступен');
    }

    const primaryName = Lib.CONFIG.SHEETS.PRIMARY;

    const primarySheet = ss.getSheetByName(primaryName);
    const targetSheet = ss.getSheetByName(targetSheetName);
    if (!primarySheet || !targetSheet) {
      throw new Error(
        'Не найдены листы «' +
          primaryName +
          '» или «' +
          targetSheetName +
          '».'
      );
    }

    // Читаем данные С ЦЕЛЕВОГО ЛИСТА
    const targetRange = targetSheet.getDataRange();
    const targetValues = targetRange.getValues();
    if (!targetValues || targetValues.length <= 1) {
      throw new Error('Нет данных на листе «' + targetSheetName + '» для сортировки.');
    }

    const targetHeaders = targetValues[0].map((value) => String(value || '').trim());
    const targetIndex = _makeHeaderIndex_(targetHeaders);

    // Читаем данные из Главная для получения группировочных полей
    const primaryRange = primarySheet.getDataRange();
    const primaryValues = primaryRange.getValues();
    if (!primaryValues || primaryValues.length <= 1) {
      throw new Error('Нет данных на листе «' + primaryName + '» для определения групп.');
    }

    const primaryHeaders = primaryValues[0].map((value) => String(value || '').trim());
    const primaryIndex = _makeHeaderIndex_(primaryHeaders);

    // Создаем карту ID → (ID-G, ID-L, Линия, Группа) из листа Главная
    const idToGroupMap = new Map();
    for (let i = 1; i < primaryValues.length; i++) {
      const row = primaryValues[i];
      if (_isRowEmpty_(row)) continue;

      const idIndex = primaryIndex['ID'];
      if (idIndex === undefined) continue;

      const id = String(row[idIndex] || '').trim();
      if (!id) continue;

      const idgIndex = primaryIndex[FIELDS.IDG];
      const idlIndex = primaryIndex['ID-L'];
      const lineIndex = primaryIndex[FIELDS.LINE];
      const groupIndex = primaryIndex[FIELDS.GROUP];

      idToGroupMap.set(id, {
        idg: idgIndex !== undefined ? row[idgIndex] : null,
        idl: idlIndex !== undefined ? row[idlIndex] : null,
        line: lineIndex !== undefined ? row[lineIndex] : null,
        group: groupIndex !== undefined ? row[groupIndex] : null,
      });
    }

    // Для режима byPrice создаем две карты из листа Прайс:
    // 1. ID → (ID-L, Группа линии, Линия Прайс) - для получения ID-L и названия группы по ID товара
    // 2. ID-L → (Группа линии, Линия Прайс) - для формирования названия группы
    const idToPriceInfoMap = new Map();
    const idlToLineInfoMap = new Map();
    if (mode === 'byPrice') {
      const priceSheet = ss.getSheetByName('Прайс');
      if (priceSheet) {
        const priceRange = priceSheet.getDataRange();
        const priceValues = priceRange.getValues();
        if (priceValues && priceValues.length > 1) {
          const priceHeaders = priceValues[0].map((value) => String(value || '').trim());
          const priceIndex = _makeHeaderIndex_(priceHeaders);

          const idPriceIndex = priceIndex['ID'];
          const idlPriceIndex = priceIndex['ID-L'];
          const groupLineIndex = priceIndex['Группа линии'];
          const linePriceIndex = priceIndex['Линия Прайс'];

          Lib.logInfo('[ORDER] Индексы колонок в листе Прайс: ID=' + idPriceIndex + ', ID-L=' + idlPriceIndex + ', Группа линии=' + groupLineIndex + ', Линия Прайс=' + linePriceIndex);

          if (idPriceIndex !== undefined && idlPriceIndex !== undefined) {
            for (let i = 1; i < priceValues.length; i++) {
              const row = priceValues[i];
              if (_isRowEmpty_(row)) continue;

              const idRaw = row[idPriceIndex];
              const idStr = String(idRaw || '').trim();
              if (!idStr) continue;

              const idlRaw = row[idlPriceIndex];
              const groupLine = groupLineIndex !== undefined ? String(row[groupLineIndex] || '').trim() : '';
              const linePrice = linePriceIndex !== undefined ? String(row[linePriceIndex] || '').trim() : '';

              // Добавляем в карту ID → информация о прайсе
              idToPriceInfoMap.set(idStr, {
                idl: idlRaw,
                groupLine: groupLine,
                linePrice: linePrice,
              });

              // Добавляем в карту ID-L → название группы (только первое вхождение)
              if (idlRaw !== '' && idlRaw !== null && !isNaN(Number(idlRaw))) {
                const idlNum = Number(idlRaw);
                if (!idlToLineInfoMap.has(idlNum)) {
                  idlToLineInfoMap.set(idlNum, {
                    groupLine: groupLine,
                    linePrice: linePrice,
                  });
                  Lib.logDebug('[ORDER] Добавлен ID-L=' + idlNum + ': "' + groupLine + '" / "' + linePrice + '"');
                }
              }
            }
            Lib.logInfo('[ORDER] Карта ID создана из листа Прайс, записей: ' + idToPriceInfoMap.size + ', уникальных линий: ' + idlToLineInfoMap.size);
          }
        }
      } else {
        Lib.logWarn('[ORDER] Лист "Прайс" не найден, сортировка будет выполнена без данных из прайса');
      }
    }

    const titleColumnIndex = Math.max(targetHeaders.indexOf(FIELDS.TITLE), 0);
    const idColumnIndex = targetIndex['ID'];
    const idpColumnIndex = targetIndex[FIELDS.IDP];
    const dsNameColumnIndex = targetIndex['Наименования рус по ДС'];

    if (idColumnIndex === undefined) {
      throw new Error('Столбец "ID" не найден на листе «' + targetSheetName + '»');
    }

    const groups = new Map();

    // Группируем строки из целевого листа в зависимости от режима
    for (let rowIndex = 1; rowIndex < targetValues.length; rowIndex++) {
      const row = targetValues[rowIndex];
      if (_isRowEmpty_(row)) continue;

      const id = String(row[idColumnIndex] || '').trim();
      if (!id) continue;

      const groupInfo = idToGroupMap.get(id);
      const idpRaw = idpColumnIndex !== undefined ? row[idpColumnIndex] : null;

      let groupKey = 'UNASSIGNED';
      let groupTitle = '';
      let groupSortValue = null;

      if (mode === 'byPrice') {
        // Сортировка по Прайсу: группировка по ID-L из листа Прайс
        const priceInfo = idToPriceInfoMap.get(id);
        const idlRaw = priceInfo ? priceInfo.idl : null;

        const hasIdL = idlRaw !== '' && idlRaw !== null && !isNaN(Number(idlRaw));

        if (hasIdL) {
          const idlNum = Number(idlRaw);
          groupKey = 'LINE_' + String(idlNum);
          groupSortValue = idlNum;

          // Формируем название группы из данных листа Прайс
          const lineInfo = idlToLineInfoMap.get(idlNum);
          if (lineInfo && (lineInfo.groupLine || lineInfo.linePrice)) {
            const parts = [];
            if (lineInfo.groupLine) parts.push(lineInfo.groupLine);
            if (lineInfo.linePrice) parts.push(lineInfo.linePrice);
            groupTitle = parts.join('\n');

            // Логируем создание новой группы
            if (!groups.has(groupKey)) {
              Lib.logDebug('[ORDER] Создается новая группа: ID-L=' + idlNum + ', название="' + groupTitle.replace(/\n/g, ' / ') + '"');
            }
          } else {
            groupTitle = 'Группа не определена';
            Lib.logWarn('[ORDER] ID-L=' + idlNum + ' не найден в карте названий линий');
          }
        } else {
          groupKey = 'UNASSIGNED';
          groupTitle = 'Группа не идентифицирована';
          if (priceInfo) {
            Lib.logDebug('[ORDER] Товар ID=' + id + ' не имеет ID-L в листе Прайс (idlRaw=' + idlRaw + ')');
          } else {
            Lib.logDebug('[ORDER] Товар ID=' + id + ' не найден в листе Прайс');
          }
        }
      } else {
        // Сортировка по Производителю: группировка по Группе (ID-G) - по умолчанию
        const idgRaw = groupInfo ? groupInfo.idg : null;
        const groupRaw = groupInfo ? groupInfo.group : null;

        const hasIdG = idgRaw !== '' && idgRaw !== null && !isNaN(Number(idgRaw));

        if (hasIdG) {
          groupKey = 'GROUP_' + String(Number(idgRaw));
          groupTitle = (groupRaw && String(groupRaw).trim()) || 'Группа не определена';
          groupSortValue = Number(idgRaw);
        } else {
          const fallbackTitle = (groupRaw && String(groupRaw).trim()) || '';
          if (fallbackTitle) {
            groupKey = 'META_' + fallbackTitle;
            groupTitle = fallbackTitle;
          } else {
            groupKey = 'UNASSIGNED';
            groupTitle = 'Группа не идентифицирована';
          }
        }
      }

      if (!groups.has(groupKey)) {
        groups.set(groupKey, {
          title: groupTitle,
          sortValue: groupSortValue,
          rows: [],
        });
      }

      const hasIdP = idpRaw !== '' && idpRaw !== null && !isNaN(Number(idpRaw));
      const idpNum = hasIdP ? Number(idpRaw) : null;
      const fallbackIndex = rowIndex + 1; // исходный порядок на целевом листе
      const dsName = dsNameColumnIndex !== undefined && dsNameColumnIndex !== -1
        ? String(row[dsNameColumnIndex] || '').trim()
        : '';

      groups.get(groupKey).rows.push({
        hasIdP: hasIdP,
        idp: idpNum,
        fallbackIndex: fallbackIndex,
        dsName: dsName,
        values: row,
      });
    }

    // Сортируем группы
    const sortedKeys = Array.from(groups.keys()).sort((a, b) => {
      if (a === 'UNASSIGNED' && b === 'UNASSIGNED') return 0;
      if (a === 'UNASSIGNED') return 1;
      if (b === 'UNASSIGNED') return -1;

      const groupA = groups.get(a);
      const groupB = groups.get(b);

      // Сначала пытаемся сортировать по sortValue (числовой ID)
      if (groupA.sortValue !== null && groupB.sortValue !== null) {
        return groupA.sortValue - groupB.sortValue;
      }
      if (groupA.sortValue !== null) return -1;
      if (groupB.sortValue !== null) return 1;

      // Fallback на строковую сортировку
      return String(a).localeCompare(String(b));
    });

    // Очищаем данные (но сохраняем заголовки)
    _clearDestination_(targetSheet);

    let writeRow = DATA_START_ROW;
    const headerRows = [];

    // Записываем отсортированные данные
    sortedKeys.forEach((key) => {
      const group = groups.get(key);
      const headerRow = new Array(targetHeaders.length).fill('');
      headerRow[titleColumnIndex] = group.title || '';
      targetSheet.getRange(writeRow, 1, 1, targetHeaders.length).setValues([headerRow]);
      headerRows.push(writeRow);
      writeRow += 1;

      let sortedRows;
      if (mode === 'byPrice') {
        // Сортировка по Прайсу: сортировка внутри группы по ID-P (от меньшего к большему)
        // Товары с ID-P идут первыми (отсортированные по ID-P)
        // Товары без ID-P идут в конце (отсортированные по исходному порядку)
        const withIdP = group.rows
          .filter((r) => r.hasIdP)
          .sort((a, b) => (a.idp - b.idp));

        const withoutIdP = group.rows
          .filter((r) => !r.hasIdP)
          .sort((a, b) => (a.fallbackIndex - b.fallbackIndex));

        sortedRows = withIdP.concat(withoutIdP).map(function (entry) {
          return entry.values;
        });
      } else {
        // Улучшенное поведение для сортировки по Производителю:
        // 1) Внутри группы держим порядок по ID-P
        // 2) Строки без ID-P вставляем сразу после совпадения по "Наименования рус по ДС" (если найдено)
        const normalizeDs = function (v) { return String(v || '').trim().toLowerCase(); };

        const withIdP = group.rows
          .filter((r) => r.hasIdP)
          .sort((a, b) => (a.idp - b.idp));

        const withoutIdP = group.rows
          .filter((r) => !r.hasIdP)
          .sort((a, b) => (a.fallbackIndex - b.fallbackIndex));

        const arranged = withIdP.slice();

        const findLastIndexByDs = function (arr, key) {
          if (!key) return -1;
          for (let i = arr.length - 1; i >= 0; i--) {
            if (normalizeDs(arr[i].dsName) === key) return i;
          }
          return -1;
        };

        withoutIdP.forEach((item) => {
          const key = normalizeDs(item.dsName);
          const lastIdx = findLastIndexByDs(arranged, key);
          if (lastIdx >= 0) {
            arranged.splice(lastIdx + 1, 0, item);
          } else {
            arranged.push(item);
          }
        });

        sortedRows = arranged.map((entry) => entry.values);
      }

      if (sortedRows.length) {
        const targetRange = targetSheet.getRange(
          writeRow,
          1,
          sortedRows.length,
          targetHeaders.length
        );
        targetRange.setValues(sortedRows);
        writeRow += sortedRows.length;
      }
    });

    // Форматируем заголовки групп
    const headerBgColor = mode === 'byPrice' ? '#7f6000' : GROUP_HEADER_BG;
    headerRows.forEach((rowNumber) => {
      targetSheet
        .getRange(rowNumber, 1, 1, targetHeaders.length)
          .setBackground(headerBgColor)
          .setFontWeight('bold')
          .setFontColor(GROUP_HEADER_FONT);
    });

    const modeLabel = mode === 'byPrice' ? 'по Прайсу' : 'по Производителю';
    Lib.logInfo(
      '[ORDER] Лист "' + targetSheetName + '" структурирован ' +
        modeLabel +
        ': групп ' +
        sortedKeys.length +
        ', строк ' +
        (writeRow - DATA_START_ROW) +
        '.'
    );
  };

  Lib.onUpdateOverride = function (sheet, row, col, debugSource) {
    try {
      if (!sheet || !Lib.CONFIG || !Lib.CONFIG.SHEETS) return;
      if (!col || col < 1) return;

      const orderSheetName = Lib.CONFIG.SHEETS.ORDER_FORM;
      if (sheet.getName() === orderSheetName) {
        // Делегируем обработку изменений на листе Заказ в модуль AvtoZakaz
        if (typeof Lib.handleAvtoZakazUpdate === 'function') {
          Lib.handleAvtoZakazUpdate(sheet, row, col, debugSource);
        }
        return;
      }

      const primarySheetName = Lib.CONFIG.SHEETS.PRIMARY;
      if (sheet.getName() !== primarySheetName) return;

      // На листе Главная при изменении Линия или Группа НЕ пересоздаём лист Заказ
      // (сортировка теперь выполняется вручную по кнопке)
    } catch (error) {
      Lib.logError('onUpdateOverride: ошибка', error);
    }
  };

  function _ensureDestinationHeader_(sheet, columns) {
    sheet.getRange(1, 1, 1, columns.length).setValues([columns]);
    const lastColumn = sheet.getLastColumn();
    if (lastColumn > columns.length) {
      sheet.deleteColumns(columns.length + 1, lastColumn - columns.length);
    }
  }

  function _clearDestination_(sheet) {
    const lastRow = Math.max(DATA_START_ROW, sheet.getMaxRows());
    const lastCol = Math.max(1, sheet.getMaxColumns());
    if (lastRow >= DATA_START_ROW) {
      sheet
        .getRange(DATA_START_ROW, 1, lastRow - (DATA_START_ROW - 1), lastCol)
        .clearContent()
        .setBackground(null)
        .setFontColor(null)
        .setFontWeight('normal');
    }
  }

  function _makeHeaderIndex_(headers) {
    const map = {};
    headers.forEach((name, index) => {
      const key = String(name).trim();
      if (key && !(key in map)) {
        map[key] = index;
      }
    });
    return map;
  }

  function _isRowEmpty_(row) {
    for (let i = 0; i < row.length; i++) {
      if (row[i] !== '' && row[i] !== null) {
        return false;
      }
    }
    return true;
  }

  function _toastOrder_(message) {
    try {
      SpreadsheetApp.getActive().toast(String(message || ''), 'Сборка заказа', 5);
    } catch (_) {
      // Используем logWithEmoji если доступна, иначе fallback
      if (typeof global.logWithEmoji === 'function') {
        global.logWithEmoji(String(message), 'INFO', null, 'OrderForm');
      } else {
        Logger.log(message);
      }
    }
  }

  /**
   * Обработчик override для onUpdate: очищает строки 4 и 5 при изменении заголовка 'Потребность на N месяцев, шт' в строке 3
   */
  Lib.handleOrderUpdate = function (sheet, row, col, debugSource) {
    try {
      if (!sheet || !global.CONFIG) {
        Lib.logDebug("[Order] handleOrderUpdate: sheet/CONFIG отсутствуют");
        return;
      }

      var ORDER_SHEET =
        (global.CONFIG.SHEETS && global.CONFIG.SHEETS.ORDER_FORM) || "Заказ";
      if (sheet.getName() !== ORDER_SHEET) return;

      var HEADER_ROW = 1; // единственная строка заголовков колонок
      var FIRST_DATA_ROW = 2; // фактические данные начинаются со 2 строки

      function readHeaderAt(r, c) {
        // value -> displayValue -> merged
        var v = String(sheet.getRange(r, c).getValue() || "").trim();
        if (v) return v;
        try {
          var d = String(sheet.getRange(r, c).getDisplayValue() || "").trim();
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
              return String(mv || "").trim();
            }
          }
        } catch (_) {}
        return "";
      }

      // 1) Правка строки 1 (заголовки колонок)
      if (row === HEADER_ROW) {
        var topHeader = readHeaderAt(HEADER_ROW, col);
        var m = topHeader.match(/потребность\s+на\s*(\d+)\s*месяц/i);
        if (topHeader && (topHeader.toLowerCase() === "кол-во месяцев" || m)) {
          Lib.logDebug(
            "[Order] handleOrderUpdate: строка 1 изменена: " + topHeader
          );

          if (m) {
            var n = parseInt(m[1], 10);
            Lib.logDebug(
              "[Order] handleOrderUpdate: найден коэффициент N=" +
                n +
                ", начинаем пересчет"
            );

            var lastRow = sheet.getLastRow();
            var lastCol = sheet.getLastColumn();
            var headersInfo = _collectOrderHeaders_(sheet, lastCol, HEADER_ROW);
            var headerMap = _buildOrderHeaderIndexMap(headersInfo);
            var needCol = col; // Текущий столбец
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
                "[Order] handleOrderUpdate: найден столбец '" +
                  (avgLabel || "Среднемесячные продажи, шт") +
                  "' #" +
                  avgCol +
                  ", пересчитываем " +
                  rowsToProcess +
                  " строк"
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
                  needOrderValues.push([need - stock - transit - reserve]);
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
              Lib.logDebug(
                "[Order] handleOrderUpdate: пересчет завершен"
              );
            } else if (rowsToProcess === 0) {
              Lib.logDebug(
                "[Order] handleOrderUpdate: данных для пересчёта нет"
              );
            } else {
              Lib.logWarn(
                "[Order] handleOrderUpdate: не найден столбец 'Среднемесячные продажи'"
              );
            }
          } else {
            _populateDerivedColumns(col);
          }
        }
        return;
      }

      // 2) Правка строк данных (row > 1) — пересчёт по конкретному полю
      if (row > HEADER_ROW) {
        // Определяем имя столбца из строки 1 (единственной шапки)
        var h = readHeaderAt(HEADER_ROW, col);
        var n = _normalizeHeaderName_(h);
        var trigger = null;
        var compactName = n.replace(/\s+/g, '');

        var shouldRefreshDerived =
          n === 'остаток' ||
          n === 'резерв' ||
          compactName.indexOf('товарвпути') === 0 ||
          compactName.indexOf('продажи') === 0 ||
          compactName.indexOf('среднемесячныепродажи') === 0 ||
          compactName.indexOf('насколькомесзапас') === 0;

        if (n.indexOf("предварительный заказ") === 0) trigger = "preOrder";
        else if (n.indexOf("заказ в упаковках") === 0) trigger = "packs";
        else if (n === "заказ") trigger = "order";
        else if (n.replace(/\s+/g, "") === "штуп" || n.indexOf("шт уп") !== -1)
          trigger = "unitsPerPack";
        else if (n.indexOf("exw alfaspa текущая") === 0) trigger = "price";
        else if (n.indexOf("предварительная сумма заказа") === 0)
          trigger = "sum";

        // Если не удалось распознать триггер по заголовкам, попробуем извлечь его из debugSource
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
            "[Order] handleOrderUpdate: пересчёт производных столбцов для R" +
              row +
              "C" +
              col +
              " (" +
              h +
              ")"
          );
          _populateDerivedColumns();
          return;
        }

        if (trigger) {
          Lib.logDebug(
            "[Order] handleOrderUpdate: trigger=" +
              trigger +
              " for R" +
              row +
              "C" +
              col
          );
          _recalculateOrderRow_(sheet, row, { trigger: trigger });
        }
      }
    } catch (err) {
      Lib.logError("handleOrderUpdate: ошибка", err);
    }
  };

  function _populateDerivedColumns(monthsColumnOverride) {
    try {
      Lib.logDebug(
        "[Order] _populateDerivedColumns: старт (override=" +
          (typeof monthsColumnOverride === "number"
            ? monthsColumnOverride
            : "нет") +
          ")"
      );
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) {
        return;
      }

      var orderSheetName =
        (global.CONFIG.SHEETS && global.CONFIG.SHEETS.ORDER_FORM) || "Заказ";
      var orderSheet = ss.getSheetByName(orderSheetName);
      if (!orderSheet) {
        Lib.logWarn("[Order] _populateDerivedColumns: лист заказа не найден");
        return;
      }

      var headerRowIndex = HEADER_ROW_INDEX;
      var firstDataRow = _resolveFirstDataRow_(orderSheet, headerRowIndex);
      var lastRow = orderSheet.getLastRow();
      if (lastRow < firstDataRow) {
        Lib.logDebug(
          "[Order] _populateDerivedColumns: нет строк для обновления"
        );
        return;
      }

      var lastColumn = orderSheet.getLastColumn();
      var headersInfo = _collectOrderHeaders_(orderSheet, lastColumn, headerRowIndex);
      var headersEffective = headersInfo.effective;

      var idx = _buildOrderHeaderIndexMap(headersInfo);

      Lib.logDebug(
        "[Order] _populateDerivedColumns: header indices: " +
          JSON.stringify(idx)
      );
      Lib.logDebug(
        '[Order] _populateDerivedColumns: header at need index="' +
          (idx.need !== -1 ? headersEffective[idx.need] || "" : "") +
          '"'
      );

      if (idx.need === -1) {
        Lib.logWarn(
          '[Order] _populateDerivedColumns: столбец "Потребность на 6 месяцев, шт" не найден'
        );
        return;
      }

      if (idx.avg === -1) {
        Lib.logWarn(
          '[Order] _populateDerivedColumns: столбец "Среднемесячные продажи, шт" не найден'
        );
        return;
      }

      var monthsColumn = null;
      if (typeof monthsColumnOverride === "number") {
        monthsColumn = monthsColumnOverride;
      } else {
        var topRow = headersInfo.top.map(function (value) {
          return String(value || "")
            .trim()
            .toLowerCase();
        });
        var found = topRow.indexOf("кол-во месяцев");
        if (found !== -1) {
          monthsColumn = found + 1;
        }
      }

      var monthsFactor = null;
      if (monthsColumn !== null) {
        // В новой модели коэффициент извлекаем только из названия колонки в строке 1
        var headerFromText = headersInfo.effective[monthsColumn - 1] || "";
        var matchFromHeader = headerFromText.match(/потребность\s+на\s*(\d+)/i);
        if (matchFromHeader) {
          monthsFactor = parseInt(matchFromHeader[1], 10);
        }
      }

      if ((monthsFactor === null || monthsFactor <= 0) && idx.need !== -1) {
        var needHeader = _resolveHeaderLabel_(headersInfo, idx.need) || "";
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
        "[Order] _populateDerivedColumns: коэффициент = " + monthsFactor
      );

      var numRows = lastRow - firstDataRow + 1;

      // Для отладки: снимем первые 10 старых значений столбца 'need' и 'needOrder'
      try {
        var sampleCount = Math.min(10, numRows);
        if (sampleCount > 0) {
          var oldNeed = orderSheet
            .getRange(firstDataRow, idx.need + 1, sampleCount, 1)
            .getValues()
            .map(function (r) {
              return r[0];
            });
          Lib.logDebug(
            "[Order] _populateDerivedColumns: old need sample (first " +
              sampleCount +
              "): " +
              JSON.stringify(oldNeed)
          );
        }
      } catch (e) {
        Lib.logDebug(
          "[Order] _populateDerivedColumns: не удалось прочитать старые значения need sample: " +
            e
        );
      }
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
          var needToOrder = need - stockValue - inTransitValue - reserveValue;
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
            Lib.logDebug(
              "[Order] _populateDerivedColumns: row " +
                (i + firstDataRow) +
                " pack computed=" +
                computedPack
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
            Lib.logDebug(
              "[Order] _populateDerivedColumns: row " +
                (i + firstDataRow) +
                " orderUnits computed=" +
                orderUnits
            );
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
            Lib.logDebug(
              "[Order] _populateDerivedColumns: row " +
                (i + firstDataRow) +
                " price=" +
                priceValue +
                ", prelimSum=" +
                preliminarySum
            );
            if (preliminarySum !== null && !isNaN(preliminarySum)) {
              sumAccumulator += preliminarySum;
              sumValues.push([preliminarySum]);
            } else {
              sumValues.push([""]);
            }
          }
        }
      }

      // Сравним старые значения целиком (если возможно) с новыми и подсчитаем изменёные
      try {
        var oldAllNeed = orderSheet
          .getRange(firstDataRow, idx.need + 1, numRows, 1)
          .getValues()
          .map(function (r) {
            return r[0];
          });
      } catch (e) {
        Lib.logDebug(
          "[Order] _populateDerivedColumns: не удалось прочитать старые значения need: " +
            e
        );
        var oldAllNeed = null;
      }

      if (avgOutputValues) {
        orderSheet
          .getRange(firstDataRow, idx.avg + 1, numRows, 1)
          .setValues(avgOutputValues)
          .setFontFamily("Century Gothic")
          .setFontSize(9)
          .setNumberFormat("0")
          .setBorder(false, false, false, false, false, false);
      }

      orderSheet
        .getRange(firstDataRow, idx.need + 1, numRows, 1)
        .setValues(needValues)
        .setFontFamily("Century Gothic")
        .setFontSize(9)
        .setNumberFormat("0")
        .setBorder(false, false, false, false, false, false);

      var updatedColumns = [];
      if (avgOutputValues) {
        updatedColumns.push("«Среднемесячные продажи…»");
      }
      updatedColumns.push("«Потребность…»");

      // Посчитаем сколько строк реально изменилось (если были старые значения)
      try {
        if (oldAllNeed) {
          var diffCount = 0;
          for (var di = 0; di < numRows; di++) {
            var oldv = oldAllNeed[di];
            var newv = needValues[di][0];
            if (oldv !== newv) diffCount++;
          }
          Lib.logDebug(
            "[Order] _populateDerivedColumns: количество изменённых строк в need = " +
              diffCount +
              " из " +
              numRows
          );
          // Если изменений нет — принудительно перезапишем столбец (часто помогает визуально обновить лист)
          if (diffCount === 0) {
            try {
              Lib.logDebug(
                "[Order] _populateDerivedColumns: diffCount==0, принудительная перезапись столбца need"
              );
              orderSheet
                .getRange(firstDataRow, idx.need + 1, numRows, 1)
                .clearContent();
              // неполное удаление форматирования, только содержимое
              orderSheet
                .getRange(firstDataRow, idx.need + 1, numRows, 1)
                .setValues(needValues);
            } catch (e) {
              Lib.logDebug(
                "[Order] _populateDerivedColumns: ошибка при принудительной перезаписи: " +
                  e
              );
            }
          }
        }
      } catch (e) {
        Lib.logDebug(
          "[Order] _populateDerivedColumns: ошибка при сравнении старых/новых значений: " +
            e
        );
      }

      if (needToOrderValues) {
        orderSheet
          .getRange(firstDataRow, idx.needOrder + 1, numRows, 1)
          .setValues(needToOrderValues)
          .setFontFamily("Century Gothic")
          .setFontSize(9)
          .setNumberFormat("0")
          .setBorder(false, false, false, false, false, false);
        updatedColumns.push("«Необходимо заказать…»");
      }

      if (monthsValues) {
        orderSheet
          .getRange(firstDataRow, idx.months + 1, numRows, 1)
          .setValues(monthsValues)
          .setFontFamily("Century Gothic")
          .setFontSize(9)
          .setNumberFormat('0.0 "мес."')
          .setBorder(false, false, false, false, false, false);
        updatedColumns.push("«На сколько мес запас»");
      }

      if (preOrderAccumulator !== null && idx.preOrder !== -1) {
        // Итог в строку 2 отключён по требованиям — ничего не пишем
      }

      if (packValues) {
        orderSheet
          .getRange(firstDataRow, idx.packs + 1, numRows, 1)
          .setValues(packValues)
          .setFontFamily("Century Gothic")
          .setFontSize(9)
          .setNumberFormat("0")
          .setBorder(false, false, false, false, false, false);
        updatedColumns.push("«заказ в упаковках»");
      }

      if (orderValues) {
        orderSheet
          .getRange(firstDataRow, idx.order + 1, numRows, 1)
          .setValues(orderValues)
          .setFontFamily("Century Gothic")
          .setFontSize(9)
          .setNumberFormat("0")
          .setBorder(false, false, false, false, false, false);
        updatedColumns.push("«ЗАКАЗ»");
      }

      if (sumValues) {
        orderSheet
          .getRange(firstDataRow, idx.sum + 1, numRows, 1)
          .setValues(sumValues)
          .setFontFamily("Century Gothic")
          .setFontSize(9)
          .setNumberFormat('#,##0.00"€"')
          .setBorder(false, false, false, false, false, false);
        // Обновляем подпись в строке 1 на основе фактической суммы колонки
        var newTotalForHeader = _recalculateColumnTotal_(
          orderSheet,
          idx.sum + 1,
          Math.max(1, firstDataRow - 1),
          '#,##0.00"€"'
        );
        var sumHeaderLabel =
          'Предварительная сумма заказа ' + _formatEuroLabel_(newTotalForHeader);
        orderSheet.getRange(HEADER_ROW_INDEX, idx.sum + 1).setValue(sumHeaderLabel);
        updatedColumns.push("«предварительная сумма заказа»");
      }

      Lib.logInfo(
        "[Order] _populateDerivedColumns: обновлены столбцы " +
          updatedColumns.join(", ") +
          " для " +
          numRows +
          " строк"
      );
    } catch (err) {
      Lib.logError("[Order] _populateDerivedColumns: ошибка", err);
    }
  }

  // Публичный хук: пересчитать производные колонки на листе заказа из других модулей
  Lib.populateOrderDerivedColumns = function () {
    try {
      _populateDerivedColumns();
    } catch (e) {
      try { Lib.logError('[Order] populateOrderDerivedColumns: ошибка', e); } catch (_) {}
    }
  };

  function _collectOrderHeaders_(sheet, lastColumn, headerRowIndex) {
    // В новой модели читаем только строку 1 - единственную строку заголовков
    var headerValues = sheet.getRange(headerRowIndex, 1, 1, lastColumn).getValues();
    var headerRow = headerValues[0]
      ? headerValues[0].map(_trimHeaderValue_)
      : [];

    // effective теперь просто копия headerRow, так как строка 1 - это источник истины
    var effective = headerRow.map(_trimHeaderValue_);

    // Для обратной совместимости оставляем структуру, но все поля указывают на одни и те же данные
    return {
      effective: effective,
      top: headerRow,       // top = строка 1
      banner: [],           // banner больше не используется
      data: [],             // data больше не используется
    };
  }

  function _trimHeaderValue_(value) {
    return String(value || "").trim();
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
    if (normalized.indexOf("потребность на") === 0) return true;
    if (normalized.indexOf("среднемесячные продажи") === 0) return true;
    if (normalized.indexOf("необходимо заказать") === 0) return true;
    if (normalized === "остаток") return true;
    if (normalized === "резерв") return true;
    if (normalized.indexOf("товар в пути") === 0) return true;
    if (normalized.indexOf("предварительный заказ") === 0) return true;
    if (normalized === "заказ") return true;
    if (normalized.indexOf("заказ в упаковках") === 0) return true;
    if (
      normalized.indexOf("шт уп") !== -1 ||
      normalized.indexOf("штуп") !== -1 ||
      normalized.replace(/\s+/g, "") === "штуп"
    )
      return true;
    if (normalized.indexOf("exw alfaspa текущая") === 0) return true;
    if (normalized.indexOf("предварительная сумма заказа") === 0) return true;
    if (normalized.indexOf("кол-во месяцев") === 0) return true;
    return false;
  }

  function _rowContainsHeaderCandidate_(row) {
    if (!row || !row.length) return false;
    for (var i = 0; i < row.length; i++) {
      if (_isHeaderCandidate_(row[i])) return true;
    }
    return false;
  }

  // --- Вспомогательные функции ---
  function _getHeaderCandidates_(headersInfo, index) {
    var candidates = [];
    if (!headersInfo || !headersInfo.effective) {
      return candidates;
    }
    // В новой модели используем только effective (строка 1 - единственный источник истины)
    var effective = headersInfo.effective;
    if (index < effective.length) {
      var candidate = _trimHeaderValue_(effective[index]);
      if (candidate) {
        candidates.push(candidate);
      }
    }
    return candidates;
  }

  function _resolveHeaderLabel_(headersInfo, index) {
    if (!headersInfo || index === undefined || index === null || index < 0) {
      return "";
    }
    var candidates = _getHeaderCandidates_(headersInfo, index);
    return candidates.length ? candidates[0] : "";
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
    if (value === null || value === undefined) return "";
    return String(value).trim().toLowerCase();
  }

  function _parseNumber(value) {
    if (value === null || value === undefined) return null;
    if (typeof value === "number") return isNaN(value) ? null : value;
    var str = String(value).trim().replace(/\s+/g, "").replace(",", ".");
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
    if (value === null || value === undefined || isNaN(value)) return "";
    return value;
  }

  function _fmtThousands_(value) {
    var num = Number(value || 0);
    var sign = num < 0 ? "-" : "";
    var abs = Math.abs(num);
    var formatted = abs.toFixed(2);
    var parts = formatted.split(".");
    parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, " ");
    return sign + parts[0] + "," + parts[1];
  }

  function _formatEuroLabel_(value) {
    return _fmtThousands_(value) + " €";
  }

  function _setColumnTotal_(sheet, columnIndex, total, numberFormat) {
    if (!sheet || !columnIndex) return;
    try {
      Lib.logDebug(
        "[Order] _setColumnTotal_: итоги в строке 2 отключены (col=" +
          columnIndex +
          ")"
      );
    } catch (_) {}
    return; // по новым правилам ничего не пишем в строку 2
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
      _setColumnTotal_(sheet, columnIndex, 0, numberFormat);
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
    var result = hasValue ? total : 0;

    // Проверка: если колонка = «Предварительная сумма заказа» → не трогаем строку 2
    try {
      var headerLabel = String(sheet.getRange(1, columnIndex).getValue() || "")
        .trim()
        .toLowerCase();
      if (headerLabel.indexOf("предварительная сумма заказа") !== 0) {
        _setColumnTotal_(sheet, columnIndex, result, numberFormat);
      } else {
        Lib.logDebug("[Order] _recalculateColumnTotal_: для суммы пишем только в шапку, строка 2 пропущена");
      }
    } catch (_) {}

    return result;
  }

  function _recalculateOrderRow_(sheet, row, options) {
    if (!sheet || !row) return;
    options = options || {};
    // Динамически определяем первую строку данных и используем её - 1 как индекс строки «шапки»
    // для корректного пересчёта итогов колонок вне зависимости от компоновки листа.
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
    var trigger = options.trigger || "";
    var forceFromPreOrder =
      trigger === "preOrder" ||
      (trigger === "unitsPerPack" && preOrderValue !== null);
    Lib.logDebug(
      "[Order] recalc row " +
        row +
        ": trigger=" +
        trigger +
        ", preOrder=" +
        (preOrderValue !== null ? preOrderValue : "∅") +
        ", unitsPerPack=" +
        (unitsPerPackValue !== null ? unitsPerPackValue : "∅") +
        ", packs(cell)=" +
        (packCellValue !== null ? packCellValue : "∅") +
        ", order(cell)=" +
        (orderCellValue !== null ? orderCellValue : "∅") +
        ", price=" +
        (priceValue !== null ? priceValue : "∅")
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
        "[Order] recalc row " +
          row +
          ": computedPack=" +
          computedPack +
          " (forceFromPreOrder=" +
          forceFromPreOrder +
          ")"
      );
      sheet
        .getRange(row, idx.packs + 1)
        .setValue(packOutput === "" ? "" : packOutput)
        .setFontFamily("Century Gothic")
        .setFontSize(9)
        .setNumberFormat("0")
        .setBorder(false, false, false, false, false, false);
      Lib.logDebug(
        "[Order] recalc row " +
          row +
          ": pack updated -> " +
          (packOutput === "" ? "∅" : packOutput)
      );
    }
    var effectiveOrder = orderCellValue;
    if (idx.order !== -1) {
      if (trigger === "order") {
        effectiveOrder = orderCellValue;
      } else {
        var calculatedOrder = _calculateOrderUnits_(
          computedPack,
          unitsPerPackValue
        );
        Lib.logDebug(
          "[Order] recalc row " +
            row +
            ": calculatedOrder=" +
            calculatedOrder +
            " from computedPack=" +
            computedPack +
            ", unitsPerPack=" +
            unitsPerPackValue
        );
        if (calculatedOrder !== null && !isNaN(calculatedOrder))
          effectiveOrder = calculatedOrder;
      }
      var orderOutput = _formatNumberOrBlank_(effectiveOrder);
      sheet
        .getRange(row, idx.order + 1)
        .setValue(orderOutput === "" ? "" : orderOutput)
        .setFontFamily("Century Gothic")
        .setFontSize(9)
        .setNumberFormat("0")
        .setBorder(false, false, false, false, false, false);
      Lib.logDebug(
        "[Order] recalc row " +
          row +
          ": order updated -> " +
          (orderOutput === "" ? "∅" : orderOutput)
      );
    }
    if (idx.sum !== -1) {
      var sumValue = _calculatePreliminarySum_(effectiveOrder, priceValue);
      Lib.logDebug(
        "[Order] recalc row " +
          row +
          ": sumValue=" +
          sumValue +
          " (order=" +
          effectiveOrder +
          ", price=" +
          priceValue +
          ")"
      );
      var sumOutput = _formatNumberOrBlank_(sumValue);
      sheet
        .getRange(row, idx.sum + 1)
        .setValue(sumOutput === "" ? "" : sumValue)
        .setFontFamily("Century Gothic")
        .setFontSize(9)
        .setNumberFormat('#,##0.00"€"')
        .setBorder(false, false, false, false, false, false);
      Lib.logDebug(
        "[Order] recalc row " +
          row +
          ": sum updated -> " +
          (sumOutput === "" ? "∅" : sumValue)
      );
    }
    // Итоги в строке 2 отключены для всех колонок — считаем только сумму для подписи в строке 1
    if (idx.sum !== -1) {
      var newTotal = _recalculateColumnTotal_(
        sheet,
        idx.sum + 1,
        headerRowIndex,
        '#,##0.00"€"'
      );
      try {
        // Обновим подпись в строке 1, чтобы отображалась актуальная сумма
        var sumHeaderLabel =
          'Предварительная сумма заказа ' + _formatEuroLabel_(newTotal);
        sheet.getRange(HEADER_ROW_INDEX, idx.sum + 1).setValue(sumHeaderLabel);
      } catch (_) {}
    }
  }

  // ========== СТАДИИ ПО ЗАКАЗ: УНИВЕРСАЛЬНЫЕ ФУНКЦИИ ФИЛЬТРАЦИИ ==========

  /**
   * Вспомогательная функция для получения индексов колонок по названиям
   * Поддерживает точное совпадение и частичное совпадение (начало строки)
   */
  function _getColumnIndexes(sheet, columnNames) {
    var maxColumn = sheet.getMaxColumns();
    if (maxColumn <= 0) return {};

    // Используем getMaxColumns() чтобы получить все колонки, включая скрытые
    var headers = sheet.getRange(1, 1, 1, maxColumn).getValues()[0];
    var indexes = {};

    for (var i = 0; i < columnNames.length; i++) {
      var colName = columnNames[i];
      var index = headers.indexOf(colName);

      // Если не найдено точное совпадение, пытаемся найти частичное (начало строки)
      if (index === -1) {
        for (var j = 0; j < headers.length; j++) {
          var header = String(headers[j] || "");
          if (header.indexOf(colName) === 0) {
            index = j;
            break;
          }
        }
      }

      if (index !== -1) {
        indexes[colName] = index + 1; // +1 для номера колонки (1-based)
      }
    }

    return indexes;
  }

  /**
   * Вспомогательная функция для скрытия колонок по именам
   */
  function _hideColumnsByNames(sheet, columnNames) {
    var indexes = _getColumnIndexes(sheet, columnNames);
    var hiddenCount = 0;
    var notFoundColumns = [];

    for (var i = 0; i < columnNames.length; i++) {
      var colName = columnNames[i];
      if (indexes[colName]) {
        try {
          sheet.hideColumns(indexes[colName]);
          hiddenCount++;
          Lib.logDebug("[OrderStages] Скрыта колонка: " + colName + " (индекс " + indexes[colName] + ")");
        } catch (e) {
          Lib.logWarn("[OrderStages] Не удалось скрыть колонку: " + colName + ", ошибка: " + e);
        }
      } else {
        notFoundColumns.push(colName);
      }
    }

    if (notFoundColumns.length > 0) {
      Lib.logWarn("[OrderStages] Не найдены колонки: " + notFoundColumns.join(", "));
    }

    Lib.logInfo("[OrderStages] Скрыто колонок: " + hiddenCount + " из " + columnNames.length);
  }

  /**
   * Вспомогательная функция для фильтрации строк по условиям
   */
  function _filterRowsByStatus(sheet, filterConfig) {
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return;

    var lastColumn = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];

    // Получаем индексы нужных колонок
    var statusIndex = headers.indexOf("Статус");
    var remainderIndex = headers.indexOf("Остаток");
    var salesIndex = headers.indexOf("ПРОДАЖИ");
    var categoryIndex = headers.indexOf("Категория товара");

    if (statusIndex === -1) {
      Lib.logWarn("[OrderStages] Не найдена колонка 'Статус'");
      return;
    }

    // Логируем настройки фильтрации для отладки
    if (filterConfig.hideCategories && Array.isArray(filterConfig.hideCategories)) {
      Lib.logInfo("[OrderStages] Фильтрация по категориям: " + filterConfig.hideCategories.join(", "));
      Lib.logInfo("[OrderStages] Индекс колонки 'Категория товара': " + categoryIndex);
    }

    // Получаем все данные
    var data = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();

    // Определяем строки для скрытия
    var rowsToHide = [];

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var status = String(row[statusIndex] || "").trim();
      var remainder = row[remainderIndex];
      var sales = row[salesIndex];
      var category = categoryIndex !== -1 ? String(row[categoryIndex] || "").trim() : "";

      var shouldHide = false;

      // Проверяем условие "Снят архив"
      if (filterConfig.hideArchived && status === "Снят архив") {
        var remainderEmpty = remainder === "" || remainder === null || remainder === undefined || remainder === 0;
        var salesEmpty = sales === "" || sales === null || sales === undefined || sales === 0;

        if (remainderEmpty && salesEmpty) {
          shouldHide = true;
        }
      }

      // Проверяем условие "Не заказываем"
      if (filterConfig.hideNotOrdered && status === "Не заказываем") {
        shouldHide = true;
      }

      // Проверяем условие скрытия по категориям
      if (filterConfig.hideCategories && Array.isArray(filterConfig.hideCategories) && category) {
        for (var k = 0; k < filterConfig.hideCategories.length; k++) {
          if (category.toLowerCase() === filterConfig.hideCategories[k].toLowerCase()) {
            shouldHide = true;
            Lib.logDebug("[OrderStages] Строка " + (i + 2) + " будет скрыта. Категория: '" + category + "'");
            break;
          }
        }
      }

      if (shouldHide) {
        rowsToHide.push(i + 2); // +2 потому что строки начинаются с 2 (1 - заголовок)
      }
    }

    // Скрываем строки
    if (rowsToHide.length > 0) {
      // Группируем последовательные строки для более эффективного скрытия
      var start = rowsToHide[0];
      var count = 1;

      for (var j = 1; j <= rowsToHide.length; j++) {
        if (j < rowsToHide.length && rowsToHide[j] === rowsToHide[j - 1] + 1) {
          count++;
        } else {
          try {
            sheet.hideRows(start, count);
          } catch (e) {
            Lib.logWarn("[OrderStages] Не удалось скрыть строки " + start + "-" + (start + count - 1) + ": " + e);
          }
          if (j < rowsToHide.length) {
            start = rowsToHide[j];
            count = 1;
          }
        }
      }

      Lib.logInfo("[OrderStages] Скрыто строк: " + rowsToHide.length);
    }
  }

  /**
   * 1. Все данные - убирает все фильтры и показывает лист полностью
   */
  Lib.showAllOrderData = function() {
    var ui = SpreadsheetApp.getUi();
    var menuTitle = "Стадии по заказ";

    try {
      Lib.logInfo("[OrderStages] Показать все данные: старт");

      var ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) {
        throw new Error("Не удалось получить активную таблицу");
      }

      var orderSheetName = Lib.CONFIG && Lib.CONFIG.SHEETS && Lib.CONFIG.SHEETS.ORDER_FORM;
      if (!orderSheetName) {
        throw new Error("Не найдена конфигурация листа Заказ");
      }

      var sheet = ss.getSheetByName(orderSheetName);
      if (!sheet) {
        throw new Error("Не найден лист: " + orderSheetName);
      }

      // Показываем все колонки
      var lastColumn = sheet.getMaxColumns();
      if (lastColumn > 0) {
        sheet.showColumns(1, lastColumn);
      }

      // Показываем все строки
      var lastRow = sheet.getMaxRows();
      if (lastRow > 1) {
        sheet.showRows(2, lastRow - 1);
      }

      Lib.logInfo("[OrderStages] Показать все данные: завершено");
      ss.toast("Все фильтры сняты", menuTitle, 3);
    } catch (error) {
      Lib.logError("[OrderStages] showAllOrderData: ошибка", error);
      ui.alert("Ошибка", error.message || String(error), ui.ButtonSet.OK);
    }
  };

  /**
   * 2. Заказ - показывает данные для заказа
   */
  Lib.showOrderStage = function() {
    var ui = SpreadsheetApp.getUi();
    var menuTitle = "Стадии по заказ";

    try {
      Lib.logInfo("[OrderStages] Показать стадию 'Заказ': старт");

      var ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) {
        throw new Error("Не удалось получить активную таблицу");
      }

      var orderSheetName = Lib.CONFIG && Lib.CONFIG.SHEETS && Lib.CONFIG.SHEETS.ORDER_FORM;
      if (!orderSheetName) {
        throw new Error("Не найдена конфигурация листа Заказ");
      }

      var sheet = ss.getSheetByName(orderSheetName);
      if (!sheet) {
        throw new Error("Не найден лист: " + orderSheetName);
      }

      // Сначала показываем все колонки и строки
      var lastColumn = sheet.getMaxColumns();
      var lastRow = sheet.getMaxRows();
      if (lastColumn > 0) sheet.showColumns(1, lastColumn);
      if (lastRow > 1) sheet.showRows(2, lastRow - 1);

      // Скрываем указанные колонки
      var columnsToHide = [
        "АКЦИИ",
        "условие#",
        "коментарий#",
        "Срок#",
        "Набор",
        "условие",
        "коментарий",
        "Срок",
        "Добавить в прайс",
        "Категория товара",
        "Группа линии",
        "Линия Прайс",
        "Примечание",
        "Продублировать"
      ];
      _hideColumnsByNames(sheet, columnsToHide);

      // Фильтруем строки: скрываем "Снят архив" с пустыми Остаток и ПРОДАЖИ
      // Не скрываем строки по категориям (показываем "Пробники", "Mini Size", "Тестер")
      _filterRowsByStatus(sheet, {
        hideArchived: true,
        hideNotOrdered: false
      });

      Lib.logInfo("[OrderStages] Показать стадию 'Заказ': завершено");
      ss.toast("Применен фильтр 'Заказ'", menuTitle, 3);
    } catch (error) {
      Lib.logError("[OrderStages] showOrderStage: ошибка", error);
      ui.alert("Ошибка", error.message || String(error), ui.ButtonSet.OK);
    }
  };

  /**
   * 3. Акции - показывает данные для акций
   */
  Lib.showPromotionsStage = function() {
    var ui = SpreadsheetApp.getUi();
    var menuTitle = "Стадии по заказ";

    try {
      Lib.logInfo("[OrderStages] Показать стадию 'Акции': старт");

      var ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) {
        throw new Error("Не удалось получить активную таблицу");
      }

      var orderSheetName = Lib.CONFIG && Lib.CONFIG.SHEETS && Lib.CONFIG.SHEETS.ORDER_FORM;
      if (!orderSheetName) {
        throw new Error("Не найдена конфигурация листа Заказ");
      }

      var sheet = ss.getSheetByName(orderSheetName);
      if (!sheet) {
        throw new Error("Не найден лист: " + orderSheetName);
      }

      // Сначала показываем все колонки и строки
      var lastColumn = sheet.getMaxColumns();
      var lastRow = sheet.getMaxRows();
      if (lastColumn > 0) sheet.showColumns(1, lastColumn);
      if (lastRow > 1) sheet.showRows(2, lastRow - 1);

      // Скрываем указанные колонки
      var columnsToHide = [
        "Необходимо заказать, шт",
        "Предварительный заказ",
        "заказ в упаковках",
        "ЗАКАЗ",
        "EXW  ALFASPA  текущая, €",
        "Предварительная сумма заказа",
        "Набор",
        "условие",
        "коментарий",
        "Срок",
        "Добавить в прайс",
        "Категория товара",
        "Группа линии",
        "Линия Прайс",
        "Примечание",
        "Продублировать"
      ];
      _hideColumnsByNames(sheet, columnsToHide);

      // Фильтруем строки: скрываем "Снят архив" и "Не заказываем"
      // Не скрываем строки по категориям (показываем "Пробники", "Mini Size", "Тестер")
      _filterRowsByStatus(sheet, {
        hideArchived: true,
        hideNotOrdered: true
      });

      Lib.logInfo("[OrderStages] Показать стадию 'Акции': завершено");
      ss.toast("Применен фильтр 'Акции'", menuTitle, 3);
    } catch (error) {
      Lib.logError("[OrderStages] showPromotionsStage: ошибка", error);
      ui.alert("Ошибка", error.message || String(error), ui.ButtonSet.OK);
    }
  };

  /**
   * 4. Набор - показывает данные для набора
   */
  Lib.showSetStage = function() {
    var ui = SpreadsheetApp.getUi();
    var menuTitle = "Стадии по заказ";

    try {
      Lib.logInfo("[OrderStages] Показать стадию 'Набор': старт");

      var ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) {
        throw new Error("Не удалось получить активную таблицу");
      }

      var orderSheetName = Lib.CONFIG && Lib.CONFIG.SHEETS && Lib.CONFIG.SHEETS.ORDER_FORM;
      if (!orderSheetName) {
        throw new Error("Не найдена конфигурация листа Заказ");
      }

      var sheet = ss.getSheetByName(orderSheetName);
      if (!sheet) {
        throw new Error("Не найден лист: " + orderSheetName);
      }

      // Сначала показываем все колонки и строки
      var lastColumn = sheet.getMaxColumns();
      var lastRow = sheet.getMaxRows();
      if (lastColumn > 0) sheet.showColumns(1, lastColumn);
      if (lastRow > 1) sheet.showRows(2, lastRow - 1);

      // Скрываем указанные колонки
      var columnsToHide = [
        "Необходимо заказать, шт",
        "Предварительный заказ",
        "заказ в упаковках",
        "ЗАКАЗ",
        "EXW  ALFASPA  текущая, €",
        "Предварительная сумма заказа",
        "АКЦИИ",
        "условие#",
        "коментарий#",
        "Срок#",
        "Добавить в прайс",
        "Категория товара",
        "Группа линии",
        "Линия Прайс",
        "Примечание",
        "Продублировать"
      ];
      _hideColumnsByNames(sheet, columnsToHide);

      // Фильтруем строки: скрываем "Снят архив" и "Не заказываем"
      // Не скрываем строки по категориям (показываем "Пробники", "Mini Size", "Тестер")
      _filterRowsByStatus(sheet, {
        hideArchived: true,
        hideNotOrdered: true
      });

      Lib.logInfo("[OrderStages] Показать стадию 'Набор': завершено");
      ss.toast("Применен фильтр 'Набор'", menuTitle, 3);
    } catch (error) {
      Lib.logError("[OrderStages] showSetStage: ошибка", error);
      ui.alert("Ошибка", error.message || String(error), ui.ButtonSet.OK);
    }
  };

  /**
   * 5. Прайс - показывает данные для прайса
   */
  Lib.showPriceStage = function() {
    var ui = SpreadsheetApp.getUi();
    var menuTitle = "Стадии по заказ";

    try {
      Lib.logInfo("[OrderStages] Показать стадию 'Прайс': старт");

      var ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) {
        throw new Error("Не удалось получить активную таблицу");
      }

      var orderSheetName = Lib.CONFIG && Lib.CONFIG.SHEETS && Lib.CONFIG.SHEETS.ORDER_FORM;
      if (!orderSheetName) {
        throw new Error("Не найдена конфигурация листа Заказ");
      }

      var sheet = ss.getSheetByName(orderSheetName);
      if (!sheet) {
        throw new Error("Не найден лист: " + orderSheetName);
      }

      // Сначала показываем все колонки и строки
      var lastColumn = sheet.getMaxColumns();
      var lastRow = sheet.getMaxRows();
      if (lastColumn > 0) sheet.showColumns(1, lastColumn);
      if (lastRow > 1) sheet.showRows(2, lastRow - 1);

      // Скрываем указанные колонки
      var columnsToHide = [
        "Остаток 1",
        "СГ 1",
        "Остаток 2",
        "СГ 2",
        "Остаток3",
        "СГ 3",
        "СПИСАНО",
        "РЕЗЕРВ",
        "Среднемесячные продажи, шт",
        "Потребность на 6 месяцев, шт",
        "На сколько мес запас",
        "Необходимо заказать, шт",
        "Предварительный заказ",
        "заказ в упаковках",
        "ЗАКАЗ",
        "EXW  ALFASPA  текущая, €",
        "Предварительная сумма заказа",
        "АКЦИИ",
        "условие#",
        "коментарий#",
        "Срок#",
        "Набор",
        "условие",
        "коментарий",
        "Срок"
      ];
      _hideColumnsByNames(sheet, columnsToHide);

      // Фильтруем строки: скрываем "Снят архив" и "Не заказываем"
      // А также скрываем строки с категориями "Пробники", "Mini Size", "Тестер"
      _filterRowsByStatus(sheet, {
        hideArchived: true,
        hideNotOrdered: true,
        hideCategories: ["Пробники", "Mini Size", "Тестер"]
      });

      Lib.logInfo("[OrderStages] Показать стадию 'Прайс': завершено");
      ss.toast("Применен фильтр 'Прайс'", menuTitle, 3);
    } catch (error) {
      Lib.logError("[OrderStages] showPriceStage: ошибка", error);
      ui.alert("Ошибка", error.message || String(error), ui.ButtonSet.OK);
    }
  };

  /**
   * ОТЛАДОЧНАЯ ФУНКЦИЯ: Показать все заголовки колонок
   */
  Lib.debugShowAllColumnHeaders = function() {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var orderSheetName = Lib.CONFIG && Lib.CONFIG.SHEETS && Lib.CONFIG.SHEETS.ORDER_FORM;
      var sheet = ss.getSheetByName(orderSheetName);

      if (!sheet) {
        SpreadsheetApp.getUi().alert("Лист Заказ не найден: " + orderSheetName);
        return;
      }

      var maxColumn = sheet.getMaxColumns();
      var headers = sheet.getRange(1, 1, 1, maxColumn).getValues()[0];

      var headersList = [];
      for (var i = 0; i < headers.length; i++) {
        if (headers[i] && String(headers[i]).trim() !== "") {
          headersList.push((i + 1) + ". \"" + headers[i] + "\"");
        }
      }

      Lib.logInfo("[DEBUG] Все заголовки колонок:\n" + headersList.join("\n"));
      SpreadsheetApp.getUi().alert("Заголовки колонок", "Всего колонок: " + headersList.length + "\n\nПроверьте логи для полного списка.", SpreadsheetApp.getUi().ButtonSet.OK);
    } catch (error) {
      Lib.logError("[DEBUG] debugShowAllColumnHeaders: ошибка", error);
      SpreadsheetApp.getUi().alert("Ошибка", error.message || String(error), SpreadsheetApp.getUi().ButtonSet.OK);
    }
  };
})(Lib, this);
