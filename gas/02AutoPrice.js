var Lib = Lib || {};

/**
 * =======================================================================================
 * МОДУЛЬ: АВТОМАТИЧЕСКОЕ ДОБАВЛЕНИЕ ГОДОВОГО БЛОКА НА ЛИСТ "Динамика цены"
 * ---------------------------------------------------------------------------------------
 * Назначение: по команде из меню "🧾 Заказ" вставляет новый блок столбцов для текущего года
 * сразу после столбца "Комментарий" на листе "Динамика цены".
 * =======================================================================================
 */

(function (Lib, global) {
  var HEADER_ROW_INDEX = 1;
  var MODULE_TAG = "[AutoPrice]";
  var _cachedLocale = null;

  var BLOCK_COLORS = [
    "#cc0000", // Темно-красный
    "#e06666", // Красный
    "#f6b26b", // Оранжевый
    "#ffd966", // Желтый
    "#93c47d", // Светло-зеленый
    "#76a5af", // Бирюзовый
    "#6fa8dc", // Голубой
    "#6d9eeb", // Синий
    "#8e7cc3", // Фиолетовый
    "#c27ba0"  // Розовый
  ];

  /**
   * Публичная функция: добавить блок столбцов для текущего года на лист "Динамика цены".
   * Вызывается из пользовательского меню (кнопка "New год для динамика").
   * ВАЖНО: Сначала сохраняет данные прошлого периода на листе "Расчет цены",
   * затем добавляет новый блок на лист "Динамика цены".
   */
  Lib.addNewYearColumnsToPriceDynamics = function () {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) {
        throw new Error("Не удалось получить активную таблицу");
      }

      // ШАГ 1: Сначала сохраняем данные прошлого периода на листе "Расчет цены"
      if (Lib.logDebug) {
        Lib.logDebug(MODULE_TAG + " Шаг 1: Сохранение данных прошлого периода на листе \"Расчет цены\"");
      }

      try {
        if (Lib.createNewYearForDynamics) {
          Lib.createNewYearForDynamics();
        } else {
          if (Lib.logWarn) {
            Lib.logWarn(MODULE_TAG + " Функция createNewYearForDynamics не найдена");
          }
        }
      } catch (saveError) {
        if (Lib.logError) {
          Lib.logError(MODULE_TAG + " Ошибка при сохранении данных прошлого периода", saveError);
        }
        // Продолжаем выполнение, даже если сохранение не удалось
      }

      // ШАГ 2: Добавляем новый блок на лист "Динамика цены"
      if (Lib.logDebug) {
        Lib.logDebug(MODULE_TAG + " Шаг 2: Добавление нового блока на лист \"Динамика цены\"");
      }

      var sheetName =
        Lib.CONFIG &&
        Lib.CONFIG.SHEETS &&
        Lib.CONFIG.SHEETS.PRICE_DYNAMICS
          ? Lib.CONFIG.SHEETS.PRICE_DYNAMICS
          : "Динамика цены";
      var sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        _showAlert(
          "Лист \"" + sheetName + "\" не найден",
          "Проверьте, что лист существует в таблице."
        );
        return;
      }

      var year = new Date().getFullYear();
      var templates = _buildColumnTemplates(year);

      var lastColumn = sheet.getLastColumn();
      if (lastColumn < 1) {
        _showAlert(
          "Лист \"" + sheetName + "\" пуст",
          "Добавление столбцов невозможно без строки заголовков."
        );
        return;
      }

      var headerValues = sheet
        .getRange(HEADER_ROW_INDEX, 1, 1, lastColumn)
        .getValues()[0]
        .map(function (value) {
          // Удаляем все пробелы в начале и конце, включая неразрывные пробелы
          return String(value || "").replace(/^[\s\u00A0]+|[\s\u00A0]+$/g, "");
        });

      var commentIndex = headerValues.indexOf("Комментарий");
      if (commentIndex === -1) {
        _showAlert(
          "Столбец \"Комментарий\" не найден",
          "Убедитесь, что на листе \"Динамика цены\" присутствует столбец \"Комментарий\"."
        );
        return;
      }

      var primaryHeader = templates[0].header;
      var blockAdded = false;
      if (headerValues.indexOf(primaryHeader) !== -1) {
        _showAlert(
          "Блок за " + year + " год уже существует",
          "Столбцы с заголовком \"" +
            primaryHeader +
            "\" уже присутствуют. Добавление нового блока не требуется."
        );
      } else {
        var insertPosition = commentIndex + 1; // индекс столбца (0-based)
        sheet.insertColumnsAfter(insertPosition, templates.length);

        var headerRange = sheet.getRange(
          HEADER_ROW_INDEX,
          insertPosition + 1,
          1,
          templates.length
        );
        headerRange.setValues([templates.map(function (tpl) { return tpl.header; })]);

        // Копируем формат заголовка с исходного столбца "Комментарий"
        var commentHeaderRange = sheet.getRange(HEADER_ROW_INDEX, commentIndex + 1);
        commentHeaderRange.copyFormatToRange(
          sheet,
          insertPosition + 1,
          insertPosition + templates.length,
          HEADER_ROW_INDEX,
          HEADER_ROW_INDEX
        );

        // Применяем цвет к заголовкам блока (новый блок всегда первый - индекс 0)
        var blockColor = BLOCK_COLORS[0];
        headerRange.setBackground(blockColor);

        var dataRowCount = sheet.getMaxRows() - HEADER_ROW_INDEX;
        if (dataRowCount > 0) {
          templates.forEach(function (tpl, idx) {
            if (tpl.numberFormat) {
              var tplRange = sheet.getRange(
                HEADER_ROW_INDEX + 1,
                insertPosition + 1 + idx,
                dataRowCount,
                1
              );
              _safeApplyNumberFormat(tplRange, {
                format: tpl.numberFormat,
                columnName: tpl.header
              });
            }
          });
        }

        SpreadsheetApp.flush();

        if (Lib.logInfo) {
          Lib.logInfo(
            MODULE_TAG +
              " Добавлен блок столбцов за " +
              year +
              " год на лист \"" +
              sheetName +
              '"'
          );
        }

        ss.toast(
          "Добавлен блок столбцов за " + year + " год на лист \"" + sheetName + "\"",
          "Авто цены",
          5
        );
        blockAdded = true;
      }

      _applyPriceDynamicsFormulas(sheet);

      // Обновляем вертикальные границы на листе "Динамика цены"
      if (typeof Lib.updatePriceDynamicsBorders === 'function') {
        Lib.updatePriceDynamicsBorders();
      }

      // ШАГ 3: Обновляем формулы на листе "Расчет цены"
      if (Lib.logDebug) {
        Lib.logDebug(MODULE_TAG + " Шаг 3: Обновление формул на листе \"Расчет цены\"");
      }

      try {
        var priceCalcSheetName =
          Lib.CONFIG &&
          Lib.CONFIG.SHEETS &&
          Lib.CONFIG.SHEETS.PRICE_CALCULATION
            ? Lib.CONFIG.SHEETS.PRICE_CALCULATION
            : "Расчет цены";
        var priceCalcSheet = ss.getSheetByName(priceCalcSheetName);
        if (priceCalcSheet) {
          _applyPriceCalculationFormulas(priceCalcSheet, sheet);
          _applyAllCalculationFormulas(priceCalcSheet, ss);

          // Обновляем вертикальные границы на листе "Расчет цены"
          if (typeof Lib.updatePriceCalculationBorders === 'function') {
            Lib.updatePriceCalculationBorders();
          }
        }
      } catch (calcError) {
        if (Lib.logWarn) {
          Lib.logWarn(MODULE_TAG + " Не удалось обновить лист \"Расчет цены\"", calcError);
        }
      }

      // ШАГ 4: Очищаем столбцы EXW/Закупочная/DDP на листе "Расчет цены"
      // Это делается ПОСЛЕ обновления формул, чтобы формулы не сломались
      if (Lib.logDebug) {
        Lib.logDebug(MODULE_TAG + " Шаг 4: Очистка столбцов EXW/Закупочная/DDP");
      }

      try {
        _clearCalculatedColumnsOnPriceCalc();
      } catch (clearError) {
        if (Lib.logWarn) {
          Lib.logWarn(MODULE_TAG + " Не удалось очистить столбцы на листе \"Расчет цены\"", clearError);
        }
      }

      if (!blockAdded) {
        ss.toast(
          "Формулы \"Динамика цены\" обновлены",
          "Авто цены",
          5
        );
      } else {
        ss.toast(
          "Добавлен блок столбцов за " + year + " год, обновлены формулы и данные прошлого периода",
          "Авто цены",
          5
        );
      }
    } catch (error) {
      if (Lib.logError) {
        Lib.logError(MODULE_TAG + " Ошибка добавления блока столбцов", error);
      } else {
        console.error(MODULE_TAG + " Ошибка добавления блока столбцов", error);
      }
      _showAlert("Ошибка", error.message || String(error));
    }
  };

  /**
   * Построить шаблон столбцов для указанного года.
   * @param {number} year - целевой год
   * @returns {Array<{header: string, numberFormat: string|null}>}
   */
  function _buildColumnTemplates(year) {
    var prevYear = year - 1;
    var prevYearShort = String(prevYear).slice(-2);
    var currentYearShort = String(year).slice(-2);
    var growthSuffix = prevYearShort + "-" + currentYearShort;

    return [
      { header: "EXW " + year + ", €", numberFormat: "0.00" },
      { header: "СКИДКА ОТ EXW " + year + ", %", numberFormat: "0%" },
      { header: "EXW ALFASPA " + year + ", €", numberFormat: "0.00" },
      { header: "Закупочная цена " + year + ", ₽", numberFormat: "#,##0" },
      { header: "DDP-МОСКВА " + year + ", ₽", numberFormat: "#,##0" },
      { header: "Прирост EXW " + growthSuffix + ", %", numberFormat: "0.0%" },
      {
        header: "Прирост DDP-МОСКВА " + growthSuffix + ", %",
        numberFormat: "0.0%",
      },
    ];
  }

  Lib.recalculatePriceDynamicsFormulas = function () {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) {
        throw new Error("Не удалось получить активную таблицу");
      }

      var sheetName =
        Lib.CONFIG &&
        Lib.CONFIG.SHEETS &&
        Lib.CONFIG.SHEETS.PRICE_DYNAMICS
          ? Lib.CONFIG.SHEETS.PRICE_DYNAMICS
          : "Динамика цены";
      var sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        _showAlert(
          "Лист \"" + sheetName + "\" не найден",
          "Проверьте, что лист существует в таблице."
        );
        return;
      }

      _applyPriceDynamicsFormulas(sheet);
      if (typeof Lib.logStep === 'function') {
        Lib.logStep("Startup", "Загрузили/обновили формулы на листе \"Динамика цены\"");
      }

      // Обновляем вертикальные границы на листе "Динамика цены"
      if (typeof Lib.updatePriceDynamicsBorders === 'function') {
        Lib.updatePriceDynamicsBorders();
      }

      // Обновляем формулы на листе "Расчет цены"
      try {
        var priceCalcSheetName =
          Lib.CONFIG &&
          Lib.CONFIG.SHEETS &&
          Lib.CONFIG.SHEETS.PRICE_CALCULATION
            ? Lib.CONFIG.SHEETS.PRICE_CALCULATION
            : "Расчет цены";
        if (typeof Lib.logStep === 'function') {
          Lib.logStep("Startup", "Обновлены формулы на листе \"Расчет цены\"");
        }
        var priceCalcSheet = ss.getSheetByName(priceCalcSheetName);
        if (priceCalcSheet) {
          _applyPriceCalculationFormulas(priceCalcSheet, sheet);
          _applyAllCalculationFormulas(priceCalcSheet, ss);

          // Обновляем вертикальные границы на листе "Расчет цены"
          if (typeof Lib.updatePriceCalculationBorders === 'function') {
            Lib.updatePriceCalculationBorders();
          }
        }
      } catch (calcError) {
        if (Lib.logWarn) {
          Lib.logWarn(MODULE_TAG + " Не удалось обновить лист \"Расчет цены\"", calcError);
        }
      }

      ss.toast("Формулы \"Динамика цены\" обновлены", "Авто цены", 5);
    } catch (error) {
      if (Lib.logError) {
        Lib.logError(MODULE_TAG + " Ошибка пересчёта формул", error);
      } else {
        console.error(MODULE_TAG + " Ошибка пересчёта формул", error);
      }
      _showAlert("Ошибка", error.message || String(error));
    }
  };

  /**
   * Публичная функция: обновить формулы на листе "Расчет цены"
   * Подтягивает данные из листа "Динамика цены" по ID
   * @param {boolean} silent - Не показывать уведомления (для автозапуска)
   */
  Lib.updatePriceCalculationFormulas = function (silent) {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) {
        throw new Error("Не удалось получить активную таблицу");
      }

      var priceCalcSheetName =
        Lib.CONFIG &&
        Lib.CONFIG.SHEETS &&
        Lib.CONFIG.SHEETS.PRICE_CALCULATION
          ? Lib.CONFIG.SHEETS.PRICE_CALCULATION
          : "Расчет цены";
      var priceCalcSheet = ss.getSheetByName(priceCalcSheetName);
      if (typeof Lib.logStep === 'function') {
        Lib.logStep("Startup", "Обновляем формулы листа \"Расчет цены\"");
      }
      if (!priceCalcSheet) {
        if (!silent) {
          _showAlert(
            "Лист \"" + priceCalcSheetName + "\" не найден",
            "Проверьте, что лист существует в таблице."
          );
        }
        return;
      }

      var priceDynamicsSheetName =
        Lib.CONFIG &&
        Lib.CONFIG.SHEETS &&
        Lib.CONFIG.SHEETS.PRICE_DYNAMICS
          ? Lib.CONFIG.SHEETS.PRICE_DYNAMICS
          : "Динамика цены";
      var priceDynamicsSheet = ss.getSheetByName(priceDynamicsSheetName);
      if (!priceDynamicsSheet) {
        if (!silent) {
          _showAlert(
            "Лист \"" + priceDynamicsSheetName + "\" не найден",
            "Проверьте, что лист существует в таблице."
          );
        }
        return;
      }

      _applyPriceCalculationFormulas(priceCalcSheet, priceDynamicsSheet);

      // Обновляем вертикальные границы после применения формул
      if (typeof Lib.updatePriceCalculationBorders === 'function') {
        Lib.updatePriceCalculationBorders();
      }

      if (!silent) {
        ss.toast("Формулы \"Расчет цены\" обновлены", "Авто цены", 5);
      }
    } catch (error) {
      if (Lib.logError) {
        Lib.logError(MODULE_TAG + " Ошибка обновления формул расчета цены", error);
      } else {
        console.error(MODULE_TAG + " Ошибка обновления формул расчета цены", error);
      }
      if (!silent) {
        _showAlert("Ошибка", error.message || String(error));
      }
    }
  };

  function _applyPriceDynamicsFormulas(sheet) {
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return;
    }

    var lastColumn = sheet.getLastColumn();
    if (lastColumn <= 0) {
      return;
    }

    var headers = sheet
      .getRange(HEADER_ROW_INDEX, 1, 1, lastColumn)
      .getValues()[0]
      .map(function (value) {
        // Удаляем все пробелы в начале и конце, включая неразрывные пробелы
        return String(value || "").replace(/^[\s\u00A0]+|[\s\u00A0]+$/g, "");
      });

    var yearBlocks = [];
    for (var c = 0; c < headers.length; c++) {
      var match = headers[c].match(/^EXW\s+(\d{4}),\s*€$/i);
      if (match) {
        var year = parseInt(match[1], 10);
        var exwCol = c + 1;
        if (exwCol + 6 > lastColumn) {
          continue;
        }
        yearBlocks.push({
          year: year,
          exwCol: exwCol,
          discountCol: exwCol + 1,
          alfaspaCol: exwCol + 2,
          purchaseCol: exwCol + 3,
          ddpCol: exwCol + 4,
          growthExwCol: exwCol + 5,
          growthDdpCol: exwCol + 6,
        });
      }
    }

    if (!yearBlocks.length) {
      return;
    }

    yearBlocks.sort(function (a, b) {
      return b.year - a.year;
    });

    var ss = sheet.getParent();
    var referenceSheet = ss.getSheetByName("Справочник");
    var currencyTableRange = null;
    var ddpRef = "'Справочник'!R2C4";

    if (referenceSheet) {
      var refLastRow = referenceSheet.getLastRow();
      var refLastCol = referenceSheet.getLastColumn();
      if (refLastRow > 1 && refLastCol > 0) {
        var refHeaders = referenceSheet
          .getRange(1, 1, 1, refLastCol)
          .getValues()[0]
          .map(function (value) {
            // Удаляем все пробелы в начале и конце, включая неразрывные пробелы
            return String(value || "").replace(/^[\s\u00A0]+|[\s\u00A0]+$/g, "");
          });
        var yearIdx = refHeaders.indexOf("Год");
        var currencyIdx = refHeaders.indexOf("Курс валюты, €");
        if (yearIdx !== -1 && currencyIdx !== -1) {
          var yearCol = yearIdx + 1;
          var currencyCol = currencyIdx + 1;
          currencyTableRange = {
            yearCol: yearCol,
            currencyCol: currencyCol,
            lastRow: refLastRow
          };
        } else if (Lib.logWarn) {
          Lib.logWarn(
            MODULE_TAG +
              ' Не найдены столбцы "Год" или "Курс валюты, €" на листе Справочник'
          );
        }
      }
    } else if (Lib.logWarn) {
      Lib.logWarn(MODULE_TAG + ' Лист "Справочник" не найден');
    }

    var dataRows = lastRow - 1;

    yearBlocks.forEach(function (block, index) {
      _applyFormulasForBlock(sheet, block, dataRows, currencyTableRange, ddpRef, yearBlocks[index + 1]);
    });

    _applyBlockColors(sheet, yearBlocks);
    _applyBlockBorders(sheet, yearBlocks);
    _applyPriceDynamicsFormatting(sheet, headers, lastRow);

    // Применяем границу перед столбцом "ЦЕНА EXW из Б/З, €"
    var exwFromBudgetIdx = headers.indexOf("ЦЕНА EXW из Б/З, €");
    if (exwFromBudgetIdx !== -1) {
      var totalRows = lastRow - HEADER_ROW_INDEX + 1;
      var borderRange = sheet.getRange(HEADER_ROW_INDEX, exwFromBudgetIdx + 1, totalRows, 1);
      borderRange.setBorder(
        null, // top
        true, // left - жирная граница
        null, // bottom
        null, // right
        null, // vertical
        null, // horizontal
        null, // color
        SpreadsheetApp.BorderStyle.SOLID_THICK
      );
    }

    if (Lib.logInfo) {
      Lib.logInfo(
        MODULE_TAG +
          " Формулы для листа \"" +
          sheet.getName() +
          "\" обновлены, блоков: " +
          yearBlocks.length
      );
    }
  }

  function _applyFormulasForBlock(sheet, block, dataRows, currencyTableRange, ddpRef, prevBlock) {
    if (dataRows <= 0) {
      return;
    }

    var alfaspaRange = sheet.getRange(HEADER_ROW_INDEX + 1, block.alfaspaCol, dataRows, 1);
    var alfaspaFormula =
      '=IF(LEN(R[0]C[-1])=0,' +
      ' R[0]C[-2],' +
      ' ROUND(R[0]C[-2]*(1-R[0]C[-1]/100), 1)' +
      ')';
    _setFormulasR1C1(alfaspaRange, alfaspaFormula);
    _safeApplyNumberFormat(alfaspaRange, {
      format: "0.0",
      columnName: "EXW ALFASPA"
    });

    var purchaseRange = sheet.getRange(HEADER_ROW_INDEX + 1, block.purchaseCol, dataRows, 1);
    var currencyFormula = "1";
    if (currencyTableRange) {
      currencyFormula =
        'VLOOKUP(' + block.year + ',' +
        ' INDIRECT("Справочник!R2C' + currencyTableRange.yearCol + ':R' + currencyTableRange.lastRow + 'C' + currencyTableRange.currencyCol + '", FALSE),' +
        ' ' + (currencyTableRange.currencyCol - currencyTableRange.yearCol + 1) + ',' +
        ' FALSE)';
    }
    var purchaseFormula =
      '=IF(LEN(R[0]C[-1])=0,' +
      ' "",' +
      ' ROUND(R[0]C[-1]*' +
      currencyFormula +
      ', 0)' +
      ')';
    _setFormulasR1C1(purchaseRange, purchaseFormula);
    _safeApplyNumberFormat(purchaseRange, {
      format: "#,##0",
      columnName: "Закупочная цена (блок динамики)"
    });

    var ddpRange = sheet.getRange(HEADER_ROW_INDEX + 1, block.ddpCol, dataRows, 1);
    var ddpFormula =
      '=IF(LEN(R[0]C[-1])=0,' +
      ' "",' +
      ' ROUND(R[0]C[-1]*' +
      ddpRef +
      ', 0)' +
      ')';
    _setFormulasR1C1(ddpRange, ddpFormula);
    _safeApplyNumberFormat(ddpRange, {
      format: "#,##0",
      columnName: "DDP (блок динамики)"
    });

    var growthExwRange = sheet.getRange(HEADER_ROW_INDEX + 1, block.growthExwCol, dataRows, 1);
    if (prevBlock) {
      var offsetToCurrent = block.exwCol - block.growthExwCol;
      var offsetToPrev = prevBlock.exwCol - block.growthExwCol;
      var growthExwFormula =
        '=IF(OR(LEN(R[0]C[' + offsetToCurrent + '])=0, LEN(R[0]C[' + offsetToPrev + '])=0),' +
        ' "",' +
        ' R[0]C[' + offsetToCurrent + '] / R[0]C[' + offsetToPrev + '] - 1' +
        ')';
      _setFormulasR1C1(growthExwRange, growthExwFormula);
      _safeApplyNumberFormat(growthExwRange, {
        format: "0.0%",
        columnName: "Прирост EXW"
      });
    } else {
      growthExwRange.clearContent();
    }

    var growthDdpRange = sheet.getRange(HEADER_ROW_INDEX + 1, block.growthDdpCol, dataRows, 1);
    if (prevBlock) {
      var offsetToCurrentDdp = block.ddpCol - block.growthDdpCol;
      var offsetToPrevDdp = prevBlock.ddpCol - block.growthDdpCol;
      var growthDdpFormula =
        '=IF(OR(LEN(R[0]C[' + offsetToCurrentDdp + '])=0, LEN(R[0]C[' + offsetToPrevDdp + '])=0),' +
        ' "",' +
        ' R[0]C[' + offsetToCurrentDdp + '] / R[0]C[' + offsetToPrevDdp + '] - 1' +
        ')';
      _setFormulasR1C1(growthDdpRange, growthDdpFormula);
      _safeApplyNumberFormat(growthDdpRange, {
        format: "0.0%",
        columnName: "Прирост DDP"
      });
    } else {
      growthDdpRange.clearContent();
    }
  }

  function _setFormulasR1C1(range, formula) {
    var rows = range.getNumRows();
    if (rows <= 0) {
      return;
    }
    var localizedFormula = _localizeFormulaIfNeeded(formula);
    var formulas = [];
    for (var i = 0; i < rows; i++) {
      formulas.push([localizedFormula]);
    }
    range.setFormulasR1C1(formulas);
  }

  /**
   * Показать пользователю alert-сообщение.
   * @param {string} title
   * @param {string} message
   */
  function _showAlert(title, message) {
    try {
      var ui = SpreadsheetApp.getUi();
      ui.alert(title, message, ui.ButtonSet.OK);
    } catch (err) {
      console.log(MODULE_TAG + " " + title + ": " + message);
    }
  }

  function _getSpreadsheetLocale() {
    try {
      if (_cachedLocale) {
        return _cachedLocale;
      }
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      _cachedLocale = ss ? String(ss.getSpreadsheetLocale() || "").toLowerCase() : "";
      return _cachedLocale;
    } catch (err) {
      return "";
    }
  }

  function _localizeFormulaIfNeeded(formula) {
    var locale = _getSpreadsheetLocale();
    if (!locale || formula.indexOf("=") !== 0) {
      return formula;
    }
    if (locale.indexOf("ru") === 0 || locale.indexOf("uk") === 0 || locale.indexOf("be") === 0) {
      return formula.replace(/,/g, ";");
    }
    return formula;
  }

  function _applyBlockColors(sheet, yearBlocks) {
    yearBlocks.forEach(function (block, index) {
      var blockColor = BLOCK_COLORS[index % BLOCK_COLORS.length];
      var headerRange = sheet.getRange(HEADER_ROW_INDEX, block.exwCol, 1, 7);
      headerRange.setBackground(blockColor);
    });
  }

  function _applyBlockBorders(sheet, yearBlocks) {
    var lastRow = sheet.getLastRow();
    if (lastRow < HEADER_ROW_INDEX) {
      return;
    }

    var totalRows = lastRow - HEADER_ROW_INDEX + 1;

    yearBlocks.forEach(function (block) {
      // Применяем жирную левую границу к первому столбцу блока (EXW)
      var leftBorderRange = sheet.getRange(HEADER_ROW_INDEX, block.exwCol, totalRows, 1);
      leftBorderRange.setBorder(
        null, // top
        true, // left - жирная граница
        null, // bottom
        null, // right
        null, // vertical
        null, // horizontal
        null, // color
        SpreadsheetApp.BorderStyle.SOLID_THICK
      );

      // Применяем жирную правую границу к последнему столбцу блока (Прирост DDP)
      var rightBorderRange = sheet.getRange(HEADER_ROW_INDEX, block.growthDdpCol, totalRows, 1);
      rightBorderRange.setBorder(
        null, // top
        null, // left
        null, // bottom
        true, // right - жирная граница
        null, // vertical
        null, // horizontal
        null, // color
        SpreadsheetApp.BorderStyle.SOLID_THICK
      );
    });
  }

  var _formatLockWarnings = Object.create(null);

  function _isDataTypeFormatError(error) {
    if (!error || typeof error.message !== "string") {
      return false;
    }
    return (
      error.message.indexOf("Нельзя выбрать числовой формат ячеек") !== -1 ||
      error.message.indexOf("Cannot set a number format") !== -1
    );
  }

  function _logFormatLockHint(columnName, errorMessage) {
    var hintKey = columnName || "__unknown__";
    if (_formatLockWarnings[hintKey]) {
      return;
    }
    _formatLockWarnings[hintKey] = true;

    var baseMessage =
      MODULE_TAG +
      ' Столбец "' +
      columnName +
      '" помечен в Google Sheets как столбец с типом данных ("База"), поэтому автоматическое форматирование заблокировано.';

    if (Lib.logWarn) {
      Lib.logWarn(baseMessage + (errorMessage ? " Детали: " + errorMessage : ""));
    }
    if (Lib.logInfo) {
      Lib.logInfo(
        MODULE_TAG +
          ' Формулы продолжат обновляться, но для восстановления форматирования нужно отключить для этого столбца режим "База" (Данные → Таблица → Преобразовать в диапазон / Тип данных → Автоматически).'
      );
    }
  }

  function _safeApplyNumberFormat(range, options) {
    options = options || {};
    try {
      if (options.format) {
        range.setNumberFormat(options.format);
      }
      if (options.align) {
        range.setHorizontalAlignment(options.align);
      }
      return true;
    } catch (error) {
      if (_isDataTypeFormatError(error)) {
        _logFormatLockHint(options.columnName || range.getA1Notation(), error.message);
        return false;
      }
      throw error;
    }
  }

  /**
   * Применить формулы на лист "Расчет цены"
   * Подтягивает данные из листа "Динамика цены" по ID
   */
  function _applyPriceCalculationFormulas(priceCalcSheet, priceDynamicsSheet) {
    var currentYear = new Date().getFullYear();
    var prevYear = currentYear - 1;

    // Находим столбцы на листе "Динамика цены"
    var dynLastCol = priceDynamicsSheet.getLastColumn();
    if (dynLastCol <= 0) {
      if (Lib.logWarn) {
        Lib.logWarn(MODULE_TAG + ' Лист "Динамика цены" пуст');
      }
      return;
    }

    var dynHeaders = priceDynamicsSheet
      .getRange(HEADER_ROW_INDEX, 1, 1, dynLastCol)
      .getValues()[0]
      .map(function (value) {
        // Удаляем все пробелы в начале и конце, включая неразрывные пробелы
        return String(value || "").replace(/^[\s\u00A0]+|[\s\u00A0]+$/g, "");
      });

    var dynIdCol = dynHeaders.indexOf("ID");
    var dynExwFromBudgetCol = dynHeaders.indexOf("ЦЕНА EXW из Б/З, €");
    if (dynExwFromBudgetCol === -1) {
      dynExwFromBudgetCol = dynHeaders.indexOf("ЦЕНА EXW из Б/З");
    }
    var dynExwCurrentCol = dynHeaders.indexOf("EXW " + currentYear + ", €");
    var dynExwPrevCol = dynHeaders.indexOf("EXW " + prevYear + ", €");
    var dynAlfaspaCurrentCol = dynHeaders.indexOf("EXW ALFASPA " + currentYear + ", €");
    var dynPurchaseCurrentCol = dynHeaders.indexOf("Закупочная цена " + currentYear + ", ₽");
    var dynDdpCurrentCol = dynHeaders.indexOf("DDP-МОСКВА " + currentYear + ", ₽");

    if (dynIdCol === -1) {
      if (Lib.logWarn) {
        Lib.logWarn(MODULE_TAG + ' Не найден столбец "ID" на листе "Динамика цены"');
      }
      return;
    }

    // Находим столбцы на листе "Расчет цены"
    var calcLastCol = priceCalcSheet.getLastColumn();
    if (calcLastCol <= 0) {
      if (Lib.logWarn) {
        Lib.logWarn(MODULE_TAG + ' Лист "Расчет цены" пуст');
      }
      return;
    }

    var calcHeaders = priceCalcSheet
      .getRange(HEADER_ROW_INDEX, 1, 1, calcLastCol)
      .getValues()[0]
      .map(function (value) {
        // Удаляем все пробелы в начале и конце, включая неразрывные пробелы
        return String(value || "").replace(/^[\s\u00A0]+|[\s\u00A0]+$/g, "");
      });

    var calcIdCol = calcHeaders.indexOf("ID");
    var calcExwFromBudgetCol = calcHeaders.indexOf("ЦЕНА EXW из Б/З, €");
    if (calcExwFromBudgetCol === -1) {
      calcExwFromBudgetCol = calcHeaders.indexOf("ЦЕНА EXW из Б/З");
    }
    var calcExwPrevCol = calcHeaders.indexOf("EXW предыдущая, €");
    var calcExwCurrentCol = calcHeaders.indexOf("EXW текущая, €");
    var calcAlfaspaCol = calcHeaders.indexOf("EXW ALFASPA текущая, €");
    var calcPurchaseCol = calcHeaders.indexOf("Закупочная цена, ₽");
    var calcDdpCol = calcHeaders.indexOf("DDP-МОСКВА, ₽");
    var calcWholesaleApprovedCol = calcHeaders.indexOf("Утвержденная цена Опт, ₽");
    var calcRrpCol = calcHeaders.indexOf("Утвержденная РРЦ, ₽");
    var calcEcommerceCol = calcHeaders.indexOf("Утвержденая Е-комерс, ₽");

    if (Lib.logInfo) {
      Lib.logInfo(MODULE_TAG + " Индексы столбцов на листе \"" + priceCalcSheet.getName() + "\":");
      Lib.logInfo("  ID: " + calcIdCol);
      Lib.logInfo("  ЦЕНА EXW из Б/З: " + calcExwFromBudgetCol);
      Lib.logInfo("  EXW предыдущая: " + calcExwPrevCol);
      Lib.logInfo("  EXW текущая: " + calcExwCurrentCol);
      Lib.logInfo("  EXW ALFASPA: " + calcAlfaspaCol);
      Lib.logInfo("  Закупочная цена: " + calcPurchaseCol);
      Lib.logInfo("  DDP-МОСКВА: " + calcDdpCol);
      Lib.logInfo("  Утвержденная цена Опт, ₽: " + calcWholesaleApprovedCol);
      Lib.logInfo("  Утвержденная РРЦ, ₽: " + calcRrpCol);
      Lib.logInfo("  Утвержденая Е-комерс, ₽: " + calcEcommerceCol);
    }

    if (calcIdCol === -1) {
      if (Lib.logWarn) {
        Lib.logWarn(MODULE_TAG + ' Не найден столбец "ID" на листе "Расчет цены"');
        Lib.logInfo(MODULE_TAG + " Доступные заголовки: " + calcHeaders.join(" | "));
      }
      return;
    }

    var calcLastRow = priceCalcSheet.getLastRow();
    if (calcLastRow <= HEADER_ROW_INDEX) {
      return;
    }

    var dataRows = calcLastRow - HEADER_ROW_INDEX;
    var dynSheetName = priceDynamicsSheet.getName();

    // Применяем формулы для каждого целевого столбца
    // ВАЖНО: Столбец "ЦЕНА EXW из Б/З, €" НЕ включен в список, так как он заполняется
    // значениями (а не формулами) через функцию copyPriceFromPrimaryToSheets
    var formulas = [
      { targetCol: calcExwPrevCol, sourceCol: dynExwPrevCol },
      { targetCol: calcExwCurrentCol, sourceCol: dynExwCurrentCol },
      { targetCol: calcAlfaspaCol, sourceCol: dynAlfaspaCurrentCol },
      { targetCol: calcPurchaseCol, sourceCol: dynPurchaseCurrentCol },
      { targetCol: calcDdpCol, sourceCol: dynDdpCurrentCol }
    ];

    formulas.forEach(function (item) {
      if (item.targetCol !== -1 && item.sourceCol !== -1) {
        var targetRange = priceCalcSheet.getRange(
          HEADER_ROW_INDEX + 1,
          item.targetCol + 1,
          dataRows,
          1
        );
        _applyIndexMatchFormula(
          targetRange,
          calcIdCol + 1,
          dynSheetName,
          dynIdCol + 1,
          item.sourceCol + 1
        );
      }
    });

    // Применяем форматирование к колонкам с INDEX/MATCH
    // ПРИМЕЧАНИЕ: Столбец "ЦЕНА EXW из Б/З, €" форматируется в copyPriceFromPrimaryToSheets
    if (calcExwPrevCol !== -1) {
      var exwPrevRange = priceCalcSheet.getRange(HEADER_ROW_INDEX + 1, calcExwPrevCol + 1, dataRows, 1);
      _safeApplyNumberFormat(exwPrevRange, {
        format: '0.00 "€"',
        align: "center",
        columnName: "EXW предыдущая, €"
      });
    }
    if (calcExwCurrentCol !== -1) {
      var exwCurrentRange = priceCalcSheet.getRange(HEADER_ROW_INDEX + 1, calcExwCurrentCol + 1, dataRows, 1);
      _safeApplyNumberFormat(exwCurrentRange, {
        format: '0.00 "€"',
        align: "center",
        columnName: "EXW текущая, €"
      });
    }
    if (calcAlfaspaCol !== -1) {
      var alfaspaRange = priceCalcSheet.getRange(HEADER_ROW_INDEX + 1, calcAlfaspaCol + 1, dataRows, 1);
      _safeApplyNumberFormat(alfaspaRange, {
        format: '0.00 "€"',
        align: "center",
        columnName: "EXW ALFASPA текущая, €"
      });
    }
    if (calcPurchaseCol !== -1) {
      var purchaseRange = priceCalcSheet.getRange(HEADER_ROW_INDEX + 1, calcPurchaseCol + 1, dataRows, 1);
      _safeApplyNumberFormat(purchaseRange, {
        format: '#,##0 "₽"',
        align: "center",
        columnName: "Закупочная цена, ₽"
      });
    }
    if (calcDdpCol !== -1) {
      var ddpRange = priceCalcSheet.getRange(HEADER_ROW_INDEX + 1, calcDdpCol + 1, dataRows, 1);
      _safeApplyNumberFormat(ddpRange, {
        format: '#,##0 "₽"',
        align: "center",
        columnName: "DDP-МОСКВА, ₽"
      });
    }

    Lib.refreshOrderSheetExwAlfaspaColumn({
      priceCalcSheet: priceCalcSheet,
      calcIdCol: calcIdCol,
      calcAlfaspaCol: calcAlfaspaCol
    });

    Lib.refreshPriceSheetCommercialColumns({
      priceCalcSheet: priceCalcSheet,
      calcIdCol: calcIdCol,
      calcWholesaleApprovedCol: calcWholesaleApprovedCol,
      calcRrpCol: calcRrpCol,
      calcEcommerceCol: calcEcommerceCol
    });

    if (Lib.logInfo) {
      Lib.logInfo(
        MODULE_TAG +
          " Формулы для листа \"" +
          priceCalcSheet.getName() +
          "\" обновлены"
      );
    }
  }

  /**
   * Применить формулу INDEX/MATCH к диапазону
   */
  function _applyIndexMatchFormula(targetRange, idColCalc, dynSheetName, idColDyn, sourceColDyn) {
    var rows = targetRange.getNumRows();
    if (rows <= 0) {
      return;
    }

    // Формула: =IFERROR(INDEX('Динамика цены'!$COL:$COL, MATCH(RC[ID_OFFSET], 'Динамика цены'!$ID_COL:$ID_COL, 0)), "")
    var idOffset = idColCalc - targetRange.getColumn();
    var formula =
      '=IFERROR(INDEX(' +
      "'" + dynSheetName + "'!C" + sourceColDyn + ':C' + sourceColDyn + ',' +
      ' MATCH(RC[' + idOffset + '], ' +
      "'" + dynSheetName + "'!C" + idColDyn + ':C' + idColDyn + ', 0)), "")';

    var localizedFormula = _localizeFormulaIfNeeded(formula);
    var formulas = [];
    for (var i = 0; i < rows; i++) {
      formulas.push([localizedFormula]);
    }
    targetRange.setFormulasR1C1(formulas);
  }

  /**
   * Заполняет столбец "EXW ALFASPA текущая, €" на листе "Заказ" формулой INDEX/MATCH
   * для подтягивания данных из листа "Расчет цены" по ID.
   */
  function _applyOrderSheetExwAlfaspaFormula(priceCalcSheet, calcIdCol, calcAlfaspaCol) {
    try {
      if (!priceCalcSheet || calcIdCol === -1 || calcAlfaspaCol === -1) {
        return;
      }

      var ss = priceCalcSheet.getParent();
      if (!ss) {
        return;
      }

      var orderSheetName =
        (Lib.CONFIG &&
          Lib.CONFIG.SHEETS &&
          Lib.CONFIG.SHEETS.ORDER_FORM) ||
        "Заказ";
      var orderSheet = ss.getSheetByName(orderSheetName);
      if (!orderSheet) {
        if (Lib.logWarn) {
          Lib.logWarn(
            MODULE_TAG +
              ' Лист "' +
              orderSheetName +
              '" не найден, пропускаем обновление столбца EXW ALFASPA текущая, €'
          );
        }
        return;
      }

      var lastRow = orderSheet.getLastRow();
      if (lastRow <= HEADER_ROW_INDEX) {
        return;
      }

      var lastCol = orderSheet.getLastColumn();
      if (lastCol <= 0) {
        return;
      }

      var headers = orderSheet
        .getRange(HEADER_ROW_INDEX, 1, 1, lastCol)
        .getValues()[0]
        .map(function (value) {
          return String(value || "").replace(/^[\s\u00A0]+|[\s\u00A0]+$/g, "");
        });

      var orderIdCol = headers.indexOf("ID");
      if (orderIdCol === -1) {
        if (Lib.logWarn) {
          Lib.logWarn(
            MODULE_TAG +
              ' Не найден столбец "ID" на листе "' +
              orderSheetName +
              '"'
          );
        }
        return;
      }

      var orderAlfaspaCol = headers.indexOf("EXW  ALFASPA  текущая, €");
      if (orderAlfaspaCol === -1) {
        orderAlfaspaCol = headers.indexOf("EXW ALFASPA текущая, €");
      }
      if (orderAlfaspaCol === -1) {
        if (Lib.logWarn) {
          Lib.logWarn(
            MODULE_TAG +
              ' Не найден столбец "EXW ALFASPA текущая, €" на листе "' +
              orderSheetName +
              '"'
          );
        }
        return;
      }

      var dataRows = lastRow - HEADER_ROW_INDEX;
      if (dataRows <= 0) {
        return;
      }

      var targetRange = orderSheet.getRange(
        HEADER_ROW_INDEX + 1,
        orderAlfaspaCol + 1,
        dataRows,
        1
      );

      _applyIndexMatchFormula(
        targetRange,
        orderIdCol + 1,
        priceCalcSheet.getName(),
        calcIdCol + 1,
        calcAlfaspaCol + 1
      );

      _safeApplyNumberFormat(targetRange, {
        format: '0.00 "€"',
        align: "center",
        columnName: "EXW ALFASPA текущая, € (Заказ)"
      });

      if (Lib.logInfo) {
        Lib.logInfo(
          MODULE_TAG +
            ' Обновлены формулы столбца "EXW ALFASPA текущая, €" на листе "' +
            orderSheetName +
            '"'
        );
      }
    } catch (error) {
      if (Lib.logWarn) {
        Lib.logWarn(
          MODULE_TAG +
            " Не удалось обновить столбец EXW ALFASPA текущая, € на листе Заказ",
          error
        );
      }
    }
  }

  function _applyPriceSheetCommercialColumns(priceCalcSheet, calcIdCol, calcCols) {
    if (!priceCalcSheet || calcIdCol === -1) {
      return;
    }

    var ss = priceCalcSheet.getParent();
    if (!ss) {
      return;
    }

    var priceSheetName =
      (Lib.CONFIG &&
        Lib.CONFIG.SHEETS &&
        Lib.CONFIG.SHEETS.PRICE) ||
      "Прайс";
    var priceSheet = ss.getSheetByName(priceSheetName);
    if (!priceSheet) {
      if (Lib.logWarn) {
        Lib.logWarn(
          MODULE_TAG +
            ' Лист "' +
            priceSheetName +
            '" не найден, пропускаем обновление ценового блока'
        );
      }
      return;
    }

    var lastRow = priceSheet.getLastRow();
    if (lastRow <= HEADER_ROW_INDEX) {
      return;
    }

    var lastCol = priceSheet.getLastColumn();
    if (lastCol <= 0) {
      return;
    }

    var headers = priceSheet
      .getRange(HEADER_ROW_INDEX, 1, 1, lastCol)
      .getValues()[0]
      .map(function (value) {
        return String(value || "").replace(/^[\s\u00A0]+|[\s\u00A0]+$/g, "");
      });

    var priceIdCol = headers.indexOf("ID");
    if (priceIdCol === -1) {
      if (Lib.logWarn) {
        Lib.logWarn(
          MODULE_TAG +
            ' Не найден столбец "ID" на листе "' +
            priceSheetName +
            '"'
        );
      }
      return;
    }

    var dataRows = lastRow - HEADER_ROW_INDEX;
    if (dataRows <= 0) {
      return;
    }

    var mappings = [
      { header: "Цена", sourceCol: calcCols.wholesale },
      { header: "RRP", sourceCol: calcCols.rrp },
      { header: "Е-comerc", sourceCol: calcCols.ecommerce }
    ];

    mappings.forEach(function (map) {
      if (map.sourceCol === -1) {
        return;
      }
      var targetCol = headers.indexOf(map.header);
      if (targetCol === -1) {
        if (Lib.logWarn) {
          Lib.logWarn(
            MODULE_TAG +
              ' Не найден столбец "' +
              map.header +
              '" на листе "' +
              priceSheetName +
              '"'
          );
        }
        return;
      }

      var targetRange = priceSheet.getRange(
        HEADER_ROW_INDEX + 1,
        targetCol + 1,
        dataRows,
        1
      );

      _applyIndexMatchFormula(
        targetRange,
        priceIdCol + 1,
        priceCalcSheet.getName(),
        calcIdCol + 1,
        map.sourceCol + 1
      );

      _safeApplyNumberFormat(targetRange, {
        format: '#,##0 "₽"',
        align: "center",
        columnName: map.header + ' (Прайс)'
      });
    });

    if (Lib.logInfo) {
      Lib.logInfo(
        MODULE_TAG +
          ' Обновлены формулы столбцов "Цена/RRP/Е-comerc" на листе "' +
          priceSheetName +
          '"'
      );
    }
  }

  /**
   * Публичная функция: синхронизировать столбец EXW ALFASPA текущая, €
   * на листе "Заказ" с данными листа "Расчет цены".
   * @param {Object} [options]
   * @param {GoogleAppsScript.Spreadsheet.Sheet} [options.priceCalcSheet] - лист "Расчет цены"
   * @param {number} [options.calcIdCol] - индекс столбца ID на листе "Расчет цены"
   * @param {number} [options.calcAlfaspaCol] - индекс столбца EXW ALFASPA текущая, €
   */
  Lib.refreshOrderSheetExwAlfaspaColumn = function (options) {
    options = options || {};
    try {
      var priceCalcSheet =
        options.priceCalcSheet ||
        (function () {
          var ss = SpreadsheetApp.getActiveSpreadsheet();
          if (!ss) {
            return null;
          }
          var priceCalcSheetName =
            (Lib.CONFIG &&
              Lib.CONFIG.SHEETS &&
              Lib.CONFIG.SHEETS.PRICE_CALCULATION) ||
            "Расчет цены";
          return ss.getSheetByName(priceCalcSheetName);
        })();

      if (!priceCalcSheet) {
        if (Lib.logWarn) {
          Lib.logWarn(
            MODULE_TAG +
              ' Не найден лист "Расчет цены", пропускаем обновление EXW ALFASPA для листа "Заказ"'
          );
        }
        return;
      }

      var calcIdCol =
        typeof options.calcIdCol === "number" && options.calcIdCol >= 0
          ? options.calcIdCol
          : -1;
      var calcAlfaspaCol =
        typeof options.calcAlfaspaCol === "number" && options.calcAlfaspaCol >= 0
          ? options.calcAlfaspaCol
          : -1;

      if (calcIdCol === -1 || calcAlfaspaCol === -1) {
        var calcLastCol = priceCalcSheet.getLastColumn();
        if (calcLastCol <= 0) {
          return;
        }
        var calcHeaders = priceCalcSheet
          .getRange(HEADER_ROW_INDEX, 1, 1, calcLastCol)
          .getValues()[0]
          .map(function (value) {
            return String(value || "").replace(/^[\s\u00A0]+|[\s\u00A0]+$/g, "");
          });

        if (calcIdCol === -1) {
          calcIdCol = calcHeaders.indexOf("ID");
        }
        if (calcAlfaspaCol === -1) {
          calcAlfaspaCol = calcHeaders.indexOf("EXW ALFASPA текущая, €");
        }

        if ((calcIdCol === -1 || calcAlfaspaCol === -1) && Lib.logWarn) {
          Lib.logWarn(
            MODULE_TAG +
              ' Не удалось найти столбцы "ID" или "EXW ALFASPA текущая, €" на листе "' +
              priceCalcSheet.getName() +
              '"'
          );
        }
      }

      if (calcIdCol === -1 || calcAlfaspaCol === -1) {
        return;
      }

      _applyOrderSheetExwAlfaspaFormula(priceCalcSheet, calcIdCol, calcAlfaspaCol);
    } catch (error) {
      if (Lib.logWarn) {
        Lib.logWarn(
          MODULE_TAG +
            " Ошибка при попытке синхронизировать столбец EXW ALFASPA текущая, € на листе Заказ",
          error
        );
      }
    }
  };

  /**
   * Публичная функция: синхронизировать столбцы "Цена", "RRP" и "Е-comerc"
   * на листе "Прайс" с листом "Расчет цены" по ID.
   * @param {Object} [options]
   * @param {GoogleAppsScript.Spreadsheet.Sheet} [options.priceCalcSheet]
   * @param {number} [options.calcIdCol]
   * @param {number} [options.calcWholesaleApprovedCol]
   * @param {number} [options.calcRrpCol]
   * @param {number} [options.calcEcommerceCol]
   */
  Lib.refreshPriceSheetCommercialColumns = function (options) {
    options = options || {};
    try {
      var priceCalcSheet =
        options.priceCalcSheet ||
        (function () {
          var ss = SpreadsheetApp.getActiveSpreadsheet();
          if (!ss) {
            return null;
          }
          var priceCalcSheetName =
            (Lib.CONFIG &&
              Lib.CONFIG.SHEETS &&
              Lib.CONFIG.SHEETS.PRICE_CALCULATION) ||
            "Расчет цены";
          return ss.getSheetByName(priceCalcSheetName);
        })();

      if (!priceCalcSheet) {
        if (Lib.logWarn) {
          Lib.logWarn(
            MODULE_TAG +
              ' Не найден лист "Расчет цены", пропускаем обновление цен для листа "Прайс"'
          );
        }
        return;
      }

      var calcIdCol =
        typeof options.calcIdCol === "number" && options.calcIdCol >= 0
          ? options.calcIdCol
          : -1;
      var calcWholesaleApprovedCol =
        typeof options.calcWholesaleApprovedCol === "number" && options.calcWholesaleApprovedCol >= 0
          ? options.calcWholesaleApprovedCol
          : -1;
      var calcRrpCol =
        typeof options.calcRrpCol === "number" && options.calcRrpCol >= 0
          ? options.calcRrpCol
          : -1;
      var calcEcommerceCol =
        typeof options.calcEcommerceCol === "number" && options.calcEcommerceCol >= 0
          ? options.calcEcommerceCol
          : -1;

      if (
        calcIdCol === -1 ||
        calcWholesaleApprovedCol === -1 ||
        calcRrpCol === -1 ||
        calcEcommerceCol === -1
      ) {
        var calcLastCol = priceCalcSheet.getLastColumn();
        if (calcLastCol <= 0) {
          return;
        }
        var calcHeaders = priceCalcSheet
          .getRange(HEADER_ROW_INDEX, 1, 1, calcLastCol)
          .getValues()[0]
          .map(function (value) {
            return String(value || "").replace(/^[\s\u00A0]+|[\s\u00A0]+$/g, "");
          });

        if (calcIdCol === -1) {
          calcIdCol = calcHeaders.indexOf("ID");
        }
        if (calcWholesaleApprovedCol === -1) {
          calcWholesaleApprovedCol = calcHeaders.indexOf("Утвержденная цена Опт, ₽");
        }
        if (calcRrpCol === -1) {
          calcRrpCol = calcHeaders.indexOf("Утвержденная РРЦ, ₽");
        }
        if (calcEcommerceCol === -1) {
          calcEcommerceCol = calcHeaders.indexOf("Утвержденая Е-комерс, ₽");
        }

        var missing = [];
        if (calcIdCol === -1) missing.push("ID");
        if (calcWholesaleApprovedCol === -1) missing.push("Утвержденная цена Опт, ₽");
        if (calcRrpCol === -1) missing.push("Утвержденная РРЦ, ₽");
        if (calcEcommerceCol === -1) missing.push("Утвержденая Е-комерс, ₽");

        if (missing.length && Lib.logWarn) {
          Lib.logWarn(
            MODULE_TAG +
              " Не найдены столбцы на листе \"" +
              priceCalcSheet.getName() +
              "\": " +
              missing.join(", ")
          );
        }
      }

      if (
        calcIdCol === -1 ||
        calcWholesaleApprovedCol === -1 ||
        calcRrpCol === -1 ||
        calcEcommerceCol === -1
      ) {
        return;
      }

      _applyPriceSheetCommercialColumns(priceCalcSheet, calcIdCol, {
        wholesale: calcWholesaleApprovedCol,
        rrp: calcRrpCol,
        ecommerce: calcEcommerceCol
      });
    } catch (error) {
      if (Lib.logWarn) {
        Lib.logWarn(
          MODULE_TAG + " Ошибка при попытке обновить цены на листе Прайс",
          error
        );
      }
    }
  };

  /**
   * Публичная функция: создать новый год для динамики на листе "Расчет цены"
   * Копирует текущие утвержденные значения в столбцы прошлого периода
   * ВАЖНО: НЕ очищает столбцы EXW/Закупочная/DDP - это делается отдельно после обновления формул
   */
  Lib.createNewYearForDynamics = function () {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) {
        throw new Error("Не удалось получить активную таблицу");
      }

      var sheetName =
        Lib.CONFIG &&
        Lib.CONFIG.SHEETS &&
        Lib.CONFIG.SHEETS.PRICE_CALCULATION
          ? Lib.CONFIG.SHEETS.PRICE_CALCULATION
          : "Расчет цены";
      var sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        _showAlert(
          "Лист \"" + sheetName + "\" не найден",
          "Проверьте, что лист существует в таблице."
        );
        return;
      }

      var lastRow = sheet.getLastRow();
      if (lastRow <= HEADER_ROW_INDEX) {
        _showAlert(
          "Лист \"" + sheetName + "\" пуст",
          "Нет данных для копирования."
        );
        return;
      }

      var lastCol = sheet.getLastColumn();
      if (lastCol <= 0) {
        _showAlert(
          "Лист \"" + sheetName + "\" пуст",
          "Нет столбцов для работы."
        );
        return;
      }

      var headers = sheet
        .getRange(HEADER_ROW_INDEX, 1, 1, lastCol)
        .getValues()[0]
        .map(function (value) {
          return String(value || "").replace(/^[\s\u00A0]+|[\s\u00A0]+$/g, "");
        });

      // Находим индексы необходимых столбцов
      var priceWholesaleApprovedIdx = headers.indexOf("Утвержденная цена Опт, ₽");
      var coefficientFactIdx = headers.indexOf("К-т ФАКТ");
      var priceRetailApprovedIdx = headers.indexOf("Утвержденная РРЦ, ₽");

      // Находим индексы столбцов прошлого периода
      var priceWholesalePrevIdx = headers.indexOf("ОПТ прошлого периода, ₽");
      var coefficientFactPrevIdx = headers.indexOf("К-т ФАКТ прош");
      var priceRetailPrevIdx = headers.indexOf("РРЦ прошлого периода, ₽");
      var dynamicsIdx = headers.indexOf("Динамика");

      // Проверяем наличие исходных столбцов
      if (priceWholesaleApprovedIdx === -1 || coefficientFactIdx === -1 || priceRetailApprovedIdx === -1) {
        _showAlert(
          "Не найдены необходимые столбцы",
          "Убедитесь, что на листе присутствуют столбцы: \"Утвержденная цена Опт, ₽\", \"К-т ФАКТ\", \"Утвержденная РРЦ, ₽\""
        );
        return;
      }

      // Если столбцы прошлого периода не существуют, создаем их
      var columnsToInsert = [];
      if (priceWholesalePrevIdx === -1) {
        columnsToInsert.push({ header: "ОПТ прошлого периода, ₽", numberFormat: '#,##0 "₽"' });
      }
      if (coefficientFactPrevIdx === -1) {
        columnsToInsert.push({ header: "К-т ФАКТ прош", numberFormat: "0.0" });
      }
      if (priceRetailPrevIdx === -1) {
        columnsToInsert.push({ header: "РРЦ прошлого периода, ₽", numberFormat: '#,##0 "₽"' });
      }
      if (dynamicsIdx === -1) {
        columnsToInsert.push({ header: "Динамика", numberFormat: "0%" });
      }

      // Вставляем новые столбцы после "Утвержденная РРЦ, ₽"
      if (columnsToInsert.length > 0) {
        var insertPosition = priceRetailApprovedIdx + 1; // после Утвержденная РРЦ
        sheet.insertColumnsAfter(insertPosition, columnsToInsert.length);

        // Устанавливаем заголовки
        var headerRange = sheet.getRange(
          HEADER_ROW_INDEX,
          insertPosition + 1,
          1,
          columnsToInsert.length
        );
        headerRange.setValues([columnsToInsert.map(function (col) { return col.header; })]);

        // Копируем формат заголовка
        var sourceHeaderRange = sheet.getRange(HEADER_ROW_INDEX, priceRetailApprovedIdx + 1);
        sourceHeaderRange.copyFormatToRange(
          sheet,
          insertPosition + 1,
          insertPosition + columnsToInsert.length,
          HEADER_ROW_INDEX,
          HEADER_ROW_INDEX
        );

        // Обновляем заголовки
        headers = sheet
          .getRange(HEADER_ROW_INDEX, 1, 1, sheet.getLastColumn())
          .getValues()[0]
          .map(function (value) {
            return String(value || "").replace(/^[\s\u00A0]+|[\s\u00A0]+$/g, "");
          });

        // Обновляем индексы
        priceWholesalePrevIdx = headers.indexOf("ОПТ прошлого периода, ₽");
        coefficientFactPrevIdx = headers.indexOf("К-т ФАКТ прош");
        priceRetailPrevIdx = headers.indexOf("РРЦ прошлого периода, ₽");
        dynamicsIdx = headers.indexOf("Динамика");

        // ВАЖНО: После вставки столбцов нужно обновить индексы исходных столбцов,
        // так как они могли сдвинуться вправо
        priceWholesaleApprovedIdx = headers.indexOf("Утвержденная цена Опт, ₽");
        coefficientFactIdx = headers.indexOf("К-т ФАКТ");
        priceRetailApprovedIdx = headers.indexOf("Утвержденная РРЦ, ₽");
      }

      var dataRows = lastRow - HEADER_ROW_INDEX;

      if (Lib.logDebug) {
        Lib.logDebug(MODULE_TAG + " Индексы столбцов после возможной вставки:");
        Lib.logInfo("  Утвержденная цена Опт, ₽: " + priceWholesaleApprovedIdx);
        Lib.logInfo("  К-т ФАКТ: " + coefficientFactIdx);
        Lib.logInfo("  Утвержденная РРЦ, ₽: " + priceRetailApprovedIdx);
        Lib.logInfo("  ОПТ прошлого периода, ₽: " + priceWholesalePrevIdx);
        Lib.logInfo("  К-т ФАКТ прош: " + coefficientFactPrevIdx);
        Lib.logInfo("  РРЦ прошлого периода, ₽: " + priceRetailPrevIdx);
        Lib.logInfo("  Динамика: " + dynamicsIdx);
      }

      // Копируем значения из текущих столбцов в столбцы прошлого периода
      if (priceWholesalePrevIdx !== -1 && priceWholesaleApprovedIdx !== -1) {
        var sourceRange = sheet.getRange(HEADER_ROW_INDEX + 1, priceWholesaleApprovedIdx + 1, dataRows, 1);
        var targetRange = sheet.getRange(HEADER_ROW_INDEX + 1, priceWholesalePrevIdx + 1, dataRows, 1);
        sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
        _safeApplyNumberFormat(targetRange, {
          format: '#,##0 "₽"',
          align: "center",
          columnName: "ОПТ прошлого периода, ₽"
        });
      }

      if (coefficientFactPrevIdx !== -1 && coefficientFactIdx !== -1) {
        var sourceRange = sheet.getRange(HEADER_ROW_INDEX + 1, coefficientFactIdx + 1, dataRows, 1);
        var targetRange = sheet.getRange(HEADER_ROW_INDEX + 1, coefficientFactPrevIdx + 1, dataRows, 1);
        sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
        _safeApplyNumberFormat(targetRange, {
          format: "0.0",
          align: "center",
          columnName: "К-т ФАКТ прош"
        });
      }

      if (priceRetailPrevIdx !== -1 && priceRetailApprovedIdx !== -1) {
        var sourceRange = sheet.getRange(HEADER_ROW_INDEX + 1, priceRetailApprovedIdx + 1, dataRows, 1);
        var targetRange = sheet.getRange(HEADER_ROW_INDEX + 1, priceRetailPrevIdx + 1, dataRows, 1);
        sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
        _safeApplyNumberFormat(targetRange, {
          format: '#,##0 "₽"',
          align: "center",
          columnName: "РРЦ прошлого периода, ₽"
        });
      }

      // Создаем формулу для столбца "Динамика"
      if (dynamicsIdx !== -1 && priceWholesalePrevIdx !== -1 && priceWholesaleApprovedIdx !== -1) {
        var dynamicsRange = sheet.getRange(HEADER_ROW_INDEX + 1, dynamicsIdx + 1, dataRows, 1);
        var offsetPrev = priceWholesalePrevIdx - dynamicsIdx;
        var offsetCurrent = priceWholesaleApprovedIdx - dynamicsIdx;
        var formula = '=IF(OR(LEN(R[0]C[' + offsetPrev + '])=0, LEN(R[0]C[' + offsetCurrent + '])=0), "", R[0]C[' + offsetPrev + '] / R[0]C[' + offsetCurrent + '] - 1)';
        _setFormulasR1C1(dynamicsRange, formula);
        _safeApplyNumberFormat(dynamicsRange, {
          format: "0%",
          align: "center",
          columnName: "Динамика"
        });
      }

      SpreadsheetApp.flush();

      if (Lib.logInfo) {
        Lib.logInfo(
          MODULE_TAG +
            " Данные прошлого периода обновлены на листе \"" +
            sheetName +
            "\" (без очистки столбцов EXW/Закупочная/DDP)"
        );
      }
    } catch (error) {
      if (Lib.logError) {
        Lib.logError(MODULE_TAG + " Ошибка создания нового года для динамики", error);
      } else {
        console.error(MODULE_TAG + " Ошибка создания нового года для динамики", error);
      }
      _showAlert("Ошибка", error.message || String(error));
    }
  };

  /**
   * Вспомогательная функция: очистить столбцы EXW/Закупочная/DDP на листе "Расчет цены"
   * после обновления формул. Вызывается из addNewYearColumnsToPriceDynamics.
   * @private
   */
  function _clearCalculatedColumnsOnPriceCalc() {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) {
        return;
      }

      var sheetName =
        Lib.CONFIG &&
        Lib.CONFIG.SHEETS &&
        Lib.CONFIG.SHEETS.PRICE_CALCULATION
          ? Lib.CONFIG.SHEETS.PRICE_CALCULATION
          : "Расчет цены";
      var sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        return;
      }

      var lastRow = sheet.getLastRow();
      if (lastRow <= HEADER_ROW_INDEX) {
        return;
      }

      var lastCol = sheet.getLastColumn();
      if (lastCol <= 0) {
        return;
      }

      var headers = sheet
        .getRange(HEADER_ROW_INDEX, 1, 1, lastCol)
        .getValues()[0]
        .map(function (value) {
          return String(value || "").replace(/^[\s\u00A0]+|[\s\u00A0]+$/g, "");
        });

      var dataRows = lastRow - HEADER_ROW_INDEX;

      // Очищаем столбцы с текущими данными (оставляем только формулы)
      var columnsToClear = [
        { name: "EXW текущая, €", format: '0.00 "€"' },
        { name: "EXW ALFASPA текущая, €", format: '0.00 "€"' },
        { name: "Закупочная цена, ₽", format: '#,##0 "₽"' },
        { name: "DDP-МОСКВА, ₽", format: '#,##0 "₽"' }
      ];

      columnsToClear.forEach(function (col) {
        var colIdx = headers.indexOf(col.name);
        if (colIdx !== -1) {
          var range = sheet.getRange(HEADER_ROW_INDEX + 1, colIdx + 1, dataRows, 1);

          // Получаем текущие формулы (они уже должны быть обновлены на шаге 3)
          var formulas = range.getFormulas();
          var values = range.getValues();

          // Проходим по каждой ячейке и очищаем только значения (не формулы)
          for (var i = 0; i < dataRows; i++) {
            // Если в ячейке нет формулы, но есть значение - очищаем
            if (!formulas[i][0] && values[i][0]) {
              sheet.getRange(HEADER_ROW_INDEX + 1 + i, colIdx + 1).clearContent();
            }
            // Если есть формула - не трогаем, она уже обновлена
          }

          // Применяем форматирование ко всему столбцу
          _safeApplyNumberFormat(range, {
            format: col.format,
            align: "center",
            columnName: col.name
          });
        }
      });

      SpreadsheetApp.flush();

      if (Lib.logInfo) {
        Lib.logInfo(
          MODULE_TAG +
            " Очищены столбцы EXW/Закупочная/DDP на листе \"" +
            sheetName +
            '"'
        );
      }
    } catch (error) {
      if (Lib.logError) {
        Lib.logError(MODULE_TAG + " Ошибка очистки столбцов", error);
      }
    }
  }

  /**
   * Публичная функция: обновить вертикальные границы на листе "Динамика цены"
   * Вызывается после добавления новых строк для обновления границ на всю высоту таблицы
   */
  Lib.updatePriceDynamicsBorders = function() {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) {
        return;
      }

      var sheetName =
        Lib.CONFIG &&
        Lib.CONFIG.SHEETS &&
        Lib.CONFIG.SHEETS.PRICE_DYNAMICS
          ? Lib.CONFIG.SHEETS.PRICE_DYNAMICS
          : "Динамика цены";
      var sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        return;
      }

      var lastRow = sheet.getLastRow();
      if (lastRow <= HEADER_ROW_INDEX) {
        return;
      }

      var lastCol = sheet.getLastColumn();
      if (lastCol <= 0) {
        return;
      }

      var headers = sheet
        .getRange(HEADER_ROW_INDEX, 1, 1, lastCol)
        .getValues()[0]
        .map(function (value) {
          return String(value || "").replace(/^[\s\u00A0]+|[\s\u00A0]+$/g, "");
        });

      var totalRows = lastRow - HEADER_ROW_INDEX + 1;
      var bordersApplied = 0;

      // Применяем границу перед столбцом "ЦЕНА EXW из Б/З, €"
      var exwFromBudgetIdx = headers.indexOf("ЦЕНА EXW из Б/З, €");
      if (exwFromBudgetIdx !== -1) {
        var borderRange = sheet.getRange(HEADER_ROW_INDEX, exwFromBudgetIdx + 1, totalRows, 1);
        borderRange.setBorder(
          null, // top
          true, // left - жирная граница
          null, // bottom
          null, // right
          null, // vertical
          null, // horizontal
          null, // color
          SpreadsheetApp.BorderStyle.SOLID_THICK
        );
        bordersApplied++;
      }

      // Динамически находим все столбцы вида "EXW YYYY, €" и применяем к ним границы
      for (var i = 0; i < headers.length; i++) {
        var header = headers[i];
        // Проверяем паттерн "EXW YYYY, €" где YYYY - год (4 цифры)
        if (/^EXW \d{4}, €$/.test(header)) {
          var borderRange = sheet.getRange(HEADER_ROW_INDEX, i + 1, totalRows, 1);
          borderRange.setBorder(
            null, // top
            true, // left - жирная граница
            null, // bottom
            null, // right
            null, // vertical
            null, // horizontal
            null, // color
            SpreadsheetApp.BorderStyle.SOLID_THICK
          );
          bordersApplied++;
        }
      }

      if (Lib.logInfo) {
        Lib.logInfo(
          MODULE_TAG +
            " Обновлены вертикальные границы на листе \"" +
            sheetName +
            "\" для " + totalRows + " строк (" + bordersApplied + " границ)"
        );
      }
    } catch (error) {
      if (Lib.logError) {
        Lib.logError(MODULE_TAG + " Ошибка обновления границ на Динамика цены", error);
      }
    }
  };

  /**
   * Публичная функция: обновить вертикальные границы на листе "Расчет цены"
   * Вызывается после добавления новых строк для обновления границ на всю высоту таблицы
   */
  Lib.updatePriceCalculationBorders = function() {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) {
        return;
      }

      var sheetName =
        Lib.CONFIG &&
        Lib.CONFIG.SHEETS &&
        Lib.CONFIG.SHEETS.PRICE_CALCULATION
          ? Lib.CONFIG.SHEETS.PRICE_CALCULATION
          : "Расчет цены";
      var sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        return;
      }

      var lastRow = sheet.getLastRow();
      if (lastRow <= HEADER_ROW_INDEX) {
        return;
      }

      var lastCol = sheet.getLastColumn();
      if (lastCol <= 0) {
        return;
      }

      var headers = sheet
        .getRange(HEADER_ROW_INDEX, 1, 1, lastCol)
        .getValues()[0]
        .map(function (value) {
          return String(value || "").replace(/^[\s\u00A0]+|[\s\u00A0]+$/g, "");
        });

      var totalRows = lastRow - HEADER_ROW_INDEX + 1;

      // Список столбцов, перед которыми нужно применить вертикальную границу
      var borderColumns = [
        "ЦЕНА EXW из Б/З, €",
        "EXW предыдущая, €",
        "EXW текущая, €",
        "К-т",
        "% наценка для РРЦ",
        "Утвержденная цена Опт, ₽",
        "РРЦ прошлого периода, ₽",
        "Утвержденая Е-комерс, ₽",
        "Цена дистрибьютора -30%, ₽",
        "Цена крупный опт -10 %, ₽",
        "Max скидка-50%, ₽",
        "ОПТ прошлого периода, ₽"
      ];

      var bordersApplied = 0;

      // Применяем жирную границу перед каждым из указанных столбцов
      borderColumns.forEach(function(columnName) {
        var colIdx = headers.indexOf(columnName);
        if (colIdx !== -1) {
          try {
            var borderRange = sheet.getRange(HEADER_ROW_INDEX, colIdx + 1, totalRows, 1);
            borderRange.setBorder(
              null, // top
              true, // left - жирная граница
              null, // bottom
              null, // right
              null, // vertical
              null, // horizontal
              null, // color
              SpreadsheetApp.BorderStyle.SOLID_THICK
            );
            bordersApplied++;
          } catch (e) {
            if (Lib.logWarn) {
              Lib.logWarn(MODULE_TAG + ' Не удалось установить границу для столбца "' + columnName + '": ' + e.message);
            }
          }
        }
      });

      if (Lib.logInfo) {
        Lib.logInfo(
          MODULE_TAG +
            " Обновлены вертикальные границы на листе \"" +
            sheetName +
            "\" для " + totalRows + " строк (" + bordersApplied + " границ)"
        );
      }
    } catch (error) {
      if (Lib.logError) {
        Lib.logError(MODULE_TAG + " Ошибка обновления границ", error);
      }
    }
  };

  /**
   * Публичная функция: обновить вертикальные границы на листе "Заказ"
   * Вызывается после добавления новых строк для обновления границ на всю высоту таблицы
   */
  Lib.updateOrderFormBorders = function() {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) {
        return;
      }

      var sheetName =
        Lib.CONFIG &&
        Lib.CONFIG.SHEETS &&
        Lib.CONFIG.SHEETS.ORDER_FORM
          ? Lib.CONFIG.SHEETS.ORDER_FORM
          : "Заказ";
      var sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        return;
      }

      var lastRow = sheet.getLastRow();
      if (lastRow <= HEADER_ROW_INDEX) {
        return;
      }

      var lastCol = sheet.getLastColumn();
      if (lastCol <= 0) {
        return;
      }

      var headers = sheet
        .getRange(HEADER_ROW_INDEX, 1, 1, lastCol)
        .getValues()[0]
        .map(function (value) {
          return String(value || "").replace(/^[\s\u00A0]+|[\s\u00A0]+$/g, "");
        });

      var totalRows = lastRow - HEADER_ROW_INDEX + 1;

      // Список столбцов, перед которыми нужно применить вертикальную границу
      var borderColumns = [
        "ID",
        "ПРОДАЖИ",
        "Среднемесячные продажи, шт",
        "Необходимо заказать, шт",
        "АКЦИИ",
        "Набор",
        "Добавить в прайс"
      ];

      var bordersApplied = 0;

      // Применяем жирную границу перед каждым из указанных столбцов
      borderColumns.forEach(function(columnName) {
        var colIdx = headers.indexOf(columnName);
        if (colIdx !== -1) {
          try {
            var borderRange = sheet.getRange(HEADER_ROW_INDEX, colIdx + 1, totalRows, 1);
            borderRange.setBorder(
              null, // top
              true, // left - жирная граница
              null, // bottom
              null, // right
              null, // vertical
              null, // horizontal
              null, // color
              SpreadsheetApp.BorderStyle.SOLID_THICK
            );
            bordersApplied++;
          } catch (e) {
            if (Lib.logWarn) {
              Lib.logWarn(MODULE_TAG + ' Не удалось установить границу для столбца "' + columnName + '": ' + e.message);
            }
          }
        }
      });

      if (Lib.logInfo) {
        Lib.logInfo(
          MODULE_TAG +
            " Обновлены вертикальные границы на листе \"" +
            sheetName +
            "\" для " + totalRows + " строк (" + bordersApplied + " границ)"
        );
      }
    } catch (error) {
      if (Lib.logError) {
        Lib.logError(MODULE_TAG + " Ошибка обновления границ на листе Заказ", error);
      }
    }
  };

  /**
   * Публичная функция: применить расчетные формулы на листе "Расчет цены"
   * @param {boolean} silent - Не показывать уведомления (для автозапуска)
   */
  Lib.applyCalculationFormulas = function (silent) {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) {
        throw new Error("Не удалось получить активную таблицу");
      }

      var sheetName =
        Lib.CONFIG &&
        Lib.CONFIG.SHEETS &&
        Lib.CONFIG.SHEETS.PRICE_CALCULATION
          ? Lib.CONFIG.SHEETS.PRICE_CALCULATION
          : "Расчет цены";
      var sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        if (!silent) {
          _showAlert(
            "Лист \"" + sheetName + "\" не найден",
            "Проверьте, что лист существует в таблице."
          );
        }
        return;
      }

      var priceDynamicsSheetName =
        Lib.CONFIG &&
        Lib.CONFIG.SHEETS &&
        Lib.CONFIG.SHEETS.PRICE_DYNAMICS
          ? Lib.CONFIG.SHEETS.PRICE_DYNAMICS
          : "Динамика цены";
      var priceDynamicsSheet = ss.getSheetByName(priceDynamicsSheetName);
      if (priceDynamicsSheet) {
        // Сначала применяем формулы INDEX/MATCH для подтягивания данных из "Динамика цены"
        try {
          _applyPriceCalculationFormulas(sheet, priceDynamicsSheet);
        } catch (error) {
          if (_isDataTypeFormatError(error)) {
            _logFormatLockHint("Расчет цены (INDEX/MATCH)", error.message);
          } else {
            throw error;
          }
        }
      }

      // Затем применяем расчетные формулы
      _applyAllCalculationFormulas(sheet, ss);

      // Обновляем вертикальные границы после применения формул
      if (typeof Lib.updatePriceCalculationBorders === 'function') {
        Lib.updatePriceCalculationBorders();
      }

      if (!silent) {
        ss.toast("Формулы расчета цены обновлены", "Авто цены", 5);
      }
    } catch (error) {
      if (Lib.logError) {
        Lib.logError(MODULE_TAG + " Ошибка применения расчетных формул", error);
      } else {
        console.error(MODULE_TAG + " Ошибка применения расчетных формул", error);
      }
      if (!silent) {
        _showAlert("Ошибка", error.message || String(error));
      }
    }
  };

  /**
   * Парсит значение объёма из строки (упрощенная формула)
   * Примеры: "50 мл." → 50, "6 х 130 гр." → 780, "6 х 4,5 гр. / 6 х 18 гр." → 135
   */
  function _parseVolumeFormula(volumeColOffset) {
    // Упрощенная формула для парсинга объёма
    // Для случаев "6 х 4,5 гр. / 6 х 18 гр." требуется сложная логика через Apps Script
    // Здесь используем базовую версию для простых случаев
    return (
      'IF(LEN(R[0]C[' + volumeColOffset + '])=0, 1, ' +
      'IF(ISNUMBER(SEARCH("х", R[0]C[' + volumeColOffset + '])), ' +
      '  VALUE(REGEXEXTRACT(R[0]C[' + volumeColOffset + '], "\\d+")) * VALUE(REGEXEXTRACT(SUBSTITUTE(R[0]C[' + volumeColOffset + '], " ", ""), "х([\\d,]+)")),' +
      '  VALUE(REGEXEXTRACT(R[0]C[' + volumeColOffset + '], "[\\d,]+"))' +
      '))'
    );
  }

  /**
   * Применить все расчетные формулы на листе "Расчет цены"
   */
  function _applyAllCalculationFormulas(sheet, ss) {
    if (Lib.logInfo) {
      Lib.logInfo(MODULE_TAG + " Начало применения расчетных формул на листе \"" + sheet.getName() + "\"");
    }

    var currentYear = new Date().getFullYear();
    var lastRow = sheet.getLastRow();
    if (lastRow <= HEADER_ROW_INDEX) {
      if (Lib.logWarn) {
        Lib.logWarn(MODULE_TAG + " Лист \"" + sheet.getName() + "\" не содержит данных для применения формул");
      }
      return;
    }

    var lastCol = sheet.getLastColumn();
    if (lastCol <= 0) {
      return;
    }

    var headers = sheet
      .getRange(HEADER_ROW_INDEX, 1, 1, lastCol)
      .getValues()[0]
      .map(function (value) {
        // Удаляем все пробелы в начале и конце, включая неразрывные пробелы
        return String(value || "").replace(/^[\s\u00A0]+|[\s\u00A0]+$/g, "");
      });

    // Находим индексы всех нужных столбцов
    var cols = {
      volume: headers.indexOf("Объём"),
      exwCurrent: headers.indexOf("EXW текущая, €"),
      exwFromBudget: headers.indexOf("ЦЕНА EXW из Б/З, €"),
      coefficient: headers.indexOf("К-т"),
      priceWholesaleEur: headers.indexOf("Расчетная цена Опт, €"),
      priceWholesaleRub: headers.indexOf("Расчетная цена Опт, ₽"),
      costPerMl: headers.indexOf("Стоимость за 1 мл, ₽"),
      markupPercent: headers.indexOf("% наценка для РРЦ"),
      priceRetailCalc: headers.indexOf("Расчетная цена РРЦ, ₽"),
      priceWholesaleApproved: headers.indexOf("Утвержденная цена Опт, ₽"),
      costPerMlApproved: headers.indexOf("Стоимость за 1 мл итог, ₽"),
      coefficientFact: headers.indexOf("К-т ФАКТ"),
      priceRetailApproved: headers.indexOf("Утвержденная РРЦ, ₽"),
      priceEcommerce: headers.indexOf("Утвержденая Е-комерс, ₽"),
      priceDistributor: headers.indexOf("Цена дистрибьютора -30%, ₽"),
      coefficientDistributor: headers.indexOf("К-т дистр."),
      priceWholesaleLarge: headers.indexOf("Цена крупный опт -10 %, ₽"),
      coefficientLarge: headers.indexOf("К-т круп. опт."),
      priceMaxDiscount: headers.indexOf("Max скидка-50%, ₽"),
      coefficientMaxDiscount: headers.indexOf("К-т мах скидка")
    };

    if (Lib.logInfo) {
      var missingCols = [];
      for (var key in cols) {
        if (cols[key] === -1) {
          missingCols.push(key);
        }
      }
      if (missingCols.length > 0) {
        Lib.logWarn(MODULE_TAG + " Не найдены столбцы на листе \"" + sheet.getName() + "\": " + missingCols.join(", "));
        Lib.logInfo(MODULE_TAG + " Доступные заголовки: " + headers.join(" | "));
      }
    }

    var dataRows = lastRow - HEADER_ROW_INDEX;

    // Получаем ссылку на справочник
    var referenceSheet = ss.getSheetByName("Справочник");
    var currencyRef = "100";
    var markupRef = "1";

    if (referenceSheet) {
      var refHeaders = referenceSheet
        .getRange(1, 1, 1, referenceSheet.getLastColumn())
        .getValues()[0]
        .map(function (v) {
          // Удаляем все пробелы в начале и конце, включая неразрывные пробелы
          return String(v || "").replace(/^[\s\u00A0]+|[\s\u00A0]+$/g, "");
        });

      var currencyIdx = refHeaders.indexOf("Курс валюты, €");
      var markupIdx = refHeaders.indexOf("% наценка для РРЦ");
      var yearIdx = refHeaders.indexOf("Год");

      if (yearIdx !== -1 && currencyIdx !== -1) {
        // Используем VLOOKUP для поиска курса по году
        currencyRef = 'VLOOKUP(' + currentYear + ', INDIRECT("Справочник!R2C' + (yearIdx + 1) + ':R' + referenceSheet.getLastRow() + 'C' + (currencyIdx + 1) + '", FALSE), ' + (currencyIdx - yearIdx + 1) + ', FALSE)';
      }

      if (markupIdx !== -1) {
        markupRef = "'Справочник'!R2C" + (markupIdx + 1);
      }
    }

    // Применяем формулы к каждому столбцу
    _applyCalculationFormulasByColumn(sheet, cols, dataRows, currencyRef, markupRef);

    // Применяем форматирование
    _applyCalculationFormatting(sheet, cols, dataRows);

    // Применяем границу перед столбцом "ЦЕНА EXW из Б/З, €"
    if (cols.exwFromBudget !== -1) {
      var totalRows = lastRow - HEADER_ROW_INDEX + 1;
      var borderRange = sheet.getRange(HEADER_ROW_INDEX, cols.exwFromBudget + 1, totalRows, 1);
      borderRange.setBorder(
        null, // top
        true, // left - жирная граница
        null, // bottom
        null, // right
        null, // vertical
        null, // horizontal
        null, // color
        SpreadsheetApp.BorderStyle.SOLID_THICK
      );
    }

    if (Lib.logInfo) {
      Lib.logInfo(MODULE_TAG + " Расчетные формулы применены к листу \"" + sheet.getName() + "\"");
    }
  }

  function _applyCalculationFormulasByColumn(sheet, cols, dataRows, currencyRef, markupRef) {
    // 1. Расчетная цена Опт, € = EXW текущая, € * К-т
    if (cols.priceWholesaleEur !== -1 && cols.exwCurrent !== -1 && cols.coefficient !== -1) {
      var offset1 = cols.exwCurrent - cols.priceWholesaleEur;
      var offset2 = cols.coefficient - cols.priceWholesaleEur;
      var formula = '=IF(LEN(R[0]C[' + offset1 + '])=0, "", R[0]C[' + offset1 + '] * R[0]C[' + offset2 + '])';
      _setFormulasR1C1(sheet.getRange(HEADER_ROW_INDEX + 1, cols.priceWholesaleEur + 1, dataRows, 1), formula);
    }

    // 2. Расчетная цена Опт, ₽ = Расчетная цена Опт, € * Курс валюты
    if (cols.priceWholesaleRub !== -1 && cols.priceWholesaleEur !== -1) {
      var offset = cols.priceWholesaleEur - cols.priceWholesaleRub;
      var formula = '=IF(LEN(R[0]C[' + offset + '])=0, "", ROUND(R[0]C[' + offset + '] * ' + currencyRef + ', 0))';
      _setFormulasR1C1(sheet.getRange(HEADER_ROW_INDEX + 1, cols.priceWholesaleRub + 1, dataRows, 1), formula);
    }

    // 3. стоимость за 1, мл = Расчетная цена Опт, ₽ / Объём
    if (cols.costPerMl !== -1 && cols.priceWholesaleRub !== -1 && cols.volume !== -1) {
      var offsetPrice = cols.priceWholesaleRub - cols.costPerMl;
      var offsetVolume = cols.volume - cols.costPerMl;
      var volumeFormula = _parseVolumeFormula(offsetVolume);
      var formula = '=IF(LEN(R[0]C[' + offsetPrice + '])=0, "", ROUND(R[0]C[' + offsetPrice + '] / (' + volumeFormula + '), 0))';
      _setFormulasR1C1(sheet.getRange(HEADER_ROW_INDEX + 1, cols.costPerMl + 1, dataRows, 1), formula);
    }

    // 4. % наценка для РРЦ - из справочника, если есть EXW текущая
    if (cols.markupPercent !== -1 && cols.exwCurrent !== -1) {
      var offset = cols.exwCurrent - cols.markupPercent;
      var formula = '=IF(LEN(R[0]C[' + offset + '])=0, "", ' + markupRef + ' / 100)';
      _setFormulasR1C1(sheet.getRange(HEADER_ROW_INDEX + 1, cols.markupPercent + 1, dataRows, 1), formula);
    }

    // 5. Расчетная цена РРЦ, ₽ = Расчетная цена Опт, ₽ + (Расчетная цена Опт, ₽ * % наценка)
    if (cols.priceRetailCalc !== -1 && cols.priceWholesaleRub !== -1 && cols.markupPercent !== -1) {
      var offsetPrice = cols.priceWholesaleRub - cols.priceRetailCalc;
      var offsetMarkup = cols.markupPercent - cols.priceRetailCalc;
      var formula = '=IF(LEN(R[0]C[' + offsetPrice + '])=0, "", ROUND(R[0]C[' + offsetPrice + '] + (R[0]C[' + offsetPrice + '] * R[0]C[' + offsetMarkup + ']), 0))';
      _setFormulasR1C1(sheet.getRange(HEADER_ROW_INDEX + 1, cols.priceRetailCalc + 1, dataRows, 1), formula);
    }

    // 6. Утвержденная цена Опт, ₽ = ROUND(Расчетная цена Опт, ₽, -2) - округление до 100
    if (cols.priceWholesaleApproved !== -1 && cols.priceWholesaleRub !== -1) {
      var offset = cols.priceWholesaleRub - cols.priceWholesaleApproved;
      var formula = '=IF(LEN(R[0]C[' + offset + '])=0, "", ROUND(R[0]C[' + offset + '], -2))';
      _setFormulasR1C1(sheet.getRange(HEADER_ROW_INDEX + 1, cols.priceWholesaleApproved + 1, dataRows, 1), formula);
    }

    // 7. Стоимость за 1 мл, ₽ = Утвержденная цена Опт, ₽ / Объём
    if (cols.costPerMlApproved !== -1 && cols.priceWholesaleApproved !== -1 && cols.volume !== -1) {
      var offsetPrice = cols.priceWholesaleApproved - cols.costPerMlApproved;
      var offsetVolume = cols.volume - cols.costPerMlApproved;
      var volumeFormula = _parseVolumeFormula(offsetVolume);
      var formula = '=IF(LEN(R[0]C[' + offsetPrice + '])=0, "", ROUND(R[0]C[' + offsetPrice + '] / (' + volumeFormula + '), 0))';
      _setFormulasR1C1(sheet.getRange(HEADER_ROW_INDEX + 1, cols.costPerMlApproved + 1, dataRows, 1), formula);
    }

    // 8. К-т ФАКТ = (Утвержденная цена Опт, ₽ / Курс валюты) / EXW текущая, €
    if (cols.coefficientFact !== -1 && cols.priceWholesaleApproved !== -1 && cols.exwCurrent !== -1) {
      var offsetPrice = cols.priceWholesaleApproved - cols.coefficientFact;
      var offsetExw = cols.exwCurrent - cols.coefficientFact;
      var formula = '=IF(OR(LEN(R[0]C[' + offsetPrice + '])=0, LEN(R[0]C[' + offsetExw + '])=0), "", ROUND((R[0]C[' + offsetPrice + '] / (' + currencyRef + ')) / R[0]C[' + offsetExw + '], 1))';
      _setFormulasR1C1(sheet.getRange(HEADER_ROW_INDEX + 1, cols.coefficientFact + 1, dataRows, 1), formula);
    }

    // 9. Утвержденная РРЦ, ₽ = ROUND(Утвержденная цена Опт, ₽ * (1 + % наценка), -1) - округление до 10
    if (cols.priceRetailApproved !== -1 && cols.priceWholesaleApproved !== -1 && cols.markupPercent !== -1) {
      var offsetPrice = cols.priceWholesaleApproved - cols.priceRetailApproved;
      var offsetMarkup = cols.markupPercent - cols.priceRetailApproved;
      var formula = '=IF(LEN(R[0]C[' + offsetPrice + '])=0, "", ROUND(R[0]C[' + offsetPrice + '] * (1 + R[0]C[' + offsetMarkup + ']), -1))';
      _setFormulasR1C1(sheet.getRange(HEADER_ROW_INDEX + 1, cols.priceRetailApproved + 1, dataRows, 1), formula);
    }

    // 10. Утвержденая Е-комерс, ₽ = Утвержденная цена Опт, ₽ * 2 (100% наценка)
    if (cols.priceEcommerce !== -1 && cols.priceWholesaleApproved !== -1) {
      var offset = cols.priceWholesaleApproved - cols.priceEcommerce;
      var formula = '=IF(LEN(R[0]C[' + offset + '])=0, "", R[0]C[' + offset + '] * 2)';
      _setFormulasR1C1(sheet.getRange(HEADER_ROW_INDEX + 1, cols.priceEcommerce + 1, dataRows, 1), formula);
    }

    // 11. Цена дистрибьютора -30%, ₽ = Утвержденная цена Опт, ₽ * 0.7
    if (cols.priceDistributor !== -1 && cols.priceWholesaleApproved !== -1) {
      var offset = cols.priceDistributor - cols.priceWholesaleApproved;
      var formula = '=IF(LEN(INDIRECT(ADDRESS(ROW(); COLUMN()-' + offset + ')))=0; ""; ROUND(INDIRECT(ADDRESS(ROW(); COLUMN()-' + offset + '))*0,7; 0))';
      var formulas = [];
      for (var i = 0; i < dataRows; i++) {
        formulas.push([formula]);
      }
      sheet.getRange(HEADER_ROW_INDEX + 1, cols.priceDistributor + 1, dataRows, 1).setFormulas(formulas);
    }

    // 12. К-т дистр. = (Цена дистрибьютора / Курс валюты) / EXW текущая, €
    if (cols.coefficientDistributor !== -1 && cols.priceDistributor !== -1 && cols.exwCurrent !== -1) {
      var offsetPrice = cols.priceDistributor - cols.coefficientDistributor;
      var offsetExw = cols.exwCurrent - cols.coefficientDistributor;
      var formula = '=IF(OR(LEN(R[0]C[' + offsetPrice + '])=0, LEN(R[0]C[' + offsetExw + '])=0), "", ROUND((R[0]C[' + offsetPrice + '] / (' + currencyRef + ')) / R[0]C[' + offsetExw + '], 1))';
      _setFormulasR1C1(sheet.getRange(HEADER_ROW_INDEX + 1, cols.coefficientDistributor + 1, dataRows, 1), formula);
    }

    // 13. Цена крупный опт -10%, ₽ = Утвержденная цена Опт, ₽ * 0.9
    if (cols.priceWholesaleLarge !== -1 && cols.priceWholesaleApproved !== -1) {
      var offset = cols.priceWholesaleLarge - cols.priceWholesaleApproved;
      var formula = '=IF(LEN(INDIRECT(ADDRESS(ROW(); COLUMN()-' + offset + ')))=0; ""; ROUND(INDIRECT(ADDRESS(ROW(); COLUMN()-' + offset + '))*0,9; 0))';
      var formulas = [];
      for (var i = 0; i < dataRows; i++) {
        formulas.push([formula]);
      }
      sheet.getRange(HEADER_ROW_INDEX + 1, cols.priceWholesaleLarge + 1, dataRows, 1).setFormulas(formulas);
    }

    // 14. К-т круп. опт. = (Цена крупный опт / Курс валюты) / EXW текущая, €
    if (cols.coefficientLarge !== -1 && cols.priceWholesaleLarge !== -1 && cols.exwCurrent !== -1) {
      var offsetPrice = cols.priceWholesaleLarge - cols.coefficientLarge;
      var offsetExw = cols.exwCurrent - cols.coefficientLarge;
      var formula = '=IF(OR(LEN(R[0]C[' + offsetPrice + '])=0, LEN(R[0]C[' + offsetExw + '])=0), "", ROUND((R[0]C[' + offsetPrice + '] / (' + currencyRef + ')) / R[0]C[' + offsetExw + '], 1))';
      _setFormulasR1C1(sheet.getRange(HEADER_ROW_INDEX + 1, cols.coefficientLarge + 1, dataRows, 1), formula);
    }

    // 15. Max скидка-50%, ₽ = Утвержденная цена Опт, ₽ * 0.5
    if (cols.priceMaxDiscount !== -1 && cols.priceWholesaleApproved !== -1) {
      var offset = cols.priceMaxDiscount - cols.priceWholesaleApproved;
      var formula = '=IF(LEN(INDIRECT(ADDRESS(ROW(); COLUMN()-' + offset + ')))=0; ""; ROUND(INDIRECT(ADDRESS(ROW(); COLUMN()-' + offset + '))*0,5; 0))';
      var formulas = [];
      for (var i = 0; i < dataRows; i++) {
        formulas.push([formula]);
      }
      sheet.getRange(HEADER_ROW_INDEX + 1, cols.priceMaxDiscount + 1, dataRows, 1).setFormulas(formulas);
    }

    // 16. К-т мах скидка = (Max скидка / Курс валюты) / EXW текущая, €
    if (cols.coefficientMaxDiscount !== -1 && cols.priceMaxDiscount !== -1 && cols.exwCurrent !== -1) {
      var offsetPrice = cols.priceMaxDiscount - cols.coefficientMaxDiscount;
      var offsetExw = cols.exwCurrent - cols.coefficientMaxDiscount;
      var formula = '=IF(OR(LEN(R[0]C[' + offsetPrice + '])=0, LEN(R[0]C[' + offsetExw + '])=0), "", ROUND((R[0]C[' + offsetPrice + '] / (' + currencyRef + ')) / R[0]C[' + offsetExw + '], 1))';
      _setFormulasR1C1(sheet.getRange(HEADER_ROW_INDEX + 1, cols.coefficientMaxDiscount + 1, dataRows, 1), formula);
    }
  }

  function _applyCalculationFormatting(sheet, cols, dataRows) {
    // Форматирование столбцов с рублями
    var rubColumns = [
      cols.priceWholesaleRub,
      cols.priceRetailCalc,
      cols.priceWholesaleApproved,
      cols.priceRetailApproved,
      cols.priceEcommerce,
      cols.priceDistributor,
      cols.priceWholesaleLarge,
      cols.priceMaxDiscount,
      cols.costPerMl,
      cols.costPerMlApproved
    ];

    rubColumns.forEach(function (colIndex) {
      if (colIndex !== -1) {
        var range = sheet.getRange(HEADER_ROW_INDEX + 1, colIndex + 1, dataRows, 1);
        _safeApplyNumberFormat(range, {
          format: '#,##0 "₽"',
          align: "center",
          columnName: "₽ столбец индекс " + colIndex
        });
      }
    });

    // Форматирование столбцов с евро
    var eurColumns = [cols.priceWholesaleEur, cols.exwFromBudget];
    eurColumns.forEach(function (colIndex) {
      if (colIndex !== -1) {
        var range = sheet.getRange(HEADER_ROW_INDEX + 1, colIndex + 1, dataRows, 1);
        _safeApplyNumberFormat(range, {
          format: '0.00 "€"',
          align: "center",
          columnName: "€ столбец индекс " + colIndex
        });
      }
    });

    // Форматирование процентов
    if (cols.markupPercent !== -1) {
      var markupRange = sheet.getRange(HEADER_ROW_INDEX + 1, cols.markupPercent + 1, dataRows, 1);
      _safeApplyNumberFormat(markupRange, {
        format: "0%",
        align: "center",
        columnName: "% наценка для РРЦ"
      });
    }

    // Форматирование коэффициентов
    var coeffColumns = [
      cols.coefficient,
      cols.coefficientFact,
      cols.coefficientDistributor,
      cols.coefficientLarge,
      cols.coefficientMaxDiscount
    ];
    coeffColumns.forEach(function (colIndex) {
      if (colIndex !== -1) {
        var range = sheet.getRange(HEADER_ROW_INDEX + 1, colIndex + 1, dataRows, 1);
        _safeApplyNumberFormat(range, {
          format: "0.0",
          align: "center",
          columnName: "Коэффициент индекс " + colIndex
        });
      }
    });
  }

  /**
   * Применить форматирование к листу "Динамика цены"
   */
  function _applyPriceDynamicsFormatting(sheet, headers, lastRow) {
    if (lastRow <= HEADER_ROW_INDEX) {
      return;
    }

    var dataRows = lastRow - HEADER_ROW_INDEX;

    for (var i = 0; i < headers.length; i++) {
      var header = headers[i];

      var range = sheet.getRange(HEADER_ROW_INDEX + 1, i + 1, dataRows, 1);

      // Столбцы с евро - формат 0.00 "€"
      if (header.indexOf(", €") !== -1) {
        _safeApplyNumberFormat(range, {
          format: '0.00 "€"',
          align: "center",
          columnName: header
        });
      }
      // Столбцы с рублями - формат #,##0 "₽"
      else if (header.indexOf(", ₽") !== -1) {
        _safeApplyNumberFormat(range, {
          format: '#,##0 "₽"',
          align: "center",
          columnName: header
        });
      }
      // Столбцы с процентами
      else if (header.indexOf(", %") !== -1 || header.indexOf("Прирост") !== -1) {
        _safeApplyNumberFormat(range, {
          format: "0.0%",
          align: "center",
          columnName: header
        });
      }
    }
  }

  function _safeRunSheetRefresh(fn, message) {
    try {
      if (Lib && typeof fn === "function") {
        fn();
      }
    } catch (hookError) {
      if (Lib && Lib.logWarn) {
        Lib.logWarn(message, hookError);
      }
    }
  }

  function _ensureSheetRefreshHooksOnOpen() {
    if (!global || !global.Lib || global.Lib.__autoPriceSheetHooksInitialized) {
      return;
    }

    var originalOnOpen = typeof global.Lib.onOpen === "function" ? global.Lib.onOpen : null;
    global.Lib.onOpen = function () {
      try {
        if (originalOnOpen) {
          originalOnOpen.apply(this, arguments);
        }
      } finally {
        _safeRunSheetRefresh(
          Lib && Lib.refreshOrderSheetExwAlfaspaColumn,
          MODULE_TAG + " Ошибка автозапуска EXW ALFASPA при onOpen"
        );
        _safeRunSheetRefresh(
          Lib && Lib.refreshPriceSheetCommercialColumns,
          MODULE_TAG + " Ошибка автозапуска цен листа Прайс при onOpen"
        );
        _safeRunSheetRefresh(
          function() {
            if (Lib && Lib.applyCalculationFormulas) {
              Lib.applyCalculationFormulas(true); // silent=true для автозапуска
            }
          },
          MODULE_TAG + " Ошибка автозапуска формул листа Расчет цены при onOpen"
        );
      }
    };

    global.Lib.__autoPriceExwOrderHooked = true; // сохраняем старый флаг для совместимости
  }

  // _ensureSheetRefreshHooksOnOpen(); // ОТКЛЮЧЕНО: Миграция на Python (load-functions)

  // =======================================================================================
  // ГЛОБАЛЬНЫЕ ПРОКСИ-ФУНКЦИИ ДЛЯ МЕНЮ
  // =======================================================================================

  /**
   * Прокси-функция для вызова из меню: добавить новый год на лист "Динамика цены"
   */
  global.addNewYearColumnsToPriceDynamics_proxy = function () {
    if (Lib.addNewYearColumnsToPriceDynamics) {
      Lib.addNewYearColumnsToPriceDynamics();
    }
  };

  /**
   * Прокси-функция для вызова из меню: создать новый год для динамики на листе "Расчет цены"
   */
  global.createNewYearForDynamics_proxy = function () {
    if (Lib.createNewYearForDynamics) {
      Lib.createNewYearForDynamics();
    }
  };

  /**
   * Прокси-функция для вызова из меню: пересчитать формулы "Динамика цены"
   */
  global.recalculatePriceDynamicsFormulas_proxy = function () {
    if (Lib.recalculatePriceDynamicsFormulas) {
      Lib.recalculatePriceDynamicsFormulas();
    }
  };

  /**
   * Прокси-функция для вызова из меню: обновить формулы "Расчет цены"
   */
  global.updatePriceCalculationFormulas_proxy = function () {
    if (Lib.updatePriceCalculationFormulas) {
      Lib.updatePriceCalculationFormulas();
    }
  };

  /**
   * Прокси-функция для вызова из меню: применить расчетные формулы
   */
  global.applyCalculationFormulas_proxy = function () {
    if (Lib.applyCalculationFormulas) {
      Lib.applyCalculationFormulas();
    }
  };

})(Lib, this);
