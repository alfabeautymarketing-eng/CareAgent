/** =============================================================================
 * E_InvoiceLogic.gs — Инвойсы (кнопка: "Форматировать лист 'Ордер'")
 * =============================================================================
 */
var Lib = Lib || {};

(function (Lib, global) {
  const SHEET_NAME = (Lib.CONFIG && Lib.CONFIG.SHEETS && Lib.CONFIG.SHEETS.ORDER) || 'Ордер';

  // Заголовки, по которым работаем (строго по ИМЕНАМ, не по индексам!)
  const H = {
    QTY:        'Кол-во поставка',
    PRICE:      'Цена ед.',
    DISCOUNT:   'Скидка, %',
    TOTAL:      'Итого цена',
    EXPIRY:     'Годен до',
  };

  /** ПУБЛИЧНАЯ: форматировать лист "Ордер" */
  Lib.formatOrderSheet = function () {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_NAME);

    if (!sh) {
      ui.alert(`Лист "${SHEET_NAME}" не найден.`);
      return;
    }

    try {
      ss.toast(`Форматирование листа "${SHEET_NAME}"…`, 'Выполнение', 2);
      // 1) Эксклюзивная очистка под поставщика
      _exclusiveCleanup_Order_(sh);

      // 2) Проверка данных
      let lastRow = sh.getLastRow();
      const lastCol = sh.getLastColumn();
      if (lastRow <= 1 || lastCol < 1) {
        ui.alert(`Лист "${SHEET_NAME}" пуст после очистки.`);
        return;
      }

      // 3) Карту заголовков строим по факту
      const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v || '').trim());
      const col = (name) => {
        const i0 = headers.indexOf(name);
        return i0 === -1 ? 0 : (i0 + 1);
      };

      // 4) Нормализация числовых колонок
      const moneyFmt = (Lib.CONFIG && Lib.CONFIG.VALUES && Lib.CONFIG.VALUES.DEFAULT_NUMBER_FORMAT) || '#,##0.00';
      _normalizeNumberColumn_(sh, col(H.PRICE),   lastRow, moneyFmt, /*isPercent=*/false);
      _normalizeNumberColumn_(sh, col(H.TOTAL),   lastRow, moneyFmt, /*isPercent=*/false);
      _normalizeNumberColumn_(sh, col(H.QTY),     lastRow, '#,##0',   /*isPercent=*/false);
      _normalizeNumberColumn_(sh, col(H.DISCOUNT),lastRow, '0.00%',   /*isPercent=*/true);

      // 5) Выравнивание по центру для ключевых колонок
      [H.QTY, H.PRICE, H.DISCOUNT, H.TOTAL, H.EXPIRY].forEach(h => {
        const c = col(h);
        if (c) sh.getRange(2, c, lastRow - 1, 1).setHorizontalAlignment('center');
      });

      // 6) Автоширина + зачистка "хвостов"
      sh.autoResizeColumns(1, lastCol);
      lastRow = sh.getLastRow();
      const maxRows = sh.getMaxRows();
      if (maxRows > lastRow) sh.deleteRows(lastRow + 1, maxRows - lastRow);
      const maxCols = sh.getMaxColumns();
      if (maxCols > lastCol) sh.deleteColumns(lastCol + 1, maxCols - lastCol);

      ss.toast('Форматирование успешно завершено!', 'Готово', 5);
      Lib.logInfo(`[ORDER] Форматирование завершено на листе "${SHEET_NAME}"`);
    } catch (e) {
      Lib.logError('formatOrderSheet: критическая ошибка', e);
      ui.alert(`Ошибка при форматировании: ${e.message}`);
      ss.toast('Ошибка форматирования', 'Ошибка', 3);
    } finally {
      // Гарантированно закрываем любой висящий toast
      ss.toast('', '', 1);
    }
  };

  // ------------------------ ВНУТРЕННИЕ ПОМОЩНИКИ ------------------------

  /** Эксклюзивная очистка данных под конкретного поставщика */
  function _exclusiveCleanup_Order_(sh) {
    const PAGE_STRING_TO_DELETE = 'Страница';
    const QTY_SUFFIX_TO_REMOVE = 'UN';

    // Удаляем строки, содержащие "Страница"
    try {
      const hits = sh.createTextFinder(PAGE_STRING_TO_DELETE).matchCase(false).findAll() || [];
      if (hits.length) {
        hits.slice().reverse().forEach(r => sh.deleteRow(r.getRow()));
      }
    } catch(_) {}

    // Удаляем суффикс "UN" из "Кол-во поставка"
    const lastRow = sh.getLastRow();
    if (lastRow <= 1) return;

    const lastCol = sh.getLastColumn();
    const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v || '').trim());
    const qtyCol = headers.indexOf(H.QTY) + 1;
    if (qtyCol > 0) {
      const rng = sh.getRange(2, qtyCol, lastRow - 1, 1);
      const disp = rng.getDisplayValues();
      const cleaned = disp.map(([v]) => [String(v || '').replace(new RegExp(`\\s*${QTY_SUFFIX_TO_REMOVE}\\s*`, 'i'), '').trim()]);
      rng.setValues(cleaned);
    }
  }

  /**
   * Нормализация числовой колонки:
   * - парсит числа из любых строковых представлений (с точкой/запятой, пробелами, %);
   * - для процентов приводит к 0..1 и задаёт формат '0.00%';
   * - для денег/количества задаёт формат из аргумента.
   */
  function _normalizeNumberColumn_(sh, col, lastRow, numberFormat, isPercent) {
    if (!col || lastRow <= 1) return;

    const rng = sh.getRange(2, col, lastRow - 1, 1);
    const raw = rng.getValues();
    const out = new Array(raw.length);

    for (let i = 0; i < raw.length; i++) {
      let v = raw[i][0];

      if (v === '' || v === null || v === undefined) {
        out[i] = [''];
        continue;
      }

      if (typeof v === 'number') {
        // Уже число
        let n = v;
        if (isPercent) n = n > 1 ? n / 100 : n; // если пришло 5 → 5%
        out[i] = [n];
        continue;
      }

      // Строка: убираем пробелы/неразрывные, выносим %, нормализуем разделители
      let s = String(v).trim();
      const hadPercent = /%$/.test(s);
      s = s.replace(/\u00A0/g, '').replace(/\s/g, '').replace('%', '');

      // Если есть и точка и запятая → считаем, что запятые — тысячные
      if (s.includes(',') && s.includes('.')) s = s.replace(/,/g, '');
      else s = s.replace(',', '.');

      const parsed = parseFloat(s);
      if (isNaN(parsed)) {
        out[i] = [''];
        continue;
      }

      let n = parsed;
      if (isPercent) n = (hadPercent || parsed > 1) ? parsed / 100 : parsed;
      out[i] = [n];
    }

    rng.setValues(out).setNumberFormat(numberFormat);
  }

})(Lib, this);

/** =============================================================================
 * E_InvoiceLogic.gs — Инвойсы
 * Кнопка: "1. Создать лист 'Для инвойса'"
 * =============================================================================
 */
var Lib = Lib || {};

(function (Lib, global) {
  const S = (Lib.CONFIG && Lib.CONFIG.SHEETS) || {};
  const H_ALL = (Lib.CONFIG && Lib.CONFIG.HEADERS) || {};
  const TGT_NAME = S.INVOICE_FULL || 'Для инвойса';

  // Имена колонок (берём из эталонных шапок конфигурации)
  const H_ORDER = H_ALL[S.ORDER] || ['Арт. произв.', 'Название ENG', 'Кол-во поставка', 'Цена ед.', 'Скидка, %', 'Итого цена', 'Годен до'];
  const H_CERT  = H_ALL[S.CERTIFICATION] || [];
  const H_LABEL = H_ALL[S.LABELS] || [];
  const H_INV   = H_ALL[S.INVOICE_FULL] || ['ID', 'Арт. Рус','Арт. произв.','Наименование','Наименование Англ','кол-во','Цена ед.','Сумма', 'ДС','ДС до','номер ДС','Спирт','Спирт до','статус','Примечание','Назначение Этикетка','Применение Этикетка','Активные ингредиенты без %'];

  // Ключевые заголовки в исходных листах
  const COLS_ORDER = {
    PRODUCER_ART: 'Арт. произв.',
    QTY:          'Кол-во поставка',
    PRICE:        'Цена ед.',
    TOTAL:        'Итого цена',
  };

  const COLS_CERT = {
    PRODUCER_ART: 'Арт. произв.',
    RUS_ART:      'Арт. Рус',
    STATUS:       'Статус',
    INVOICE_NAME:       'Наименование для инвойса',
    INVOICE_NAME_ENG:   'Наименование для инвойса Англ',
    DS_UNTIL:     'ДС до',
    SPIRIT_UNTIL: 'Спирт До',
    DS_SCAN:      'ДС',
    SPIRIT_SCAN:  'Скан спирт',
    ID:           'ID',
    DS_NUMBER:    'номер ДС',
  };

  const COLS_LABELS = {
    RUS_ART:      'Арт. Рус',
    PURPOSE:      'Назначение Этикетка',
    APPLICATION:  'Применение Этикетка',
    INGREDIENTS:  'Активные ингредиенты без %',
  };

  // Заголовки целевого листа (жёстко равны конфигу)
  const COLS_INV = {
    ID:           'ID',
    RUS_ART:      'Арт. Рус',
    PRODUCER_ART: 'Арт. произв.',
    NAME_RU:      'Наименование',
    NAME_EN:      'Наименование Англ',
    QTY:          'кол-во',
    PRICE:        'Цена ед.',
    TOTAL:        'Сумма',
    DS_LINK:      'ДС',
    DS_UNTIL:     'ДС до',
    DS_NUMBER:    'номер ДС',
    SPIRIT_LINK:  'Спирт',
    SPIRIT_UNTIL: 'Спирт до',
    STATUS:       'статус',
    NOTE:         'Примечание',
    PURPOSE:      'Назначение Этикетка',
    APPLICATION:  'Применение Этикетка',
    INGREDIENTS:  'Активные ингредиенты без %',
  };

  const norm = v => String(v || '').toLowerCase().trim();

  /** ПУБЛИЧНАЯ: создать лист "Для инвойса" заново */
  Lib.createFullInvoice = function () {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    try {
      ss.toast('Этап 1/5: Подготовка…', 'Выполнение', 2);

      const shOrder = ss.getSheetByName(S.ORDER);
      const shCert  = ss.getSheetByName(S.CERTIFICATION);
      const shLabel = ss.getSheetByName(S.LABELS);
      if (!shOrder || !shCert || !shLabel) {
        throw new Error(`Не найден один из листов: "${S.ORDER}", "${S.CERTIFICATION}" или "${S.LABELS}".`);
      }

      // Пересоздаём целевой лист
      let shTarget = ss.getSheetByName(TGT_NAME);
      if (shTarget) ss.deleteSheet(shTarget);
      shTarget = ss.insertSheet(TGT_NAME);

      // Пишем шапку как в конфиге
      shTarget.getRange(1, 1, 1, H_INV.length).setValues([H_INV]).setFontWeight('bold').setWrap(true);
      shTarget.setFrozenRows(1);

      // --- 2) СПРАВОЧНИКИ --------------------------------------------------
      ss.toast('Этап 2/5: Анализ справочников…', 'Выполнение', 2);

      // Сертификация: берём и значения и richText (для ссылок)
      const certDataRange = shCert.getDataRange();
      const certVals = certDataRange.getValues();
      const certRT   = certDataRange.getRichTextValues();

      const certHdr = (certVals.shift() || []).map(x => String(x || '').trim());
      certRT.shift();

      const idxC = headerIndex(certHdr, COLS_CERT, 'Сертификация');
      const certLookup = new Map();
      certVals.forEach((row, i) => {
        const key = norm(row[idxC.PRODUCER_ART]);
        if (!key) return;
        certLookup.set(key, {
          rusArt:     row[idxC.RUS_ART],
          status:     row[idxC.STATUS],
          name:       row[idxC.INVOICE_NAME],
          nameEng:    row[idxC.INVOICE_NAME_ENG],
          dsUntil:    row[idxC.DS_UNTIL],
          spiritUntil:row[idxC.SPIRIT_UNTIL],
          dsLink:     (certRT[i][idxC.DS_SCAN] || null),
          spiritLink: (certRT[i][idxC.SPIRIT_SCAN] || null),
          id:          idxC.ID !== -1 ? row[idxC.ID] : '',
          dsNumber:    idxC.DS_NUMBER !== -1 ? row[idxC.DS_NUMBER] : '',
        });
      });

      // Этикетки
      const lblVals = shLabel.getDataRange().getValues();
      const lblHdr  = (lblVals.shift() || []).map(x => String(x || '').trim());
      const idxL = headerIndex(lblHdr, COLS_LABELS, 'Этикетки');
      const labelLookup = new Map();
      lblVals.forEach(row => {
        const key = norm(row[idxL.RUS_ART]);
        if (!key) return;
        labelLookup.set(key, {
          purpose:     row[idxL.PURPOSE],
          application: row[idxL.APPLICATION],
          ingredients: row[idxL.INGREDIENTS],
        });
      });

      // --- 3) ОБРАБОТКА ОРДЕРА --------------------------------------------
      ss.toast('Этап 3/5: Обработка "Ордера"…', 'Выполнение', 2);

      const rngOrd   = shOrder.getDataRange();
      const ordVals  = rngOrd.getValues();
      const ordFmt   = rngOrd.getNumberFormats();

      const ordHdr   = (ordVals.shift() || []).map(x => String(x || '').trim());
      ordFmt.shift();

      const idxO = headerIndex(ordHdr, COLS_ORDER, 'Ордер');

      // Построим быстрый индекс столбцов целевого листа
      const invHdr = H_INV.slice();
      const idxInv = {};
      invHdr.forEach((h, i) => idxInv[h] = i);

      const resultValues   = [];
      const resultRichText = [];
      const resultFormats  = [];

      ordVals.forEach((row, rIdx0) => {
        const prodArtOrig = row[idxO.PRODUCER_ART];
        const prodArtKey  = norm(prodArtOrig);
        if (!prodArtKey) return;

        const cert = certLookup.get(prodArtKey);

        const outRow   = new Array(invHdr.length).fill('');
        const outRTr   = new Array(invHdr.length).fill(null);
        const outFmt   = new Array(invHdr.length).fill(null);

        // 1) Из Сертификации
        if (cert) {
          outRow[idxInv[COLS_INV.RUS_ART]]   = cert.rusArt || '';
          outRow[idxInv[COLS_INV.NAME_RU]]   = cert.name || '';
          outRow[idxInv[COLS_INV.NAME_EN]]   = cert.nameEng || '';
          outRow[idxInv[COLS_INV.STATUS]]    = cert.status || '';
          outRow[idxInv[COLS_INV.DS_UNTIL]]  = cert.dsUntil || '';
          outRow[idxInv[COLS_INV.SPIRIT_UNTIL]] = cert.spiritUntil || '';
          outRow[idxInv[COLS_INV.ID]]        = cert ? (cert.id || '') : '';
          outRow[idxInv[COLS_INV.DS_NUMBER]] = cert ? (cert.dsNumber || '') : '';
          if (cert.dsLink)     outRTr[idxInv[COLS_INV.DS_LINK]]     = cert.dsLink;
          if (cert.spiritLink) outRTr[idxInv[COLS_INV.SPIRIT_LINK]] = cert.spiritLink;
        } else {
          // если из сертификации не нашли — хотя бы пометим названием
          outRow[idxInv[COLS_INV.NAME_RU]] = Lib.CONFIG.VALUES.NOT_FOUND_TEXT || '!!! НЕ НАЙДЕНО !!!';
        }

        // 2) Из Ордера (исходные числовые форматы сохраняем!)
        outRow[idxInv[COLS_INV.PRODUCER_ART]] = prodArtOrig;
        outRow[idxInv[COLS_INV.QTY]]   = row[idxO.QTY];
        outRow[idxInv[COLS_INV.PRICE]] = row[idxO.PRICE];
        outRow[idxInv[COLS_INV.TOTAL]] = row[idxO.TOTAL];

        outFmt[idxInv[COLS_INV.PRICE]] = ordFmt[rIdx0][idxO.PRICE];
        outFmt[idxInv[COLS_INV.TOTAL]] = ordFmt[rIdx0][idxO.TOTAL];

        // 3) Этикетки — ключ по русскому артикулу (если он есть)
        const rusArtKey = norm(outRow[idxInv[COLS_INV.RUS_ART]]);
        if (rusArtKey) {
          const lbl = labelLookup.get(rusArtKey);
          if (lbl) {
            outRow[idxInv[COLS_INV.PURPOSE]]     = lbl.purpose;
            outRow[idxInv[COLS_INV.APPLICATION]] = lbl.application;
            outRow[idxInv[COLS_INV.INGREDIENTS]] = lbl.ingredients;
          }
        }

        resultValues.push(outRow);
        resultRichText.push(outRTr);
        resultFormats.push(outFmt);
      });

      // --- 4) Запись -------------------------------------------------------
      ss.toast('Этап 4/5: Запись данных…', 'Завершение', 2);

      if (resultValues.length) {
        const dataRng = shTarget.getRange(2, 1, resultValues.length, invHdr.length);
        dataRng.setValues(resultValues);

        // RichText (ссылки ДС/Спирт)
        resultRichText.forEach((rowRT, i) => {
          rowRT.forEach((rt, j) => { if (rt) shTarget.getRange(i + 2, j + 1).setRichTextValue(rt); });
        });

        SpreadsheetApp.flush();

        // Применяем числовые форматы цены/суммы на уровне строк
        const priceCol  = idxInv[COLS_INV.PRICE] + 1;
        const totalCol  = idxInv[COLS_INV.TOTAL] + 1;
        resultFormats.forEach((fmtRow, i) => {
          const r = i + 2;
          const pf = fmtRow[priceCol - 1] || (Lib.CONFIG.VALUES.DEFAULT_NUMBER_FORMAT || '#,##0.00');
          const tf = fmtRow[totalCol - 1] || (Lib.CONFIG.VALUES.DEFAULT_NUMBER_FORMAT || '#,##0.00');
          shTarget.getRange(r, priceCol).setNumberFormat(pf);
          shTarget.getRange(r, totalCol).setNumberFormat(tf);
        });

        // Границы, ширины, выравнивание, даты
        const lastRow = shTarget.getLastRow();
        shTarget.getRange(1, 1, lastRow, invHdr.length)
                .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
        shTarget.autoResizeColumns(1, invHdr.length);

        const dsUntilCol     = idxInv[COLS_INV.DS_UNTIL] + 1;
        const spiritUntilCol = idxInv[COLS_INV.SPIRIT_UNTIL] + 1;
        if (dsUntilCol > 0)     shTarget.getRange(2, dsUntilCol, resultValues.length, 1).setNumberFormat('dd.MM.yyyy');
        if (spiritUntilCol > 0) shTarget.getRange(2, spiritUntilCol, resultValues.length, 1).setNumberFormat('dd.MM.yyyy');

        const nameRuCol  = idxInv[COLS_INV.NAME_RU] + 1;
        const nameEnCol  = idxInv[COLS_INV.NAME_EN] + 1;
        if (nameRuCol > 0) shTarget.setColumnWidth(nameRuCol, 350), shTarget.getRange(2, nameRuCol, resultValues.length, 1).setWrap(true);
        if (nameEnCol > 0) shTarget.setColumnWidth(nameEnCol, 350), shTarget.getRange(2, nameEnCol, resultValues.length, 1).setWrap(true);

        // Клип длинных текстов «этикеточных» колонок
        [COLS_INV.PURPOSE, COLS_INV.APPLICATION, COLS_INV.INGREDIENTS].forEach(h => {
          const c = idxInv[h] + 1;
          if (c > 0) {
            shTarget.getRange(2, c, resultValues.length, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
            shTarget.setColumnWidth(c, 200);
          }
        });

        // Центр для ключевых числовых/кодов
        [COLS_INV.RUS_ART, COLS_INV.PRODUCER_ART, COLS_INV.QTY, COLS_INV.PRICE, COLS_INV.TOTAL].forEach(h => {
          const c = idxInv[h] + 1;
          if (c > 0) shTarget.getRange(2, c, lastRow - 1, 1).setHorizontalAlignment('center');
        });

        // Эксклюзивный пересчёт/подсветка
        _applyExclusiveFormatting_(shTarget);

        // Срез «хвостов»
        const maxRows = shTarget.getMaxRows();
        if (maxRows > lastRow) shTarget.deleteRows(lastRow + 1, maxRows - lastRow);
        const lastCol = shTarget.getLastColumn();
        const maxCols = shTarget.getMaxColumns();
        if (maxCols > lastCol) shTarget.deleteColumns(lastCol + 1, maxCols - lastCol);
      }

      ss.toast('Готово', 'OK', 3);
      ui.alert('Успех!', `Лист "${TGT_NAME}" успешно сформирован.`, ui.ButtonSet.OK);
      Lib.logInfo(`[INVOICE] Лист "${TGT_NAME}" создан`);
    } catch (e) {
      ss.toast('Ошибка', 'Ошибка', 3);
      Lib.logError('createFullInvoice: критическая ошибка', e);
      SpreadsheetApp.getUi().alert('Произошла ошибка!', `Подробности: ${e.message}`, ui.ButtonSet.OK);
    } finally {
      // Гарантированно закрываем любой висящий toast
      SpreadsheetApp.flush();
      ss.toast('', '', 1);
    }
  };

  // ------------------------ helpers ------------------------

  function headerIndex(headers, dict, sheetHumanName) {
    const res = {};
    const missing = [];
    Object.keys(dict).forEach(k => {
      const name = dict[k];
      const i = headers.indexOf(name);
      if (i === -1) { missing.push(name); }
      res[k] = i;
    });
    if (missing.length) {
      throw new Error(`На листе "${sheetHumanName}" отсутствуют столбцы: ${missing.join(', ')}`);
    }
    return res;
  }

  // Эксклюзив: подсветка и пересчёт для «Пробники» и «Mini Size»
  function _applyExclusiveFormatting_(targetSheet) {
    const CATEGORY_MINI = 'Mini Size';
    const CATEGORY_SAMPLES = 'Пробники';
    const COLOR_MINI = '#f4cccc';
    const COLOR_SAMPLES = '#fce5cd';

    try {
      const main = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S.PRIMARY);
      if (!main) return;

      const mv = main.getDataRange().getValues();
      const mh = (mv.shift() || []).map(x => String(x || '').trim());

      const idxArt  = mh.indexOf('Арт. произв.');
      const idxCat  = mh.indexOf('Категория товара');
      const idxPack = mh.indexOf('шт./уп.');
      if (idxArt === -1 || idxCat === -1 || idxPack === -1) return;

      const mp = new Map();
      mv.forEach(r => {
        const art = String(r[idxArt] || '').trim();
        if (!art) return;
        mp.set(art, { cat: String(r[idxCat] || '').trim(), perPack: r[idxPack] });
      });

      const th = (targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0] || []).map(x => String(x || '').trim());
      const iArt = th.indexOf(COLS_INV.PRODUCER_ART);
      const iQty = th.indexOf(COLS_INV.QTY);
      const toHighlight = [COLS_INV.PURPOSE, COLS_INV.APPLICATION, COLS_INV.INGREDIENTS]
        .map(h => th.indexOf(h)).filter(i => i !== -1);

      if (iArt === -1 || iQty === -1) return;

      const lastRow = targetSheet.getLastRow();
      if (lastRow <= 1) return;

      const data = targetSheet.getRange(2, 1, lastRow - 1, targetSheet.getLastColumn()).getValues();

      data.forEach((r, idx0) => {
        const art = String(r[iArt] || '').trim();
        const found = mp.get(art);
        if (!found) return;

        let color = null;
        if (found.cat.toLowerCase() === CATEGORY_SAMPLES.toLowerCase()) {
          color = COLOR_SAMPLES;
          const perPack = parseFloat(found.perPack);
          const q = parseFloat(r[iQty]);
          if (!isNaN(perPack) && perPack > 0 && !isNaN(q)) {
            targetSheet.getRange(idx0 + 2, iQty + 1).setValue(q * perPack);
          }
        } else if (found.cat.toLowerCase() === CATEGORY_MINI.toLowerCase()) {
          color = COLOR_MINI;
        }

        if (color && toHighlight.length) {
          toHighlight.forEach(ci => targetSheet.getRange(idx0 + 2, ci + 1).setBackground(color));
        }
      });
    } catch (_) { /* не критично */ }
  }

})(Lib, this);






// ============================================================================
//  СОБРАТЬ ДОКУМЕНТЫ И ОПИСАНИЕ ИЗ ИНВОЙСА
//  Экспорт: Lib.collectAndCopyDocuments()
// ============================================================================
(function (Lib, global) {
  Lib.collectAndCopyDocuments = function () {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // --- Имена колонок (как в реальных листах / конфиге) ---
    const COL_INV = {
      PRODUCT_NAME: 'Наименование',
      PURPOSE: 'Назначение Этикетка',
      APPLICATION: 'Применение Этикетка',
      ACTIVE_INGREDIENTS: 'Активные ингредиенты без %',
      DS_LINK: 'ДС',
      SPIRIT_LINK: 'Спирт'
    };
    const COL_CERT = {
      INVOICE_NAME: 'Наименование для инвойса',
      DS_NAME: 'Наименование ДС'
    };

    try {
      const invoiceSheet = ss.getSheetByName(Lib.CONFIG.SHEETS.INVOICE_FULL);
      const certSheet    = ss.getSheetByName(Lib.CONFIG.SHEETS.CERTIFICATION);
      if (!invoiceSheet || !certSheet) {
        throw new Error(
          `Не найден один из листов: "${Lib.CONFIG.SHEETS.INVOICE_FULL}" или "${Lib.CONFIG.SHEETS.CERTIFICATION}".`
        );
      }

      ui.alert('Начинаю сбор документов и создание описания... Это может занять некоторое время.');

      // --- Папки/документ назначения ---
      const parentFolderId = CONFIG.DRIVE.INVOICES.COLLECTION_PARENT_FOLDER_ID;
      if (!parentFolderId) throw new Error('В конфиге не задано DRIVE.INVOICES.COLLECTION_PARENT_FOLDER_ID');
      const parentFolder = DriveApp.getFolderById(parentFolderId);

      const now = new Date();
      const monthName = Lib.CONFIG.VALUES.MONTH_NAMES[now.getMonth()];
      const mainFolderName = `${Lib.CONFIG.SETTINGS.BRAND_PREFIX}${monthName} ${now.getFullYear()}`;

      const mainFolder   = _getOrCreateFolder_(parentFolder, mainFolderName);
      const dsFolder     = _getOrCreateFolder_(mainFolder, 'ДС');
      const spiritsFolder= _getOrCreateFolder_(mainFolder, 'Спирты');

      const docName = `${Lib.CONFIG.SETTINGS.BRAND_PREFIX}${Lib.CONFIG.DRIVE.INVOICE_DESCRIPTION.FILE_NAME_SUFFIX}`;
      const descriptionDoc = _getOrCreateDescriptionDoc_(mainFolder, docName);
      const body = descriptionDoc.getBody().clear();

      // --- Справочник "Наименование для инвойса" → "Наименование ДС" ---
      const certValues = certSheet.getDataRange().getValues();
      const certHeaders = (certValues.shift() || []).map(h => String(h).trim());
      const certIdx = {
        invoiceName: certHeaders.indexOf(COL_CERT.INVOICE_NAME),
        dsName:      certHeaders.indexOf(COL_CERT.DS_NAME),
      };
      if (certIdx.invoiceName === -1 || certIdx.dsName === -1) {
        throw new Error('На листе "Сертификация" нет колонок "Наименование для инвойса" или "Наименование ДС".');
      }

      const nameLookup = new Map();
      certValues.forEach(row => {
        const key = String(row[certIdx.invoiceName] || '').trim();
        const val = String(row[certIdx.dsName] || '').trim();
        if (key && val) nameLookup.set(key, val);
      });

      // --- Данные с листа "Для инвойса" ---
      const rng = invoiceSheet.getDataRange();
      const values = rng.getValues();
      const rich   = rng.getRichTextValues();

      const headers = (values.shift() || []).map(h => String(h).trim());
      rich.shift();

      const idx = {};
      [COL_INV.PRODUCT_NAME, COL_INV.PURPOSE, COL_INV.APPLICATION, COL_INV.ACTIVE_INGREDIENTS,
       COL_INV.DS_LINK, COL_INV.SPIRIT_LINK
      ].forEach(name => {
        const i = headers.indexOf(name);
        if (i === -1) throw new Error(`На листе "${Lib.CONFIG.SHEETS.INVOICE_FULL}" не найден столбец "${name}"`);
        idx[name] = i;
      });

      // --- Обработка строк: ссылки и формирование описания ---
      const dsLinks = new Set();
      const spiritLinks = new Set();
      const processedNames = new Set();

      values.forEach((row, i0) => {
        const rtRow = rich[i0] || [];

        // Линки ДС/Спирт из RichText
        const dsRT = rtRow[idx[COL_INV.DS_LINK]];
        const spRT = rtRow[idx[COL_INV.SPIRIT_LINK]];
        const dsUrl = dsRT && dsRT.getLinkUrl ? dsRT.getLinkUrl() : null;
        const spUrl = spRT && spRT.getLinkUrl ? spRT.getLinkUrl() : null;
        if (dsUrl) dsLinks.add(dsUrl);
        if (spUrl) spiritLinks.add(spUrl);

        // Имя для описания: «чистим» через справочник
        const invoiceName = String(row[idx[COL_INV.PRODUCT_NAME]] || '').trim();
        if (!invoiceName || invoiceName === Lib.CONFIG.VALUES.NOT_FOUND_TEXT) return;

        const cleanName = nameLookup.get(invoiceName) || invoiceName;
        if (processedNames.has(cleanName)) {
          _logAndHighlightInvoiceProcessing_(invoiceSheet, i0 + 2, idx[COL_INV.PRODUCT_NAME] + 1, '#fff2cc',
            `Строка ${i0 + 2}: "${cleanName}" пропущен (дубликат)`);
          return;
        }
        processedNames.add(cleanName);
        _logAndHighlightInvoiceProcessing_(invoiceSheet, i0 + 2, idx[COL_INV.PRODUCT_NAME] + 1, '#d9ead3',
          `Строка ${i0 + 2}: "${cleanName}" добавлен в отчет.`);

        // Пишем описание
        body.appendParagraph(cleanName)
            .setHeading(DocumentApp.ParagraphHeading.HEADING1)
            .setBold(true).setFontSize(14).setFontFamily('Times New Roman');

        _appendDescriptionParagraph_(body, 'Назначение: ', row[idx[COL_INV.PURPOSE]]);
        _appendDescriptionParagraph_(body, 'Применение: ', row[idx[COL_INV.APPLICATION]]);
        _appendDescriptionParagraph_(body, 'Активные ингредиенты: ', row[idx[COL_INV.ACTIVE_INGREDIENTS]]);
        body.appendParagraph('');
      });

      descriptionDoc.saveAndClose();

      // --- Копирование файлов ---
      const dsCopied      = _copyFilesFromLinks_(dsLinks, dsFolder);
      const spiritsCopied = _copyFilesFromLinks_(spiritLinks, spiritsFolder);

      const msg =
        `Сбор завершен!\n\n` +
        `- Создан/обновлён документ "${docName}" с ${processedNames.size} уникальными описаниями.\n` +
        `- Скопировано новых ДС: ${dsCopied}\n` +
        `- Скопировано новых Спиртов: ${spiritsCopied}`;

      ui.alert('Операция завершена!', msg, ui.ButtonSet.OK);

    } catch (e) {
      // Используем logWithEmoji для критических ошибок
      if (typeof global.logWithEmoji === 'function') {
        global.logWithEmoji(`КРИТИЧЕСКАЯ ОШИБКА в collectAndCopyDocuments: ${e.message}`, 'ERROR', '❌', 'Инвойс', e.stack);
      } else {
        Logger.log(`КРИТИЧЕСКАЯ ОШИБКА в collectAndCopyDocuments: ${e.message}\n${e.stack}`);
      }
      SpreadsheetApp.getUi().alert(`Произошла ошибка: ${e.message}.`);
    }
  };

  // ====== ЛОКАЛЬНЫЕ ХЕЛПЕРЫ (используются и другими методами этого модуля) ======
  function _getOrCreateDescriptionDoc_(parentFolder, docName) {
    const files = parentFolder.getFilesByName(docName);
    if (files.hasNext()) {
      const file = files.next();
      if (file.getMimeType() === MimeType.GOOGLE_DOCS) return DocumentApp.openById(file.getId());
      file.setTrashed(true);
    }
    const doc = DocumentApp.create(docName);
    DriveApp.getFileById(doc.getId()).moveTo(parentFolder);
    return doc;
  }

  function _appendDescriptionParagraph_(body, title, text) {
    if (!text || String(text).trim() === '') return;
    const p = body.appendParagraph(title);
    p.setBold(true).setFontSize(12).setFontFamily('Times New Roman');
    p.appendText(String(text)).setBold(false);
  }

  function _logAndHighlightInvoiceProcessing_(sheet, rowIndex, colIndex, color, message) {
    // Используем logWithEmoji если доступна
    if (typeof global.logWithEmoji === 'function') {
      global.logWithEmoji(message, 'INFO', null, 'Инвойс');
    } else {
      Logger.log(message);
    }
    sheet.getRange(rowIndex, colIndex).setBackground(color);
  }

  function _copyFilesFromLinks_(linkSet, destinationFolder) {
    let copiedCount = 0;
    const idRegex = /(?:\/d\/|\?id=)([a-zA-Z0-9_-]{25,})/;
    for (const link of linkSet) {
      const m = String(link || '').match(idRegex);
      const fileId = m ? m[1] : null;
      if (!fileId) continue;
      try {
        const file = DriveApp.getFileById(fileId);
        if (!destinationFolder.getFilesByName(file.getName()).hasNext()) {
          file.makeCopy(file.getName(), destinationFolder);
          copiedCount++;
        }
      } catch (e) { /* игнорируем нет доступа/удалено */ }
    }
    return copiedCount;
  }

  function _getOrCreateFolder_(parent, name) {
    const it = parent.getFoldersByName(name);
    return it.hasNext() ? it.next() : parent.createFolder(name);
  }
})(Lib || (Lib = {}), this);
