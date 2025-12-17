var Lib = Lib || {};

/**
 * =======================================================================================
 * МОДУЛЬ: УПРАВЛЕНИЕ СТАТУСАМИ (StatusManager)
 * ---------------------------------------------------------------------------------------
 * Назначение: Автоматическое обновление статусов и синхронизация с листом "ТЗ по статусам"
 *
 * Функции:
 * 1. Обновление статусов после обработки прайсов
 * 2. Синхронизация данных с листом "ТЗ по статусам"
 * 3. Автоматическое удаление отработанных строк из ТЗ
 * =======================================================================================
 */

(function (Lib, global) {
  const TASK_SHEET_NAME = 'ТЗ по статусам';
  const STATUS_NEW = 'Снят в работу';

  // Столбцы для копирования из Главная в ТЗ
  const COLUMNS_TO_COPY_FROM_PRIMARY = [
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

  const TASK_SHEET_HEADERS = [
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

  /**
   * ПУБЛИЧНАЯ: Обновить статусы и синхронизировать с ТЗ после обработки
   * Вызывается после завершения последней функции обработки прайса
   */
  Lib.updateStatusesAfterProcessing = function () {
    try {
      Lib.logInfo('[StatusManager] Начало обновления статусов');

      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const primarySheet = ss.getSheetByName(Lib.CONFIG.SHEETS.PRIMARY);

      if (!primarySheet) {
        Lib.logError('[StatusManager] Лист Главная не найден');
        return;
      }

      const primaryData = primarySheet.getDataRange().getValues();
      if (primaryData.length <= 1) {
        Lib.logInfo('[StatusManager] Нет данных на листе Главная');
        return;
      }

      const headers = primaryData[0].map((h) => String(h || '').trim());
      const headerIndex = _makeHeaderIndex_(headers);

      const idpIndex = headerIndex['ID-P'];
      const statusIndex = headerIndex['Статус'];

      if (idpIndex === undefined || statusIndex === undefined) {
        Lib.logError('[StatusManager] Не найдены столбцы ID-P или Статус');
        return;
      }

      let updatedCount = 0;
      const rowsToUpdate = [];
      const updatedRowIndices = []; // Индексы строк, для которых мы ТОЛЬКО ЧТО изменили статус

      // Проходим по всем строкам и находим те, что нужно обновить
      for (let i = 1; i < primaryData.length; i++) {
        const row = primaryData[i];
        const idpValue = row[idpIndex];
        const statusValue = String(row[statusIndex] || '').trim();

        // Проверяем условие: ID-P пустой И Статус пустой
        const isIdpEmpty = idpValue === '' || idpValue === null || idpValue === undefined;
        const isStatusEmpty = statusValue === '';

        if (isIdpEmpty && isStatusEmpty) {
          row[statusIndex] = STATUS_NEW;
          rowsToUpdate.push({ rowNumber: i + 1, newStatus: STATUS_NEW });
          updatedRowIndices.push(i); // Запоминаем индекс строки
          updatedCount++;
        }
      }

      // Обновляем статусы на листе
      if (rowsToUpdate.length > 0) {
        rowsToUpdate.forEach((update) => {
          primarySheet.getRange(update.rowNumber, statusIndex + 1).setValue(update.newStatus);
        });
        Lib.logInfo('[StatusManager] Обновлено статусов: ' + updatedCount);

        // ВАЖНО: Прямая синхронизация статусов через правила
        try {
          _syncStatusDirectlyFromPrimary(ss, primarySheet, rowsToUpdate, STATUS_NEW);
        } catch (e) {
          Lib.logWarn('[StatusManager] Не удалось выполнить прямую синхронизацию статусов', e);
        }
      } else {
        Lib.logInfo('[StatusManager] Нет строк для обновления статусов');
      }

      // Синхронизируем с листом ТЗ ТОЛЬКО строки, которые мы только что обновили
      _syncToTaskSheet(ss, primarySheet, primaryData, headerIndex, updatedRowIndices);

      SpreadsheetApp.getActive().toast(
        'Обновлено статусов: ' + updatedCount,
        'Обновление статусов',
        3
      );
    } catch (error) {
      Lib.logError('[StatusManager] updateStatusesAfterProcessing: ошибка', error);
      SpreadsheetApp.getUi().alert(
        'Ошибка обновления статусов',
        error.message || String(error),
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  };

  /**
   * ПУБЛИЧНАЯ: Обработчик изменений для автоматического удаления строк из ТЗ
   * Вызывается из onEdit при установке галочки в столбце "Отработано"
   */
  Lib.handleTaskSheetEdit = function (e) {
    try {
      if (!e || !e.range) return;

      const sheet = e.range.getSheet();
      if (sheet.getName() !== TASK_SHEET_NAME) return;

      const row = e.range.getRow();
      const col = e.range.getColumn();

      // Проверяем, что редактируется не заголовок
      if (row <= 1) return;

      // Получаем заголовки
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const headerIndex = _makeHeaderIndex_(headers.map((h) => String(h || '').trim()));

      const completedIndex = headerIndex['Отработано'];
      if (completedIndex === undefined) return;

      // Проверяем, что изменение произошло в столбце "Отработано"
      if (col !== completedIndex + 1) return;

      const value = e.range.getValue();

      // Если установлена галочка (TRUE)
      if (value === true) {
        Lib.logInfo('[StatusManager] Удаление строки ' + row + ' из ТЗ (отработано)');

        // Удаляем строку
        sheet.deleteRow(row);

        SpreadsheetApp.getActive().toast(
          'Строка удалена из ТЗ',
          'Задача выполнена',
          2
        );
      }
    } catch (error) {
      Lib.logError('[StatusManager] handleTaskSheetEdit: ошибка', error);
    }
  };

  // =======================================================================================
  // ВНУТРЕННИЕ ФУНКЦИИ
  // =======================================================================================

  /**
   * Прямая синхронизация статусов через правила синхронизации
   * Работает с листом Главная как источником
   */
  function _syncStatusDirectlyFromPrimary(ss, primarySheet, rowsToUpdate, newStatusValue) {
    if (!rowsToUpdate || rowsToUpdate.length === 0) return;

    Lib.logInfo('[StatusManager] Начало прямой синхронизации статусов из Главная');

    try {
      // Получаем правила синхронизации
      const rulesSheetName = Lib.CONFIG.SHEETS.RULES || 'Правила синхро';
      const rulesSheet = ss.getSheetByName(rulesSheetName);

      if (!rulesSheet) {
        Lib.logWarn('[StatusManager] Лист "' + rulesSheetName + '" не найден');
        return;
      }

      const rulesLastRow = rulesSheet.getLastRow();
      if (rulesLastRow <= 1) {
        Lib.logInfo('[StatusManager] Нет правил синхронизации');
        return;
      }

      // Читаем все правила
      const rulesData = rulesSheet.getRange(2, 1, rulesLastRow - 1, 10).getValues();
      const primarySheetName = primarySheet.getName();
      const applicableRules = [];

      // Находим правила для синхронизации статуса из Главная
      for (let i = 0; i < rulesData.length; i++) {
        const rule = rulesData[i];
        const enabled = rule[1] === true || String(rule[1] || '').trim().toLowerCase() === 'да';
        const sourceSheet = String(rule[4] || '').trim();
        const sourceHeader = String(rule[5] || '').trim();
        const targetSheet = String(rule[6] || '').trim();
        const targetHeader = String(rule[7] || '').trim();
        const isExternal = rule[8] === true || String(rule[8] || '').trim().toLowerCase() === 'да';
        const targetDocId = String(rule[9] || '').trim();

        // Фильтруем: активные, источник = Главная, колонка = Статус
        if (enabled && sourceSheet === primarySheetName && sourceHeader === 'Статус' && targetSheet && targetHeader === 'Статус') {
          applicableRules.push({
            targetSheet: targetSheet,
            targetHeader: targetHeader,
            isExternal: isExternal,
            targetDocId: targetDocId
          });
        }
      }

      if (applicableRules.length === 0) {
        Lib.logInfo('[StatusManager] Нет применимых правил для синхронизации статуса');
        return;
      }

      Lib.logInfo('[StatusManager] Найдено применимых правил: ' + applicableRules.length);

      // Применяем изменения для каждого правила
      for (let i = 0; i < applicableRules.length; i++) {
        const rule = applicableRules[i];
        try {
          _applyStatusToTargetSheetFromPrimary(ss, primarySheet, rowsToUpdate, newStatusValue, rule);
        } catch (e) {
          Lib.logWarn('[StatusManager] Ошибка применения правила для листа "' + rule.targetSheet + '"', e);
        }
      }

      Lib.logInfo('[StatusManager] Прямая синхронизация статусов завершена');
    } catch (e) {
      Lib.logError('[StatusManager] _syncStatusDirectlyFromPrimary: критическая ошибка', e);
    }
  }

  /**
   * Применяет новый статус на целевой лист по правилу (из Главная)
   */
  function _applyStatusToTargetSheetFromPrimary(ss, primarySheet, rowsToUpdate, newStatusValue, rule) {
    // Получаем целевую таблицу
    const targetSS = rule.isExternal && rule.targetDocId
      ? SpreadsheetApp.openById(rule.targetDocId)
      : ss;

    const targetSheet = targetSS.getSheetByName(rule.targetSheet);
    if (!targetSheet) {
      Lib.logWarn('[StatusManager] Целевой лист "' + rule.targetSheet + '" не найден');
      return;
    }

    // Находим столбец ID и Статус на целевом листе
    const targetLastColumn = targetSheet.getLastColumn();
    const targetHeaders = targetSheet.getRange(1, 1, 1, targetLastColumn).getValues()[0];
    let targetIdIndex = -1;
    let targetStatusIndex = -1;

    for (let i = 0; i < targetHeaders.length; i++) {
      const header = String(targetHeaders[i] || '').trim();
      if (header === 'ID') targetIdIndex = i;
      if (header === rule.targetHeader) targetStatusIndex = i;
    }

    if (targetIdIndex === -1 || targetStatusIndex === -1) {
      Lib.logWarn('[StatusManager] На листе "' + rule.targetSheet + '" не найдены столбцы ID или ' + rule.targetHeader);
      return;
    }

    // Создаём карту ID целевого листа
    const targetLastRow = targetSheet.getLastRow();
    if (targetLastRow <= 1) return;

    const targetData = targetSheet.getRange(2, targetIdIndex + 1, targetLastRow - 1, 1).getValues();
    const targetIdMap = {};
    for (let i = 0; i < targetData.length; i++) {
      const id = String(targetData[i][0] || '').trim();
      if (id) {
        targetIdMap[id] = i + 2; // +2 для учёта заголовка
      }
    }

    // Применяем изменения
    let updatedCount = 0;
    for (let i = 0; i < rowsToUpdate.length; i++) {
      const update = rowsToUpdate[i];
      const rowNumber = update.rowNumber;

      // Получаем ID из листа Главная
      const id = String(primarySheet.getRange(rowNumber, 1).getValue() || '').trim();
      if (!id) continue;

      // Находим строку на целевом листе
      const targetRow = targetIdMap[id];
      if (targetRow) {
        targetSheet.getRange(targetRow, targetStatusIndex + 1).setValue(newStatusValue);
        updatedCount++;
      }
    }

    if (updatedCount > 0) {
      Lib.logInfo('[StatusManager] Обновлено статусов на листе "' + rule.targetSheet + '": ' + updatedCount);
    }
  }

  /**
   * Синхронизация данных с листом "ТЗ по статусам"
   * Данные берутся из листа Главная (основная информация) и листа Заказ (остатки и продажи)
   * @param {Set} updatedRowIndices - Индексы строк, для которых был только что изменён статус
   */
  function _syncToTaskSheet(ss, primarySheet, primaryData, primaryHeaderIndex, updatedRowIndices) {
    try {
      Lib.logInfo('[StatusManager] Начало синхронизации с ТЗ по статусам');

      // Если нет обновлённых строк, ничего не делаем
      if (!updatedRowIndices || updatedRowIndices.length === 0) {
        Lib.logInfo('[StatusManager] Нет обновлённых строк для синхронизации');
        return;
      }

      // Получаем или создаём лист ТЗ
      let taskSheet = ss.getSheetByName(TASK_SHEET_NAME);
      if (!taskSheet) {
        taskSheet = ss.insertSheet(TASK_SHEET_NAME);
        Lib.logInfo('[StatusManager] Создан новый лист: ' + TASK_SHEET_NAME);
      }

      // Проверяем/создаём заголовки
      _ensureTaskSheetHeaders(taskSheet);

      // Получаем текущие данные из ТЗ
      const taskData = taskSheet.getDataRange().getValues();
      const taskHeaders = taskData[0].map((h) => String(h || '').trim());
      const taskHeaderIndex = _makeHeaderIndex_(taskHeaders);

      // Создаём карту ID из ТЗ для быстрой проверки (на случай дублей)
      const existingIds = new Set();
      for (let i = 1; i < taskData.length; i++) {
        const idIndex = taskHeaderIndex['ID'];
        if (idIndex !== undefined) {
          const id = taskData[i][idIndex];
          if (id) existingIds.add(String(id));
        }
      }

      // Получаем лист Заказ для остатков и продаж
      const orderSheetName = (global.CONFIG && global.CONFIG.SHEETS && global.CONFIG.SHEETS.ORDER_FORM) || 'Заказ';
      const orderSheet = ss.getSheetByName(orderSheetName);

      if (!orderSheet) {
        Lib.logWarn('[StatusManager] Лист "' + orderSheetName + '" не найден, продолжаем без данных остатков');
      }

      // Получаем данные листа Заказ
      let orderDataMap = {};
      let orderStockIndex = -1;
      let orderSalesIndex = -1;

      if (orderSheet) {
        const orderData = orderSheet.getDataRange().getValues();
        if (orderData.length > 1) {
          const orderHeaders = orderData[0].map((h) => String(h || '').trim());
          const orderHeaderIndex = _makeHeaderIndex_(orderHeaders);
          const orderIdIndex = orderHeaderIndex['ID'];
          orderStockIndex = orderHeaderIndex['Остаток'];
          orderSalesIndex = orderHeaderIndex['ПРОДАЖИ'];

          // Создаём карту данных из листа Заказ по ID
          if (orderIdIndex !== undefined) {
            for (let i = 1; i < orderData.length; i++) {
              const id = orderData[i][orderIdIndex];
              if (id) {
                orderDataMap[String(id)] = {
                  stock: (orderStockIndex !== undefined && orderStockIndex !== -1) ? orderData[i][orderStockIndex] : '',
                  sales: (orderSalesIndex !== undefined && orderSalesIndex !== -1) ? orderData[i][orderSalesIndex] : ''
                };
              }
            }
          }
        }
      }

      // Собираем строки для добавления в ТЗ - ТОЛЬКО те, что были обновлены
      const rowsToAdd = [];
      const idIndex = primaryHeaderIndex['ID'];

      for (let i = 0; i < updatedRowIndices.length; i++) {
        const rowIndex = updatedRowIndices[i];
        const row = primaryData[rowIndex];
        const id = row[idIndex];

        // Проверяем, что строка ещё не добавлена в ТЗ (защита от дублей)
        if (id && !existingIds.has(String(id))) {
          // Получаем остаток и продажи из листа Заказ
          const orderInfo = orderDataMap[String(id)] || { stock: '', sales: '' };

          const newRow = _extractRowForTaskSheet(row, primaryData[0], primaryHeaderIndex, orderInfo.stock, orderInfo.sales);
          if (newRow) {
            rowsToAdd.push(newRow);
            existingIds.add(String(id)); // Добавляем в set, чтобы избежать дублей в текущем запуске
          }
        } else {
          Lib.logDebug('[StatusManager] Строка с ID=' + id + ' уже есть в ТЗ, пропускаем');
        }
      }

      // Добавляем новые строки в ТЗ
      if (rowsToAdd.length > 0) {
        const lastRow = taskSheet.getLastRow();
        const targetRange = taskSheet.getRange(
          lastRow + 1,
          1,
          rowsToAdd.length,
          TASK_SHEET_HEADERS.length
        );
        targetRange.setValues(rowsToAdd);

        // Добавляем чекбоксы в столбец "Отработано"
        const completedColIndex = taskHeaderIndex['Отработано'];
        if (completedColIndex !== undefined) {
          const checkboxRange = taskSheet.getRange(
            lastRow + 1,
            completedColIndex + 1,
            rowsToAdd.length,
            1
          );
          checkboxRange.insertCheckboxes();
        }

        Lib.logInfo('[StatusManager] Добавлено в ТЗ: ' + rowsToAdd.length + ' строк');
      } else {
        Lib.logInfo('[StatusManager] Нет новых строк для добавления в ТЗ');
      }
    } catch (error) {
      Lib.logError('[StatusManager] _syncToTaskSheet: ошибка', error);
    }
  }

  /**
   * Проверяет и создаёт заголовки на листе ТЗ
   */
  function _ensureTaskSheetHeaders(sheet) {
    const lastColumn = sheet.getLastColumn();

    // Если лист пустой, создаём заголовки
    if (lastColumn === 0 || sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, TASK_SHEET_HEADERS.length).setValues([TASK_SHEET_HEADERS]);

      // Форматируем заголовки
      sheet.getRange(1, 1, 1, TASK_SHEET_HEADERS.length)
        .setFontWeight('bold')
        .setBackground('#4a86e8')
        .setFontColor('#ffffff');

      // Замораживаем первую строку
      sheet.setFrozenRows(1);

      Lib.logInfo('[StatusManager] Созданы заголовки для ТЗ по статусам');
    }
  }

  /**
   * Извлекает данные строки для листа ТЗ
   * Данные берутся из листа Главная (основная информация) и листа Заказ (остатки и продажи)
   */
  function _extractRowForTaskSheet(row, headers, headerIndex, stock, sales) {
    const newRow = [];

    // Копируем столбцы из Главная
    for (let i = 0; i < COLUMNS_TO_COPY_FROM_PRIMARY.length; i++) {
      const columnName = COLUMNS_TO_COPY_FROM_PRIMARY[i];
      const colIndex = headerIndex[columnName];

      if (colIndex !== undefined) {
        newRow.push(row[colIndex]);
      } else {
        newRow.push('');
      }
    }

    // Добавляем Остаток и Продажи из листа Заказ
    newRow.push(stock !== undefined && stock !== null ? stock : ''); // Остаток
    newRow.push(sales !== undefined && sales !== null ? sales : ''); // Продажи

    // Добавляем пустое значение для комментария
    newRow.push('');

    // Добавляем пустое значение для чекбокса "Отработано"
    newRow.push(false);

    return newRow;
  }

  /**
   * Создаёт карту индексов заголовков
   */
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
})(Lib, this);
