/**
 * Интеграция с Google Sheets для анализа составов
 * Использует Gemini API напрямую (без внешних сервисов)
 *
 * @file 02SheetsIntegration.js
 * @requires 01GeminiAPI.js
 */

(function (global) {

// ============================================================================
// ЗАВИСИМОСТИ
// ============================================================================

// Используем глобальный AutomationLib, который должен быть определен в 01GeminiAPI.js
// Если скрипты загружаются в неправильном порядке, это может вызвать ошибку.
// В GAS порядок загрузки определяется алфавитным порядком файлов, так что 01 был раньше 02.

const AutomationLib = global.AutomationLib || {};
const analyzePdfWithGemini = AutomationLib.analyzePdfWithGemini;

if (!analyzePdfWithGemini) {
    console.warn('⚠️ analyzePdfWithGemini не найден в AutomationLib. Возможно, 01GeminiAPI.js еще не загружен или произошла ошибка.');
}

// ============================================================================
// ФУНКЦИИ ДЛЯ РАБОТЫ С ТАБЛИЦЕЙ
// ============================================================================

/**
 * Анализировать выбранную строку в таблице
 *
 * @param {number} rowNumber - Номер строки
 * @param {string} sheetName - Название листа
 * @param {string} spreadsheetId - ID таблицы
 * @returns {Object} Результат анализа
 */
function analyzeRowInSheet(rowNumber, sheetName, spreadsheetId) {
  try {
    Logger.log('📋 Анализ строки ' + rowNumber + ' в листе "' + sheetName + '"');

    // Получить таблицу
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      throw new Error('Лист "' + sheetName + '" не найден');
    }

    // Прочитать данные строки
    const range = sheet.getRange(rowNumber, 1, 1, 25); // A-Y
    const values = range.getValues()[0];

    // Извлечь данные
    const rowData = {
      id: values[0],                  // A - ID
      productName: values[4],         // E - Наименование ДС
      inci_link: values[7],           // H - INCI ссылка
      purpose: values[8] || '',       // I - Назначение
      application: values[9] || ''    // J - Применение
    };

    if (!rowData.inci_link) {
      throw new Error('Отсутствует ссылка на INCI документ (колонка H)');
    }

    Logger.log('📄 Продукт: ' + rowData.productName);
    Logger.log('🔗 INCI: ' + rowData.inci_link);

    // Выполнить анализ через Gemini. Используем глобальную функцию из namespace
    const apiFunc = global.AutomationLib.analyzePdfWithGemini;
    if (!apiFunc) throw new Error('AutomationLib.analyzePdfWithGemini не доступен');

    const result = apiFunc(rowData.inci_link, {
      purpose: rowData.purpose,
      application: rowData.application
    });

    // Записать результаты в таблицу
    writeResultsToSheet(sheet, rowNumber, result);

    Logger.log('✅ Результаты записаны в строку ' + rowNumber);

    return result;

  } catch (error) {
    Logger.log('❌ Ошибка анализа строки: ' + error.message);
    throw error;
  }
}

/**
 * Записать результаты анализа в таблицу
 *
 * @param {Sheet} sheet - Лист таблицы
 * @param {number} rowNumber - Номер строки
 * @param {Object} result - Результат анализа
 */
function writeResultsToSheet(sheet, rowNumber, result) {
  // Подготовить данные для записи (колонки L-Y)
  const values = [[
    result.category_code,      // L - код категории
    result.category,           // M - категория
    result.product_type,       // N - вид продукции
    result.product_reason,     // O - обоснование
    result.reg_doc,            // P - документ
    result.reg_reason,         // Q - обоснование документа
    result.active_ru_no_pct,   // R - активные без % (рус)
    result.active_en_no_pct,   // S - активные без % (англ)
    result.active_ru_pct,      // T - активные с % (рус)
    result.active_en_pct,      // U - активные с % (англ)
    result.full_ru_no_pct,     // V - полный без % (рус)
    result.full_en_no_pct,     // W - полный без % (англ)
    result.full_ru_pct,        // X - полный с % (рус)
    result.full_en_pct         // Y - полный с % (англ)
  ]];

  // Записать в колонки L-Y (12-25)
  const range = sheet.getRange(rowNumber, 12, 1, 14);
  range.setValues(values);

  Logger.log('📝 Данные записаны в колонки L-Y');
}

/**
 * Найти и проанализировать все пустые строки
 *
 * @param {string} sheetName - Название листа
 * @param {string} spreadsheetId - ID таблицы
 * @returns {Object} Статистика обработки
 */
function analyzeEmptyRowsInSheet(sheetName, spreadsheetId) {
  try {
    Logger.log('📦 Поиск пустых строк в листе "' + sheetName + '"');

    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheet = ss.getSheetByName(sheetName);

    const lastRow = sheet.getLastRow();
    const data = sheet.getRange(2, 1, lastRow - 1, 25).getValues();

    const emptyRows = [];

    // Найти строки с пустыми результатами
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const hasInci = row[7];           // H - INCI link
      const hasCategory = row[11];      // L - категория
      const hasActiveRu = row[17];      // R - активные рус

      if (hasInci && (!hasCategory || !hasActiveRu)) {
        emptyRows.push(i + 2); // +2 т.к. начинаем со 2 строки
      }
    }

    Logger.log('📋 Найдено строк для обработки: ' + emptyRows.length);

    if (emptyRows.length === 0) {
      Logger.log('✅ Все строки уже обработаны');
      return {
        total: 0,
        success: 0,
        failed: 0
      };
    }

    // Обработать построчно
    const stats = {
      total: emptyRows.length,
      success: 0,
      failed: 0,
      errors: []
    };

    for (let i = 0; i < emptyRows.length; i++) {
      const rowNum = emptyRows[i];

      try {
        Logger.log('Обработка ' + (i + 1) + '/' + emptyRows.length + ' (строка ' + rowNum + ')');
        analyzeRowInSheet(rowNum, sheetName, spreadsheetId);
        stats.success++;

        // Пауза между запросами
        Utilities.sleep(2000);

      } catch (error) {
        stats.failed++;
        stats.errors.push({
          row: rowNum,
          error: error.message
        });
        Logger.log('⚠️ Ошибка в строке ' + rowNum + ': ' + error.message);
      }
    }

    Logger.log('✅ Обработка завершена');
    Logger.log('Успешно: ' + stats.success + '/' + stats.total);
    Logger.log('Ошибок: ' + stats.failed);

    return stats;

  } catch (error) {
    Logger.log('❌ Ошибка обработки пустых строк: ' + error.message);
    throw error;
  }
}

// ============================================================================
// ЭКСПОРТ В GLOBAL
// ============================================================================

  global.AutomationLib = global.AutomationLib || {};

  global.AutomationLib.analyzeRowInSheet = analyzeRowInSheet;
  global.AutomationLib.analyzeEmptyRowsInSheet = analyzeEmptyRowsInSheet;
  
  console.log('✅ AutomationLib: SheetsIntegration loaded');

})(this);
