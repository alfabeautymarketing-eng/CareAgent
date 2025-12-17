/**
 * Интеграция с AutomationLib для AI-анализа косметических составов
 * Версия 2.0 - прямой вызов Gemini API (без внешних сервисов)
 *
 * @file 09AgentIntegration_v2.js
 * @requires AutomationLib (подключается как библиотека Apps Script)
 */

console.log('LOADING: 09AgentIntegration.js');
console.log('Function type check: ' + typeof menuCheckService_proxy);

// ============================================================================
// КОНФИГУРАЦИЯ
// ============================================================================

/**
 * ID таблицы МТ с данными для анализа
 */
const MT_SPREADSHEET_ID = '1fMOjUE7oZV96fCY5j5rPxnhWGJkDqg-GfwPZ8jUVgPw';

/**
 * Название листа с информацией
 */
const MT_SHEET_NAME = 'Информация';

// ============================================================================
// ФУНКЦИИ МЕНЮ
// ============================================================================

/**
 * Проверить статус Gemini API
 */
function menuCheckService() {
  try {
    if (!AutomationLib || !AutomationLib.callGeminiAPI) {
      throw new Error('AutomationLib не инициализирован');
    }

    Logger.log('=== ПРОВЕРКА GEMINI API ===');

    // Сначала попробуем получить список доступных моделей
    let availableModels = [];
    try {
      if (AutomationLib.listGeminiModels) {
        availableModels = AutomationLib.listGeminiModels();
        Logger.log('✅ Доступные модели: ' + availableModels.join(', '));
      }
    } catch (e) {
      Logger.log('⚠️ Не удалось получить список моделей: ' + e.message);
    }

    // Получаем текущую модель
    const currentModel = AutomationLib.getGeminiModel ? AutomationLib.getGeminiModel() : 'gemini-2.5-flash';
    Logger.log('📌 Текущая модель: ' + currentModel);

    // Используем прямой вызов через AutomationLib
    const response = AutomationLib.callGeminiAPI('Привет! Ответь: OK', {
      maxOutputTokens: 10
    });

    let message = '✅ Gemini API онлайн\n\n' +
      'Модель: ' + currentModel + '\n' +
      'Ответ: ' + response + '\n' +
      'Статус: Готов к работе';

    if (availableModels.length > 0) {
      message += '\n\nДоступно моделей: ' + availableModels.length;
    }

    SpreadsheetApp.getUi().alert(message);

  } catch (error) {
    Logger.log('❌ ОШИБКА: ' + error.message);
    Logger.log('Stack: ' + error.stack);

    SpreadsheetApp.getUi().alert(
      '❌ Gemini API недоступен\n\n' +
      'Ошибка: ' + error.message + '\n\n' +
      'Проверьте:\n' +
      '1. GEMINI_API_KEY в Script Properties\n' +
      '2. Логи (View → Logs) для деталей\n' +
      '3. Настройки → Показать доступные модели'
    );
  }
}

// Proxy-функция
function menuCheckService_proxy() {
  try {
    menuCheckService();
  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Ошибка: ' + error.message);
    Logger.log('Ошибка в menuCheckService: ' + error.message);
  }
}

/**
 * Анализировать выбранную строку
 */
function menuAnalyzeSelected() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const row = sheet.getActiveRange().getRow();

  if (row < 2) {
    SpreadsheetApp.getUi().alert('Выберите строку с данными (не заголовок)');
    return;
  }

  try {
    // Вызов функции из локального AutomationLib
    if (AutomationLib && AutomationLib.analyzeRowInSheet) {
      AutomationLib.analyzeRowInSheet(row, sheet.getName(), MT_SPREADSHEET_ID);
    } else {
      throw new Error('AutomationLib.analyzeRowInSheet не найдена. Проверьте загрузку модулей Automation.');
    }

    SpreadsheetApp.getUi().alert(
      '✅ Анализ завершен!\n\n' +
      'Результаты записаны в строку ' + row + '\n' +
      'Колонки L-Y обновлены'
    );

  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Ошибка: ' + error.message);
    Logger.log('Ошибка анализа: ' + error.message);
  }
}

// Proxy-функция
function menuAnalyzeSelected_proxy() {
  try {
    menuAnalyzeSelected();
  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Ошибка: ' + error.message);
    Logger.log('Ошибка в menuAnalyzeSelected: ' + error.message);
  }
}

/**
 * Анализировать пустые строки
 */
function menuAnalyzeEmpty() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Анализ пустых строк',
    'Это может занять много времени.\n' +
    'Обработка будет идти построчно с паузами.\n\n' +
    'Продолжить?',
    ui.ButtonSet.YES_NO
  );

  if (response === ui.Button.YES) {
    try {
      const stats = AutomationLib.analyzeEmptyRowsInSheet(
        SpreadsheetApp.getActiveSheet().getName(),
        MT_SPREADSHEET_ID
      );

      ui.alert(
        '✅ Обработка завершена!\n\n' +
        'Всего строк: ' + stats.total + '\n' +
        'Успешно: ' + stats.success + '\n' +
        'Ошибок: ' + stats.failed
      );

    } catch (error) {
      ui.alert('❌ Ошибка: ' + error.message);
      Logger.log('Ошибка обработки: ' + error.message);
    }
  }
}

// Proxy-функция
function menuAnalyzeEmpty_proxy() {
  try {
    menuAnalyzeEmpty();
  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Ошибка: ' + error.message);
    Logger.log('Ошибка в menuAnalyzeEmpty: ' + error.message);
  }
}

/**
 * Показать информацию о категориях
 */
function menuShowCategories() {
  const message =
    'Доступные категории ТН ВЭД:\n\n' +
    '3304 - Средства косметические или для макияжа\n' +
    '       и средства для ухода за кожей\n\n' +
    '3305 - Средства для ухода за волосами\n\n' +
    '3307 - Средства для бритья, дезодоранты,\n' +
    '       соли для ванн и прочие средства';

  SpreadsheetApp.getUi().alert('Категории ТН ВЭД', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

// Proxy-функция
function menuShowCategories_proxy() {
  try {
    menuShowCategories();
  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Ошибка: ' + error.message);
    Logger.log('Ошибка в menuShowCategories: ' + error.message);
  }
}

// ============================================================================
// ЭКСПОРТ В БИБЛИОТЕКУ (global.Lib)
// ============================================================================
(function (global) {
  // Гарантируем существование Lib
  global.Lib = global.Lib || {};

  // Привязываем функции к Lib
  global.Lib.menuCheckService = menuCheckService;
  global.Lib.menuAnalyzeSelected = menuAnalyzeSelected;
  global.Lib.menuAnalyzeEmpty = menuAnalyzeEmpty;
  global.Lib.menuShowCategories = menuShowCategories;



  // Тестовые
  global.Lib.testGeminiConnection = testGeminiConnection;
  global.Lib.testAnalyzeRow = testAnalyzeRow;
  
  // Убедимся, что AutomationLib существует (если файлы Automation/ еще не загрузились, создадим заглушку, но они должны загрузиться раньше по алфавиту 01.. перед 09..)
  // Но так как папка Automation/..., то порядок файлов GAS:
  // Automation/01GeminiAPI.js
  // Automation/02SheetsIntegration.js
  // 09AgentIntegration.js (корневая папка идет ПОСЛЕ папок? Нет, в GAS плоская структура. Имена файлов с путями.)
  // Обычно сортировка по имени файла.
  
  if (!global.AutomationLib) {
     console.warn('⚠️ AutomationLib object is missing in 09AgentIntegration.js. Check file load order.');
  }

  console.log('✅ 09AgentIntegration: Functions exported to Lib');
})(this);

// ============================================================================
// ТЕСТОВЫЕ ФУНКЦИИ
// ============================================================================

/**
 * Тест подключения к Gemini API
 */
function testGeminiConnection() {
  Logger.log('=== ТЕСТ ПОДКЛЮЧЕНИЯ К GEMINI ===');
  menuCheckService();
}

/**
 * Тест анализа строки
 */
function testAnalyzeRow() {
  Logger.log('=== ТЕСТ АНАЛИЗА СТРОКИ ===');

  try {
    const result = AutomationLib.analyzeRowInSheet(2, MT_SHEET_NAME, MT_SPREADSHEET_ID);
    Logger.log('✅ Тест успешен!');
    Logger.log(JSON.stringify(result, null, 2));
  } catch (error) {
    Logger.log('❌ Тест провален: ' + error.message);
  }
}
