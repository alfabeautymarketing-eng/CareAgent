/**
 * Настройка Gemini API
 * Скрипт для установки и проверки настроек Gemini
 *
 * @file 03SetupGemini.js
 */

console.log('LOADING: Automation/03SetupGemini.js');

// ============================================================================
// ФУНКЦИИ НАСТРОЙКИ
// ============================================================================

/**
 * Установить модель Gemini в Script Properties AutomationLib
 */
/**
 * Установить модель Gemini
 */
function setGeminiModel() {
  const ui = SpreadsheetApp.getUi();

  // Получаем текущую модель из Script Properties
  const props = PropertiesService.getScriptProperties();
  const currentModel = props.getProperty('GEMINI_MODEL') || 'gemini-2.5-flash (default)';

  const response = ui.prompt(
    'Настройка модели Gemini',
    'Текущая модель: ' + currentModel + '\n\n' +
    'Введите название модели (например, gemini-2.5-flash, gemini-2.0-flash, gemini-2.5-pro):\n' +
    'Оставьте пустым для сброса на gemini-2.5-flash.',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() === ui.Button.OK) {
    const model = response.getResponseText().trim();

    if (model) {
      props.setProperty('GEMINI_MODEL', model);
      ui.alert('✅ Модель установлена: ' + model);
    } else {
      props.deleteProperty('GEMINI_MODEL');
      ui.alert('✅ Модель сброшена на по умолчанию (gemini-2.5-flash)');
    }
  }
}

/**
 * Диагностика: показать доступные модели
 */
function showAvailableModels() {
  const ui = SpreadsheetApp.getUi();
  try {
    // Используем this для доступа к глобальной области
    const AutomationLib = this.AutomationLib;

    if (!AutomationLib || !AutomationLib.listGeminiModels) {
      throw new Error('AutomationLib не готов');
    }

    Logger.log('=== ПОЛУЧЕНИЕ СПИСКА МОДЕЛЕЙ ===');
    const models = AutomationLib.listGeminiModels();

    if (models.length === 0) {
      ui.alert(
        '⚠️ Модели не найдены',
        'Возможные причины:\n' +
        '1. API ключ неверный\n' +
        '2. API ключ не имеет доступа к Gemini\n' +
        '3. Проблема с сетью\n\n' +
        'Проверьте логи (View → Logs) для деталей.',
        ui.ButtonSet.OK
      );
      return;
    }

    const currentModel = AutomationLib.getGeminiModel();
    const message =
      'Текущая модель: ' + currentModel + '\n\n' +
      'Доступные модели (' + models.length + '):\n' +
      models.map((m, i) => (i + 1) + '. ' + m).join('\n');

    Logger.log(message);

    // Проверяем, доступна ли текущая модель
    const isCurrentAvailable = models.includes(currentModel);
    let alertMessage = message.substring(0, 1000);

    if (!isCurrentAvailable) {
      alertMessage += '\n\n⚠️ ВНИМАНИЕ: Текущая модель НЕ в списке доступных!';
    }

    ui.alert(
      'Список моделей Gemini',
      alertMessage + (message.length > 1000 ? '\n...' : ''),
      ui.ButtonSet.OK
    );

  } catch (e) {
    Logger.log('❌ Ошибка showAvailableModels: ' + e.message);
    Logger.log('Stack: ' + e.stack);
    ui.alert(
      '❌ Ошибка получения списка моделей',
      'Ошибка: ' + e.message + '\n\n' +
      'Возможно:\n' +
      '1. API ключ не установлен\n' +
      '2. API ключ неверный\n' +
      '3. Нет доступа к интернету\n\n' +
      'Проверьте логи для деталей.',
      ui.ButtonSet.OK
    );
  }
}

/**
 * Показать текущие настройки Gemini
 */
function showGeminiSettings() {
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('GEMINI_API_KEY');
  const model = props.getProperty('GEMINI_MODEL') || 'gemini-1.5-flash (default)';

  const message =
    'Текущие настройки Gemini API:\n\n' +
    'API Key: ' + (apiKey ? ('***' + apiKey.slice(-8)) : 'НЕ УСТАНОВЛЕН') + '\n' +
    'Модель: ' + model + '\n\n' +
    'Endpoint: v1beta\n';

  Logger.log(message);

  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    'Настройки Gemini API',
    message + '\n\nПоказать доступные модели?',
    ui.ButtonSet.YES_NO
  );
  
  if (result === ui.Button.YES) {
    showAvailableModels();
  }
}

/**
 * Установить API ключ (интерактивно)
 */
function setGeminiApiKey() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.prompt(
    'Настройка API ключа',
    'Введите ваш Gemini API ключ:\n\n' +
    'Получить ключ: https://aistudio.google.com/apikey\n\n' +
    'Ключ будет сохранён в Script Properties библиотеки EcosystemLib.',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() === ui.Button.OK) {
    const apiKey = response.getResponseText().trim();

    if (apiKey) {
      PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', apiKey);

      ui.alert(
        '✅ API ключ установлен',
        'Ключ успешно сохранён в Script Properties.\n\n' +
        'Теперь вы можете использовать функции AI Agent.',
        ui.ButtonSet.OK
      );

      Logger.log('✅ GEMINI_API_KEY установлен');
    } else {
      ui.alert('❌ Ошибка', 'Ключ не может быть пустым', ui.ButtonSet.OK);
    }
  }
}

/**
 * Полная настройка Gemini (ключ + модель)
 */
function setupGeminiComplete() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('GEMINI_API_KEY');

  if (!apiKey) {
    // Ключ не установлен - сразу предлагаем его установить
    const response = ui.alert(
      '⚙️ Настройка Gemini API',
      'API ключ не найден.\n\n' +
      'Для работы AI Agent необходимо установить GEMINI_API_KEY.\n\n' +
      'Хотите установить ключ сейчас?',
      ui.ButtonSet.YES_NO
    );

    if (response === ui.Button.YES) {
      setGeminiApiKey();
    }
  } else {
    // Ключ уже установлен
    const response = ui.alert(
      'Настройка Gemini API',
      'Выберите действие:\n' +
      'Да - Изменить API ключ\n' +
      'Нет - Изменить Модель\n' +
      'Отмена - Выход',
      ui.ButtonSet.YES_NO_CANCEL
    );

    if (response === ui.Button.YES) {
      setGeminiApiKey();
    } else if (response === ui.Button.NO) {
      setGeminiModel();
    }
  }
}

/**
 * Сбросить все настройки Gemini
 */
function resetGeminiSettings() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    'Сброс настроек',
    'Вы уверены, что хотите удалить все настройки Gemini?\n' +
    '(API ключ и модель будут удалены)',
    ui.ButtonSet.YES_NO
  );

  if (response === ui.Button.YES) {
    const props = PropertiesService.getScriptProperties();
    props.deleteProperty('GEMINI_API_KEY');
    props.deleteProperty('GEMINI_MODEL');

    ui.alert(
      '✅ Настройки сброшены',
      'Все настройки Gemini удалены из Script Properties',
      ui.ButtonSet.OK
    );

    Logger.log('✅ Настройки Gemini сброшены');
  }
}

// ============================================================================
// ЭКСПОРТ В БИБЛИОТЕКУ (global.Lib)
// ============================================================================
(function (global) {
  global.Lib = global.Lib || {};

  global.Lib.setGeminiModel = setGeminiModel;
  global.Lib.showGeminiSettings = showGeminiSettings;
  global.Lib.showAvailableModels = showAvailableModels;
  global.Lib.setGeminiApiKey = setGeminiApiKey;
  global.Lib.setupGeminiComplete = setupGeminiComplete;
  global.Lib.resetGeminiSettings = resetGeminiSettings;

  console.log('✅ 03SetupGemini: Setup functions loaded');
})(this);
