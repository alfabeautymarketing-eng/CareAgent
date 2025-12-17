/**
 * Модуль для работы с Google Gemini API
 * Прямая интеграция без внешних сервисов
 *
 * @file 01GeminiAPI.js
 * @requires PropertiesService для хранения API ключа
 */

(function (global) {

// ============================================================================
// КОНФИГУРАЦИЯ
// ============================================================================

/**
 * Получить API ключ Gemini из свойств скрипта
 * Установить ключ: File → Project properties → Script properties
 * Ключ: GEMINI_API_KEY
 */
function getGeminiApiKey() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    throw new Error('Gemini API ключ не настроен. Добавьте GEMINI_API_KEY в Script Properties.');
  }
  return apiKey;
}

/**
 * Базовый URL для Gemini API
 */
/**
 * Базовый URL для Gemini API
 * v1beta поддерживает все современные модели (1.5 Flash, Pro, etc)
 */
const GEMINI_API_BASE_URL = 'https://generativelanguage.googleapis.com/v1beta';

/**
 * Получить текущую модель из настроек или вернуть дефолтную
 */
function getGeminiModel() {
  const model = PropertiesService.getScriptProperties().getProperty('GEMINI_MODEL');
  // Используем gemini-2.5-flash как наиболее актуальную и быструю модель (Dec 2024/2025)
  return model || 'gemini-2.5-flash'; 
}

// ============================================================================
// ОСНОВНЫЕ ФУНКЦИИ
// ============================================================================

/**
 * Получить список доступных моделей
 * Полезно для отладки ошибок "404 Not Found"
 */
function listGeminiModels() {
  try {
    const apiKey = getGeminiApiKey();
    const url = `${GEMINI_API_BASE_URL}/models?key=${apiKey}`;
    
    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      muteHttpExceptions: true
    });
    
    const code = response.getResponseCode();
    const text = response.getContentText();
    
    if (code !== 200) {
      throw new Error(`Ошибка получения списка моделей (${code}): ${text}`);
    }
    
    const data = JSON.parse(text);
    if (!data.models) return [];
    
    // Фильтруем только those that support generateContent
    return data.models
      .filter(m => m.supportedGenerationMethods && m.supportedGenerationMethods.includes('generateContent'))
      .map(m => m.name.replace('models/', '')); // remove prefix
      
  } catch (e) {
    Logger.log('❌ Ошибка listGeminiModels: ' + e.message);
    throw e;
  }
}

/**
 * Вызов Gemini API с текстовым промптом
 *
 * @param {string} prompt - Текст промпта
 * @param {Object} options - Дополнительные параметры
 * @returns {string} Ответ от Gemini
 */
/**
 * Вызов Gemini API с текстовым промптом
 *
 * @param {string} prompt - Текст промпта
 * @param {Object} options - Дополнительные параметры
 * @returns {string} Ответ от Gemini
 */
function callGeminiAPI(prompt, options) {
  options = options || {};

  try {
    const apiKey = getGeminiApiKey();
    // Используем переданную модель или из настроек
    const model = options.model || getGeminiModel();
    const url = `${GEMINI_API_BASE_URL}/models/${model}:generateContent?key=${apiKey}`;

    // Формат ответа: по умолчанию text/plain, JSON только когда явно запрошено
    const responseMimeType = options.jsonResponse ? "application/json" : "text/plain";

    const payload = {
      contents: [{
        parts: [{
          text: prompt
        }]
      }],
      generationConfig: {
        temperature: options.temperature || 0.1,
        topK: options.topK || 40,
        topP: options.topP || 0.95,
        maxOutputTokens: options.maxOutputTokens || 8192,
        responseMimeType: responseMimeType
      }
    };

    const requestOptions = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    Logger.log(`🤖 [API] Подготовка запроса к Gemini...`);
    Logger.log(`📌 Модель: ${model}`);
    Logger.log(`📏 Размер промпта: ${prompt.length} символов`);
    Logger.log(`🔗 URL: ${url.replace(/key=.*/, 'key=***')}`);

    // Retry логика для обработки временных ошибок (503, 429)
    // Увеличено до 5 попыток с экспоненциальной задержкой
    const maxRetries = 5;
    let lastError = null;

    for (let attempt = 1; attempt <= maxRetries; attempt++) {
      const startTime = new Date().getTime();
      Logger.log(`🚀 [API] Попытка ${attempt}/${maxRetries}...`);
      
      const response = UrlFetchApp.fetch(url, requestOptions);
      const responseCode = response.getResponseCode();
      const endTime = new Date().getTime();
      const duration = (endTime - startTime) / 1000;

      if (responseCode === 200) {
        Logger.log(`✅ [API] Успех! (${duration} сек)`);
        const result = JSON.parse(response.getContentText());

        // Логируем структуру ответа для отладки
        Logger.log(`📦 [API] Структура ответа: ${JSON.stringify(Object.keys(result))}`);

        if (!result.candidates || !result.candidates[0]) {
          Logger.log(`⚠️ [API] Полный ответ: ${JSON.stringify(result).substring(0, 500)}`);
          throw new Error('Пустой ответ от Gemini API (нет candidates)');
        }

        const candidate = result.candidates[0];
        Logger.log(`📦 [API] Candidate keys: ${JSON.stringify(Object.keys(candidate))}`);

        // Безопасное извлечение текста с проверками
        let text = '';
        if (candidate.content && candidate.content.parts && candidate.content.parts[0]) {
          text = candidate.content.parts[0].text || '';
        } else if (candidate.text) {
          // Альтернативный формат ответа
          text = candidate.text;
        } else if (candidate.output) {
          // Ещё один возможный формат
          text = candidate.output;
        } else {
          Logger.log(`⚠️ [API] Candidate структура: ${JSON.stringify(candidate).substring(0, 500)}`);
          throw new Error('Не удалось извлечь текст из ответа Gemini. Структура: ' + JSON.stringify(Object.keys(candidate)));
        }

        Logger.log(`📝 [API] Получено ${text.length} символов ответа.`);

        return text;
      }

      const errorText = response.getContentText();
      Logger.log(`⚠️ [API] Ответ сервера (код ${responseCode}): ${errorText}`);

      // Повторяем при 503 (перегрузка) или 429 (rate limit) или 500 (server error)
      if ((responseCode === 503 || responseCode === 429 || responseCode >= 500) && attempt < maxRetries) {
        // Экспоненциальная задержка: 2, 4, 8, 16 секунд...
        // let waitSeconds = Math.pow(2, attempt); 
        // Но чтобы не ждать вечность, сделаем линейно-прогрессивную для Gemini: 2, 5, 10, 20
        const waitSeconds = [2, 5, 10, 20, 30][attempt - 1] || 30;
        
        Logger.log(`⏳ [API] Сервер перегружен (503/429/5xx). Ждём ${waitSeconds} сек перед повтором...`);
        Utilities.sleep(waitSeconds * 1000);
        continue;
      }

      // Если модель не найдена (404), предлагаем решение
      if (responseCode === 404 && errorText.includes('not found')) {
        throw new Error(
          `Модель "${model}" не найдена или не доступна.\n\n` +
          `Попробуйте:\n` +
          `1. Изменить модель через меню "Настройки → Установить модель"\n` +
          `2. Использовать актуальные модели: gemini-2.5-flash, gemini-2.0-flash, gemini-2.5-pro\n` +
          `3. Проверить доступные модели через "Настройки → Показать модели"\n` +
          `4. Обновить API ключ в Google AI Studio (https://aistudio.google.com/apikey)\n\n` +
          `Детали: ${errorText}`
        );
      }

      lastError = new Error(`Gemini API error (${responseCode}): ${errorText}`);
      // Если это не 503/429 и не 500+, то прерываем цикл (например 400 Bad Request)
      if (responseCode < 500 && responseCode !== 429) {
         break;
      }
    }

    Logger.log('❌ [API] Все попытки исчерпаны.');
    throw lastError || new Error('Не удалось получить ответ от Gemini после ' + maxRetries + ' попыток');

  } catch (error) {
    Logger.log('❌ [API] Критическая ошибка: ' + error.message);
    throw error;
  }
}

/**
 * Анализ PDF документа с INCI составом
 *
 * @param {string} pdfUrl - URL PDF документа в Google Drive
 * @param {Object} context - Дополнительный контекст (назначение, применение)
 * @returns {Object} Результат анализа
 */
/**
 * Анализ PDF документа с INCI составом
 *
 * @param {string} pdfUrl - URL PDF документа в Google Drive
 * @param {Object} context - Дополнительный контекст (назначение, применение)
 * @returns {Object} Результат анализа
 */
function analyzePdfWithGemini(pdfUrl, context) {
  try {
    Logger.log('==============================================');
    Logger.log('🚀 [START] Начало процесса анализа PDF');
    Logger.log('🔗 Ссылка: ' + pdfUrl);
    
    // 1. Извлечь текст из PDF
    Logger.log('--- ШАГ 1: Извлечение текста ---');
    const pdfText = extractTextFromPdf(pdfUrl);
    
    if (!pdfText || pdfText.length < 10) {
      throw new Error('Извлеченный текст слишком короткий или пустой. Возможно, это скан без OCR.');
    }
    Logger.log(`📝 Успешно извлечено ${pdfText.length} символов.`);

    // 2. Сформировать промпт для анализа
    Logger.log('--- ШАГ 2: Подготовка промпта ---');
    const prompt = buildAnalysisPrompt(pdfText, context);
    Logger.log(`🧠 Промпт сформирован.`);

    // 3. Вызвать Gemini API
    Logger.log('--- ШАГ 3: Запрос к AI (может занять 10-30 сек) ---');
    const response = callGeminiAPI(prompt, {
      temperature: 0.1,
      maxOutputTokens: 4096,
      jsonResponse: true  // Ожидаем JSON для анализа
    });

    // 4. Парсить JSON ответ
    Logger.log('--- ШАГ 4: Обработка ответа ---');
    const result = parseGeminiResponse(response);

    Logger.log('✅ [SUCCESS] Анализ успешно завершен!');
    Logger.log('==============================================');
    return result;

  } catch (error) {
    Logger.log('❌ [ERROR] Ошибка в analyzePdfWithGemini: ' + error.message);
    // Пробрасываем ошибку выше, чтобы она попала в UI
    throw error;
  }
}

// ============================================================================
// ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
// ============================================================================

/**
 * Извлечь текст из PDF по URL
 *
 * @param {string} pdfUrl - URL PDF в Google Drive
 * @returns {string} Текст документа
 */
function extractTextFromPdf(pdfUrl) {
  try {
    // Извлечь ID файла из URL
    const fileId = extractFileIdFromUrl(pdfUrl);

    if (!fileId) {
      throw new Error('Не удалось извлечь ID файла из URL: ' + pdfUrl);
    }

    // Получить файл
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    const fileName = file.getName();

    Logger.log('📄 Конвертация PDF в текст: ' + fileName);

    // Используем Drive REST API v3 для конвертации PDF в Google Doc
    const token = ScriptApp.getOAuthToken();

    // Создаём Google Doc из PDF через multipart upload
    const boundary = '-------314159265358979323846';
    const metadata = {
      name: fileName + '_temp_ocr',
      mimeType: 'application/vnd.google-apps.document'
    };

    const requestBody =
      '--' + boundary + '\r\n' +
      'Content-Type: application/json; charset=UTF-8\r\n\r\n' +
      JSON.stringify(metadata) + '\r\n' +
      '--' + boundary + '\r\n' +
      'Content-Type: application/pdf\r\n' +
      'Content-Transfer-Encoding: base64\r\n\r\n' +
      Utilities.base64Encode(blob.getBytes()) + '\r\n' +
      '--' + boundary + '--';

    const response = UrlFetchApp.fetch(
      'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id',
      {
        method: 'POST',
        headers: {
          'Authorization': 'Bearer ' + token,
          'Content-Type': 'multipart/related; boundary=' + boundary
        },
        payload: requestBody,
        muteHttpExceptions: true
      }
    );

    if (response.getResponseCode() !== 200) {
      throw new Error('Ошибка создания документа: ' + response.getContentText());
    }

    const tempDocId = JSON.parse(response.getContentText()).id;

    // Читаем текст из созданного документа
    const doc = DocumentApp.openById(tempDocId);
    const text = doc.getBody().getText();

    // Удаляем временный документ
    DriveApp.getFileById(tempDocId).setTrashed(true);

    Logger.log('✅ Извлечено ' + text.length + ' символов');
    return text;

  } catch (error) {
    Logger.log('❌ Ошибка извлечения текста из PDF: ' + error.message);
    throw new Error('Не удалось извлечь текст из PDF: ' + error.message);
  }
}

/**
 * Извлечь File ID из URL Google Drive или найти файл по имени
 *
 * @param {string} urlOrName - URL файла или имя файла
 * @returns {string} File ID
 */
function extractFileIdFromUrl(urlOrName) {
  if (!urlOrName) return null;

  const input = urlOrName.trim();

  // Паттерны URL Google Drive
  const patterns = [
    /\/file\/d\/([a-zA-Z0-9_-]+)/,
    /id=([a-zA-Z0-9_-]+)/,
    /^([a-zA-Z0-9_-]{25,})$/  // Прямой ID
  ];

  for (const pattern of patterns) {
    const match = input.match(pattern);
    if (match) {
      return match[1];
    }
  }

  // Если это похоже на имя файла (содержит расширение или пробелы)
  if (input.includes('.pdf') || input.includes('.PDF') || input.includes(' ')) {
    Logger.log('🔍 Поиск файла по имени: ' + input);

    // Поиск файла по имени в Google Drive
    const files = DriveApp.getFilesByName(input);
    if (files.hasNext()) {
      const file = files.next();
      Logger.log('✅ Найден файл: ' + file.getName() + ' (ID: ' + file.getId() + ')');
      return file.getId();
    }

    // Попробуем поискать с частичным совпадением
    const searchQuery = 'title contains "' + input.replace(/"/g, '\\"').split('.')[0] + '" and mimeType = "application/pdf"';
    try {
      const searchResults = DriveApp.searchFiles(searchQuery);
      if (searchResults.hasNext()) {
        const file = searchResults.next();
        Logger.log('✅ Найден файл по частичному совпадению: ' + file.getName() + ' (ID: ' + file.getId() + ')');
        return file.getId();
      }
    } catch (e) {
      Logger.log('⚠️ Ошибка поиска: ' + e.message);
    }

    Logger.log('❌ Файл не найден: ' + input);
  }

  return null;
}

/**
 * Построить промпт для анализа состава
 *
 * @param {string} text - Текст из PDF
 * @param {Object} context - Контекст (назначение, применение)
 * @returns {string} Промпт для Gemini
 */
function buildAnalysisPrompt(text, context) {
  context = context || {};

  return `Ты эксперт по косметической продукции и классификации товаров по кодам ТН ВЭД.

ЗАДАЧА: Проанализируй состав косметического продукта и определи:
1. Код ТН ВЭД (3304, 3305 или 3307)
2. Категорию продукта
3. Вид продукции
4. Разрешительный документ (СГР или Декларация)
5. Составы на русском и английском языках

ИСХОДНЫЕ ДАННЫЕ:

Состав INCI:
${text}

${context.purpose ? 'Назначение: ' + context.purpose : ''}
${context.application ? 'Применение: ' + context.application : ''}

КАТЕГОРИИ ТН ВЭД:
- 3304: Средства косметические или для макияжа и средства для ухода за кожей
- 3305: Средства для ухода за волосами
- 3307: Средства для бритья, дезодоранты, соли для ванн и прочие косметические средства

ОТВЕТ должен быть в формате JSON:

{
  "category_code": "3304",
  "category": "Средства косметические или для макияжа...",
  "product_type": "Крем для лица",
  "product_reason": "Обоснование почему это крем для лица",
  "reg_doc": "СГР" или "Точно декларация",
  "reg_reason": "Обоснование выбора документа",
  "active_ru_no_pct": "Активные ингредиенты без % (рус)",
  "active_en_no_pct": "Active ingredients without % (eng)",
  "active_ru_pct": "Активные с % (рус)",
  "active_en_pct": "Active with % (eng)",
  "full_ru_no_pct": "Полный состав без % (рус)",
  "full_en_no_pct": "Full composition without % (eng)",
  "full_ru_pct": "Полный состав с % (рус)",
  "full_en_pct": "Full composition with % (eng)"
}

Верни ТОЛЬКО валидный JSON, без дополнительных комментариев.`;
}

/**
 * Парсить ответ от Gemini
 *
 * @param {string} response - Текст ответа
 * @returns {Object} Распарсенный JSON
 */
function parseGeminiResponse(response) {
  try {
    // Удалить markdown код блоки если есть
    let cleaned = response.trim();

    if (cleaned.startsWith('```json')) {
      cleaned = cleaned.replace(/```json\s*/g, '').replace(/```\s*$/g, '');
    } else if (cleaned.startsWith('```')) {
      cleaned = cleaned.replace(/```\s*/g, '').replace(/```\s*$/g, '');
    }

    const result = JSON.parse(cleaned);

    // Валидация обязательных полей
    const required = [
      'category_code',
      'category',
      'product_type',
      'reg_doc',
      'active_ru_no_pct',
      'full_ru_no_pct'
    ];

    for (const field of required) {
      if (!result[field]) {
        throw new Error('Отсутствует обязательное поле: ' + field);
      }
    }

    return result;

  } catch (error) {
    Logger.log('❌ Ошибка парсинга ответа Gemini: ' + error.message);
    Logger.log('Ответ: ' + response);
    throw new Error('Не удалось распарсить ответ Gemini: ' + error.message);
  }
}

// ============================================================================
// ЭКСПОРТ В GLOBAL (Имитация библиотеки)
// ============================================================================

  global.AutomationLib = global.AutomationLib || {};
  
  // Экспортируем функции API
  global.AutomationLib.callGeminiAPI = callGeminiAPI;
  global.AutomationLib.analyzePdfWithGemini = analyzePdfWithGemini;
  global.AutomationLib.listGeminiModels = listGeminiModels;
  global.AutomationLib.getGeminiModel = getGeminiModel;
  
  // Экспортируем внутренние функции только если нужно (как protected)
  global.AutomationLib._extractTextFromPdf = extractTextFromPdf;
  global.AutomationLib._buildAnalysisPrompt = buildAnalysisPrompt;
  global.AutomationLib._parseGeminiResponse = parseGeminiResponse;

  console.log('✅ AutomationLib: GeminiAPI loaded');

})(this);
