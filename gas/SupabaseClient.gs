/**
 * =============================================================================
 * SUPABASE CLIENT FOR GOOGLE APPS SCRIPT
 * =============================================================================
 * Интеграция с Supabase для синхронизации данных
 *
 * Версия: 1.0
 * Дата: 2026-01-14
 * Автор: AgentCare System
 * =============================================================================
 */

// ============================================================================
// КОНФИГУРАЦИЯ
// ============================================================================

const SUPABASE_CONFIG = {
  // ⚠️ ВАЖНО: Замените на ваши реальные значения из Supabase console
  URL: "https://jfpnhkcqvriblsiqqjis.supabase.co",
  ANON_KEY: "YOUR_ANON_KEY_HERE", // Замените на Anon Key из Supabase
  // ⚠️ НИКОГДА не выкладывайте Service Role Key в GAS!
  // Используйте только ANON_KEY для клиентского кода
  // Service Role Key используется только в backend (Python)
};

// ============================================================================
// SUPABASE CLIENT
// ============================================================================

/**
 * Основной класс для работы с Supabase
 */
class SupabaseClient {
  constructor(config = SUPABASE_CONFIG) {
    this.url = config.URL;
    this.anonKey = config.ANON_KEY;
    this.baseUrl = `${this.url}/rest/v1`;
  }

  /**
   * Получить заголовки для запроса
   */
  getHeaders() {
    return {
      "apikey": this.anonKey,
      "Authorization": `Bearer ${this.anonKey}`,
      "Content-Type": "application/json"
    };
  }

  /**
   * Выполнить GET запрос
   * @param {string} table - имя таблицы
   * @param {Object} params - параметры запроса
   * @returns {Array} данные из таблицы
   */
  select(table, params = {}) {
    const url = this.buildUrl(table, params);

    try {
      const options = {
        method: 'get',
        headers: this.getHeaders(),
        muteHttpExceptions: true
      };

      const response = UrlFetchApp.fetch(url, options);
      const code = response.getResponseCode();

      if (code === 200) {
        return JSON.parse(response.getContentText());
      } else {
        const error = response.getContentText();
        Logger.log(`❌ Error ${code}: ${error}`);
        return null;
      }
    } catch (e) {
      Logger.log(`❌ Select error: ${e.message}`);
      return null;
    }
  }

  /**
   * Выполнить INSERT запрос
   * @param {string} table - имя таблицы
   * @param {Object|Array} data - данные для вставки
   * @returns {Array} вставленные записи
   */
  insert(table, data) {
    const url = `${this.baseUrl}/${table}`;
    const payload = JSON.stringify(Array.isArray(data) ? data : [data]);

    try {
      const options = {
        method: 'post',
        headers: this.getHeaders(),
        payload: payload,
        muteHttpExceptions: true
      };

      const response = UrlFetchApp.fetch(url, options);
      const code = response.getResponseCode();

      if (code === 201) {
        return JSON.parse(response.getContentText());
      } else {
        const error = response.getContentText();
        Logger.log(`❌ Insert error ${code}: ${error}`);
        return null;
      }
    } catch (e) {
      Logger.log(`❌ Insert error: ${e.message}`);
      return null;
    }
  }

  /**
   * Выполнить UPDATE запрос
   * @param {string} table - имя таблицы
   * @param {Object} data - данные для обновления
   * @param {Object} match - условия WHERE
   * @returns {Array} обновленные записи
   */
  update(table, data, match) {
    let url = `${this.baseUrl}/${table}`;

    // Добавить условия WHERE
    const conditions = [];
    for (const [key, value] of Object.entries(match)) {
      if (typeof value === 'string') {
        conditions.push(`${key}=eq.${encodeURIComponent(value)}`);
      } else {
        conditions.push(`${key}=eq.${value}`);
      }
    }
    url += '?' + conditions.join('&');

    const payload = JSON.stringify(data);

    try {
      const options = {
        method: 'patch',
        headers: this.getHeaders(),
        payload: payload,
        muteHttpExceptions: true
      };

      const response = UrlFetchApp.fetch(url, options);
      const code = response.getResponseCode();

      if (code === 200) {
        return JSON.parse(response.getContentText());
      } else {
        const error = response.getContentText();
        Logger.log(`❌ Update error ${code}: ${error}`);
        return null;
      }
    } catch (e) {
      Logger.log(`❌ Update error: ${e.message}`);
      return null;
    }
  }

  /**
   * Выполнить DELETE запрос
   * @param {string} table - имя таблицы
   * @param {Object} match - условия WHERE
   * @returns {boolean} успешно ли удалено
   */
  delete(table, match) {
    let url = `${this.baseUrl}/${table}`;

    // Добавить условия WHERE
    const conditions = [];
    for (const [key, value] of Object.entries(match)) {
      if (typeof value === 'string') {
        conditions.push(`${key}=eq.${encodeURIComponent(value)}`);
      } else {
        conditions.push(`${key}=eq.${value}`);
      }
    }
    url += '?' + conditions.join('&');

    try {
      const options = {
        method: 'delete',
        headers: this.getHeaders(),
        muteHttpExceptions: true
      };

      const response = UrlFetchApp.fetch(url, options);
      const code = response.getResponseCode();

      return code === 204;
    } catch (e) {
      Logger.log(`❌ Delete error: ${e.message}`);
      return false;
    }
  }

  /**
   * Построить URL с параметрами
   * @private
   */
  buildUrl(table, params) {
    let url = `${this.baseUrl}/${table}`;
    const queryParts = [];

    for (const [key, value] of Object.entries(params)) {
      if (key === 'select') {
        queryParts.push(`select=${encodeURIComponent(value)}`);
      } else if (key === 'filter') {
        queryParts.push(`${value}`); // фильтры уже формируются как "col=eq.value"
      } else if (key === 'limit') {
        queryParts.push(`limit=${value}`);
      } else if (key === 'offset') {
        queryParts.push(`offset=${value}`);
      } else if (key === 'order') {
        queryParts.push(`order=${value}`);
      }
    }

    if (queryParts.length > 0) {
      url += '?' + queryParts.join('&');
    }

    return url;
  }
}

// ============================================================================
// SYNC OPERATIONS
// ============================================================================

/**
 * Вставить лог в Supabase
 * @param {Object} logData - данные лога
 */
function insertLogToSupabase(logData) {
  const client = new SupabaseClient();

  const log = {
    timestamp: logData.timestamp || new Date().toISOString(),
    level: logData.level || 'INFO',
    category: logData.category || 'Unknown',
    message: logData.message || '',
    function_name: logData.functionName || '',
    details: logData.details || '',
    variables: logData.variables ? JSON.stringify(logData.variables) : null,
    result: logData.result ? JSON.stringify(logData.result) : null,
    project_key: logData.projectKey || 'MT'
  };

  const result = client.insert('logs', log);

  if (result) {
    Logger.log(`✅ Log inserted to Supabase: ${log.message}`);
    return true;
  } else {
    Logger.log(`❌ Failed to insert log to Supabase`);
    return false;
  }
}

/**
 * Получить артикулы из Supabase
 * @param {string} projectKey - ключ проекта (MT, SS, SK)
 * @returns {Array} список артикулов
 */
function getArticlesFromSupabase(projectKey = 'MT') {
  const client = new SupabaseClient();

  const articles = client.select('articles', {
    select: '*',
    filter: `project_key=eq.${projectKey}`,
    limit: 10000
  });

  if (articles) {
    Logger.log(`✅ Retrieved ${articles.length} articles from Supabase`);
    return articles;
  } else {
    Logger.log(`❌ Failed to retrieve articles from Supabase`);
    return [];
  }
}

/**
 * Вставить артикулы в Supabase
 * @param {Array} articles - массив артикулов
 */
function insertArticlesToSupabase(articles) {
  const client = new SupabaseClient();

  const articlesData = articles.map(article => ({
    article_id: article.article_id || '',
    project_key: article.project_key || 'MT',
    article_rus: article.article_rus || null,
    article_producer: article.article_producer || null,
    code_base: article.code_base || null,
    code_1c: article.code_1c || null,
    name_eng_price: article.name_eng_price || null,
    name_eng_ds: article.name_eng_ds || null,
    name_rus_ds: article.name_rus_ds || null,
    volume: article.volume || null,
    release_form: article.release_form || null,
    units_per_pack: article.units_per_pack || null,
    category: article.category || null,
    line_name: article.line_name || null,
    group_name: article.group_name || null,
    barcode: article.barcode || null,
    status: article.status || 'ACTIVE',
    notes: article.notes || null,
    synced_from_sheet: new Date().toISOString(),
    google_sheet_id: article.google_sheet_id || ''
  }));

  // Вставьте батчами
  const batchSize = 100;
  let insertedCount = 0;

  for (let i = 0; i < articlesData.length; i += batchSize) {
    const batch = articlesData.slice(i, i + batchSize);
    const result = client.insert('articles', batch);

    if (result) {
      insertedCount += batch.length;
      Logger.log(`✅ Inserted batch of ${batch.length} articles`);
    } else {
      Logger.log(`❌ Failed to insert batch starting at index ${i}`);
    }

    // Подождите между батчами, чтобы не перегрузить API
    Utilities.sleep(500);
  }

  Logger.log(`✅ Total inserted: ${insertedCount} articles`);
  return insertedCount;
}

/**
 * Обновить артикулы в Supabase
 * @param {Array} articles - массив артикулов
 */
function updateArticlesInSupabase(articles) {
  const client = new SupabaseClient();

  let updatedCount = 0;

  for (const article of articles) {
    const data = {
      article_rus: article.article_rus || null,
      article_producer: article.article_producer || null,
      status: article.status || 'ACTIVE',
      notes: article.notes || null,
      synced_from_sheet: new Date().toISOString()
    };

    const result = client.update('articles', data, {
      article_id: article.article_id
    });

    if (result) {
      updatedCount++;
      Logger.log(`✅ Updated article: ${article.article_id}`);
    } else {
      Logger.log(`❌ Failed to update article: ${article.article_id}`);
    }

    // Небольшая задержка
    Utilities.sleep(100);
  }

  Logger.log(`✅ Total updated: ${updatedCount} articles`);
  return updatedCount;
}

// ============================================================================
// ТЕСТИРОВАНИЕ
// ============================================================================

/**
 * Функция для тестирования подключения к Supabase
 */
function testSupabaseConnection() {
  Logger.log("🧪 Testing Supabase connection...");

  const client = new SupabaseClient();

  // Тест 1: Получить категории логирования
  Logger.log("\n📋 Test 1: Getting log categories...");
  const categories = client.select('log_categories', { limit: 5 });
  if (categories) {
    Logger.log(`✅ Retrieved ${categories.length} categories`);
    Logger.log(JSON.stringify(categories, null, 2));
  } else {
    Logger.log("❌ Failed to retrieve categories");
  }

  // Тест 2: Вставить тестовый лог
  Logger.log("\n📝 Test 2: Inserting test log...");
  const testLog = {
    timestamp: new Date().toISOString(),
    level: 'INFO',
    category: 'Settings',
    message: 'Supabase connection test',
    function_name: 'testSupabaseConnection',
    details: 'Testing Google Apps Script integration'
  };

  if (insertLogToSupabase(testLog)) {
    Logger.log("✅ Test log inserted successfully");
  } else {
    Logger.log("❌ Failed to insert test log");
  }

  // Тест 3: Получить логи
  Logger.log("\n📊 Test 3: Retrieving logs...");
  const logs = client.select('logs', {
    select: '*',
    order: 'timestamp.desc',
    limit: 5
  });
  if (logs) {
    Logger.log(`✅ Retrieved ${logs.length} logs`);
    logs.forEach((log, index) => {
      Logger.log(`  ${index + 1}. ${log.timestamp} - ${log.level} - ${log.message}`);
    });
  } else {
    Logger.log("❌ Failed to retrieve logs");
  }

  Logger.log("\n✅ All tests completed!");
}

/**
 * Запустить тест из меню
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚡ Supabase')
    .addItem('Test Connection', 'testSupabaseConnection')
    .addItem('Get Articles', 'testGetArticles')
    .addItem('Insert Test Log', 'testInsertLog')
    .addSeparator()
    .addItem('Sync All Articles', 'testSyncAllArticles')
    .addToUi();
}

function testGetArticles() {
  const articles = getArticlesFromSupabase('MT');
  Logger.log(JSON.stringify(articles.slice(0, 3), null, 2));
}

function testInsertLog() {
  insertLogToSupabase({
    timestamp: new Date().toISOString(),
    level: 'INFO',
    category: 'Settings',
    message: 'Manual test from GAS',
    functionName: 'testInsertLog',
    projectKey: 'MT'
  });
}

function testSyncAllArticles() {
  Logger.log("🔄 Starting article sync...");
  const projectKeys = ['MT', 'SS', 'SK'];

  for (const projectKey of projectKeys) {
    Logger.log(`\n📦 Syncing project ${projectKey}...`);
    const articles = getArticlesFromSupabase(projectKey);
    Logger.log(`Retrieved ${articles.length} articles`);
  }

  Logger.log("\n✅ Sync completed!");
}
