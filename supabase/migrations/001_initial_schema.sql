-- ============================================================================
-- AGENTCARE SUPABASE SCHEMA
-- ============================================================================
-- Полная схема для синхронизации данных из Google Sheets
-- Дата: 2026-01-14
-- ============================================================================

-- ============================================================================
-- 1. ТАБЛИЦЫ ПРОЕКТОВ (Articles, Prices, Stocks)
-- ============================================================================

-- Таблица основных артикулов (из листа "Главная")
CREATE TABLE IF NOT EXISTS articles (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  project_key VARCHAR(10) NOT NULL, -- 'MT', 'SS', 'SK'

  -- IDs
  article_id VARCHAR(50) UNIQUE NOT NULL,
  article_id_p VARCHAR(50),
  article_id_g VARCHAR(50),
  article_id_l VARCHAR(50),

  -- Артикулы
  article_rus VARCHAR(255),
  article_producer VARCHAR(255),
  code_base VARCHAR(50),
  code_1c VARCHAR(50),

  -- Названия
  name_eng_price VARCHAR(500),
  name_eng_ds VARCHAR(500),
  name_rus_ds VARCHAR(500),

  -- Характеристики
  volume VARCHAR(50),
  release_form VARCHAR(100),
  units_per_pack VARCHAR(50),
  category VARCHAR(255),
  line_name VARCHAR(255),
  group_name VARCHAR(255),
  barcode VARCHAR(50),

  -- Статусы
  status VARCHAR(50) DEFAULT 'ACTIVE', -- ACTIVE, ARCHIVED, DRAFT
  notes TEXT,

  -- Аудит
  created_at TIMESTAMP DEFAULT NOW(),
  updated_at TIMESTAMP DEFAULT NOW(),
  synced_from_sheet TIMESTAMP,
  last_synced_at TIMESTAMP,
  google_sheet_id VARCHAR(500) -- ID листа в Google Sheets для обратной синхронизации
);

CREATE INDEX IF NOT EXISTS idx_articles_project ON articles(project_key);
CREATE INDEX IF NOT EXISTS idx_articles_status ON articles(status);
CREATE INDEX IF NOT EXISTS idx_articles_article_id ON articles(article_id);

-- Таблица цен и прайсов
CREATE TABLE IF NOT EXISTS prices (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  article_id UUID NOT NULL REFERENCES articles(id) ON DELETE CASCADE,
  project_key VARCHAR(10) NOT NULL,

  -- Цены по типам
  price_base DECIMAL(15, 2),
  price_sk DECIMAL(15, 2),
  price_mt DECIMAL(15, 2),
  price_ss DECIMAL(15, 2),

  -- Маржи
  margin_percent DECIMAL(5, 2),
  markup_percent DECIMAL(5, 2),

  -- Сроки
  date_from DATE,
  date_to DATE,

  -- История
  created_at TIMESTAMP DEFAULT NOW(),
  updated_at TIMESTAMP DEFAULT NOW(),
  last_synced_at TIMESTAMP
);

CREATE INDEX IF NOT EXISTS idx_prices_article ON prices(article_id);
CREATE INDEX IF NOT EXISTS idx_prices_project ON prices(project_key);
CREATE INDEX IF NOT EXISTS idx_prices_date_range ON prices(date_from, date_to);

-- Таблица остатков (stocks)
CREATE TABLE IF NOT EXISTS stocks (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  article_id UUID NOT NULL REFERENCES articles(id) ON DELETE CASCADE,
  project_key VARCHAR(10) NOT NULL,

  -- Остатки
  quantity_total INT,
  quantity_available INT,
  quantity_reserved INT,
  quantity_damaged INT,

  -- Поставщики
  supplier_name VARCHAR(255),
  supplier_code VARCHAR(50),

  -- Локации
  warehouse_location VARCHAR(255),
  shelf_number VARCHAR(50),

  -- История движений
  last_received_date DATE,
  last_sold_date DATE,

  -- Аудит
  created_at TIMESTAMP DEFAULT NOW(),
  updated_at TIMESTAMP DEFAULT NOW(),
  last_synced_at TIMESTAMP
);

CREATE INDEX IF NOT EXISTS idx_stocks_article ON stocks(article_id);
CREATE INDEX IF NOT EXISTS idx_stocks_project ON stocks(project_key);
CREATE INDEX IF NOT EXISTS idx_stocks_quantity ON stocks(quantity_available);

-- ============================================================================
-- 2. ТАБЛИЦЫ ЛОГИРОВАНИЯ
-- ============================================================================

-- Основной журнал логов
CREATE TABLE IF NOT EXISTS logs (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),

  -- Информация о логе
  timestamp TIMESTAMP DEFAULT NOW(),
  level VARCHAR(20), -- DEBUG, INFO, WARN, ERROR
  category VARCHAR(50), -- Article, Sync, Price, etc.
  message TEXT,

  -- Контекст
  function_name VARCHAR(255),
  details TEXT,

  -- Параметры и результаты (JSON)
  variables JSONB,
  result JSONB,

  -- Связь с проектом
  project_key VARCHAR(10),

  -- Поиск
  search_text TSVECTOR,

  created_at TIMESTAMP DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_logs_timestamp ON logs(timestamp DESC);
CREATE INDEX IF NOT EXISTS idx_logs_level ON logs(level);
CREATE INDEX IF NOT EXISTS idx_logs_category ON logs(category);
CREATE INDEX IF NOT EXISTS idx_logs_function ON logs(function_name);
CREATE INDEX IF NOT EXISTS idx_logs_project ON logs(project_key);
CREATE INDEX IF NOT EXISTS idx_logs_search ON logs USING GIN(search_text);

-- Таблица архивов логов
CREATE TABLE IF NOT EXISTS log_archives (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),

  -- Метаданные архива
  archive_date DATE NOT NULL,
  archive_type VARCHAR(50), -- 'DAILY', 'WEEKLY', 'MONTHLY'

  -- Данные (JSON)
  logs_data JSONB,

  -- Статистика
  total_logs INT,
  error_count INT,
  warning_count INT,

  -- Аудит
  created_at TIMESTAMP DEFAULT NOW(),
  created_by VARCHAR(255),

  -- Хранилище (опционально)
  storage_path VARCHAR(500), -- путь в облаке/диск
  compressed BOOLEAN DEFAULT FALSE
);

CREATE INDEX IF NOT EXISTS idx_log_archives_date ON log_archives(archive_date DESC);
CREATE INDEX IF NOT EXISTS idx_log_archives_type ON log_archives(archive_type);

-- ============================================================================
-- 3. ТАБЛИЦЫ КОНФИГУРАЦИИ И НАСТРОЕК
-- ============================================================================

-- Правила синхронизации
CREATE TABLE IF NOT EXISTS sync_rules (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),

  -- Основная информация
  name VARCHAR(255) NOT NULL,
  description TEXT,
  enabled BOOLEAN DEFAULT TRUE,

  -- Источник и назначение
  source_type VARCHAR(50), -- 'GOOGLE_SHEETS', 'SUPABASE'
  source_id VARCHAR(500), -- ID листа или таблицы
  target_type VARCHAR(50), -- 'GOOGLE_SHEETS', 'SUPABASE'
  target_id VARCHAR(500),

  -- Направление
  sync_direction VARCHAR(50), -- 'ONE_WAY', 'TWO_WAY'

  -- Расписание
  schedule VARCHAR(255), -- CRON или 'MANUAL', 'ON_CHANGE'

  -- Фильтры
  filters JSONB, -- условия для синхронизации

  -- Преобразование полей
  field_mapping JSONB, -- маппинг между полями

  -- История
  last_synced_at TIMESTAMP,
  next_sync_at TIMESTAMP,

  created_at TIMESTAMP DEFAULT NOW(),
  updated_at TIMESTAMP DEFAULT NOW(),

  created_by VARCHAR(255),
  updated_by VARCHAR(255)
);

CREATE INDEX IF NOT EXISTS idx_sync_rules_enabled ON sync_rules(enabled);
CREATE INDEX IF NOT EXISTS idx_sync_rules_source ON sync_rules(source_type, source_id);
CREATE INDEX IF NOT EXISTS idx_sync_rules_target ON sync_rules(target_type, target_id);

-- Категории логирования (справочник)
CREATE TABLE IF NOT EXISTS log_categories (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  code VARCHAR(50) UNIQUE NOT NULL,
  name VARCHAR(255) NOT NULL,
  emoji VARCHAR(10),
  description TEXT,
  color_hex VARCHAR(7), -- #RRGGBB
  created_at TIMESTAMP DEFAULT NOW()
);

-- Статусы операций (справочник)
CREATE TABLE IF NOT EXISTS operation_statuses (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  code VARCHAR(50) UNIQUE NOT NULL,
  name VARCHAR(100) NOT NULL,
  emoji VARCHAR(10),
  description TEXT,
  severity VARCHAR(20), -- INFO, SUCCESS, WARNING, ERROR
  created_at TIMESTAMP DEFAULT NOW()
);

-- Конфигурация приложения
CREATE TABLE IF NOT EXISTS app_config (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),

  -- Ключи конфигурации
  config_key VARCHAR(100) UNIQUE NOT NULL,
  config_value TEXT,

  -- Тип
  value_type VARCHAR(50), -- string, number, boolean, json

  -- Описание
  description TEXT,

  -- Видимость
  is_public BOOLEAN DEFAULT FALSE,

  created_at TIMESTAMP DEFAULT NOW(),
  updated_at TIMESTAMP DEFAULT NOW(),
  updated_by VARCHAR(255)
);

CREATE INDEX IF NOT EXISTS idx_app_config_key ON app_config(config_key);

-- ============================================================================
-- 4. ТАБЛИЦЫ СИНХРОНИЗАЦИИ (СЛУЖЕБНЫЕ)
-- ============================================================================

-- История синхронизации
CREATE TABLE IF NOT EXISTS sync_history (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),

  -- Информация о синхронизации
  sync_rule_id UUID REFERENCES sync_rules(id) ON DELETE SET NULL,

  -- Время
  started_at TIMESTAMP DEFAULT NOW(),
  completed_at TIMESTAMP,

  -- Результаты
  status VARCHAR(50), -- 'PENDING', 'IN_PROGRESS', 'SUCCESS', 'FAILED', 'PARTIAL'
  total_records INT,
  processed_records INT,
  failed_records INT,
  skipped_records INT,

  -- Логирование
  error_message TEXT,
  error_details JSONB,

  -- Идентификатор
  sync_batch_id VARCHAR(100) UNIQUE
);

CREATE INDEX IF NOT EXISTS idx_sync_history_rule ON sync_history(sync_rule_id);
CREATE INDEX IF NOT EXISTS idx_sync_history_status ON sync_history(status);
CREATE INDEX IF NOT EXISTS idx_sync_history_started_at ON sync_history(started_at DESC);

-- Очередь синхронизации
CREATE TABLE IF NOT EXISTS sync_queue (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),

  -- Информация об элементе
  table_name VARCHAR(100) NOT NULL,
  record_id UUID NOT NULL,
  operation VARCHAR(20), -- 'INSERT', 'UPDATE', 'DELETE'

  -- Приоритет
  priority INT DEFAULT 0,

  -- Статус
  status VARCHAR(50) DEFAULT 'PENDING', -- 'PENDING', 'PROCESSING', 'DONE', 'FAILED'
  retry_count INT DEFAULT 0,
  max_retries INT DEFAULT 3,

  -- Данные
  data_before JSONB,
  data_after JSONB,

  -- Время
  created_at TIMESTAMP DEFAULT NOW(),
  processed_at TIMESTAMP,

  -- Идентификатор
  sync_batch_id VARCHAR(100)
);

CREATE INDEX IF NOT EXISTS idx_sync_queue_status ON sync_queue(status);
CREATE INDEX IF NOT EXISTS idx_sync_queue_priority ON sync_queue(priority DESC);
CREATE INDEX IF NOT EXISTS idx_sync_queue_created_at ON sync_queue(created_at);

-- ============================================================================
-- 5. ТАБЛИЦЫ ПОЛЬЗОВАТЕЛЕЙ И ПРАВ
-- ============================================================================

-- Пользователи системы
CREATE TABLE IF NOT EXISTS users (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),

  -- Профиль
  email VARCHAR(255) UNIQUE NOT NULL,
  display_name VARCHAR(255),
  avatar_url TEXT,

  -- Статус
  status VARCHAR(50) DEFAULT 'ACTIVE', -- 'ACTIVE', 'INACTIVE', 'BANNED'

  -- Роли
  roles TEXT[] DEFAULT ARRAY['user'], -- 'admin', 'manager', 'analyst', 'user'

  -- Параметры
  settings JSONB,

  -- Действия
  last_login_at TIMESTAMP,
  created_at TIMESTAMP DEFAULT NOW(),
  updated_at TIMESTAMP DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_users_email ON users(email);
CREATE INDEX IF NOT EXISTS idx_users_status ON users(status);

-- Логирование действий пользователей
CREATE TABLE IF NOT EXISTS audit_logs (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),

  -- Пользователь
  user_id UUID REFERENCES users(id) ON DELETE SET NULL,
  user_email VARCHAR(255),

  -- Действие
  action VARCHAR(100),
  object_type VARCHAR(100),
  object_id VARCHAR(500),

  -- Данные
  changes JSONB,

  -- Результат
  success BOOLEAN DEFAULT TRUE,
  error_message TEXT,

  -- Время
  created_at TIMESTAMP DEFAULT NOW(),

  -- IP и браузер (опционально)
  ip_address INET,
  user_agent TEXT
);

CREATE INDEX IF NOT EXISTS idx_audit_logs_user ON audit_logs(user_id);
CREATE INDEX IF NOT EXISTS idx_audit_logs_object ON audit_logs(object_type, object_id);
CREATE INDEX IF NOT EXISTS idx_audit_logs_created_at ON audit_logs(created_at DESC);

-- ============================================================================
-- 6. ФУНКЦИИ И ТРИГГЕРЫ
-- ============================================================================

-- Функция автоматического обновления updated_at
CREATE OR REPLACE FUNCTION update_updated_at_column()
RETURNS TRIGGER AS $$
BEGIN
  NEW.updated_at = NOW();
  RETURN NEW;
END;
$$ LANGUAGE plpgsql;

-- Триггеры для обновления updated_at
CREATE TRIGGER trigger_articles_updated_at
BEFORE UPDATE ON articles
FOR EACH ROW
EXECUTE FUNCTION update_updated_at_column();

CREATE TRIGGER trigger_prices_updated_at
BEFORE UPDATE ON prices
FOR EACH ROW
EXECUTE FUNCTION update_updated_at_column();

CREATE TRIGGER trigger_stocks_updated_at
BEFORE UPDATE ON stocks
FOR EACH ROW
EXECUTE FUNCTION update_updated_at_column();

-- Функция для логирования изменений
CREATE OR REPLACE FUNCTION log_table_changes()
RETURNS TRIGGER AS $$
BEGIN
  INSERT INTO audit_logs (object_type, object_id, action, changes)
  VALUES (TG_TABLE_NAME, NEW.id::TEXT, TG_OP, TO_JSONB(NEW));
  RETURN NEW;
END;
$$ LANGUAGE plpgsql;

-- ============================================================================
-- 7. INITIAL DATA - Справочники
-- ============================================================================

-- Категории логирования
INSERT INTO log_categories (code, name, emoji, description) VALUES
('Article', 'Операции с артикулами', '📦', 'Добавление, удаление, редактирование артикулов'),
('Sync', 'Синхронизация', '🔄', 'Синхронизация данных между системами'),
('Price', 'Работа с ценами', '💰', 'Обновление и пересчет цен'),
('Stock', 'Управление остатками', '📊', 'Изменение остатков товара'),
('Export', 'Выгрузка данных', '📤', 'Экспорт в различные форматы'),
('Import', 'Импорт данных', '📥', 'Загрузка данных из внешних источников'),
('Agent', 'AI агент', '🤖', 'Операции умного анализа и подбора'),
('Certificate', 'Сертификация', '📋', 'Работа с сертификатами и документами'),
('Invoice', 'Счета и накладные', '🧾', 'Создание и управление счетами'),
('Delivery', 'Доставка', '🚚', 'Операции с доставкой'),
('Settings', 'Настройки', '⚙️', 'Изменение конфигурации системы'),
('Dashboard', 'Дашборд', '📊', 'Работа с визуализацией данных'),
('Database', 'База данных', '🗄️', 'Операции с БД'),
('Integration', 'Интеграция', '🔗', 'Интеграция с внешними сервисами'),
('Auth', 'Авторизация', '🔐', 'Аутентификация и управление доступом')
ON CONFLICT (code) DO NOTHING;

-- Статусы операций
INSERT INTO operation_statuses (code, name, emoji, severity, description) VALUES
('START', 'Начало операции', '▶️', 'INFO', 'Операция начала выполнение'),
('PROGRESS', 'В процессе', '⏳', 'INFO', 'Операция выполняется'),
('SUCCESS', 'Успешно', '✅', 'SUCCESS', 'Операция выполнена успешно'),
('WARNING', 'Предупреждение', '⚠️', 'WARNING', 'Операция выполнена с предупреждениями'),
('ERROR', 'Ошибка', '❌', 'ERROR', 'Операция завершилась с ошибкой'),
('RETRY', 'Повторная попытка', '🔁', 'WARNING', 'Выполняется повторная попытка'),
('TIMEOUT', 'Превышено время', '⏱️', 'ERROR', 'Операция превысила лимит времени'),
('CANCELLED', 'Отменено', '🛑', 'WARNING', 'Операция была отменена пользователем')
ON CONFLICT (code) DO NOTHING;

-- ============================================================================
-- КОНЕЦ SCHEMA (Функции поиска добавлены позже)
-- ============================================================================
-- Версия: 1.0
-- Дата создания: 2026-01-14
-- Автор: AgentCare System
-- ============================================================================
