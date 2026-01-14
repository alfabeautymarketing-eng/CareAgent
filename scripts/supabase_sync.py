#!/usr/bin/env python3
"""
=============================================================================
AGENTCARE SUPABASE SYNCHRONIZATION MODULE
=============================================================================
Синхронизация данных между Google Sheets и Supabase

Версия: 1.0
Дата: 2026-01-14
Автор: AgentCare System
=============================================================================
"""

import os
import json
import logging
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Tuple
from dataclasses import dataclass, asdict
from enum import Enum

import dotenv
from supabase import create_client, Client
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth import default as google_auth_default

# ============================================================================
# ЛОГИРОВАНИЕ И КОНФИГУРАЦИЯ
# ============================================================================

# Загрузить переменные окружения
dotenv.load_dotenv('.env.local')

# Настройка логирования
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('supabase_sync.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


class SyncDirection(Enum):
    """Направление синхронизации"""
    ONE_WAY_TO_SUPABASE = "one_way_to_supabase"
    ONE_WAY_TO_SHEETS = "one_way_to_sheets"
    TWO_WAY = "two_way"


class OperationStatus(Enum):
    """Статусы операций"""
    PENDING = "PENDING"
    IN_PROGRESS = "IN_PROGRESS"
    SUCCESS = "SUCCESS"
    PARTIAL = "PARTIAL"
    FAILED = "FAILED"


# ============================================================================
# КЛАССЫ ДАННЫХ
# ============================================================================

@dataclass
class Article:
    """Модель артикула"""
    article_id: str
    project_key: str
    article_rus: Optional[str] = None
    article_producer: Optional[str] = None
    code_base: Optional[str] = None
    code_1c: Optional[str] = None
    name_eng_price: Optional[str] = None
    name_eng_ds: Optional[str] = None
    name_rus_ds: Optional[str] = None
    volume: Optional[str] = None
    release_form: Optional[str] = None
    units_per_pack: Optional[str] = None
    category: Optional[str] = None
    line_name: Optional[str] = None
    group_name: Optional[str] = None
    barcode: Optional[str] = None
    status: str = "ACTIVE"
    notes: Optional[str] = None
    synced_from_sheet: Optional[datetime] = None
    google_sheet_id: Optional[str] = None


@dataclass
class SyncResult:
    """Результат синхронизации"""
    status: OperationStatus
    total_records: int
    processed_records: int
    failed_records: int
    skipped_records: int
    error_message: Optional[str] = None
    error_details: Optional[Dict] = None
    started_at: datetime = None
    completed_at: datetime = None

    def __post_init__(self):
        if self.started_at is None:
            self.started_at = datetime.now()
        if self.completed_at is None:
            self.completed_at = datetime.now()


# ============================================================================
# SUPABASE CLIENT WRAPPER
# ============================================================================

class SupabaseClient:
    """Wrapper для Supabase клиента"""

    def __init__(self):
        """Инициализация клиента"""
        self.url = os.getenv('SUPABASE_URL')
        self.key = os.getenv('SUPABASE_SERVICE_ROLE_KEY')

        if not self.url or not self.key:
            raise ValueError(
                "SUPABASE_URL и SUPABASE_SERVICE_ROLE_KEY "
                "должны быть установлены в .env.local"
            )

        self.client: Client = create_client(self.url, self.key)
        logger.info(f"✅ Supabase клиент инициализирован: {self.url}")

    def insert_articles(self, articles: List[Article]) -> SyncResult:
        """Вставить артикулы в таблицу"""
        logger.info(f"📥 Вставка {len(articles)} артикулов...")

        result = SyncResult(
            status=OperationStatus.IN_PROGRESS,
            total_records=len(articles),
            processed_records=0,
            failed_records=0,
            skipped_records=0
        )

        try:
            # Подготовьте данные (конвертируйте datetime в ISO string)
            articles_data = []
            for a in articles:
                data = asdict(a)
                # Конвертируйте datetime в ISO строки и удалите None значения
                cleaned_data = {}
                for key, value in data.items():
                    if value is not None:
                        if isinstance(value, datetime):
                            cleaned_data[key] = value.isoformat()
                        else:
                            cleaned_data[key] = value
                articles_data.append(cleaned_data)

            # Вставьте в батче по 100 записей
            batch_size = 100
            for i in range(0, len(articles_data), batch_size):
                batch = articles_data[i:i + batch_size]

                try:
                    response = self.client.table('articles').insert(batch).execute()
                    result.processed_records += len(batch)
                    logger.info(f"✅ Вставлено {len(batch)} артикулов")
                except Exception as e:
                    result.failed_records += len(batch)
                    logger.error(f"❌ Ошибка при вставке батча: {e}")

            result.status = (
                OperationStatus.SUCCESS
                if result.failed_records == 0
                else OperationStatus.PARTIAL
            )

        except Exception as e:
            result.status = OperationStatus.FAILED
            result.error_message = str(e)
            logger.error(f"❌ Ошибка при вставке артикулов: {e}")

        return result

    def update_articles(self, articles: List[Article]) -> SyncResult:
        """Обновить артикулы"""
        logger.info(f"🔄 Обновление {len(articles)} артикулов...")

        result = SyncResult(
            status=OperationStatus.IN_PROGRESS,
            total_records=len(articles),
            processed_records=0,
            failed_records=0,
            skipped_records=0
        )

        try:
            for article in articles:
                try:
                    # Конвертируйте datetime в ISO строки и удалите None значения
                    article_data = asdict(article)
                    data = {}
                    for key, value in article_data.items():
                        if value is not None:
                            if isinstance(value, datetime):
                                data[key] = value.isoformat()
                            else:
                                data[key] = value

                    response = (
                        self.client.table('articles')
                        .update(data)
                        .eq('article_id', article.article_id)
                        .execute()
                    )

                    result.processed_records += 1
                    logger.debug(f"✅ Обновлен артикул: {article.article_id}")

                except Exception as e:
                    result.failed_records += 1
                    logger.error(f"❌ Ошибка при обновлении {article.article_id}: {e}")

            result.status = (
                OperationStatus.SUCCESS
                if result.failed_records == 0
                else OperationStatus.PARTIAL
            )

        except Exception as e:
            result.status = OperationStatus.FAILED
            result.error_message = str(e)
            logger.error(f"❌ Ошибка при обновлении артикулов: {e}")

        return result

    def get_articles(self, project_key: str) -> List[Dict]:
        """Получить артикулы проекта"""
        try:
            response = (
                self.client.table('articles')
                .select('*')
                .eq('project_key', project_key)
                .execute()
            )
            return response.data
        except Exception as e:
            logger.error(f"❌ Ошибка при получении артикулов: {e}")
            return []

    def insert_log(self, log_data: Dict) -> bool:
        """Вставить запись в лог"""
        try:
            self.client.table('logs').insert(log_data).execute()
            return True
        except Exception as e:
            logger.error(f"❌ Ошибка при вставке лога: {e}")
            return False

    def get_sync_rules(self, enabled_only: bool = True) -> List[Dict]:
        """Получить правила синхронизации"""
        try:
            query = self.client.table('sync_rules').select('*')
            if enabled_only:
                query = query.eq('enabled', True)
            response = query.execute()
            return response.data
        except Exception as e:
            logger.error(f"❌ Ошибка при получении правил синхронизации: {e}")
            return []

    def save_sync_history(self, sync_history: Dict) -> bool:
        """Сохранить историю синхронизации"""
        try:
            self.client.table('sync_history').insert(sync_history).execute()
            return True
        except Exception as e:
            logger.error(f"❌ Ошибка при сохранении истории синхронизации: {e}")
            return False


# ============================================================================
# GOOGLE SHEETS CLIENT
# ============================================================================

class GoogleSheetsClient:
    """Клиент для работы с Google Sheets"""

    SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

    def __init__(self, credentials_file: str = 'credentials.json'):
        """Инициализация клиента"""
        self.credentials_file = credentials_file
        self.service = None
        self._authenticate()

    def _authenticate(self):
        """Аутентификация в Google Sheets API"""
        try:
            from google.oauth2.service_account import Credentials as SACredentials

            # Попробуйте использовать service account
            if os.path.exists('service-account.json'):
                creds = SACredentials.from_service_account_file(
                    'service-account.json',
                    scopes=self.SCOPES
                )
                logger.info("✅ Использован service account для Google Sheets")
            else:
                # Иначе используйте OAuth2
                creds, _ = google_auth_default(scopes=self.SCOPES)
                logger.info("✅ Использован OAuth2 для Google Sheets")

            from googleapiclient.discovery import build
            self.service = build('sheets', 'v4', credentials=creds)

        except Exception as e:
            logger.error(f"❌ Ошибка аутентификации Google Sheets: {e}")
            raise

    def get_sheet_values(self, spreadsheet_id: str, range_name: str) -> List[List]:
        """Получить значения из листа"""
        try:
            result = (
                self.service.spreadsheets()
                .values()
                .get(spreadsheetId=spreadsheet_id, range=range_name)
                .execute()
            )
            return result.get('values', [])
        except Exception as e:
            logger.error(f"❌ Ошибка при получении данных из Sheets: {e}")
            return []

    def parse_articles_from_sheet(
        self,
        spreadsheet_id: str,
        sheet_range: str = "Главная!A1:Z1000",
        project_key: str = "MT"
    ) -> List[Article]:
        """Распарсить артикулы из листа Sheets"""
        logger.info(f"📊 Чтение артикулов из Sheets: {sheet_range}")

        values = self.get_sheet_values(spreadsheet_id, sheet_range)

        if not values or len(values) < 2:
            logger.warning("⚠️  Нет данных в листе")
            return []

        # Первая строка - заголовки
        headers = values[0]
        articles = []

        # Данные начинаются со второй строки
        for row_idx, row in enumerate(values[1:], start=2):
            try:
                # Создайте словарь для маппинга
                row_data = {}
                for col_idx, header in enumerate(headers):
                    if col_idx < len(row):
                        row_data[header] = row[col_idx]
                    else:
                        row_data[header] = None

                # Создайте объект Article
                article = Article(
                    article_id=row_data.get('ID', f"article_{row_idx}"),
                    project_key=project_key,
                    article_rus=row_data.get('Арт. Рус'),
                    article_producer=row_data.get('Арт. произв.'),
                    code_base=row_data.get('Код база'),
                    code_1c=row_data.get('Код в 1С'),
                    name_eng_price=row_data.get('Название  ENG прайс произв'),
                    name_eng_ds=row_data.get('Наименования англ по ДС'),
                    name_rus_ds=row_data.get('Наименования рус по ДС'),
                    volume=row_data.get('Объём'),
                    release_form=row_data.get('Форма выпуска'),
                    units_per_pack=row_data.get('шт./уп.'),
                    category=row_data.get('Категория товара'),
                    line_name=row_data.get('Линия'),
                    group_name=row_data.get('Группа'),
                    barcode=row_data.get('BAR CODE'),
                    google_sheet_id=spreadsheet_id,
                    synced_from_sheet=datetime.now()
                )

                articles.append(article)

            except Exception as e:
                logger.warning(f"⚠️  Ошибка при парсинге строки {row_idx}: {e}")
                continue

        logger.info(f"✅ Распарсено {len(articles)} артикулов из Sheets")
        return articles


# ============================================================================
# СИНХРОНИЗАЦИЯ
# ============================================================================

class DataSynchronizer:
    """Основной класс для синхронизации"""

    def __init__(self):
        """Инициализация"""
        self.supabase = SupabaseClient()
        self.sheets = GoogleSheetsClient()

    def sync_articles_from_sheets(
        self,
        spreadsheet_id: str,
        sheet_range: str,
        project_key: str,
        upsert: bool = True
    ) -> SyncResult:
        """Синхронизировать артикулы из Google Sheets в Supabase"""

        logger.info(
            f"🔄 Начало синхронизации артикулов "
            f"({project_key}): {spreadsheet_id}"
        )

        # 1. Прочитайте данные из Sheets
        articles = self.sheets.parse_articles_from_sheet(
            spreadsheet_id,
            sheet_range,
            project_key
        )

        if not articles:
            return SyncResult(
                status=OperationStatus.FAILED,
                total_records=0,
                processed_records=0,
                failed_records=0,
                skipped_records=0,
                error_message="Нет артикулов для синхронизации"
            )

        # 2. Получите существующие артикулы в Supabase
        existing = {a['article_id']: a for a in self.supabase.get_articles(project_key)}

        # 3. Разделите на новые и обновляемые
        new_articles = [a for a in articles if a.article_id not in existing]
        update_articles = [a for a in articles if a.article_id in existing]

        result = SyncResult(
            status=OperationStatus.IN_PROGRESS,
            total_records=len(articles),
            processed_records=0,
            failed_records=0,
            skipped_records=0
        )

        # 4. Вставьте новые
        if new_articles:
            insert_result = self.supabase.insert_articles(new_articles)
            result.processed_records += insert_result.processed_records
            result.failed_records += insert_result.failed_records

        # 5. Обновите существующие
        if update_articles and upsert:
            update_result = self.supabase.update_articles(update_articles)
            result.processed_records += update_result.processed_records
            result.failed_records += update_result.failed_records

        # 6. Установите финальный статус
        result.status = (
            OperationStatus.SUCCESS
            if result.failed_records == 0
            else OperationStatus.PARTIAL
        )
        result.completed_at = datetime.now()

        # 7. Сохраните историю
        sync_history = {
            'total_records': result.total_records,
            'processed_records': result.processed_records,
            'failed_records': result.failed_records,
            'skipped_records': result.skipped_records,
            'status': result.status.value,
            'started_at': result.started_at.isoformat() if result.started_at else None,
            'completed_at': result.completed_at.isoformat() if result.completed_at else None,
            'sync_batch_id': f"{project_key}_{datetime.now().isoformat()}"
        }
        self.supabase.save_sync_history(sync_history)

        logger.info(f"✅ Синхронизация завершена: {result.status.value}")
        return result

    def sync_logs_from_sheets(
        self,
        spreadsheet_id: str,
        sheet_range: str = "LOG_DEBUG!A1:H1000"
    ) -> SyncResult:
        """Синхронизировать логи из Google Sheets в Supabase"""

        logger.info(f"📋 Начало синхронизации логов: {sheet_range}")

        values = self.sheets.get_sheet_values(spreadsheet_id, sheet_range)

        if not values or len(values) < 2:
            return SyncResult(
                status=OperationStatus.FAILED,
                total_records=0,
                processed_records=0,
                failed_records=0,
                skipped_records=0,
                error_message="Нет логов для синхронизации"
            )

        headers = values[0]
        result = SyncResult(
            status=OperationStatus.IN_PROGRESS,
            total_records=len(values) - 1,
            processed_records=0,
            failed_records=0,
            skipped_records=0
        )

        for row_idx, row in enumerate(values[1:], start=2):
            try:
                row_data = {}
                for col_idx, header in enumerate(headers):
                    if col_idx < len(row):
                        row_data[header] = row[col_idx]

                # Создайте запись лога
                timestamp_str = row_data.get('Время', datetime.now().isoformat())
                try:
                    # Если это строка ISO, преобразуйте в datetime и обратно в строку
                    if isinstance(timestamp_str, str):
                        timestamp_obj = datetime.fromisoformat(timestamp_str)
                        timestamp_iso = timestamp_obj.isoformat()
                    else:
                        timestamp_iso = datetime.now().isoformat()
                except:
                    timestamp_iso = datetime.now().isoformat()

                log_entry = {
                    'timestamp': timestamp_iso,
                    'category': row_data.get('Категория'),
                    'status': row_data.get('Статус'),
                    'message': row_data.get('Сообщение'),
                    'function_name': row_data.get('Функция'),
                    'details': row_data.get('Детали'),
                    'variables': json.loads(row_data.get('Параметры (JSON)', '{}')) if row_data.get('Параметры (JSON)') else None,
                    'result': json.loads(row_data.get('Результаты (JSON)', '{}')) if row_data.get('Результаты (JSON)') else None,
                }

                if self.supabase.insert_log(log_entry):
                    result.processed_records += 1
                else:
                    result.failed_records += 1

            except Exception as e:
                result.failed_records += 1
                logger.warning(f"⚠️  Ошибка при обработке лога в строке {row_idx}: {e}")

        result.status = (
            OperationStatus.SUCCESS
            if result.failed_records == 0
            else OperationStatus.PARTIAL
        )
        result.completed_at = datetime.now()

        logger.info(f"✅ Синхронизация логов завершена: {result.status.value}")
        return result


# ============================================================================
# MAIN FUNCTION
# ============================================================================

def main():
    """Основная функция"""

    logger.info("=" * 80)
    logger.info("AGENTCARE SUPABASE SYNCHRONIZATION")
    logger.info("=" * 80)

    try:
        synchronizer = DataSynchronizer()

        # Синхронизируйте артикулы для каждого проекта
        spreadsheet_ids = {
            'MT': '13kB77R67GJOZQ3vsLcwR1nUaRsupR8ZnEaTdDd66CTQ',
            'SS': '12yIL1CuESZxeUUd-oKK2brtN1FnXE9q95N7SqzNc7vk',
            'SK': '1CpYYLvRYslsyCkuLzL9EbbjsvbNpWCEZcmhKqMoX5zw'
        }

        for project_key, spreadsheet_id in spreadsheet_ids.items():
            logger.info(f"\n📦 Синхронизация проекта {project_key}...")
            result = synchronizer.sync_articles_from_sheets(
                spreadsheet_id=spreadsheet_id,
                sheet_range=f"Главная!A1:Z10000",
                project_key=project_key,
                upsert=True
            )

            logger.info(f"   Всего: {result.total_records}")
            logger.info(f"   Обработано: {result.processed_records}")
            logger.info(f"   Ошибок: {result.failed_records}")
            logger.info(f"   Статус: {result.status.value}")

        logger.info("\n" + "=" * 80)
        logger.info("✅ СИНХРОНИЗАЦИЯ ЗАВЕРШЕНА")
        logger.info("=" * 80)

    except Exception as e:
        logger.error(f"❌ КРИТИЧЕСКАЯ ОШИБКА: {e}")
        exit(1)


if __name__ == '__main__':
    main()
