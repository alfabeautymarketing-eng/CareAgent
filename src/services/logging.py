from typing import Optional, List, Any
from datetime import datetime
from .sheets import SheetsService
from src.utils.logger import logger

class LoggingService:
    def __init__(self, sheets_service: SheetsService):
        self.sheets_service = sheets_service
        # Единое имя листа логов
        self.log_sheet_name = "Логи"
        # Даем доступ SheetsService к логгеру, чтобы сервисы могли логировать изнутри
        self.sheets_service.logging_service = self
        self.sheets_service.log_sheet_name = self.log_sheet_name

    def init_session_log(self, spreadsheet_id: str):
        """
        Ensures the 'Логи' sheet exists with proper headers.
        Called on spreadsheet load (onOpen).

        ВАЖНО: Лист НЕ очищается при каждом открытии!
        Очистка происходит только в полночь через триггер midnightLogRotation_trigger
        после архивирования в месячную таблицу.
        """
        try:
            headers = ["🕒 Время", "🏷️ Категория", "💬 Действие", "📝 Детали", "🔘 Статус"]

            # 1. Try to get existing sheet or create new one
            try:
                ws = self.sheets_service.get_worksheet(spreadsheet_id, self.log_sheet_name)
                # Sheet exists - check if headers are correct
                existing_headers = ws.row_values(1) if ws.row_count > 0 else []
                if existing_headers != headers:
                    # Fix headers if they're wrong
                    ws.update('A1:E1', [headers])
                    ws.format("A1:E1", {
                        "textFormat": {"bold": True},
                        "backgroundColor": {"red": 0.9, "green": 0.9, "blue": 0.9}
                    })
            except Exception:
                # Sheet doesn't exist - create it
                ws = self.sheets_service.create_worksheet(spreadsheet_id, self.log_sheet_name)
                ws.append_row(headers)
                ws.format("A1:E1", {
                    "textFormat": {"bold": True},
                    "backgroundColor": {"red": 0.9, "green": 0.9, "blue": 0.9}
                })

            # 2. Move to front (index 0)
            self.sheets_service.reorder_sheets(spreadsheet_id, [self.log_sheet_name, "Главная"])

            # 3. Log session start (append, not replace)
            self.add_log(
                spreadsheet_id,
                "СИСТЕМА",
                "Сессия открыта",
                "Пользователь открыл таблицу",
                "✅ СТАРТ"
            )

            return True
        except Exception as e:
            logger.error("init_session_log_failed", error=str(e))
            return False

    def add_log(self, spreadsheet_id: str, category: str, action: str, details: str, status: str):
        """
        Append a row to 'Логи' sheet.
        """
        try:
            ws = self.sheets_service.get_worksheet(spreadsheet_id, self.log_sheet_name)
            timestamp = datetime.now().strftime("%d.%m.%Y %H:%M:%S")
            row = [timestamp, category, action, details, status]
            ws.append_row(row)
        except Exception as e:
            # Fallback to internal logger if sheet logging fails
            logger.error("internal_log_failed", error=str(e), action=action)

logging_service = None # Dependency injection hint
