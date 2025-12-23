import logging
from typing import Optional, List, Any
from datetime import datetime
from .sheets import SheetsService

logger = logging.getLogger(__name__)

class LoggingService:
    def __init__(self, sheets_service: SheetsService):
        self.sheets_service = sheets_service
        self.log_sheet_name = "Логи"

    def init_session_log(self, spreadsheet_id: str):
        """
        Re-creates the 'Логи' sheet with headers.
        Called on spreadsheet load (onOpen).
        """
        try:
            # 1. Try to get or create sheet
            try:
                ws = self.sheets_service.get_worksheet(spreadsheet_id, self.log_sheet_name)
                # If exists, clear it
                ws.clear()
            except Exception:
                # If not exists, create it
                ws = self.sheets_service.create_worksheet(spreadsheet_id, self.log_sheet_name)
            
            # 2. Move to front (index 0)
            self.sheets_service.reorder_sheets(spreadsheet_id, [self.log_sheet_name, "Главная"])
            
            # 3. Set Headers
            # 🕒 Время | 🏷️ Категория | 💬 Действие | 📝 Детали | 🔘 Статус
            headers = ["🕒 Время", "🏷️ Категория", "💬 Действие", "📝 Детали", "🔘 Статус"]
            ws.append_row(headers)
            
            # Format headers
            ws.format("A1:E1", {
                "textFormat": {"bold": True},
                "backgroundColor": {"red": 0.9, "green": 0.9, "blue": 0.9}
            })
            
            self.add_log(
                spreadsheet_id, 
                "СИСТЕМА", 
                "Сессия логирования инициализирована", 
                "Лист Логи очищен и перемещен в начало", 
                "✅ СТАРТ"
            )
            
            self.add_log(
                spreadsheet_id,
                "СИСТЕМА",
                "Выстраиваем листы по порядку",
                "Логи установлен первым, остальные по стандарту",
                "🔄 В ПРОЦЕССЕ"
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
