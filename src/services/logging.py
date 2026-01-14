from typing import Optional, List, Any
from datetime import datetime
from .sheets import SheetsService
from src.utils.logger import logger
from src.services.drive import get_drive_service
import asyncio
from typing import Dict, Any

class LoggingService:
    """
    ЗАГЛУШКА (STUB) - Логирование в лист 'Логи' отключено.
    Все методы пустые, чтобы избежать ошибок при вызовах из других частей кода.
    """
    def __init__(self, sheets_service: SheetsService):
        self.sheets_service = sheets_service
        # Сохраняем атрибуты для совместимости
        self.log_sheet_name = "Логи"
        self.log_debug_sheet_name = "Журнал логов"
        self.sync_log_sheet_name = "Журнал синхро"
        self.archive_folder_id = "1oWoxwuOMlriwgHuua6MpLheROEIT4SIP"
        
        # ОТКЛЮЧЕНО: Больше не даем доступ к логгеру
        # self.sheets_service.logging_service = self
        # self.sheets_service.log_sheet_name = self.log_sheet_name

    def init_session_log(self, spreadsheet_id: str):
        """ОТКЛЮЧЕНО: Лист 'Логи' больше не создается."""
        return True

    def recreate_log_sheet(self, spreadsheet_id: str, force_clear: bool = False, sheet_name: Optional[str] = None):
        """ОТКЛЮЧЕНО: Лист 'Логи' не пересоздается."""
        return True

    async def archive_logs(self, spreadsheet_id: str, project_prefix: str = "Common") -> Dict[str, Any]:
        """ОТКЛЮЧЕНО: Архивирование логов не выполняется."""
        return {
            "success": True,
            "total_rows": 0,
            "archive_name": "disabled",
            "details": {}
        }

    def get_archive_status(self, project_prefix: str = "Common") -> Dict[str, Any]:
        """ОТКЛЮЧЕНО: Статус архива недоступен."""
        return {
            "archive_name": "disabled",
            "folder_id": self.archive_folder_id,
            "exists": False,
            "last_modified": None
        }

    def add_log(self, spreadsheet_id: str, category: str, action: str, details: str, status: str):
        """ОТКЛЮЧЕНО: Логи не записываются в лист."""
        pass

    def add_summary_log(
        self,
        spreadsheet_id: str,
        category: str,
        action: str,
        details: str = "",
        status: str = "✅ OK",
        extra_data: Optional[Any] = None,
    ):
        """ОТКЛЮЧЕНО: Сводные логи не записываются."""
        pass

    def log_sync_summary(
        self,
        spreadsheet_id: str,
        source_sheet: str,
        target_sheet: str,
        status: str = "success",
        details: str = "",
        **kwargs
    ):
        """ОТКЛЮЧЕНО: Логирование синхронизации отключено."""
        pass

    def log_function_summary(
        self,
        spreadsheet_id: str,
        function_name: str,
        status: str = "completed",
        step_count: int = 0,
        duration_ms: Optional[float] = None,
        error: Optional[str] = None,
    ):
        """ОТКЛЮЧЕНО: Логирование функций отключено."""
        pass
    
    # Дополнительные методы для совместимости (если вызываются из других мест)
    async def archive_logs_daily(self, *args, **kwargs):
        """ОТКЛЮЧЕНО"""
        return {"success": True}
    
    async def reset_daily_log_sheet(self, *args, **kwargs):
        """ОТКЛЮЧЕНО"""
        return {"success": True}
    
    async def midnight_log_rotation(self, *args, **kwargs):
        """ОТКЛЮЧЕНО"""
        return {"success": True}
    
    async def write_log_entry(self, *args, **kwargs):
        """ОТКЛЮЧЕНО"""
        return {"success": True}

logging_service = None # Dependency injection hint
