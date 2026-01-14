from typing import Optional, List, Any, Dict
from datetime import datetime
from .sheets import SheetsService
from src.utils.logger import logger
from src.services.drive import get_drive_service
import asyncio

class LoggingService:
    def __init__(self, sheets_service: SheetsService):
        self.sheets_service = sheets_service
        # Names kept for potential reference, but not used for writing
        self.log_sheet_name = "Логи"
        
    def init_session_log(self, spreadsheet_id: str):
        """
        No-op: Session logs on sheet are disabled.
        """
        return True

    def recreate_log_sheet(self, spreadsheet_id: str, force_clear: bool = False, sheet_name: Optional[str] = None):
        """
        No-op: Sheet logging disabled.
        """
        return True

    async def archive_logs(self, spreadsheet_id: str, project_prefix: str = "Common") -> Dict[str, Any]:
        """
        No-op: Archiving disabled.
        """
        return {"success": True, "message": "Archiving disabled"}

    def get_archive_status(self, project_prefix: str = "Common") -> Dict[str, Any]:
        """
        No-op.
        """
        return {"exists": False, "message": "Archiving disabled"}

    def add_log(self, spreadsheet_id: str, category: str, action: str, details: str, status: str):
        """
        Log to system logger only. No sheet writing.
        """
        level = "ERROR" if "ОШИБКА" in status or status == "❌ ERR" else "INFO"
        
        log_msg = f"[{category}] {action} | {details} | {status}"
        
        if level == "ERROR":
            logger.error("system_log_entry", spreadsheet_id=spreadsheet_id, msg=log_msg)
        else:
            logger.info("system_log_entry", spreadsheet_id=spreadsheet_id, msg=log_msg)

    def add_summary_log(
        self,
        spreadsheet_id: str,
        category: str,
        action: str,
        details: str = "",
        status: str = "✅ OK",
        extra_data: Optional[Any] = None,
    ):
        """
        Redirect to add_log (system logger).
        """
        if details and len(details) > 200:
            details = details[:197] + "..."
        if extra_data:
            details = f"{details} | {str(extra_data)[:100]}"

        self.add_log(spreadsheet_id, category, action, details, status)

    def log_sync_summary(
        self,
        spreadsheet_id: str,
        source_sheet: str,
        target_sheet: str,
        status: str = "success",
        details: str = "",
        **kwargs
    ):
        """
        Log sync summary to system logger.
        """
        category = "СИНХРО"
        action = f"Синхро: {source_sheet} → {target_sheet}"
        status_emoji = "✅ OK" if status == "success" else "⚠️ ОШИБКА" if status == "error" else "⏳ ОЖИДАНИЕ"

        if not details and kwargs:
            details = " | ".join(f"{k}={v}" for k, v in list(kwargs.items())[:3])

        self.add_log(spreadsheet_id, category, action, details, status_emoji)

    def log_function_summary(
        self,
        spreadsheet_id: str,
        function_name: str,
        status: str = "completed",
        step_count: int = 0,
        duration_ms: Optional[float] = None,
        error: Optional[str] = None,
    ):
        """
        Log function summary to system logger.
        """
        category = "ФУНКЦИЯ"
        action = function_name
        status_emoji = "✅ OK" if status == "completed" else "❌ ОШИБКА"

        details = f"Шаги: {step_count}"
        if duration_ms:
            details += f" | Время: {duration_ms:.0f}ms"
        if error:
            details += f" | Ошибка: {error[:50]}"

        self.add_log(spreadsheet_id, category, action, details, status_emoji)

logging_service = None # Dependency injection hint
