from typing import Optional, List, Any, Dict
from datetime import datetime
from .sheets import SheetsService
from src.utils.logger import logger
from src.services.drive import get_drive_service
import asyncio

if TYPE_CHECKING:
    from src.services.sync_log_service import SyncLogService

class LoggingService:
    def __init__(self, sheets_service: SheetsService, sync_log_service: Optional["SyncLogService"] = None):
        """
        Initialize the logger.
        
        Args:
            sheets_service: Service for interacting with Google Sheets
            sync_log_service: Optional service for persistent sync logs (dashboard)
        """
        self.sheets_service = sheets_service
        self.sync_log_service = sync_log_service
        
    def init_session_log(self, spreadsheet_id: str):
        """
        DEPRECATED: Initialize session log sheet.
        Now a no-op as all logging is server-side.
        """
        pass

    def recreate_log_sheet(self, spreadsheet_id: str, debug: bool = False):
        """
        DEPRECATED: Recreate the log sheet.
        Now a no-op as all logging is server-side.
        """
        pass

    async def archive_logs(self, spreadsheet_id: str) -> dict:
        """
        Archive logs locally if needed (managed by SyncLogService usually).
        This method is kept for API compatibility but might be redundant.
        """
        logger.info("logging_archive_requested", spreadsheet_id=spreadsheet_id)
        return {"status": "ok", "message": "Archiving handled by server logs"}

    def add_log(
        self,
        spreadsheet_id: str,
        message: str,
        level: str = "INFO",
        category: str = "General",
        status: str = "NB",
        details: str = "",
        function_name: str = "",
    ):
        """
        Add a log entry. 
        Logs to system logger and optionally to SyncLogService for dashboard visibility.
        """
        logger.info(
            "general_log",
            spreadsheet_id=spreadsheet_id,
            level=level,
            message=message,
            category=category,
            status=status,
            details=details
        )
        
        if self.sync_log_service:
            # Map general log fields to SyncLogEntry fields
            # We use 'source_info' for function name + message
            # 'status' is mapped directly
            self.sync_log_service.add_entry(
                spreadsheet_id=spreadsheet_id,
                source_info=f"{function_name}: {message}",
                target_info=details,
                category=category,
                status=status,
                project="SYSTEM",
                event="LOG",
                extra={"level": level}
            )

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
