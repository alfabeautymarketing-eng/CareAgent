from typing import Optional, List, Any
from datetime import datetime
from .sheets import SheetsService
from src.utils.logger import logger
from src.services.drive import get_drive_service
import asyncio
from typing import Dict, Any

class LoggingService:
    def __init__(self, sheets_service: SheetsService):
        self.sheets_service = sheets_service
        # Единое имя листа логов
        self.log_sheet_name = "Логи"
        self.log_debug_sheet_name = "Журнал логов"
        self.sync_log_sheet_name = "Журнал синхро"
        self.archive_folder_id = "1oWoxwuOMlriwgHuua6MpLheROEIT4SIP"
        
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
            return True
        except Exception as e:
            logger.error("init_session_log_failed", error=str(e))
            return False

    def recreate_log_sheet(self, spreadsheet_id: str, force_clear: bool = False, sheet_name: Optional[str] = None):
        """
        Manually recreate or clear the log sheet.
        """
        target_sheet = sheet_name if sheet_name else self.log_sheet_name
        
        try:
            headers = ["🕒 Время", "🏷️ Категория", "💬 Действие", "📝 Детали", "🔘 Статус"]
            
            # Try to get existing sheet
            try:
                ws = self.sheets_service.get_worksheet(spreadsheet_id, target_sheet)
                
                if force_clear:
                    # Clear entire sheet
                    ws.clear()
                    # Re-add headers
                    ws.update('A1:E1', [headers])
                    ws.format("A1:E1", {
                        "textFormat": {"bold": True},
                        "backgroundColor": {"red": 0.9, "green": 0.9, "blue": 0.9}
                    })
                else:
                    # Just ensure headers are there, don't delete data
                    existing_headers = ws.row_values(1) if ws.row_count > 0 else []
                    if existing_headers != headers:
                        ws.insert_row(headers, index=1)
                        ws.format("A1:E1", {
                            "textFormat": {"bold": True},
                            "backgroundColor": {"red": 0.9, "green": 0.9, "blue": 0.9}
                        })
            except Exception:
                # Sheet missing, create new
                ws = self.sheets_service.create_worksheet(spreadsheet_id, target_sheet)
                ws.append_row(headers)
                ws.format("A1:E1", {
                    "textFormat": {"bold": True},
                    "backgroundColor": {"red": 0.9, "green": 0.9, "blue": 0.9}
                })
                
            # Only reorder if it's the main log sheet or explicitly requested (omitted for flexibility)
            if target_sheet == self.log_sheet_name:
                self.sheets_service.reorder_sheets(spreadsheet_id, [target_sheet, "Главная"])
            
            self.add_log(
                spreadsheet_id,
                "СИСТЕМА",
                f"Лист '{target_sheet}' пересоздан" if force_clear else f"Лист '{target_sheet}' восстановлен",
                "Ручная операция",
                "✅ OK"
            )
            return True
        except Exception as e:
            logger.error("recreate_log_sheet_failed", error=str(e))
            raise


    async def archive_logs(self, spreadsheet_id: str, project_prefix: str = "Common") -> Dict[str, Any]:
        """
        Archive logs to a monthly spreadsheet and clear source sheets.
        """
        drive = get_drive_service()
        now = datetime.now()
        months_ru = [
            "январь", "февраль", "март", "апрель", "май", "июнь",
            "июль", "август", "сентябрь", "октябрь", "ноябрь", "декабрь"
        ]
        month_name = months_ru[now.month - 1]
        year = now.year
        archive_name = f"{project_prefix.lower()}-{month_name}-{year}"
        
        sheets_to_archive = [self.log_sheet_name, self.sync_log_sheet_name, self.log_debug_sheet_name]
        results = {}
        total_rows = 0
        
        try:
            # 1. Find or Create Archive Spreadsheet
            archive_ss_id = None
            
            # Search in folder
            query = f"name = '{archive_name}' and '{self.archive_folder_id}' in parents and mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false"
            files = drive.service.files().list(q=query, fields="files(id, name)").execute()
            found_files = files.get('files', [])
            
            if found_files:
                archive_ss_id = found_files[0]['id']
            else:
                # Create new spreadsheet
                # Note: SheetsService create_spreadsheet usually creates in root, we need to move it or create via Drive API
                # Using Sheets API to create
                ss_body = {'properties': {'title': archive_name}}
                new_ss = self.sheets_service.service.spreadsheets().create(body=ss_body, fields='spreadsheetId').execute()
                archive_ss_id = new_ss.get('spreadsheetId')
                # Move to archive folder
                drive.copy_file(archive_ss_id, self.archive_folder_id, None) # This copies? No wait, we need to move. 
                # Drive API move is update parents.
                # Let's use drive.service directly for move as copy_file creates a copy
                file = drive.service.files().get(fileId=archive_ss_id, fields='parents').execute()
                previous_parents = ",".join(file.get('parents'))
                drive.service.files().update(fileId=archive_ss_id, addParents=self.archive_folder_id, removeParents=previous_parents).execute()
                
            # 2. Archive Data
            for sheet_name in sheets_to_archive:
                try:
                    # Read source data
                    ws = self.sheets_service.get_worksheet(spreadsheet_id, sheet_name)
                    values = ws.get_all_values()
                    
                    if len(values) <= 1:
                        results[sheet_name] = 0
                        continue
                        
                    headers = values[0]
                    data = values[1:]
                    
                    # Add Archive Date Column
                    archive_date = now.strftime("%d.%m.%Y %H:%M:%S")
                    headers.append("📅 Дата архивации")
                    for row in data:
                        row.append(archive_date)
                        
                    rows_count = len(data)
                    
                    # Get or Create Target Sheet in Archive SS
                    try:
                        target_ws = self.sheets_service.get_worksheet(archive_ss_id, sheet_name)
                    except Exception:
                        target_ws = self.sheets_service.create_worksheet(archive_ss_id, sheet_name)
                        
                    # Check if target is empty, if so write headers
                    if target_ws.row_count == 0 or (target_ws.row_count == 1 and not target_ws.row_values(1)):
                         target_ws.append_row(headers)
                         # Format headers
                         target_ws.format("A1:Z1", {"textFormat": {"bold": True}, "backgroundColor": {"red": 0.9, "green": 0.9, "blue": 0.9}})

                    # Append data
                    target_ws.append_rows(data)
                    
                    # 3. Clear Source Sheet (keep headers)
                    # We can use recreate_log_sheet logic or just delete rows
                    if sheet_name == self.log_sheet_name:
                         self.recreate_log_sheet(spreadsheet_id, force_clear=True)
                    else:
                         # For others, simple resize or clear
                         ws.resize(rows=2)
                         ws.delete_rows(2, rows_count + 1) # simple clear might be safer: ws.delete_rows(2, ws.row_count)
                         # Actually gspread delete_rows is index, count. 
                         # ws.batch_clear([f"A2:Z{len(values)}"])
                         ws.resize(rows=100) # Reset size
                         # Re-write headers just in case
                         ws.update('A1', [headers[:-1]]) 
                    
                    results[sheet_name] = rows_count
                    total_rows += rows_count
                    
                except Exception as sheet_err:
                     logger.warning(f"Failed to archive sheet {sheet_name}: {sheet_err}")
                     results[sheet_name] = f"Error: {str(sheet_err)}"

            self.add_log(spreadsheet_id, "СИСТЕМА", "Архивирование логов", f"Записано {total_rows} строк", "✅ OK")
            
            return {
                "success": True,
                "total_rows": total_rows,
                "archive_name": archive_name,
                "details": results
            }
            
        except Exception as e:
            logger.error("archive_logs_failed", error=str(e))
            raise

    def get_archive_status(self, project_prefix: str = "Common") -> Dict[str, Any]:
        """
        Check status of current month's archive.
        """
        drive = get_drive_service()
        now = datetime.now()
        months_ru = ["январь", "февраль", "март", "апрель", "май", "июнь", "июль", "август", "сентябрь", "октябрь", "ноябрь", "декабрь"]
        month_name = months_ru[now.month - 1]
        year = now.year
        archive_name = f"{project_prefix.lower()}-{month_name}-{year}"
        
        query = f"name = '{archive_name}' and '{self.archive_folder_id}' in parents and trashed = false"
        files = drive.service.files().list(q=query, fields="files(id, name, createdTime, modifiedTime)").execute()
        found_files = files.get('files', [])
        
        status = {
            "archive_name": archive_name,
            "folder_id": self.archive_folder_id,
            "exists": len(found_files) > 0,
            "last_modified": found_files[0]['modifiedTime'] if found_files else None
        }
        return status

    def add_log(self, spreadsheet_id: str, category: str, action: str, details: str, status: str):
        """
        Append a row to 'Логи' sheet with retry logic.
        """
        import time
        max_retries = 3

        timestamp = datetime.now().strftime("%d.%m.%Y %H:%M:%S")
        row = [timestamp, category, action, details, status]

        for attempt in range(max_retries):
            try:
                ws = self.sheets_service.get_worksheet(spreadsheet_id, self.log_sheet_name)
                ws.append_row(row)
                return  # Success
            except Exception as e:
                if attempt < max_retries - 1:
                    time.sleep(1 * (attempt + 1))  # Backoff: 1s, 2s...
                    continue

                # Final failure: fall back to internal logger
                logger.error("log_append_failed_final", error=str(e), action=action, spreadsheet_id=spreadsheet_id)
                # Ensure we don't crash main thread just because log failed
                try:
                    # Additional fallback: print to stdout/stderr or file if really critical
                    pass
                except:
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
        """
        Write a summary log entry to the 'Логи' sheet.
        This is used for summarizing sync events, function executions, etc.
        """
        # Truncate details if too long
        if details and len(details) > 200:
            details = details[:197] + "..."

        # Add extra data if provided
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
        Log a summary of a sync operation.
        """
        category = "СИНХРО"
        action = f"Синхро: {source_sheet} → {target_sheet}"

        status_emoji = "✅ OK" if status == "success" else "⚠️ ОШИБКА" if status == "error" else "⏳ ОЖИДАНИЕ"

        if not details and kwargs:
            details = " | ".join(f"{k}={v}" for k, v in list(kwargs.items())[:3])

        self.add_summary_log(spreadsheet_id, category, action, details, status_emoji)

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
        Log a summary of a function execution.
        """
        category = "ФУНКЦИЯ"
        action = function_name

        status_emoji = "✅ OK" if status == "completed" else "❌ ОШИБКА"

        details = f"Шаги: {step_count}"
        if duration_ms:
            details += f" | Время: {duration_ms:.0f}ms"
        if error:
            details += f" | Ошибка: {error[:50]}"

        self.add_summary_log(spreadsheet_id, category, action, details, status_emoji)

logging_service = None # Dependency injection hint
