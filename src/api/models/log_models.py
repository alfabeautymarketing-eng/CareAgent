from typing import Optional, List, Dict
from pydantic import BaseModel, Field

class LogArchiveRequest(BaseModel):
    """Запрос на архивацию логов."""
    spreadsheet_id: str = Field(..., description="ID таблицы")
    archive_folder_id: str = Field(..., description="ID папки на Google Диске для архивов")
    project_code: str = Field("project", description="Код проекта")
    dry_run: bool = Field(False, description="Тестовый запуск")

    class Config:
        title = "Запрос архивации логов"

class LogResetRequest(BaseModel):
    """Запрос на сброс листа логов."""
    spreadsheet_id: str = Field(..., description="ID таблицы")
    sheet_name: str = Field("Логи", description="Имя листа")
    dry_run: bool = Field(False, description="Тестовый запуск")

    class Config:
        title = "Запрос очистки логов"

class LogRotationRequest(BaseModel):
    """Запрос на ночную ротацию логов."""
    spreadsheet_id: str = Field(..., description="ID таблицы")
    archive_folder_id: str = Field(..., description="ID папки на Google Диске")
    project_code: str = Field("project", description="Код проекта")
    force: bool = Field(False, description="Принудительная ротация")
    dry_run: bool = Field(False, description="Тестовый запуск")

    class Config:
        title = "Запрос ротации логов"

class LogEntryRequest(BaseModel):
    """Запрос на запись одной строки в лог."""
    spreadsheet_id: str = Field(..., description="ID таблицы")
    sheet_name: str = Field("Логи", description="Имя листа")
    category: str = Field("SYSTEM", description="Категория")
    action: str = Field(..., description="Действие")
    details: str = Field("", description="Подробности")
    level: str = Field("INFO", description="Уровень (INFO, WARN, ERROR)")

    class Config:
        title = "Запрос записи в лог"

class LogArchiveResponse(BaseModel):
    """Ответ после архивации логов."""
    status: str = Field(..., description="Статус")
    total_rows: int = Field(0, description="Всего строк")
    sheets_archived: Optional[Dict[str, int]] = Field(None, description="Архивированные листы")
    archive_name: Optional[str] = Field(None, description="Название архива")
    message: str = Field("", description="Сообщение")

    class Config:
        title = "Ответ архивации логов"

class LogStatusResponse(BaseModel):
    """Ответ о статусе архивации логов."""
    status: str = Field(..., description="Статус")
    last_archive_date: str = Field("Never", description="Дата последней архивации")
    current_archive_name: str = Field("", description="Имя текущего архива")
    project_code: str = Field("", description="Код проекта")
    sheets_to_archive: List[str] = Field([], description="Листы для архивации")
    current_row_counts: Optional[Dict[str, int]] = Field(None, description="Количество строк в листах")
    total_pending_rows: int = Field(0, description="Всего строк к архивации")

    class Config:
        title = "Ответ о статусе логов"
