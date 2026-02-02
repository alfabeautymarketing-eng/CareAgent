from typing import Optional
from pydantic import BaseModel, Field

class ExportRequest(BaseModel):
    """Запрос на экспорт данных."""
    spreadsheet_id: str = Field(..., description="ID исходной таблицы")
    project: str = Field(..., description="Код проекта (mt, sk, ss)")
    target_doc_id: Optional[str] = Field(None, description="ID целевого документа (переопределение)")
    dry_run: bool = Field(False, description="Тестовый запуск")

    class Config:
        title = "Запрос на экспорт"

class ExportResponse(BaseModel):
    """Ответ после выполнения экспорта."""
    status: str = Field(..., description="Статус выполнения")
    export_type: str = Field(..., description="Тип экспорта")
    exported_rows: int = Field(0, description="Количество экспортированных строк")
    target_doc_id: str = Field("", description="ID целевого документа")
    target_sheet_name: str = Field("", description="Имя целевого листа")
    target_url: str = Field("", description="Ссылка на целевой документ")
    message: str = Field("", description="Сообщение результата")

    class Config:
        title = "Ответ на экспорт"
