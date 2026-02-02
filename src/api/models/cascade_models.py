from typing import List, Optional, Dict, Any
from pydantic import BaseModel, Field

class CascadeProcessRequest(BaseModel):
    """Запрос на каскадную обработку данных."""
    spreadsheet_id: str = Field(..., description="ID таблицы")
    sheet_name: str = Field("Сертификация", description="Имя листа")
    row: Optional[int] = Field(None, description="Номер строки (если для одной)")
    changed_column: Optional[str] = Field(None, description="Заголовок измененной колонки")
    new_value: Optional[str] = Field(None, description="Новое значение")
    dry_run: bool = Field(False, description="Тестовый запуск")

    class Config:
        title = "Запрос каскадной обработки"

class CascadeProcessResponse(BaseModel):
    """Ответ после каскадной обработки."""
    status: str = Field(..., description="Статус выполнения")
    row: int = Field(0, description="Номер обработанной строки")
    changes: List[Dict[str, Any]] = Field([], description="Список внесенных изменений")
    applied: bool = Field(False, description="Применены ли изменения")
    message: str = Field("", description="Сообщение")

    class Config:
        title = "Ответ каскадной обработки"
