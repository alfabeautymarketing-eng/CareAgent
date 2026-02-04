from typing import Optional, List
from pydantic import BaseModel, Field

class RuleItem(BaseModel):
    mode: str = Field("unidirectional", description="Режим синхронизации: в одну сторону или в обе")
    enabled: bool = Field(True, description="Включено ли правило")
    category: str = Field("", description="Категория данных")
    source_sheet: Optional[str] = Field(None, description="Лист-источник")
    source_header: Optional[str] = Field(None, description="Заголовок-источник")
    target_sheet: Optional[str] = Field(None, description="Лист-назначение")
    target_header: Optional[str] = Field(None, description="Заголовок-назначение")
    sheet_a: Optional[str] = Field(None, description="Лист A")
    header_a: Optional[str] = Field(None, description="Заголовок A")
    sheet_b: Optional[str] = Field(None, description="Лист B")
    header_b: Optional[str] = Field(None, description="Заголовок B")
    is_external: bool = Field(False, description="Внешняя ли таблица")
    target_doc_id: Optional[str] = Field(None, description="ID внешней таблицы")

    class Config:
        title = "Элемент правила"

class RulesSaveRequest(BaseModel):
    rules: List[RuleItem]

    class Config:
        title = "Запрос на сохранение правил"

class RuleCreateRequest(BaseModel):
    """Запрос на создание нового правила синхронизации."""
    mode: str = Field("unidirectional", description="Режим (unidirectional/bidirectional)")
    enabled: bool = Field(True, description="Включено ли")
    category: str = Field("", description="Категория (например, 'Pricing')")
    source_sheet: Optional[str] = Field(None, description="Лист-источник")
    source_header: Optional[str] = Field(None, description="Заголовок-источник")
    target_sheet: Optional[str] = Field(None, description="Целевой лист")
    target_header: Optional[str] = Field(None, description="Целевой заголовок")
    sheet_a: Optional[str] = Field(None, description="Лист A")
    header_a: Optional[str] = Field(None, description="Заголовок A")
    sheet_b: Optional[str] = Field(None, description="Лист B")
    header_b: Optional[str] = Field(None, description="Заголовок B")
    is_external: bool = Field(False, description="Внешняя ли таблица")
    target_doc_id: Optional[str] = Field(None, description="ID внешней таблицы")
    projects: List[str] = Field(default=["ALL"], description="Список проектов (например ['ss', 'mt']) или ['ALL']")

    class Config:
        title = "Запрос на создание правила"

class RuleUpdateRequest(BaseModel):
    """Запрос на обновление существующего правила синхронизации."""
    enabled: Optional[bool] = Field(None, description="Включено ли")
    category: Optional[str] = Field(None, description="Категория")
    mode: Optional[str] = Field(None, description="Режим")
    source_sheet: Optional[str] = Field(None, description="Лист-источник")
    source_header: Optional[str] = Field(None, description="Заголовок-источник")
    target_sheet: Optional[str] = Field(None, description="Целевой лист")
    target_header: Optional[str] = Field(None, description="Целевой заголовок")
    sheet_a: Optional[str] = Field(None, description="Лист A")
    header_a: Optional[str] = Field(None, description="Заголовок A")
    sheet_b: Optional[str] = Field(None, description="Лист B")
    header_b: Optional[str] = Field(None, description="Заголовок B")
    is_external: Optional[bool] = Field(None, description="Внешняя ли")
    target_doc_id: Optional[str] = Field(None, description="ID внешней таблицы")
    projects: Optional[List[str]] = Field(None, description="Список проектов")

    class Config:
        title = "Запрос на обновление правила"

class RuleToggleRequest(BaseModel):
    """Запрос на переключение статуса активности правила."""
    enabled: bool = Field(..., description="Новый статус активности (True/False)")

    class Config:
        title = "Запрос на переключение правила"
