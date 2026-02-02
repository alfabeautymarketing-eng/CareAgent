from typing import Optional
from pydantic import BaseModel, Field

class FormulaPriceDynamicsRequest(BaseModel):
    """Запрос на расчет формул динамики цен."""
    spreadsheet_id: str = Field(..., description="ID таблицы")
    sheet_name: str = Field("Динамика цены", description="Имя листа")
    dry_run: bool = Field(False, description="Тестовый запуск")

    class Config:
        title = "Запрос динамики цен"

class FormulaPriceCalcRequest(BaseModel):
    """Запрос на обновление формул расчета цены."""
    spreadsheet_id: str = Field(..., description="ID таблицы")
    price_calc_sheet: str = Field("Расчет цены", description="Имя листа расчета")
    price_dynamics_sheet: str = Field("Динамика цены", description="Имя листа динамики")
    silent: bool = Field(False, description="Не выводить уведомления")
    dry_run: bool = Field(False, description="Тестовый запуск")

    class Config:
        title = "Запрос расчета цен"

class FormulaAddYearRequest(BaseModel):
    """Запрос на добавление колонок нового года."""
    spreadsheet_id: str = Field(..., description="ID таблицы")
    sheet_name: str = Field("Динамика цены", description="Имя листа")
    year: Optional[int] = Field(None, description="Год (например, 2026)")
    dry_run: bool = Field(False, description="Тестовый запуск")

    class Config:
        title = "Запрос добавления года"

class FormulaResponse(BaseModel):
    """Ответ операций с формулами."""
    status: str = Field(..., description="Статус")
    blocks_processed: int = Field(0, description="Обработано блоков")
    rows_updated: int = Field(0, description="Обновлено строк")
    columns_added: int = Field(0, description="Добавлено колонок")
    year: Optional[int] = Field(None, description="Год")
    message: str = Field("", description="Сообщение")

    class Config:
        title = "Ответ по формулам"
