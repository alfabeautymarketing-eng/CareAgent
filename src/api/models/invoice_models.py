from pydantic import BaseModel, Field

class InvoiceFormatRequest(BaseModel):
    """Запрос на форматирование листа заказа."""
    spreadsheet_id: str = Field(..., description="ID таблицы")
    sheet_name: str = Field("Ордер", description="Имя листа для форматирования")
    dry_run: bool = Field(False, description="Тестовый запуск")

    class Config:
        title = "Запрос форматирования инвойса"

class InvoiceCreateRequest(BaseModel):
    """Запрос на создание полного инвойса."""
    spreadsheet_id: str = Field(..., description="ID таблицы")
    order_sheet: str = Field("Ордер", description="Имя листа заказа")
    certification_sheet: str = Field("Сертификация", description="Имя листа сертификации")
    labels_sheet: str = Field("Этикетки", description="Имя листа этикеток")
    target_sheet: str = Field("Для инвойса", description="Имя целевого листа")
    dry_run: bool = Field(False, description="Тестовый запуск")

    class Config:
        title = "Запрос создания инвойса"

class InvoiceResponse(BaseModel):
    """Ответ операций по инвойсу."""
    status: str = Field(..., description="Статус")
    rows_processed: int = Field(0, description="Обработано строк")
    target_sheet: str = Field("", description="Название целевого листа")
    message: str = Field("", description="Сообщение")

    class Config:
        title = "Ответ по инвойсу"
