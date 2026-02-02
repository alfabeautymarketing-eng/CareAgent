from pydantic import BaseModel, Field

class ExternalDocAddRequest(BaseModel):
    """Запрос на добавление внешнего документа."""
    name: str = Field(..., description="Название документа")
    doc_id: str = Field(..., description="ID документа (Google Doc ID)")

    class Config:
        title = "Запрос добавления документа"
