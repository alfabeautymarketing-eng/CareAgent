"""
AI Service for AgentCare.
Orchestrates INCI analysis using Gemini and PDF processing.
"""

import time
from typing import Optional, Dict, Any, List
from dataclasses import dataclass

from src.services.gemini_client import GeminiClient, get_gemini_client
from src.services.pdf_processor import PDFProcessor, get_pdf_processor
from src.services.sheets import SheetsService
from src.utils.logger import logger


@dataclass
class AnalysisResult:
    """Result of INCI analysis."""
    success: bool
    row_number: int
    data: Optional[Dict[str, Any]] = None
    error: Optional[str] = None
    duration: float = 0.0


class AIService:
    """
    Service for AI-powered INCI analysis.
    Integrates Gemini API with Google Sheets.
    """

    # Column mapping for writing results (L-Y, 0-indexed from column L)
    RESULT_COLUMNS = [
        "category_code",      # L (12) - код категории
        "category",           # M (13) - категория
        "product_type",       # N (14) - вид продукции
        "product_reason",     # O (15) - обоснование
        "reg_doc",            # P (16) - документ
        "reg_reason",         # Q (17) - обоснование документа
        "active_ru_no_pct",   # R (18) - активные без % (рус)
        "active_en_no_pct",   # S (19) - активные без % (англ)
        "active_ru_pct",      # T (20) - активные с % (рус)
        "active_en_pct",      # U (21) - активные с % (англ)
        "full_ru_no_pct",     # V (22) - полный без % (рус)
        "full_en_no_pct",     # W (23) - полный без % (англ)
        "full_ru_pct",        # X (24) - полный с % (рус)
        "full_en_pct"         # Y (25) - полный с % (англ)
    ]

    # Input columns (0-indexed)
    COL_ID = 0           # A - ID
    COL_NAME = 4         # E - Наименование ДС
    COL_INCI_LINK = 7    # H - INCI ссылка
    COL_PURPOSE = 8      # I - Назначение
    COL_APPLICATION = 9  # J - Применение
    COL_CATEGORY = 11    # L - Категория (first result column)
    COL_ACTIVE_RU = 17   # R - Активные рус (to check if already processed)

    RESULT_START_COL = 12  # Column L (1-indexed)

    def __init__(
        self,
        sheets_service: Optional[SheetsService] = None,
        gemini_client: Optional[GeminiClient] = None,
        pdf_processor: Optional[PDFProcessor] = None
    ):
        self.sheets = sheets_service or SheetsService()
        self.gemini = gemini_client or get_gemini_client()
        self.pdf = pdf_processor or get_pdf_processor()

    def check_service(self) -> Dict[str, Any]:
        """
        Check Gemini service status.

        Returns:
            Status dict with connection info and available models
        """
        return self.gemini.test_connection()

    def get_available_categories(self) -> List[Dict[str, str]]:
        """
        Get available ТН ВЭД categories.

        Returns:
            List of category dicts
        """
        return [
            {
                "code": "3304",
                "name": "Средства косметические или для макияжа и средства для ухода за кожей",
                "description": "Кремы, лосьоны, маски для лица и тела, декоративная косметика"
            },
            {
                "code": "3305",
                "name": "Средства для ухода за волосами",
                "description": "Шампуни, кондиционеры, маски для волос, средства для укладки"
            },
            {
                "code": "3307",
                "name": "Средства для бритья, дезодоранты, соли для ванн и прочие",
                "description": "Дезодоранты, антиперспиранты, средства для бритья, соли и бомбочки для ванн"
            }
        ]

    def analyze_pdf(
        self,
        pdf_url: str,
        purpose: Optional[str] = None,
        application: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Analyze PDF document with INCI composition.

        Args:
            pdf_url: Google Drive URL or file ID
            purpose: Product purpose
            application: Product application

        Returns:
            Analysis result dict
        """
        start_time = time.time()

        logger.info("analyze_pdf_start", url=pdf_url[:50] if pdf_url else None)

        # 1. Extract text from PDF
        try:
            inci_text = self.pdf.extract_text(pdf_url)
        except Exception as e:
            # Try OCR fallback
            logger.warning("pdf_extraction_failed_trying_ocr", error=str(e))
            try:
                inci_text = self.pdf.extract_text_with_ocr(pdf_url)
            except Exception as ocr_error:
                logger.error("pdf_ocr_also_failed", error=str(ocr_error))
                raise ValueError(f"Could not extract text from PDF: {e}")

        # 2. Analyze with Gemini
        result = self.gemini.analyze_inci(inci_text, purpose, application)

        duration = time.time() - start_time
        logger.info("analyze_pdf_complete",
                   duration=f"{duration:.2f}s",
                   category=result.get("category_code"))

        return result

    def analyze_row(
        self,
        spreadsheet_id: str,
        sheet_name: str,
        row_number: int
    ) -> AnalysisResult:
        """
        Analyze a single row in the spreadsheet.

        Args:
            spreadsheet_id: Google Sheets document ID
            sheet_name: Sheet name (e.g., "Информация")
            row_number: Row number (1-indexed)

        Returns:
            AnalysisResult with success status and data
        """
        start_time = time.time()

        logger.info("analyze_row_start",
                   spreadsheet_id=spreadsheet_id[:20],
                   sheet=sheet_name,
                   row=row_number)

        try:
            # Get worksheet
            ws = self.sheets.get_worksheet(spreadsheet_id, sheet_name)

            # Read row data (columns A-Y, 25 columns)
            row_range = ws.range(row_number, 1, row_number, 25)
            values = [cell.value for cell in row_range]

            # Extract input data
            row_data = {
                "id": values[self.COL_ID] if len(values) > self.COL_ID else None,
                "product_name": values[self.COL_NAME] if len(values) > self.COL_NAME else None,
                "inci_link": values[self.COL_INCI_LINK] if len(values) > self.COL_INCI_LINK else None,
                "purpose": values[self.COL_PURPOSE] if len(values) > self.COL_PURPOSE else None,
                "application": values[self.COL_APPLICATION] if len(values) > self.COL_APPLICATION else None
            }

            # If inci_link is just text but might be a hyperlink, try to get the formula
            if row_data["inci_link"] and not row_data["inci_link"].startswith("http") and not "/" in row_data["inci_link"]:
                try:
                    # Fetch only the INCI cell with formula to extract URL
                    cell_a1 = rowcol_to_a1(row_number, self.COL_INCI_LINK + 1)
                    cell_with_formula = ws.acell(cell_a1, value_render_option='FORMULA').value
                    if "HYPERLINK" in cell_with_formula:
                        import re
                        match = re.search(r'HYPERLINK\("([^"]+)"', cell_with_formula)
                        if match:
                            row_data["inci_link"] = match.group(1)
                            logger.info("extracted_link_from_formula", url=row_data["inci_link"][:50])
                except Exception as fe:
                    logger.warning("failed_to_extract_formula_link", error=str(fe))

            if not row_data["inci_link"]:
                return AnalysisResult(
                    success=False,
                    row_number=row_number,
                    error="Missing INCI link (column H)"
                )

            logger.info("analyze_row_data",
                       product=row_data.get("product_name"),
                       inci_link=row_data.get("inci_link", "")[:50])

            # Analyze PDF
            result = self.analyze_pdf(
                pdf_url=row_data["inci_link"],
                purpose=row_data.get("purpose"),
                application=row_data.get("application")
            )

            # Write results to sheet
            self._write_results(ws, row_number, result)

            duration = time.time() - start_time

            return AnalysisResult(
                success=True,
                row_number=row_number,
                data=result,
                duration=duration
            )

        except Exception as e:
            duration = time.time() - start_time
            logger.error("analyze_row_failed",
                        row=row_number,
                        error=str(e))

            return AnalysisResult(
                success=False,
                row_number=row_number,
                error=str(e),
                duration=duration
            )

    def _write_results(
        self,
        worksheet,
        row_number: int,
        result: Dict[str, Any]
    ):
        """Write analysis results to the worksheet."""
        # Prepare values for columns L-Y (14 columns)
        values = []
        for col_name in self.RESULT_COLUMNS:
            values.append(result.get(col_name, ""))

        # Update cells L through Y
        cell_list = worksheet.range(row_number, self.RESULT_START_COL,
                                    row_number, self.RESULT_START_COL + len(values) - 1)

        for i, cell in enumerate(cell_list):
            cell.value = values[i]

        worksheet.update_cells(cell_list)

        logger.info("results_written", row=row_number, columns=len(values))

    def analyze_empty_rows(
        self,
        spreadsheet_id: str,
        sheet_name: str,
        delay_between: float = 2.0
    ) -> Dict[str, Any]:
        """
        Find and analyze all empty rows (rows with INCI link but no results).

        Args:
            spreadsheet_id: Google Sheets document ID
            sheet_name: Sheet name
            delay_between: Delay between requests in seconds

        Returns:
            Statistics dict with total, success, failed counts
        """
        logger.info("analyze_empty_rows_start",
                   spreadsheet_id=spreadsheet_id[:20],
                   sheet=sheet_name)

        try:
            ws = self.sheets.get_worksheet(spreadsheet_id, sheet_name)

            # Get all data
            all_values = ws.get_all_values()

            if len(all_values) <= 1:
                return {
                    "total": 0,
                    "success": 0,
                    "failed": 0,
                    "message": "Sheet is empty or has only header row"
                }

            # Find rows to process (skip header row)
            empty_rows = []
            for i, row in enumerate(all_values[1:], start=2):  # Start from row 2
                # Check if row has INCI link but no category
                has_inci = len(row) > self.COL_INCI_LINK and row[self.COL_INCI_LINK]
                has_category = len(row) > self.COL_CATEGORY and row[self.COL_CATEGORY]
                has_active_ru = len(row) > self.COL_ACTIVE_RU and row[self.COL_ACTIVE_RU]

                if has_inci and (not has_category or not has_active_ru):
                    empty_rows.append(i)

            logger.info("empty_rows_found", count=len(empty_rows))

            if not empty_rows:
                return {
                    "total": 0,
                    "success": 0,
                    "failed": 0,
                    "message": "All rows already processed"
                }

            # Process rows
            stats = {
                "total": len(empty_rows),
                "success": 0,
                "failed": 0,
                "errors": []
            }

            for idx, row_num in enumerate(empty_rows):
                logger.info("processing_row",
                           current=idx + 1,
                           total=len(empty_rows),
                           row=row_num)

                result = self.analyze_row(spreadsheet_id, sheet_name, row_num)

                if result.success:
                    stats["success"] += 1
                else:
                    stats["failed"] += 1
                    stats["errors"].append({
                        "row": row_num,
                        "error": result.error
                    })

                # Delay between requests (except for last one)
                if idx < len(empty_rows) - 1:
                    time.sleep(delay_between)

            logger.info("analyze_empty_rows_complete",
                       success=stats["success"],
                       failed=stats["failed"])

            return stats

        except Exception as e:
            logger.error("analyze_empty_rows_failed", error=str(e))
            raise

    def get_gemini_settings(self) -> Dict[str, Any]:
        """
        Get current Gemini settings.

        Returns:
            Dict with model and API key status
        """
        from src.utils.config import settings

        api_key = settings.gemini_api_key
        key_status = "configured" if api_key else "not_set"

        # Mask key for display
        if api_key and len(api_key) > 8:
            masked_key = api_key[:4] + "****" + api_key[-4:]
        else:
            masked_key = "****" if api_key else "not set"

        return {
            "model": settings.gemini_model,
            "api_key_status": key_status,
            "api_key_masked": masked_key,
            "max_retries": settings.gemini_max_retries
        }


# Singleton instance
_ai_service: Optional[AIService] = None


def get_ai_service() -> AIService:
    """Get or create the AI service singleton."""
    global _ai_service
    if _ai_service is None:
        _ai_service = AIService()
    return _ai_service
