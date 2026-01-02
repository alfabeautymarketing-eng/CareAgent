"""
Product Matcher Service.

Handles fetching base products from Google Sheets and matching new products
using Gemini AI.
"""

import json
from typing import List, Dict, Optional, Any

from src.services.gemini_client import get_gemini_client
from src.services.sheets import SheetsService
from src.utils.logger import logger


class ProductMatcher:
    """Service for finding matching base products."""

    def __init__(self, spreadsheet_id: str, sheet_name: str = "Сертификация"):
        self.spreadsheet_id = spreadsheet_id
        self.sheet_name = sheet_name
        self.sheets_service = SheetsService()

    def fetch_base_products(self) -> List[Dict[str, str]]:
        """
        Fetches rows from the sheet where 'Базовый' is checked.
        Returns a list of simplified product dictionaries.
        """
        logger.info(f"Fetching base products from '{self.sheet_name}'...")
        try:
            ws = self.sheets_service.get_worksheet(self.spreadsheet_id, self.sheet_name)
            all_values = ws.get_all_values()
        except Exception as e:
            logger.error(f"Error fetching sheet data: {e}")
            return []

        if not all_values:
            logger.warning("Sheet is empty.")
            return []

        headers = all_values[0]
        data_rows = all_values[1:]

        try:
            idx_base = headers.index("Базовый")
            idx_name_en = headers.index("Наименования англ по ДС")
            idx_name_ru = headers.index("Наименования рус по ДС")
        except ValueError as e:
            logger.error(f"Missing required headers: {e}")
            return []

        base_products = []

        for i, row in enumerate(data_rows):
            # Pad row if short
            if len(row) <= max(idx_base, idx_name_en, idx_name_ru):
                row = row + [""] * (len(headers) - len(row))

            is_base = str(row[idx_base]).lower() in ('true', 'yes', 'да', '1')

            if is_base:
                name_en = row[idx_name_en].strip()
                name_ru = row[idx_name_ru].strip()

                if not name_en and not name_ru:
                    continue

                base_products.append({
                    "id": f"ROW-{i + 2}",
                    "name_eng_ds": name_en,
                    "name_rus_ds": name_ru,
                    # "keywords": f"{name_en} {name_ru}" # Not strictly needed for JSON passed to AI
                })

        logger.info(f"Loaded {len(base_products)} Base Products.")
        return base_products

    def find_best_match(self, new_product_name: str, candidates: Optional[List[Dict]] = None) -> Optional[Dict[str, Any]]:
        """
        Uses Gemini to find the best semantic match.
        If candidates are not provided, fetches them first.
        """
        if candidates is None:
            candidates = self.fetch_base_products()

        if not candidates:
            return None

        client = get_gemini_client()

        # Simplify candidates for prompt to save tokens/reduce noise
        candidates_simplified = [
            {"id": c["id"], "en": c["name_eng_ds"], "ru": c["name_rus_ds"]}
            for c in candidates
        ]

        candidates_str = json.dumps(candidates_simplified, ensure_ascii=False, indent=2)

        prompt = f"""
START_ROLE
You are an expert Data Matching Assistant for a cosmetic company.
END_ROLE

TASK
Match a raw product name from a supplier price list (New Product) to our internal Base Product Database.
Find the best semantic match.

NEW PRODUCT NAME:
"{new_product_name}"

BASE PRODUCTS DATABASE:
{candidates_str}

INSTRUCTIONS
1. Analyze the New Product Name.
2. Look for a Base Product that represents the SAME core product (ignoring volume, 'tester', 'sample' suffixes, or slight naming variations).
3. If a match is found, return the ID of the Base Product and a confidence score (0-100).
4. If no match is found (confidence < 50), return null.
5. Return ONLY valid JSON.

JSON FORMAT:
{{
  "match_found": true/false,
  "best_match_id": "ROW-xxx" or null,
  "confidence": 95,
  "reasoning": "Explanation why..."
}}
"""
        try:
            response_text = client.call_api(
                prompt=prompt,
                temperature=0.1,
                json_response=True,
                max_output_tokens=2000
            )
            
            # Clean content if markdown blocks exist
            cleaned_text = response_text.strip()
            if cleaned_text.startswith("```json"):
                cleaned_text = cleaned_text[7:]
            if cleaned_text.startswith("```"):
                cleaned_text = cleaned_text[3:]
            if cleaned_text.endswith("```"):
                cleaned_text = cleaned_text[:-3]
            
            result = json.loads(cleaned_text.strip())
            
            # Enrich the result with the actual matched product data if found
            if result.get("match_found") and result.get("best_match_id"):
                 best_id = result["best_match_id"]
                 matched_product = next((p for p in candidates if p["id"] == best_id), None)
                 if matched_product:
                     result["matched_product_details"] = matched_product

            return result

        except Exception as e:
            logger.error(f"AI Matching Error: {e}")
            return None
