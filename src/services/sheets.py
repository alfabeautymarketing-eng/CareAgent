import gspread
from gspread.utils import rowcol_to_a1
from src.utils.config import settings
from src.utils.logger import logger

class SheetsService:
    def __init__(self):
        self.gc = None
        self._connect()

    def _connect(self):
        """Connect to Google Sheets using service account."""
        try:
            self.gc = gspread.service_account(filename=settings.google_credentials_file)
            logger.info("connected_to_google_sheets")
        except Exception as e:
            logger.error("google_sheets_connection_failed", error=str(e))
            raise

    def get_worksheet(self, spreadsheet_id: str, sheet_name: str):
        """Get a worksheet object."""
        try:
            sh = self.gc.open_by_key(spreadsheet_id)
            return sh.worksheet(sheet_name)
        except Exception as e:
            logger.error("worksheet_not_found", spreadsheet_id=spreadsheet_id, sheet_name=sheet_name, error=str(e))
            raise

    def sort_by_header(self, spreadsheet_id: str, sheet_name: str, header_name: str, ascending: bool = True):
        """
        Sort a sheet by a specific header column.
        Assumes the first row is the header.
        """

        def _normalize(value: str) -> str:
            """Normalize header for comparison (trim, lower, collapse spaces and slashes)."""
            if value is None:
                return ""
            import re
            val = str(value).strip().lower()
            val = re.sub(r"\s+", " ", val)
            val = val.replace(" / ", "/")
            val = val.replace("/", " / ")  # ensure slash separated words are comparable
            val = re.sub(r"\s+", " ", val)
            return val

        def _tokens(value: str) -> set:
            """Tokenize header into meaningful parts (letters/digits, ignoring punctuation)."""
            import re
            norm = _normalize(value)
            return set(filter(None, re.split(r"[\\s,;:/\\\\]+", norm)))

        def _find_header_index(headers: list, target: str) -> int:
            """Find header index using tolerant matching. Returns -1 if not found."""
            norm_target = _normalize(target)
            tokens_target = _tokens(target)

            # 1) Exact match
            for i, h in enumerate(headers):
                if _normalize(h) == norm_target:
                    return i

            # 2) Contains "производ" fallback for manufacturer
            if "производ" in norm_target:
                for i, h in enumerate(headers):
                    if "производ" in _normalize(h):
                        return i

            # 3) Substring match
            for i, h in enumerate(headers):
                norm_h = _normalize(h)
                if norm_target and norm_target in norm_h:
                    return i

            # 4) Token subset match: all target tokens present in header tokens
            if tokens_target:
                for i, h in enumerate(headers):
                    header_tokens = _tokens(h)
                    if tokens_target.issubset(header_tokens):
                        return i

            # 5) Best token overlap (fallback for similar headers, e.g., extra words)
            best_idx = -1
            best_score = 0
            if tokens_target:
                for i, h in enumerate(headers):
                    header_tokens = _tokens(h)
                    if not header_tokens:
                        continue
                    overlap = len(tokens_target.intersection(header_tokens))
                    score = overlap / max(len(tokens_target), 1)
                    if score > best_score and overlap > 0:
                        best_score = score
                        best_idx = i
            if best_idx != -1 and best_score >= 0.5:
                return best_idx

            # 6) EXW heuristic: pick first column containing "exw" if nothing else matched
            if "exw" in norm_target:
                for i, h in enumerate(headers):
                    if "exw" in _normalize(h):
                        return i
            return -1

        try:
            ws = self.get_worksheet(spreadsheet_id, sheet_name)
            
            # Find the header column index
            headers = ws.row_values(1)
            col_index = _find_header_index(headers, header_name)
            if col_index == -1:
                raise ValueError(f"Header '{header_name}' not found in the first row. Headers: {headers}")
            col_index += 1  # convert to 1-based index

            # Calculate the range to sort (exclude header row)
            # A2 to last_col last_row
            num_rows = ws.row_count
            num_cols = ws.col_count
            last_cell = rowcol_to_a1(num_rows, num_cols)
            sort_range = f"A2:{last_cell}"

            # Sort
            # gspread sort takes specs: validation_result list of (col_index, 'asc'|'des')
            order = 'asc' if ascending else 'des'
            ws.sort((col_index, order), range=sort_range)
            
            logger.info("sheet_sorted", sheet=sheet_name, column=header_name, order=order)
            return True

        except Exception as e:
            logger.error("sort_failed", error=str(e))
            raise

    def process_load_logic(self, spreadsheet_id: str):
        """
        Run the heavy logic that was previously in GAS onOpen.
        1. Apply calculation formulas to "Расчет цены"
        """
        results = {}
        
        # 1. Apply Calculation Formulas
        try:
            self.apply_calculation_formulas(spreadsheet_id)
            results["price_calculation"] = "Success"
        except Exception as e:
            logger.error("apply_calculation_formulas_failed", error=str(e))
            results["price_calculation"] = f"Failed: {str(e)}"
            
        return results

    def apply_calculation_formulas(self, spreadsheet_id: str):
        """
        Replicates Lib.applyCalculationFormulas logic.
        Updates formulas in "Расчет цены" based on "Динамика цены" and headers.
        """
        SHEET_NAME = "Расчет цены"
        REF_SHEET_NAME = "Справочник"
        
        ws = self.get_worksheet(spreadsheet_id, SHEET_NAME)
        
        # Get all values to map headers
        values = ws.get_all_values()
        if not values:
            logger.warning("sheet_empty", sheet=SHEET_NAME)
            return
            
        headers = values[0]
        num_rows = len(values)
        if num_rows <= 1:
            return

        # Map headers to 0-based indices
        cols = {h.strip(): i for i, h in enumerate(headers) if h}
        
        # Helper to get col letter (A, B, ...) or R1C1 offset
        def get_col_idx(name): 
            return cols.get(name)

        # Logic for Reference data (Currency, Markup)
        # We need to build the formulas string dynamically.
        # In GAS: VLOOKUP(currentYear, INDIRECT(...), ...)
        # Here we will use the same formula logic because it runs IN sheets.
        
        # Current Year
        import datetime
        current_year = datetime.datetime.now().year
        
        # We need to find "Год" and "Курс валюты, €" in Reference sheet to build the VLOOKUP
        # For simplicity, we assume standard positions or fetching them if needed. 
        # But to be robust like GAS, we should fetch reference headers. 
        # For now, we will use the safer 'INDIRECT' approach if we want to be dynamic, 
        # or simplified hardcoded ranges if we trust structure.
        # Let's try to fetch reference headers to be safe.
        
        try:
            ref_ws = self.get_worksheet(spreadsheet_id, REF_SHEET_NAME)
            ref_headers = ref_ws.row_values(1)
            try:
                ref_year_idx = ref_headers.index("Год") + 1
                ref_curr_idx = ref_headers.index("Курс валюты, €") + 1
                ref_last_row = ref_ws.row_count
                
                # Formula: VLOOKUP(2025, 'Справочник'!R2C1:R100C5, 5, FALSE) but using R1C1 notation in formulas might be tricky if not consistent.
                # Let's use A1 notation for the reference part inside the formula.
                # Actually, GAS used INDIRECT + R1C1 string construction.
                # We can construct a standard VLOOKUP: VLOOKUP(2025, 'Справочник'!A:E, 5, FALSE)
                
                # Column letters
                from gspread.utils import rowcol_to_a1
                # We need column letters for the reference range
                # range: Col(ref_year_idx) : Col(ref_curr_idx)
                # It's safer to use the exact columns.
                
                # Simplified robust approach: 'Справочник'!A:Z is usually fine if we know indices.
                # But let's stick to the generated formula structure if possible.
                
                # Let's use a simpler static formula for now or replicates the dynamic builder if critical.
                # The GAS code uses: INDIRECT("Справочник!R2C" & yearCol & ":R" & lastRow & "C" & currCol, FALSE)
                # We can produce exactly that string.
                
                currency_ref = f'VLOOKUP({current_year}, INDIRECT("Справочник!R2C{ref_year_idx}:R{ref_last_row}C{ref_curr_idx}", FALSE), {ref_curr_idx - ref_year_idx + 1}, FALSE)'
            except ValueError:
                # Fallback
                currency_ref = "1"
                
            # Markup ref
            try:
                markup_idx = ref_headers.index("% наценка для РРЦ") + 1
                markup_ref = f"'Справочник'!R2C{markup_idx}"
            except ValueError:
                markup_ref = "0"
                
        except Exception:
            # If reference sheet missing
            currency_ref = "1"
            markup_ref = "0"

        # Apply formulas column by column
        # list of (col_name, formula_template_lambda)
        
        updates = [] # list of {range, values}
        
        # Helper R1C1 generator
        # R[0]C[x] relative to current cell.
        
        def make_r1c1(target_col_idx, source_col_idx):
            offset = source_col_idx - target_col_idx
            return f"R[0]C[{offset}]"
            
        # 1. Price Wholesale EUR = EXW * Coeff
        # "Расчетная цена Опт, €" = "EXW текущая, €" * "К-т"
        tgt = get_col_idx("Расчетная цена Опт, €")
        src_exw = get_col_idx("EXW текущая, €")
        src_coeff = get_col_idx("К-т")
        
        if tgt is not None and src_exw is not None and src_coeff is not None:
            # Formula: =IF(LEN(R[0]C[exw_off])=0, "", R[0]C[exw_off] * R[0]C[coeff_off])
            exw_ref = make_r1c1(tgt, src_exw)
            coeff_ref = make_r1c1(tgt, src_coeff)
            formula = f'=IF(LEN({exw_ref})=0, "", {exw_ref} * {coeff_ref})'
            
            # Prepare batch data: (row_idx, col_idx, value) - gspread needs range.
            # We want to update from row 2 to num_rows
            # range: R2C(tgt+1):R(num_rows)C(tgt+1)
            # We will use batch_update with data
            
            # Construct a list of [formula] for all rows
            formulas = [[formula]] * (num_rows - 1)
            updates.append({
                'range': f"{rowcol_to_a1(2, tgt+1)}:{rowcol_to_a1(num_rows, tgt+1)}",
                'values': formulas
            })

        # 2. Price Wholesale RUB = Price EUR * Currency
        tgt_rub = get_col_idx("Расчетная цена Опт, ₽")
        src_eur = get_col_idx("Расчетная цена Опт, €")
        
        if tgt_rub is not None and src_eur is not None:
            eur_ref = make_r1c1(tgt_rub, src_eur)
            formula = f'=IF(LEN({eur_ref})=0, "", ROUND({eur_ref} * {currency_ref}, 0))'
            formulas = [[formula]] * (num_rows - 1)
            updates.append({
                'range': f"{rowcol_to_a1(2, tgt_rub+1)}:{rowcol_to_a1(num_rows, tgt_rub+1)}",
                'values': formulas
            })
            
        # 3. Cost per ml = Price RUB / Volume
        tgt_cost = get_col_idx("Стоимость за 1 мл, ₽")
        src_rub = get_col_idx("Расчетная цена Опт, ₽")
        src_vol = get_col_idx("Объём")
        
        if tgt_cost is not None and src_rub is not None and src_vol is not None:
            rub_ref = make_r1c1(tgt_cost, src_rub)
            vol_off = src_vol - tgt_cost
            # Reusing the regex logic for volume from GAS
            # IF(LEN(R[0]C[off])=0, 1, IF(ISNUMBER(SEARCH("х", ...
            # We construct the volume formula string
            vol_cell = f"R[0]C[{vol_off}]"
            vol_formula = (
                f'IF(LEN({vol_cell})=0, 1, '
                f'IF(ISNUMBER(SEARCH("х", {vol_cell})), '
                f'VALUE(REGEXEXTRACT({vol_cell}, "\\d+")) * VALUE(REGEXEXTRACT(SUBSTITUTE({vol_cell}, " ", ""), "х([\\d,]+)")), '
                f'VALUE(REGEXEXTRACT({vol_cell}, "[\\d,]+"))))'
            )
            
            formula = f'=IF(LEN({rub_ref})=0, "", ROUND({rub_ref} / ({vol_formula}), 0))'
            formulas = [[formula]] * (num_rows - 1)
            updates.append({
                'range': f"{rowcol_to_a1(2, tgt_cost+1)}:{rowcol_to_a1(num_rows, tgt_cost+1)}",
                'values': formulas
            })

        # 4. Markup % = Reference
        tgt_mark = get_col_idx("% наценка для РРЦ")
        # Needs EXW check
        src_exw_chk = get_col_idx("EXW текущая, €")
        
        if tgt_mark is not None and src_exw_chk is not None:
            exw_ref = make_r1c1(tgt_mark, src_exw_chk)
            formula = f'=IF(LEN({exw_ref})=0, "", {markup_ref} / 100)'
            formulas = [[formula]] * (num_rows - 1)
            updates.append({
                'range': f"{rowcol_to_a1(2, tgt_mark+1)}:{rowcol_to_a1(num_rows, tgt_mark+1)}",
                'values': formulas
            })
            
        # Execute Batch Update
        if updates:
            ws.batch_update(updates, value_input_option='USER_ENTERED')
            logger.info("formulas_applied", sheet=SHEET_NAME, columns_updated=len(updates))

    def reorder_sheets(self, spreadsheet_id: str, desired_order: list):
        """
        Reorder sheets according to desired order.
        Sheets not in the list are moved to the end.

        Args:
            spreadsheet_id: Google Sheets document ID
            desired_order: List of sheet names in desired order

        Returns:
            dict with reordering results
        """
        try:
            sh = self.gc.open_by_key(spreadsheet_id)
            worksheets = sh.worksheets()

            # Build current sheet name -> worksheet mapping
            sheet_map = {ws.title: ws for ws in worksheets}
            current_names = [ws.title for ws in worksheets]

            # Build target order: known sheets first, then unknown
            target_order = []

            # Add sheets from desired_order that exist
            for name in desired_order:
                if name in sheet_map:
                    target_order.append(name)

            # Add remaining sheets (not in desired_order) at the end
            for name in current_names:
                if name not in target_order:
                    target_order.append(name)

            # Check if reordering is needed
            if target_order == current_names:
                logger.info("sheets_already_ordered", spreadsheet_id=spreadsheet_id)
                return {
                    "sheets_moved": 0,
                    "final_order": target_order,
                    "message": "Sheets already in correct order"
                }

            # Reorder sheets using batch update
            # gspread uses sheet IDs for reordering
            requests = []
            for new_index, sheet_name in enumerate(target_order):
                ws = sheet_map[sheet_name]
                requests.append({
                    "updateSheetProperties": {
                        "properties": {
                            "sheetId": ws.id,
                            "index": new_index
                        },
                        "fields": "index"
                    }
                })

            if requests:
                sh.batch_update({"requests": requests})

            sheets_moved = sum(1 for i, name in enumerate(target_order) if i < len(current_names) and current_names[i] != name)

            logger.info("sheets_reordered",
                       spreadsheet_id=spreadsheet_id,
                       sheets_moved=sheets_moved)

            return {
                "sheets_moved": sheets_moved,
                "final_order": target_order,
                "message": f"Reordered {sheets_moved} sheets"
            }

        except Exception as e:
            logger.error("reorder_sheets_failed", error=str(e))
            raise
