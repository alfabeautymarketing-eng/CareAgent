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
        try:
            ws = self.get_worksheet(spreadsheet_id, sheet_name)
            
            # Find the header column index
            headers = ws.row_values(1)
            try:
                col_index = headers.index(header_name) + 1  # 1-based index
            except ValueError:
                raise ValueError(f"Header '{header_name}' not found in the first row.")

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
