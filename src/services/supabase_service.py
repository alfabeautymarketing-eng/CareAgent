from typing import Any, Union

from supabase import Client, create_client

from src.utils.config import settings
from src.utils.logger import logger


class SupabaseService:
    def __init__(self):
        self.client: Client | None = None
        if settings.supabase_url and settings.supabase_service_key:
            try:
                self.client = create_client(settings.supabase_url, settings.supabase_service_key)
            except Exception as e:
                logger.error("supabase_init_failed", error=str(e))

    def is_configured(self) -> bool:
        return self.client is not None

    def get_tables(self) -> list[str]:
        """
        Returns a list of public tables.
        Since PostgREST doesn't expose a 'list tables' endpoint directly without
        permission adjustments, we return the known schema or a static list if
        dynamic listing fails.
        """
        if not self.client:
            return ["users", "products", "orders", "sync_log"]  # Fallback

        # Try to infer from a known metadata query or just return the standard set
        # Real dynamic listing would require querying information_schema via
        # a specific function or SQL
        return ["users", "products", "orders", "sync_log"]

    def execute_query(self, query: str) -> dict[str, Any]:
        """
        Executes a simplified SQL query.
        WARNING: This implementation primarily supports SELECT queries maps them to PostgREST calls.
        It is NOT a full SQL executor (which requires direct Postgres connection).
        """
        if not self.client:
            raise Exception("Supabase client not configured")

        query = query.strip()
        lower_query = query.lower()

        try:
            # Very basic parser for "SELECT * FROM table LIMIT n"
            if lower_query.startswith("select"):
                # Extract table name
                parts = query.split()
                try:
                    from_index = [p.lower() for p in parts].index("from")
                    table_name = parts[from_index + 1]
                    # Remove any punctuation from table name
                    table_name = table_name.strip(';"')
                except ValueError:
                    raise Exception("Invalid SELECT query: Could not find FROM clause")

                # Build query
                db_query = self.client.table(table_name).select("*")

                # Check for LIMIT
                if "limit" in lower_query:
                    try:
                        limit_index = [p.lower() for p in parts].index("limit")
                        limit_val = int(parts[limit_index + 1].strip(';'))
                        db_query = db_query.limit(limit_val)
                    except (ValueError, IndexError):
                        pass  # Ignore limit parsing errors

                response = db_query.execute()
                return {"data": response.data, "count": len(response.data)}

            else:
                raise Exception("Only SELECT queries are currently supported via this interface")

        except Exception as e:
            logger.error("supabase_query_failed", query=query, error=str(e))
            raise e

    def get_table_data(self, table_name: str, limit: int = 100) -> list[dict[str, Any]]:
        if not self.client:
             raise Exception("Supabase client not configured")

        response = self.client.table(table_name).select("*").limit(limit).execute()
        return response.data

    def upsert_data(self, table_name: str, data: Union[dict[str, Any], list[dict[str, Any]]], on_conflict: str = 'id') -> dict[str, Any]:
        """
        Upsert data into a table.
        Args:
            table_name: Name of the table
            data: Dictionary or list of dictionaries to upsert
            on_conflict: Column to check for conflicts (default: 'id')
        """
        if not self.client:
             raise Exception("Supabase client not configured")
        
        try:
            response = self.client.table(table_name).upsert(data, on_conflict=on_conflict).execute()
            return {"data": response.data, "count": len(response.data) if response.data else 0}
        except Exception as e:
            logger.error("supabase_upsert_failed", table=table_name, error=str(e))
            raise e

_supabase_service = None

def get_supabase_service() -> SupabaseService:
    global _supabase_service
    if not _supabase_service:
        _supabase_service = SupabaseService()
    return _supabase_service
