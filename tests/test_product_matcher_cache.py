"""
Tests for ProductMatcher caching functionality.
"""

import time
import unittest
from unittest.mock import Mock, patch, MagicMock

from src.services.product_matcher import ProductMatcher


class TestProductMatcherCache(unittest.TestCase):
    """Tests for ProductMatcher cache behavior."""

    def setUp(self):
        """Set up test fixtures."""
        self.spreadsheet_id = "test-sheet-id"
        self.matcher = ProductMatcher(self.spreadsheet_id)

    def tearDown(self):
        """Clean up after tests."""
        self.matcher.invalidate_cache()

    @patch('src.services.product_matcher.SheetsService')
    def test_cache_hit_avoids_api_call(self, mock_sheets_service_class):
        """Verify that cache hit doesn't make an API call."""
        # Setup mock
        mock_sheets_service = MagicMock()
        mock_sheets_service_class.return_value = mock_sheets_service

        mock_ws = MagicMock()
        mock_sheets_service.get_worksheet.return_value = mock_ws

        test_data = [
            ["Базовый", "Наименования англ по ДС", "Наименования рус по ДС"],
            ["Да", "English Name 1", "Russian Name 1"],
            ["Да", "English Name 2", "Russian Name 2"],
        ]
        mock_ws.get_all_values.return_value = test_data

        # Create fresh matcher with mocked service
        self.matcher.sheets_service = mock_sheets_service

        # First call - should fetch from API
        result1 = self.matcher.fetch_base_products()
        self.assertEqual(mock_sheets_service.get_worksheet.call_count, 1)
        self.assertEqual(len(result1), 2)

        # Second call - should use cache
        result2 = self.matcher.fetch_base_products()
        self.assertEqual(mock_sheets_service.get_worksheet.call_count, 1)  # Still 1, not 2
        self.assertEqual(result1, result2)

    @patch('src.services.product_matcher.SheetsService')
    def test_cache_ttl_expiration(self, mock_sheets_service_class):
        """Verify that cache expires after TTL."""
        # Setup mock
        mock_sheets_service = MagicMock()
        mock_sheets_service_class.return_value = mock_sheets_service

        mock_ws = MagicMock()
        mock_sheets_service.get_worksheet.return_value = mock_ws

        test_data = [
            ["Базовый", "Наименования англ по ДС", "Наименования рус по ДС"],
            ["Да", "English Name", "Russian Name"],
        ]
        mock_ws.get_all_values.return_value = test_data

        self.matcher.sheets_service = mock_sheets_service
        self.matcher._cache_ttl = 1  # 1 second for testing

        # First call
        result1 = self.matcher.fetch_base_products()
        self.assertEqual(mock_sheets_service.get_worksheet.call_count, 1)

        # Wait for TTL to expire
        time.sleep(1.1)

        # Second call - should fetch again
        result2 = self.matcher.fetch_base_products()
        self.assertEqual(mock_sheets_service.get_worksheet.call_count, 2)  # Cache expired

    @patch('src.services.product_matcher.SheetsService')
    def test_force_refresh_bypasses_cache(self, mock_sheets_service_class):
        """Verify that force_refresh=True bypasses cache."""
        # Setup mock
        mock_sheets_service = MagicMock()
        mock_sheets_service_class.return_value = mock_sheets_service

        mock_ws = MagicMock()
        mock_sheets_service.get_worksheet.return_value = mock_ws

        test_data = [
            ["Базовый", "Наименования англ по ДС", "Наименования рус по ДС"],
            ["Да", "English Name", "Russian Name"],
        ]
        mock_ws.get_all_values.return_value = test_data

        self.matcher.sheets_service = mock_sheets_service

        # First call
        result1 = self.matcher.fetch_base_products()
        self.assertEqual(mock_sheets_service.get_worksheet.call_count, 1)

        # Second call with force_refresh=True
        result2 = self.matcher.fetch_base_products(force_refresh=True)
        self.assertEqual(mock_sheets_service.get_worksheet.call_count, 2)  # Fetched again

    def test_invalidate_cache_clears_cache(self):
        """Verify that invalidate_cache() clears the cache."""
        # Manually set cache
        self.matcher._base_products_cache = [{"id": "test"}]
        self.matcher._cache_timestamp = time.time()

        # Verify cache is set
        self.assertIsNotNone(self.matcher._base_products_cache)
        self.assertGreater(self.matcher._cache_timestamp, 0)

        # Invalidate
        self.matcher.invalidate_cache()

        # Verify cache is cleared
        self.assertIsNone(self.matcher._base_products_cache)
        self.assertEqual(self.matcher._cache_timestamp, 0)

    @patch('src.services.product_matcher.SheetsService')
    def test_cache_after_invalidation(self, mock_sheets_service_class):
        """Verify that cache works again after invalidation."""
        # Setup mock
        mock_sheets_service = MagicMock()
        mock_sheets_service_class.return_value = mock_sheets_service

        mock_ws = MagicMock()
        mock_sheets_service.get_worksheet.return_value = mock_ws

        test_data = [
            ["Базовый", "Наименования англ по ДС", "Наименования рус по ДС"],
            ["Да", "English Name", "Russian Name"],
        ]
        mock_ws.get_all_values.return_value = test_data

        self.matcher.sheets_service = mock_sheets_service

        # First call
        self.matcher.fetch_base_products()
        self.assertEqual(mock_sheets_service.get_worksheet.call_count, 1)

        # Invalidate cache
        self.matcher.invalidate_cache()

        # Verify cache is cleared
        self.assertIsNone(self.matcher._base_products_cache)

        # Second call - should fetch again
        self.matcher.fetch_base_products()
        self.assertEqual(mock_sheets_service.get_worksheet.call_count, 2)

    @patch('src.services.product_matcher.SheetsService')
    def test_cache_with_multiple_matchers(self, mock_sheets_service_class):
        """Verify that cache is instance-level, not class-level."""
        # Setup mock
        mock_sheets_service = MagicMock()
        mock_sheets_service_class.return_value = mock_sheets_service

        mock_ws = MagicMock()
        mock_sheets_service.get_worksheet.return_value = mock_ws

        test_data = [
            ["Базовый", "Наименования англ по ДС", "Наименования рус по ДС"],
            ["Да", "English Name", "Russian Name"],
        ]
        mock_ws.get_all_values.return_value = test_data

        # Create two matchers
        matcher1 = ProductMatcher("sheet-1")
        matcher2 = ProductMatcher("sheet-2")

        matcher1.sheets_service = mock_sheets_service
        matcher2.sheets_service = mock_sheets_service

        # Reset call count
        mock_sheets_service.get_worksheet.reset_mock()

        # Fetch from matcher1
        matcher1.fetch_base_products()
        self.assertEqual(mock_sheets_service.get_worksheet.call_count, 1)

        # Fetch from matcher2 - should make another call (separate cache)
        matcher2.fetch_base_products()
        self.assertEqual(mock_sheets_service.get_worksheet.call_count, 2)


if __name__ == "__main__":
    unittest.main()
