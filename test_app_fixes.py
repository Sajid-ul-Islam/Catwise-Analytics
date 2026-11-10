import unittest
import pandas as pd
import numpy as np
import os
import json
import re
from io import BytesIO

import app

class TestCatwiseBugFixes(unittest.TestCase):
    def setUp(self):
        self.sample_df = pd.DataFrame({
            'Product Title': ['Panjabi Modern Cut', 'Classic T-Shirt', 'Denim Shirt Slim', 'Unknown Item'],
            'Selling Price': [1500, 450, 1200, 300],
            'Qty Sold': [2, 5, 1, 3],
            'COGS Amount': [900, 250, 700, 180],
            'Order Date': ['2026-05-01', '2026-05-02', '2026-05-03', '2026-05-04'],
            'Order #': ['ORD-101', 'ORD-102', 'ORD-103', 'ORD-104'],
            'Phone': ['01711111111', '01822222222', '01933333333', '01711111111']
        })
        self.valid_mapping = {
            'name': 'Product Title',
            'cost': 'Selling Price',
            'qty': 'Qty Sold',
            'cogs': 'COGS Amount',
            'date': 'Order Date',
            'order_id': 'Order #',
            'phone': 'Phone'
        }

    def test_bug1_no_duplicate_preview_expander(self):
        """Bug #1: Ensure 'Preview Data' appears at most ONCE in app.py."""
        with open("app.py", "r", encoding="utf-8") as f:
            content = f.read()
        matches = re.findall(r'expander\([^)]*Preview Data[^)]*\)', content, re.IGNORECASE)
        self.assertEqual(len(matches), 1, f"Expected exactly 1 'Preview Data' expander in app.py, found {len(matches)}.")

    def test_bug4_input_validation_catches_missing_required_columns(self):
        """Bug #4: Ensure validation detects missing or None required columns before processing."""
        # Test missing 'name'
        invalid_mapping_1 = {'name': None, 'cost': 'Selling Price', 'qty': 'Qty Sold'}
        is_valid, errors, _ = app.validate_column_mapping(self.sample_df, invalid_mapping_1)
        self.assertFalse(is_valid)
        self.assertTrue(any('Product Name' in e for e in errors))

        # Test 'None' string
        invalid_mapping_2 = {'name': 'Product Title', 'cost': 'None', 'qty': 'Qty Sold'}
        is_valid, errors, _ = app.validate_column_mapping(self.sample_df, invalid_mapping_2)
        self.assertFalse(is_valid)
        self.assertTrue(any('Unit Price' in e for e in errors))

        # Test empty string
        invalid_mapping_empty = {'name': 'Product Title', 'cost': 'Selling Price', 'qty': '   '}
        is_valid, errors, _ = app.validate_column_mapping(self.sample_df, invalid_mapping_empty)
        self.assertFalse(is_valid)
        self.assertTrue(any('Quantity' in e for e in errors))

        # Test column not in dataframe
        invalid_mapping_3 = {'name': 'Product Title', 'cost': 'Selling Price', 'qty': 'NonExistentColumn'}
        is_valid, errors, _ = app.validate_column_mapping(self.sample_df, invalid_mapping_3)
        self.assertFalse(is_valid)
        self.assertTrue(any('NonExistentColumn' in e for e in errors))

    def test_bug4_input_validation_catches_duplicate_required_mappings(self):
        """Bug #4: Warn when the same column is mapped to multiple required fields."""
        dup_mapping = {'name': 'Selling Price', 'cost': 'Selling Price', 'qty': 'Qty Sold'}
        is_valid, errors, warnings = app.validate_column_mapping(self.sample_df, dup_mapping)
        self.assertTrue(is_valid)
        self.assertTrue(any('multiple required fields' in w for w in warnings))

    def test_bug4_input_validation_catches_non_numeric_values(self):
        """Bug #4: Warn when non-numeric text column is selected for cost or qty."""
        df_with_text = pd.DataFrame({
            'Name': ['Item A', 'Item B'],
            'TextCost': ['Free', 'Gift'],
            'Qty': [1, 2]
        })
        mapping = {'name': 'Name', 'cost': 'TextCost', 'qty': 'Qty'}
        is_valid, errors, warnings = app.validate_column_mapping(df_with_text, mapping)
        self.assertTrue(any('TextCost' in w and 'numerical' in w for w in warnings))

    def test_bug4_process_analytics_guards_against_missing_columns(self):
        """Bug #4: process_analytics should raise ValueError/KeyError with clear message rather than unhandled KeyError."""
        invalid_mapping = {'name': 'Product Title', 'cost': 'Selling Price'}
        with self.assertRaises((ValueError, KeyError)) as ctx:
            app.process_analytics(self.sample_df, invalid_mapping)
        self.assertTrue('qty' in str(ctx.exception).lower() or 'missing' in str(ctx.exception).lower())

        invalid_col_mapping = {'name': 'Product Title', 'cost': 'NonExistent', 'qty': 'Qty Sold'}
        with self.assertRaises(KeyError) as ctx:
            app.process_analytics(self.sample_df, invalid_col_mapping)
        self.assertIn('NonExistent', str(ctx.exception))

    def test_bug3_error_handling_empty_dataset(self):
        """Bug #3: Handling empty datasets gracefully."""
        empty_df = pd.DataFrame()
        is_valid, errors, _ = app.validate_column_mapping(empty_df, self.valid_mapping)
        self.assertFalse(is_valid)
        self.assertTrue(any('empty' in e.lower() for e in errors))

        with self.assertRaises(ValueError):
            app.process_analytics(empty_df, self.valid_mapping)

    def test_bug3_load_data_csv_encoding_fallback(self):
        """Bug #3: Verify load_data works with UTF-8, Latin-1, etc."""
        csv_data_utf8 = "Product Name,Price,Quantity\nPanjabi,1500,2\n".encode("utf-8")
        df_utf8 = app.load_data(csv_data_utf8, "test.csv")
        self.assertEqual(len(df_utf8), 1)

        csv_data_latin1 = "Product Name,Price,Quantity\nPanjabi,1500,2\n".encode("latin1")
        df_latin1 = app.load_data(csv_data_latin1, "test_latin.csv")
        self.assertEqual(len(df_latin1), 1)

        with self.assertRaises(ValueError):
            app.load_data(b"", "empty.csv")

    def test_bug2_preset_and_preference_preservation(self):
        """Bug #2: Verify preset and preference save/load functions."""
        test_preset_name = "__test_preset_123__"
        saved = app.save_preset(test_preset_name, self.valid_mapping)
        self.assertTrue(saved)

        presets = app.load_presets()
        self.assertIn(test_preset_name, presets)
        self.assertEqual(presets[test_preset_name]['name'], 'Product Title')

        app.save_last_mapping(self.valid_mapping)
        last_map = app.load_last_mapping()
        self.assertIsNotNone(last_map)
        self.assertEqual(last_map['name'], 'Product Title')

        deleted = app.delete_preset(test_preset_name)
        self.assertTrue(deleted)
        self.assertNotIn(test_preset_name, app.load_presets())

    def test_clean_numeric_series(self):
        """Ensure currency, commas, negative parentheses, and dirty inputs are handled."""
        series = pd.Series(["$1,500.50", "TK (250)", " 3,000 ", "-100", "invalid"])
        cleaned = app.clean_numeric_series(series)
        expected = [1500.50, -250.0, 3000.0, -100.0, 0.0]
        np.testing.assert_allclose(cleaned.values, expected)

    def test_successful_analytics_processing_and_excel(self):
        """Verify normal processing and Excel report generation work without error."""
        results = app.process_analytics(self.sample_df, self.valid_mapping)
        self.assertIn('summary', results)
        self.assertIn('drilldown', results)
        self.assertIn('product_summary', results)
        self.assertIn('total_rev', results)
        self.assertEqual(results['total_rev'], (1500*2 + 450*5 + 1200*1 + 300*3))
        self.assertEqual(results['total_qty'], (2 + 5 + 1 + 3))

        sheet_dict = {
            'Category Summary': results['summary'],
            'Price-wise Category': results['drilldown'],
            'Category-wise Product': results['product_summary']
        }
        excel_bytes = app.generate_excel_report(sheet_dict)
        self.assertIsInstance(excel_bytes, bytes)
        self.assertGreater(len(excel_bytes), 0)

if __name__ == '__main__':
    unittest.main()
