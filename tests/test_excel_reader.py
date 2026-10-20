import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook

from utils.excel_reader import load_test_data


class LoadTestDataTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp_dir.cleanup)
        self.workbook_path = Path(self.temp_dir.name) / "test-data.xlsx"

    def create_workbook(self, rows, headers=("Item", "Quantities", "Promo_code", "Country")):
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(headers)
        for row in rows:
            sheet.append(row)
        workbook.save(self.workbook_path)
        workbook.close()
        return self.workbook_path

    def test_loads_items_quantities_and_optional_values(self):
        path = self.create_workbook(
            [
                ("Potato", 2, "valid-code", "India"),
                ("Beans", 1, None, None),
            ]
        )

        data = load_test_data(path)

        self.assertEqual(data["vegetables"], ["Potato", "Beans"])
        self.assertEqual(data["quantities"], [2, 1])
        self.assertEqual(data["promocode"], "valid-code")
        self.assertEqual(data["country"], "India")

    def test_uses_test_data_file_environment_variable(self):
        path = self.create_workbook([("Tomato", 3, None, "India")])

        with patch.dict(os.environ, {"TEST_DATA_FILE": str(path)}):
            data = load_test_data()

        self.assertEqual(data["vegetables"], ["Tomato"])
        self.assertEqual(data["quantities"], [3])

    def test_missing_workbook_has_an_actionable_error(self):
        missing_path = Path(self.temp_dir.name) / "missing.xlsx"

        with self.assertRaisesRegex(FileNotFoundError, "missing.xlsx"):
            load_test_data(missing_path)

    def test_corrupt_workbook_has_an_actionable_error(self):
        self.workbook_path.write_text("not an Excel workbook", encoding="utf-8")

        with self.assertRaisesRegex(ValueError, "Could not read test data workbook"):
            load_test_data(self.workbook_path)

    def test_rejects_unexpected_headers(self):
        path = self.create_workbook(
            [("Potato", 2, None, None)],
            headers=("Vegetable", "Count", "Promo", "Country"),
        )

        with self.assertRaisesRegex(ValueError, "Item.*Quantities"):
            load_test_data(path)

    def test_rejects_missing_non_integer_and_non_positive_quantities(self):
        invalid_quantities = (None, "many", 0, -1)
        for quantity in invalid_quantities:
            with self.subTest(quantity=quantity):
                path = self.create_workbook([("Potato", quantity, None, None)])
                with self.assertRaisesRegex(ValueError, "Invalid quantity at row 2"):
                    load_test_data(path)

    def test_rejects_workbook_without_product_rows(self):
        path = self.create_workbook([])

        with self.assertRaisesRegex(ValueError, "no product rows"):
            load_test_data(path)


if __name__ == "__main__":
    unittest.main()