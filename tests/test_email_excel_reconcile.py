import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from email_excel_reconcile import is_excel_file, normalize_title, safe_str


class TestSafeStr(unittest.TestCase):
    def test_none_returns_empty_string(self):
        self.assertEqual(safe_str(None), "")

    def test_strips_whitespace(self):
        self.assertEqual(safe_str("  hello  "), "hello")


class TestNormalizeTitle(unittest.TestCase):
    def test_empty_string_returns_empty(self):
        self.assertEqual(normalize_title(""), "")

    def test_collapses_internal_whitespace(self):
        self.assertEqual(normalize_title("hello   world"), "hello world")

    def test_strips_leading_and_trailing_whitespace(self):
        self.assertEqual(normalize_title("  hello world  "), "hello world")


class TestIsExcelFile(unittest.TestCase):
    def test_accepts_xlsx(self):
        self.assertTrue(is_excel_file(Path("data.xlsx")))

    def test_rejects_temp_lock_files(self):
        self.assertFalse(is_excel_file(Path("~$data.xlsx")))

    def test_rejects_non_excel_extension(self):
        self.assertFalse(is_excel_file(Path("data.txt")))


if __name__ == "__main__":
    unittest.main()
