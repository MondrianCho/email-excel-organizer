import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from email_excel_organize import get_new_filename, is_excel_file, match_e_rules, safe_str


class TestSafeStr(unittest.TestCase):
    def test_none_returns_empty_string(self):
        self.assertEqual(safe_str(None), "")

    def test_strips_whitespace(self):
        self.assertEqual(safe_str("  hello  "), "hello")

    def test_non_string_is_stringified(self):
        self.assertEqual(safe_str(123), "123")


class TestIsExcelFile(unittest.TestCase):
    def test_accepts_xlsx(self):
        self.assertTrue(is_excel_file(Path("report.xlsx")))

    def test_accepts_xlsm(self):
        self.assertTrue(is_excel_file(Path("report.xlsm")))

    def test_rejects_other_extensions(self):
        self.assertFalse(is_excel_file(Path("report.csv")))

    def test_rejects_temp_lock_files(self):
        self.assertFalse(is_excel_file(Path("~$report.xlsx")))


class TestMatchERules(unittest.TestCase):
    def test_ack_keyword_is_case_insensitive(self):
        self.assertEqual(match_e_rules("automatic reply: out of office"), ("ACK", None, None, None))

    def test_maintenance_reminder_takes_priority_over_reminder(self):
        k, l, n, o = match_e_rules("Maintenance Reminder for patent 12345")
        self.assertEqual((k, l), ("비용", "(비용)연차료/RMD"))

    def test_annuity_keyword_maps_to_cost_category(self):
        k, l, n, o = match_e_rules("Annuity payment due")
        self.assertEqual(k, "비용")

    def test_no_match_returns_none(self):
        self.assertIsNone(match_e_rules("전혀 관련 없는 제목"))


class TestGetNewFilename(unittest.TestCase):
    def test_uses_current_date_when_not_using_mod_date(self):
        from datetime import datetime

        name = get_new_filename(Path("dummy.xlsx"), use_mod_date=False)
        today = datetime.now().strftime("%Y-%m-%d")
        self.assertEqual(name, f"{today} 이메일 접수_초안.xlsx")


if __name__ == "__main__":
    unittest.main()
