import unittest
from decimal import Decimal
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch

from expense_converter import (
    convert_many,
    expense_detail,
    parse_money_items,
    parse_txt,
    validate_many,
)
from game_converter import convert_game_txt_files, parse_game_txt, validate_game_txt_files
from models import DayRecord
from year_merge import _rewrite_year_formula, infer_year_from_workbook_name


class ParsingTests(unittest.TestCase):
    def test_game_txt_accepts_optional_year_and_colon_after_year(self):
        with TemporaryDirectory() as directory:
            source = Path(directory) / "2025年游戏.txt"
            source.write_text(
                "1月1日:原神:68\n2025年2月2日:崩坏:60\n2025年:3月3日:原神:30",
                encoding="utf-8",
            )
            records = parse_game_txt(source)
            self.assertEqual([record.purchase_date.month for record in records], [1, 2, 3])
            self.assertEqual(sum(record.amount for record in records), Decimal("158"))

    def test_game_validation_reports_file_and_line_for_year_mismatch(self):
        with TemporaryDirectory() as directory:
            source = Path(directory) / "2025年游戏.txt"
            source.write_text("2024年1月1日:原神:68", encoding="utf-8")
            message = "\n".join(validate_game_txt_files([source]))
            self.assertIn("2025年游戏.txt", message)
            self.assertIn("第 1 行", message)
            self.assertIn("不一致", message)

    def test_game_conversion_combines_multiple_years(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            first = root / "2024年游戏.txt"
            second = root / "2025年游戏.txt"
            first.write_text("1月1日:原神:68", encoding="utf-8")
            second.write_text("2月2日:崩坏:60", encoding="utf-8")
            output = root / "游戏消费统计.xlsx"
            with patch("game_converter.create_game_workbook") as create:
                result = convert_game_txt_files([first, second], output)
            self.assertEqual(result, output)
            records_by_year, target = create.call_args.args
            self.assertEqual(set(records_by_year), {2024, 2025})
            self.assertEqual(target, output)

    def test_income_only_when_plus_prefix_is_present(self):
        detail = "父母:+600,妈妈:-520,购物:98.6"
        record = DayRecord(1, detail, parse_money_items(detail))
        self.assertEqual(record.income, Decimal("600"))
        self.assertEqual(record.expense, Decimal("618.6"))

    def test_chinese_punctuation(self):
        items = parse_money_items("早餐：8，工资：+100")
        self.assertEqual(len(items), 2)
        self.assertFalse(items[0].is_income)
        self.assertTrue(items[1].is_income)

    def test_income_is_removed_from_daily_detail(self):
        detail = "早饭:4.3,午饭:14,父母:+600"
        record = DayRecord(8, detail, parse_money_items(detail))
        self.assertEqual(expense_detail(record), "早饭:4.3,午饭:14")

    def test_missing_colon_before_amount_is_tolerated(self):
        detail = "早饭:5.5,午饭:17,晚饭10,父母:+600,夜宵:19.9"
        record = DayRecord(14, detail, parse_money_items(detail))
        self.assertEqual(expense_detail(record), "早饭:5.5,午饭:17,晚饭:10,夜宵:19.9")
        self.assertEqual(record.expense, Decimal("52.4"))

    def test_missing_comma_between_items_is_tolerated(self):
        detail = "午饭:17晚饭10父母:+600,夜宵:19.9"
        record = DayRecord(14, detail, parse_money_items(detail))
        self.assertEqual(expense_detail(record), "午饭:17,晚饭:10,夜宵:19.9")
        self.assertEqual(record.expense, Decimal("46.9"))
        self.assertEqual(record.income, Decimal("600"))

    def test_unrecognized_text_raises_instead_of_being_dropped(self):
        with self.assertRaisesRegex(ValueError, "无法识别明细项目"):
            parse_money_items("午饭:17,这一项没有金额")

    def test_comma_used_as_decimal_point_is_tolerated(self):
        items = parse_money_items("晚饭:13,1,生鲜:6")
        self.assertEqual(items[0].amount, Decimal("13.1"))
        self.assertEqual(items[1].amount, Decimal("6"))

    def test_comma_used_instead_of_colon_is_tolerated(self):
        items = parse_money_items("早饭:4,午饭,14.7,晚饭:12.9")
        self.assertEqual([(item.label, item.amount) for item in items], [
            ("早饭", Decimal("4")),
            ("午饭", Decimal("14.7")),
            ("晚饭", Decimal("12.9")),
        ])

    def test_multiple_month_files_are_combined(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            january = root / "2025年1月消费.txt"
            february = root / "2025年2月消费.txt"
            january.write_text("1日:早饭:8", encoding="utf-8")
            february.write_text("1日:午饭:12", encoding="utf-8")
            output = root / "result.xlsx"
            with patch("expense_converter.create_workbook") as create:
                result = convert_many([january, february], output, None, None)
            self.assertEqual(result, output)
            month_records, year, target = create.call_args.args
            self.assertEqual(set(month_records), {1, 2})
            self.assertEqual(year, 2025)
            self.assertEqual(target, output)

    def test_parse_error_includes_filename_line_and_day(self):
        with TemporaryDirectory() as directory:
            source = Path(directory) / "2025年7月消费.txt"
            source.write_text("1日:早饭:8\n2日:晚饭:7+2+20+4", encoding="utf-8")
            with self.assertRaises(ValueError) as raised:
                parse_txt(source)
            message = str(raised.exception)
            self.assertIn("2025年7月消费.txt", message)
            self.assertIn("第 2 行", message)
            self.assertIn("2日", message)

    def test_validation_reports_file_line_day_and_bad_item(self):
        with TemporaryDirectory() as directory:
            source = Path(directory) / "2025年7月消费.txt"
            source.write_text("1日:早饭:8\n2日:晚饭7+2+20+4", encoding="utf-8")
            errors = validate_many([source], None, None)
            message = "\n".join(errors)
            self.assertIn("2025年7月消费.txt", message)
            self.assertIn("第 2 行", message)
            self.assertIn("2日", message)
            self.assertIn("晚饭7+2+20+4", message)

    def test_validation_accepts_multiple_standard_files(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            january = root / "2025年1月消费.txt"
            february = root / "2025年2月消费.txt"
            january.write_text("1日:早饭:8,午饭:12.5", encoding="utf-8")
            february.write_text("1日:工资:+600,晚饭:10", encoding="utf-8")
            self.assertEqual(validate_many([january, february], None, None), [])

    def test_year_is_inferred_from_workbook_filename(self):
        self.assertEqual(infer_year_from_workbook_name(Path("2025年消费统计.xlsx")), 2025)

    def test_month_formula_is_rewritten_with_year_prefix(self):
        self.assertEqual(
            _rewrite_year_formula("='1月'!C33", 2025),
            "='2025-1月'!C33",
        )


if __name__ == "__main__":
    unittest.main()
