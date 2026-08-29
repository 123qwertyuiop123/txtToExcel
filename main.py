"""程序启动入口；具体业务分别位于独立模块中。"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Iterable

from expense_converter import convert_many
from gui import launch_gui

# 保留旧版从 main 导入这些名称的兼容性，同时不在入口文件重复实现业务逻辑。
from expense_converter import (  # noqa: F401
    FILE_DATE_RE,
    convert,
    create_workbook,
    expense_detail,
    find_txt_files,
    infer_year_month,
    parse_money_items,
    parse_txt,
    validate_many,
    validate_txt_standard,
)
from game_converter import (  # noqa: F401
    convert_game_txt_files,
    create_game_workbook,
    infer_game_year,
    parse_game_txt,
    validate_game_txt_files,
)
from models import DayRecord, GameRecord, MoneyItem  # noqa: F401
from money_utils import (  # noqa: F401
    decimal_to_number as _decimal_to_number,
    decimal_to_text as _decimal_to_text,
    round_one_decimal as _round_one_decimal,
)
from year_merge import (  # noqa: F401
    _rewrite_year_formula,
    infer_year_from_workbook_name,
    merge_yearly_workbooks,
    validate_yearly_workbooks,
)


def build_parser() -> argparse.ArgumentParser:
    """创建命令行参数解析器。"""
    parser = argparse.ArgumentParser(description="把消费 TXT 转换为带月度工作表和年度总计的 Excel")
    parser.add_argument("input", nargs="*", type=Path, help="一个或多个 TXT 文件，或包含月份 TXT 的目录")
    parser.add_argument("-o", "--output", type=Path, help="输出的 .xlsx 路径")
    parser.add_argument("--year", type=int, help="文件名不含年份时手动指定")
    parser.add_argument("--month", type=int, help="转换单个文件且文件名不含月份时手动指定")
    parser.add_argument("--gui", action="store_true", help="打开图形窗口")
    return parser


def main(argv: Iterable[str] | None = None) -> int:
    """根据参数启动 GUI，或执行普通消费命令行转换。"""
    args = build_parser().parse_args(argv)
    if args.gui or not args.input:
        return launch_gui()
    try:
        result = convert_many(args.input, args.output, args.year, args.month)
    except (OSError, ValueError, RuntimeError) as exc:
        print(f"转换失败：{exc}", file=sys.stderr)
        return 1
    print(f"转换完成：{result.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
