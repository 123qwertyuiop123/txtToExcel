"""金额的统一舍入、显示和 Excel 数值转换工具。"""

from decimal import ROUND_HALF_UP, Decimal, InvalidOperation


def decimal_to_number(value: Decimal) -> int | float:
    """整数转为 int，其余转为 float，使 Excel 不显示无意义的小数点。"""
    return int(value) if value == value.to_integral_value() else float(value)


def round_one_decimal(value: Decimal) -> Decimal:
    """按财务常见的四舍五入规则保留一位小数。"""
    try:
        return value.quantize(Decimal("0.1"), rounding=ROUND_HALF_UP)
    except InvalidOperation as exc:
        # 超长金额也必须给出可展示的错误，而不是让 GUI 回调抛出未处理异常。
        raise ValueError(f"金额超出可处理范围：{value}") from exc


def decimal_to_text(value: Decimal) -> str:
    """生成 TXT 明细文本，删除小数末尾多余的零。"""
    text = format(value, "f")
    return text.rstrip("0").rstrip(".") if "." in text else text
