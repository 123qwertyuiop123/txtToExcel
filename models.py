"""消费转换过程中使用的数据模型。"""

from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal


@dataclass(frozen=True)
class MoneyItem:
    """普通消费 TXT 中的一条“项目:金额”。"""

    label: str
    amount: Decimal
    is_income: bool


@dataclass
class DayRecord:
    """某一天的全部普通消费与收入项目。"""

    day: int
    detail: str
    items: list[MoneyItem] = field(default_factory=list)

    @property
    def expense(self) -> Decimal:
        """当天支出合计；负数支出按绝对值计入。"""
        return sum((abs(item.amount) for item in self.items if not item.is_income), Decimal("0"))

    @property
    def income(self) -> Decimal:
        """当天收入合计；只有原始金额以“+”开头时才属于收入。"""
        return sum((abs(item.amount) for item in self.items if item.is_income), Decimal("0"))


@dataclass(frozen=True)
class GameRecord:
    """一笔游戏消费；使用真实日期便于 Excel 排序和筛选。"""

    purchase_date: date
    game: str
    amount: Decimal
