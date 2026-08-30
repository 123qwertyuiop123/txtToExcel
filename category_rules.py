"""消费原因的保守归类规则。

只合并含义明确的名称；未命中规则的项目返回原名称，避免主观猜测。
"""


CATEGORY_GROUPS: dict[str, frozenset[str]] = {
    "餐饮": frozenset({
        "早饭", "早餐", "午饭", "午餐", "晚饭", "晚餐", "夜宵",
        "kfc", "肯德基", "麦当劳", "烧烤", "饮料", "奶茶", "咖啡",
        "外卖", "餐饮",
    }),
    "交通": frozenset({
        "火车票", "高铁票", "飞机票", "机票", "车票", "打车", "出租车",
        "公交", "公交车", "地铁", "交通", "滴滴",
    }),
    "游戏消费": frozenset({"原神", "崩坏", "崩坏3", "游戏", "游戏消费"}),
    "医疗": frozenset({"挂号", "药品", "买药", "医院", "看病", "医疗"}),
    "网络通信": frozenset({"宽带", "话费", "流量", "网费", "网络通信"}),
}


def _normalize_label(label: str) -> str:
    """忽略首尾空白和英文字母大小写，不改变中文内容。"""
    return label.strip().casefold()


_CATEGORY_BY_LABEL = {
    _normalize_label(label): category
    for category, labels in CATEGORY_GROUPS.items()
    for label in labels
}


def categorize_expense(label: str) -> str:
    """返回明确归类；不明确的项目保留原名称。"""
    cleaned = label.strip()
    return _CATEGORY_BY_LABEL.get(_normalize_label(cleaned), cleaned)
