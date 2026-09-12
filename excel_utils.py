"""Excel 输出的安全工具：文本防公式注入与临时文件原子保存。"""

import os
import tempfile
from pathlib import Path
from typing import Any, Iterable


FORMULA_PREFIXES = ("=", "+", "-", "@")


def validate_output_path(output_path: Path, sources: Iterable[Path] = ()) -> Path:
    """限制输出扩展名，并阻止结果覆盖输入文件（包括符号链接和硬链接）。"""
    output_path = Path(output_path)
    if output_path.suffix.lower() != ".xlsx":
        raise ValueError(f"输出文件必须使用 .xlsx 扩展名：{output_path}")
    for source in sources:
        source = Path(source)
        same_path = output_path.resolve() == source.resolve()
        same_file = output_path.exists() and source.exists() and os.path.samefile(output_path, source)
        if same_path or same_file:
            raise ValueError(f"输出文件不能覆盖输入文件：{source}")
    return output_path


def set_safe_text(cell: Any, value: str | None) -> None:
    """把外部文字明确写成字符串，阻止 Excel 将其解释为公式。

    XLSX 单元格本身带有数据类型，因此不需要在内容前添加会显示出来的
    单引号；把 data_type 固定为 ``s`` 即可保留原文并禁用公式执行。
    """
    cell.value = value
    if isinstance(value, str) and value.startswith(FORMULA_PREFIXES):
        cell.data_type = "s"


def save_workbook_safely(workbook: Any, output_path: Path) -> None:
    """先写同目录临时文件，完整成功后再替换目标工作簿。

    临时文件与目标位于同一磁盘，使 ``os.replace`` 可以原子替换。同名旧文件
    在新文件完全写好之前不会受到影响；保存或替换失败时会删除临时文件。
    """
    output_path = validate_output_path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    file_descriptor, temporary_name = tempfile.mkstemp(
        prefix=f".{output_path.stem}-",
        suffix=".tmp.xlsx",
        dir=output_path.parent,
    )
    os.close(file_descriptor)
    temporary_path = Path(temporary_name)
    try:
        workbook.save(temporary_path)
        # 同目录替换在 Windows 上具有原子性；目标被 Excel 占用时会安全失败。
        os.replace(temporary_path, output_path)
    finally:
        # replace 成功后临时路径已不存在；失败时清理残留，不触碰旧目标。
        temporary_path.unlink(missing_ok=True)
