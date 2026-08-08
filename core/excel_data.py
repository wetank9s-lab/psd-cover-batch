# -*- coding: utf-8 -*-
"""
excel_data.py —— Excel 行数据的纯解析函数（Stage 0：仅抽取纯函数，不改行为）。

Stage 0 目标：把「对 openpyxl 行元组的解析」抽成不依赖文件、不依赖 GUI 的
纯函数，便于 pytest 覆盖（表头 / 无表头 / 门店非 A 列 / 空行 / 电话类型 /
列越界 等）。GUI 与 CLI 暂不切换（后续 Stage 4 再统一接入 load_dataset）。

电话规范（与文档第 8 节一致）：
  - int           -> str
  - float 且为整数 -> int -> str
  - 其他数值       -> 保留原样 str（避免 1.23e10 失真）
  - str / 其他     -> strip
"""

from dataclasses import dataclass, field
from typing import Any, Optional


@dataclass
class ExcelRow:
    """从 Excel 一行解析出的业务数据（Stage 0 纯数据，供测试与后续复用）。"""

    excel_row: int            # Excel 中的物理行号（1 起，含表头行；便于用户核对）
    store: str
    name: str
    phone: str
    role: Optional[str] = None


def format_phone(value: Any) -> str:
    """把 Excel 单元格的电话值规范成字符串。

    - int           -> str(value)
    - float 且是整数 -> str(int(value))     （如 13800000001.0 -> '13800000001'）
    - float 非整数   -> str(value)          （保留原样，避免失真）
    - str / 其他     -> str(value).strip()
    """
    if value is None:
        return ""
    if isinstance(value, bool):
        return "1" if value else "0"
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return str(value)
    return str(value).strip()


def parse_row(
    raw: tuple,
    *,
    excel_row: int,
    col_store: int = 0,
    col_name: int = 1,
    col_phone: int = 3,
    col_role: int = 2,
    require_name: bool = True,
    require_phone: bool = True,
) -> Optional[ExcelRow]:
    """把 openpyxl 的一行（tuple）解析为 ExcelRow。

    - 列索引越界时按「空值」处理；
    - 空白字符串视为空值；
    - name 为空（或未配置必填 phone 时 phone 为空）返回 None（跳过）；
    - 保留物理行号 excel_row。
    """
    if raw is None:
        return None
    n = len(raw)

    def cell(idx):
        return raw[idx] if 0 <= idx < n else None

    store = cell(col_store)
    name = cell(col_name)
    phone = cell(col_phone)
    role = cell(col_role) if col_role is not None and col_role >= 0 else None

    store_s = str(store).strip() if store is not None else ""
    name_s = str(name).strip() if name is not None else ""
    role_s = str(role).strip() if role is not None else ""
    phone_s = format_phone(phone)

    if require_name and not name_s:
        return None
    if require_phone and not phone_s:
        return None
    return ExcelRow(
        excel_row=excel_row,
        store=store_s,
        name=name_s,
        phone=phone_s,
        role=role_s if role_s else None,
    )


def parse_rows(
    rows,
    *,
    has_header: bool = True,
    col_store: int = 0,
    col_name: int = 1,
    col_phone: int = 3,
    col_role: int = 2,
    require_name: bool = True,
    require_phone: bool = True,
):
    """批量解析 openpyxl 行（iter_rows(values_only=True) 的结果）。

    返回 (valid: list[ExcelRow], skipped: list[int])：
      valid  —— 通过校验的行；
      skipped—— 被跳过的物理行号列表（1 起）。
    """
    if rows is None:
        rows = []
    start = 1 if has_header else 0
    valid = []
    skipped = []
    for offset, raw in enumerate(rows[start:]):
        physical = start + offset + 1  # 1 起物理行号（含表头行）
        r = parse_row(
            raw,
            excel_row=physical,
            col_store=col_store,
            col_name=col_name,
            col_phone=col_phone,
            col_role=col_role,
            require_name=require_name,
            require_phone=require_phone,
        )
        if r is None:
            skipped.append(physical)
        else:
            valid.append(r)
    return valid, skipped


def unique_stores(rows, col_store: int = 0):
    """从 openpyxl 行中提取去重门店名（保序，跳过空值）。"""
    out = []
    for r in rows:
        if r and len(r) > col_store and r[col_store] is not None:
            s = str(r[col_store]).strip()
            if s and s not in out:
                out.append(s)
    return out
