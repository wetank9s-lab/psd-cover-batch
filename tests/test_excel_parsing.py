# -*- coding: utf-8 -*-
"""core.excel_data 的单元测试（Stage 0）。"""

import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from core.excel_data import parse_row, parse_rows, format_phone, unique_stores


# ---------------------------------------------------------------------------
# format_phone
# ---------------------------------------------------------------------------
def test_phone_int():
    assert format_phone(13800000001) == "13800000001"


def test_phone_float_integer():
    # Excel 常见：13800000001.0 应为 '13800000001'，不出现 .0
    assert format_phone(13800000001.0) == "13800000001"


def test_phone_float_non_integer():
    assert format_phone(123.45) == "123.45"


def test_phone_string_strip():
    assert format_phone("  18090082479  ") == "18090082479"


def test_phone_none():
    assert format_phone(None) == ""


# ---------------------------------------------------------------------------
# parse_row
# ---------------------------------------------------------------------------
def test_parse_row_basic():
    r = parse_row(("易田电器", "王兵", "销售顾问", 13800000001),
                  excel_row=2, col_store=0, col_name=1, col_phone=3, col_role=2)
    assert r is not None
    assert r.excel_row == 2
    assert r.store == "易田电器"
    assert r.name == "王兵"
    assert r.phone == "13800000001"
    assert r.role == "销售顾问"


def test_parse_row_phone_float():
    r = parse_row(("康乐", "张三", "", 13800000001.0), excel_row=3,
                  col_store=0, col_name=1, col_phone=3, col_role=2)
    assert r is not None
    assert r.phone == "13800000001"


def test_parse_row_empty_name_skipped():
    r = parse_row(("易田电器", "   ", "销售顾问", 13800000001), excel_row=2)
    assert r is None


def test_parse_row_empty_phone_skipped():
    r = parse_row(("易田电器", "王兵", "销售顾问", None), excel_row=2)
    assert r is None


def test_parse_row_col_out_of_range():
    # 只有 2 列，phone 列索引 3 越界 -> phone 为空 -> 按必填跳过
    r = parse_row(("易田电器", "王兵"), excel_row=2, col_store=0, col_name=1, col_phone=3)
    assert r is None


def test_parse_row_store_not_in_col_a():
    # 门店在 B 列（索引 1）
    r = parse_row(("王兵", "易田电器", "销售顾问", 13800000001), excel_row=2,
                  col_store=1, col_name=0, col_phone=3, col_role=2)
    assert r is not None
    assert r.store == "易田电器"
    assert r.name == "王兵"


# ---------------------------------------------------------------------------
# parse_rows（含表头逻辑）
# ---------------------------------------------------------------------------
def test_parse_rows_with_header():
    rows = [
        ("门店", "姓名", "销售顾问", "电话"),
        ("易田电器", "王兵", "销售顾问", 13800000001),
        ("康乐", "张三", "销售顾问", 13800000002),
    ]
    valid, skipped = parse_rows(rows, has_header=True)
    assert len(valid) == 2
    assert valid[0].excel_row == 2
    assert valid[1].excel_row == 3
    assert skipped == []


def test_parse_rows_without_header():
    rows = [
        ("易田电器", "王兵", "销售顾问", 13800000001),
        ("康乐", "张三", "销售顾问", 13800000002),
    ]
    valid, skipped = parse_rows(rows, has_header=False)
    assert len(valid) == 2
    assert valid[0].excel_row == 1  # 无表头时第一行就是数据


def test_parse_rows_skips_empty():
    rows = [
        ("门店", "姓名", "销售顾问", "电话"),
        ("易田电器", "王兵", "销售顾问", 13800000001),
        (None, None, None, None),                       # 空行
        ("康乐", "", "销售顾问", 13800000002),           # 姓名为空
    ]
    valid, skipped = parse_rows(rows, has_header=True)
    assert len(valid) == 1
    assert valid[0].name == "王兵"
    assert skipped == [3, 4]


def test_parse_rows_20plus_columns():
    wide = ["门店", "姓名", "销售顾问", "电话"] + [f"X{i}" for i in range(16)]
    data = ["易田电器", "王兵", "销售顾问", 13800000001] + [f"x{i}" for i in range(16)]
    valid, skipped = parse_rows([wide, data], has_header=True)
    assert len(valid) == 1
    assert valid[0].phone == "13800000001"


def test_unique_stores_dedup_order():
    rows = [
        ("易田电器", "王兵", "销售顾问", 13800000001),
        ("康乐", "张三", "销售顾问", 13800000002),
        ("易田电器", "李四", "销售顾问", 13800000003),
    ]
    assert unique_stores(rows, col_store=0) == ["易田电器", "康乐"]


def test_unique_stores_skips_empty():
    rows = [
        ("易田电器", "王兵", "销售顾问", 13800000001),
        (None, "张三", "销售顾问", 13800000002),
    ]
    assert unique_stores(rows, col_store=0) == ["易田电器"]
