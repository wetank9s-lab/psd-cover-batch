# -*- coding: utf-8 -*-
"""core.util.sanitize_filename 的单元测试（Stage 0）。"""

import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from core.util import sanitize_filename, normalize_name, fuzzy_contains


# ---------------------------------------------------------------------------
# sanitize_filename
# ---------------------------------------------------------------------------
def test_removes_illegal_chars():
    # Windows 非法字符 < > : " / \ | ? * 与控制字符
    assert sanitize_filename('A/B') == "AB"
    assert sanitize_filename('A:B') == "AB"
    assert sanitize_filename('张三?') == "张三"
    assert sanitize_filename('a<b>c|d"e') == "abcde"
    assert sanitize_filename('x\x00y') == "xy"


def test_strips_leading_trailing_whitespace():
    assert sanitize_filename("  张三  ") == "张三"


def test_removes_trailing_dot_and_space():
    # Windows 不允许文件名以点或空格结尾
    assert sanitize_filename("李四.") == "李四"
    assert sanitize_filename("李四. ") == "李四"
    assert sanitize_filename("王五..") == "王五"


def test_reserved_device_names():
    assert sanitize_filename("CON") == "_CON"
    assert sanitize_filename("PRN") == "_PRN"
    assert sanitize_filename("AUX") == "_AUX"
    assert sanitize_filename("NUL") == "_NUL"
    assert sanitize_filename("COM1") == "_COM1"
    assert sanitize_filename("LPT3") == "_LPT3"


def test_empty_and_none_fallback():
    assert sanitize_filename("") == "unnamed"
    assert sanitize_filename("///") == "unnamed"
    assert sanitize_filename(None) == "unnamed"
    assert sanitize_filename("", fallback="未命名") == "未命名"


def test_long_name_truncated():
    long_name = "很" * 200
    out = sanitize_filename(long_name, max_len=120)
    assert len(out) == 120


def test_chinese_and_space_preserved():
    assert sanitize_filename("易田电器") == "易田电器"
    assert sanitize_filename("易田 电器") == "易田 电器"  # 内部空格保留


# ---------------------------------------------------------------------------
# normalize_name / fuzzy_contains
# ---------------------------------------------------------------------------
def test_normalize_fullwidth_and_whitespace():
    # 全角→半角、全角空格、大小写
    assert normalize_name("ＡＢＣ") == "abc"
    assert normalize_name("康 乐") == "康乐"
    assert normalize_name("　康乐　") == "康乐"
    assert normalize_name("KangLe") == "kangle"


def test_fuzzy_contains_mutual():
    # 康乐 <-> 康乐电器
    assert fuzzy_contains("康乐", "康乐电器")
    assert fuzzy_contains("康乐电器", "康乐")
    # 九兴 <-> 九兴电器
    assert fuzzy_contains("九兴", "九兴电器")
    # 完全一致
    assert fuzzy_contains("易田电器", "易田电器")
    # 空值不匹配
    assert not fuzzy_contains("", "康乐电器")
    assert not fuzzy_contains("康乐", "")


def test_fuzzy_not_contains_unrelated():
    assert not fuzzy_contains("东山", "易田电器")
    assert not fuzzy_contains("张三", "李四")
