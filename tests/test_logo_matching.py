# -*- coding: utf-8 -*-
"""core.util 的 Logo 匹配单元测试（Stage 0）。

覆盖：归一化、互含匹配、打分选优、歧义不自动选。
"""

import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from core.util import (
    normalize_name,
    fuzzy_contains,
    match_store_to_layers,
    pick_best_logo,
)


def test_normalize_strips_spaces_and_case():
    assert normalize_name("康 乐") == "康乐"
    assert normalize_name("ＡＢＣ") == "abc"
    assert normalize_name(" 易田电器 ") == "易田电器"


def test_match_kanle_to_kangle_dianqi():
    hits = match_store_to_layers("康乐", ["康乐电器", "易田电器", "九兴电器"])
    assert hits == ["康乐电器"]


def test_match_mutual_contains_both_directions():
    # 门店名短、图层名长；以及门店名长、图层名短
    assert fuzzy_contains("康乐", "康乐电器")
    assert fuzzy_contains("康乐电器", "康乐")


def test_match_jiuxing():
    hits = match_store_to_layers("九兴", ["九兴电器", "欣盛电器"])
    assert hits == ["九兴电器"]


def test_match_empty_store():
    assert match_store_to_layers("", ["康乐电器"]) == []
    assert match_store_to_layers("康乐", []) == []


def test_pick_best_unique():
    # 唯一最优：康乐 -> 康乐电器（startswith 90 分）
    assert pick_best_logo("康乐", ["康乐电器", "易田电器"]) == "康乐电器"


def test_pick_best_exact():
    # 归一化后完全相等：100 分
    assert pick_best_logo("易田电器", ["易田电器", "康乐电器"]) == "易田电器"


def test_pick_best_ambiguous_returns_none():
    # 康乐 -> 康乐电器 / 康乐家电 都是 90 分，歧义 -> None（绝不自动选第一个）
    assert pick_best_logo("康乐", ["康乐电器", "康乐家电"]) is None


def test_pick_best_no_match_returns_none():
    assert pick_best_logo("东山", ["易田电器", "欣盛电器"]) is None


def test_pick_best_empty_input():
    assert pick_best_logo("", ["康乐电器"]) is None
    assert pick_best_logo("康乐", []) is None
