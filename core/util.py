# -*- coding: utf-8 -*-
"""
util.py —— 通用纯函数：文件名清洗、门店名/图层名归一化、Logo 模糊匹配。

Stage 0：从 GUI（qifang_cover_maker.py）与 CLI（psd_cover_batch.py）中
抽取出的「不依赖 Photoshop COM、不依赖 Tk」的纯函数。抽取时保持行为一致，
供两侧复用，并为后续 Stage 提供可单元测试的基础。
"""

import os
import re
import unicodedata

# ---------------------------------------------------------------------------
# 文件名清洗（Windows）
# ---------------------------------------------------------------------------
# Windows 非法字符（含 CLI 原实现已过滤的 / \ : * ? " < > | 与 GUI 只过滤的 / \）
_INVALID_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
# Windows 保留设备名（不含扩展名比较）
_RESERVED_NAMES = {
    "CON", "PRN", "AUX", "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}


def sanitize_filename(name, fallback="unnamed", max_len=120):
    """清洗为可在 Windows 使用的文件名片段。

    - 移除 Windows 非法字符 `<>:"/\\|?*` 与控制字符；
    - 去除首尾空白；
    - 去除结尾的 `.` 与空格（Windows 不允许文件名以点或空格结尾）；
    - 保留设备名（CON / PRN / AUX / NUL / COM1.. / LPT1..）时前面加 `_` 前缀；
    - 超长截断到 max_len；
    - 清洗后为空则返回 fallback（默认 "unnamed"）。

    注意：本函数只清洗「单个文件名片段」，不含扩展名与路径分隔符。
    """
    if name is None:
        return fallback
    s = str(name)
    # 移除非法字符 + 控制字符
    s = _INVALID_FILENAME_CHARS.sub("", s)
    # 去首尾空白；再单独去除结尾点/空格（strip('.') 会去掉多个，这里保留内部点）
    s = s.strip()
    while s.endswith((".", " ")):
        s = s[:-1]
    # 设备名保护：保留名加下划线前缀，避免 Windows 拒绝
    stem = s.upper()
    if stem in _RESERVED_NAMES:
        s = "_" + s
    # 截断（避免路径过长）
    if max_len and len(s) > max_len:
        s = s[:max_len]
    # 空结果回退
    if not s:
        s = fallback
    return s


# ---------------------------------------------------------------------------
# 门店名 / 图层名归一化
# ---------------------------------------------------------------------------
def normalize_name(text):
    """归一化门店名 / 图层名，用于模糊匹配。

    - NFKC 归一化（全角→半角、兼容字符折叠）；
    - 去首尾空白；
    - 转小写；
    - 移除所有 Unicode 空白字符（含全角空格、不断行空格）。
    """
    if text is None:
        return ""
    t = unicodedata.normalize("NFKC", str(text))
    t = t.strip().lower()
    # 移除所有空白字符（含 \u3000 全角空格、\xa0 等）
    t = re.sub(r"\s+", "", t)
    return t


# ---------------------------------------------------------------------------
# Logo 模糊匹配
# ---------------------------------------------------------------------------
def fuzzy_contains(a, b):
    """互含匹配：a 与 b 归一化后互相包含即视为命中。

    保留与旧实现一致的行为（归一化互含），供 CLI / GUI 沿用。
    注意：本函数不处理「歧义」；有多个候选时由调用方决定。
    """
    na, nb = normalize_name(a), normalize_name(b)
    return bool(na) and bool(nb) and (na in nb or nb in na)


def match_store_to_layers(store, layer_names):
    """返回与门店互含匹配的图层名列表（按传入顺序，保留全部命中）。

    参数：
      store:       Excel 门店名（str）
      layer_names: 可迭代的候选图层名（list[str]）

    返回：
      list[str] —— 命中的图层名。可能有 0 个、1 个或多个（歧义时由调用方处理）。
    """
    if not store:
        return []
    hits = []
    for nm in layer_names:
        if nm and fuzzy_contains(store, nm):
            hits.append(nm)
    return hits


def pick_best_logo(store, layer_names):
    """从候选中选一个最合适的 Logo 图层名。

    优先级（评分制，与文档 P1-02 一致）：
      - 归一化后完全相等         100
      - 图层名以门店名开头         90
      - 门店名以图层名开头         85
      - 互含匹配                  70
    如果最高分出现并列（如 康乐 → 康乐电器 / 康乐家电 同分 90），返回 None，
    表示「需要人工确认」，绝不自动取第一个。

    返回：
      str | None —— 唯一最优的图层名；无匹配或歧义时返回 None。
    """
    target = normalize_name(store)
    if not target:
        return None
    best_name = None
    best_score = -1
    tied = False
    for nm in layer_names:
        n = normalize_name(nm)
        if not n:
            continue
        if n == target:
            score = 100
        elif n.startswith(target):
            score = 90
        elif target.startswith(n):
            score = 85
        elif target in n or n in target:
            score = 70
        else:
            continue
        if score > best_score:
            best_score = score
            best_name = nm
            tied = False
        elif score == best_score:
            tied = True  # 同分：歧义，不自动选
    if best_name is None or tied:
        return None
    return best_name


# 保留旧名（CLI 的 sanitize 语义与 GUI 文件名片段一致，供最小改动）
def sanitize(name):
    """兼容旧 CLI 的 sanitize()：只移除非法字符并 strip。"""
    return sanitize_filename(name)
