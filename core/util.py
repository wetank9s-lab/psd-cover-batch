# -*- coding: utf-8 -*-
"""
util.py —— 通用纯函数：文件名清洗、门店名/图层名归一化。

Stage 0：从 GUI（qifang_cover_maker.py）与 CLI（psd_cover_batch.py）中
抽取出的「不依赖 Photoshop COM、不依赖 Tk」的纯函数。抽取时保持行为一致，
供两侧复用，并为后续 Stage 提供可单元测试的基础。

Stage 7：删除历史 name-based Logo 匹配（fuzzy_contains / match_store_to_layers /
pick_best_logo）——已由 core.logo_mapping.match_store_logo（评分制、确定性、
歧义不自动选）完全取代；本模块只保留文件名清洗与名称归一化。
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
# 兼容旧名（CLI 的 sanitize 语义与 GUI 文件名片段一致）
# ---------------------------------------------------------------------------
def sanitize(name):
    """兼容旧 CLI 的 sanitize()：只移除非法字符并 strip。"""
    return sanitize_filename(name)
