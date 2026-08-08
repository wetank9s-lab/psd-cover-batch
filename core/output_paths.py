# -*- coding: utf-8 -*-
"""
output_paths.py —— 输出分组核心（Stage 4.5 重构）。

职责：把「Excel 任意一列的分组值」安全、稳定地映射为输出子文件夹。
属于 output path 范畴，与 Excel 解析（core/excel_data.py）解耦：
  - ExcelRow 只是输入数据（其 values 已是完整规范化行，见 Stage 4 红线）；
  - 本模块不读取 Excel 文件，也绝不重新访问 Workbook / Worksheet / Cell；
  - 不依赖 core.excel_data（ExcelRow 在此按鸭子类型使用：row.values）。

核心概念：
  - sanitize_group_component : 单个分组值 -> 安全目录片段（不含碰撞后缀）
  - GroupFolderMap          : 批次稳定的「分组值 -> 目录名」映射（只读视图）
  - build_group_folder_map  : 批次开始前一次性建立完整映射（稳定、可统计）
  - resolve_output_directory: 唯一目录解析入口（GUI Batch / GUI Preview / CLI 共用），
                              base_dir 不同（batch=out_dir, preview=out_dir/_preview），
                              resolver 相同，并做 containment 校验

碰撞与复用规则（禁止静默合并不同分组）：
  - A/B  -> A_B
  - A\\B -> A_B_2
  - A:B  -> A_B_3
  - 同一 normalized source value（NFKC + strip）再次出现 -> 复用原目录（绝不变成 _4）
  - 后缀 _2/_3 按数据稳定顺序（rows 顺序）分配，不随机

空值规则：
  - None / "" / "   " 统一进入 EMPTY_GROUP_FOLDER = "未分类"
  - 真实 Excel 值本身是「未分类」时，第一版允许与空值共用该目录
"""

import os
import re
import unicodedata
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Sequence, Tuple

# 空值默认分组目录（前面改造方案约定的统一名称）
EMPTY_GROUP_FOLDER = "未分类"


class OutputPathError(Exception):
    """输出路径异常。

    code 取值：
      INVALID_GROUP_COLUMN 分组功能已启用但分组列无效（None / 越界）
      UNSAFE_PATH          解析出的路径逃逸了基础输出目录
    """

    def __init__(self, message: str, code: str = "GENERIC"):
        super().__init__(message)
        self.code = code


def _col_letter(index: int) -> str:
    """0 起列索引 -> Excel 列字母（0->A, 25->Z, 26->AA）。

    与 core.excel_data.index_to_excel_column 等价；此处独立实现以保持
    output_paths 与 excel_data 零耦合（仅用于错误提示文案）。
    """
    if index < 0:
        return "?"
    n = index
    out = []
    while True:
        out.append(chr(ord("A") + n % 26))
        n = n // 26 - 1
        if n < 0:
            break
    return "".join(reversed(out))


# ---------------------------------------------------------------------------
# 目录名清洗
# ---------------------------------------------------------------------------
# 路径分隔符 / : \\ 在目录名里不合法，统一替换为下划线（A/B -> A_B），
# 其余 Windows 非法字符（< > " | ? * 与控制字符）删除。
_PATH_SEP_RE = re.compile(r"[/\\:]")
_OTHER_INVALID_RE = re.compile(r'[<>"|?*\x00-\x1f]')
# Windows 保留设备名（不含扩展名比较；与 core/util.py 一致）
_RESERVED_NAMES = {
    "CON", "PRN", "AUX", "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}
_MAX_GROUP_DIR_LEN = 120

# 空值 key 哨兵（真实值经 NFKC+strip 后不可能等于它）
_EMPTY = object()


def sanitize_group_component(value: Any, fallback: str = EMPTY_GROUP_FOLDER) -> str:
    """把单个分组值清洗为安全目录片段（不含碰撞后缀）。

    - None / 空串 / 纯空白 -> fallback（默认「未分类」）；
    - Unicode NFKC 归一化（全角 -> 半角，如 Ａ／Ｂ -> A/B -> A_B）；
    - / : \\ 替换为 _（A/B -> A_B，禁止路径分隔符产生层级）；
    - 其余非法字符 < > " | ? * 与控制字符删除；
    - 去除首尾空白与结尾的 `.` / 空格（Windows 目录名限制）；
    - 保留设备名（CON / PRN / COM1 ...）加 `_` 前缀；
    - 清洗结果为空 / 仅为 `.` / `..` -> fallback；
    - 超长截断。
    """
    if value is None:
        return fallback
    s = unicodedata.normalize("NFKC", str(value))
    s = s.strip()
    if not s:
        return fallback
    # 路径分隔符/冒号 -> 下划线；其余非法字符删除
    s = _PATH_SEP_RE.sub("_", s)
    s = _OTHER_INVALID_RE.sub("", s)
    # 去结尾点/空格（Windows 不允许目录名以点或空格结尾）
    while s.endswith((".", " ")):
        s = s[:-1]
    # 设备名保护：保留名加下划线前缀，避免 Windows 拒绝
    stem = s.upper()
    if stem in _RESERVED_NAMES:
        s = "_" + s
    # '.' / '..' 防御（清洗后正常不可能出现，兜底）
    if s in (".", ".."):
        return fallback
    # 截断
    if len(s) > _MAX_GROUP_DIR_LEN:
        s = s[:_MAX_GROUP_DIR_LEN]
    return s or fallback


def _value_key(value: Any):
    """规范化分组值 -> 映射 key（NFKC + strip；空值统一哨兵）。

    同一 normalized source value（如两次出现的 'A/B'）必须映射同一目录；
    'Ａ' 与 'A'（NFKC 后相同）也视为同一源值。
    """
    if value is None:
        return _EMPTY
    s = unicodedata.normalize("NFKC", str(value)).strip()
    return s or _EMPTY


def _key_display(key: Any, fallback: str) -> str:
    return fallback if key is _EMPTY else str(key)


# ---------------------------------------------------------------------------
# GroupFolderMap / build_group_folder_map
# ---------------------------------------------------------------------------
@dataclass
class GroupFolderMap:
    """批次稳定的「分组值 -> 目录名」映射（只读视图）。

    - key_to_dir            : 规范化源值 key -> 最终目录名（不含路径分隔符）
    - dir_keys              : 目录名 -> 其下源值 key 列表（用于碰撞统计；fallback 组不计）
    - distinct_folder_count : 预计创建的目录数（含「未分类」，去重后）
    - empty_group_count     : 空值行数（None / "" / 空白，全部进入 fallback 目录）
    - collision_count       : 清洗后同名的冲突组数（如 A/B 与 A\\B -> 1 组）
    - collision_groups      : [(目录名, [源值, ...]), ...]（仅冲突组，按稳定顺序）
    """

    fallback: str = EMPTY_GROUP_FOLDER
    key_to_dir: Dict[Any, str] = field(default_factory=dict)
    dir_keys: Dict[str, List[Any]] = field(default_factory=dict)
    distinct_folder_count: int = 0
    empty_group_count: int = 0
    collision_count: int = 0
    collision_groups: List[Tuple[str, List[Any]]] = field(default_factory=list)

    def subdir_for(self, row, group_column: Optional[int]) -> str:
        """返回该行在分组配置下的目录名。

        - 未启用（group_column None）或该行 values 越界 -> fallback；
        - 否则按规范化源值查表（映射在整批中稳定，不会因行序变化而变化）。
        """
        if row is None or group_column is None:
            return self.fallback
        if not (0 <= group_column < len(row.values)):
            return self.fallback
        return self.key_to_dir.get(_value_key(row.values[group_column]), self.fallback)

    def collision_summary(self) -> List[str]:
        """供 GUI / CLI 日志展示的冲突明细行。"""
        out = []
        for base, keys in self.collision_groups:
            vals = [_key_display(k, self.fallback) for k in keys]
            out.append(f"{base} ← {' / '.join(vals)}")
        return out


def build_group_folder_map(rows: Sequence, group_column: Optional[int],
                           fallback: str = EMPTY_GROUP_FOLDER) -> GroupFolderMap:
    """批次开始前一次性建立完整 folder map（GUI / CLI 共用）。

    - 稳定顺序：按 rows 顺序首次出现分配，后缀 _2/_3 不随机；
    - 同一源值复用同一目录；不同源值清洗后同名 -> 碰撞后缀；
    - 空值（None / "" / 空白）统一进入 fallback 目录；真实值清洗后恰为
      fallback（如本身就是「未分类」）也与之共用（允许）；
    - group_column 为 None，或没有任何一行 values 覆盖该列 ->
      抛 OutputPathError(INVALID_GROUP_COLUMN)（配置错误，不静默 fallback）。
    """
    rows = list(rows or [])
    if group_column is None:
        raise OutputPathError("分组功能已启用，但未指定分组列。", code="INVALID_GROUP_COLUMN")
    if not rows or not any(0 <= group_column < len(r.values) for r in rows):
        raise OutputPathError(
            f"分组字段 {_col_letter(group_column)} 已超出当前数据的有效列范围。",
            code="INVALID_GROUP_COLUMN")

    key_to_dir: Dict[Any, str] = {}
    dir_keys: Dict[str, List[Any]] = {}
    base_keys: Dict[str, List[Any]] = {}   # 清洗后 base 名 -> 源值 key（碰撞检测）
    empty_count = 0
    fallback_used = False

    for r in rows:
        val = r.values[group_column] if 0 <= group_column < len(r.values) else None
        key = _value_key(val)
        if key is _EMPTY:
            empty_count += 1          # 按行计数（多次空值都计入）
        if key in key_to_dir:
            continue                  # 同一源值复用原目录（稳定）
        if key is _EMPTY:
            key_to_dir[key] = fallback
            fallback_used = True
            continue
        base = sanitize_group_component(val, fallback=fallback)
        if base == fallback:
            # 真实值清洗后恰为 fallback（如本身就是「未分类」）：允许与空值共用
            key_to_dir[key] = fallback
            fallback_used = True
            continue
        base_keys.setdefault(base, []).append(key)
        candidate = base
        n = 2
        while candidate in dir_keys:
            candidate = f"{base}_{n}"
            n += 1
        key_to_dir[key] = candidate
        dir_keys.setdefault(candidate, []).append(key)

    collision_groups = [(base, ks) for base, ks in base_keys.items() if len(ks) > 1]
    return GroupFolderMap(
        fallback=fallback,
        key_to_dir=key_to_dir,
        dir_keys=dir_keys,
        distinct_folder_count=len(dir_keys) + (1 if fallback_used else 0),
        empty_group_count=empty_count,
        collision_count=len(collision_groups),
        collision_groups=collision_groups,
    )


# ---------------------------------------------------------------------------
# 预检 + 统一目录解析入口 + containment 校验
# ---------------------------------------------------------------------------
def assert_group_column_valid(max_columns: int, enabled: bool,
                              group_column: Optional[int]) -> None:
    """Preflight：分组功能已启用时，分组列必须有效。

    - enabled=False -> 直接通过（不需要 group column）；
    - enabled=True  -> group_column 必须非 None 且 0 <= group_column < max_columns，
      否则抛 OutputPathError(INVALID_GROUP_COLUMN)。
    GUI 用它在 Preview / Start 前阻止；CLI / run_batch 在数据集加载后同样调用。
    """
    if not enabled:
        return
    if group_column is None:
        raise OutputPathError(
            "分组功能已启用，但未指定分组列。", code="INVALID_GROUP_COLUMN")
    if not (0 <= group_column < max_columns):
        raise OutputPathError(
            f"分组字段 {_col_letter(group_column)} 已超出当前 Excel 的有效列范围"
            f"（共 {max_columns} 列），请重新选择。",
            code="INVALID_GROUP_COLUMN")


def _check_containment(base_dir: str, resolved: str) -> None:
    """校验 resolved 绝对路径必须位于 base_dir 之下（base_dir 本身允许）。"""
    base = os.path.abspath(base_dir)
    target = os.path.abspath(resolved)
    try:
        common = os.path.commonpath([base, target])
    except ValueError:  # 不同盘符等
        common = ""
    if common != base:
        raise OutputPathError(
            f"输出路径逃逸了基础输出目录：{resolved!r} 不在 {base!r} 之下。",
            code="UNSAFE_PATH")


def resolve_output_directory(base_dir: str, row, group_column: Optional[int],
                             folder_map: Optional[GroupFolderMap] = None,
                             fallback: str = EMPTY_GROUP_FOLDER) -> str:
    """返回该行在 base_dir 下的最终输出目录（绝对路径），并确保已创建。

    这是 GUI Batch / GUI Preview / CLI **唯一**的目录解析入口：
      - 未启用（group_column None）-> 直接返回 base_dir（不建子目录）；
      - 启用 -> base_dir / map.subdir_for(row)；
      - 每次解析都执行 containment 校验（UNSAFE_PATH）；
      - folder_map 为 None 时按单行临时建立（Preview 单行场景兜底）。

    禁止任何调用方自行 os.path.join(out, sanitized_value)。
    """
    if group_column is None:
        os.makedirs(base_dir, exist_ok=True)
        return base_dir
    if folder_map is None:
        folder_map = build_group_folder_map([row], group_column, fallback=fallback)
    sub = folder_map.subdir_for(row, group_column)
    resolved = os.path.join(base_dir, sub)
    _check_containment(base_dir, resolved)
    os.makedirs(resolved, exist_ok=True)
    return resolved
