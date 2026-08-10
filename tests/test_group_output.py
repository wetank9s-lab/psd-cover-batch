# -*- coding: utf-8 -*-
"""Stage 4.5：按 Excel 任意列分组输出子文件夹 —— 稳定性回归测试（BLOCKED 验收 16 项）。

红线：分组值必须来自 ExcelRow.values[group_column]（Stage 4 统一管线），
测试同样只走 load_excel_dataset / core.output_paths，绝不直接 openpyxl 读分组值。
core/output_paths.py 本身禁止 import openpyxl。

覆盖用户要求的 40 个关键回归场景（node id 见各测试注释中的 #scenario N）。
"""
import json
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import pytest  # noqa: E402

from core.excel_data import load_excel_dataset  # noqa: E402
from core.output_paths import (  # noqa: E402
    EMPTY_GROUP_FOLDER, OutputPathError, GroupFolderMap,
    sanitize_group_component, build_group_folder_map,
    resolve_output_directory, assert_group_column_valid,
)
import openpyxl  # noqa: E402  (仅测试 fixture 用)


def _make_xlsx(tmp_path, rows, name="data.xlsx"):
    p = tmp_path / name
    wb = openpyxl.Workbook()
    ws = wb.active
    for row in rows:
        ws.append(row)
    wb.save(p)
    return str(p)


def _ds(tmp_path, rows, col_store=0, col_name=1, col_phone=2):
    p = _make_xlsx(tmp_path, rows)
    return load_excel_dataset(p, col_store=col_store, col_name=col_name,
                              col_phone=col_phone)


# ---------------------------------------------------------------------------
# sanitize_group_component：目录名清洗（含 NFKC / 前后空格 / 设备名 / 点）
# ---------------------------------------------------------------------------
def test_sanitize_normal():  # #scenario 2 普通 A门店
    assert sanitize_group_component("A门店") == "A门店"
    assert sanitize_group_component(" 康乐  ") == "康乐"


def test_sanitize_path_seps():  # #scenario 9 A/B -> A_B
    assert sanitize_group_component("A/B") == "A_B"
    assert sanitize_group_component("A\\B") == "A_B"
    assert sanitize_group_component("A:B") == "A_B"


def test_sanitize_nfkc():  # #scenario 7 Unicode NFKC
    # 全角 Ａ／Ｂ NFKC -> A/B -> A_B；全角空格 → 半角空格 → strip
    assert sanitize_group_component("Ａ／Ｂ") == "A_B"
    assert sanitize_group_component("\u3000A门店\u3000") == "A门店"


def test_sanitize_reserved():  # #scenario 12/13/14 CON PRN COM1
    assert sanitize_group_component("CON") == "_CON"
    assert sanitize_group_component("PRN") == "_PRN"
    assert sanitize_group_component("COM1") == "_COM1"


def test_sanitize_trailing():  # #scenario 15 trailing dot / 16 trailing space
    assert sanitize_group_component("门店.") == "门店"
    assert sanitize_group_component("门店 ") == "门店"


def test_sanitize_dot_dot():  # #scenario 17/18 "." / ".."
    assert sanitize_group_component(".") == EMPTY_GROUP_FOLDER
    assert sanitize_group_component("..") == EMPTY_GROUP_FOLDER


def test_sanitize_evil_paths():  # #scenario 19-22 绝对/UNC 路径的目录片段清洗
    # 这些值如果直接 os.path.join 会逃逸；本清洗把分隔符/冒号转 _，只产生单层目录名
    for evil in ["../evil", "..\\evil", "C:\\Temp", "\\server\\share"]:
        comp = sanitize_group_component(evil)
        assert os.sep not in comp and "/" not in comp and "\\" not in comp
        assert ":" not in comp


def test_sanitize_empty():  # #scenario 8 空值 -> 未分类
    assert sanitize_group_component(None) == "未分类"
    assert sanitize_group_component("") == "未分类"
    assert sanitize_group_component("   ") == "未分类"
    # 全分隔符值：/ : \ 都被替换为 _（保留信息、不产生层级），非空
    assert sanitize_group_component("///") == "___"


# ---------------------------------------------------------------------------
# GroupFolderMap / build_group_folder_map：碰撞、复用、统计
# ---------------------------------------------------------------------------
class _Row:
    """最小 row 替身（鸭子类型：只需 .values）。"""
    def __init__(self, values):
        self.values = tuple(values)


def test_map_no_collision():  # #scenario 2/4 普通 A门店、任意非 store 列
    rows = [_Row(["A门店", "上海"]), _Row(["A门店", "上海"]), _Row(["B门店", "杭州"])]
    m = build_group_folder_map(rows, 0)
    assert [m.subdir_for(r, 0) for r in rows] == ["A门店", "A门店", "B门店"]
    assert m.distinct_folder_count == 2
    assert m.collision_count == 0
    m2 = build_group_folder_map(rows, 1)  # 任意列（城市）
    assert [m2.subdir_for(r, 1) for r in rows] == ["上海", "上海", "杭州"]


def test_map_same_value_reuse():  # #scenario 27 同原始值再次出现复用原目录
    rows = [_Row(["A/B", "x"]), _Row(["A\\B", "y"]), _Row(["A/B", "z"])]
    m = build_group_folder_map(rows, 0)
    dirs = [m.subdir_for(r, 0) for r in rows]
    # 第1行 A/B -> A_B；第2行 A\B 冲突 -> A_B_2；第3行 A/B 复用 A_B（绝不变成 A_B_3/A_B_4）
    assert dirs == ["A_B", "A_B_2", "A_B"]


def test_map_three_collision():  # #scenario 25/26 collision -> _2、三冲突 -> _3
    rows = [_Row(["A/B"]), _Row(["A\\B"]), _Row(["A:B"])]
    m = build_group_folder_map(rows, 0)
    dirs = [m.subdir_for(r, 0) for r in rows]
    assert dirs == ["A_B", "A_B_2", "A_B_3"]
    assert m.collision_count == 1           # 1 组冲突（三个源值）
    assert len(m.collision_groups) == 1
    assert m.collision_groups[0][0] == "A_B"


def test_map_stable_order():  # #scenario 24 A/B 与 A\B collision；后缀按数据顺序
    rows = [_Row(["A\\B"]), _Row(["A/B"])]
    m = build_group_folder_map(rows, 0)
    assert m.subdir_for(rows[0], 0) == "A_B"     # 先出现者无后缀
    assert m.subdir_for(rows[1], 0) == "A_B_2"


def test_map_empty_counts():  # #scenario 8 空值统计
    rows = [_Row(["", "x"]), _Row([None, "y"]), _Row(["   ", "z"]),
            _Row(["A门店", "w"])]
    m = build_group_folder_map(rows, 0)
    assert [m.subdir_for(r, 0) for r in rows] == ["未分类", "未分类", "未分类", "A门店"]
    assert m.empty_group_count == 3       # 按行计数
    assert m.distinct_folder_count == 2   # 未分类 + A门店


def test_map_real_value_equals_fallback():  # 真实值恰为「未分类」允许共用
    rows = [_Row([""]), _Row(["未分类"]), _Row(["A门店"])]
    m = build_group_folder_map(rows, 0)
    assert [m.subdir_for(r, 0) for r in rows] == ["未分类", "未分类", "A门店"]


def test_map_invalid_column_raises():  # #scenario 28/29 enabled + None/越界 -> error
    with pytest.raises(OutputPathError) as ei:
        build_group_folder_map([_Row(["a", "b"])], None)
    assert ei.value.code == "INVALID_GROUP_COLUMN"
    with pytest.raises(OutputPathError) as ei2:
        build_group_folder_map([_Row(["a", "b"])], 5)
    assert ei2.value.code == "INVALID_GROUP_COLUMN"
    # 空 rows 也算无效
    with pytest.raises(OutputPathError):
        build_group_folder_map([], 0)


# ---------------------------------------------------------------------------
# containment：任何 Excel 分组值都不能逃离 base output dir
# ---------------------------------------------------------------------------
def test_containment_evil_values(tmp_path):  # #scenario 19-23 containment
    base = str(tmp_path / "base")
    os.makedirs(base, exist_ok=True)
    for evil in ["../evil", "..\\evil", "C:\\Temp", "\\server\\share", ".", ".."]:
        rows = [_Row([evil])]
        m = build_group_folder_map(rows, 0)
        d = m.subdir_for(rows[0], 0)
        assert d == "未分类" or (os.sep not in d and "/" not in d and "\\" not in d)
        resolved = resolve_output_directory(base, rows[0], 0, folder_map=m)
        # 必须位于 base 之下（commonpath 相等或更深）
        common = os.path.commonpath([os.path.abspath(base), os.path.abspath(resolved)])
        assert common == os.path.abspath(base), f"逃逸: {evil} -> {resolved}"


def test_containment_direct_check():  # #scenario 23 直接校验
    rows = [_Row(["../evil"])]
    m = build_group_folder_map(rows, 0)
    # ../evil 清洗后为 _.._evil（不含分隔符），仍是 base 下的单层目录
    sub = m.subdir_for(rows[0], 0)
    assert "/" not in sub and "\\" not in sub and ":" not in sub


# ---------------------------------------------------------------------------
# resolve_output_directory：统一入口、禁用直返 base、containment、建目录
# ---------------------------------------------------------------------------
def test_resolve_disabled_returns_base(tmp_path):  # #scenario 1 disabled -> base dir
    base = str(tmp_path / "out")
    r = _Row(["A门店", "x"])
    assert resolve_output_directory(base, r, None) == base
    assert os.path.isdir(base)


def test_resolve_grouped_creates_subdir(tmp_path):  # #scenario 2/3 同值多行共用目录
    base = str(tmp_path / "out")
    rows = [_Row(["A门店", "x"]), _Row(["A门店", "y"])]
    m = build_group_folder_map(rows, 0)
    d1 = resolve_output_directory(base, rows[0], 0, folder_map=m)
    d2 = resolve_output_directory(base, rows[1], 0, folder_map=m)
    assert d1 == d2 == os.path.join(base, "A门店")
    assert os.path.isdir(d1)


def test_resolve_unsafe_path_raises(tmp_path):  # #scenario 23 UNSAFE_PATH
    # 直接构造逃逸子目录（绕过 map，验证 resolver 的 containment 兜底）
    from core.output_paths import _check_containment
    base = str(tmp_path / "out")
    with pytest.raises(OutputPathError) as ei:
        _check_containment(base, os.path.join(base, "..", "evil"))
    assert ei.value.code == "UNSAFE_PATH"


# ---------------------------------------------------------------------------
# 多格式同目录：同一 ExcelRow 的所有 enabled formats 共用同一 row_output_dir
# ---------------------------------------------------------------------------
def test_multiformat_same_row_dir(tmp_path):  # #scenario 33/34/35 PNG JPG PSD+PNG 同目录
    base = str(tmp_path / "out")
    rows = [_Row(["A门店", "x"]), _Row(["B门店", "y"])]
    m = build_group_folder_map(rows, 0)

    def emit(r, fmt, also_png=False):
        row_dir = resolve_output_directory(base, r, 0, folder_map=m)
        files = [f"001_{fmt.lower()}.{fmt.lower()}"]
        if also_png:
            files.append("001_also.png")
        for fn in files:
            with open(os.path.join(row_dir, fn), "w") as f:
                f.write("x")
        return row_dir

    a_dir = emit(rows[0], "PNG")
    a_dir2 = emit(rows[0], "JPG")
    a_dir3 = emit(rows[0], "PSD", also_png=True)
    b_dir = emit(rows[1], "PNG")
    # 同一行所有格式 -> 同一分组目录
    assert a_dir == a_dir2 == a_dir3 == os.path.join(base, "A门店")
    assert b_dir == os.path.join(base, "B门店")
    assert sorted(os.listdir(a_dir)) == ["001_also.png", "001_jpg.jpg", "001_png.png", "001_psd.psd"]


# ---------------------------------------------------------------------------
# Preflight：assert_group_column_valid
# ---------------------------------------------------------------------------
def test_preflight_disabled_ok():  # #scenario 1 disabled 不校验列
    assert_group_column_valid(3, False, None)
    assert_group_column_valid(3, False, 99)   # 未启用时列无效也无所谓


def test_preflight_enabled_requires_valid():  # #scenario 28/29/39
    with pytest.raises(OutputPathError) as ei:
        assert_group_column_valid(5, True, None)
    assert ei.value.code == "INVALID_GROUP_COLUMN"
    with pytest.raises(OutputPathError) as ei2:
        assert_group_column_valid(5, True, 26)   # 旧配置 AA 越界（新 Excel 只有 5 列）
    assert ei2.value.code == "INVALID_GROUP_COLUMN"
    assert "AA" in str(ei2.value)                 # 提示列字母
    assert_group_column_valid(30, True, 26)       # 有效（AA < 30 列）
    assert_group_column_valid(5, True, 4)


# ---------------------------------------------------------------------------
# 不重读 Excel：同一 ExcelDataset 切换分组列 -> 不同目录，无 Excel IO
# ---------------------------------------------------------------------------
def test_no_excel_reload_on_group_column_change(tmp_path):  # #scenario 30/4
    ds = _ds(tmp_path, [
        ["门店", "姓名", "电话", "城市"],
        ["A门店", "张三", 13800000001, "上海"],
        ["B门店", "李四", 13800000002, "杭州"],
    ])
    # group=门店 -> A门店/B门店
    m1 = build_group_folder_map(ds.valid_rows, 0)
    # group=城市 -> 上海/杭州（不调用 load_excel_dataset，纯 map 计算）
    m2 = build_group_folder_map(ds.valid_rows, 3)
    assert [m1.subdir_for(r, 0) for r in ds.valid_rows] == ["A门店", "B门店"]
    assert [m2.subdir_for(r, 3) for r in ds.valid_rows] == ["上海", "杭州"]


def test_grouping_core_no_openpyxl_import():
    """grouping core 禁止 import openpyxl（规格 #9/#15）。"""
    import re
    import core.output_paths as op
    with open(op.__file__, encoding="utf-8") as f:
        src = f.read()
    # 检测 import 语句（docstring 提到该禁令字样不影响）
    imports = re.findall(r"^\s*(?:import|from)\s+([\w.]+)", src, re.M)
    assert not any("openpyxl" in m for m in imports)
    assert "Workbook" not in imports and "Worksheet" not in imports


# ---------------------------------------------------------------------------
# Preview 隔离：Preview 走 out_dir/_preview，与 Batch 同一 resolver
# ---------------------------------------------------------------------------
def test_preview_uses_same_resolver_and_preview_base(tmp_path):  # #scenario 31/32
    out = str(tmp_path / "out")
    preview_base = os.path.join(out, "_preview")
    ds = _ds(tmp_path, [
        ["门店", "姓名", "电话"],
        ["A门店", "张三", 13800000001],
        ["B门店", "王五", 13800000003],
    ])
    m = build_group_folder_map(ds.valid_rows, 0)
    # Preview：base=out_dir/_preview（同一 resolver，仅 base 不同）
    preview_dir = resolve_output_directory(preview_base, ds.valid_rows[0], 0, folder_map=m)
    assert preview_dir == os.path.join(preview_base, "A门店")
    assert os.path.isdir(preview_dir)
    # 预览只写 _preview 子目录；正式 out_dir 下不产生 A门店（未做 Batch）
    assert not os.path.exists(os.path.join(out, "A门店"))
    # Batch 后再验证正式目录（用同一 resolver + 同一 map）与预览目录分开
    batch_dir = resolve_output_directory(out, ds.valid_rows[0], 0, folder_map=m)
    assert batch_dir == os.path.join(out, "A门店")
    assert batch_dir != preview_dir


def test_preview_name_does_not_pollute_batch(tmp_path):  # #scenario 32
    out = str(tmp_path / "out")
    preview_base = os.path.join(out, "_preview")
    ds = _ds(tmp_path, [
        ["门店", "姓名", "电话"],
        ["A门店", "张三", 13800000001],
    ])
    m = build_group_folder_map(ds.valid_rows, 0)
    # 预览文件真实写入 preview 子目录
    pdir = resolve_output_directory(preview_base, ds.valid_rows[0], 0, folder_map=m)
    with open(os.path.join(pdir, "preview_张三.png"), "w") as f:
        f.write("x")
    # 正式目录不得出现任何文件
    assert not os.path.exists(os.path.join(out, "A门店"))
    assert sorted(os.listdir(pdir)) == ["preview_张三.png"]


def test_preview_and_batch_share_collision_folder_mapping(tmp_path):
    """CONDITIONAL PASS 边界核验：Preview 与 Batch 在 collision 场景共享同一 folder mapping。

    构造 dataset（group value 发生清洗碰撞）：
      第1有效行 group value = A/B
      第2有效行 group value = A\\B
      第3有效行 group value = A:B
    完整 batch folder map 必须：
      A/B -> A_B
      A\\B -> A_B_2
      A:B -> A_B_3

    Preview 与 Batch 必须都基于完整 dataset 建 map（build_group_folder_map(valid_rows)），
    不能只针对 preview row 单独建 map（否则碰撞后缀不一致）。

    验证：
      - Batch row2 subdir   == A_B_2
      - Preview row2 subdir == A_B_2
      - Preview 完整路径位于 _preview/A_B_2
      - Preview row3（三冲突）位于 _preview/A_B_3
      - Preview 不污染正式目录
    """
    out = str(tmp_path / "out")
    preview_base = os.path.join(out, "_preview")
    # 真实 Excel（4 列：门店/姓名/电话/分组列）——分组值 A/B、A\B、A:B
    ds = _ds(tmp_path, [
        ["门店", "姓名", "电话", "分组"],
        ["M1", "张三", 13800000001, "A/B"],
        ["M2", "李四", 13800000002, "A\\B"],
        ["M3", "王五", 13800000003, "A:B"],
    ])
    group_col = 3

    # —— 完整 batch folder map（GUI run_batch / CLI run 都这样建）——
    folder_map = build_group_folder_map(ds.valid_rows, group_col)
    assert [folder_map.subdir_for(r, group_col) for r in ds.valid_rows] == \
        ["A_B", "A_B_2", "A_B_3"]

    # —— Batch（base=out_dir）：row2 -> out/A_B_2 ——
    batch_row2 = resolve_output_directory(out, ds.valid_rows[1], group_col, folder_map=folder_map)
    assert batch_row2 == os.path.join(out, "A_B_2")
    # 与「只按单行建 map」的歧义实现对比：单行 map 会给 row2 无后缀的 A_B（证明差异存在）
    single_map = build_group_folder_map([ds.valid_rows[1]], group_col)
    assert single_map.subdir_for(ds.valid_rows[1], group_col) == "A_B"

    # —— Preview（base=out_dir/_preview，同一 folder_map）：row2 -> _preview/A_B_2 ——
    preview_row2 = resolve_output_directory(preview_base, ds.valid_rows[1], group_col,
                                            folder_map=folder_map)
    assert preview_row2 == os.path.join(preview_base, "A_B_2")
    # 同一 ExcelRow 的 subdir 在 Batch 与 Preview 必须完全一致
    assert os.path.basename(preview_row2) == os.path.basename(batch_row2) == "A_B_2"

    # —— 三冲突：Preview row3 -> _preview/A_B_3 ——
    preview_row3 = resolve_output_directory(preview_base, ds.valid_rows[2], group_col,
                                            folder_map=folder_map)
    assert preview_row3 == os.path.join(preview_base, "A_B_3")

    # —— 预览不污染正式目录（out/A_B_2 只由 Batch 创建，无 preview 文件）——
    assert os.path.isdir(batch_row2)            # Batch 目录存在
    assert not os.path.exists(os.path.join(batch_row2, "preview.png"))
    # 正式 out_dir 下不存在 A_B（碰撞行绝不落入无后缀目录）
    assert not os.path.exists(os.path.join(out, "A_B"))


# ---------------------------------------------------------------------------
# 配置存取：old config 缺 group 字段 -> disabled；new config roundtrip；0-based
# ---------------------------------------------------------------------------
class _V:
    """轻量 Var 替身：get/set + value 语义（避免 tk root 依赖）。"""
    def __init__(self, value=None):
        self._v = value
    def get(self):
        return self._v
    def set(self, v):
        self._v = v
    def is_set(self):
        return bool(self._v)


def _gui_min():
    """构造最小 App 实例（App.__new__，不跑 __init__），mock 所需控件（无 tk root）。"""
    import qifang_cover_maker as g
    app = g.App.__new__(g.App)
    app.col_store_var = _V("B")
    app.col_name_var = _V("C")
    app.col_phone_var = _V("D")
    app.col_role_var = _V("（不替换）")
    app.group_output_var = _V(True)
    app.group_col_var = _V("E")
    app.header_var = _V(True)
    app.psd_var = _V("t.psd")
    app.xlsx_var = _V("t.xlsx")
    app.out_var = _V("out")
    app.fmt_var = _V("png")
    app.also_png_var = _V(False)
    app.excel_dataset = None
    app._column_labels = ["A", "B", "C", "D", "E"]
    app.col_store_cb = app.col_name_cb = app.col_phone_cb = app.col_role_cb = None
    app.group_col_cb = None
    app.logo_checks = {}
    app.map_combos = {}
    app.brand_checks = {}
    app.layer_index = None
    app.all_psd_layers = []
    app.tm_name_var = _V("（不替换）")
    app.tm_phone_var = _V("（不替换）")
    app.tm_role_var = _V("（不替换）")
    return app


def test_config_old_without_group_fields(tmp_path):  # #scenario 36 old config -> disabled
    import qifang_cover_maker as g
    app = _gui_min()
    cfg = {
        "psd_path": "", "xlsx_path": "", "out_dir": "",
        "has_header": True,
        "col_store": 1, "col_name": 2, "col_phone": 3, "col_role": -1,
        "text_map": {}, "fmt": "png", "also_png": False,
        # 没有 group_output_enabled / group_output_column（旧配置）
    }
    app.config_path = str(tmp_path / "cfg_old.json")
    with open(app.config_path, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False)
    app._load_config()
    assert app.group_output_var.get() == False
    assert app.group_output_var.get() is not True


def test_config_roundtrip(tmp_path):  # #scenario 37/38 config roundtrip, 0-based
    import qifang_cover_maker as g
    app = _gui_min()
    cfg = {
        "psd_path": "", "xlsx_path": "", "out_dir": "",
        "has_header": True,
        "col_store": 1, "col_name": 2, "col_phone": 3, "col_role": -1,
        "group_output_enabled": True, "group_output_column": 0,
        "text_map": {}, "fmt": "png", "also_png": False,
    }
    app.config_path = str(tmp_path / "cfg_new.json")
    with open(app.config_path, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False)
    app._load_config()
    assert app.group_output_var.get() == True
    assert app.group_col_var.get() == "A"          # 0 -> A
    # save/reload roundtrip：重新 _collect_cfg 仍为 True / 0（0-based，不存 "A"）
    cfg2 = app._collect_cfg()
    assert cfg2["group_output_enabled"] == True
    assert cfg2["group_output_column"] == 0


def test_config_aa_26():  # #scenario 5/38 AA -> 26（0-based）
    import qifang_cover_maker as g
    app = _gui_min()
    app.group_col_var = _V("AA")
    cfg = app._collect_cfg()
    assert cfg["group_output_column"] == 26
    assert cfg["group_output_column"] != "AA"      # 不保存列字母
    assert cfg["group_output_column"] != "A - 门店"


def test_config_old_column_letter_bad_data_safe(tmp_path):
    """旧版坏数据（group_output_column 存成列字母 "AA"）load 时不崩溃、按 disabled 兜底。"""
    import qifang_cover_maker as g
    app = _gui_min()
    cfg = {"group_output_enabled": True, "group_output_column": "AA"}
    app.config_path = str(tmp_path / "_cfg_bad.json")
    with open(app.config_path, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False)
    # _load_config 内部 except Exception 兜底，不会抛；group_col_var 保持默认
    app._load_config()
    assert app.group_col_var.get() in ("A", "E")   # 未被 "AA" 污染


def test_gui_collect_cfg_group_output_index():  # #scenario 38 E -> 4
    app = _gui_min()
    cfg = app._collect_cfg()
    assert cfg["group_output_enabled"] == True
    assert cfg["group_output_column"] == 4   # E 列 = 0-based 4


def test_gui_collect_cfg_group_disabled():  # disabled -> column None
    app = _gui_min()
    app.group_output_var.set(False)
    cfg = app._collect_cfg()
    assert cfg["group_output_enabled"] == False
    assert cfg["group_output_column"] is None


def test_gui_load_config_restore_group(tmp_path):  # index 4 -> "E"
    import qifang_cover_maker as g
    app = _gui_min()
    cfg = {
        "psd_path": "", "xlsx_path": "", "out_dir": "",
        "has_header": True,
        "col_store": 1, "col_name": 2, "col_phone": 3, "col_role": -1,
        "group_output_enabled": True, "group_output_column": 4,
        "text_map": {}, "fmt": "png", "also_png": False,
    }
    app.config_path = str(tmp_path / "cfg.json")
    with open(app.config_path, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False)
    app._load_config()
    assert app.group_output_var.get() == True
    assert app.group_col_var.get() == "E"


# ---------------------------------------------------------------------------
# 真实 Excel：任意列 / AA 列 / 空值（场景 4/5/6/8，走 load_excel_dataset）
# ---------------------------------------------------------------------------
def test_real_excel_store_and_city(tmp_path):  # #scenario 4/30 门店列 vs 城市列
    ds = _ds(tmp_path, [
        ["门店", "姓名", "电话", "城市"],
        ["A门店", "张三", 13800000001, "上海"],
        ["B门店", "李四", 13800000002, "杭州"],
    ])
    m_store = build_group_folder_map(ds.valid_rows, 0)
    m_city = build_group_folder_map(ds.valid_rows, 3)
    assert [m_store.subdir_for(r, 0) for r in ds.valid_rows] == ["A门店", "B门店"]
    assert [m_city.subdir_for(r, 3) for r in ds.valid_rows] == ["上海", "杭州"]


def test_real_excel_aa_column(tmp_path):  # #scenario 5 AA 列（26）
    header = [f"C{i}" for i in range(27)]
    header[26] = "区域"
    rows = [header]
    for i, area in enumerate(["东区", "东区", "西区"], start=1):
        data = [f"x{i}-{j}" for j in range(27)]
        data[0] = f"门店{i}"
        data[1] = f"姓名{i}"
        data[3] = 13800000000 + i
        data[26] = area
        rows.append(data)
    ds = _ds(tmp_path, rows, col_store=0, col_name=1, col_phone=3)
    m = build_group_folder_map(ds.valid_rows, 26)
    assert [m.subdir_for(r, 26) for r in ds.valid_rows] == ["东区", "东区", "西区"]


def test_real_excel_trim_and_empty(tmp_path):  # #scenario 6/8 前后空格 / 空值
    ds = _ds(tmp_path, [
        ["门店", "姓名", "电话", "城市"],
        ["  A门店  ", "张三", 13800000001, ""],
        ["B门店", "李四", 13800000002, "  杭州  "],
    ])
    m = build_group_folder_map(ds.valid_rows, 0)
    assert [m.subdir_for(r, 0) for r in ds.valid_rows] == ["A门店", "B门店"]
    m2 = build_group_folder_map(ds.valid_rows, 3)
    assert [m2.subdir_for(r, 3) for r in ds.valid_rows] == ["未分类", "杭州"]


def test_gui_min_fixture_has_group_col_cb():  # 防挂起回归：_set_column_options 引用 group_col_cb
    app = _gui_min()
    assert app.group_col_cb is None
    # _set_column_options 在 cb=None 时跳过（不 AttributeError）
    app.excel_dataset = None
    app._set_column_options(5)


# ---------------------------------------------------------------------------
# Stage 7.5：PSD 指纹 / 变更检测 / force_reload 快照
# ---------------------------------------------------------------------------
def test_psd_fingerprint_none_when_missing(tmp_path):
    import qifang_cover_maker as g
    app = _gui_min()
    assert g.App._psd_fingerprint_of(str(tmp_path / "no.psd")) is None


def test_psd_fingerprint_records_mtime_size(tmp_path):
    import qifang_cover_maker as g
    p = tmp_path / "t.psd"
    p.write_bytes(b"PSD-BYTES-1")
    fp = g.App._psd_fingerprint_of(str(p))
    assert fp is not None and len(fp) == 2
    mtime, size = fp
    assert size == len(b"PSD-BYTES-1")
    # 修改后指纹变化
    p.write_bytes(b"PSD-BYTES-2-LONGER")
    fp2 = g.App._psd_fingerprint_of(str(p))
    assert fp2 != fp


def test_psd_changed_since_load_false_when_no_fingerprint():
    import qifang_cover_maker as g
    app = _gui_min()
    # 未加载过（无指纹）-> 不认为变更
    assert app._psd_changed_since_load() is False


def test_psd_changed_since_load_true_after_external_modify(tmp_path):
    import qifang_cover_maker as g
    app = _gui_min()
    p = tmp_path / "t.psd"
    p.write_bytes(b"V1")
    app.psd_var.set(str(p))
    # 模拟成功 Load 后记录指纹
    app._psd_fingerprint = (str(p), g.App._psd_fingerprint_of(str(p)))
    assert app._psd_changed_since_load() is False
    # 外部修改（新增图层 -> 文件变大 / mtime 变化）
    import time as _t
    p.write_bytes(b"V2-LONGER-NEW-LAYERS")
    # mtime 粒度问题：size 变化已足够触发
    assert app._psd_changed_since_load() is True


def test_psd_changed_since_load_false_when_path_switched(tmp_path):
    import qifang_cover_maker as g
    app = _gui_min()
    p1 = tmp_path / "a.psd"
    p2 = tmp_path / "b.psd"
    p1.write_bytes(b"V1")
    p2.write_bytes(b"V1")
    app.psd_var.set(str(p1))
    app._psd_fingerprint = (str(p1), g.App._psd_fingerprint_of(str(p1)))
    # 用户切换了 PSD 路径 -> 指纹不适用（视为未变更，等重新加载）
    app.psd_var.set(str(p2))
    assert app._psd_changed_since_load() is False


def test_snapshot_for_load_force_reload_flag(tmp_path):
    import qifang_cover_maker as g
    app = _gui_min()
    snap = app._snapshot_for_load(force_reload=True)
    assert snap["force_reload"] is True
    snap2 = app._snapshot_for_load(force_reload=False)
    assert snap2["force_reload"] is False
    snap3 = app._snapshot_for_load()
    assert snap3["force_reload"] is False
