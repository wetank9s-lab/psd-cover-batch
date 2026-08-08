# -*- coding: utf-8 -*-
"""Stage 4.5：按 Excel 任意列分组输出子文件夹 —— 核心纯函数测试。

红线：分组值必须来自 ExcelRow.values[group_column]（Stage 4 统一管线），
测试同样只走 load_excel_dataset / resolve_group_subdir，绝不直接 openpyxl 读分组值。
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import pytest  # noqa: E402

from core.excel_data import (  # noqa: E402
    load_excel_dataset, group_dir_name, resolve_group_subdir,
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


# ---------------------------------------------------------------------------
# group_dir_name：目录名清洗
# ---------------------------------------------------------------------------
def test_group_dir_name_normal():
    assert group_dir_name("A门店") == "A门店"
    assert group_dir_name(" 康乐  ") == "康乐"
    assert group_dir_name("张三/李四") == "张三李四"      # 移除 Windows 非法字符


def test_group_dir_name_empty_fallback():
    assert group_dir_name(None) == "未分组"
    assert group_dir_name("") == "未分组"
    assert group_dir_name("   ") == "未分组"
    assert group_dir_name("///") == "未分组"              # 清洗后为空


def test_group_dir_name_reserved_device():
    assert group_dir_name("CON") == "_CON"               # 保留设备名加下划线
    assert group_dir_name("NUL") == "_NUL"


def test_group_dir_name_custom_fallback():
    assert group_dir_name("", fallback="其他") == "其他"


# ---------------------------------------------------------------------------
# resolve_group_subdir：从 ExcelRow.values 取分组值（Stage 4.5 红线）
# ---------------------------------------------------------------------------
def test_resolve_group_subdir_store_column(tmp_path):
    """按门店列（col_store=B=1）分组：分组值直接来自 values[1]。"""
    p = _make_xlsx(tmp_path, [
        ["编号", "门店", "姓名", "电话"],
        [1, "A门店", "张三", 13800000001],
        [2, "A门店", "李四", 13800000002],
        [3, "B门店", "王五", 13800000003],
    ])
    ds = load_excel_dataset(p, col_store=1, col_name=2, col_phone=3)
    dirs = [resolve_group_subdir(r, 1) for r in ds.valid_rows]
    assert dirs == ["A门店", "A门店", "B门店"]


def test_resolve_group_subdir_arbitrary_column(tmp_path):
    """任意列（非门店列）也可分组：城市列（D=3）。"""
    p = _make_xlsx(tmp_path, [
        ["门店", "姓名", "电话", "城市"],
        ["易田", "张三", 13800000001, "成都"],
        ["康乐", "李四", 13800000002, "成都"],
        ["红旗", "王五", 13800000003, "重庆"],
    ])
    ds = load_excel_dataset(p, col_store=0, col_name=1, col_phone=2)
    dirs = [resolve_group_subdir(r, 3) for r in ds.valid_rows]
    assert dirs == ["成都", "成都", "重庆"]


def test_resolve_group_subdir_aa_column(tmp_path):
    """AA 列（26）作为分组列。"""
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
    p = _make_xlsx(tmp_path, rows)
    ds = load_excel_dataset(p, col_store=0, col_name=1, col_phone=3)
    dirs = [resolve_group_subdir(r, 26) for r in ds.valid_rows]
    assert dirs == ["东区", "东区", "西区"]


def test_resolve_group_subdir_empty_value_fallback(tmp_path):
    """分组值为空 -> 「未分组」。"""
    p = _make_xlsx(tmp_path, [
        ["门店", "姓名", "电话", "城市"],
        ["易田", "张三", 13800000001, ""],
        ["康乐", "李四", 13800000002, "成都"],
    ])
    ds = load_excel_dataset(p, col_store=0, col_name=1, col_phone=2)
    dirs = [resolve_group_subdir(r, 3) for r in ds.valid_rows]
    assert dirs == ["未分组", "成都"]


def test_resolve_group_subdir_disabled_or_out_of_range(tmp_path):
    """未启用（None）或列越界 -> fallback，不崩溃。"""
    p = _make_xlsx(tmp_path, [
        ["门店", "姓名", "电话"],
        ["易田", "张三", 13800000001],
    ])
    ds = load_excel_dataset(p, col_store=0, col_name=1, col_phone=2)
    r0 = ds.valid_rows[0]
    assert resolve_group_subdir(r0, None) == "未分组"
    assert resolve_group_subdir(r0, 99) == "未分组"     # 越界
    assert resolve_group_subdir(r0, 99, fallback="其他") == "其他"


# ---------------------------------------------------------------------------
# GUI / CLI 共用：输出路径 = os.path.join(out_dir, resolve_group_subdir(...))
# ---------------------------------------------------------------------------
def test_group_dir_path_build(tmp_path):
    """组装规则：out_dir / 分组子目录 / 文件 —— Preview/Batch/CLI 同一规则。"""
    p = _make_xlsx(tmp_path, [
        ["编号", "门店", "姓名", "电话"],
        [1, "A门店", "张三", 13800000001],
        [2, "B门店", "王五", 13800000003],
    ])
    ds = load_excel_dataset(p, col_store=1, col_name=2, col_phone=3)
    out = tmp_path / "out"
    paths = []
    for i, r in enumerate(ds.valid_rows, start=1):
        sub = resolve_group_subdir(r, 1)
        row_dir = os.path.join(str(out), sub)
        os.makedirs(row_dir, exist_ok=True)
        fp = os.path.join(row_dir, f"{i:03d}_x.png")
        with open(fp, "w") as f:      # 真实写入文件，验证目录归属
            f.write("x")
        paths.append(fp)
    assert os.path.exists(os.path.join(str(out), "A门店"))
    assert os.path.exists(os.path.join(str(out), "B门店"))
    assert sorted(os.listdir(os.path.join(str(out), "A门店"))) == ["001_x.png"]
    assert sorted(os.listdir(os.path.join(str(out), "B门店"))) == ["002_x.png"]


# ---------------------------------------------------------------------------
# GUI 配置存取：group_output_enabled / group_output_column 保存 0-based index
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
    # 下拉刷新（_set_column_options 用到）
    app.excel_dataset = None
    app._column_labels = ["A", "B", "C", "D", "E"]
    app.col_store_cb = app.col_name_cb = app.col_phone_cb = app.col_role_cb = None
    app.group_col_cb = None
    # _collect_cfg 依赖
    app.logo_checks = {}
    app.map_combos = {}
    app.brand_checks = {}
    app.layer_index = None
    app.all_psd_layers = []
    app.tm_name_var = _V("（不替换）")
    app.tm_phone_var = _V("（不替换）")
    app.tm_role_var = _V("（不替换）")
    return app


def test_gui_collect_cfg_group_output_index():
    """_collect_cfg 保存 group_output_enabled=True + group_output_column=E -> 4（0-based）。"""
    app = _gui_min()
    cfg = app._collect_cfg()
    assert cfg["group_output_enabled"] == True
    assert cfg["group_output_column"] == 4   # E 列 = 0-based 4


def test_gui_collect_cfg_group_disabled():
    """未勾选分组 -> group_output_column 为 None。"""
    app = _gui_min()
    app.group_output_var.set(False)
    cfg = app._collect_cfg()
    assert cfg["group_output_enabled"] == False
    assert cfg["group_output_column"] is None


def test_gui_load_config_restore_group(tmp_path):
    """_load_config 恢复 group_output_enabled / group_output_column（0-based index -> 列字母）。"""
    import json
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
