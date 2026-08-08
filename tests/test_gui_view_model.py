# -*- coding: utf-8 -*-
"""Stage 6.5B：gui_view_model 纯逻辑 ViewModel 单元测试（不依赖 Tk / worker）。

覆盖：
  - AppState -> 状态文本/bootstyle/进度模式 映射（规格 6.5B 第 6 节短文案）
  - 门店映射状态（mapping_status）
  - 搜索过滤（Logo 图层 / 门店 / 完整路径）
  - selected vs brand 展示（logo_display_state，规格 6.5B 第 13 节：禁 BRAND 术语）
  - 进度显示模型（progress_display）
  - BatchResult -> Summary 模型（batch_summary_model / summary_duration_text）
  - 配置检查模型（config_check_model）
  - LayerRef -> display label（layer_display_label，规格 6.5B 第 10 节：dict 不泄露）
  - 规格 6.5B 第 30 节：视觉纯逻辑回归（状态短文案 / summary headline / button 文案）
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import pytest  # noqa: E402

from core.task_events import AppState  # noqa: E402
from gui_view_model import (  # noqa: E402
    state_text, state_bootstyle, state_progress_mode,
    mapping_status, mapping_review_needed,
    filter_items, filter_logo_labels, filter_stores,
    logo_display_state, progress_display, batch_summary_model,
    summary_duration_text, config_check_model, layer_display_label,
)
from gui_styles import (  # noqa: E402
    BS_DANGER, BS_INFO, BS_PRIMARY, BS_SECONDARY, BS_SUCCESS, BS_WARNING,
)


# ---------------------------------------------------------------------------
# 1. AppState -> 视觉映射（规格 6.5B 第 6 节：短文案）
# ---------------------------------------------------------------------------
def test_state_text_all_states():
    """#1 7 个状态都有短中文文案。"""
    assert state_text(AppState.IDLE) == "待配置"
    assert state_text(AppState.LOADING) == "正在加载"
    assert state_text(AppState.READY) == "准备就绪"
    assert state_text(AppState.PREVIEWING) == "正在预览"
    assert state_text(AppState.RUNNING) == "正在生成"
    assert state_text(AppState.STOPPING) == "正在停止"
    assert state_text(AppState.ERROR) == "需要处理"


def test_state_text_short_no_old_wording():
    """#1b 状态文案不含旧长文本（等待配置/正在分析/正在批量生成/发生错误 已精简）。"""
    for st in AppState:
        t = state_text(st)
        assert "等待配置" not in t
        assert "分析" not in t
        assert "批量生成" not in t
        assert "发生错误" not in t
        assert "..." not in t and "…" not in t      # 无省略号
        assert len(t) <= 5                          # 全部 3~4 字短文案


def test_state_bootstyle_all_states():
    """#2 7 个状态都有徽标色（纯色，tb.Label 可用；RUNNING 用 primary 蓝）。"""
    mapping = {
        AppState.IDLE: BS_SECONDARY,
        AppState.LOADING: BS_INFO,
        AppState.READY: BS_SUCCESS,
        AppState.PREVIEWING: BS_INFO,
        AppState.RUNNING: BS_PRIMARY,       # 规格 6.5B：正在生成 = 主操作蓝，非绿
        AppState.STOPPING: BS_WARNING,
        AppState.ERROR: BS_DANGER,
    }
    for st, bs in mapping.items():
        assert state_bootstyle(st) == bs
        assert "-outline" not in bs   # 徽标必须纯色（tb.Label 不支持 outline）


def test_running_not_success_bootstyle():
    """#2b RUNNING 不再用 success 绿（与「成功完成」区分）。"""
    assert state_bootstyle(AppState.RUNNING) == BS_PRIMARY
    assert state_bootstyle(AppState.RUNNING) != BS_SUCCESS


def test_state_progress_mode():
    """#3 LOADING/PREVIEWING indeterminate，其余 determinate。"""
    assert state_progress_mode(AppState.LOADING) == "indeterminate"
    assert state_progress_mode(AppState.PREVIEWING) == "indeterminate"
    assert state_progress_mode(AppState.RUNNING) == "determinate"
    assert state_progress_mode(AppState.READY) == "determinate"
    assert state_progress_mode(AppState.IDLE) == "determinate"
    assert state_progress_mode(AppState.STOPPING) == "determinate"
    assert state_progress_mode(AppState.ERROR) == "determinate"


# ---------------------------------------------------------------------------
# 2. 门店映射状态
# ---------------------------------------------------------------------------
def test_mapping_status_missing():
    """#4 未映射 -> danger。"""
    text, bs = mapping_status("店A", "", False)
    assert bs == BS_DANGER
    assert "未映射" in text
    text2, bs2 = mapping_status("店A", "（无）", True)
    assert bs2 == BS_DANGER


def test_mapping_status_auto():
    """#5 自动匹配 -> info（徽标「● 手动」）。"""
    text, bs = mapping_status("店A", "logo_x", True)
    assert bs == BS_INFO
    assert "手动" in text


def test_mapping_status_confirmed():
    """#6 手动确认 -> success（徽标「✓ 已匹配」）。"""
    text, bs = mapping_status("店A", "logo_x", False)
    assert bs == BS_SUCCESS
    assert "已匹配" in text


def test_mapping_review_needed():
    """#7 未映射需要确认。"""
    assert mapping_review_needed("店A", "", False) is True
    assert mapping_review_needed("店A", "（无）", True) is True
    assert mapping_review_needed("店A", "logo_x", False) is False


# ---------------------------------------------------------------------------
# 3. 搜索过滤
# ---------------------------------------------------------------------------
def test_filter_items_empty_query():
    """#8 空查询返回全部。"""
    items = ["a", "b", "c"]
    assert filter_items(items, "") == items
    assert filter_items(items, None) == items
    assert filter_items(items, "  ") == items


def test_filter_items_substring():
    """#9 大小写不敏感子串匹配。"""
    items = ["Apple", "Banana", "Apricot"]
    assert filter_items(items, "ap") == ["Apple", "Apricot"]
    assert filter_items(items, "BAN") == ["Banana"]


def test_filter_logo_labels_by_label():
    """#10 按 label 过滤。"""
    labels = ["logo_七方", "logo_圣大", "名字"]
    assert filter_logo_labels(labels, "logo") == ["logo_七方", "logo_圣大"]
    assert filter_logo_labels(labels, "名字") == ["名字"]
    assert filter_logo_labels(labels, "") == labels


def test_filter_logo_labels_by_display_path():
    """#11 按完整 display_path 过滤。"""
    labels = ["logo_七方", "logo_圣大"]
    path_of = {"logo_七方": "组1/品牌区/logo_七方",
               "logo_圣大": "组2/其他/logo_圣大"}
    assert filter_logo_labels(labels, "品牌区", path_of) == ["logo_七方"]
    assert filter_logo_labels(labels, "组2", path_of) == ["logo_圣大"]


def test_filter_stores():
    """#12 门店过滤。"""
    stores = ["江华电器", "大华超市", "华美商行"]
    assert filter_stores(stores, "华") == ["江华电器", "大华超市", "华美商行"]
    assert filter_stores(stores, "超市") == ["大华超市"]
    assert filter_stores(stores, "") == stores


# ---------------------------------------------------------------------------
# 4. selected vs brand 展示（规格 6.5B 第 13 节：禁 BRAND 术语，改「固定显示」）
# ---------------------------------------------------------------------------
def test_logo_display_state_all_combos():
    """#13 selected/brand 四种组合展示（用户语言，无 BRAND）。"""
    assert logo_display_state("x", True, True) == "固定+参与"
    assert logo_display_state("x", False, True) == "固定"
    assert logo_display_state("x", True, False) == "参与"
    assert logo_display_state("x", False, False) == "候选"


def test_logo_display_state_no_brand_term():
    """#13b 任何输出都不含 BRAND（规格 6.5B 第 13 节）。"""
    for selected in (True, False):
        for is_brand in (True, False):
            out = logo_display_state("x", selected, is_brand)
            assert "BRAND" not in out.upper()
            assert "brand" not in out.lower()


# ---------------------------------------------------------------------------
# 5. 进度显示模型
# ---------------------------------------------------------------------------
def test_progress_display_normal():
    """#14 78/221 -> 35%。"""
    m = progress_display(78, 221, phase="正在导出 PNG…")
    assert m["text"] == "78 / 221"
    assert m["percent"] == 35
    assert m["phase_text"] == "正在导出 PNG…"


def test_progress_display_rounding():
    """#15 百分比四舍五入。"""
    m = progress_display(1, 3)
    assert m["percent"] == 33
    m2 = progress_display(2, 3)
    assert m2["percent"] == 67


def test_progress_display_zero_total():
    """#16 total<=0 -> 0%。"""
    m = progress_display(0, 0)
    assert m["percent"] == 0
    assert m["text"] == "0 / 0"


# ---------------------------------------------------------------------------
# 6. BatchResult -> Summary 模型（规格 6.5B 第 18 节）
# ---------------------------------------------------------------------------
def test_batch_summary_success():
    """#17 全成功 -> ok + success bootstyle。"""
    m = batch_summary_model({
        "success": 3, "failed": 0, "skipped": 1, "cancelled": False,
        "duration_seconds": 112.0, "rows": [],
    })
    assert m["ok"] is True
    assert m["headline"] == "批量生成完成"
    assert m["bootstyle"] == BS_SUCCESS
    assert m["rows_failed"] == []


def test_batch_summary_failed():
    """#18 有失败 -> 非 ok + warning + rows_failed + headline 含失败数。"""
    m = batch_summary_model({
        "success": 2, "failed": 1, "skipped": 0, "cancelled": False,
        "duration_seconds": 30.0,
        "rows": [
            {"excel_row": 3, "status": "FAILED", "store": "店A", "name": "张三",
             "errors": ["PSD 打开失败"]},
            {"excel_row": 4, "status": "SUCCESS", "store": "店B", "name": "李四", "errors": []},
        ],
    })
    assert m["ok"] is False
    assert m["headline"] == "批量生成完成，但有 1 条失败"
    assert m["bootstyle"] == BS_WARNING
    assert len(m["rows_failed"]) == 1
    assert m["rows_failed"][0]["excel_row"] == 3
    assert m["rows_failed"][0]["errors"] == ["PSD 打开失败"]


def test_batch_summary_failed_plural():
    """#18b 多条失败 headline 显示数量。"""
    m = batch_summary_model({
        "success": 0, "failed": 3, "skipped": 0, "cancelled": False,
        "duration_seconds": 10.0, "rows": [],
    })
    assert m["headline"] == "批量生成完成，但有 3 条失败"


def test_batch_summary_cancelled():
    """#19 已停止 -> 部分完成。"""
    m = batch_summary_model({
        "success": 5, "failed": 0, "skipped": 0, "cancelled": True,
        "duration_seconds": 60.0, "rows": [],
    })
    assert m["cancelled"] is True
    assert m["headline"] == "已停止（部分完成）"
    assert m["bootstyle"] == BS_WARNING


def test_batch_summary_empty():
    """#20 空 summary 不崩。"""
    m = batch_summary_model(None)
    assert m["success"] == 0 and m["failed"] == 0
    assert m["ok"] is True
    assert m["rows_failed"] == []


def test_summary_duration_text():
    """#21 秒 -> 人类可读。"""
    assert summary_duration_text(112) == "1 分 52 秒"
    assert summary_duration_text(2 * 60 + 3) == "2 分 3 秒"
    assert summary_duration_text(59) == "59 秒"
    assert summary_duration_text(0) == "0 秒"
    assert summary_duration_text(3700) == "1 小时 1 分 40 秒"


# ---------------------------------------------------------------------------
# 7. 配置检查模型
# ---------------------------------------------------------------------------
def test_config_check_model_missing_all(tmp_path):
    """#22 全空配置 -> 关键项不 ok（分组关闭/无门店 Logo 视为正常）。"""
    items = config_check_model({}, False, [], {}, {}, None, False, "")
    assert len(items) == 6
    # 分组输出关闭 = 正常；无门店 = Logo 映射无未映射项（视为正常）
    for it in items:
        if it["label"] in ("分组输出", "Logo 映射"):
            assert it["ok"] is True, it["label"]
        else:
            assert not it["ok"], it["label"]


def test_config_check_model_psd_exists(tmp_path):
    """#23 PSD 存在 -> ok。"""
    p = tmp_path / "t.psd"
    p.write_bytes(b"x")
    items = config_check_model({"psd_path": str(p)}, False, [], {}, {}, None, False, "")
    psd = next(it for it in items if it["label"] == "PSD 模板")
    assert psd["ok"] is True
    assert psd["bootstyle"] == BS_SUCCESS


def test_config_check_model_store_missing(tmp_path):
    """#24 有门店未映射 -> Logo 映射 warning。"""
    items = config_check_model({"out_dir": str(tmp_path)}, True,
                               ["店A", "店B"],
                               {}, {"店A": "logo_x"}, 2, False, "")
    logo = next(it for it in items if it["label"] == "Logo 映射")
    assert logo["ok"] is False
    assert logo["bootstyle"] == BS_WARNING
    assert "1 个门店未映射" in logo["detail"]


def test_config_check_model_all_ok(tmp_path):
    """#25 全部就绪。"""
    psd = tmp_path / "t.psd"
    psd.write_bytes(b"x")
    xlsx = tmp_path / "t.xlsx"
    xlsx.write_bytes(b"x")
    items = config_check_model(
        {"psd_path": str(psd), "xlsx_path": str(xlsx), "out_dir": str(tmp_path)},
        True, ["店A"], {}, {"店A": "logo_x"}, 10, True, "E")
    assert all(it["ok"] for it in items)


# ---------------------------------------------------------------------------
# 8. LayerRef -> display label（规格 6.5B 第 10 节：UI 永不显示 dict）
# ---------------------------------------------------------------------------
class _FakeLayerRef:
    """最小 LayerRef 替身（有 id/name 属性，模拟 core.layer_index.LayerRef）。"""

    def __init__(self, lid, name):
        self.id = lid
        self.name = name


def test_layer_display_label_plain_string():
    """#26 普通字符串原样返回。"""
    assert layer_display_label("文字图层A") == "文字图层A"
    assert layer_display_label("") == ""


def test_layer_display_label_none():
    """#27 None -> 空字符串。"""
    assert layer_display_label(None) == ""


def test_layer_display_label_dict_with_id_and_name():
    """#28 serialize_ref dict（layer_id+name）-> label_of_id 优先。"""
    ref = {"layer_id": "12", "name": "文字A", "index_path": "1/2"}
    assert layer_display_label(ref, {"12": "文字图层A"}) == "文字图层A"


def test_layer_display_label_dict_without_map():
    """#29 dict 无 label 映射 -> 回退 name（绝不显示 dict 原样）。"""
    ref = {"layer_id": "12", "name": "文字A"}
    out = layer_display_label(ref)
    assert out == "文字A"
    assert "layer_id" not in out
    assert "{" not in out and "}" not in out


def test_layer_display_label_dict_no_name():
    """#30 dict 无 name 无映射 -> 空字符串（不崩、不显示 dict）。"""
    out = layer_display_label({"layer_id": "99"})
    assert out == ""
    assert "{'layer_id'" not in out


def test_layer_display_label_layerref():
    """#31 LayerRef 对象 -> label_of_id 优先，否则 name。"""
    ref = _FakeLayerRef("7", "文字B")
    assert layer_display_label(ref, {"7": "文字图层B"}) == "文字图层B"
    assert layer_display_label(ref) == "文字B"


def test_layer_display_label_id_fallback_name():
    """#32 LayerRef 有 id 但 label 映射缺该 id -> name。"""
    ref = _FakeLayerRef("7", "文字B")
    assert layer_display_label(ref, {"8": "别的"}) == "文字B"


def test_layer_display_label_json_string_defensive():
    """#33 dict 字符串（历史配置）-> 防御性返回空，不显示 dict。"""
    out = layer_display_label("{'layer_id': '12', 'name': '文字A'}")
    assert out == ""
    assert "layer_id" not in out


def test_layer_display_label_never_leaks_dict_syntax():
    """#34 所有输入形态都不输出 dict/JSON 语法字符（{ } [ ] '）。"""
    cases = [
        None,
        {"layer_id": "12", "name": "A"},
        {"id": "12"},
        {"name": "B"},
        "{'layer_id': '12'}",
        _FakeLayerRef("12", "C"),
        _FakeLayerRef("", ""),
    ]
    for c in cases:
        out = layer_display_label(c)
        for ch in ("{", "}", "[", "]", "'layer_id'", "layer_id:"):
            assert ch not in out, f"{c!r} -> {out!r} 泄露内部结构"


# ---------------------------------------------------------------------------
# 9. 规格 6.5B 第 30 节：源码级视觉回归（不实例化 GUI）
# ---------------------------------------------------------------------------
_UI_SOURCE = os.path.join(os.path.dirname(__file__), "..", "qifang_cover_maker.py")
with open(_UI_SOURCE, encoding="utf-8") as _f:
    _UI_CODE = _f.read()


def test_ui_never_contains_dict_display_of_layerref():
    """#35 UI 源码不出现 '{'layer_id' 形态的 dict 拼接（回归测试）。"""
    assert "{'layer_id'" not in _UI_CODE
    assert "layer_id':" not in _UI_CODE


def test_ui_header_single_status_badge():
    """#36 Header 只保留单个状态徽标（status_label 隐藏）。"""
    assert "status_label.pack_forget" in _UI_CODE
    assert "status_badge" in _UI_CODE


def test_ui_file_tab_new_sections():
    """#37 文件页四个 Section 名称齐全（文件与输出/数据字段/输出分组/文字图层）。"""
    for sec in ("文件与输出", "数据字段", "输出分组", "文字图层"):
        assert f'text="{sec}"' in _UI_CODE, f"缺少 Section: {sec}"


def test_ui_run_tab_new_sections():
    """#38 生成页五个 Section 名称齐全（输出格式/生成准备/当前任务/本次结果/运行日志）。"""
    for sec in ("输出格式", "生成准备", "当前任务", "本次结果", "运行日志"):
        assert f'text="{sec}"' in _UI_CODE, f"缺少 Section: {sec}"


def test_ui_logo_tab_hint_and_columns():
    """#39 Logo 页顶部说明 + 双栏标题（Logo 图层 / 门店映射）+ 固定显示。"""
    assert "选择参与生成的 Logo，并为每个门店指定对应图层" in _UI_CODE
    assert "Logo 图层" in _UI_CODE
    assert "门店映射" in _UI_CODE
    assert "固定显示" in _UI_CODE


def test_ui_data_status_lightweight():
    """#40 数据状态轻量文案（✓ Excel 已读取 · N 条数据 · N 个门店）。"""
    assert "Excel 已读取" in _UI_CODE
    assert "条数据" in _UI_CODE
    assert "个门店" in _UI_CODE


def test_ui_button_labels_no_old_long_wording():
    """#41 Button labels 无旧长文案（规格 6.5B 第 21 节）。"""
    for old in ("浏览...", "浏览…", "加载 PSD 并分析图层", "试做第 1 张预览"):
        assert old not in _UI_CODE, f"旧按钮文案残留: {old}"
    for new in ("加载并分析", "生成预览", "开始生成", "保存配置", "选择", "停止"):
        assert f'text="{new}"' in _UI_CODE or new in _UI_CODE, f"缺少新按钮: {new}"
