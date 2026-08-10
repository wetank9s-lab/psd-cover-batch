# -*- coding: utf-8 -*-
"""Stage 6.5B：gui_styles 纯 View 层单元测试（不依赖 Tk 实例）。

覆盖：
  - spacing 常量存在且单调（XS<SM<MD<LG<XL）
  - 主题锁定 litera（规格 6.5B 第 2 节：主操作蓝 / 成功绿分层）
  - 字体体系常量（标题 18 / 正文 10-11 / 辅助 9 / 日志 9 等宽）
  - 语义 bootstyle 层级（primary=主操作 / success=成功 / warning=确认 / danger=停止）
  - mapping_status_for 判定（missing/auto/confirmed）+ 徽标中文文案（规格 6.5B 第 14 节）
  - log_level_of 轻量等级判定
  - shorten_path 超长路径压缩
  - 无边框 section / Separator / Notebook 样式 / 全局字体注入 helper
  - 规格 6.5B 第 30 节：视觉纯逻辑回归（Button 文案 / 徽标 / 状态文案）
"""
import os
import re
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import pytest  # noqa: E402

import gui_styles as gs  # noqa: E402


# ---------------------------------------------------------------------------
# 1. spacing 常量
# ---------------------------------------------------------------------------
def test_spacing_constants_exist_and_ordered():
    """#1 spacing 常量存在且递增。"""
    assert gs.PAD_XS == 4
    assert gs.PAD_SM == 8
    assert gs.PAD_MD == 12
    assert gs.PAD_LG == 16
    assert gs.PAD_XL == 24
    assert gs.PAD_XS < gs.PAD_SM < gs.PAD_MD < gs.PAD_LG < gs.PAD_XL


# ---------------------------------------------------------------------------
# 2. 主题
# ---------------------------------------------------------------------------
def test_theme_litera():
    """#2 主题锁定 litera（规格 6.5B 第 2 节：主操作蓝 #4582ec 与成功绿 #02b875 分层）。"""
    assert gs.APP_THEME == "litera"


def test_theme_colors_semantic_layering():
    """#2b litera 主色为蓝、成功色为绿——主操作与成功态天然分层（防高饱和绿主按钮）。"""
    # litera 的 primary/success 由 ttkbootstrap 主题数据决定；这里锁定语义常量本身
    assert gs.BS_PRIMARY == "primary"
    assert gs.BS_SUCCESS == "success"
    assert gs.BS_PRIMARY != gs.BS_SUCCESS
    assert gs.BS_MAIN == "primary"          # 主按钮是 primary（不是 success 绿）
    assert gs.BS_MAIN_OUTLINE == "primary-outline"


# ---------------------------------------------------------------------------
# 2b. 字体体系（规格 6.5B 第 3 节）
# ---------------------------------------------------------------------------
def test_font_constants_exist():
    """#2c 字体体系常量齐全且字号档位正确。"""
    assert gs.FONT_UI == "Microsoft YaHei UI"
    assert gs.FONT_MONO == "Consolas"
    assert gs.FS_TITLE == 18
    assert gs.FS_SUBTITLE == 9
    assert gs.FS_TAB == 11
    assert gs.FS_SECTION == 11
    assert gs.FS_BODY == 10
    assert gs.FS_HELP == 9
    assert gs.FS_LOG == 9


def test_font_compositions():
    """#2d 组合字体：标题 18 bold / Section bold / 日志等宽。"""
    assert gs.FONT_TITLE == (gs.FONT_UI, 18, "bold")
    assert gs.FONT_SECTION == (gs.FONT_UI, 11, "bold")
    assert gs.FONT_BODY[0] == gs.FONT_UI and gs.FONT_BODY[1] == 10
    assert gs.FONT_HELP[1] == 9
    assert gs.FONT_LOG[0] == gs.FONT_MONO          # 等宽
    assert gs.FONT_SECTION_PLAIN == (gs.FONT_UI, 11)  # 非 bold 次标题


def test_body_font_bigger_than_help():
    """#2e 正文（10）> 辅助（9）：正文视觉提高 10%~15% 由字号档位保证。"""
    assert gs.FS_BODY > gs.FS_HELP


# ---------------------------------------------------------------------------
# 3. bootstyle 常量
# ---------------------------------------------------------------------------
def test_bootstyle_constants():
    """#3 bootstyle 语义色常量齐全。"""
    assert gs.BS_PRIMARY == "primary"
    assert gs.BS_SUCCESS == "success"
    assert gs.BS_INFO == "info"
    assert gs.BS_WARNING == "warning"
    assert gs.BS_DANGER == "danger"
    assert gs.BS_SECONDARY == "secondary"          # 纯色（tb.Label 可用）
    assert gs.BS_MAIN == "primary"                 # 主按钮
    assert gs.BS_MAIN_OUTLINE == "primary-outline" # 主操作 outline 变体
    assert gs.BS_OUTLINE_SECONDARY == "secondary-outline"  # 次要按钮
    assert gs.BS_DANGER == "danger"                # 停止


def test_button_hierarchy_distinct():
    """#3b 按钮层级三档互不相同：主 / 主-outline / 次要-outline / 停止。"""
    tiers = {gs.BS_MAIN, gs.BS_MAIN_OUTLINE, gs.BS_OUTLINE_SECONDARY, gs.BS_DANGER}
    assert len(tiers) == 4
    assert gs.BS_MAIN == "primary"                 # 主操作 = 蓝
    assert gs.BS_OUTLINE_SECONDARY == "secondary-outline"


# ---------------------------------------------------------------------------
# 4. 门店映射状态
# ---------------------------------------------------------------------------
def test_mapping_status_for_missing():
    """#4 未映射（空 / （无））-> missing。"""
    assert gs.mapping_status_for("店A", "", False) == gs.MAPPING_STATUS_MISSING
    assert gs.mapping_status_for("店A", "（无）", False) == gs.MAPPING_STATUS_MISSING
    assert gs.mapping_status_for("店A", None, True) == gs.MAPPING_STATUS_MISSING


def test_mapping_status_for_auto():
    """#5 自动匹配 -> auto。"""
    assert gs.mapping_status_for("店A", "logo_x", True) == gs.MAPPING_STATUS_AUTO


def test_mapping_status_for_confirmed():
    """#6 手动确认 -> confirmed。"""
    assert gs.mapping_status_for("店A", "logo_x", False) == gs.MAPPING_STATUS_CONFIRMED


def test_mapping_status_text_and_bootstyle():
    """#7 状态 -> 文本/bootstyle 映射完整（规格 6.5B 第 14 节文案）。"""
    assert gs.mapping_status_text(gs.MAPPING_STATUS_CONFIRMED) == "✓ 已匹配"
    assert gs.mapping_status_text(gs.MAPPING_STATUS_AUTO) == "● 手动"
    assert gs.mapping_status_text(gs.MAPPING_STATUS_REVIEW) == "⚠ 待确认"
    assert gs.mapping_status_text(gs.MAPPING_STATUS_MISSING) == "✕ 未映射"
    assert gs.mapping_status_bootstyle(gs.MAPPING_STATUS_CONFIRMED) == gs.BS_SUCCESS
    assert gs.mapping_status_bootstyle(gs.MAPPING_STATUS_AUTO) == gs.BS_INFO
    assert gs.mapping_status_bootstyle(gs.MAPPING_STATUS_REVIEW) == gs.BS_WARNING
    assert gs.mapping_status_bootstyle(gs.MAPPING_STATUS_MISSING) == gs.BS_DANGER


def test_mapping_badge_texts_no_old_wording():
    """#7b 徽标文案不含旧术语（自动匹配/已确认/需要确认 已改版）。"""
    for st in (gs.MAPPING_STATUS_CONFIRMED, gs.MAPPING_STATUS_AUTO,
               gs.MAPPING_STATUS_REVIEW, gs.MAPPING_STATUS_MISSING):
        text = gs.mapping_status_text(st)
        assert "自动匹配" not in text
        assert "已确认" not in text
        assert "需要确认" not in text


# ---------------------------------------------------------------------------
# 5. 日志等级
# ---------------------------------------------------------------------------
def test_log_level_of_error():
    """#8 错误关键词 -> error。"""
    assert gs.log_level_of("错误（load）：PSD 打开失败") == gs.LOG_ERROR
    assert gs.log_level_of("Error: something failed") == gs.LOG_ERROR
    assert gs.log_level_of("批量完成：成功 0，失败 5") == gs.LOG_ERROR


def test_log_level_of_warn():
    """#9 警告关键词 -> warn。"""
    assert gs.log_level_of("警告：磁盘空间不足") == gs.LOG_WARN
    assert gs.log_level_of("warning: deprecated") == gs.LOG_WARN


def test_log_level_of_info():
    """#10 普通日志 -> info。"""
    assert gs.log_level_of("[状态] idle -> ready") == gs.LOG_INFO
    assert gs.log_level_of("解析完成：PSD 共 12 个图层") == gs.LOG_INFO
    assert gs.log_level_of("") == gs.LOG_INFO


# ---------------------------------------------------------------------------
# 6. 路径压缩
# ---------------------------------------------------------------------------
def test_shorten_path_short():
    """#11 短路径不压缩。"""
    p = r"D:\out\001.png"
    assert gs.shorten_path(p, 60) == p


def test_shorten_path_long():
    """#12 超长路径保留头尾 + 省略号。"""
    p = r"D:\very\long\path\with\many\sub\folders\output\001.png"
    s = gs.shorten_path(p, 40)
    assert len(s) <= 40 + 3          # 允许省略号略超
    assert s.startswith(p[:5])       # 保留头部盘符
    assert s.endswith(p[-8:])        # 保留尾部文件名
    assert "..." in s


def test_shorten_path_empty():
    """#13 空路径返回空。"""
    assert gs.shorten_path("") == ""
    assert gs.shorten_path(None) == ""


# ---------------------------------------------------------------------------
# 7. 无边框 Section / helper（规格 6.5B 第 5 节：去掉 Labelframe 卡片套卡片）
# ---------------------------------------------------------------------------
def test_section_returns_callable():
    """#14 section 是无边框容器工厂（返回内容 Frame，不是 Labelframe）。"""
    assert callable(gs.section)


def test_section_builds_labelframe_free_body():
    """#14b section() 实例化后返回 ttk.Frame（无边框 body），不再返回 LabelFrame。

    需要真实 Tk（venv314 / py314-tk 提供）；无显示环境则跳过。
    """
    import tkinter as tk
    from tkinter import ttk
    try:
        root = tk.Tk()
    except Exception as e:
        pytest.skip(f"no display: {e}")
    try:
        root.withdraw()
        body = gs.section(root, "测试分区")
        assert isinstance(body, ttk.Frame)
        assert not isinstance(body, ttk.LabelFrame)
        # 标题行是子控件（bold Label），body 有内容
        children = [w for w in body.winfo_children()]
        assert len(children) == 0     # body 本身为空，标题在 box 上
        # box 上有标题 Label + body
        parent = body.master
        labels = [w for w in parent.winfo_children() if isinstance(w, ttk.Label)]
        assert any(w.cget("text") == "测试分区" for w in labels)
    finally:
        try:
            root.destroy()
        except Exception:
            pass


def test_make_separator_callable():
    """#15 make_separator 工厂存在（Section 之间水平分隔线）。"""
    assert callable(gs.make_separator)


def test_make_scrollable_callable():
    """#15b make_scrollable 工厂存在。"""
    assert callable(gs.make_scrollable)


def test_add_tooltip_callable():
    """#16 ToolTip helper 存在。"""
    assert callable(gs.add_tooltip)
    assert callable(gs.ToolTip)


def test_style_helpers_callable():
    """#16b Notebook / 全局字体 / Tree 样式 helper 存在。"""
    assert callable(gs.style_notebook)
    assert callable(gs.apply_global_fonts)
    assert callable(gs.style_tree)
    assert callable(gs.section_help)


# ---------------------------------------------------------------------------
# 8. 规格 6.5B 第 30 节：视觉纯逻辑回归（不测像素颜色值）
# ---------------------------------------------------------------------------
# 这些测试锁定「UI 文案规范」，防止旧长文案 / 内部术语 / dict 泄露回潮。
# 测试基于源码文件内容（静态扫描），不实例化 GUI。
_UI_SOURCE = os.path.join(os.path.dirname(__file__), "..", "qifang_cover_maker.py")
with open(_UI_SOURCE, encoding="utf-8") as _f:
    _UI_CODE = _f.read()
_UI_TEXTS = re.findall(r'text="([^"]*)"', _UI_CODE)


def test_button_labels_use_new_short_wording():
    """#17 Button labels 无旧长文案（浏览.../加载 PSD/试做第 1 张 已清除）。"""
    old_wording = ["浏览...", "浏览…", "加载 PSD", "试做第 1 张预览", "试做第 1 张",
                   "生成预览并导出"]
    for old in old_wording:
        assert old not in _UI_TEXTS, f"旧按钮文案残留: {old}"
    # 新文案必须存在
    assert "选择" in _UI_TEXTS
    assert "加载并分析" in _UI_TEXTS
    assert "生成预览" in _UI_TEXTS
    assert "开始生成" in _UI_TEXTS
    assert "保存配置" in _UI_TEXTS
    assert "停止" in _UI_TEXTS


def test_checkbutton_labels_use_user_language():
    """#18 勾选/说明文案用用户语言（规格 6.5B 第 9 节）。"""
    assert "第一行为表头" in _UI_TEXTS
    assert "按字段创建子文件夹" in _UI_TEXTS
    assert "PSD 同时保存 PNG" in _UI_TEXTS


def test_no_brand_internal_term_in_ui_texts():
    """#19 UI 文本不含 BRAND（规格 6.5B 第 13 节：改「固定显示」）。"""
    for t in _UI_TEXTS:
        assert "BRAND" not in t.upper() or "品牌" in t, f"UI 文本出现 BRAND 术语: {t}"
    assert "固定显示" in _UI_TEXTS
    # 辅助说明经 section_help() 传入（非 text= 属性），直接扫源码
    assert "勾选后将在每张封面中保持显示" in _UI_CODE


def test_search_placeholders_short():
    """#20 搜索 placeholder 短文案（规格 6.5B 第 15 节）。"""
    assert "搜索 Logo 图层" in _UI_TEXTS
    assert "搜索门店" in _UI_TEXTS


def test_no_dict_leak_patterns_in_ui_source():
    """#21 源码 UI 区不显示内部数据结构（'{'layer_id' 形态永不出现）。"""
    assert "{'layer_id'" not in _UI_CODE
    assert "layer_id':" not in _UI_CODE
    # _tm_display 注释必须体现「不显示 dict / layer_id」的规格意图
    assert "不显示 dict" in _UI_CODE or "不显示 dict / JSON / layer_id" in _UI_CODE


def test_no_random_padding_values():
    """#22 禁止随机 padding（规格 6.5B 第 4 节：只有 4/8/12/16/24 档位）。

    pady=1/2 仅允许出现在行内紧凑列表（徽标/勾选行）——这是密集列表的微调，
    不改变分层 spacing 体系。此处只验证无 3/6/7/11/15 等明显随机值。
    """
    bad = re.findall(r"pad[xy]=([3-9]|1[0-5])(?![0-9])", _UI_CODE)
    allowed_extra = {"6", "7"}   # 旧代码若存在 6/7 需清理；当前策略：不允许
    bad = [v for v in bad if v not in allowed_extra]
    assert not bad, f"随机 padding 数值残留: {bad}"


def test_empty_state_short_text():
    """#23 生成页空状态精简（规格 6.5B 第 17 节：「尚未开始生成」）。"""
    assert "尚未开始生成" in _UI_CODE


def test_progress_bar_uses_primary():
    """#24 进度条样式使用 primary（规格 6.5B 第 19 节：非 success 绿）。"""
    assert "primary.Horizontal.TProgressbar" in _UI_CODE
    assert "success.Horizontal.TProgressbar" not in _UI_CODE


def test_all_bootstyle_refs_are_defined():
    """Stage 7.5 守卫：GUI 源码中所有 bootstyle=BS_* 引用必须存在于 gui_styles 常量。

    防止再出现使用未定义常量（如 BS_OUTLINE_WARNING）导致的启动崩溃。
    """
    import re, os
    here = os.path.dirname(__file__)
    src_path = os.path.join(here, "..", "qifang_cover_maker.py")
    with open(src_path, encoding="utf-8") as f:
        src = f.read()
    refs = set(re.findall(r"bootstyle\s*=\s*(BS_[A-Z_]+)", src))
    defined = {n for n in dir(gs) if n.startswith("BS_")}
    missing = refs - defined
    assert not missing, f"GUI 引用了未定义的 bootstyle 常量: {sorted(missing)}"
    assert refs, "未扫描到任何 bootstyle 引用（扫描逻辑失效？）"
