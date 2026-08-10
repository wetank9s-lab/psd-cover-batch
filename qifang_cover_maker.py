# -*- coding: utf-8 -*-
"""
七方视频封面批量制作
===================
通过 Adobe Photoshop (COM 自动化) 批量替换 PSD 中的文字图层（姓名 / 电话 / 销售顾问）
并按门店切换对应的 Logo 图层，导出视频封面图。

依赖：
  - 本机已安装 Adobe Photoshop (2020+)
  - Python 包：openpyxl, pywin32, tkinter(内置)
"""
import os
import sys
import json
import time
import threading
import queue

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

import ttkbootstrap as tb
from ttkbootstrap import Style as TbStyle

import win32com.client
import pythoncom

# Stage 6.5/6.5B：GUI 视觉层（spacing / theme / 字体 / 徽标 / tooltip / 滚动框架 / 纯逻辑 ViewModel）
from gui_styles import (
    APP_THEME, PAD_XS, PAD_SM, PAD_MD, PAD_LG, PAD_XL,
    BS_MAIN, BS_SECONDARY, BS_DANGER, BS_OUTLINE,
    BS_OUTLINE_SECONDARY,
    BS_PRIMARY, BS_SUCCESS, BS_INFO, BS_WARNING,
    add_tooltip, section, make_scrollable,
    shorten_path, log_level_of,
    APP_TITLE_TEXT, APP_SUBTITLE_TEXT,
    FONT_TITLE, FONT_SUBTITLE, FONT_SECTION, FONT_BODY, FONT_HELP, FONT_LOG,
    FONT_TAB, FONT_UI, FS_TAB, FS_BODY, FS_SECTION, FS_HELP,
    FONT_SECTION_PLAIN,
    section_help, make_separator, style_notebook, apply_global_fonts,
)
from gui_view_model import (
    state_text, state_bootstyle, state_progress_mode,
    mapping_status, mapping_review_needed,
    filter_items, filter_logo_labels, filter_stores,
    logo_display_state, progress_display, batch_summary_model,
    summary_duration_text, config_check_model, state_line,
    layer_display_label,
)

# Stage 0：引入 core 纯函数（不依赖 COM / Tk），保持行为一致，供测试与复用
from core import util as core_util
# Stage 1：Photoshop 安全资源管理（Session 只关自己 open/duplicate 的文档，绝不关用户文档）
from core.photoshop import PhotoshopSession, PhotoshopSessionError
# Stage 2：唯一 LayerRef（index_path 精确定位，取消「图层名 = 图层 ID」）
from core.layer_index import (
    LayerRef, LayerIndex, collect_layer_index, resolve_layer,
    rebind_layer_reference, ref_from_config, serialize_ref,
    VALID, MIGRATED, AMBIGUOUS, MISSING,
)
# Stage 3：Logo 模型与映射（selection -> effective -> store/brand -> visibility plan）
from core.logo_mapping import (
    LogoMapping, LogoRole, LogoMatchResult, LogoValidationError, LogoVisibilityError,
    resolve_effective_logo_layers, match_store_logo, validate_logo_mapping,
    prepare_logo_visibility, verify_logo_visibility, migrate_old_config,
    suggest_brand_logos, recommend_logo_selection,
    EXACT, AUTO, AMBIGUOUS as LM_AMBIGUOUS, NO_MATCH,
)
# Stage 4：统一 Excel 数据管线（GUI 全部入口统一走 load_excel_dataset）
from core.excel_data import (
    ExcelRow, ExcelDataset, SkippedRow, ExcelDataError,
    load_excel_dataset, index_to_excel_column, excel_column_to_index,
)
# Stage 4.5：输出分组核心（output path 范畴；与 Excel 解析解耦）
from core.output_paths import (
    OutputPathError, build_group_folder_map, resolve_output_directory,
    assert_group_column_valid,
)
# Stage 5：统一 Renderer（Preview / Batch / CLI 共用 render_one）
from core.renderer import (
    render_one, run_batch as renderer_run_batch,
)

# ---------------- Photoshop 资源管理 ----------------
# Stage 1 起由 core.photoshop.PhotoshopSession 统一管理：
#   - Session 实例持有 app_started_by_tool / owned_documents / initial_documents；
#   - 只关闭 owned documents（本工具 Open / Duplicate 出来的）；
#   - 绝不遍历 Photoshop 当前全部 Documents 执行 Close；
#   - 不再使用模块级粘性全局 _PS_LAUNCHED_BY_US；
#   - 第一版保守 Quit 策略：默认完全不调用 app.Quit()（用户数据安全优先）。


APP_TITLE = "七方视频封面批量制作"
CONFIG_NAME = "qifang_cover_config.json"

# 导出格式
FMT_PNG = "PNG"
FMT_JPG = "JPG"
FMT_PSD = "PSD"

# 文本图层固定名称（PSD 内）
LAYER_NAME = "姓名"
LAYER_PHONE = "电话"
LAYER_ROLE = "销售顾问"


# ----------------------------------------------------------------------------
# Photoshop 辅助函数
# ----------------------------------------------------------------------------
def collect_layer_names(doc):
    """递归收集文档中所有图层。返回 [(name, is_group, parent_name), ...]（去重保序）。"""
    out = []

    def walk(parent, parent_name=""):
        try:
            n = parent.Layers.Count
        except Exception:
            return
        for i in range(1, n + 1):
            try:
                L = parent.Layers[i]
            except Exception:
                continue
            try:
                is_group = L.Layers.Count > 0
            except Exception:
                is_group = False
            out.append((L.Name, bool(is_group), parent_name))
            if is_group:
                walk(L, L.Name)

    walk(doc)
    # 去重但保持顺序
    seen = set()
    res = []
    for t in out:
        if t[0] not in seen:
            seen.add(t[0])
            res.append(t)
    return res


def find_layer(doc, target):
    """递归按名称查找图层对象，忽略名称内的空格差异。返回 Layer 或 None。"""
    t = target.replace(" ", "")

    def walk(parent):
        try:
            n = parent.Layers.Count
        except Exception:
            return None
        for i in range(1, n + 1):
            try:
                L = parent.Layers[i]
            except Exception:
                continue
            if L.Name.replace(" ", "") == t:
                return L
            try:
                if L.Layers.Count > 0:
                    r = walk(L)
                    if r is not None:
                        return r
            except Exception:
                pass
        return None

    return walk(doc)


def find_text_layer(doc, target):
    """查找文字图层，找不到或不是文字层时抛出清晰的中文错误。"""
    L = find_layer(doc, target)
    if L is None:
        raise RuntimeError(f"在 PSD 中找不到文字图层「{target}」（已尝试忽略空格差异）。")
    try:
        _ = L.TextItem.Contents
    except Exception:
        raise RuntimeError(f"图层「{target}」存在，但它不是文字图层，无法写入文本。")
    return L


def collect_text_layer_names(doc):
    """递归收集所有文字图层名称（去重保序）。"""
    out = []

    def walk(parent):
        try:
            n = parent.Layers.Count
        except Exception:
            return
        for i in range(1, n + 1):
            try:
                L = parent.Layers[i]
            except Exception:
                continue
            try:
                _ = L.TextItem.Contents
                if L.Name not in out:
                    out.append(L.Name)
            except Exception:
                pass
            try:
                if L.Layers.Count > 0:
                    walk(L)
            except Exception:
                pass

    walk(doc)
    return out


def safe_replace_text(doc, name, value, log=None):
    """安全替换文字图层文本：找不到或非文字层时记录警告并跳过，不抛异常。"""
    if not name:
        return False
    L = find_layer(doc, name)
    if L is None:
        if log:
            log(f"  警告：找不到文字图层「{name}」，已跳过该字段。")
        return False
    try:
        _ = L.TextItem.Contents
    except Exception:
        if log:
            log(f"  警告：图层「{name}」不是文字图层，已跳过该字段。")
        return False
    set_text_safe(L, value)
    return True


def set_layer_visible_by_name(doc, target, visible):
    """递归按名称设置图层可见性（只匹配第一个同名图层）。
    设为可见时，一并启用其所有父组，确保该图层真正显示出来。"""
    found = [False]

    def walk(parent):
        try:
            n = parent.Layers.Count
        except Exception:
            return
        for i in range(1, n + 1):
            try:
                L = parent.Layers[i]
            except Exception:
                continue
            if not found[0] and L.Name == target:
                try:
                    L.Visible = visible
                    found[0] = True
                    if visible:
                        _enable_parents(parent)
                except Exception:
                    pass
                return
            try:
                if L.Layers.Count > 0:
                    walk(L)
            except Exception:
                pass

    walk(doc)


def _ref_key(x):
    """把配置项（LayerRef dict / 旧 name 字符串 / LayerRef）归一化为唯一 key（id 优先，退化 name）。"""
    if isinstance(x, LayerRef):
        return x.id if x.id else x.name
    r = ref_from_config(x)
    if r is None:
        return None
    return r.id if r.id else r.name


def brand_logos_for(logo_layers, store_logo_map):
    """品牌 Logo：名字含 'logo' 且未被用作任何门店 Logo 的图层 —— 每张封面强制显示。

    与门店 Logo 区分：门店 Logo 会被映射给某个门店（出现在 store_logo_map 的值里），
    按门店切换显隐；品牌 Logo（如「七方logo」）不属于任何门店，应始终可见。

    Stage 2 语义：门店映射判定用「name 级别」——只要该 name 有任一图层被映射为门店，
    所有同名图层都视为门店候选（避免同名 Logo 一个被映射、另一个被误判为品牌）。
    logo_layers / store_logo_map 可能是 LayerRef dict 或旧 name 字符串。
    """
    mapped_names = set()
    for v in store_logo_map.values():
        r = ref_from_config(v)
        if r is not None:
            mapped_names.add(r.name)
    out = []
    for ln in logo_layers:
        r = ref_from_config(ln)
        if r is None:
            continue
        nm = r.name or ""
        if nm.lower().find("logo") >= 0 and nm not in mapped_names:
            out.append(ln)
    return out


# ----------------------------------------------------------------------------
# Stage 3：Logo 运行时辅助（替代 brand_logos_for 的 name 猜测）
# ----------------------------------------------------------------------------
def _build_logo_mapping(cfg, index):
    """从配置构造 LogoMapping（运行时唯一可信来源）。

    读取新字段 logo_selection / brand_logo_layers（LayerRef dict 列表）；
    兼容旧字段 logo_layers / store_logo_map（str 或 LayerRef dict）。

    返回 LogoMapping：
      - store_logo_map 的 value 全部转成叶子 LayerRef（旧 name 通过 LayerIndex rebind；
        歧义/缺失 -> None，绝不在运行时猜）；
      - brand_logo_refs 全部为叶子 LayerRef；
      - logo_selection_refs 保留用户勾选（可能含组，仅供保存/展开）。
    """
    index = index if index is not None else _empty_index()
    store_map = {}

    # 1) 门店映射：优先新格式（dict store -> LayerRef dict/str），兼容旧
    raw_store = cfg.get("store_logo_map") or {}
    for s, v in raw_store.items():
        if v in (None, "", "（无）"):
            store_map[s] = None
            continue
        ref = ref_from_config(v) if not isinstance(v, LayerRef) else v
        if ref is None:
            store_map[s] = None
            continue
        # 迁移到当前 index：唯一命中用；歧义/缺失 -> None（不自动选）
        cand = _rebind_leaf(index, ref)
        store_map[s] = cand

    # 2) 品牌 Logo：优先新字段 brand_logo_layers；兼容旧：从 logo_layers 中
    #    名字含 logo 且未被门店使用的叶子（仅首载建议，运行时不再调用）
    brand = []
    raw_brand = cfg.get("brand_logo_layers")
    if raw_brand:
        for v in raw_brand:
            ref = ref_from_config(v) if not isinstance(v, LayerRef) else v
            if ref is None:
                continue
            cand = _rebind_leaf(index, ref)
            if cand is not None:
                brand.append(cand)
    else:
        # 旧配置兜底：从 logo_layers 中按 name 含 logo 且未被门店使用
        mapped_ids = {r.id for r in store_map.values() if r is not None}
        for v in (cfg.get("logo_layers") or []):
            ref = ref_from_config(v) if not isinstance(v, LayerRef) else v
            if ref is None:
                continue
            if "logo" in ref.name.lower() and ref.id not in mapped_ids:
                cand = _rebind_leaf(index, ref)
                if cand is not None:
                    brand.append(cand)

    # 3) selection：优先新字段 logo_selection；兼容旧 logo_layers
    selected = []
    raw_sel = cfg.get("logo_selection")
    if raw_sel is None:
        raw_sel = cfg.get("logo_layers") or []
    for v in raw_sel:
        ref = ref_from_config(v) if not isinstance(v, LayerRef) else v
        if ref is None:
            continue
        cand = index.get(ref.id) if ref.id and index.get(ref.id) else _rebase_any(index, ref)
        if cand is not None and cand.id not in {x.id for x in selected}:
            selected.append(cand)

    return LogoMapping(store_logo_map=store_map, brand_logo_refs=brand,
                       logo_selection_refs=selected)


def _empty_index():
    from core.layer_index import LayerIndex
    return LayerIndex([])


def _rebind_leaf(index, ref):
    """把 ref 绑定到当前 index 的叶子；歧义/缺失/是组 -> None（绝不自动选）。"""
    if ref is None:
        return None
    if ref.id and index.get(ref.id):
        cand = index.get(ref.id)
        if cand.is_group:
            return None
        return cand
    # id 失效 / 纯 name：按 name 唯一命中
    matches = index.find_matching(ref.name)
    if len(matches) == 1 and not matches[0].is_group:
        return matches[0]
    return None


def _rebase_any(index, ref):
    """selection 迁移允许组（旧配置 logo_layers 可能是组名）。"""
    if ref is None:
        return None
    if ref.id and index.get(ref.id):
        return index.get(ref.id)
    matches = index.find_matching(ref.name)
    if len(matches) == 1:
        return matches[0]
    return None


def _validate_runtime_logo(logo_map, index, log_cb):
    """运行前校验 Logo 配置；返回错误消息字符串（无错误返回 None）。

    校验通过后把 effective leaves 缓存在 logo_map 上（运行时 prepare 需要它隐藏
    「候选但未被当前门店使用」的叶子）。
    """
    try:
        effective = resolve_effective_logo_layers(index, logo_map.logo_selection_refs)
        validate_logo_mapping(logo_map, effective_leaf_refs=effective,
                              allow_duplicate_store_targets=True)
        logo_map._effective_leaf_refs = effective
        return None
    except LogoValidationError as e:
        return f"Logo 配置校验失败：{e}"


def _verify_applied_visibility(doc, index, plan, applied, log_cb):
    """写 Visible 后 read-back 校验：不一致抛 LogoVisibilityError（不静默成功）。"""
    readback = {}
    for ref, _expected in plan:
        try:
            layer = resolve_layer(doc, ref)
            readback[ref.id] = bool(layer.Visible)
        except Exception:
            readback[ref.id] = applied.get(ref.id)   # 无法 resolve：以 applied 兜底（日志提示）
    verify_logo_visibility(readback, plan)





def _enable_parents(container):
    """启用 container 及其所有祖先组（直到文档层）。"""
    p = container
    while True:
        try:
            p.Visible = True
            p = p.Parent
        except Exception:
            break


# ----------------------------------------------------------------------------
# Stage 2：LayerRef 精确操作辅助（替代「按 name 找第一个」的运行时定位）
# ----------------------------------------------------------------------------
def _find_indexed_layer(doc, index, ref_or_name):
    """把配置值（LayerRef dict / 旧 name 字符串 / LayerRef 对象）解析为 COM Layer。

    - 优先按 LayerRef.index_path 精确定位（resolve_layer，绝不 fallback 同名）；
    - 纯 name（旧配置 / 无 index_path）时走 rebind：唯一 -> 用之；歧义/缺失 -> None；
    - 返回 (layer, status)；status 用于日志（VALID/MIGRATED/AMBIGUOUS/MISSING/STALE）。
    """
    from core.layer_index import resolve_layer
    if isinstance(ref_or_name, LayerRef):
        ref = ref_or_name
    else:
        ref = ref_from_config(ref_or_name)
    if ref is None:
        return None, MISSING
    if ref.index_path:
        try:
            return resolve_layer(doc, ref), VALID
        except Exception as e:
            # index_path 失效（结构变化）：尝试按 name rebind，标记 MIGRATED/AMBIGUOUS
            status, cand = rebind_layer_reference(index, ref.name)
            if cand is not None:
                try:
                    return resolve_layer(doc, cand), MIGRATED
                except Exception:
                    return None, status
            return None, status
    # 纯 name（旧配置）
    status, cand = rebind_layer_reference(index, ref.name)
    if cand is None:
        return None, status
    try:
        return resolve_layer(doc, cand), status
    except Exception:
        return None, status


def set_text_by_ref(doc, index, ref_or_name, value, log=None, label=""):
    """按 LayerRef 精确替换文字层（找不到/歧义时记日志并跳过，不抛异常）。"""
    if not ref_or_name:
        return False
    layer, status = _find_indexed_layer(doc, index, ref_or_name)
    if layer is None:
        if log:
            reason = {"AMBIGUOUS": "同名图层多个，无法唯一确定",
                      "MISSING": "找不到该图层",
                      "MIGRATED": "已按唯一名称迁移"}.get(status, str(status))
            log(f"  警告：文字层 {label or ref_or_name} 定位失败（{reason}），已跳过该字段。")
        return False
    try:
        _ = layer.TextItem.Contents
    except Exception:
        if log:
            log(f"  警告：图层 {label or ref_or_name} 不是文字图层，已跳过该字段。")
        return False
    return set_text_safe(layer, value)


def set_visible_by_ref(doc, index, ref_or_name, visible, log=None, label=""):
    """按 LayerRef 精确设置图层可见性；可见时启用父组。返回 (ok, status)。"""
    if ref_or_name is None:
        return False, MISSING
    layer, status = _find_indexed_layer(doc, index, ref_or_name)
    if layer is None:
        return False, status
    try:
        layer.Visible = visible
        if visible:
            try:
                _enable_parents(layer.Parent)
            except Exception:
                pass
        return True, status
    except Exception as e:
        if log:
            log(f"  警告：设置可见性失败：{label or ref_or_name}：{e}")
        return False, status


def set_text_safe(layer, text, retries=6):
    for attempt in range(retries):
        try:
            layer.TextItem.Contents = text
            return True
        except Exception:
            if attempt == retries - 1:
                return False
            time.sleep(0.3 * (attempt + 1))


def export_doc(doc, path, fmt):
    if fmt == FMT_PNG:
        opt = win32com.client.Dispatch("Photoshop.PNGSaveOptions")
        opt.Interlaced = False
        opt.Compression = 6
    elif fmt == FMT_JPG:
        opt = win32com.client.Dispatch("Photoshop.JPGSaveOptions")
        opt.Quality = 9
        opt.EmbedColorProfile = False
    else:  # PSD
        opt = win32com.client.Dispatch("Photoshop.PhotoshopSaveOptions")
    doc.SaveAs(path, opt, True)


def _group_plan_summary(folder_map, group_column):
    """分组配置 -> 日志摘要（批次开始前调用；无分组返回 None）。"""
    if group_column is None:
        return None
    lines = [f"分组字段：{index_to_excel_column(group_column)}",
             f"预计创建：{folder_map.distinct_folder_count} 个目录"]
    if folder_map.empty_group_count:
        lines.append(f"空值：{folder_map.empty_group_count} 条 → {folder_map.fallback}")
    if folder_map.collision_count:
        lines.append(f"名称冲突：{folder_map.collision_count} 组，已自动区分")
        for cl in folder_map.collision_summary():
            lines.append(f"    {cl}")
    return "；".join(lines)


# ----------------------------------------------------------------------------
# 与 GUI 解耦的核心批处理逻辑（在 worker 线程调用）
# ----------------------------------------------------------------------------
def run_batch(cfg, progress_cb, log_cb, stop_flag):
    """
    cfg 字段：
      psd_path, xlsx_path, out_dir, has_header(bool),
      col_store, col_name, col_phone, col_role (0-based; col_role=-1 表示不替换),
      text_map (dict: "姓名"/"电话"/"销售顾问" -> 实际图层名 或 ""表示不替换),
      logo_layers (list[str]), store_logo_map (dict store->logo_layer or ""),
      fmt, also_png(bool)
    """
    # COM 初始化/反初始化由 PhotoshopSession 的 __enter__/__exit__ 统一负责
    try:
        # ---- 读取 Excel（Stage 4：统一入口 load_excel_dataset）----
        try:
            dataset = load_excel_dataset(
                cfg["xlsx_path"],
                has_header=cfg.get("has_header", True),
                col_store=cfg["col_store"],
                col_name=cfg["col_name"],
                col_phone=cfg["col_phone"],
                col_role=cfg["col_role"] if cfg.get("col_role", -1) >= 0 else None,
            )
        except ExcelDataError as e:
            log_cb(f"Excel 错误：{e}")
            return
        data = dataset.valid_rows
        log_cb(f"Excel 读取完成：{len(data)} 行有效数据，"
               f"跳过 {len(dataset.skipped_rows)} 行（sheet：{dataset.sheet_name}）。")
        if not data:
            log_cb("没有可处理的数据，已停止。")
            return

        # ---- 连接 Photoshop（Session 只管理自己打开/复制的文档）----
        log_cb("正在启动 / 连接 Photoshop ...")
        with PhotoshopSession() as ps:
            app = ps.app
            doc0 = ps.open_document(cfg["psd_path"])
            time.sleep(0.6)

            # Stage 2：运行时建立 LayerIndex（对当前打开的模板文档）
            index = collect_layer_index(doc0)

            text_map = cfg.get("text_map", {})
            name_ref = text_map.get("姓名", "")
            phone_ref = text_map.get("电话", "")
            role_ref = text_map.get("销售顾问", "")

            log_cb(f"模板已加载（{doc0.Width}x{doc0.Height}）。"
                   f"文字层映射：姓名→{name_ref.get('display_path') if isinstance(name_ref, dict) else (name_ref or '（不替换）')}，"
                   f"电话→{phone_ref.get('display_path') if isinstance(phone_ref, dict) else (phone_ref or '（不替换）')}，"
                   f"销售顾问→{role_ref.get('display_path') if isinstance(role_ref, dict) else (role_ref or '（不替换）')}。")

            out_dir = cfg["out_dir"]
            os.makedirs(out_dir, exist_ok=True)
            # Stage 4.5：分组功能启用时，批次开始前 Preflight 校验分组列有效性
            group_enabled = bool(cfg.get("group_output_enabled"))
            group_column = cfg.get("group_output_column")
            if group_enabled:
                try:
                    assert_group_column_valid(dataset.max_columns, True, group_column)
                except OutputPathError as e:
                    log_cb(f"分组配置错误：{e}")
                    return
                folder_map = build_group_folder_map(data, group_column)
                summary = _group_plan_summary(folder_map, group_column)
                if summary:
                    log_cb(f"[分组计划] {summary}")
            else:
                folder_map = None
            # Stage 3：Logo 运行时数据 = LogoMapping（store_logo_map / brand_logo_refs / logo_selection_refs）。
            # 不再在运行时用 name heuristic 判断品牌/门店（brand_logos_for 已废弃）。
            logo_map = _build_logo_mapping(cfg, index)
            logo_validation = _validate_runtime_logo(logo_map, index, log_cb)
            if logo_validation:
                log_cb(logo_validation)
                return
            fmt = cfg["fmt"]
            also_png = cfg["also_png"]

            # ---- Stage 5：统一 Renderer（render_one/run_batch 收敛全部单行逻辑）----
            # Batch 只负责：打开模板 → 建 index / logo_map / folder_map → 调 renderer_run_batch
            # 汇总 BatchResult → 按 RowResult 展示成功/失败/跳过。不再自行
            # Duplicate / 替换文字 / Logo / SaveAs / Close。
            total = len(data)
            t0 = time.time()
            br = renderer_run_batch(
                ps_session=ps,
                template_doc=doc0,
                rows=data,
                config={
                    "fmt": fmt,
                    "also_png": also_png,
                    "text_map": cfg.get("text_map", {}),
                    "group_output_enabled": group_enabled,
                    "group_output_column": group_column,
                },
                layer_index=index,
                logo_mapping=logo_map,
                out_dir=out_dir,
                folder_map=folder_map,
                cancel_event=stop_flag,
                progress_cb=progress_cb,
                log=log_cb,
                com_dispatch=win32com.client.Dispatch,
            )
            el = time.time() - t0
            log_cb(
                f"完成！成功 {br.success} 张，失败 {br.failed} 张"
                + (f"，跳过 {br.skipped} 张" if br.skipped else "")
                + (f"，用户已停止（剩余 {total - len(br.rows)} 张未处理）" if br.cancelled else "")
                + f"，耗时 {el:.1f}s。输出目录：{out_dir}")
            # 部分失败明细（GUI 展示失败行错误，供用户定位）
            for r in br.rows:
                if r.failed:
                    log_cb(f"  [失败 行{r.excel_row}] {'；'.join(r.errors)}")
    except Exception as e:
        log_cb(f"错误：{e}")
        import traceback
        log_cb(traceback.format_exc())


# ----------------------------------------------------------------------------
# GUI
# ----------------------------------------------------------------------------
class App:
    def __init__(self, root):
        self.root = root
        # Stage 6.5：ttkbootstrap Window（tb.Window 是 tk.Tk 子类，接口兼容）
        root.title(APP_TITLE)
        # 规格第 30 节：默认 1050x780；第 33 节 1366x768 硬验收 ——
        # 屏幕高度不足时（768 减任务栏 40px），自动压缩到可用高度，保证完整可用。
        try:
            sw = root.winfo_screenwidth()
            sh = root.winfo_screenheight()
            win_w = min(1050, max(980, sw - 80))
            # 内容高 = 屏幕高 - 任务栏(40) - 标题栏(≈15)；下限 660 保证可用。
            # 1366x768 -> 713（Tab1 固定内容堆叠不溢出，按钮区完整可见）
            win_h = min(780, max(660, sh - 55))
        except Exception:
            win_w, win_h = 1050, 780
        root.geometry(f"{win_w}x{win_h}")
        root.minsize(980, 640)
        root.resizable(True, True)

        self.config_path = self._find_config_path()
        self.cfg = {
            "psd_path": "", "xlsx_path": "", "out_dir": "",
            "has_header": True,
            "col_store": 0, "col_name": 1, "col_phone": 3, "col_role": 2,
            "text_map": {"姓名": "", "电话": "", "销售顾问": ""},
            "logo_layers": [], "store_logo_map": {},
            "fmt": FMT_PNG, "also_png": False,
        }
        # 运行时状态
        self.all_psd_layers = []
        self.all_psd_is_group = {}
        self.all_psd_parent = {}
        self.all_text_layers = []
        # Stage 2：LayerRef 状态（唯一图层身份）
        self.layer_index = None          # 当前 PSD 的 LayerIndex（不丢同名）
        self._ref_by_id = {}             # layer_id -> LayerRef
        self.layer_labels = {}           # layer_id -> 展示 label（同名时含 [id=...] 后缀，保证肉眼可区分）
        self.text_label_to_ref = {}      # label -> LayerRef（文字层）
        self.logo_label_to_ref = {}      # label -> LayerRef（当前有效 Logo 叶子）
        # Stage 3 补充：品牌 Logo 人工指定（selected 叶子 -> BRAND）
        #   brand_checks: label -> BooleanVar（勾选 = 该 leaf 是固定品牌 Logo）
        #   brand_widgets: label -> Checkbutton（重建时记录，供冲突回滚刷新）
        self.brand_checks = {}
        self.brand_widgets = {}
        self.excel_stores = []
        self.excel_headers = []
        # Stage 4：当前加载的 ExcelDataset（统一数据源；_load 后才有值）
        self.excel_dataset = None
        # Stage 4 补充（BLOCKED 修复）：Dataset 一致性 key —— path/has_header/4 列任一变化即失效
        self._ds_key = None
        # Stage 7.5：PSD 指纹（mtime）—— 记录最近一次成功 Load 时的磁盘状态，
        # 用于检测「PSD 文件被外部修改（新增门店图层等）但 GUI 仍显示旧数据」。
        self._psd_fingerprint = None     # (path, mtime, size) 或 None（未加载过）
        self.running = False
        self.stop_flag = threading.Event()
        self.worker = None
        # Stage 6：Worker / Queue / AppState
        from core.task_events import AppState as _AppState
        from core.worker_base import WorkerAlreadyRunningError
        from gui_workers import TaskWorker
        self._AppState = _AppState
        self._WorkerAlreadyRunningError = WorkerAlreadyRunningError
        self._state = _AppState.IDLE
        self._task_worker = TaskWorker()
        self._worker_polling = False      # 防止重复启动 poll
        self._pending_close = False       # 窗口关闭请求
        self._start_polling()

        # Stage 6.5B：注入统一字体体系 + Notebook 样式（litera 主题已由 main() 创建）
        try:
            apply_global_fonts(root.style)
            style_notebook(root.style)
        except Exception:
            pass

        self._build_ui()
        # 初始状态应用 IDLE 控件规则（UI 构建后控件默认可点，需按状态机收口）
        from core.app_state import controls_for as _controls_for
        self._apply_controls(_controls_for(_AppState.IDLE))
        self._load_config()

    # ---------------- 路径 ----------------
    def _find_config_path(self):
        if getattr(sys, "frozen", False):
            base = os.path.dirname(sys.executable)
        else:
            base = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(base, CONFIG_NAME)

    # ---------------- Stage 6：状态机 ----------------
    def _set_state(self, new_state):
        """集中状态转换：校验 + 更新控件（仅 main thread 调用）。

        Stage 6.5：视觉仍由 AppState + StateControls 决定（不复制状态机）；
        本方法只把状态翻译为视觉（状态文本 / 徽标 / 进度模式）。
        """
        from core.app_state import controls_for, transition
        from core.task_events import is_valid_transition
        old = self._state
        if not is_valid_transition(old, new_state):
            # 非法转换：记录但不崩溃（例如 STOPPING 中收到重复事件）
            self._log_gui(f"[状态] 忽略非法转换 {old.value}->{new_state.value}")
            return
        self._state = new_state
        self._apply_controls(controls_for(new_state))
        # Stage 6.5 视觉映射：AppState -> 状态文本 / 徽标 / 进度模式
        self._apply_state_visual(new_state)
        self._log_gui(f"[状态] {old.value} -> {new_state.value}")

    def _apply_state_visual(self, state):
        """AppState -> 视觉呈现（仅 main thread；消费状态机，不复制状态机）。"""
        from core.task_events import AppState
        # 状态文字（含徽标色）
        if hasattr(self, "status_label"):
            self.status_label.config(text=state_line(state))
        if hasattr(self, "status_badge"):
            try:
                # 徽标文本 + 颜色都随状态刷新（Stage 6.5 修复：此前 text 停留初始值）
                self.status_badge.config(text=state_text(state),
                                         bootstyle=state_bootstyle(state))
            except Exception:
                pass
        # 进度条模式：LOADING/PREVIEWING 用 indeterminate，RUNNING 用 determinate
        if hasattr(self, "progress") and self._state is not None:
            mode = state_progress_mode(self._state)
            try:
                cur = str(self.progress.cget("mode"))
                if cur != mode:
                    self.progress.config(mode=mode)
                    if mode == "indeterminate":
                        self.progress.start(12)   # 仅 main thread 控制
                    else:
                        self.progress.stop()
            except Exception:
                pass

    def _apply_controls(self, c):
        """把控件规则应用到 widget（仅 main thread）。

        Stage 6.5：控件 enabled/disabled 仍完全由 StateControls 决定。
        """
        def _set(btn, enabled):
            if btn is not None:
                btn.config(state="normal" if enabled else "disabled")
        # 控件名 -> StateControls 字段名
        _map = {
            "btn_load": "load",
            "btn_preview": "preview",
            "btn_run": "run",
            "btn_stop": "stop",
        }
        for wname, cname in _map.items():
            if hasattr(self, wname):
                _set(getattr(self, wname), getattr(c, cname))
        # 文件选择 / tab 页
        for var_btn in (getattr(self, "btn_pick_psd", None),
                        getattr(self, "btn_pick_xlsx", None),
                        getattr(self, "btn_pick_out", None)):
            _set(var_btn, c.pick_files)
        try:
            self.notebook.state(["disabled"] if not c.tabs_enabled else ["!disabled"])
        except Exception:
            pass

    def _start_polling(self):
        """启动事件队列轮询（仅一次）。"""
        if self._worker_polling:
            return
        self._worker_polling = True
        self.root.after(80, self._poll_worker_events)

    def _poll_worker_events(self):
        """main thread 轮询 worker 事件队列（Stage 6 #18）。"""
        q = self._task_worker.event_queue
        try:
            while True:
                ev = q.get_nowait()
                self._handle_worker_event(ev)
        except queue.Empty:
            pass
        if not self._pending_close or self._task_worker.worker_alive:
            self.root.after(80, self._poll_worker_events)

    def _handle_worker_event(self, ev):
        """处理一个 worker 事件（仅 main thread；payload 纯 Python）。"""
        from core.task_events import WorkerEvent, AppState
        t = ev.type
        if t == WorkerEvent.STATE:
            st = AppState.from_str(ev.payload) if isinstance(ev.payload, str) else ev.payload
            self._set_state(st)
        elif t == WorkerEvent.LOG:
            self._log_gui(ev.payload)
        elif t == WorkerEvent.PROGRESS:
            self._update_progress(ev.payload.current, ev.payload.total,
                                  phase=ev.payload.phase,
                                  excel_row=ev.payload.excel_row,
                                  store=ev.payload.store,
                                  name=ev.payload.name)
        elif t == WorkerEvent.ROW_STARTED:
            pass  # 日志已由 renderer 输出；如需可在此显示当前行
        elif t == WorkerEvent.ROW_FINISHED:
            pass
        elif t == WorkerEvent.LOAD_DONE:
            self._on_load_done(ev.payload)
        elif t == WorkerEvent.PREVIEW_DONE:
            self._on_preview_done(ev.payload)
        elif t == WorkerEvent.BATCH_DONE:
            self._on_batch_done(ev.payload)
        elif t == WorkerEvent.CANCELLED:
            self._log_gui(f"已停止：{ev.payload}")
        elif t == WorkerEvent.ERROR:
            p = ev.payload
            self._log_gui(f"错误（{p.operation}）：{p.message}")
            if p.fatal:
                self._set_state(AppState.ERROR)
        elif t == WorkerEvent.WORKER_DONE:
            self._on_worker_done()

    def _on_worker_done(self):
        """worker 正常退出（仅 main thread）。"""
        from core.task_events import AppState
        if self._state in (AppState.LOADING,):
            self._set_state(AppState.READY if self.layer_index is not None else AppState.IDLE)
        elif self._state in (AppState.RUNNING, AppState.STOPPING):
            self._set_state(AppState.READY)
        elif self._state in (AppState.PREVIEWING,):
            self._set_state(AppState.READY)
        # 窗口关闭请求且 worker 已退出 -> 允许 destroy
        self._maybe_destroy()

    def _maybe_destroy(self):
        """窗口关闭条件：worker 已退出才 destroy（Stage 6 #15）。"""
        if self._pending_close and not self._task_worker.worker_alive:
            try:
                self.root.destroy()
            except Exception:
                pass

    def _snapshot_for_load(self, force_reload=False):
        """Load 任务快照（纯 Python；main thread 收集）。"""
        return {
            "psd_path": self.psd_var.get().strip(),
            "xlsx_path": self.xlsx_var.get().strip(),
            "has_header": bool(self.header_var.get()),
            "col_store": self._col_of(self.col_store_var.get()),
            "col_name": self._col_of(self.col_name_var.get()),
            "col_phone": self._col_of(self.col_phone_var.get()),
            "col_role": self._col_of(self.col_role_var.get()),
            # Stage 7.5：force_reload=True 时 worker 用 PhotoshopSession.open_document
            # 的 force_reload 参数强制从磁盘重读 PSD（拿到外部新增的门店图层）。
            "force_reload": bool(force_reload),
        }

    @staticmethod
    def _psd_fingerprint_of(path):
        """PSD 磁盘指纹 (mtime, size)；不存在返回 None。"""
        try:
            import os
            st = os.stat(path)
            return (st.st_mtime, st.st_size)
        except OSError:
            return None

    def _psd_changed_since_load(self):
        """PSD 是否自上次成功 Load 后被外部修改（新增图层等）。

        返回 True 仅当：上次成功 Load 的路径与当前 psd_var 一致，且
        磁盘 (mtime, size) 与记录指纹不同。用于「开始/预览前提示刷新」。
        """
        fp = getattr(self, "_psd_fingerprint", None)
        if fp is None:
            return False
        cur = self.psd_var.get().strip()
        if not cur or fp[0] != cur:
            return False
        disk = self._psd_fingerprint_of(cur)
        return disk is not None and disk != fp[1]

    def _snapshot_for_preview(self):
        cfg = self._collect_cfg()
        cfg["_task"] = "preview"
        return cfg

    def _snapshot_for_batch(self):
        cfg = self._collect_cfg()
        cfg["_task"] = "batch"
        return cfg

    # ---------------- UI 构建 ----------------
    def _build_ui(self):
        # Stage 6.5B：顶部标题栏 + Notebook 三 Tab（规格第 6/7 节）
        self._build_header(self.root)
        self.notebook = ttk.Notebook(self.root, style="S65.TNotebook")
        self.notebook.pack(fill="both", expand=True, padx=PAD_LG, pady=(PAD_XS, PAD_LG))

        # Tab 1: 文件与数据
        f1 = ttk.Frame(self.notebook, padding=(PAD_LG, PAD_MD))
        self.notebook.add(f1, text="文件与数据", padding=(PAD_XS, PAD_SM))
        self._build_tab_files(f1)

        # Tab 2: Logo 映射
        f2 = ttk.Frame(self.notebook, padding=(PAD_LG, PAD_MD))
        self.notebook.add(f2, text="Logo 映射", padding=(PAD_XS, PAD_SM))
        self._build_tab_logo(f2)

        # Tab 3: 生成与导出
        f3 = ttk.Frame(self.notebook, padding=(PAD_LG, PAD_MD))
        self.notebook.add(f3, text="生成与导出", padding=(PAD_XS, PAD_SM))
        self._build_tab_run(f3)

    def _build_header(self, parent):
        """顶部标题栏（规格 6.5B 第 6 节）：标题 + 副标题 + 单个状态徽标。

        不再重复显示「状态：等待配置 等待配置」——只保留一个徽标（文字 + 语义色）。
        """
        hdr = ttk.Frame(parent, padding=(PAD_XL, PAD_MD, PAD_XL, PAD_SM))
        hdr.pack(fill="x")
        left = ttk.Frame(hdr)
        left.pack(side="left", fill="x", expand=True)
        ttk.Label(left, text=APP_TITLE_TEXT, font=FONT_TITLE).pack(anchor="w")
        tb.Label(left, text=APP_SUBTITLE_TEXT, bootstyle="secondary",
                 font=FONT_SUBTITLE).pack(anchor="w", pady=(2, 0))
        # 右侧单个状态徽标（AppState 视觉映射；_set_state -> _apply_state_visual 刷新）
        right = ttk.Frame(hdr)
        right.pack(side="right")
        self.status_badge = tb.Label(right, text=state_text(self._state),
                                     bootstyle=state_bootstyle(self._state),
                                     font=FONT_BODY, padding=(PAD_MD, PAD_XS))
        self.status_badge.pack(side="right")
        # status_label 不再单独显示（徽标文字即状态）；保留属性兼容 _apply_state_visual
        self.status_label = ttk.Label(right, text=state_text(self._state), font=FONT_BODY)
        self.status_label.pack_forget()

    def _build_tab_files(self, parent):
        # ---- Section 一：文件与输出（规格 6.5B 第 5/8 节：无 Labelframe，标题+内容+Separator） ----
        # 三行 Entry 严格对齐：label 固定宽 / Entry 占满 / 按钮固定宽
        sec1 = ttk.Frame(parent)
        sec1.pack(fill="x")
        ttk.Label(sec1, text="文件与输出", font=FONT_SECTION).pack(anchor="w")

        grid = ttk.Frame(sec1)
        grid.pack(fill="x", pady=(PAD_XS, 0))
        grid.columnconfigure(1, weight=1)   # Entry 列占满剩余

        ttk.Label(grid, text="PSD 模板", font=FONT_BODY, width=10, anchor="w").grid(
            row=0, column=0, sticky="w", pady=PAD_XS)
        self.psd_var = tk.StringVar()
        self.psd_entry = ttk.Entry(grid, textvariable=self.psd_var, font=FONT_BODY)
        self.psd_entry.grid(row=0, column=1, padx=(PAD_SM, PAD_SM), sticky="we")
        self.btn_pick_psd = tb.Button(grid, text="选择", command=self._pick_psd,
                                      bootstyle=BS_OUTLINE_SECONDARY, width=8)
        self.btn_pick_psd.grid(row=0, column=2, sticky="e")

        ttk.Label(grid, text="Excel 数据", font=FONT_BODY, width=10, anchor="w").grid(
            row=1, column=0, sticky="w", pady=PAD_XS)
        self.xlsx_var = tk.StringVar()
        self.xlsx_entry = ttk.Entry(grid, textvariable=self.xlsx_var, font=FONT_BODY)
        self.xlsx_entry.grid(row=1, column=1, padx=(PAD_SM, PAD_SM), sticky="we")
        self.btn_pick_xlsx = tb.Button(grid, text="选择", command=self._pick_xlsx,
                                       bootstyle=BS_OUTLINE_SECONDARY, width=8)
        self.btn_pick_xlsx.grid(row=1, column=2, sticky="e")

        ttk.Label(grid, text="输出目录", font=FONT_BODY, width=10, anchor="w").grid(
            row=2, column=0, sticky="w", pady=PAD_XS)
        self.out_var = tk.StringVar()
        self.out_entry = ttk.Entry(grid, textvariable=self.out_var, font=FONT_BODY)
        self.out_entry.grid(row=2, column=1, padx=(PAD_SM, PAD_SM), sticky="we")
        self.btn_pick_out = tb.Button(grid, text="选择", command=self._pick_out,
                                      bootstyle=BS_OUTLINE_SECONDARY, width=8)
        self.btn_pick_out.grid(row=2, column=2, sticky="e")

        # Stage 4：has_header 默认 True（已有 config 优先，见 _load_config）
        self.header_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(grid, text="第一行为表头", variable=self.header_var).grid(row=3, column=1, sticky="w", pady=(PAD_XS, 0))
        section_help(grid, "数据从第 2 行开始读取").grid(
            row=3, column=2, sticky="w")

        make_separator(parent, pady=PAD_XS)

        # ---- Section 二：数据字段（规格 6.5B 第 8 节） ----
        sec2 = ttk.Frame(parent)
        sec2.pack(fill="x")
        ttk.Label(sec2, text="数据字段", font=FONT_SECTION).pack(anchor="w")
        mf = ttk.Frame(sec2)
        mf.pack(fill="x", pady=(PAD_XS, 0))
        mf.columnconfigure(1, weight=1)
        mf.columnconfigure(3, weight=1)

        cols = [index_to_excel_column(i) for i in range(26)]  # A..Z（加载后按实际列数刷新）
        self.col_store_var = tk.StringVar(value="A")
        self.col_name_var = tk.StringVar(value="B")
        self.col_phone_var = tk.StringVar(value="D")
        self.col_role_var = tk.StringVar(value="C")
        self._column_labels = cols

        ttk.Label(mf, text="门店", font=FONT_BODY, width=10, anchor="w").grid(
            row=0, column=0, sticky="w", pady=PAD_XS)
        self.col_store_cb = ttk.Combobox(mf, textvariable=self.col_store_var, values=cols,
                                         width=12, state="readonly", font=FONT_BODY)
        self.col_store_cb.grid(row=0, column=1, padx=(PAD_SM, PAD_LG), sticky="w")
        ttk.Label(mf, text="姓名", font=FONT_BODY, width=10, anchor="w").grid(
            row=0, column=2, sticky="w", pady=PAD_XS)
        self.col_name_cb = ttk.Combobox(mf, textvariable=self.col_name_var, values=cols,
                                        width=12, state="readonly", font=FONT_BODY)
        self.col_name_cb.grid(row=0, column=3, padx=(PAD_SM, 0), sticky="w")
        ttk.Label(mf, text="电话", font=FONT_BODY, width=10, anchor="w").grid(
            row=1, column=0, sticky="w", pady=PAD_XS)
        self.col_phone_cb = ttk.Combobox(mf, textvariable=self.col_phone_var, values=cols,
                                         width=12, state="readonly", font=FONT_BODY)
        self.col_phone_cb.grid(row=1, column=1, padx=(PAD_SM, PAD_LG), sticky="w")
        ttk.Label(mf, text="销售顾问", font=FONT_BODY, width=10, anchor="w").grid(
            row=1, column=2, sticky="w", pady=PAD_XS)
        self.col_role_cb = ttk.Combobox(mf, textvariable=self.col_role_var,
                                        values=cols + ["（不替换）"], width=12,
                                        state="readonly", font=FONT_BODY)
        self.col_role_cb.grid(row=1, column=3, padx=(PAD_SM, 0), sticky="w")

        section_help(mf, "选「不替换」则保留 PSD 原文字").grid(
            row=2, column=1, columnspan=3, sticky="w")

        make_separator(parent, pady=PAD_XS)

        # ---- Section 三：输出分组 ----
        sec3 = ttk.Frame(parent)
        sec3.pack(fill="x")
        ttk.Label(sec3, text="输出分组", font=FONT_SECTION).pack(anchor="w")
        gf = ttk.Frame(sec3)
        gf.pack(fill="x", pady=(PAD_XS, 0))
        gf.columnconfigure(1, weight=1)

        # Stage 4.5：按 Excel 任意列创建输出子文件夹（group_output_enabled / group_output_column）
        self.group_output_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(gf, text="按字段创建子文件夹", variable=self.group_output_var).grid(row=0, column=0, sticky="w", pady=PAD_XS)
        ttk.Label(gf, text="分组字段", font=FONT_BODY, width=10, anchor="w").grid(
            row=0, column=1, sticky="w", pady=PAD_XS)
        self.group_col_var = tk.StringVar(value="A")
        self.group_col_cb = ttk.Combobox(gf, textvariable=self.group_col_var, values=["A"],
                                         width=14, state="readonly", font=FONT_BODY)
        self.group_col_cb.grid(row=0, column=2, sticky="w", padx=(PAD_SM, 0))
        section_help(gf, "示例：输出目录 / 门店名 / 001_张三.png").grid(
            row=1, column=1, columnspan=2, sticky="w")

        make_separator(parent, pady=PAD_XS)

        # ---- Section 四：文字图层 ----
        sec4 = ttk.Frame(parent)
        sec4.pack(fill="x")
        ttk.Label(sec4, text="文字图层", font=FONT_SECTION).pack(anchor="w")
        tf = ttk.Frame(sec4)
        tf.pack(fill="x", pady=(PAD_XS, 0))
        tf.columnconfigure(1, weight=1)
        tf.columnconfigure(3, weight=1)

        self.tm_name_var = tk.StringVar(value="（不替换）")
        self.tm_phone_var = tk.StringVar(value="（不替换）")
        self.tm_role_var = tk.StringVar(value="（不替换）")
        tm_opts = ["（不替换）"]

        ttk.Label(tf, text="姓名", font=FONT_BODY, width=10, anchor="w").grid(
            row=0, column=0, sticky="w", pady=PAD_XS)
        self.tm_name_cb = ttk.Combobox(tf, textvariable=self.tm_name_var, values=tm_opts,
                                       width=24, state="readonly", font=FONT_BODY)
        self.tm_name_cb.grid(row=0, column=1, padx=(PAD_SM, PAD_LG), sticky="w")
        ttk.Label(tf, text="电话", font=FONT_BODY, width=10, anchor="w").grid(
            row=0, column=2, sticky="w", pady=PAD_XS)
        self.tm_phone_cb = ttk.Combobox(tf, textvariable=self.tm_phone_var, values=tm_opts,
                                        width=24, state="readonly", font=FONT_BODY)
        self.tm_phone_cb.grid(row=0, column=3, padx=(PAD_SM, 0), sticky="w")
        ttk.Label(tf, text="销售顾问", font=FONT_BODY, width=10, anchor="w").grid(
            row=1, column=0, sticky="w", pady=PAD_XS)
        self.tm_role_cb = ttk.Combobox(tf, textvariable=self.tm_role_var, values=tm_opts,
                                       width=24, state="readonly", font=FONT_BODY)
        self.tm_role_cb.grid(row=1, column=1, padx=(PAD_SM, PAD_LG), sticky="w")

        section_help(tf, "图层名与默认不同时在此手动指定；没有对应图层保持「不替换」").grid(
            row=2, column=1, columnspan=3, sticky="w")

        make_separator(parent, pady=PAD_XS)

        # ---- 加载按钮区（规格 6.5B 第 20/21 节：主按钮 primary + 次要 outline） ----
        bf = ttk.Frame(parent)
        bf.pack(fill="x")
        self.btn_load = tb.Button(bf, text="加载并分析",
                                  command=self._load, bootstyle=BS_MAIN, padding=(PAD_LG, PAD_SM))
        self.btn_load.pack(side="left")
        # Stage 7.5：刷新图层 —— PSD 被外部修改（新增门店图层等）后强制从磁盘重读。
        # 普通「加载并分析」在 PS 中已打开旧版时会复用旧文档；刷新会先关闭旧文档重开。
        self.btn_refresh = tb.Button(bf, text="刷新图层", command=self._refresh_layers,
                                     bootstyle=BS_WARNING)
        self.btn_refresh.pack(side="left", padx=(PAD_SM, 0))
        self.btn_save_cfg = tb.Button(bf, text="保存配置", command=self._save_config,
                                      bootstyle=BS_OUTLINE_SECONDARY)
        self.btn_save_cfg.pack(side="left", padx=(PAD_SM, 0))
        # 文件数据状态反馈（规格 6.5B 第 11 节：轻量单行，secondary 文字）
        self.file_status = tb.Label(bf, text="", bootstyle="secondary",
                                    font=FONT_HELP, padding=(PAD_SM, 0))
        self.file_status.pack(side="left", padx=(PAD_SM, 0))

    def _build_tab_logo(self, parent):
        # 顶部说明行（规格 6.5B 第 12 节：一句话说明，secondary）
        info = tb.Label(parent,
                        text="选择参与生成的 Logo，并为每个门店指定对应图层",
                        bootstyle="secondary", font=FONT_HELP, padding=(PAD_XS, 2))
        info.pack(fill="x", pady=(0, PAD_XS))
        self.logo_info = info

        # 搜索行（规格 6.5B 第 15 节：placeholder 短文案，只过滤不修改映射；清空恢复）
        sbar = ttk.Frame(parent)
        sbar.pack(fill="x", pady=(0, PAD_SM))
        sbar.columnconfigure(1, weight=1)
        sbar.columnconfigure(3, weight=1)
        ttk.Label(sbar, text="搜索 Logo 图层", font=FONT_BODY).grid(
            row=0, column=0, sticky="w")
        self.logo_search_var = tk.StringVar()
        self.logo_search_var.trace_add("write", lambda *a: self._refresh_logo_filtered())
        self.logo_search_entry = ttk.Entry(sbar, textvariable=self.logo_search_var,
                                           font=FONT_BODY)
        self.logo_search_entry.grid(row=0, column=1, sticky="we", padx=(PAD_SM, PAD_LG))
        ttk.Label(sbar, text="搜索门店", font=FONT_BODY).grid(
            row=0, column=2, sticky="w")
        self.store_search_var = tk.StringVar()
        self.store_search_var.trace_add("write", lambda *a: self._refresh_store_filtered())
        self.store_search_entry = ttk.Entry(sbar, textvariable=self.store_search_var,
                                            font=FONT_BODY)
        self.store_search_entry.grid(row=0, column=3, sticky="we", padx=(PAD_SM, 0))

        # 左右分栏（规格 6.5B 第 12 节：保持双栏）
        paned = ttk.PanedWindow(parent, orient="horizontal")
        paned.pack(fill="both", expand=True, pady=(0, PAD_XS))

        left = ttk.Frame(paned)
        right = ttk.Frame(paned)
        paned.add(left, weight=1)
        paned.add(right, weight=1)

        # 左：Logo 图层清单（规格 6.5B：无边框 section 标题 + 滚动区）
        lsec = ttk.Frame(left)
        lsec.pack(fill="both", expand=True)
        ttk.Label(lsec, text="Logo 图层", font=FONT_SECTION).pack(anchor="w")
        self.logo_scroll = make_scrollable(lsec, height=0)
        self.logo_scroll.pack(fill="both", expand=True, pady=(PAD_XS, 0))
        self.logo_inner = self.logo_scroll.inner
        self.logo_checks = {}  # name -> BooleanVar
        # label -> 展示标记（selected/brand）徽标缓存，便于搜索刷新时更新
        self.logo_badges = {}

        # 右：门店 -> Logo 映射
        rsec = ttk.Frame(right)
        rsec.pack(fill="both", expand=True)
        ttk.Label(rsec, text="门店映射", font=FONT_SECTION).pack(anchor="w")

        # 固定显示（规格 6.5B 第 13 节：禁止内部术语 BRAND，改「固定显示」）
        bsec = ttk.Frame(rsec)
        bsec.pack(fill="x", pady=(PAD_XS, 0))
        ttk.Label(bsec, text="固定显示", font=FONT_SECTION_PLAIN).pack(anchor="w")
        self.brand_scroll = make_scrollable(bsec, height=92)
        self.brand_scroll.pack(fill="x", expand=False, pady=(PAD_XS, 0))
        self.brand_inner = self.brand_scroll.inner
        section_help(bsec, "勾选后将在每张封面中保持显示").pack(anchor="w", pady=(PAD_XS, 0))

        self.map_scroll = make_scrollable(rsec, height=0)
        self.map_scroll.pack(fill="both", expand=True, pady=(PAD_XS, 0))
        self.map_inner = self.map_scroll.inner
        self.map_combos = {}  # store -> StringVar
        self.map_combo_widgets = {}  # store -> Combobox widget
        self.map_badges = {}  # store -> tb.Label 状态徽标

    # ---- 搜索刷新（只过滤显示，不改数据；规格第 13/14 节） ----
    def _refresh_logo_filtered(self):
        """图层搜索：只改 View 中可见的勾选列表（不改变 logo_checks 数据）。"""
        query = self.logo_search_var.get() if hasattr(self, "logo_search_var") else ""
        labels = self.all_psd_layers_labels()
        display_path_of = {}
        if self.layer_index is not None:
            display_path_of = {self.layer_labels[r.id]: r.display_path
                               for r in self.layer_index.layers}
        visible = filter_logo_labels(labels, query, display_path_of)
        self._rebuild_logo_list_rows(visible)

    def _refresh_store_filtered(self):
        """门店搜索：只过滤显示（不删除 mapping 数据；规格第 14 节）。"""
        query = self.store_search_var.get() if hasattr(self, "store_search_var") else ""
        visible = filter_stores(self.excel_stores, query)
        self._rebuild_map_rows(visible)

    def _build_tab_run(self, parent):
        # ---- Section 一：输出格式（规格 6.5B 第 16 节） ----
        sec1 = ttk.Frame(parent)
        sec1.pack(fill="x")
        ttk.Label(sec1, text="输出格式", font=FONT_SECTION).pack(anchor="w")
        opt = ttk.Frame(sec1)
        opt.pack(fill="x", pady=(PAD_XS, 0))
        self.fmt_var = tk.StringVar(value=FMT_PNG)
        ttk.Label(opt, text="格式", font=FONT_BODY, width=10, anchor="w").grid(
            row=0, column=0, sticky="w")
        self.fmt_cb = ttk.Combobox(opt, textvariable=self.fmt_var,
                                   values=[FMT_PNG, FMT_JPG, FMT_PSD],
                                   width=10, state="readonly", font=FONT_BODY)
        self.fmt_cb.grid(row=0, column=1, sticky="w", padx=(PAD_SM, PAD_LG))
        self.also_png_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(opt, text="PSD 同时保存 PNG", variable=self.also_png_var).grid(row=0, column=2, sticky="w")

        make_separator(parent)

        # ---- Section 二：生成准备（配置检查 inline 列表，规格 6.5B 第 16 节） ----
        sec2 = ttk.Frame(parent)
        sec2.pack(fill="x")
        head2 = ttk.Frame(sec2)
        head2.pack(fill="x")
        ttk.Label(head2, text="生成准备", font=FONT_SECTION).pack(side="left")
        self.btn_check_cfg = tb.Button(head2, text="检查配置", command=self._check_config,
                                       bootstyle=BS_OUTLINE_SECONDARY)
        self.btn_check_cfg.pack(side="right")
        self.cfg_check_box = ttk.Frame(sec2)
        self.cfg_check_box.pack(fill="x", pady=(PAD_XS, 0))
        self._render_cfg_check_placeholder()

        make_separator(parent)

        # ---- Section 三：当前任务（进度，规格 6.5B 第 16/19 节） ----
        sec3 = ttk.Frame(parent)
        sec3.pack(fill="x")
        ttk.Label(sec3, text="当前任务", font=FONT_SECTION).pack(anchor="w")
        pf = ttk.Frame(sec3)
        pf.pack(fill="x", pady=(PAD_XS, 0))
        # 第一行：进度条（主题 Primary）+ 百分比 + 78/221
        row1 = ttk.Frame(pf)
        row1.pack(fill="x")
        self.progress = ttk.Progressbar(row1, mode="determinate", maximum=100,
                                        style="primary.Horizontal.TProgressbar")
        self.progress.pack(fill="x", side="left", expand=True)
        self.progress_pct = tb.Label(row1, text="0%", bootstyle="secondary",
                                     font=FONT_BODY, width=6)
        self.progress_pct.pack(side="left", padx=(PAD_SM, 0))
        self.progress_label = tb.Label(row1, text="0 / 0", bootstyle="primary",
                                       font=FONT_BODY)
        self.progress_label.pack(side="left", padx=(PAD_SM, 0))
        # 第二行：门店 · 姓名 + 阶段
        row2 = ttk.Frame(pf)
        row2.pack(fill="x", pady=(PAD_XS, 0))
        self.progress_cur = tb.Label(row2, text="", bootstyle="secondary", font=FONT_BODY)
        self.progress_cur.pack(side="left")
        self.progress_phase = tb.Label(row2, text="", bootstyle="info", font=FONT_BODY)
        self.progress_phase.pack(side="left", padx=(PAD_SM, 0))

        # ---- 按钮区（规格 6.5B 第 20/21 节：主按钮 primary / 停止 danger / 预览 outline） ----
        bf = ttk.Frame(parent)
        bf.pack(fill="x", pady=(PAD_MD, 0))
        self.btn_run = tb.Button(bf, text="开始生成", command=self._start,
                                 bootstyle=BS_MAIN,
                                 padding=(PAD_LG, PAD_SM))
        self.btn_run.pack(side="left")
        self.btn_preview = tb.Button(bf, text="生成预览", command=self._preview,
                                     bootstyle=BS_OUTLINE_SECONDARY)
        self.btn_preview.pack(side="left", padx=(PAD_SM, 0))
        self.btn_stop = tb.Button(bf, text="停止", command=self._stop,
                                  bootstyle=BS_DANGER, state="disabled")
        self.btn_stop.pack(side="left", padx=(PAD_SM, 0))

        make_separator(parent)

        # ---- Section 四：本次结果（空状态精简，规格 6.5B 第 17 节） ----
        sec4 = ttk.Frame(parent)
        sec4.pack(fill="x")
        ttk.Label(sec4, text="本次结果", font=FONT_SECTION).pack(anchor="w")
        self.summary_box = ttk.Frame(sec4)
        self.summary_box.pack(fill="x", pady=(PAD_XS, 0))
        self._render_summary_placeholder()

        # ---- Section 五：运行日志（规格 6.5B：允许保留的大区域框——日志文本区） ----
        sec5 = ttk.Frame(parent)
        sec5.pack(fill="both", expand=True, pady=(PAD_MD, 0))
        head5 = ttk.Frame(sec5)
        head5.pack(fill="x")
        ttk.Label(head5, text="运行日志", font=FONT_SECTION).pack(side="left")
        tbrow = ttk.Frame(head5)
        tbrow.pack(side="right")
        self.btn_log_clear = tb.Button(tbrow, text="清空", command=self._log_clear,
                                       bootstyle=BS_OUTLINE_SECONDARY)
        self.btn_log_clear.pack(side="left")
        self.btn_log_copy = tb.Button(tbrow, text="复制", command=self._log_copy,
                                      bootstyle=BS_OUTLINE_SECONDARY)
        self.btn_log_copy.pack(side="left", padx=(PAD_XS, 0))
        self.btn_log_open = tb.Button(tbrow, text="打开目录", command=self._log_open_dir,
                                      bootstyle=BS_OUTLINE_SECONDARY)
        self.btn_log_open.pack(side="left", padx=(PAD_XS, 0))
        self.log = scrolledtext.ScrolledText(sec5, height=10, state="disabled",
                                             font=FONT_LOG, relief="solid", borderwidth=1)
        self.log.pack(fill="both", expand=True, pady=(PAD_XS, 0))

    # ---------------- 文件选择 ----------------
    def _pick_psd(self):
        p = filedialog.askopenfilename(title="选择 PSD 模板",
                                       filetypes=[("PSD", "*.psd"), ("所有", "*.*")])
        if p:
            self.psd_var.set(p)

    def _pick_xlsx(self):
        # Stage 4：仅支持 .xlsx / .xlsm（.xls 在 load_excel_dataset 中明确拒绝）
        p = filedialog.askopenfilename(title="选择 Excel 数据",
                                       filetypes=[("Excel", "*.xlsx *.xlsm"), ("所有", "*.*")])
        if p:
            self.xlsx_var.set(p)
            # 默认输出目录
            if not self.out_var.get():
                d = os.path.splitext(os.path.basename(p))[0] + "_封面输出"
                self.out_var.set(os.path.join(os.path.dirname(p), d))

    def _pick_out(self):
        d = filedialog.askdirectory(title="选择输出目录")
        if d:
            self.out_var.set(d)

    # ---------------- 列名工具（Stage 4）----------------
    def _col_of(self, label):
        """GUI 下拉 label（'A'/'B'/.../'AA'、'B - 门店' 或 '（不替换）'）-> 0 起列索引。"""
        if label == "（不替换）":
            return -1
        # 兼容「B - 门店」格式（有表头时下拉显示列字母 - 表头名）
        if label and " - " in label:
            label = label.split(" - ")[0]
        return excel_column_to_index(label)

    def _set_column_options(self, max_columns):
        """按实际工作表最大列数动态生成列下拉（支持 A..Z/AA/AB...）。

        有表头时显示「列字母 - 表头名」格式（如 B - 门店）；无表头只显示列字母。
        """
        if max_columns <= 0:
            max_columns = 1
        cols = [index_to_excel_column(i) for i in range(max_columns)]
        ds = self.excel_dataset
        if ds is not None and ds.headers:
            labels = [f"{c} - {ds.headers[i]}" if ds.headers[i] else c
                      for i, c in enumerate(cols)]
        else:
            labels = cols
        self._column_labels = labels
        for cb in (self.col_store_cb, self.col_name_cb, self.col_phone_cb,
                   self.col_role_cb, self.group_col_cb):
            if cb is not None:
                cb["values"] = labels + (["（不替换）"] if cb is self.col_role_cb else [])

    # ---------------- 加载解析 ----------------
    def _ds_cache_key(self):
        """Dataset 一致性 key：path + has_header + 4 列配置。任一变化即失效。"""
        return (
            self.xlsx_var.get().strip(),
            bool(self.header_var.get()),
            self._col_of(self.col_store_var.get()),
            self._col_of(self.col_name_var.get()),
            self._col_of(self.col_phone_var.get()),
            self._col_of(self.col_role_var.get()),
        )

    def _load_excel_data(self, xlsx=None):
        """仅加载 Excel -> 更新 self.excel_dataset / excel_stores / 列下拉。

        返回 True 成功；False 失败（已弹窗）。供 _load 与 _ensure_dataset_fresh 共用。
        """
        if xlsx is None:
            xlsx = self.xlsx_var.get().strip()
        try:
            ds = load_excel_dataset(
                xlsx,
                has_header=self.header_var.get(),
                col_store=self._col_of(self.col_store_var.get()),
                col_name=self._col_of(self.col_name_var.get()),
                col_phone=self._col_of(self.col_phone_var.get()),
                col_role=self._col_of(self.col_role_var.get()),
            )
            self.excel_dataset = ds
            self._ds_key = self._ds_cache_key()
            # 动态生成列下拉（按实际最大列数，支持 A..Z/AA/AB...）
            self._set_column_options(ds.max_columns)
            self.excel_headers = [index_to_excel_column(i) for i in range(ds.max_columns)]
            self.excel_stores = ds.stores
            return True
        except ExcelDataError as e:
            messagebox.showerror("Excel 错误", str(e))
            return False
        except Exception as e:
            messagebox.showerror("Excel 错误", str(e))
            return False

    def _ensure_dataset_fresh(self, trigger="start"):
        """Stage 4 补充（BLOCKED B/C）：开始/预览前保证 Dataset 与当前控件配置一致。

        若 path / has_header / col_store / col_name / col_phone / col_role 任一
        与上次加载不一致 -> 自动重解析 Excel（不必等用户重新点「加载」），
        并同步重建门店->Logo 映射区（避免 Logo 页旧门店列表 + Batch 新列错配）。
        返回 True 可继续；False 已弹窗报错。
        """
        if self.excel_dataset is None:
            return True   # 未加载过（_start/_preview 要求先点加载，这里兜底）
        cur = self._ds_cache_key()
        if cur == self._ds_key:
            return True
        # 配置变化 -> 重新解析 Excel
        self._log_gui(f"[Excel 重解析] 字段映射/表头已变化（{trigger}前自动刷新），正在重新读取 Excel ...")
        if not self._load_excel_data():
            return False
        # 若 PSD 已加载（有 layer_index），同步重建门店映射区
        if self.layer_index is not None and self.map_combos:
            self._rebuild_store_map_ui()
        return True

    def _rebuild_store_map_ui(self):
        """按当前 excel_stores 重建门店->Logo 映射下拉区。

        - 保留旧门店（同名）已选映射值（用户人工选择不丢失）；
        - 新增门店：自动 match_store_logo（EXACT/AUTO 唯一命中才选，歧义/无匹配 -> 无）。
        - 消失的门店：其映射自然移除（不再是 stores 一员）。
        """
        stores = self.excel_stores
        prev_combos = self.map_combos
        prev_map = {}
        for s, var in prev_combos.items():
            prev_map[s] = var.get()
        self.map_combos = {}
        selected_refs = self._selected_logo_refs()
        try:
            effective_leaves = resolve_effective_logo_layers(self.layer_index, selected_refs)
        except Exception:
            effective_leaves = []
        for s in stores:
            default_label = prev_map.get(s, "（无）")
            if default_label != "（无）" and self._label_to_ref(default_label) is not None:
                pass   # 保留旧值
            else:
                mr = match_store_logo(s, effective_leaves)
                if mr.status in (EXACT, AUTO) and mr.best is not None:
                    default_label = self.layer_labels[mr.best.id]
                else:
                    default_label = "（无）"
            self.map_combos[s] = tk.StringVar(value=default_label)
        self._rebuild_logo_lists()
        self._log_gui(f"门店映射已按新 Excel 列刷新：{len(stores)} 个门店。")

    def _refresh_layers(self):
        """Stage 7.5：刷新图层 —— PSD 被外部修改后强制从磁盘重读。

        与「加载并分析」的区别：force_reload=True 会先关闭 Photoshop 中已打开的
        旧版本文档再重新 Open，确保拿到磁盘最新图层（新增门店/文字层等）。
        """
        psd = self.psd_var.get().strip()
        if not psd or not os.path.exists(psd):
            messagebox.showerror("错误", "请先选择有效的 PSD 模板文件。")
            return
        if not self.layer_index:
            # 尚未加载过：等同首次加载（无需强制重读，正常加载即可）
            self._load(force_reload=False)
            return
        # 已加载过：强制重读磁盘最新版本
        self._log_gui("[刷新图层] 正在强制从磁盘重新读取 PSD（关闭旧文档并重开）...")
        self._load(force_reload=True)

    def _load(self, force_reload=False):
        """Stage 6: Load into worker. main thread only collects snapshot + starts worker.

        force_reload=True（Stage 7.5）：PSD 文件已在 Photoshop 中打开过（可能被
        外部修改新增图层），强制关闭旧文档并从磁盘重新读取最新版本。
        """
        from core.task_events import AppState
        psd = self.psd_var.get().strip()
        xlsx = self.xlsx_var.get().strip()
        if not psd or not os.path.exists(psd):
            messagebox.showerror("错误", "请先选择有效的 PSD 模板文件。")
            return
        if not xlsx or not os.path.exists(xlsx):
            messagebox.showerror("错误", "请先选择有效的 Excel 数据文件。")
            return
        if self._task_worker.worker_alive:
            messagebox.showwarning("提示", "已有任务正在运行，请稍候。")
            return
        cfg = self._snapshot_for_load(force_reload=force_reload)
        self._set_state(AppState.LOADING)
        try:
            self._task_worker.run_load(cfg)
        except Exception as e:
            self._log_gui(f"启动加载任务失败：{e}")
            self._set_state(AppState.IDLE)

    def _on_load_done(self, result):
        """Load done (main thread): rebuild GUI from pure data (no COM objects)."""
        from core.task_events import AppState
        from core.layer_index import LayerIndex, ref_from_config
        if not result.ok:
            messagebox.showerror("加载错误", result.error or "加载失败。")
            self._set_state(AppState.IDLE)
            return
        # Stage 7.5：成功 Load 后记录 PSD 磁盘指纹（检测外部修改用）
        self._psd_fingerprint = (self.psd_var.get().strip(),
                                 self._psd_fingerprint_of(self.psd_var.get().strip()))
        refs = [ref_from_config(d) for d in result.layer_refs]
        refs = [r for r in refs if r is not None]
        index = LayerIndex(refs)
        self.layer_index = index
        self._ref_by_id = {r.id: r for r in index.layers}
        self.layer_labels = index.labels()
        self.all_psd_layers = [r.name for r in index.layers]
        self.all_psd_is_group = {r.id: r.is_group for r in index.layers}
        self.all_psd_parent = {r.id: self._parent_name_of(r) for r in index.layers}
        self.all_text_layers = [r.display_path for r in index.layers if r.is_text]
        self.text_label_to_ref = {}
        for r in index.layers:
            if r.is_text:
                self.text_label_to_ref[self.layer_labels[r.id]] = r
        self.logo_label_to_ref = {}
        self.excel_stores = result.store_logo_defaults.get("_stores", []) or self.excel_stores
        opts = ["（不替换）"] + [self.layer_labels[r.id] for r in index.layers if r.is_text]
        self.tm_name_cb["values"] = opts
        self.tm_phone_cb["values"] = opts
        self.tm_role_cb["values"] = opts
        tdef = result.text_defaults or {}
        self.tm_name_var.set(self._label_of_text(tdef.get("姓名")) or "（不替换）")
        self.tm_phone_var.set(self._label_of_text(tdef.get("电话")) or "（不替换）")
        self.tm_role_var.set(self._label_of_text(tdef.get("销售顾问")) or "（不替换）")
        self._init_logo_checks(index)
        self._rebuild_logo_lists()
        # Stage 6.5 修复（基线遗留）：Load 完成后同步构建 GUI 侧 excel_dataset，
        # 使 _start/_preview 的分组预检能拿到 max_columns（此前始终 None -> 分组被误拦截）。
        # worker 已验证 xlsx 有效；此处直接 load_excel_dataset 同步读取（不弹窗，静默兜底）。
        try:
            if self.excel_dataset is None and os.path.exists(
                    self.xlsx_var.get().strip()):
                ds = load_excel_dataset(
                    self.xlsx_var.get().strip(),
                    has_header=self.header_var.get(),
                    col_store=self._col_of(self.col_store_var.get()),
                    col_name=self._col_of(self.col_name_var.get()),
                    col_phone=self._col_of(self.col_phone_var.get()),
                    col_role=self._col_of(self.col_role_var.get()),
                )
                self.excel_dataset = ds
                self._ds_key = self._ds_cache_key()
                self._set_column_options(ds.max_columns)
                self.excel_headers = [index_to_excel_column(i)
                                      for i in range(ds.max_columns)]
                self.excel_stores = ds.stores
        except Exception:
            pass
        stores = self.excel_stores
        self._log_gui(f"解析完成：PSD 共 {len(index)} 个图层，Excel 共 {len(stores)} 个门店。")
        # 规格 6.5B 第 11 节：轻量单行状态（不拆多个冗长提示）
        try:
            self.file_status.config(
                text=f"✓ Excel 已读取 · {len(stores)} 条数据 · {len(set(stores))} 个门店")
        except Exception:
            pass
        self.logo_info.config(text="请在下方勾选 Logo 并为每个门店指定对应图层")
        self._set_state(AppState.READY)

    def _label_of_text(self, display_path):
        """find label by display_path for text dropdown auto-suggestion."""
        if not display_path or self.layer_index is None:
            return ""
        for r in self.layer_index.layers:
            if r.display_path == display_path:
                return self.layer_labels.get(r.id, r.display_path)
        return ""

    def _init_logo_checks(self, index):
        """first-load recommended check + store->logo auto match (equivalent to v1.1.0 _load)."""
        from core.logo_mapping import (
            recommend_logo_selection, match_store_logo, EXACT, AUTO)
        stores = self.excel_stores

        def is_logo_candidate(ref):
            if ref.is_group:
                return False
            if "logo" in ref.name.lower():
                return True
            parent = (self._parent_name_of(ref) or "").lower()
            if "logo" in parent:
                return True
            return ref.name in stores

        saved_items_raw = self.cfg.get("logo_selection")
        if saved_items_raw is None:
            saved_items_raw = self.cfg.get("logo_layers", [])
        has_saved = bool(saved_items_raw)
        saved_logo_ids = set()
        saved_names = set()
        for x in saved_items_raw or []:
            if isinstance(x, str):
                saved_names.add(x)
            else:
                rf = ref_from_config(x)
                if rf is not None and rf.id:
                    saved_logo_ids.add(rf.id)
        recommend = []
        if not has_saved:
            recommend = recommend_logo_selection(
                index.layers, stores, is_logo_heuristic=is_logo_candidate)
        self.logo_checks = {}
        for ref in index.layers:
            label = self.layer_labels[ref.id]
            checked = (ref.id in saved_logo_ids or ref.name in saved_names
                       or (not has_saved and ref.id in {r.id for r in recommend}))
            self.logo_checks[label] = tk.BooleanVar(value=checked)
        self.map_combos = {}
        selected = self._selected_logo_refs()
        eff = resolve_effective_logo_layers(index, selected)
        prev_map = self.cfg.get("store_logo_map", {})
        for s in stores:
            default_label = "（无）"
            if s in prev_map:
                prev = ref_from_config(prev_map[s])
                if prev is not None and prev.id and prev.id in self._ref_by_id:
                    default_label = self.layer_labels[prev.id]
            else:
                mr = match_store_logo(s, eff)
                if mr.status in (EXACT, AUTO) and mr.best is not None:
                    default_label = self.layer_labels[mr.best.id]
            self.map_combos[s] = tk.StringVar(value=default_label)
        self._init_brand_checks(index, eff)

    def _init_brand_checks(self, index, eff):
        """brand logo initial checks (suggest_brand_logos or saved config)."""
        from core.logo_mapping import suggest_brand_logos
        saved_brand = self.cfg.get("brand_logo_layers")
        self.brand_checks = {}
        if saved_brand:
            saved_ids = set()
            for v in saved_brand:
                rf = ref_from_config(v)
                if rf is not None and rf.id and rf.id in self._ref_by_id:
                    saved_ids.add(rf.id)
            for label in self._effective_logo_layers():
                ref = self._label_to_ref(label)
                self.brand_checks[label] = tk.BooleanVar(
                    value=ref is not None and ref.id in saved_ids)
        else:
            store_map_now = {}
            for s, var in self.map_combos.items():
                v = var.get()
                store_map_now[s] = self._label_to_ref(v) if v != "（无）" else None
            try:
                suggest = suggest_brand_logos(eff, store_map_now)
            except Exception:
                suggest = []
            suggest_ids = {r.id for r in suggest}
            for label in self._effective_logo_layers():
                ref = self._label_to_ref(label)
                self.brand_checks[label] = tk.BooleanVar(
                    value=ref is not None and ref.id in suggest_ids)

    def _parent_name_of(self, ref):
        """返回 LayerRef 的父组名（display_path 倒数第二段）。"""
        if not ref or not ref.display_path:
            return ""
        parts = [p.strip() for p in ref.display_path.split(">")]
        return parts[-2] if len(parts) >= 2 else ""

    def _parent_name_of_str(self, name):
        """兼容：按 name 找 ref 再取父名（旧逻辑用）。"""
        if self.layer_index is None:
            return ""
        m = self.layer_index.find_matching(name)
        if len(m) == 1:
            return self._parent_name_of(m[0])
        return ""

    def _label_to_ref(self, label):
        """把展示 label（display_path 或 display_path [id=...]）解析回 LayerRef。"""
        if self.layer_index is None:
            return None
        if not label:
            return None
        # 优先精确 id 后缀
        for r in self.layer_index.layers:
            if self.layer_labels.get(r.id) == label:
                return r
        # 退化为 display_path 匹配
        for r in self.layer_index.layers:
            if r.display_path == label:
                return r
        return None

    def _build_children_map(self):
        """根据 all_psd_parent 构建 父名 -> [子名,...] 映射（兼容旧逻辑）。"""
        children = {}
        for rid, parent in self.all_psd_parent.items():
            if parent:
                children.setdefault(parent, []).append(rid)
        return children

    def _effective_logo_layers(self):
        """返回当前应作为候选 Logo 的『叶子图层 label 列表』（含被勾选组的子图层）。

        Stage 3：走 core.logo_mapping.resolve_effective_logo_layers（纯函数、按 id 去重、
        组递归展开、组本身不进结果）。
        """
        selected = self._selected_logo_refs()
        eff = resolve_effective_logo_layers(self.layer_index, selected)
        result = []
        for r in eff:
            label = self.layer_labels.get(r.id)
            if label and label not in result:
                result.append(label)
        return result

    def _selected_logo_refs(self):
        """当前 GUI 勾选的 Logo 图层（含组）-> LayerRef 列表（Stage 3 selection）。"""
        out = []
        seen = set()
        for label, var in self.logo_checks.items():
            if var.get():
                ref = self._label_to_ref(label)
                if ref is not None and ref.id not in seen:
                    seen.add(ref.id)
                    out.append(ref)
        return out

    def _rebuild_logo_lists(self):
        """全量重建 Logo 页（Stage 6.5：搜索区保留，列表行走拆分渲染）。"""
        self.map_combo_widgets = {}
        self.brand_widgets = {}
        self.logo_badges = {}
        self.map_badges = {}
        self._rebuild_brand_list()
        self._refresh_logo_filtered()
        self._refresh_store_filtered()

    # ---- Logo 图层行渲染（规格第 13/16/17 节） ----
    def _rebuild_logo_list_rows(self, visible_labels):
        """按 visible_labels 重建左侧 Logo 勾选列表（含 selected/brand 徽标 + 完整路径）。"""
        for w in self.logo_inner.winfo_children():
            w.destroy()
        self.logo_badges = {}
        ttk.Label(self.logo_inner, text="☑ 图层（完整路径）",
                  font=FONT_SECTION_PLAIN).grid(
            row=0, column=0, columnspan=3, sticky="w", padx=PAD_XS, pady=PAD_XS)
        if not visible_labels:
            tb.Label(self.logo_inner, text="（无匹配图层）", bootstyle="secondary",
                     font=FONT_HELP).grid(
                row=1, column=0, columnspan=3, sticky="w", padx=PAD_XS, pady=2)
            return
        for i, label in enumerate(visible_labels, start=1):
            var = self.logo_checks.get(label)
            if var is None:
                var = tk.BooleanVar(value=False)
                self.logo_checks[label] = var
            # selected vs brand 视觉区分（规格 6.5B 第 13 节：固定/参与/候选）
            selected = bool(var.get())
            is_brand = bool((self.brand_checks.get(label) or tk.BooleanVar(value=False)).get())
            disp = logo_display_state(label, selected, is_brand)
            if is_brand:
                badge_bs = BS_PRIMARY
            elif selected:
                badge_bs = BS_SUCCESS
            else:
                badge_bs = BS_SECONDARY
            # 完整路径：label + 父路径（缩短显示，tooltip 全路径）
            cb = ttk.Checkbutton(self.logo_inner, text=label, variable=var)
            cb.grid(row=i, column=0, sticky="w", padx=PAD_XS, pady=1)
            badge = tb.Label(self.logo_inner, text=disp, bootstyle=badge_bs,
                             font=FONT_HELP, padding=(PAD_XS, 0))
            badge.grid(row=i, column=1, sticky="w", padx=(0, PAD_XS))
            self.logo_badges[label] = badge
            # 完整路径：优先显示 display_path（label 已是最终形式，路径给 tooltip）
            full = ""
            if self.layer_index is not None and label in self.logo_label_to_ref:
                ref = self.logo_label_to_ref[label]
                full = getattr(ref, "display_path", "") or ""
            elif self.layer_index is not None:
                for r in self.layer_index.layers:
                    if self.layer_labels.get(r.id) == label:
                        full = getattr(r, "display_path", "") or ""
                        break
            if full:
                path_lbl = ttk.Label(self.logo_inner, text=shorten_path(full, 52),
                                     foreground="#888888", font=FONT_HELP)
                path_lbl.grid(row=i, column=2, sticky="w", padx=(0, PAD_XS))
                add_tooltip(path_lbl, full)
            else:
                ttk.Label(self.logo_inner, text="", font=FONT_BODY).grid(
                    row=i, column=2, sticky="w")
            # 勾选状态变化 -> 同步右侧映射 + 刷新徽标（trace 去重：避免搜索重建累积）
            if not self._has_trace(self.logo_checks[label], "write"):
                self.logo_checks[label].trace_add(
                    "write", lambda *a: self._on_logo_checks_changed())

    @staticmethod
    def _has_trace(var, mode):
        """StringVar.trace_info 判重。

        trace_info() 返回 [( (mode,), callback_id ), ...]（mode 可能是 'write'/'w'）。
        """
        try:
            infos = var.trace_info()
            for modes, _cid in infos:
                for m in modes:
                    if m == mode or m == mode[0]:
                        return True
            return False
        except Exception:
            return False

    # ---- 门店映射行渲染（规格第 14/15 节） ----
    def _rebuild_map_rows(self, visible_stores):
        """按 visible_stores 重建右侧门店→Logo 下拉 + 状态徽标（不删除数据）。"""
        for w in self.map_inner.winfo_children():
            w.destroy()
        self.map_badges = {}
        ttk.Label(self.map_inner, text="门店", font=FONT_SECTION_PLAIN).grid(
            row=0, column=0, sticky="w", padx=PAD_XS, pady=PAD_XS)
        ttk.Label(self.map_inner, text="→ Logo 图层", font=FONT_SECTION_PLAIN).grid(
            row=0, column=1, sticky="w", padx=PAD_XS, pady=PAD_XS)
        ttk.Label(self.map_inner, text="状态", font=FONT_SECTION_PLAIN).grid(
            row=0, column=2, sticky="w", padx=PAD_XS, pady=PAD_XS)
        if not visible_stores:
            tb.Label(self.map_inner, text="（无匹配门店）", bootstyle="secondary",
                     font=FONT_HELP).grid(
                row=1, column=0, columnspan=3, sticky="w", padx=PAD_XS, pady=2)
            return
        logo_opts = ["（无）"] + self._effective_logo_layers()
        for i, s in enumerate(visible_stores, start=1):
            ttk.Label(self.map_inner, text=s, font=FONT_BODY).grid(
                row=i, column=0, sticky="w", padx=PAD_XS, pady=1)
            var = self.map_combos.get(s)
            if var is None:
                var = tk.StringVar(value="（无）")
                self.map_combos[s] = var
            cb = ttk.Combobox(self.map_inner, textvariable=var, values=logo_opts,
                              width=32, state="readonly", font=FONT_BODY)
            cb.grid(row=i, column=1, sticky="w", padx=PAD_XS, pady=1)
            self.map_combo_widgets[s] = cb
            # 状态徽标（规格 6.5B 第 14 节：✓已匹配 / ●手动 / ⚠待确认 / ✕未映射）
            badge = self._make_map_badge(s, var.get())
            badge.grid(row=i, column=2, sticky="w", padx=(0, PAD_XS))
            self.map_badges[s] = badge
            # 下拉变化 -> 刷新徽标 + 品牌冲突联动（trace 去重：避免搜索重建累积）
            if not self._has_trace(var, "write"):
                var.trace_add("write", lambda *a, st=s: self._on_map_changed(st))

    def _make_map_badge(self, store, mapped_label):
        """根据映射状态生成徽标（规格 6.5B 第 14 节）。auto_matched 由 core 自动匹配判定。"""
        auto_matched = self._is_auto_matched(store, mapped_label)
        text, bs = mapping_status(store, mapped_label, auto_matched)
        return tb.Label(self.map_inner, text=text, bootstyle=bs,
                        font=FONT_HELP, padding=(PAD_XS, 0))

    def _is_auto_matched(self, store, mapped_label):
        """判定该门店映射是否来自自动匹配（EXACT/AUTO）。"""
        if not mapped_label or mapped_label == "（无）" or self.layer_index is None:
            return False
        try:
            from core.logo_mapping import match_store_logo
            eff = self._effective_logo_refs_internal()
            mr = match_store_logo(store, eff)
            return mr.status in ("exact", "auto") and mr.best is not None \
                and self.layer_labels.get(mr.best.id) == mapped_label
        except Exception:
            return False

    def _effective_logo_refs_internal(self):
        """当前 effective 叶子 LayerRef 列表（供自动匹配判定）。"""
        out = []
        for label in self._effective_logo_layers():
            ref = self._label_to_ref(label)
            if ref is not None:
                out.append(ref)
        return out

    def _on_map_changed(self, store):
        """门店下拉变化：刷新徽标 + 清除该 leaf 的 brand 冲突提示。"""
        mapped = self.map_combos[store].get() if store in self.map_combos else "（无）"
        if store in self.map_badges:
            badge = self._make_map_badge(store, mapped)
            self.map_badges[store].grid_forget()
            badge.grid(row=self._map_row_of(store), column=2, sticky="w", padx=(0, PAD_XS))
            self.map_badges[store] = badge
        # 若该 leaf 已被 brand 指定 -> 提示冲突（与 _on_brand_toggle 对称）
        if mapped and mapped != "（无）":
            for label, bvar in self.brand_checks.items():
                if bvar.get() and label == mapped:
                    self._log_gui(f"注意：图层「{mapped}」同时被选为门店 Logo 与固定品牌。")

    def _map_row_of(self, store):
        """当前可见门店列表中的行号（用于徽标刷新定位）。"""
        for i, s in enumerate(self.excel_stores, start=1):
            if s == store:
                return i
        return 1

    # ---- 固定显示区渲染（规格 6.5B 第 13 节：禁 BRAND 术语，改「固定显示」） ----
    def _rebuild_brand_list(self):
        """重建固定显示 Logo 勾选区。

        可选项 = 当前 effective 叶子（勾选的候选集合）；勾选 = 该 leaf 是固定显示。
        语义：selected 只是候选集合，不等于固定显示；这里人工勾选才真正写入 brand。
        """
        for w in self.brand_inner.winfo_children():
            w.destroy()
        self.brand_widgets = {}
        eff_labels = self._effective_logo_layers()
        if not eff_labels:
            ttk.Label(self.brand_inner,
                      text="（先勾选左侧 Logo 图层后，这里可指定固定显示）",
                      foreground="#888888", font=FONT_HELP).grid(
                row=0, column=0, sticky="w", padx=PAD_XS, pady=2)
            return
        ttk.Label(self.brand_inner, text="☑ 固定显示（叶子）",
                  font=FONT_SECTION_PLAIN).grid(
            row=0, column=0, sticky="w", padx=PAD_XS, pady=2)
        for i, label in enumerate(eff_labels, start=1):
            # 该 leaf 若已被某门店映射为 store target，标记（冲突提示用）
            store_owner = self._store_owner_of_leaf(label)
            suffix = f"（门店:{store_owner}）" if store_owner else ""
            var = self.brand_checks.get(label, tk.BooleanVar(value=False))
            var.set(var.get())   # 保持既有状态
            cb = ttk.Checkbutton(
                self.brand_inner, text=label + suffix, variable=var,
                command=lambda lb=label: self._on_brand_toggle(lb))
            cb.grid(row=i, column=0, sticky="w", padx=PAD_XS, pady=1)
            self.brand_widgets[label] = cb
            self.brand_checks[label] = var

    def _store_owner_of_leaf(self, label):
        """返回把该 leaf 选为门店 Logo 的门店名（无则空）。冲突提示用。"""
        ref = self._label_to_ref(label)
        if ref is None:
            return ""
        for s, var in self.map_combos.items():
            if var.get() == label:
                return s
            # 兼容：下拉值可能是 ref 序列化 label（含 id 后缀）
            if var.get() and self._label_to_ref(var.get()) is not None \
                    and self._label_to_ref(var.get()).id == ref.id:
                return s
        return ""

    def _on_brand_toggle(self, label):
        """品牌勾选变化：同 leaf 若已是 store target -> 冲突提示并回滚。

        规则（用户要求）：
          - 同一 LayerRef 不能同时是 store target 与 brand；
          - 勾选时若发现已是某门店 store target，立即弹提示并恢复未勾选。
        """
        var = self.brand_checks.get(label)
        if var is None:
            return
        if not var.get():
            return   # 取消勾选：合法，直接放行
        # 勾选动作：检查冲突
        owner = self._store_owner_of_leaf(label)
        if owner:
            var.set(False)   # 回滚
            messagebox.showwarning(
                "品牌 Logo 冲突",
                f"图层「{label}」已被门店「{owner}」选为门店 Logo，"
                f"不能同时作为固定品牌 Logo。\n请先在右侧门店映射中把它改为其他图层。")
            return
        self._log_gui(f"品牌 Logo 已指定：{label}")

    def _on_logo_checks_changed(self):
        """左侧勾选 Logo 图层变动时，实时同步右侧门店→Logo 下拉的可选项。

        Stage 3：若某个门店当前映射的叶子已被取消勾选（不再在 effective 集合内），
        该映射自动重置为「（无）」并提示（绝不静默继续使用失效映射）。
        品牌勾选同样联动：被取消勾选的叶子不再是候选（其 brand 标记一并清除并提示）。
        """
        logo_opts = ["（无）"] + self._effective_logo_layers()
        eff_ids = {self._label_to_ref(lb).id for lb in logo_opts[1:]}
        for s, var in self.map_combos.items():
            w = self.map_combo_widgets.get(s)
            if w is None:
                continue
            cur = var.get()
            if cur not in logo_opts:
                if cur != "（无）":
                    self._log_gui(f"“{s}”的原 Logo 映射已不在当前选择范围，已清除。")
                var.set("（无）")
            w["values"] = logo_opts
        # 品牌：已勾选但不再在 effective 的叶子 -> 清除 brand 标记并提示
        for label, var in list(self.brand_checks.items()):
            if var.get() and label not in logo_opts[1:]:
                var.set(False)
                self._log_gui(f"品牌 Logo「{label}」已不在当前选择范围，已清除。")

    # ---------------- 运行 ----------------
    def _collect_cfg(self):
        # Stage 4：列下拉可能是「B - 门店」格式，_col_of 内部已 normalize
        col_of = lambda v: self._col_of(v)
        # 保存用户显式勾选的图层（含组），用于下次加载时恢复
        # Stage 2：以 LayerRef dict 保存（layer_id/name/display_path），并保留 name 供旧版兼容
        def ref_to_cfg(label):
            ref = self._label_to_ref(label)
            if ref is None:
                return label   # 旧 name（极端情况）
            return serialize_ref(ref)

        logo_layers = [ref_to_cfg(label)
                       for label in self.all_psd_layers_labels()
                       if self.logo_checks.get(label, tk.BooleanVar()).get()]
        store_logo_map = {}
        for s, var in self.map_combos.items():
            val = var.get()
            store_logo_map[s] = "" if val == "（无）" else ref_to_cfg(val)
        # Stage 3：品牌 Logo 明确保存（叶子 LayerRef）。
        # 以「固定品牌 Logo」区块的人工勾选为准（brand_checks）——selected 只是候选，
        # 人工勾选才真正写入 brand_logo_refs；首次加载的 suggest_brand_logos 建议
        # 会在 _load 中预填这些勾选（仅默认建议，不是唯一入口）。
        # 运行时只读取这里保存的 brand_logo_layers，不再每行 name 猜测。
        brand_logo_layers = []
        for label, var in self.brand_checks.items():
            if var.get():
                ref = self._label_to_ref(label)
                if ref is not None:
                    brand_logo_layers.append(serialize_ref(ref))
        def tm_val(var):
            v = var.get()
            if v == "（不替换）":
                return ""
            return ref_to_cfg(v)
        text_map = {
            "姓名": tm_val(self.tm_name_var),
            "电话": tm_val(self.tm_phone_var),
            "销售顾问": tm_val(self.tm_role_var),
        }
        return {
            "psd_path": self.psd_var.get().strip(),
            "xlsx_path": self.xlsx_var.get().strip(),
            "out_dir": self.out_var.get().strip(),
            "has_header": self.header_var.get(),
            "col_store": col_of(self.col_store_var.get()),
            "col_name": col_of(self.col_name_var.get()),
            "col_phone": col_of(self.col_phone_var.get()),
            "col_role": col_of(self.col_role_var.get()),
            # Stage 4.5：按 Excel 任意列分组输出（内部保存 0-based index）
            "group_output_enabled": bool(self.group_output_var.get()),
            "group_output_column": col_of(self.group_col_var.get()) if self.group_output_var.get() else None,
            "text_map": text_map,
            # Stage 3 新字段（运行时唯一来源）
            "logo_selection": [ref_to_cfg(label)
                               for label in self.all_psd_layers_labels()
                               if self.logo_checks.get(label, tk.BooleanVar()).get()],
            "store_logo_map": store_logo_map,
            "brand_logo_layers": brand_logo_layers,
            # 旧字段保留（兼容旧版读取）
            "logo_layers": logo_layers,
            "fmt": self.fmt_var.get(),
            "also_png": self.also_png_var.get(),
            # Stage 7.5：PSD 已被外部修改且用户确认继续时，batch/preview 也强制
            # 从磁盘重读最新 PSD（否则 open_document 会复用 PS 中旧版本文档）。
            "psd_force_reload": bool(self._psd_changed_since_load()),
        }

    def all_psd_layers_labels(self):
        """当前 PSD 全部图层的展示 label（与 logo_checks 的 key 一致）。"""
        if self.layer_index is not None:
            return [self.layer_labels[r.id] for r in self.layer_index.layers]
        return list(self.all_psd_layers)

    def _ensure_psd_fresh(self, trigger="开始"):
        """Stage 7.5：PSD 是否被外部修改（新增图层）且 GUI 仍显示旧数据。

        返回 True 可继续；False 时已提示用户「先刷新图层」（不自动弹错）。
        """
        if not self._psd_changed_since_load():
            return True
        self._log_gui(f"[PSD 变更检测] PSD 文件自上次加载后已被外部修改"
                      f"（新增图层？），当前界面仍显示旧图层数据。"
                      f"建议先点击「刷新图层」重新读取，再{trigger}。")
        # 非阻塞提示：继续可能漏掉新图层，但由用户决定
        ret = messagebox.askyesno(
            "PSD 已变更",
            f"PSD 模板文件自上次加载后已被修改（可能新增了门店/文字图层），\n"
            f"当前界面仍显示旧图层数据。\n\n"
            f"是否现在刷新图层？\n"
            f"（选「否」将按当前旧图层继续{trigger}，新增图层不会被识别）")
        if ret:
            self._refresh_layers()
            return False   # 已启动刷新任务，本次{trigger}取消
        return True        # 用户选择不刷新，继续（按旧图层）

    def _start(self):
        """Stage 6: Batch into worker (single-worker constraint)."""
        from core.task_events import AppState
        if self._state in (AppState.RUNNING, AppState.STOPPING, AppState.LOADING,
                           AppState.PREVIEWING):
            messagebox.showwarning("提示", "已有任务正在运行，请稍候。")
            return
        if not self._ensure_dataset_fresh(trigger="开始"):
            return
        if not self._ensure_psd_fresh(trigger="开始"):
            return
        cfg = self._collect_cfg()
        if not cfg["psd_path"] or not os.path.exists(cfg["psd_path"]):
            messagebox.showerror("错误", "请先选择 PSD 模板。")
            return
        if not cfg["xlsx_path"] or not os.path.exists(cfg["xlsx_path"]):
            messagebox.showerror("错误", "请先选择 Excel 数据。")
            return
        if not cfg["out_dir"]:
            messagebox.showerror("错误", "请先选择输出目录。")
            return
        if cfg.get("group_output_enabled"):
            try:
                assert_group_column_valid(
                    self.excel_dataset.max_columns if self.excel_dataset is not None else 0,
                    True, cfg.get("group_output_column"))
            except OutputPathError as e:
                messagebox.showerror("分组配置错误", str(e))
                return
        if self._task_worker.worker_alive:
            messagebox.showwarning("提示", "已有任务正在运行，请稍候。")
            return
        self.progress["value"] = 0
        self.progress_label.config(text="0 / 0")
        self._set_state(AppState.RUNNING)
        try:
            self._task_worker.run_batch(cfg)
        except Exception as e:
            self._log_gui(f"启动批量任务失败：{e}")
            self._set_state(AppState.READY)

    def _worker_run(self, cfg):
        """v1.1.0 兼容保留（不再使用）：worker 化后由 TaskWorker._do_batch 替代。"""
        raise RuntimeError("_worker_run 已由 TaskWorker 替代（Stage 6）")

    def _preview(self):
        """Stage 6: Preview into worker (entire Photoshop lifecycle in worker)."""
        from core.task_events import AppState
        if self._state in (AppState.RUNNING, AppState.STOPPING, AppState.LOADING,
                           AppState.PREVIEWING):
            messagebox.showwarning("提示", "已有任务正在运行，请稍候。")
            return
        if not self._ensure_dataset_fresh(trigger="预览"):
            return
        if not self._ensure_psd_fresh(trigger="预览"):
            return
        cfg = self._collect_cfg()
        if not cfg["psd_path"] or not os.path.exists(cfg["psd_path"]) or \
           not cfg["xlsx_path"] or not os.path.exists(cfg["xlsx_path"]):
            messagebox.showerror("错误", "请先选择 PSD 与 Excel，并点击【加载】。")
            return
        if not cfg["out_dir"]:
            messagebox.showerror("错误", "请先选择输出目录（预览将生成在 输出目录/_preview 下，不污染正式输出）。")
            return
        if cfg.get("group_output_enabled"):
            try:
                assert_group_column_valid(
                    self.excel_dataset.max_columns if self.excel_dataset is not None else 0,
                    True, cfg.get("group_output_column"))
            except OutputPathError as e:
                messagebox.showerror("分组配置错误", str(e))
                return
        if self._task_worker.worker_alive:
            messagebox.showwarning("提示", "已有任务正在运行，请稍候。")
            return
        self._set_state(AppState.PREVIEWING)
        try:
            self._task_worker.run_preview(cfg)
        except Exception as e:
            self._log_gui(f"启动预览任务失败：{e}")
            self._set_state(AppState.READY)

    def _on_preview_done(self, result):
        """Preview done (main thread): os.startfile 是用户可见动作，由 main thread 执行。"""
        from core.task_events import AppState
        if result.ok and result.preview_path:
            self._log_gui(f"预览已生成：{result.preview_path}")
            try:
                os.startfile(result.preview_path)
            except Exception:
                pass
        else:
            for e in (result.errors or []):
                self._log_gui(f"预览失败：{e}")
            if result.error and not result.errors:
                self._log_gui(f"预览失败：{result.error}")
        # worker 已退出；PREVIEWING -> READY
        self._set_state(AppState.READY)

    def _on_batch_done(self, summary):
        """Batch done (main thread): 页面内 Summary（BatchResult 唯一数据源，规格第 23/24 节）。"""
        from core.task_events import AppState
        s = summary or {}
        self._log_gui(
            f"批量完成：成功 {s.get('success', 0)}，失败 {s.get('failed', 0)}，"
            f"跳过 {s.get('skipped', 0)}，耗时 {summary_duration_text(s.get('duration_seconds', 0))}"
            + ("（已停止）" if s.get("cancelled") else ""))
        # 页面内 Summary（规格第 23/24 节：BatchResult 唯一数据源）
        self._render_summary(s)
        # 若仍有失败，列出明细
        for r in s.get("rows", []):
            if r.get("status") == "FAILED" and r.get("errors"):
                self._log_gui(f"  [失败 行{r.get('excel_row')}] {'；'.join(r['errors'])}")
        # Batch 单行失败 ≠ GUI ERROR：回 READY（Stage 6 #24）
        self._set_state(AppState.READY)
        # Stage 7.5：batch 成功消费了最新 PSD（若有变更已用 psd_force_reload 重读），
        # 同步更新指纹，避免下次开始/预览误报「PSD 已变更」。
        try:
            cur = self.psd_var.get().strip()
            if cur and os.path.exists(cur):
                self._psd_fingerprint = (cur, self._psd_fingerprint_of(cur))
        except Exception:
            pass

    # ---------------- Summary 渲染（规格 6.5B 第 17/18 节） ----------------
    def _render_summary_placeholder(self):
        """初始占位（未执行时显示简单灰字，不占大高度）。"""
        for w in self.summary_box.winfo_children():
            w.destroy()
        section_help(self.summary_box, "尚未开始生成").pack(anchor="w")

    def _render_summary(self, summary):
        """把 BatchResult summary dict 渲染为页面内 Summary 卡片（唯一数据源）。"""
        for w in self.summary_box.winfo_children():
            w.destroy()
        m = batch_summary_model(summary)
        # 标题行：headline + 徽标
        head = ttk.Frame(self.summary_box)
        head.pack(fill="x")
        tb.Label(head, text=m["headline"], bootstyle=m["bootstyle"],
                 font=FONT_BODY, padding=(PAD_SM, PAD_XS)).pack(side="left")
        dur = summary_duration_text(m["duration_seconds"])
        tb.Label(head, text=f"耗时 {dur}", bootstyle="secondary",
                 font=FONT_HELP).pack(side="left", padx=(PAD_SM, 0))
        # 统计行：成功 / 失败 / 跳过（数字用语义色）
        stats = ttk.Frame(self.summary_box)
        stats.pack(fill="x", pady=(PAD_XS, 0))
        tb.Label(stats, text=f"成功 {m['success']}", bootstyle=BS_SUCCESS,
                 font=FONT_BODY).pack(side="left")
        tb.Label(stats, text=f"失败 {m['failed']}",
                 bootstyle=(BS_DANGER if m["failed"] else BS_SECONDARY),
                 font=FONT_BODY).pack(side="left", padx=(PAD_SM, 0))
        tb.Label(stats, text=f"跳过 {m['skipped']}", bootstyle=BS_SECONDARY,
                 font=FONT_BODY).pack(side="left", padx=(PAD_SM, 0))
        # 输出目录行（规格 6.5B 第 18 节）+ 打开输出目录
        out_dir = ""
        if summary:
            out_dir = (summary.get("out_dir") or summary.get("output_dir") or "")
        if out_dir:
            outrow = ttk.Frame(self.summary_box)
            outrow.pack(fill="x", pady=(PAD_XS, 0))
            tb.Label(outrow, text="输出目录", bootstyle="secondary",
                     font=FONT_HELP).pack(side="left")
            tb.Label(outrow, text=shorten_path(out_dir, 70), bootstyle="secondary",
                     font=FONT_HELP).pack(side="left", padx=(PAD_SM, 0))
            self.btn_open_out = tb.Button(outrow, text="打开输出目录",
                                          command=lambda: self._open_out_dir(out_dir),
                                          bootstyle=BS_OUTLINE_SECONDARY)
            self.btn_open_out.pack(side="left", padx=(PAD_SM, 0))
        # 失败明细（有失败时展示，最多 8 行，其余进日志）
        if m["rows_failed"]:
            fl = ttk.Frame(self.summary_box)
            fl.pack(fill="x", pady=(PAD_XS, 0))
            tb.Label(fl, text="失败明细：", bootstyle=BS_DANGER,
                     font=FONT_HELP).pack(side="left", anchor="n")
            inner = ttk.Frame(fl)
            inner.pack(side="left", fill="x", expand=True)
            for rf in m["rows_failed"][:8]:   # 最多展示 8 行，其余进日志
                who = rf.get("name") or rf.get("store") or f"行{rf.get('excel_row')}"
                tb.Label(inner, text=f"· {who}：{'；'.join(rf['errors'])[:80]}",
                         bootstyle="secondary", font=FONT_HELP,
                         anchor="w").pack(fill="x")
            if len(m["rows_failed"]) > 8:
                tb.Label(inner, text=f"… 其余 {len(m['rows_failed']) - 8} 行见日志",
                         bootstyle="secondary", font=FONT_HELP).pack(fill="x")

    def _open_out_dir(self, out_dir):
        """打开输出目录（不存在则用系统默认目录兜底）。"""
        target = out_dir if out_dir and os.path.isdir(out_dir) else os.getcwd()
        try:
            os.startfile(target)
        except Exception as e:
            self._log_gui(f"打开目录失败：{e}")

    # ---------------- 配置检查（规格第 18 节：inline 列表） ----------------
    def _render_cfg_check_placeholder(self):
        for w in self.cfg_check_box.winfo_children():
            w.destroy()
        tb.Label(self.cfg_check_box, text="点击「检查配置」查看是否满足生成条件。",
                 bootstyle="secondary").pack(side="left")

    def _check_config(self):
        """配置检查：只读检查 + inline 列表展示（非弹窗；规格第 18 节）。"""
        from core.task_events import AppState
        cfg = self._collect_cfg()
        # 分组列标签（展示用）：group_output_column 是 0-based 索引，转列字母
        gcol_label = ""
        if cfg.get("group_output_enabled") and cfg.get("group_output_column") is not None:
            gcol_label = index_to_excel_column(cfg.get("group_output_column"))
        items = config_check_model(
            cfg, layer_index_loaded=self.layer_index is not None,
            stores=self.excel_stores,
            text_map={"姓名": cfg.get("text_map", {}).get("姓名", "")},
            store_logo_map=cfg.get("store_logo_map", {}),
            ds_valid_rows=len(self.excel_stores) if self.excel_stores else None,
            group_enabled=cfg.get("group_output_enabled", False),
            group_column_label=gcol_label,
        )
        for w in self.cfg_check_box.winfo_children():
            w.destroy()
        grid = self.cfg_check_box
        for i, it in enumerate(items):
            tb.Label(grid, text=it["label"], bootstyle="secondary", width=12, anchor="w").grid(
                row=i, column=0, sticky="w", padx=(0, PAD_XS), pady=1)
            ok_badge = tb.Label(grid, text="✓" if it["ok"] else "✕",
                                bootstyle=(BS_SUCCESS if it["ok"] else BS_DANGER),
                                padding=(PAD_XS, 0))
            ok_badge.grid(row=i, column=1, sticky="w", padx=(0, PAD_XS), pady=1)
            tb.Label(grid, text=shorten_path(str(it["detail"]), 60), bootstyle="secondary",
                     anchor="w").grid(row=i, column=2, sticky="w", pady=1)
        # 全部 ok -> 绿色提示；否则提示未就绪（不弹窗）
        if all(it["ok"] for it in items):
            tb.Label(grid, text="全部就绪，可以开始生成。", bootstyle=BS_SUCCESS).grid(
                row=len(items), column=0, columnspan=3, sticky="w", pady=(PAD_XS, 0))
        else:
            tb.Label(grid, text="部分未就绪，请按上述提示调整。", bootstyle=BS_WARNING).grid(
                row=len(items), column=0, columnspan=3, sticky="w", pady=(PAD_XS, 0))

    # ---------------- 日志工具栏（规格第 25 节） ----------------
    def _log_clear(self):
        self.log.config(state="normal")
        self.log.delete("1.0", "end")
        self.log.config(state="disabled")

    def _log_copy(self):
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append(self.log.get("1.0", "end-1c"))
        except Exception as e:
            self._log_gui(f"复制日志失败：{e}")

    def _log_open_dir(self):
        """打开日志输出目录（输出目录或当前目录）。"""
        out = self.out_var.get().strip() if hasattr(self, "out_var") else ""
        target = out if out and os.path.isdir(out) else os.getcwd()
        try:
            os.startfile(target)
        except Exception as e:
            self._log_gui(f"打开目录失败：{e}")

    def _stop(self):
        """协作式取消：RUNNING -> STOPPING，set cancel_event（不杀线程/Photoshop）。"""
        from core.task_events import AppState
        if self._state == AppState.RUNNING:
            self._set_state(AppState.STOPPING)
            self._task_worker.request_stop()
            self._log_gui("正在停止，将在当前 Photoshop 操作完成后结束...")

    def _on_finish(self):
        """v1.1.0 兼容保留（不再使用）。"""
        pass

    def _update_progress(self, done, total, phase="", excel_row=0, store="", name=""):
        # 规格第 20/21/22 节：78/221 + 35% + 门店·姓名 + 阶段
        pm = progress_display(done, total, phase)
        if total > 0:
            self.progress["value"] = pm["percent"]
            self.progress_pct.config(text=f"{pm['percent']}%")
        else:
            self.progress["value"] = 0
            self.progress_pct.config(text="")
        self.progress_label.config(text=pm["text"])
        cur_parts = []
        if store:
            cur_parts.append(store)
        if name:
            cur_parts.append(name)
        self.progress_cur.config(text=" · ".join(cur_parts))
        self.progress_phase.config(text=pm["phase_text"])

    def _log_gui(self, msg, level=""):
        """写日志（规格第 25/26 节：INFO/WARN/ERROR 视觉区分）。"""
        if not level:
            level = log_level_of(msg)
        self.log.config(state="normal")
        tag = f"log_{level}"
        self.log.tag_configure("log_error", foreground="#c0392b")
        self.log.tag_configure("log_warn", foreground="#b9770e")
        self.log.tag_configure("log_info", foreground="#2c3e50")
        self.log.insert("end", msg + "\n", (tag,))
        self.log.see("end")
        self.log.config(state="disabled")

    # ---------------- 配置存取 ----------------
    def _save_config(self):
        cfg = self._collect_cfg()
        # 即使未加载，也保存路径与列设置
        try:
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(cfg, f, ensure_ascii=False, indent=2)
            self._log_gui(f"配置已保存：{self.config_path}")
        except Exception as e:
            self._log_gui(f"保存配置失败：{e}")

    def _load_config(self):
        if not os.path.exists(self.config_path):
            return
        try:
            with open(self.config_path, "r", encoding="utf-8") as f:
                c = json.load(f)
            self.psd_var.set(c.get("psd_path", ""))
            self.xlsx_var.set(c.get("xlsx_path", ""))
            self.out_var.set(c.get("out_dir", ""))
            self.header_var.set(c.get("has_header", True))
            # Stage 4：列名用 index_to_excel_column（支持 AA/AB...）
            self.col_store_var.set(index_to_excel_column(c.get("col_store", 0)))
            self.col_name_var.set(index_to_excel_column(c.get("col_name", 1)))
            self.col_phone_var.set(index_to_excel_column(c.get("col_phone", 3)))
            rv = c.get("col_role", 2)
            self.col_role_var.set("（不替换）" if rv < 0 else index_to_excel_column(rv))
            # Stage 4.5：分组配置恢复（group_output_column 为 0-based index）
            self.group_output_var.set(c.get("group_output_enabled", False))
            gcol = c.get("group_output_column")
            if gcol is not None:
                self.group_col_var.set(index_to_excel_column(gcol))
            elif self._column_labels:
                self.group_col_var.set(self._column_labels[0])
            self.fmt_var.set(c.get("fmt", FMT_PNG))
            self.also_png_var.set(c.get("also_png", False))
            # 文字图层映射（加载后会在 _load 中按实际图层校正，这里仅预填）
            tm = c.get("text_map", {})
            self.tm_name_var.set(self._tm_display(tm.get("姓名", "")))
            self.tm_phone_var.set(self._tm_display(tm.get("电话", "")))
            self.tm_role_var.set(self._tm_display(tm.get("销售顾问", "")))
            # 保留上次的 logo 勾选与映射，供 _load 默认填充
            self.cfg = c
        except Exception:
            pass

    def _tm_display(self, v):
        """把配置值（LayerRef dict 或旧 name 字符串）转成 GUI 下拉显示 label。

        规格 6.5B 第 10 节：任何情况下 UI 不显示 dict / JSON / layer_id / index_path。
        """
        if not v:
            return "（不替换）"
        # 防御：历史配置可能直接存了 serialize_ref dict / dict 字符串
        if isinstance(v, dict) or (isinstance(v, str) and v.startswith("{")):
            label = layer_display_label(v, self.layer_labels if self.layer_index else None)
            if label and label != "（不替换）":
                return label
            return "（不替换）"
        if self.layer_index is None:
            # 未加载 PSD：原样字符串（已排除 dict 形态）
            return v
        ref = ref_from_config(v)
        if ref is None:
            return "（不替换）"
        if ref.id and ref.id in self._ref_by_id:
            return self.layer_labels[ref.id]
        m = self.layer_index.find_matching(ref.name)
        if len(m) == 1:
            return self.layer_labels[m[0].id]
        return "（不替换）"   # ambiguous / missing：不自动选


def main():
    # Stage 6.5B：ttkbootstrap Window（litera 主题在 APP_THEME 定义，主题数据随 PyInstaller 打包）
    root = tb.Window(themename=APP_THEME)
    try:
        root.tk.call("encoding", "system", "utf-8")
    except Exception:
        pass
    app = App(root)

    def _on_exit():
        # Stage 6：安全关闭 —— worker 运行时先协作取消，等 worker 自然退出再 destroy。
        # 绝不 kill thread / kill Photoshop；也不在 worker 存活时直接 destroy。
        if app._task_worker.worker_alive:
            app._pending_close = True
            app._task_worker.request_close()   # closing flag + cancel_event
            from core.task_events import AppState
            if app._state == AppState.RUNNING:
                app._set_state(AppState.STOPPING)
            app._log_gui("正在停止并退出（等待当前 Photoshop 操作完成）...")
            # 轮询仍在运行；worker 退出后 _maybe_destroy 会 destroy root
            return
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", _on_exit)
    root.mainloop()


if __name__ == "__main__":
    main()
