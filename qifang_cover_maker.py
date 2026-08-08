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

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

import win32com.client
import pythoncom

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
        root.title(APP_TITLE)
        root.geometry("820x760")
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
        self.running = False
        self.stop_flag = threading.Event()
        self.worker = None

        self._build_ui()
        self._load_config()

    # ---------------- 路径 ----------------
    def _find_config_path(self):
        if getattr(sys, "frozen", False):
            base = os.path.dirname(sys.executable)
        else:
            base = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(base, CONFIG_NAME)

    # ---------------- UI 构建 ----------------
    def _build_ui(self):
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=8, pady=8)

        # Tab 1: 文件与列映射
        f1 = ttk.Frame(self.notebook)
        self.notebook.add(f1, text="① 文件与字段")
        self._build_tab_files(f1)

        # Tab 2: Logo 映射
        f2 = ttk.Frame(self.notebook)
        self.notebook.add(f2, text="② 门店Logo映射")
        self._build_tab_logo(f2)

        # Tab 3: 运行
        f3 = ttk.Frame(self.notebook)
        self.notebook.add(f3, text="③ 生成")
        self._build_tab_run(f3)

    def _build_tab_files(self, parent):
        # 文件选择
        row = ttk.LabelFrame(parent, text="文件选择", padding=8)
        row.pack(fill="x", padx=6, pady=6)

        ttk.Label(row, text="PSD 模板:").grid(row=0, column=0, sticky="e", pady=3)
        self.psd_var = tk.StringVar()
        ttk.Entry(row, textvariable=self.psd_var, width=55).grid(row=0, column=1, padx=4)
        ttk.Button(row, text="浏览...", command=self._pick_psd).grid(row=0, column=2)

        ttk.Label(row, text="Excel 数据:").grid(row=1, column=0, sticky="e", pady=3)
        self.xlsx_var = tk.StringVar()
        ttk.Entry(row, textvariable=self.xlsx_var, width=55).grid(row=1, column=1, padx=4)
        ttk.Button(row, text="浏览...", command=self._pick_xlsx).grid(row=1, column=2)

        ttk.Label(row, text="输出目录:").grid(row=2, column=0, sticky="e", pady=3)
        self.out_var = tk.StringVar()
        ttk.Entry(row, textvariable=self.out_var, width=55).grid(row=2, column=1, padx=4)
        ttk.Button(row, text="浏览...", command=self._pick_out).grid(row=2, column=2)

        # Stage 4：has_header 默认 True（已有 config 优先，见 _load_config）
        self.header_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(row, text="Excel 首行为表头（数据从第2行开始）",
                        variable=self.header_var).grid(row=3, column=1, sticky="w", pady=3)

        # Stage 4.5：按 Excel 任意列创建输出子文件夹（group_output_enabled / group_output_column）
        self.group_output_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(row, text="☑ 按 Excel 列创建输出子文件夹",
                        variable=self.group_output_var).grid(row=4, column=1, sticky="w", pady=3)
        ttk.Label(row, text="分组字段:").grid(row=4, column=2, sticky="e", pady=3)
        self.group_col_var = tk.StringVar(value=cols[0] if cols else "A")
        self.group_col_cb = ttk.Combobox(row, textvariable=self.group_col_var, values=cols,
                                         width=14, state="readonly")
        self.group_col_cb.grid(row=4, column=3, sticky="w", padx=4)

        # 字段映射（Stage 4：列下拉按实际工作表动态生成，这里只放默认 A..Z 兜底）
        mf = ttk.LabelFrame(parent, text="字段映射（选择 Excel 列）", padding=8)
        mf.pack(fill="x", padx=6, pady=6)

        cols = [index_to_excel_column(i) for i in range(26)]  # A..Z（加载后按实际列数刷新）
        self.col_store_var = tk.StringVar(value="A")
        self.col_name_var = tk.StringVar(value="B")
        self.col_phone_var = tk.StringVar(value="D")
        self.col_role_var = tk.StringVar(value="C")
        self._column_labels = cols

        ttk.Label(mf, text="门店列（用于选 Logo）:").grid(row=0, column=0, sticky="e", pady=3)
        self.col_store_cb = ttk.Combobox(mf, textvariable=self.col_store_var, values=cols,
                                         width=8, state="readonly")
        self.col_store_cb.grid(row=0, column=1, padx=4, sticky="w")
        ttk.Label(mf, text="姓名列:").grid(row=0, column=2, sticky="e", pady=3)
        self.col_name_cb = ttk.Combobox(mf, textvariable=self.col_name_var, values=cols,
                                        width=8, state="readonly")
        self.col_name_cb.grid(row=0, column=3, padx=4, sticky="w")
        ttk.Label(mf, text="电话列:").grid(row=1, column=0, sticky="e", pady=3)
        self.col_phone_cb = ttk.Combobox(mf, textvariable=self.col_phone_var, values=cols,
                                         width=8, state="readonly")
        self.col_phone_cb.grid(row=1, column=1, padx=4, sticky="w")
        ttk.Label(mf, text="销售顾问列:").grid(row=1, column=2, sticky="e", pady=3)
        self.col_role_cb = ttk.Combobox(mf, textvariable=self.col_role_var, values=cols + ["（不替换）"],
                                        width=12, state="readonly")
        self.col_role_cb.grid(row=1, column=3, padx=4, sticky="w")

        ttk.Label(mf, text="提示：销售顾问列选「不替换」则保留 PSD 原文字。").grid(
            row=2, column=0, columnspan=4, sticky="w", pady=(4, 0))

        # 文字图层映射（PSD 内实际文字图层 -> 字段）
        tf = ttk.LabelFrame(parent, text="文字图层映射（PSD 内实际文字图层，加载后自动识别）", padding=8)
        tf.pack(fill="x", padx=6, pady=6)

        self.tm_name_var = tk.StringVar(value="（不替换）")
        self.tm_phone_var = tk.StringVar(value="（不替换）")
        self.tm_role_var = tk.StringVar(value="（不替换）")
        tm_opts = ["（不替换）"]

        ttk.Label(tf, text="姓名 →").grid(row=0, column=0, sticky="e", pady=3)
        self.tm_name_cb = ttk.Combobox(tf, textvariable=self.tm_name_var, values=tm_opts,
                                       width=30, state="readonly")
        self.tm_name_cb.grid(row=0, column=1, padx=4, sticky="w")
        ttk.Label(tf, text="电话 →").grid(row=0, column=2, sticky="e", pady=3)
        self.tm_phone_cb = ttk.Combobox(tf, textvariable=self.tm_phone_var, values=tm_opts,
                                        width=30, state="readonly")
        self.tm_phone_cb.grid(row=0, column=3, padx=4, sticky="w")
        ttk.Label(tf, text="销售顾问 →").grid(row=1, column=0, sticky="e", pady=3)
        self.tm_role_cb = ttk.Combobox(tf, textvariable=self.tm_role_var, values=tm_opts,
                                       width=30, state="readonly")
        self.tm_role_cb.grid(row=1, column=1, padx=4, sticky="w")

        ttk.Label(tf, text="提示：若 PSD 内文字层名与「姓名/电话/销售顾问」不同（如带空格），"
                            "在此手动指定；没有对应图层请保持「不替换」。").grid(
            row=2, column=0, columnspan=4, sticky="w", pady=(4, 0))

        # 加载按钮
        bf = ttk.Frame(parent)
        bf.pack(fill="x", padx=6, pady=6)
        ttk.Button(bf, text="加载 PSD / Excel 并分析图层", command=self._load).pack(side="left")
        ttk.Button(bf, text="保存配置", command=self._save_config).pack(side="left", padx=6)

    def _build_tab_logo(self, parent):
        info = ttk.Label(parent, text="先到「① 文件与字段」点击【加载】按钮，解析出 PSD 图层与 Excel 门店后再配置。")
        info.pack(fill="x", padx=8, pady=8)
        self.logo_info = info

        # 左右分栏
        paned = ttk.PanedWindow(parent, orient="horizontal")
        paned.pack(fill="both", expand=True, padx=6, pady=6)

        left = ttk.Frame(paned)
        right = ttk.Frame(paned)
        paned.add(left, weight=1)
        paned.add(right, weight=1)

        # 左：Logo 图层清单（可勾选）
        lf = ttk.LabelFrame(left, text="标记为 Logo 图层（PSD 内的候选 Logo）", padding=6)
        lf.pack(fill="both", expand=True, padx=4, pady=4)
        self.logo_canvas = tk.Canvas(lf)
        self.logo_scroll = ttk.Scrollbar(lf, orient="vertical", command=self.logo_canvas.yview)
        self.logo_inner = ttk.Frame(self.logo_canvas)
        self.logo_inner.bind("<Configure>", lambda e: self.logo_canvas.configure(
            scrollregion=self.logo_canvas.bbox("all")))
        self.logo_canvas.create_window((0, 0), window=self.logo_inner, anchor="nw")
        self.logo_canvas.configure(yscrollcommand=self.logo_scroll.set)
        self.logo_canvas.pack(side="left", fill="both", expand=True)
        self.logo_scroll.pack(side="right", fill="y")
        self.logo_checks = {}  # name -> BooleanVar

        # 右：门店 -> Logo 映射
        rf = ttk.LabelFrame(right, text="门店 → Logo 图层 映射", padding=6)
        rf.pack(fill="both", expand=True, padx=4, pady=4)

        # 品牌 Logo 人工指定（Stage 3 补充）：
        #   「固定品牌 Logo」= 每张封面强制显示的叶子（如 七方logo / 圣大）。
        #   独立于门店映射：勾选 selected 只是候选集合，不等于 brand；
        #   这里人工勾选某个 leaf 才真正把它写入 brand_logo_refs。
        bf2 = ttk.LabelFrame(rf, text="固定品牌 Logo（每张封面都显示；勾选 = 人工指定 BRAND）",
                             padding=4)
        bf2.pack(fill="x", padx=4, pady=(0, 4))
        self.brand_canvas = tk.Canvas(bf2, height=90)
        self.brand_scroll = ttk.Scrollbar(bf2, orient="vertical", command=self.brand_canvas.yview)
        self.brand_inner = ttk.Frame(self.brand_canvas)
        self.brand_inner.bind("<Configure>", lambda e: self.brand_canvas.configure(
            scrollregion=self.brand_canvas.bbox("all")))
        self.brand_canvas.create_window((0, 0), window=self.brand_inner, anchor="nw")
        self.brand_canvas.configure(yscrollcommand=self.brand_scroll.set)
        self.brand_canvas.pack(side="left", fill="both", expand=True)
        self.brand_scroll.pack(side="right", fill="y")

        self.map_canvas = tk.Canvas(rf)
        self.map_scroll = ttk.Scrollbar(rf, orient="vertical", command=self.map_canvas.yview)
        self.map_inner = ttk.Frame(self.map_canvas)
        self.map_inner.bind("<Configure>", lambda e: self.map_canvas.configure(
            scrollregion=self.map_canvas.bbox("all")))
        self.map_canvas.create_window((0, 0), window=self.map_inner, anchor="nw")
        self.map_canvas.configure(yscrollcommand=self.map_scroll.set)
        self.map_canvas.pack(side="left", fill="both", expand=True)
        self.map_scroll.pack(side="right", fill="y")
        self.map_combos = {}  # store -> StringVar
        self.map_combo_widgets = {}  # store -> Combobox widget

    def _build_tab_run(self, parent):
        # 选项
        opt = ttk.LabelFrame(parent, text="导出选项", padding=8)
        opt.pack(fill="x", padx=6, pady=6)
        self.fmt_var = tk.StringVar(value=FMT_PNG)
        ttk.Label(opt, text="导出格式:").grid(row=0, column=0, sticky="e", pady=3)
        ttk.Combobox(opt, textvariable=self.fmt_var, values=[FMT_PNG, FMT_JPG, FMT_PSD],
                     width=8, state="readonly").grid(row=0, column=1, sticky="w", padx=4)
        self.also_png_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(opt, text="导出 PSD 时另外再生成 PNG（便于预览）",
                        variable=self.also_png_var).grid(row=0, column=2, sticky="w", padx=10)

        # 进度
        pf = ttk.Frame(parent)
        pf.pack(fill="x", padx=6, pady=4)
        self.progress = ttk.Progressbar(pf, mode="determinate", maximum=100)
        self.progress.pack(fill="x", side="left", expand=True)
        self.progress_label = ttk.Label(pf, text="0 / 0")
        self.progress_label.pack(side="left", padx=6)

        # 按钮
        bf = ttk.Frame(parent)
        bf.pack(fill="x", padx=6, pady=4)
        self.btn_run = ttk.Button(bf, text="▶ 开始生成", command=self._start)
        self.btn_run.pack(side="left")
        self.btn_stop = ttk.Button(bf, text="■ 停止", command=self._stop, state="disabled")
        self.btn_stop.pack(side="left", padx=6)
        self.btn_preview = ttk.Button(bf, text="🖼 试做第1张预览", command=self._preview)
        self.btn_preview.pack(side="left", padx=6)

        # 日志
        lf = ttk.LabelFrame(parent, text="运行日志", padding=6)
        lf.pack(fill="both", expand=True, padx=6, pady=6)
        self.log = scrolledtext.ScrolledText(lf, height=12, state="disabled",
                                             font=("Consolas", 9))
        self.log.pack(fill="both", expand=True)

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

    def _load(self):
        psd = self.psd_var.get().strip()
        xlsx = self.xlsx_var.get().strip()
        if not psd or not os.path.exists(psd):
            messagebox.showerror("错误", "请先选择有效的 PSD 模板文件。")
            return
        if not xlsx or not os.path.exists(xlsx):
            messagebox.showerror("错误", "请先选择有效的 Excel 数据文件。")
            return

        # 读取 Excel 门店 + 列头（Stage 4：统一入口 load_excel_dataset，修复 P1-01 ——
        # 门店不再固定取 A 列，而是按用户配置的 col_store 列读取）
        if not self._load_excel_data(xlsx):
            return

        # 读取 PSD 图层（通过 Photoshop COM；Session 负责 COM 初始化并只关闭自己打开的文档）
        try:
            self._log_gui(f"正在用 Photoshop 解析 PSD 图层：{os.path.basename(psd)}")
            with PhotoshopSession() as ps:
                doc = ps.open_document(psd)
                time.sleep(0.6)
                # Stage 2：建立唯一 LayerIndex（不丢同名图层），替代旧的按 name 去重
                self.layer_index = collect_layer_index(doc)
                self._ref_by_id = {r.id: r for r in self.layer_index.layers}
                self.layer_labels = self.layer_index.labels()   # id -> 唯一 label
                # 兼容缓存（旧逻辑仍读取；新建逻辑优先用 layer_index）
                self.all_psd_layers = [r.name for r in self.layer_index.layers]
                self.all_psd_is_group = {r.id: r.is_group for r in self.layer_index.layers}
                self.all_psd_parent = {r.id: self._parent_name_of(r) for r in self.layer_index.layers}
                self.all_text_layers = [r.display_path for r in self.layer_index.layers if r.is_text]
                # 文字层 label -> LayerRef（同名 display_path 已由 labels 附加 id 后缀）
                self.text_label_to_ref = {}
                for r in self.layer_index.layers:
                    if r.is_text:
                        self.text_label_to_ref[self.layer_labels[r.id]] = r
                self.logo_label_to_ref = {}
                # with 退出：Session 只关闭自己打开的 doc（即上面的模板文档）
        except Exception as e:
            messagebox.showerror("PSD 错误", f"无法解析 PSD（请确认 Photoshop 已安装并可启动）：\n{e}")
            return

        # ---- Logo 勾选与映射（Stage 3 数据流）----
        # 概念分离：
        #   勾选 = selected Logo refs（可含组，仅 GUI selection）
        #   展开 = resolve_effective_logo_layers(selected) -> 叶子候选
        #   映射 = store_logo_map（叶子 LayerRef）+ brand_logo_refs（叶子 LayerRef）
        # 名称启发式只用于「首次自动推荐」；运行时绝不再次 name 猜测。
        stores = self.excel_stores   # Stage 4：来自 _load_excel_data 的 dataset.stores
        def is_logo_candidate(ref):
            # 仅用于首次加载的推荐勾选（不是运行时规则）
            if ref.is_group:
                return False
            if "logo" in ref.name.lower():
                return True
            parent = (self._parent_name_of(ref) or "").lower()
            if "logo" in parent:
                return True
            return ref.name in stores

        # 首次加载推荐勾选（修复 P1-02）：
        #   recommend_logo_selection = logo heuristic 勾选 ∪ 与任一门店唯一命中叶子
        #   —— 保证「普通素材 > 康乐电器」这类无 logo 关键字的门店目标叶子进入
        #      effective，match_store_logo 才能看到它（AUTO 推荐康乐电器）。
        #   注意：只在「无保存配置」时推荐；有保存的 logo_selection 一律以保存值为准。
        saved_items_raw = self.cfg.get("logo_selection")
        if saved_items_raw is None:
            saved_items_raw = self.cfg.get("logo_layers", [])
        has_saved_selection = bool(saved_items_raw)

        # 恢复上次保存的勾选状态（优先新字段 logo_selection，兼容旧 logo_layers）
        saved_items = saved_items_raw
        saved_names = {x for x in saved_items if isinstance(x, str)}
        saved_logo_ids = {ref_from_config(x).id for x in saved_items
                          if ref_from_config(x) is not None and ref_from_config(x).id}

        def inherited_checked(ref):
            if ref.id in saved_logo_ids or ref.name in saved_names:
                return True
            parent = self._parent_name_of(ref)
            seen = set()
            while parent:
                if parent in saved_names or parent in saved_logo_ids:
                    return True
                if parent in seen:
                    break
                seen.add(parent)
                parent = self._parent_name_of_str(parent)
            return False

        self.logo_checks = {}
        # 首次推荐（无保存配置时）：recommend_logo_selection（含门店匹配目标）
        recommend = []
        if not has_saved_selection:
            recommend = recommend_logo_selection(
                self.layer_index.layers, stores, is_logo_heuristic=is_logo_candidate)
        for ref in self.layer_index.layers:
            label = self.layer_labels[ref.id]
            if has_saved_selection:
                checked = inherited_checked(ref)
            else:
                checked = ref.id in {r.id for r in recommend} or inherited_checked(ref)
            self.logo_checks[label] = tk.BooleanVar(value=checked)

        # 自动匹配 门店->Logo：候选 = 当前勾选展开后的全部叶子（含组展开）。
        # 不再用 is_logo_candidate 预过滤（修复「康乐 -> 康乐电器」P0）：
        #   先展开 effective leaves，再 match_store_logo 评分；
        #   AMBIGUOUS / NO_MATCH -> 保持（无），绝不自动选第一个。
        prev_map = self.cfg.get("store_logo_map", {})
        self.map_combos = {}
        selected_refs = self._selected_logo_refs()
        effective_leaves = resolve_effective_logo_layers(self.layer_index, selected_refs)
        for s in stores:
            default_label = "（无）"
            if s in prev_map:
                prev = prev_map[s]
                prev_ref = ref_from_config(prev)
                # 旧配置 name 唯一命中 -> 用；否则视为未配置
                if prev_ref is not None:
                    if prev_ref.id and prev_ref.id in self._ref_by_id:
                        default_label = self.layer_labels[prev_ref.id]
                    else:
                        m = self.layer_index.find_matching(prev_ref.name)
                        if len(m) == 1:
                            default_label = self.layer_labels[m[0].id]
                        elif len(m) > 1:
                            default_label = "（无）"   # ambiguous：不自动选
            else:
                # 首次自动推荐：只用当前 effective 叶子候选（不预过滤 name 含 logo）
                mr = match_store_logo(s, effective_leaves)
                if mr.status in (EXACT, AUTO) and mr.best is not None:
                    default_label = self.layer_labels[mr.best.id]
                else:
                    default_label = "（无）"   # AMBIGUOUS / NO_MATCH：不自动选
            self.map_combos[s] = tk.StringVar(value=default_label)

        # 文字图层映射：Stage 2 候选显示 display_path（同名肉眼可区分）；
        # 自动匹配：名称（忽略空格）全局唯一才自动选，歧义保持「不替换」
        opts = ["（不替换）"] + [self.layer_labels[r.id] for r in self.layer_index.layers if r.is_text]
        self.tm_name_cb["values"] = opts
        self.tm_phone_cb["values"] = opts
        self.tm_role_cb["values"] = opts
        prev_tm = self.cfg.get("text_map", {})

        def auto_text(field):
            t = field.replace(" ", "")
            # 先用上次的显式配置（LayerRef dict 或旧 name）
            if prev_tm.get(field):
                prev = prev_tm[field]
                prev_ref = ref_from_config(prev)
                if prev_ref is not None:
                    if prev_ref.id and prev_ref.id in self._ref_by_id:
                        return self.layer_labels[prev_ref.id]
                    m = self.layer_index.find_matching(prev_ref.name)
                    if len(m) == 1:
                        return self.layer_labels[m[0].id]
                    return "（不替换）"   # ambiguous / missing
            # 再按名称（去空格）自动匹配：仅当全局唯一
            m = self.layer_index.find_matching(field)
            if len(m) == 1:
                return self.layer_labels[m[0].id]
            return "（不替换）"

        self.tm_name_var.set(auto_text("姓名"))
        self.tm_phone_var.set(auto_text("电话"))
        self.tm_role_var.set(auto_text("销售顾问"))

        # ---- 品牌 Logo 初始勾选（Stage 3 补充）----
        # 优先级：
        #   1) 保存配置 brand_logo_layers（重启恢复，人工指定为准）；
        #   2) 否则 suggest_brand_logos 建议（仅默认建议，非唯一入口）。
        # 运行时只读保存后的 brand_logo_refs，此处只是预填 GUI 勾选。
        saved_brand = self.cfg.get("brand_logo_layers")
        self.brand_checks = {}
        if saved_brand:
            saved_brand_ids = set()
            for v in saved_brand:
                ref = ref_from_config(v)
                if ref is not None and ref.id and ref.id in self._ref_by_id:
                    saved_brand_ids.add(ref.id)
            for label in self._effective_logo_layers():
                ref = self._label_to_ref(label)
                checked = ref is not None and ref.id in saved_brand_ids
                self.brand_checks[label] = tk.BooleanVar(value=checked)
        else:
            # 首次：suggest_brand_logos 建议（name 含 logo 且未映射）仅默认勾选
            # 构造 store -> ref（当前下拉已选值，可能为「（无）」）
            store_map_now = {}
            for s, var in self.map_combos.items():
                v = var.get()
                store_map_now[s] = self._label_to_ref(v) if v != "（无）" else None
            try:
                suggest = suggest_brand_logos(effective_leaves, store_map_now)
            except Exception:
                suggest = []
            suggest_ids = {r.id for r in suggest}
            for label in self._effective_logo_layers():
                ref = self._label_to_ref(label)
                self.brand_checks[label] = tk.BooleanVar(
                    value=ref is not None and ref.id in suggest_ids)

        self._rebuild_logo_lists()
        self._log_gui(f"解析完成：PSD 共 {len(self.layer_index)} 个图层，Excel 共 {len(stores)} 个门店。")
        self.logo_info.config(text=f"PSD 图层 {len(self.layer_index)} 个 ｜ Excel 门店 {len(stores)} 个 ｜ 请在下方勾选 Logo 并完成映射。")

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
        # 清空
        for w in self.logo_inner.winfo_children():
            w.destroy()
        for w in self.map_inner.winfo_children():
            w.destroy()
        for w in self.brand_inner.winfo_children():
            w.destroy()
        self.map_combo_widgets = {}
        self.brand_widgets = {}

        ttk.Label(self.logo_inner, text="☑ 图层（显示完整路径）", font=("Microsoft YaHei", 9, "bold")).grid(
            row=0, column=0, sticky="w", padx=4, pady=2)
        for i, label in enumerate(self.logo_checks.keys(), start=1):
            cb = ttk.Checkbutton(self.logo_inner, text=label, variable=self.logo_checks[label])
            cb.grid(row=i, column=0, sticky="w", padx=4, pady=1)
            # 勾选状态变化时，同步刷新右侧门店→Logo 下拉列表
            self.logo_checks[label].trace_add(
                "write", lambda *a: self._on_logo_checks_changed())

        # ---- 品牌 Logo 人工指定（Stage 3 补充）----
        # 可选项 = 当前 effective 叶子（勾选的候选集合）；勾选 = 该 leaf 是固定品牌。
        # 语义：selected 只是候选集合，不等于 brand；这里人工勾选才真正写入 brand。
        eff_labels = self._effective_logo_layers()
        if not eff_labels:
            ttk.Label(self.brand_inner, text="（先勾选左侧 Logo 图层后，这里可指定品牌）",
                      foreground="#888").grid(row=0, column=0, sticky="w", padx=4, pady=2)
        else:
            ttk.Label(self.brand_inner, text="☑ 固定品牌 Logo（叶子）",
                      font=("Microsoft YaHei", 9, "bold")).grid(
                row=0, column=0, sticky="w", padx=4, pady=2)
            for i, label in enumerate(eff_labels, start=1):
                # 该 leaf 若已被某门店映射为 store target，标记（冲突提示用）
                store_owner = self._store_owner_of_leaf(label)
                suffix = f"（门店:{store_owner}）" if store_owner else ""
                var = self.brand_checks.get(label, tk.BooleanVar(value=False))
                var.set(var.get())   # 保持既有状态
                cb = ttk.Checkbutton(
                    self.brand_inner, text=label + suffix, variable=var,
                    command=lambda lb=label: self._on_brand_toggle(lb))
                cb.grid(row=i, column=0, sticky="w", padx=4, pady=1)
                self.brand_widgets[label] = cb
                self.brand_checks[label] = var

        # 映射表
        ttk.Label(self.map_inner, text="门店", font=("Microsoft YaHei", 9, "bold")).grid(
            row=0, column=0, sticky="w", padx=4, pady=2)
        ttk.Label(self.map_inner, text="→ Logo 图层", font=("Microsoft YaHei", 9, "bold")).grid(
            row=0, column=1, sticky="w", padx=4, pady=2)
        logo_opts = ["（无）"] + self._effective_logo_layers()
        for i, s in enumerate(self.excel_stores, start=1):
            ttk.Label(self.map_inner, text=s).grid(row=i, column=0, sticky="w", padx=4, pady=1)
            cb = ttk.Combobox(self.map_inner, textvariable=self.map_combos[s],
                              values=logo_opts, width=40, state="readonly")
            cb.grid(row=i, column=1, sticky="w", padx=4, pady=1)
            self.map_combo_widgets[s] = cb

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
        }

    def all_psd_layers_labels(self):
        """当前 PSD 全部图层的展示 label（与 logo_checks 的 key 一致）。"""
        if self.layer_index is not None:
            return [self.layer_labels[r.id] for r in self.layer_index.layers]
        return list(self.all_psd_layers)

    def _start(self):
        if self.running:
            return
        # Stage 4 补充（BLOCKED B/C）：开始前自动重解析，保证 Dataset 与当前列配置一致
        if not self._ensure_dataset_fresh(trigger="开始"):
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
        # Stage 4.5：Preflight —— 分组功能已启用但分组列无效时，阻止开始
        if cfg.get("group_output_enabled"):
            try:
                assert_group_column_valid(
                    self.excel_dataset.max_columns if self.excel_dataset is not None else 0,
                    True, cfg.get("group_output_column"))
            except OutputPathError as e:
                messagebox.showerror("分组配置错误", str(e))
                return

        self.running = True
        self.stop_flag.clear()
        self.btn_run.config(state="disabled")
        self.btn_stop.config(state="normal")
        self.btn_preview.config(state="disabled")
        self.progress["value"] = 0
        self._log_gui("开始批量生成...")

        self.worker = threading.Thread(target=self._worker_run, args=(cfg,), daemon=True)
        self.worker.start()

    def _worker_run(self, cfg):
        def prog(done, total):
            self.root.after(0, lambda: self._update_progress(done, total))
        def log(msg):
            self.root.after(0, lambda: self._log_gui(msg))
        run_batch(cfg, prog, log, self.stop_flag)
        self.root.after(0, self._on_finish)

    def _preview(self):
        if self.running:
            return
        # Stage 4 补充（BLOCKED B/C）：预览前自动重解析，保证 Dataset 与当前列配置一致
        if not self._ensure_dataset_fresh(trigger="预览"):
            return
        cfg = self._collect_cfg()
        if not cfg["psd_path"] or not os.path.exists(cfg["psd_path"]) or \
           not cfg["xlsx_path"] or not os.path.exists(cfg["xlsx_path"]):
            messagebox.showerror("错误", "请先选择 PSD 与 Excel，并点击【加载】。")
            return
        if not cfg["out_dir"]:
            messagebox.showerror("错误", "请先选择输出目录（预览将生成在 输出目录/_preview 下，不污染正式输出）。")
            return
        # Stage 4.5：Preview 隔离 —— base = out_dir/_preview（与 Batch 同一 resolver，仅 base 不同）
        preview_base = os.path.join(cfg["out_dir"], "_preview")
        os.makedirs(preview_base, exist_ok=True)
        self._log_gui("生成预览（第1行数据）...")
        # 预览只做第一行（Stage 4：统一入口 load_excel_dataset，用 valid_rows[0]）
        try:
            ds = load_excel_dataset(
                cfg["xlsx_path"],
                has_header=cfg.get("has_header", True),
                col_store=cfg["col_store"],
                col_name=cfg["col_name"],
                col_phone=cfg["col_phone"],
                col_role=cfg["col_role"] if cfg.get("col_role", -1) >= 0 else None,
            )
            if not ds.valid_rows:
                self._log_gui("Excel 中没有可生成的有效数据。")
                return
            r = ds.valid_rows[0]
            # Stage 4.5：Preflight —— 分组功能已启用但分组列无效时，阻止 Preview
            group_enabled = bool(cfg.get("group_output_enabled"))
            group_column = cfg.get("group_output_column")
            if group_enabled:
                try:
                    assert_group_column_valid(ds.max_columns, True, group_column)
                except OutputPathError as e:
                    messagebox.showerror("分组配置错误", str(e))
                    return
                # 用整批数据建 map：Preview 目录与 Batch 完全一致（碰撞后缀相同）
                folder_map = build_group_folder_map(ds.valid_rows, group_column)
            else:
                folder_map = None

            with PhotoshopSession() as ps:
                doc0 = ps.open_document(cfg["psd_path"])
                time.sleep(0.6)

                # Stage 2：运行时 LayerIndex
                index = collect_layer_index(doc0)

                # Stage 3：Logo 运行时数据（与 Batch 同一构造/校验）
                logo_map = _build_logo_mapping(cfg, index)
                logo_err = _validate_runtime_logo(logo_map, index, self._log_gui)
                if logo_err:
                    raise LogoVisibilityError(logo_err)

                # ---- Stage 5：Preview 只调 render_one(preview=True) ----
                # 只负责 ensure fresh/取首行/建 folder map/打开文件/按 RowResult 显示错误。
                # 不再自行 Duplicate / 替换文字 / Logo / SaveAs / Close。
                res = render_one(
                    ps_session=ps,
                    template_doc=doc0,
                    row=r,
                    config={
                        "fmt": cfg.get("fmt", FMT_PNG),
                        "also_png": bool(cfg.get("also_png")),
                        "text_map": cfg.get("text_map", {}),
                        "group_output_enabled": group_enabled,
                        "group_output_column": group_column,
                    },
                    layer_index=index,
                    logo_mapping=logo_map,
                    output_context={
                        "base_dir": preview_base,
                        "folder_map": folder_map,
                    },
                    index=1,
                    preview=True,
                    com_dispatch=win32com.client.Dispatch,
                    log=self._log_gui,
                )
                if res.failed:
                    for e in res.errors:
                        self._log_gui(f"预览失败：{e}")
                    return
                if not res.output_paths:
                    self._log_gui("预览失败：未生成任何输出文件。")
                    return
                p = res.output_paths[0]
                self._log_gui(f"预览已生成：{p}")
                try:
                    os.startfile(p)
                except Exception:
                    pass
        except Exception as e:
            self._log_gui(f"预览失败：{e}")

    def _stop(self):
        if self.running:
            self.stop_flag.set()
            self._log_gui("正在停止...")

    def _on_finish(self):
        self.running = False
        self.btn_run.config(state="normal")
        self.btn_stop.config(state="disabled")
        self.btn_preview.config(state="normal")

    def _update_progress(self, done, total):
        if total > 0:
            self.progress["value"] = int(done / total * 100)
        self.progress_label.config(text=f"{done} / {total}")

    def _log_gui(self, msg):
        self.log.config(state="normal")
        self.log.insert("end", msg + "\n")
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
        """把配置值（LayerRef dict 或旧 name 字符串）转成 GUI 下拉显示 label（未加载 PSD 时原样返回）。"""
        if not v:
            return "（不替换）"
        if self.layer_index is None:
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
    root = tk.Tk()
    try:
        root.tk.call("encoding", "system", "utf-8")
    except Exception:
        pass
    App(root)

    def _on_exit():
        # Stage 1：不再依赖模块级 _PS_LAUNCHED_BY_US 做全局 Quit 判断。
        # Session 的 ownership 只属于各自实例；窗口退出时不做任何跨 Session 的
        # Photoshop Quit / 关闭用户文档动作（避免误伤用户正在使用的 Photoshop）。
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", _on_exit)
    root.mainloop()


if __name__ == "__main__":
    main()
