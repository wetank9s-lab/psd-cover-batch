# -*- coding: utf-8 -*-
"""Stage 6.5：GUI 纯逻辑 ViewModel（不依赖 Tk，可独立测试）。

职责：
  - AppState -> 状态文本 / bootstyle（消费 Stage 6 状态机，不复制状态机）；
  - 门店映射状态 -> success/info/warning/danger；
  - BatchResult -> Summary 展示模型；
  - ProgressEvent -> 显示模型（当前行/百分比）；
  - 搜索过滤（Logo 图层 / 门店列表）+ 清空恢复；
  - selected vs brand 展示状态。

红线：本模块绝不 import tkinter / ttkbootstrap / worker / Queue ——
只做纯数据转换，UI 层消费结果。
"""
from __future__ import annotations

from typing import Any, Dict, List, Optional, Sequence, Tuple

from core.task_events import AppState
from gui_styles import (
    MAPPING_STATUS_AUTO, MAPPING_STATUS_CONFIRMED, MAPPING_STATUS_MISSING,
    MAPPING_STATUS_REVIEW, mapping_status_bootstyle, mapping_status_text,
    BS_INFO, BS_PRIMARY, BS_SUCCESS, BS_WARNING, BS_DANGER, BS_SECONDARY,
)

# ---------------- AppState -> 状态文本 / bootstyle ----------------
# 只读映射：唯一数据源仍是 Stage 6 的 AppState 枚举（不新建第二套视觉状态）
# 规格 6.5B 第 6 节：短文案，Header 只显示一个状态
_STATE_TEXT: Dict[str, str] = {
    AppState.IDLE.value: "待配置",
    AppState.LOADING.value: "正在加载",
    AppState.READY.value: "准备就绪",
    AppState.PREVIEWING.value: "正在预览",
    AppState.RUNNING.value: "正在生成",
    AppState.STOPPING.value: "正在停止",
    AppState.ERROR.value: "需要处理",
}

# AppState -> 状态徽标 bootstyle（低饱和语义色）
_STATE_BOOTSTYLE: Dict[str, str] = {
    AppState.IDLE.value: BS_SECONDARY,
    AppState.LOADING.value: BS_INFO,
    AppState.READY.value: BS_SUCCESS,
    AppState.PREVIEWING.value: BS_INFO,
    AppState.RUNNING.value: BS_PRIMARY,
    AppState.STOPPING.value: BS_WARNING,
    AppState.ERROR.value: BS_DANGER,
}

# AppState -> 进度条模式（LOADING/PREVIEWING 无可计算百分比 -> indeterminate）
_STATE_PROGRESS_MODE: Dict[str, str] = {
    AppState.IDLE.value: "determinate",
    AppState.LOADING.value: "indeterminate",
    AppState.READY.value: "determinate",
    AppState.PREVIEWING.value: "indeterminate",
    AppState.RUNNING.value: "determinate",
    AppState.STOPPING.value: "determinate",
    AppState.ERROR.value: "determinate",
}


def state_text(state: AppState) -> str:
    """AppState -> 展示状态文本。"""
    return _STATE_TEXT.get(state.value, state.value)


def state_bootstyle(state: AppState) -> str:
    """AppState -> 徽标 bootstyle。"""
    return _STATE_BOOTSTYLE.get(state.value, BS_SECONDARY)


def state_progress_mode(state: AppState) -> str:
    """AppState -> 进度条模式（determinate / indeterminate）。"""
    return _STATE_PROGRESS_MODE.get(state.value, "determinate")


# ---------------- 门店映射状态 ----------------
def mapping_status(store: str, mapped_label: str, auto_matched: bool) -> Tuple[str, str]:
    """门店映射 -> (状态文本, bootstyle)。

    - mapped_label 为空 / （无）  -> ✕ 未映射 (danger)
    - auto_matched               -> ● 手动 (info)
    - 手动确认                   -> ✓ 已匹配 (success)
    返回 (text, bootstyle)，UI 层直接消费。
    """
    if not mapped_label or mapped_label == "（无）":
        return mapping_status_text(MAPPING_STATUS_MISSING), \
            mapping_status_bootstyle(MAPPING_STATUS_MISSING)
    if auto_matched:
        return mapping_status_text(MAPPING_STATUS_AUTO), \
            mapping_status_bootstyle(MAPPING_STATUS_AUTO)
    return mapping_status_text(MAPPING_STATUS_CONFIRMED), \
        mapping_status_bootstyle(MAPPING_STATUS_CONFIRMED)


def mapping_review_needed(store: str, mapped_label: str, auto_matched: bool) -> bool:
    """是否需要确认（review 状态 = 映射为空但可能有候选 / 或候选歧义）。"""
    return (not mapped_label or mapped_label == "（无）")


# ---------------- 搜索过滤（第 13/14 节） ----------------
def filter_items(items: Sequence[str], query: str) -> List[str]:
    """按查询过滤字符串列表（大小写不敏感、子串匹配）；空查询返回全部。"""
    q = (query or "").strip().lower()
    if not q:
        return list(items)
    return [it for it in items if q in it.lower()]


def filter_logo_labels(labels: Sequence[str], query: str,
                       display_path_of: Optional[Dict[str, str]] = None) -> List[str]:
    """按 query 过滤 Logo 图层 label（匹配 label 或完整 display_path）。

    display_path_of: label -> display_path 映射（第 17 节完整路径搜索）。
    只改变 View，不修改 selected_logo_refs / brand refs / store mapping。
    """
    q = (query or "").strip().lower()
    if not q:
        return list(labels)
    out = []
    for lb in labels:
        if q in lb.lower():
            out.append(lb)
            continue
        if display_path_of and display_path_of.get(lb) and \
                q in display_path_of[lb].lower():
            out.append(lb)
    return out


def filter_stores(stores: Sequence[str], query: str) -> List[str]:
    """按 query 过滤门店列表（只过滤显示，不删除 mapping）。"""
    return filter_items(stores, query)


# ---------------- selected vs brand 展示（规格 6.5B 第 13 节：禁止内部术语 BRAND） ----------------
def logo_display_state(label: str, selected: bool, is_brand: bool) -> str:
    """Logo 图层展示状态：selected（参与 Logo 管理）与 brand（固定显示）必须区分。

    返回展示标记（用户语言，不含 BRAND / runtime / LayerRef 等开发术语）：
      - 固定且参与  -> "固定+参与"（每张封面固定显示 + 参与管理）
      - 固定非参与  -> "固定"（仅固定显示）
      - 参与非固定  -> "参与"（仅参与 Logo 管理）
      - 都不是      -> "候选"（仅候选，未参与）
    """
    if is_brand and selected:
        return "固定+参与"
    if is_brand:
        return "固定"
    if selected:
        return "参与"
    return "候选"


# ---------------- 进度显示模型（第 21 节） ----------------
def progress_display(current: int, total: int, phase: str = "") -> Dict[str, Any]:
    """ProgressEvent -> 显示模型。

    返回 {text, percent, phase_text}：
      - text: "78 / 221"
      - percent: 0-100 整数（total<=0 时 0）
      - phase_text: 阶段描述（如 "正在导出 PNG…"）
    """
    if total <= 0:
        percent = 0
    else:
        percent = int(round(current / total * 100))
    return {
        "text": f"{current} / {total}",
        "percent": percent,
        "phase_text": phase or "",
    }


# ---------------- BatchResult -> Summary 模型（第 24 节） ----------------
def batch_summary_model(summary: Optional[Dict[str, Any]]) -> Dict[str, Any]:
    """BatchResult 摘要 -> 展示模型（唯一数据源仍是 Stage 6 BatchResult）。

    返回：
      {success, failed, skipped, duration_seconds, ok, bootstyle, headline,
       rows_failed: [ {excel_row, store, name, errors} ]}
    """
    s = summary or {}
    success = int(s.get("success", 0))
    failed = int(s.get("failed", 0))
    skipped = int(s.get("skipped", 0))
    cancelled = bool(s.get("cancelled", False))
    dur = float(s.get("duration_seconds", 0.0))
    ok = failed == 0 and not cancelled

    rows_failed = []
    for r in s.get("rows", []) or []:
        if r.get("status") == "FAILED":
            rows_failed.append({
                "excel_row": r.get("excel_row"),
                "store": r.get("store", ""),
                "name": r.get("name", ""),
                "errors": list(r.get("errors", []) or []),
            })

    if cancelled:
        headline = "已停止（部分完成）"
        bootstyle = BS_WARNING
    elif ok:
        headline = "批量生成完成"
        bootstyle = BS_SUCCESS
    else:
        headline = f"批量生成完成，但有 {failed} 条失败"
        bootstyle = BS_WARNING

    return {
        "success": success,
        "failed": failed,
        "skipped": skipped,
        "cancelled": cancelled,
        "duration_seconds": dur,
        "ok": ok,
        "headline": headline,
        "bootstyle": bootstyle,
        "rows_failed": rows_failed,
    }


def summary_duration_text(seconds: float) -> str:
    """秒 -> 人类可读耗时（"112 秒" / "2 分 3 秒"）。"""
    secs = int(round(seconds))
    if secs < 60:
        return f"{secs} 秒"
    m, s = divmod(secs, 60)
    if m < 60:
        return f"{m} 分 {s} 秒"
    h, m2 = divmod(m, 60)
    return f"{h} 小时 {m2} 分 {s} 秒"


# ---------------- 配置检查模型（第 18 节） ----------------
def config_check_model(cfg: Optional[Dict[str, Any]], layer_index_loaded: bool,
                       stores: Sequence[str],
                       text_map: Optional[Dict[str, str]] = None,
                       store_logo_map: Optional[Dict[str, str]] = None,
                       ds_valid_rows: Optional[int] = None,
                       group_enabled: bool = False,
                       group_column_label: str = "") -> List[Dict[str, Any]]:
    """配置检查 -> 展示模型：每项 {label, ok, detail, bootstyle}。

    纯展示辅助：不在这里做实际校验（校验仍在业务层），
    只把业务层已有事实翻译为展示模型。
    """
    items: List[Dict[str, Any]] = []
    psd = (cfg or {}).get("psd_path", "")
    xlsx = (cfg or {}).get("xlsx_path", "")
    out_dir = (cfg or {}).get("out_dir", "")
    import os as _os

    items.append({
        "label": "PSD 模板", "ok": bool(psd) and _os.path.exists(psd),
        "detail": _os.path.basename(psd) if psd else "未选择",
        "bootstyle": BS_SUCCESS if psd and _os.path.exists(psd) else BS_DANGER,
    })
    items.append({
        "label": "Excel 数据",
        "ok": bool(xlsx) and _os.path.exists(xlsx),
        "detail": (f"{_os.path.basename(xlsx)} · {ds_valid_rows} 条" if ds_valid_rows
                   else (_os.path.basename(xlsx) if xlsx else "未选择")),
        "bootstyle": BS_SUCCESS if xlsx and _os.path.exists(xlsx) else BS_DANGER,
    })
    items.append({
        "label": "文字映射", "ok": bool(layer_index_loaded),
        "detail": "已加载" if layer_index_loaded else "未加载 PSD",
        "bootstyle": BS_SUCCESS if layer_index_loaded else BS_DANGER,
    })
    # Logo 映射：全部门店都有映射才 ok
    missing = [s for s in stores
               if not (store_logo_map or {}).get(s) or (store_logo_map or {}).get(s) == "（无）"]
    items.append({
        "label": "Logo 映射",
        "ok": not missing,
        "detail": "全部已映射" if not missing else f"{len(missing)} 个门店未映射",
        "bootstyle": BS_SUCCESS if not missing else BS_WARNING,
    })
    items.append({
        "label": "输出目录", "ok": bool(out_dir),
        "detail": out_dir if out_dir else "未选择",
        "bootstyle": BS_SUCCESS if out_dir else BS_DANGER,
    })
    items.append({
        "label": "分组输出",
        "ok": True,
        "detail": (f"按 {group_column_label}" if group_enabled else "关闭"),
        "bootstyle": BS_SUCCESS if group_enabled else BS_SECONDARY,
    })
    return items


# ---------------- 状态机文本 ----------------
def state_line(state: AppState) -> str:
    """返回状态文本（Header 右侧徽标直接用；保留函数兼容旧调用）。"""
    return state_text(state)


# ---------------- LayerRef -> display label（规格 6.5B 第 10 节：禁止 dict 泄露） ----------------
def layer_display_label(value: Any, label_of_id: Optional[Dict[str, str]] = None) -> str:
    """把任意 LayerRef / serialize_ref dict / name 字符串安全转为 GUI 显示 label。

    - LayerRef / dict 含 id   -> label_of_id[id]（若无映射则回退 name）
    - dict 无 id / 字符串      -> name 原样（不再出现 "{'layer_id': ...}" 形态）
    - 解析失败                 -> "（不替换）"

    保证 UI 永不显示 dict / JSON / layer_id / index_path。
    """
    if value is None:
        return ""
    # 已是字符串
    if isinstance(value, str):
        if value.startswith("{") or value.startswith("["):
            return ""          # 防御：历史配置可能直接存了 dict 字符串
        return value
    # dict（serialize_ref 形态）
    if isinstance(value, dict):
        name = str(value.get("name") or "")
        lid = str(value.get("layer_id") or value.get("id") or "")
        if lid and label_of_id and lid in label_of_id:
            return label_of_id[lid]
        return name
    # LayerRef（有 id / name 属性）
    lid = getattr(value, "id", "") or ""
    name = getattr(value, "name", "") or ""
    if lid and label_of_id and lid in label_of_id:
        return label_of_id[lid]
    return name
