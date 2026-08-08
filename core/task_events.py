# -*- coding: utf-8 -*-
"""Stage 6：GUI Worker 事件模型（纯 Python dataclass，无 Tk / COM 对象）。

红线约束：
  - 事件 payload 只允许 str / int / float / bool / Path / list / dict / dataclass；
  - 禁止把 Photoshop COM 对象（Application / Document / Layer / TextItem /
    SaveOptions / CDispatch 包装）放入 Queue / Event；
  - 禁止把 Tk widget / Tk Variable 放入 Queue / Event；
  - Exception 必须先 string 化（str(exc) / traceback.format_exc()）再入队。

事件类型：
  STATE         AppState 变化（payload: str 状态名）
  LOG           普通日志文本（payload: str）
  PROGRESS      进度（payload: ProgressPayload）
  ROW_STARTED   单行开始（payload: RowPayload）
  ROW_FINISHED  单行完成（payload: RowPayload + 行结果摘要）
  LOAD_DONE     Load 完成（payload: LoadResult 纯数据）
  PREVIEW_DONE  Preview 完成（payload: PreviewResult）
  BATCH_DONE    Batch 完成（payload: BatchResult 纯数据摘要）
  WORKER_DONE   所有任务完成（worker 即将退出；payload: None）
  ERROR         错误（payload: ErrorPayload）
  CANCELLED     用户取消（payload: str 说明）
"""
from __future__ import annotations

import traceback as _traceback
from dataclasses import dataclass, field
from enum import Enum
from typing import Any, Dict, List, Optional


class AppState(Enum):
    """GUI 应用状态机（Stage 6）。"""

    IDLE = "idle"
    LOADING = "loading"
    READY = "ready"
    PREVIEWING = "previewing"
    RUNNING = "running"
    STOPPING = "stopping"
    ERROR = "error"

    @classmethod
    def from_str(cls, s: str) -> "AppState":
        for st in cls:
            if st.value == s:
                return st
        raise ValueError(f"未知 AppState: {s!r}")


# 合法状态转换表（Stage 6 #5：集中管理）
_STATE_TRANSITIONS: Dict[AppState, set] = {
    AppState.IDLE: {AppState.LOADING, AppState.READY, AppState.ERROR},
    AppState.LOADING: {AppState.READY, AppState.ERROR, AppState.IDLE},
    AppState.READY: {AppState.LOADING, AppState.PREVIEWING, AppState.RUNNING,
                     AppState.ERROR, AppState.IDLE},
    AppState.PREVIEWING: {AppState.READY, AppState.ERROR},
    AppState.RUNNING: {AppState.READY, AppState.STOPPING, AppState.ERROR},
    AppState.STOPPING: {AppState.READY, AppState.ERROR},
    AppState.ERROR: {AppState.READY, AppState.IDLE, AppState.ERROR},
}


def is_valid_transition(frm: AppState, to: AppState) -> bool:
    """状态机合法转换校验（纯函数，供测试）。"""
    if frm == to:
        return True
    return to in _STATE_TRANSITIONS.get(frm, set())


@dataclass
class ProgressPayload:
    """进度信息（Stage 6 #10）。phase 为阶段描述（准备/复制模板/写入文字/...）。"""

    current: int = 0
    total: int = 0
    excel_row: int = 0
    store: str = ""
    name: str = ""
    phase: str = ""


@dataclass
class RowPayload:
    """单行开始/完成信息。"""

    excel_row: int = 0
    store: str = ""
    name: str = ""
    index: int = 0          # 批次内序号（1-based）
    total: int = 0
    status: str = ""        # success / failed / skipped / cancelled
    errors: List[str] = field(default_factory=list)
    output_paths: List[str] = field(default_factory=list)


@dataclass
class LoadResult:
    """Load worker 返回的纯数据（Stage 6 #27）。

    禁止包含 Document / Layer COM 对象；LayerRef 序列化为 dict。
    """

    ok: bool = False
    psd_path: str = ""
    layer_count: int = 0
    text_layers: List[str] = field(default_factory=list)      # label 列表
    logo_labels: List[str] = field(default_factory=list)      # label 列表
    store_logo_defaults: Dict[str, str] = field(default_factory=dict)  # store -> label
    text_defaults: Dict[str, str] = field(default_factory=dict)        # 字段 -> label
    brand_defaults: List[str] = field(default_factory=list)
    layer_refs: List[Dict[str, Any]] = field(default_factory=list)     # serialize_ref
    error: str = ""


@dataclass
class PreviewResult:
    """Preview worker 返回的纯数据。"""

    ok: bool = False
    preview_path: str = ""
    error: str = ""
    errors: List[str] = field(default_factory=list)


@dataclass
class ErrorPayload:
    """错误事件（Stage 6 #20/#21）。exception 已 string 化。"""

    message: str = ""
    operation: str = ""
    fatal: bool = False
    traceback_text: str = ""


@dataclass
class WorkerEvent:
    """Queue 中的统一事件（Stage 6 #9）。type ∈ 事件类型常量。"""

    type: str
    payload: Any = None

    # 类型常量
    STATE = "STATE"
    LOG = "LOG"
    PROGRESS = "PROGRESS"
    ROW_STARTED = "ROW_STARTED"
    ROW_FINISHED = "ROW_FINISHED"
    LOAD_DONE = "LOAD_DONE"
    PREVIEW_DONE = "PREVIEW_DONE"
    BATCH_DONE = "BATCH_DONE"
    WORKER_DONE = "WORKER_DONE"
    ERROR = "ERROR"
    CANCELLED = "CANCELLED"


def exception_to_error_payload(exc: BaseException, operation: str = "",
                               fatal: bool = True) -> ErrorPayload:
    """把 Exception 转成可入队的 ErrorPayload（exception 绝不直接进 Queue）。"""
    return ErrorPayload(
        message=str(exc) or exc.__class__.__name__,
        operation=operation,
        fatal=fatal,
        traceback_text=_traceback.format_exc(),
    )


def batch_result_to_summary(br: Any) -> Dict[str, Any]:
    """把 BatchResult 转成纯 dict 摘要（不进 COM 对象；RowResult 只留纯字段）。"""
    rows = []
    for r in getattr(br, "rows", []) or []:
        rows.append({
            "excel_row": r.excel_row,
            "status": getattr(r, "status", ""),
            "store": r.store,
            "name": r.name,
            "output_paths": list(getattr(r, "output_paths", []) or []),
            "errors": list(getattr(r, "errors", []) or []),
            "warnings": list(getattr(r, "warnings", []) or []),
            "duration_seconds": getattr(r, "duration_seconds", 0.0),
        })
    return {
        "total": getattr(br, "total", 0),
        "success": getattr(br, "success", 0),
        "failed": getattr(br, "failed", 0),
        "skipped": getattr(br, "skipped", 0),
        "cancelled": getattr(br, "cancelled", False),
        "duration_seconds": getattr(br, "duration_seconds", 0.0),
        "rows": rows,
    }
