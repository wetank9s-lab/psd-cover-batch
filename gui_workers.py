# -*- coding: utf-8 -*-
"""Stage 6：GUI 三个长任务 worker（Load / Preview / Batch）。

设计要点（对齐 Stage 6 红线）：
  - 全部 Photoshop COM 都在 worker 线程内执行：worker 线程内
    `with PhotoshopSession()`（内部 CoInitializeEx/CoUninitialize）、
    创建 SaveOptions（经 com_dispatch=win32com.client.Dispatch，即当前线程
    Dispatch —— 绝不跨线程缓存 SaveOptions / Document / Layer）；
  - worker 绝不碰 Tk widget / Tk Variable —— 只向 event_queue.put(WorkerEvent)；
  - Queue 只放纯 Python 数据（str/int/list/dict/dataclass）；
  - Exception 先 string 化再入队（ErrorPayload）；
  - 取消：cancel_event 传给 renderer.run_batch / render_one（协作式）；
  - 输入为快照：task_cfg 由 main thread 收集（_collect_cfg 等），worker 只读；
  - 任务开始前 main thread 已做 worker_alive 检查（本类 start_task 再兜底）。
"""
from __future__ import annotations

import os
import threading
import time
import traceback
from typing import Optional

from core.worker_base import WorkerBase
from core.task_events import (
    WorkerEvent, AppState, ProgressPayload, RowPayload, LoadResult,
    PreviewResult, ErrorPayload, exception_to_error_payload,
    batch_result_to_summary,
)

# 任务类型
TASK_LOAD = "load"
TASK_PREVIEW = "preview"
TASK_BATCH = "batch"


class TaskWorker(WorkerBase):
    """三种任务的统一 worker（单 worker 约束由 WorkerBase 保证）。"""

    def __init__(self):
        super().__init__()
        # 由 main thread 设置的回调钩子（可选；默认全走事件队列）
        self.load_done_cb = None   # load_done_cb(LoadResult)  —— main thread 调用
        self.preview_done_cb = None
        self.batch_done_cb = None

    # ---------------- 三种任务入口（main thread 调用） ----------------
    def run_load(self, cfg: dict) -> None:
        self.start_task({"task_type": TASK_LOAD, "cfg": cfg})

    def run_preview(self, cfg: dict) -> None:
        self.start_task({"task_type": TASK_PREVIEW, "cfg": cfg})

    def run_batch(self, cfg: dict) -> None:
        self.start_task({"task_type": TASK_BATCH, "cfg": cfg})

    # ---------------- worker 线程内执行 ----------------
    def _run_task(self, task_cfg: dict) -> None:
        ttype = task_cfg.get("task_type")
        cfg = task_cfg.get("cfg", {})
        if ttype == TASK_LOAD:
            self._do_load(cfg)
        elif ttype == TASK_PREVIEW:
            self._do_preview(cfg)
        elif ttype == TASK_BATCH:
            self._do_batch(cfg)
        else:
            raise ValueError(f"未知任务类型: {ttype!r}")

    # ---------------- Load worker ----------------
    def _do_load(self, cfg: dict) -> None:
        """Load：解析 PSD 图层 + 读取 Excel 门店（COM 全在 worker 内）。

        返回纯数据 LoadResult（LayerRef 序列化为 dict，绝不返回 COM 对象）。
        """
        import qifang_cover_maker as g
        from core.photoshop import PhotoshopSession
        from core.layer_index import collect_layer_index, serialize_ref
        from core.excel_data import load_excel_dataset, ExcelDataError

        result = LoadResult()
        try:
            psd = cfg.get("psd_path", "")
            xlsx = cfg.get("xlsx_path", "")
            self.put_event(WorkerEvent(WorkerEvent.STATE, AppState.LOADING.value))

            # Excel（纯 Python，无 COM；但为了快照一致性也放 worker 里读）
            try:
                ds = load_excel_dataset(
                    xlsx,
                    has_header=cfg.get("has_header", True),
                    col_store=cfg.get("col_store", 0),
                    col_name=cfg.get("col_name", 1),
                    col_phone=cfg.get("col_phone", 3),
                    col_role=cfg.get("col_role") if cfg.get("col_role", -1) >= 0 else None,
                )
                stores = list(ds.stores)
                max_columns = ds.max_columns
                headers = list(ds.headers)
            except ExcelDataError as e:
                result.ok = False
                result.error = str(e)
                self.put_event(WorkerEvent(WorkerEvent.ERROR,
                                           ErrorPayload(message=str(e),
                                                        operation="load_excel",
                                                        fatal=False)))
                return

            # PSD（COM —— worker 线程内 Session）
            self.put_event(WorkerEvent(WorkerEvent.LOG,
                                       f"正在用 Photoshop 解析 PSD 图层：{os.path.basename(psd)}"))
            with PhotoshopSession() as ps:
                doc = ps.open_document(psd)
                time.sleep(0.6)
                index = collect_layer_index(doc)
                layers = index.layers
                result.ok = True
                result.psd_path = psd
                result.layer_count = len(layers)
                result.text_layers = [r.display_path for r in layers if r.is_text]
                result.layer_refs = [serialize_ref(r) for r in layers]
                # 门店 -> 文字/Logo label 建议（main thread 组装 GUI）
                # 文字 label 用 display_path（同名由 GUI 的 labels() 处理）
                result.text_defaults = self._suggest_text_defaults(index)
                # 注：Logo 推荐 / 门店映射建议由 main thread 基于 layer_refs 重建，
                #     这里只回传原始数据（避免 worker 做 GUI 决策）。
            # 回传 Excel 信息（通过 payload 附带）
            result.store_logo_defaults = {"_stores": stores,
                                          "_max_columns": max_columns,
                                          "_headers": list(headers)}
            self.put_event(WorkerEvent(WorkerEvent.LOAD_DONE, result))
        except Exception as exc:
            result.ok = False
            result.error = str(exc)
            payload = exception_to_error_payload(exc, operation="load", fatal=True)
            self.put_event(WorkerEvent(WorkerEvent.ERROR, payload))
            self.put_event(WorkerEvent(WorkerEvent.LOAD_DONE, result))

    def _suggest_text_defaults(self, index) -> dict:
        """按去空格全局唯一自动建议文字层（与 GUI 原 auto_text 一致）。"""
        out = {}
        for field in ("姓名", "电话", "销售顾问"):
            cands = index.find_matching(field)
            if len(cands) == 1:
                out[field] = cands[0].display_path
        return out

    # ---------------- Preview worker ----------------
    def _do_preview(self, cfg: dict) -> None:
        """Preview：整个 Photoshop 生命周期都在 worker 内。

        main thread 负责 os.startfile（用户可见动作）。
        """
        from core.renderer import render_one
        from core.photoshop import PhotoshopSession
        from core.layer_index import collect_layer_index
        from core.excel_data import load_excel_dataset, ExcelDataError
        import qifang_cover_maker as g

        result = PreviewResult()
        try:
            self.put_event(WorkerEvent(WorkerEvent.STATE, AppState.PREVIEWING.value))
            self.put_event(WorkerEvent(WorkerEvent.LOG, "生成预览（第1行数据）..."))
            psd = cfg.get("psd_path", "")
            xlsx = cfg.get("xlsx_path", "")
            out_dir = cfg.get("out_dir", "")
            preview_base = os.path.join(out_dir, "_preview")
            os.makedirs(preview_base, exist_ok=True)

            ds = load_excel_dataset(
                xlsx,
                has_header=cfg.get("has_header", True),
                col_store=cfg.get("col_store", 0),
                col_name=cfg.get("col_name", 1),
                col_phone=cfg.get("col_phone", 3),
                col_role=cfg.get("col_role") if cfg.get("col_role", -1) >= 0 else None,
            )
            if not ds.valid_rows:
                result.ok = False
                result.error = "Excel 中没有可生成的有效数据。"
                self.put_event(WorkerEvent(WorkerEvent.ERROR,
                                           ErrorPayload(message=result.error,
                                                        operation="preview", fatal=False)))
                return
            r = ds.valid_rows[0]

            group_enabled = bool(cfg.get("group_output_enabled"))
            group_column = cfg.get("group_output_column")
            folder_map = None
            if group_enabled:
                from core.output_paths import assert_group_column_valid, build_group_folder_map, OutputPathError
                try:
                    assert_group_column_valid(ds.max_columns, True, group_column)
                except OutputPathError as e:
                    result.ok = False
                    result.error = str(e)
                    self.put_event(WorkerEvent(WorkerEvent.ERROR,
                                               ErrorPayload(message=str(e),
                                                            operation="preview_group",
                                                            fatal=False)))
                    return
                folder_map = build_group_folder_map(ds.valid_rows, group_column)

            # COM 生命周期全部在 worker 线程内
            with PhotoshopSession() as ps:
                doc0 = ps.open_document(psd)
                time.sleep(0.6)
                index = collect_layer_index(doc0)
                logo_map = g._build_logo_mapping(cfg, index)
                logo_err = g._validate_runtime_logo(logo_map, index, self._log_event)
                if logo_err:
                    raise g.LogoVisibilityError(logo_err)
                res = render_one(
                    ps_session=ps,
                    template_doc=doc0,
                    row=r,
                    config={
                        "fmt": cfg.get("fmt", "PNG"),
                        "also_png": bool(cfg.get("also_png")),
                        "text_map": cfg.get("text_map", {}),
                        "group_output_enabled": group_enabled,
                        "group_output_column": group_column,
                    },
                    layer_index=index,
                    logo_mapping=logo_map,
                    output_context={"base_dir": preview_base, "folder_map": folder_map},
                    index=1,
                    preview=True,
                    cancel_event=self.cancel_event,
                    com_dispatch=None,   # renderer 用 _default_dispatch（当前线程）
                    log=self._log_event,
                )
            # 结果（COM 已释放）
            if res.failed:
                result.ok = False
                result.errors = list(res.errors)
                result.error = "；".join(res.errors)[:500]
                self.put_event(WorkerEvent(WorkerEvent.PREVIEW_DONE, result))
                return
            if not res.output_paths:
                result.ok = False
                result.error = "预览失败：未生成任何输出文件。"
                self.put_event(WorkerEvent(WorkerEvent.PREVIEW_DONE, result))
                return
            result.ok = True
            result.preview_path = res.output_paths[0]
            self.put_event(WorkerEvent(WorkerEvent.LOG,
                                       f"预览已生成：{result.preview_path}"))
            self.put_event(WorkerEvent(WorkerEvent.PREVIEW_DONE, result))
        except Exception as exc:
            result.ok = False
            result.error = str(exc)
            payload = exception_to_error_payload(exc, operation="preview", fatal=True)
            self.put_event(WorkerEvent(WorkerEvent.ERROR, payload))
            self.put_event(WorkerEvent(WorkerEvent.PREVIEW_DONE, result))

    # ---------------- Batch worker ----------------
    def _do_batch(self, cfg: dict) -> None:
        """Batch：全量渲染（COM 全在 worker 内）。"""
        import qifang_cover_maker as g
        from core.excel_data import load_excel_dataset, ExcelDataError
        from core.photoshop import PhotoshopSession
        from core.layer_index import collect_layer_index
        from core.output_paths import (
            assert_group_column_valid, build_group_folder_map, OutputPathError)

        self.put_event(WorkerEvent(WorkerEvent.STATE, AppState.RUNNING.value))
        self.put_event(WorkerEvent(WorkerEvent.LOG, "开始批量生成..."))
        try:
            # Excel（快照读）
            try:
                dataset = load_excel_dataset(
                    cfg["xlsx_path"],
                    has_header=cfg.get("has_header", True),
                    col_store=cfg["col_store"],
                    col_name=cfg["col_name"],
                    col_phone=cfg["col_phone"],
                    col_role=cfg.get("col_role") if cfg.get("col_role", -1) >= 0 else None,
                )
            except ExcelDataError as e:
                self.put_event(WorkerEvent(WorkerEvent.ERROR,
                                           ErrorPayload(message=str(e),
                                                        operation="batch_excel",
                                                        fatal=False)))
                return
            data = dataset.valid_rows
            self.put_event(WorkerEvent(WorkerEvent.LOG,
                                       f"Excel 读取完成：{len(data)} 行有效数据"
                                       f"（sheet：{dataset.sheet_name}）。"))
            if not data:
                self.put_event(WorkerEvent(WorkerEvent.LOG, "没有可处理的数据，已停止。"))
                return

            out_dir = cfg["out_dir"]
            os.makedirs(out_dir, exist_ok=True)
            group_enabled = bool(cfg.get("group_output_enabled"))
            group_column = cfg.get("group_output_column")
            folder_map = None
            if group_enabled:
                try:
                    assert_group_column_valid(dataset.max_columns, True, group_column)
                except OutputPathError as e:
                    self.put_event(WorkerEvent(WorkerEvent.ERROR,
                                               ErrorPayload(message=str(e),
                                                            operation="batch_group",
                                                            fatal=False)))
                    return
                folder_map = build_group_folder_map(data, group_column)

            # COM 全在 worker 线程内
            self.put_event(WorkerEvent(WorkerEvent.LOG, "正在启动 / 连接 Photoshop ..."))
            with PhotoshopSession() as ps:
                doc0 = ps.open_document(cfg["psd_path"])
                time.sleep(0.6)
                index = collect_layer_index(doc0)
                logo_map = g._build_logo_mapping(cfg, index)
                logo_err = g._validate_runtime_logo(logo_map, index, self._log_event)
                if logo_err:
                    self.put_event(WorkerEvent(WorkerEvent.ERROR,
                                               ErrorPayload(message=logo_err,
                                                            operation="batch_logo",
                                                            fatal=False)))
                    return

                total = len(data)
                t0 = time.time()
                br = g.renderer_run_batch(
                    ps_session=ps,
                    template_doc=doc0,
                    rows=data,
                    config={
                        "fmt": cfg.get("fmt", "PNG"),
                        "also_png": bool(cfg.get("also_png")),
                        "text_map": cfg.get("text_map", {}),
                        "group_output_enabled": group_enabled,
                        "group_output_column": group_column,
                    },
                    layer_index=index,
                    logo_mapping=logo_map,
                    out_dir=out_dir,
                    folder_map=folder_map,
                    cancel_event=self.cancel_event,
                    progress_cb=lambda done, tot: self._emit_progress(done, tot, data),
                    log=self._log_event,
                    com_dispatch=None,   # 导出 SaveOptions 在 worker 线程内创建
                )
                el = time.time() - t0
            # COM 已释放（with 退出）
            summary = batch_result_to_summary(br)
            self.put_event(WorkerEvent(WorkerEvent.BATCH_DONE, summary))
            if br.cancelled:
                self.put_event(WorkerEvent(WorkerEvent.CANCELLED,
                                           f"用户已停止（剩余 {total - len(br.rows)} 张未处理）"))
            # 完成日志
            self.put_event(WorkerEvent(WorkerEvent.LOG,
                                       f"完成！成功 {br.success} 张，失败 {br.failed} 张"
                                       + (f"，跳过 {br.skipped} 张" if br.skipped else "")
                                       + (f"，用户已停止（剩余 {total - len(br.rows)} 张未处理）"
                                          if br.cancelled else "")
                                       + f"，耗时 {el:.1f}s。输出目录：{out_dir}"))
            for r in br.rows:
                if r.failed:
                    self.put_event(WorkerEvent(WorkerEvent.LOG,
                                               f"  [失败 行{r.excel_row}] {'；'.join(r.errors)}"))
        except Exception as exc:
            payload = exception_to_error_payload(exc, operation="batch", fatal=True)
            self.put_event(WorkerEvent(WorkerEvent.ERROR, payload))

    # ---------------- 事件辅助 ----------------
    def _log_event(self, msg: str) -> None:
        self.put_event(WorkerEvent(WorkerEvent.LOG, msg))

    def _emit_progress(self, done: int, total: int, rows) -> None:
        row = rows[done - 1] if done - 1 < len(rows) else None
        p = ProgressPayload(
            current=done, total=total,
            excel_row=row.excel_row if row is not None else 0,
            store=row.store if row is not None else "",
            name=row.name if row is not None else "",
            phase="完成",
        )
        self.put_event(WorkerEvent(WorkerEvent.PROGRESS, p))
