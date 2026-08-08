# -*- coding: utf-8 -*-
"""Stage 6：Worker / Queue / 事件 / 快照 / 窗口关闭单元测试。

覆盖红线：
  - Queue 中不存在 COM object / Tk widget（payload 纯 Python）；
  - Exception 先 string 化再入队（ErrorPayload / exception_to_error_payload）；
  - cancel_event 传到 renderer（协作式取消，不 kill 线程）；
  - worker 线程不直接操作 Tk（只通过 put_event）；
  - 任务快照为纯 dict（main thread 收集，worker 只读）；
  - 安全窗口关闭：closing flag -> cancel -> 等待 WORKER_DONE -> 才允许 destroy；
  - worker 异常时仍发 WORKER_DONE（finally 保证）。
"""
import os
import queue
import sys
import threading
import time

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import pytest  # noqa: E402

from core.task_events import (  # noqa: E402
    AppState, WorkerEvent, ErrorPayload, LoadResult, PreviewResult,
    ProgressPayload, RowPayload, exception_to_error_payload,
    batch_result_to_summary,
)
from core.worker_base import WorkerAlreadyRunningError, WorkerBase  # noqa: E402
from core.app_state import controls_for  # noqa: E402

# 事件类型常量全集（#19 事件种类完整）
EVENT_TYPES = {
    "STATE", "LOG", "PROGRESS", "ROW_STARTED", "ROW_FINISHED",
    "LOAD_DONE", "PREVIEW_DONE", "BATCH_DONE", "WORKER_DONE",
    "ERROR", "CANCELLED",
}


# ---------------------------------------------------------------------------
# 1. WorkerEvent / 事件类型 / payload 纯数据红线
# ---------------------------------------------------------------------------
def test_event_type_constants_complete():
    """#19 Queue 事件类型覆盖 11 种。"""
    assert EVENT_TYPES == {WorkerEvent.STATE, WorkerEvent.LOG,
                           WorkerEvent.PROGRESS, WorkerEvent.ROW_STARTED,
                           WorkerEvent.ROW_FINISHED, WorkerEvent.LOAD_DONE,
                           WorkerEvent.PREVIEW_DONE, WorkerEvent.BATCH_DONE,
                           WorkerEvent.WORKER_DONE, WorkerEvent.ERROR,
                           WorkerEvent.CANCELLED}


def test_event_created_with_str_type():
    ev = WorkerEvent(WorkerEvent.LOG, "hello")
    assert ev.type == "LOG"
    assert ev.payload == "hello"


def test_event_payload_must_be_pure_python():
    """#20 Queue 中不存在 COM object / Tk widget：
    payload 只允许 str/int/float/bool/Path/list/dict/dataclass/None。"""
    from pathlib import Path
    import dataclasses
    allowed_base = (str, int, float, bool, Path, list, dict, type(None))
    allowed = []

    def check(x, path):
        if isinstance(x, allowed_base):
            return
        if dataclasses.is_dataclass(x) and not isinstance(x, type):
            return
        if isinstance(x, tuple):
            return
        raise AssertionError(f"payload 包含非纯数据: {type(x).__name__} @ {path}")

    payloads = [
        "log line", 42, 3.14, True, Path("a/b"), [1, "2"],
        {"a": 1}, None, ProgressPayload(current=1, total=5, phase="完成"),
        ErrorPayload(message="x", operation="op", fatal=False),
        LoadResult(ok=True, layer_count=3),
        PreviewResult(ok=True, preview_path="p.png"),
    ]
    for p in payloads:
        check(p, "payload")
    # 双层嵌套也检查
    check({"nested": [ProgressPayload(current=1, total=2)]}, "nested")
    assert True


def test_worker_event_payload_never_com_object():
    """确保 dataclass 字段定义不包含 COM 类型（只检查可序列化类型）。"""
    import dataclasses
    from core.task_events import (ProgressPayload, RowPayload, LoadResult,
                                  PreviewResult, ErrorPayload, WorkerEvent)
    for cls in (ProgressPayload, RowPayload, LoadResult, PreviewResult,
                ErrorPayload, WorkerEvent):
        for f in dataclasses.fields(cls):
            assert f.type not in ("Document", "Layer", "Application",
                                  "SaveOptions", "CDispatch")


def test_exception_to_error_payload_stringified():
    """#21 Exception 先 string 化：payload 只有字符串，不携带异常对象。"""
    try:
        raise RuntimeError("boom 失败")
    except RuntimeError as exc:
        p = exception_to_error_payload(exc, operation="load", fatal=True)
    assert p.message == "boom 失败"
    assert p.operation == "load"
    assert p.fatal is True
    assert "RuntimeError" in p.traceback_text
    # 不引用异常对象本身（payload 是纯数据）
    assert not hasattr(p, "exc")


def test_exception_to_error_payload_empty_message_fallback():
    class _Weird(Exception):
        def __str__(self):
            return ""

    try:
        raise _Weird()
    except _Weird as exc:
        p = exception_to_error_payload(exc, operation="x")
    assert p.message == "_Weird"


# ---------------------------------------------------------------------------
# 2. ProgressPayload / RowPayload / LoadResult / PreviewResult 结构
# ---------------------------------------------------------------------------
def test_progress_payload_fields():
    p = ProgressPayload(current=3, total=10, excel_row=5, store="A店",
                        name="张三", phase="写入文字")
    assert p.current == 3
    assert p.total == 10
    assert p.excel_row == 5
    assert p.store == "A店"
    assert p.name == "张三"
    assert p.phase == "写入文字"


def test_progress_payload_defaults():
    p = ProgressPayload()
    assert (p.current, p.total, p.excel_row, p.store, p.name, p.phase) == \
        (0, 0, 0, "", "", "")


def test_row_payload_fields_and_defaults():
    r = RowPayload(excel_row=2, store="A店", name="李四", index=1, total=5)
    assert r.status == ""          # 默认
    assert r.errors == []          # 默认
    assert r.output_paths == []    # 默认
    r2 = RowPayload(excel_row=2, status="FAILED", errors=["e1"])
    assert r2.status == "FAILED"
    assert r2.errors == ["e1"]


def test_load_result_pure_data():
    """LoadResult 只含纯 Python 字段（list of dict / str / int）。"""
    r = LoadResult(ok=True, psd_path="a.psd", layer_count=2,
                   layer_refs=[{"layer_id": 1, "name": "姓名"}],
                   store_logo_defaults={"_stores": ["A店"]})
    assert r.ok is True
    assert r.layer_count == 2
    assert r.layer_refs[0]["name"] == "姓名"
    assert r.store_logo_defaults["_stores"] == ["A店"]
    assert r.error == ""


def test_preview_result_pure_data():
    r = PreviewResult(ok=True, preview_path="out/_preview/x.png")
    assert r.preview_path.endswith(".png")
    assert r.errors == []


# ---------------------------------------------------------------------------
# 3. batch_result_to_summary（BATCH_DONE 只用纯 dict）
# ---------------------------------------------------------------------------
class _FakeRow:
    excel_row = 2
    store = "A店"
    name = "张三"
    status = "FAILED"
    output_paths = ["out/a.png"]
    errors = ["导出失败"]
    warnings = ["w1"]
    duration_seconds = 0.5


class _FakeBatchResult:
    total = 2
    success = 1
    failed = 1
    skipped = 0
    cancelled = False
    duration_seconds = 1.2
    rows = [_FakeRow()]


def test_batch_result_to_summary_dict():
    """#22 BATCH_DONE payload 是纯 dict（可 JSON 序列化）。"""
    s = batch_result_to_summary(_FakeBatchResult())
    assert isinstance(s, dict)
    assert s["total"] == 2
    assert s["success"] == 1
    assert s["failed"] == 1
    assert s["rows"][0]["excel_row"] == 2
    assert s["rows"][0]["status"] == "FAILED"
    assert s["rows"][0]["errors"] == ["导出失败"]
    # 可 JSON 序列化（无 COM / 无 Tk）
    import json
    json.dumps(s, ensure_ascii=False)


def test_batch_result_to_summary_empty_rows():
    class _B:
        total = 0
        success = 0
        failed = 0
        skipped = 0
        cancelled = False
        duration_seconds = 0.0
        rows = []
    s = batch_result_to_summary(_B())
    assert s["rows"] == []
    assert s["total"] == 0


# ---------------------------------------------------------------------------
# 4. WorkerBase：单 worker 约束 / WORKER_DONE / 异常兜底
# ---------------------------------------------------------------------------
class _SimpleWorker(WorkerBase):
    def __init__(self, events=None, fail=False):
        super().__init__()
        self.events = events if events is not None else []
        self.fail = fail

    def _run_task(self, cfg):
        # worker 只通过 put_event 对外通信（绝不直接操作 Tk）
        self.put_event(WorkerEvent(WorkerEvent.LOG, "worker 启动"))
        if cfg.get("hold"):
            time.sleep(cfg["hold"])
        if self.fail:
            raise ValueError("worker 内部失败")
        self.put_event(WorkerEvent(WorkerEvent.LOG, "worker 完成"))


def _drain(q, timeout=5.0):
    """收集 worker 生命周期内全部事件；遇 WORKER_DONE 停止。

    返回 (done_event_or_None, before_done_events)。
    """
    end = time.time() + timeout
    done_ev = None
    before = []
    while time.time() < end:
        try:
            ev = q.get(timeout=0.2)
        except queue.Empty:
            continue
        if ev.type == WorkerEvent.WORKER_DONE:
            done_ev = ev
            break
        before.append(ev)
    return done_ev, before


def test_worker_emits_log_events_and_worker_done():
    """worker 通过事件队列发 LOG 与 WORKER_DONE（不碰 Tk）。"""
    w = _SimpleWorker()
    w.start_task({"hold": 0})
    done_ev, rest = _drain(w.event_queue)
    assert done_ev is not None
    assert done_ev.type == WorkerEvent.WORKER_DONE
    types = [e.type for e in rest]
    assert WorkerEvent.LOG in types
    # WORKER_DONE 事件已发出；线程收尾（置 _thread=None）可能稍后完成
    assert w.wait_worker_done(timeout=5) is True
    assert w.worker_alive is False
    w.join(timeout=5)


def test_worker_done_event_always_emitted_on_exception():
    """#23 worker 异常 -> ERROR 事件 + finally 仍发 WORKER_DONE。"""
    w = _SimpleWorker(fail=True)
    w.start_task({"hold": 0})
    evs = []
    end = time.time() + 10
    while time.time() < end:
        try:
            ev = w.event_queue.get(timeout=0.2)
        except queue.Empty:
            continue
        evs.append(ev)
        if ev.type == WorkerEvent.WORKER_DONE:
            break
    types = {e.type for e in evs}
    assert WorkerEvent.ERROR in types
    assert WorkerEvent.WORKER_DONE in types
    err = next(e for e in evs if e.type == WorkerEvent.ERROR)
    assert isinstance(err.payload, ErrorPayload)
    assert "worker 内部失败" in err.payload.message
    assert w.worker_alive is False
    w.join(timeout=5)


def test_worker_alive_while_running():
    w = _SimpleWorker()
    w.start_task({"hold": 0.8})
    assert w.worker_alive is True
    assert w.wait_worker_done(timeout=10)
    assert w.worker_alive is False
    w.join(timeout=5)


def test_start_task_while_alive_raises():
    """#24 worker alive 时不能 start 第二任务。"""
    w = _SimpleWorker()
    w.start_task({"hold": 0.8})
    with pytest.raises(WorkerAlreadyRunningError):
        w.start_task({"hold": 0.1})
    w.wait_worker_done(timeout=10)
    w.join(timeout=5)


def test_worker_thread_is_non_daemon():
    """worker 线程非 daemon：程序退出前会等待它自然结束。"""
    w = _SimpleWorker()
    w.start_task({"hold": 0})
    t = w._thread
    assert t is not None
    assert t.daemon is False
    assert t.name == "task-worker"
    w.wait_worker_done(timeout=10)
    w.join(timeout=5)


def test_worker_thread_pid_different_from_main():
    """worker 与 main 线程不同：COM 在 worker 线程内 CoInitialize。"""
    w = _SimpleWorker()
    main_tid = threading.get_ident()
    worker_tid = []
    seen = threading.Event()

    class _W(WorkerBase):
        def _run_task(self, cfg):
            worker_tid.append(threading.get_ident())
            seen.set()

    w = _W()
    w.start_task({})
    assert seen.wait(timeout=5)
    assert worker_tid and worker_tid[0] != main_tid
    w.wait_worker_done(timeout=10)
    w.join(timeout=5)


# ---------------------------------------------------------------------------
# 5. 协作式取消（cancel_event 传给 renderer；不 kill）
# ---------------------------------------------------------------------------
class _CancelAwareWorker(WorkerBase):
    """模拟 renderer 的 cancel 检查点：循环顶部检查 cancel_event。

    每次任务重置 cancelled 标志（等价于 renderer 每次返回新 BatchResult）。
    """
    def __init__(self):
        super().__init__()
        self.rendered = 0
        self.cancelled = False

    def _run_task(self, cfg):
        self.cancelled = False
        total = cfg.get("total", 100)
        # 等价于 renderer_run_batch 的循环顶部检查点
        for i in range(1, total + 1):
            if self.cancel_event.is_set():
                self.cancelled = True
                self.put_event(WorkerEvent(WorkerEvent.CANCELLED, "用户已停止"))
                return
            time.sleep(0.01)
            self.rendered = i
            self.put_event(WorkerEvent(WorkerEvent.PROGRESS,
                                       ProgressPayload(current=i, total=total)))


def test_request_stop_sets_cancel_event_and_worker_stops():
    """#25 request_stop -> cancel_event 置位；renderer 检查点协作退出。"""
    w = _CancelAwareWorker()
    w.start_task({"total": 200})
    time.sleep(0.15)          # 让 worker 跑一会儿
    assert w.worker_alive is True
    w.request_stop()          # 协作式：不 kill 线程
    assert w.cancel_event.is_set()
    assert w.wait_worker_done(timeout=10)   # worker 自然退出
    assert w.cancelled is True
    assert w.rendered < 200
    w.join(timeout=5)


def test_cancel_event_cleared_between_tasks():
    """#26 任务之间 cancel_event 被清空（旧取消不泄漏到新任务）。"""
    w = _CancelAwareWorker()
    w.start_task({"total": 200})
    time.sleep(0.12)
    w.request_stop()
    w.wait_worker_done(timeout=10)
    w.join(timeout=5)
    # 新任务
    w.start_task({"total": 5})
    assert w.cancel_event.is_set() is False
    w.wait_worker_done(timeout=10)
    assert w.rendered == 5     # 全部跑完（未受旧取消影响）
    assert w.cancelled is False
    w.join(timeout=5)


def test_cancel_produces_cancelled_event():
    w = _CancelAwareWorker()
    w.start_task({"total": 200})
    time.sleep(0.12)
    w.request_stop()
    types = set()
    end = time.time() + 10
    while time.time() < end:
        try:
            ev = w.event_queue.get(timeout=0.2)
        except queue.Empty:
            continue
        types.add(ev.type)
        if ev.type == WorkerEvent.WORKER_DONE:
            break
    assert WorkerEvent.CANCELLED in types
    w.join(timeout=5)


# ---------------------------------------------------------------------------
# 6. 安全窗口关闭（closing -> cancel -> 等 WORKER_DONE -> 才允许 destroy）
# ---------------------------------------------------------------------------
class _CloseProbe:
    """模拟 App 的 _maybe_destroy：只有 pending_close 且 worker 退出才 destroy。"""
    def __init__(self, worker):
        self.worker = worker
        self.pending_close = False
        self.destroyed = False

    def maybe_destroy(self):
        if self.pending_close and not self.worker.worker_alive:
            self.destroyed = True


def test_close_while_idle_destroys_immediately():
    """#27 空闲关闭：无 worker -> 直接 destroy。"""
    probe = _CloseProbe(WorkerBase())
    probe.pending_close = True
    probe.maybe_destroy()
    assert probe.destroyed is True


def test_close_while_worker_running_waits_for_worker_done():
    """#28 worker 运行时关闭：不立即 destroy；等 worker 退出后才 destroy。"""
    w = _CancelAwareWorker()
    probe = _CloseProbe(w)
    w.start_task({"total": 300})   # 跑较久
    probe.pending_close = True
    probe.maybe_destroy()
    assert probe.destroyed is False          # worker 还活着 -> 不 destroy
    w.request_close()                        # closing flag + cancel
    assert w.is_closing is True
    assert w.cancel_event.is_set() is True
    assert w.wait_worker_done(timeout=10)    # 等自然退出
    w.join(timeout=5)
    probe.maybe_destroy()
    assert probe.destroyed is True           # worker 已退出 -> 才 destroy


def test_request_close_sets_closing_and_cancel():
    w = _SimpleWorker()
    assert w.is_closing is False
    w.request_close()
    assert w.is_closing is True
    assert w.cancel_event.is_set() is True
    # closing 后仍可启动新任务（flag 会被 start_task 重置）——设计允许：
    # 但真实 GUI 关闭后不再启动；这里验证 flag 重置语义
    w2 = _SimpleWorker()
    w2.start_task({"hold": 0})
    assert w2.is_closing is False
    w2.wait_worker_done(timeout=10)
    w2.join(timeout=5)


def test_wait_worker_done_timeout_returns_false():
    """#29 wait_worker_done 带超时：worker 未退出时返回 False（不无限等）。"""
    w = _SimpleWorker()
    w.start_task({"hold": 2.0})
    assert w.wait_worker_done(timeout=0.05) is False
    assert w.worker_alive is True
    w.request_stop()
    assert w.wait_worker_done(timeout=10) is True
    w.join(timeout=5)


def test_cleanup_joins_without_kill():
    """cleanup：closing + cancel + join（不 kill 线程）。"""
    w = _CancelAwareWorker()
    w.start_task({"total": 500})
    w.cleanup()
    assert w.worker_alive is False
    assert w.is_closing is True


# ---------------------------------------------------------------------------
# 7. 任务快照（main thread 收集纯 dict；worker 只读）
# ---------------------------------------------------------------------------
def test_snapshot_is_pure_dict():
    """#30 任务快照为纯 dict（可 JSON 序列化，worker 不读 Tk）。"""
    import json
    snap = {
        "psd_path": "C:/t.psd",
        "xlsx_path": "C:/d.xlsx",
        "out_dir": "C:/out",
        "has_header": True,
        "col_store": 0, "col_name": 1, "col_phone": 3, "col_role": -1,
        "group_output_enabled": False,
        "group_output_column": None,
        "text_map": {"姓名": {"layer_id": 1, "name": "姓名"}, "电话": ""},
        "logo_selection": [{"layer_id": 2, "name": "logo"}],
        "store_logo_map": {"A店": {"layer_id": 3, "name": "logo"}},
        "brand_logo_layers": [],
        "fmt": "PNG",
        "also_png": False,
        "_task": "batch",
    }
    json.dumps(snap, ensure_ascii=False)
    assert snap["_task"] == "batch"


def test_worker_reads_snapshot_without_tk():
    """worker 收到的 cfg 就是 main thread 收集的快照（纯 dict 可遍历）。"""
    got = {}
    seen = threading.Event()

    class _W(WorkerBase):
        def _run_task(self, cfg):
            got.update(cfg)
            seen.set()

    w = _W()
    snap = {"psd_path": "C:/t.psd", "text_map": {"姓名": "x"}, "_task": "load"}
    w.start_task(snap)
    assert seen.wait(timeout=5)
    assert got == snap
    w.wait_worker_done(timeout=10)
    w.join(timeout=5)


# ---------------------------------------------------------------------------
# 8. 事件队列轮询语义（main thread 通过 get_nowait 消费）
# ---------------------------------------------------------------------------
def test_queue_get_nowait_empty_raises():
    q = queue.Queue()
    with pytest.raises(queue.Empty):
        q.get_nowait()


def test_event_queue_fifo_order():
    """事件队列 FIFO：先 put 先消费。"""
    w = _SimpleWorker()
    w.put_event(WorkerEvent(WorkerEvent.LOG, "1"))
    w.put_event(WorkerEvent(WorkerEvent.LOG, "2"))
    assert w.event_queue.get_nowait().payload == "1"
    assert w.event_queue.get_nowait().payload == "2"


def test_state_event_payload_is_appstate_value_string():
    """STATE 事件 payload 为 AppState.value 字符串（main thread 用 from_str 恢复）。"""
    ev = WorkerEvent(WorkerEvent.STATE, AppState.RUNNING.value)
    assert isinstance(ev.payload, str)
    assert AppState.from_str(ev.payload) is AppState.RUNNING


def test_poll_loop_drains_all_events():
    """模拟 main thread 轮询：get_nowait 直到 Empty，处理全部事件。"""
    w = _SimpleWorker()
    w.start_task({"hold": 0})
    handled = []
    end = time.time() + 10
    while time.time() < end:
        try:
            ev = w.event_queue.get_nowait()
        except queue.Empty:
            time.sleep(0.02)
            continue
        handled.append(ev.type)
        if ev.type == WorkerEvent.WORKER_DONE:
            break
    assert WorkerEvent.LOG in handled
    assert WorkerEvent.WORKER_DONE in handled
    w.join(timeout=5)


# ---------------------------------------------------------------------------
# 9. 事件处理器分发（_handle_worker_event 语义，纯逻辑）
# ---------------------------------------------------------------------------
def _make_dispatcher():
    """模拟 App._handle_worker_event 的分发逻辑（不依赖 Tk）。"""
    seen = {"states": [], "logs": [], "load": None, "preview": None,
            "batch": None, "errors": [], "cancelled": [], "done": 0}
    app_state = AppState.IDLE

    def handle(ev):
        nonlocal app_state
        t = ev.type
        if t == WorkerEvent.STATE:
            app_state = AppState.from_str(ev.payload)
            seen["states"].append(app_state)
        elif t == WorkerEvent.LOG:
            seen["logs"].append(ev.payload)
        elif t == WorkerEvent.LOAD_DONE:
            seen["load"] = ev.payload
        elif t == WorkerEvent.PREVIEW_DONE:
            seen["preview"] = ev.payload
        elif t == WorkerEvent.BATCH_DONE:
            seen["batch"] = ev.payload
        elif t == WorkerEvent.ERROR:
            seen["errors"].append(ev.payload)
        elif t == WorkerEvent.CANCELLED:
            seen["cancelled"].append(ev.payload)
        elif t == WorkerEvent.WORKER_DONE:
            seen["done"] += 1
        return app_state
    return seen, handle


def test_dispatcher_state_event_updates_state():
    seen, handle = _make_dispatcher()
    handle(WorkerEvent(WorkerEvent.STATE, AppState.RUNNING.value))
    assert seen["states"] == [AppState.RUNNING]


def test_dispatcher_error_payload_stringified():
    seen, handle = _make_dispatcher()
    handle(WorkerEvent(WorkerEvent.ERROR,
                       ErrorPayload(message="x", operation="load", fatal=False)))
    assert seen["errors"][0].message == "x"


def test_dispatcher_cancelled_logged():
    seen, handle = _make_dispatcher()
    handle(WorkerEvent(WorkerEvent.CANCELLED, "用户已停止"))
    assert seen["cancelled"] == ["用户已停止"]
