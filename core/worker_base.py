# -*- coding: utf-8 -*-
"""Stage 6：Worker 基类（单 worker 约束 + 事件队列 + 安全生命周期）。

红线（Stage 6 硬性要求）：
  - 一个 App 同时最多一个存活 worker（_worker_alive 检查）；
  - worker 不直接操作 Tk widget —— 只向 event_queue.put(WorkerEvent)；
  - 禁止把 COM object / Tk widget 放入 queue（事件 payload 纯 Python）；
  - worker 线程内自行 CoInitialize/CoUninitialize（子类负责 COM 部分）；
  - cancel 使用协作式 threading.Event（不 kill thread / 不 kill Photoshop）；
  - 窗口关闭：closing flag → cancel → 等待 worker 自然退出（WORKER_DONE）
    → 才允许 destroy root。

本模块与 Tk 解耦，方便纯逻辑测试。
"""
from __future__ import annotations

import queue
import threading
from typing import Optional

from core.task_events import WorkerEvent, AppState, is_valid_transition


class WorkerAlreadyRunningError(RuntimeError):
    """已有 worker 存活时启动新任务。"""


class WorkerBase:
    """泛型 worker：子类实现 _run_task(cfg) 做真实工作。"""

    # 单 worker 约束：一个实例最多一个存活线程
    _worker_lock = threading.Lock()

    def __init__(self):
        self.event_queue: "queue.Queue[WorkerEvent]" = queue.Queue()
        self.cancel_event = threading.Event()
        self._thread: Optional[threading.Thread] = None
        self._closing = False
        self._worker_done_event = threading.Event()

    # ---------------- 状态查询 ----------------
    @property
    def worker_alive(self) -> bool:
        t = self._thread
        return t is not None and t.is_alive()

    @property
    def is_closing(self) -> bool:
        return self._closing

    # ---------------- 任务启动 ----------------
    def start_task(self, task_cfg: dict) -> None:
        """启动一个新任务。已有 worker 存活则拒绝（Stage 6 #6）。"""
        with self._worker_lock:
            if self.worker_alive:
                raise WorkerAlreadyRunningError(
                    "已有任务正在运行，无法启动新任务。")
            self._closing = False
            self.cancel_event.clear()
            self._worker_done_event.clear()
            t = threading.Thread(
                target=self._run_wrapper, args=(task_cfg,),
                name="task-worker", daemon=False)
            self._thread = t
            t.start()

    def _run_wrapper(self, task_cfg: dict) -> None:
        """worker 线程入口：任务 + 收尾事件（线程内不要碰 Tk）。"""
        try:
            self._run_task(task_cfg)
        except Exception as exc:  # noqa: BLE001 —— worker 兜底
            from core.task_events import (
                ErrorPayload, exception_to_error_payload)
            payload = exception_to_error_payload(exc, operation="worker", fatal=True)
            self.put_event(WorkerEvent(WorkerEvent.ERROR, payload))
        finally:
            # 先清 _thread 再发 WORKER_DONE：main thread 收到事件时
            # worker_alive 已确定为 False（_maybe_destroy 依赖此顺序）。
            with self._worker_lock:
                self._thread = None
            self.put_event(WorkerEvent(WorkerEvent.WORKER_DONE, None))
            self._worker_done_event.set()

    # ---------------- 子类实现 ----------------
    def _run_task(self, task_cfg: dict) -> None:
        """子类实现真实任务；通过 self.put_event 发事件。"""
        raise NotImplementedError

    # ---------------- 事件 ----------------
    def put_event(self, event: WorkerEvent) -> None:
        """向事件队列放事件（worker 线程内调用；只放纯 Python 数据）。"""
        self.event_queue.put(event)

    # ---------------- 协作取消 ----------------
    def request_stop(self) -> None:
        """请求停止（协作式）：设置 cancel_event，不杀线程。"""
        self.cancel_event.set()

    # ---------------- 安全关闭 ----------------
    def request_close(self) -> None:
        """窗口关闭请求：置 closing 标志 + 请求停止，等 worker 自然退出。"""
        self._closing = True
        self.cancel_event.set()

    def wait_worker_done(self, timeout: Optional[float] = None) -> bool:
        """等待 worker 自然退出。返回是否已退出（Stage 6 #16 不无限等）。"""
        return self._worker_done_event.wait(timeout)

    def join(self, timeout: Optional[float] = None) -> None:
        t = self._thread
        if t is not None:
            t.join(timeout)

    def cleanup(self) -> None:
        """收尾：置 closing + cancel + join（不 kill）。"""
        self._closing = True
        self.cancel_event.set()
        self.join(timeout=5.0)
