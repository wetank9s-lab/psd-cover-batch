# -*- coding: utf-8 -*-
"""Stage 6：AppState 状态机 + GUI 层状态流转单元测试。

不启动真实 Tk 主循环 / 不碰真实 Photoshop：
  - AppState 枚举与转换表（IDLE/LOADING/READY/PREVIEWING/RUNNING/STOPPING/ERROR）
  - 每个状态下的控件规则（StateControls）集中管理
  - transition() 非法转换抛 ValueError
  - GUI App 的状态流转（_set_state 校验、控件应用、_on_worker_done、
    _on_load_done/_on_preview_done/_on_batch_done 回 READY、_stop 协作取消、
    _maybe_destroy 安全关闭）
  - worker alive 时不能启动第二任务（WorkerAlreadyRunningError）
"""
import os
import sys
import threading
import time

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import pytest  # noqa: E402

from core.task_events import AppState  # noqa: E402
from core.app_state import (  # noqa: E402
    StateControls, controls_for, transition,
)
from core.worker_base import WorkerAlreadyRunningError, WorkerBase  # noqa: E402
from core.task_events import (  # noqa: E402
    LoadResult, PreviewResult, WorkerEvent,
)


# ---------------------------------------------------------------------------
# 1. AppState 枚举
# ---------------------------------------------------------------------------
def test_state_enum_has_all_members():
    """#1 AppState 覆盖 7 个状态。"""
    assert {s.name for s in AppState} == {
        "IDLE", "LOADING", "READY", "PREVIEWING",
        "RUNNING", "STOPPING", "ERROR",
    }


def test_state_enum_values_unique():
    assert len({s.value for s in AppState}) == len(list(AppState))


def test_state_from_str_roundtrip():
    for s in AppState:
        assert AppState.from_str(s.value) is s


def test_state_from_str_unknown_raises():
    with pytest.raises(ValueError):
        AppState.from_str("nope")


# ---------------------------------------------------------------------------
# 2. 状态转换表
# ---------------------------------------------------------------------------
def test_transitions_idle():
    """#2 IDLE 只能去 LOADING/READY/ERROR。"""
    from core.task_events import is_valid_transition
    assert is_valid_transition(AppState.IDLE, AppState.LOADING)
    assert is_valid_transition(AppState.IDLE, AppState.READY)
    assert is_valid_transition(AppState.IDLE, AppState.ERROR)
    assert not is_valid_transition(AppState.IDLE, AppState.PREVIEWING)
    assert not is_valid_transition(AppState.IDLE, AppState.RUNNING)
    assert not is_valid_transition(AppState.IDLE, AppState.STOPPING)


def test_transitions_loading():
    """#3 LOADING 可去 READY/ERROR/IDLE（Excel 读取失败兜底回 IDLE 重试）。"""
    from core.task_events import is_valid_transition
    assert is_valid_transition(AppState.LOADING, AppState.READY)
    assert is_valid_transition(AppState.LOADING, AppState.ERROR)
    assert is_valid_transition(AppState.LOADING, AppState.IDLE)
    assert not is_valid_transition(AppState.LOADING, AppState.STOPPING)
    assert not is_valid_transition(AppState.LOADING, AppState.RUNNING)


def test_transitions_ready():
    """#4 READY 可再去 LOADING / PREVIEWING / RUNNING / ERROR / IDLE。"""
    from core.task_events import is_valid_transition
    for to in (AppState.LOADING, AppState.PREVIEWING, AppState.RUNNING,
               AppState.ERROR, AppState.IDLE):
        assert is_valid_transition(AppState.READY, to)
    assert not is_valid_transition(AppState.READY, AppState.STOPPING)


def test_transitions_previewing():
    from core.task_events import is_valid_transition
    assert is_valid_transition(AppState.PREVIEWING, AppState.READY)
    assert is_valid_transition(AppState.PREVIEWING, AppState.ERROR)
    assert not is_valid_transition(AppState.PREVIEWING, AppState.RUNNING)
    assert not is_valid_transition(AppState.PREVIEWING, AppState.STOPPING)


def test_transitions_running():
    """#5 RUNNING -> READY / STOPPING / ERROR。"""
    from core.task_events import is_valid_transition
    assert is_valid_transition(AppState.RUNNING, AppState.READY)
    assert is_valid_transition(AppState.RUNNING, AppState.STOPPING)
    assert is_valid_transition(AppState.RUNNING, AppState.ERROR)
    assert not is_valid_transition(AppState.RUNNING, AppState.PREVIEWING)
    assert not is_valid_transition(AppState.RUNNING, AppState.LOADING)


def test_transitions_stopping():
    """STOPPING 只能去 READY/ERROR（等待 worker 自然退出）。"""
    from core.task_events import is_valid_transition
    assert is_valid_transition(AppState.STOPPING, AppState.READY)
    assert is_valid_transition(AppState.STOPPING, AppState.ERROR)
    assert not is_valid_transition(AppState.STOPPING, AppState.RUNNING)
    assert not is_valid_transition(AppState.STOPPING, AppState.IDLE)


def test_transitions_error():
    """ERROR 可回 READY / IDLE（继续操作）。"""
    from core.task_events import is_valid_transition
    assert is_valid_transition(AppState.ERROR, AppState.READY)
    assert is_valid_transition(AppState.ERROR, AppState.IDLE)
    assert is_valid_transition(AppState.ERROR, AppState.ERROR)
    assert not is_valid_transition(AppState.ERROR, AppState.RUNNING)


def test_self_transition_allowed():
    from core.task_events import is_valid_transition
    for s in AppState:
        assert is_valid_transition(s, s)


def test_transition_valid():
    assert transition(AppState.IDLE, AppState.LOADING) is AppState.LOADING


def test_transition_invalid_raises():
    """非法转换必须抛 ValueError（不静默）。"""
    with pytest.raises(ValueError):
        transition(AppState.IDLE, AppState.RUNNING)
    with pytest.raises(ValueError):
        transition(AppState.RUNNING, AppState.PREVIEWING)


# ---------------------------------------------------------------------------
# 3. 控件规则
# ---------------------------------------------------------------------------
def test_controls_idle():
    """IDLE：只有 load / 文件选择 / tab 可用。"""
    c = controls_for(AppState.IDLE)
    assert c.load is True
    assert c.preview is False
    assert c.run is False
    assert c.stop is False
    assert c.pick_files is True
    assert c.tabs_enabled is True
    assert c.state_label == "空闲"


def test_controls_loading():
    """LOADING：全部控件禁用（正在解析）。"""
    c = controls_for(AppState.LOADING)
    assert c.load is False
    assert c.preview is False
    assert c.run is False
    assert c.stop is False
    assert c.pick_files is False
    assert c.tabs_enabled is False
    assert c.state_label == "正在加载模板..."


def test_controls_ready():
    """READY：load / preview / run / 文件选择 / tab 可用；stop 禁用。"""
    c = controls_for(AppState.READY)
    assert c.load is True
    assert c.preview is True
    assert c.run is True
    assert c.stop is False
    assert c.pick_files is True
    assert c.tabs_enabled is True
    assert c.state_label == "就绪"


def test_controls_previewing():
    c = controls_for(AppState.PREVIEWING)
    assert c.load is False
    assert c.preview is False
    assert c.run is False
    assert c.stop is False
    assert c.pick_files is False
    assert c.tabs_enabled is False
    assert c.state_label == "正在生成预览..."


def test_controls_running():
    """RUNNING：stop 唯一可用（可取消）。"""
    c = controls_for(AppState.RUNNING)
    assert c.load is False
    assert c.preview is False
    assert c.run is False
    assert c.stop is True
    assert c.pick_files is False
    assert c.tabs_enabled is False
    assert c.state_label == "正在批量生成..."


def test_controls_stopping():
    """STOPPING：全部禁用（正在等待 worker 退出）。"""
    c = controls_for(AppState.STOPPING)
    assert c.load is False
    assert c.preview is False
    assert c.run is False
    assert c.stop is False
    assert c.pick_files is False
    assert c.tabs_enabled is False
    assert "正在停止" in c.state_label


def test_controls_error():
    c = controls_for(AppState.ERROR)
    assert c.load is True
    assert c.preview is False
    assert c.run is False
    assert c.stop is False
    assert c.pick_files is True
    assert c.tabs_enabled is True
    assert c.state_label == "错误"


def test_controls_unknown_state_falls_back_to_error():
    c = controls_for("???")
    assert c is controls_for(AppState.ERROR)


def test_controls_as_dict():
    c = controls_for(AppState.RUNNING)
    d = c.as_dict()
    assert d["load"] is False
    assert d["preview"] is False
    assert d["run"] is False
    assert d["stop"] is True
    assert d["pick_files"] is False
    assert d["tabs_enabled"] is False


def test_controls_are_distinct_instances():
    """每个状态有独立规则实例（修改不互相影响）。"""
    a = controls_for(AppState.IDLE)
    b = controls_for(AppState.READY)
    assert a is not b
    assert a.load is True and b.load is True


# ---------------------------------------------------------------------------
# 4. GUI App 层状态流转（纯逻辑：_set_state / _apply_controls 语义）
# ---------------------------------------------------------------------------
class _FakeBtn:
    """模拟 ttk.Button：记录 state。"""
    def __init__(self, name):
        self.name = name
        self.state_ = "disabled"

    def config(self, **kw):
        if "state" in kw:
            self.state_ = kw["state"]


class _FakeLabel:
    def __init__(self):
        self.text = ""

    def config(self, **kw):
        if "text" in kw:
            self.text = kw["text"]


class _FakeNotebook:
    def __init__(self):
        self.states = []

    def state(self, *args):
        self.states.append(args)


class _FakeProgress:
    def __init__(self):
        self.value = 0

    def __setitem__(self, k, v):
        if k == "value":
            self.value = v


class _FakeApp:
    """模拟 App：只保留状态机所需的最小字段（不创建 Tk root）。"""

    def __init__(self):
        from core.app_state import controls_for
        self._state = AppState.IDLE
        self._control_sets = []
        self._logs = []
        self._last_transition = None
        self.btn_load = _FakeBtn("load")
        self.btn_preview = _FakeBtn("preview")
        self.btn_run = _FakeBtn("run")
        self.btn_stop = _FakeBtn("stop")
        self.btn_pick_psd = _FakeBtn("pick_psd")
        self.btn_pick_xlsx = _FakeBtn("pick_xlsx")
        self.btn_pick_out = _FakeBtn("pick_out")
        self.notebook = _FakeNotebook()
        self.status_label = _FakeLabel()
        self.progress = _FakeProgress()
        self.progress_label = _FakeLabel()
        self.psd_var = None
        self.xlsx_var = None
        self.out_var = None
        self.layer_index = None

    # ---- 复制 App._set_state / _apply_controls / _log_gui 的语义 ----
    def _log_gui(self, msg):
        self._logs.append(msg)

    def _apply_controls(self, c):
        def _set(btn, enabled):
            btn.config(state="normal" if enabled else "disabled")
        _map = {
            "btn_load": "load", "btn_preview": "preview",
            "btn_run": "run", "btn_stop": "stop",
        }
        for wname, cname in _map.items():
            if hasattr(self, wname):
                _set(getattr(self, wname), getattr(c, cname))
        for var_btn in (getattr(self, "btn_pick_psd", None),
                        getattr(self, "btn_pick_xlsx", None),
                        getattr(self, "btn_pick_out", None)):
            _set(var_btn, c.pick_files)
        try:
            self.notebook.state(["disabled"] if not c.tabs_enabled else ["!disabled"])
        except Exception:
            pass
        if hasattr(self, "status_label"):
            self.status_label.config(text=f"状态：{c.state_label}")

    def _set_state(self, new_state):
        """与 App._set_state 相同：非法转换记录日志但不抛。"""
        from core.task_events import is_valid_transition
        from core.app_state import controls_for
        old = self._state
        if not is_valid_transition(old, new_state):
            self._log_gui(f"[状态] 忽略非法转换 {old.value}->{new_state.value}")
            return
        self._state = new_state
        self._last_transition = (old, new_state)
        c = controls_for(new_state)
        self._apply_controls(c)
        self._control_sets.append(new_state)
        self._log_gui(f"[状态] {old.value} -> {new_state.value}")

    def _on_worker_done(self):
        from core.task_events import AppState
        if self._state in (AppState.LOADING,):
            self._set_state(AppState.READY if self.layer_index is not None else AppState.IDLE)
        elif self._state in (AppState.RUNNING, AppState.STOPPING):
            self._set_state(AppState.READY)
        elif self._state in (AppState.PREVIEWING,):
            self._set_state(AppState.READY)

    def _on_load_done(self, result):
        from core.task_events import AppState
        if not result.ok:
            self._set_state(AppState.IDLE)
            return
        self.layer_index = object()   # 非 None 即可
        self._set_state(AppState.READY)

    def _on_preview_done(self, result):
        from core.task_events import AppState
        self._set_state(AppState.READY)

    def _on_batch_done(self, summary):
        from core.task_events import AppState
        # Batch 单行失败 ≠ GUI ERROR：回 READY
        self._set_state(AppState.READY)


def test_gui_state_initial_idle():
    """GUI 初始状态 = IDLE，控件按 IDLE 规则。"""
    app = _FakeApp()
    assert app._state is AppState.IDLE
    assert app.btn_load.state_ == "disabled"   # 初始 UI 未设置
    app._set_state(AppState.IDLE)              # 幂等：自转换合法
    assert app._state is AppState.IDLE


def test_gui_state_idle_to_loading_controls():
    """#6 IDLE -> LOADING：load/run/preview/stop 全部禁用。"""
    app = _FakeApp()
    app._set_state(AppState.LOADING)
    assert app._state is AppState.LOADING
    assert app.btn_load.state_ == "disabled"
    assert app.btn_preview.state_ == "disabled"
    assert app.btn_run.state_ == "disabled"
    assert app.btn_stop.state_ == "disabled"
    assert app.btn_pick_psd.state_ == "disabled"
    assert "正在加载模板" in app.status_label.text


def test_gui_state_loading_to_ready_controls():
    """#7 LOADING -> READY：load/preview/run 可用，stop 禁用。"""
    app = _FakeApp()
    app._set_state(AppState.LOADING)
    app._set_state(AppState.READY)
    assert app._state is AppState.READY
    assert app.btn_load.state_ == "normal"
    assert app.btn_preview.state_ == "normal"
    assert app.btn_run.state_ == "normal"
    assert app.btn_stop.state_ == "disabled"
    assert "就绪" in app.status_label.text


def test_gui_state_ready_to_running_controls():
    """#8 READY -> RUNNING：stop 可用，其余禁用。"""
    app = _FakeApp()
    app._set_state(AppState.READY)
    app._set_state(AppState.RUNNING)
    assert app._state is AppState.RUNNING
    assert app.btn_stop.state_ == "normal"
    assert app.btn_run.state_ == "disabled"
    assert app.btn_preview.state_ == "disabled"
    assert app.btn_load.state_ == "disabled"


def test_gui_state_running_to_stopping_controls():
    """#9 RUNNING -> STOPPING：stop 禁用（等待退出中）。"""
    app = _FakeApp()
    app._set_state(AppState.READY)
    app._set_state(AppState.RUNNING)
    app._set_state(AppState.STOPPING)
    assert app._state is AppState.STOPPING
    assert app.btn_stop.state_ == "disabled"
    assert "正在停止" in app.status_label.text


def test_gui_state_stopping_to_ready_after_worker_done():
    """#10 STOPPING ->（worker done）-> READY：可再次启动。"""
    app = _FakeApp()
    app._set_state(AppState.READY)
    app._set_state(AppState.RUNNING)
    app._set_state(AppState.STOPPING)
    app._on_worker_done()
    assert app._state is AppState.READY
    assert app.btn_run.state_ == "normal"


def test_gui_state_illegal_transition_logged_not_crash():
    """#11 非法转换（LOADING->RUNNING）记录日志但不崩溃。"""
    app = _FakeApp()
    app._set_state(AppState.LOADING)
    app._set_state(AppState.RUNNING)   # 非法
    assert app._state is AppState.LOADING  # 状态不变
    assert any("忽略非法转换" in m for m in app._logs)


def test_gui_state_preview_done_returns_ready():
    """#12 PREVIEWING -> PREVIEW_DONE -> READY。"""
    app = _FakeApp()
    app._set_state(AppState.READY)
    app._set_state(AppState.PREVIEWING)
    r = PreviewResult(ok=True, preview_path="x.png")
    app._on_preview_done(r)
    assert app._state is AppState.READY


def test_gui_state_batch_done_returns_ready_even_failed_rows():
    """#13 Batch 有失败行 -> BATCH_DONE -> READY（单行失败 ≠ GUI ERROR）。"""
    app = _FakeApp()
    app._set_state(AppState.READY)
    app._set_state(AppState.RUNNING)
    summary = {
        "total": 2, "success": 1, "failed": 1, "skipped": 0,
        "cancelled": False, "duration_seconds": 1.2,
        "rows": [
            {"excel_row": 2, "status": "FAILED", "errors": ["PS 导出失败"]},
        ],
    }
    app._on_batch_done(summary)
    assert app._state is AppState.READY
    assert app.btn_run.state_ == "normal"


def test_gui_state_load_failed_returns_idle():
    """#14 Load 失败 -> ERROR 事件(LOADING->ERROR) -> LOAD_DONE(ok=False) -> IDLE。

    模拟真实事件顺序：_do_load 失败时先发 ERROR(fatal=True) 再发 LOAD_DONE，
    LOADING->ERROR 合法，ERROR->IDLE 合法。
    """
    app = _FakeApp()
    app._set_state(AppState.LOADING)
    app._set_state(AppState.ERROR)      # 模拟 ERROR 事件被处理
    app._on_load_done(LoadResult(ok=False, error="PS 文件无法打开"))
    assert app._state is AppState.IDLE
    assert app.btn_load.state_ == "normal"


def test_gui_state_load_done_returns_ready():
    """#15 Load 成功 -> READY（可 Preview/Run）。"""
    app = _FakeApp()
    app._set_state(AppState.LOADING)
    app._on_load_done(LoadResult(ok=True, psd_path="t.psd", layer_count=3))
    assert app._state is AppState.READY
    assert app.btn_preview.state_ == "normal"
    assert app.btn_run.state_ == "normal"


def test_gui_state_loading_worker_done_without_index_returns_idle():
    """#16 LOADING 状态 worker 退出但无 index（fatal=False 错误路径）-> IDLE 兜底。

    真实事件顺序：_do_load 的 ExcelDataError 路径发 ERROR(fatal=False)（状态不变），
    然后 finally 发 WORKER_DONE —— _on_worker_done 见 layer_index=None -> IDLE。
    """
    app = _FakeApp()
    app._set_state(AppState.LOADING)
    # fatal=False 的 ERROR 事件不改变状态（仍在 LOADING）
    app._on_worker_done()               # layer_index 仍为 None
    assert app._state is AppState.IDLE
    assert app.btn_load.state_ == "normal"


def test_gui_state_error_then_recover():
    """#17 ERROR -> READY：load 恢复可用（可重试）。"""
    app = _FakeApp()
    app._set_state(AppState.LOADING)
    app._set_state(AppState.ERROR)
    assert app._state is AppState.ERROR
    assert app.btn_load.state_ == "normal"
    app._set_state(AppState.READY)
    assert app._state is AppState.READY


def test_gui_state_worker_alive_blocks_second_task():
    """#18 worker alive 时不能启动第二任务：start_task 抛 WorkerAlreadyRunningError。"""
    w = WorkerBase()
    w.start_task({"dummy": 1})
    assert w.worker_alive is True
    with pytest.raises(WorkerAlreadyRunningError):
        w.start_task({"dummy": 2})
    w.wait_worker_done(timeout=10)
    w.join(timeout=5)


def test_gui_state_worker_alive_after_done_false():
    """worker 自然退出后 worker_alive 变 False，可再启动。"""
    w = WorkerBase()

    class _DummyWorker(WorkerBase):
        def _run_task(self, cfg):
            pass

    w = _DummyWorker()
    w.start_task({"dummy": 1})
    assert w.wait_worker_done(timeout=10)
    w.join(timeout=5)
    assert w.worker_alive is False
    w.start_task({"dummy": 2})   # 不抛异常
    assert w.worker_alive is True
    w.wait_worker_done(timeout=10)
    w.join(timeout=5)
