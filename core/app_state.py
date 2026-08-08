# -*- coding: utf-8 -*-
"""Stage 6：AppState 状态机（纯逻辑，不依赖 Tk，可独立测试）。

集中管理：
  - 状态转换合法性（is_valid_transition）；
  - 每个状态下 GUI 控件的 enabled/disabled 规则（StateControls）；
  - 状态文字（state_label）。

GUI 层 App 通过 _set_state() 调用本模块，由本模块返回控件状态，
GUI 只负责把结果应用到 widget —— worker 绝不直接操作 Tk widget。
"""
from __future__ import annotations

from typing import Dict, List

from core.task_events import AppState, is_valid_transition


class StateControls:
    """某一状态下各控件是否可用。"""

    def __init__(self, *, load: bool, preview: bool, run: bool, stop: bool,
                 pick_files: bool, tabs_enabled: bool, state_label: str):
        self.load = load
        self.preview = preview
        self.run = run
        self.stop = stop
        self.pick_files = pick_files
        self.tabs_enabled = tabs_enabled
        self.state_label = state_label

    def as_dict(self) -> Dict[str, bool]:
        return {
            "load": self.load,
            "preview": self.preview,
            "run": self.run,
            "stop": self.stop,
            "pick_files": self.pick_files,
            "tabs_enabled": self.tabs_enabled,
        }


# 每个状态 -> 控件规则（Stage 6 #5 典型转换）
_CONTROL_RULES: Dict[AppState, StateControls] = {
    AppState.IDLE: StateControls(
        load=True, preview=False, run=False, stop=False,
        pick_files=True, tabs_enabled=True, state_label="空闲"),
    AppState.LOADING: StateControls(
        load=False, preview=False, run=False, stop=False,
        pick_files=False, tabs_enabled=False, state_label="正在加载模板..."),
    AppState.READY: StateControls(
        load=True, preview=True, run=True, stop=False,
        pick_files=True, tabs_enabled=True, state_label="就绪"),
    AppState.PREVIEWING: StateControls(
        load=False, preview=False, run=False, stop=False,
        pick_files=False, tabs_enabled=False, state_label="正在生成预览..."),
    AppState.RUNNING: StateControls(
        load=False, preview=False, run=False, stop=True,
        pick_files=False, tabs_enabled=False, state_label="正在批量生成..."),
    AppState.STOPPING: StateControls(
        load=False, preview=False, run=False, stop=False,
        pick_files=False, tabs_enabled=False,
        state_label="正在停止（将在当前 Photoshop 操作完成后结束）..."),
    AppState.ERROR: StateControls(
        load=True, preview=False, run=False, stop=False,
        pick_files=True, tabs_enabled=True, state_label="错误"),
}


def controls_for(state: AppState) -> StateControls:
    """返回某状态的控件规则（未知状态按 ERROR 保守处理）。"""
    return _CONTROL_RULES.get(state, _CONTROL_RULES[AppState.ERROR])


def transition(from_state: AppState, to_state: AppState) -> AppState:
    """执行状态转换；非法转换抛 ValueError。"""
    if not is_valid_transition(from_state, to_state):
        raise ValueError(
            f"非法状态转换: {from_state.value} -> {to_state.value}")
    return to_state
