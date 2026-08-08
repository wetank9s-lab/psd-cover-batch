# -*- coding: utf-8 -*-
"""Stage 6.5B：GUI 视觉层常量与 helper（theme / 字体体系 / spacing / 语义 bootstyle / section）。

本模块只负责 View 层：
  - 主题选择（ttkbootstrap 1.9.0 litera——克制、低饱和的办公风格，primary 蓝色与 success 绿色天然分层）；
  - 统一字体体系（Microsoft YaHei UI，fallback 由系统处理；不打包第三方字体）；
  - 统一 spacing（PAD_XS~XL，禁止散落随机 padx/pady）；
  - 语义 bootstyle 常量（primary=主操作 / success=成功态 / warning=确认 / danger=停止 / secondary=次要）；
  - Notebook / Tree / 日志字体样式 helper；
  - 无边框 section（标题 + Separator 产生层级，替代 Labelframe 卡片套卡片）；
  - 门店映射状态 -> bootstyle 徽标映射；
  - tooltip / 可滚动框架 helper。

与业务逻辑解耦：本模块不 import qifang_cover_maker / core 业务模块。
"""
from __future__ import annotations

from typing import Any, Callable, List, Optional

# ---------------- 主题 ----------------
# ttkbootstrap 1.9.0 经典浅色办公主题。
# litera primary=#4582ec（蓝，主操作）/ success=#02b875（绿，成功态）——
# 语义分层清晰，避免「高饱和绿成为全局主操作色」。
APP_THEME = "litera"

# ---------------- 统一 spacing ----------------
PAD_XS = 4
PAD_SM = 8
PAD_MD = 12
PAD_LG = 16
PAD_XL = 24

# ---------------- 字体体系（Windows 系统中文 UI 字体，fallback 系统处理） ----------------
FONT_UI = "Microsoft YaHei UI"
FONT_FALLBACK = "Microsoft YaHei"
FONT_MONO = "Consolas"          # 日志等宽（Windows 自带）

# 字号档位（px 近似）
FS_TITLE = 18      # 应用标题 bold
FS_SUBTITLE = 9    # 应用副标题 secondary
FS_TAB = 11        # Tab
FS_SECTION = 11    # Section 标题 bold
FS_BODY = 10       # 正文 / Label / Button / Combobox
FS_HELP = 9        # 辅助说明 secondary
FS_LOG = 9         # 日志等宽

# 组合字体（供直接使用；ttk 风格用 FONT_* 常量注册）
FONT_TITLE = (FONT_UI, FS_TITLE, "bold")
FONT_SUBTITLE = (FONT_UI, FS_SUBTITLE)
FONT_TAB = (FONT_UI, FS_TAB)
FONT_SECTION = (FONT_UI, FS_SECTION, "bold")
FONT_BODY = (FONT_UI, FS_BODY)
FONT_HELP = (FONT_UI, FS_HELP)
FONT_LOG = (FONT_MONO, FS_LOG)
FONT_SECTION_PLAIN = (FONT_UI, FS_SECTION)   # 非 bold 的 section 标题（次要分组）

# ---------------- bootstyle 常量（ttkbootstrap 语义色） ----------------
BS_PRIMARY = "primary"
BS_SUCCESS = "success"
BS_INFO = "info"
BS_WARNING = "warning"
BS_DANGER = "danger"
BS_SECONDARY = "secondary"          # 徽标/次文本（纯色，tb.Label 可用）
BS_OUTLINE = "outline"              # outline 修饰（仅 tb.Button 支持）

# 语义层级（规格 6.5B 第 20 节：主操作/成功/停止/次要严格分层）
BS_MAIN = "primary"                 # 主按钮：开始生成 / 加载并分析
BS_MAIN_OUTLINE = "primary-outline" # 主操作的 outline 变体（预览）
BS_DANGER = "danger"                # 停止 / 错误
BS_OUTLINE_SECONDARY = "secondary-outline"  # 次要操作：选择 / 保存 / 检查 / 清空 / 复制 / 打开目录

# ---------------- 应用标题（顶部 Header，规格 6.5B 第 6 节：只显示一个状态） ----------------
APP_TITLE_TEXT = "PSD 批量封面生成器"
APP_SUBTITLE_TEXT = "批量替换文字、Logo 并导出"

# ---------------- 状态徽标（页面内轻量反馈，避免全部 MessageBox） ----------------
# 门店映射状态 -> (文本, bootstyle) 映射（规格 6.5B 第 14 节）
MAPPING_STATUS_CONFIRMED = "confirmed"    # ✓ 已匹配（success）
MAPPING_STATUS_AUTO = "auto"              # ● 手动（info）
MAPPING_STATUS_REVIEW = "review"          # ⚠ 待确认（warning）
MAPPING_STATUS_MISSING = "missing"        # ✕ 未映射（danger）

_MAPPING_STATUS_STYLE = {
    MAPPING_STATUS_CONFIRMED: ("✓ 已匹配", BS_SUCCESS),
    MAPPING_STATUS_AUTO: ("● 手动", BS_INFO),
    MAPPING_STATUS_REVIEW: ("⚠ 待确认", BS_WARNING),
    MAPPING_STATUS_MISSING: ("✕ 未映射", BS_DANGER),
}


def mapping_status_text(status: str) -> str:
    """返回映射状态徽标文本。"""
    return _MAPPING_STATUS_STYLE.get(status, ("", BS_SECONDARY))[0]


def mapping_status_bootstyle(status: str) -> str:
    """返回映射状态徽标 bootstyle。"""
    return _MAPPING_STATUS_STYLE.get(status, ("", BS_SECONDARY))[1]


def mapping_status_for(store: str, mapped_label: str, is_auto: bool) -> str:
    """纯函数：根据门店映射状态判定视觉状态（供测试）。

    - mapped_label 为空 / "（无）"  -> missing
    - is_auto=True 且 mapped        -> auto（自动匹配）
    - is_auto=False 且 mapped       -> confirmed（用户手动确认）
    """
    if not mapped_label or mapped_label in ("（无）", ""):
        return MAPPING_STATUS_MISSING
    return MAPPING_STATUS_AUTO if is_auto else MAPPING_STATUS_CONFIRMED


# ---------------- 日志等级轻量区分 ----------------
LOG_INFO = "info"
LOG_WARN = "warn"
LOG_ERROR = "error"


def log_level_of(line: str) -> str:
    """轻量日志等级判定（不解析进度；仅用于颜色/前缀区分）。"""
    if not line:
        return LOG_INFO
    low = line.lower()
    if "错误" in line or "失败" in line or "error" in low or "failed" in low:
        return LOG_ERROR
    if "警告" in line or "warning" in low or "⚠" in line:
        return LOG_WARN
    return LOG_INFO


# ---------------- 路径显示 helper（超长路径） ----------------
def shorten_path(path: str, max_len: int = 60) -> str:
    """超长路径压缩显示：保留头部盘符与尾部文件名，中间省略号。

    例如：D:\\very\\long\\...\\output\\001.png
    """
    if not path:
        return ""
    if len(path) <= max_len:
        return path
    head_len = max(6, max_len // 3)
    tail_len = max(10, max_len - head_len - 3)
    head = path[:head_len]
    tail = path[-tail_len:]
    return f"{head}...{tail}"


# ---------------- Tooltip ----------------
class ToolTip:
    """简单 Tooltip：悬停显示文本。关键规则不依赖 tooltip 才能操作。"""

    def __init__(self, widget: Any, text: str, delay_ms: int = 400):
        self.widget = widget
        self.text = text
        self.delay = delay_ms
        self._after_id: Optional[str] = None
        self._tip_win: Any = None
        widget.bind("<Enter>", self._schedule, add="+")
        widget.bind("<Leave>", self._hide, add="+")
        widget.bind("<ButtonPress>", self._hide, add="+")

    def _schedule(self, _event=None):
        self._cancel()
        self._after_id = self.widget.after(self.delay, self._show)

    def _show(self):
        self._cancel()
        if self._tip_win is not None:
            return
        x = self.widget.winfo_rootx() + 16
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 6
        import tkinter as tk
        self._tip_win = tk.Toplevel(self.widget)
        self._tip_win.wm_overrideredirect(True)
        self._tip_win.wm_geometry(f"+{x}+{y}")
        label = tk.Label(
            self._tip_win, text=self.text, justify="left",
            bg="#ffffe0", fg="#333333", relief="solid", borderwidth=1,
            font=FONT_BODY, padx=6, pady=3, wraplength=420)
        label.pack()

    def _hide(self, _event=None):
        self._cancel()
        if self._tip_win is not None:
            self._tip_win.destroy()
            self._tip_win = None

    def _cancel(self):
        if self._after_id is not None:
            try:
                self.widget.after_cancel(self._after_id)
            except Exception:
                pass
            self._after_id = None


def add_tooltip(widget: Any, text: str) -> ToolTip:
    """给 widget 挂 tooltip（返回实例供持有）。"""
    return ToolTip(widget, text)


# ---------------- 可滚动框架 helper（长列表 / 100+ 门店） ----------------
class ScrollableFrame:
    """带垂直滚动条的 Frame（Canvas + 内部 Frame）。"""

    def __init__(self, master: Any, width: int = 0, height: int = 0):
        import tkinter as tk
        from tkinter import ttk
        self.canvas = tk.Canvas(master, width=width, height=height, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(master, orient="vertical", command=self.canvas.yview)
        self.inner = ttk.Frame(self.canvas)
        self._win_id = self.canvas.create_window(
            (0, 0), window=self.inner, anchor="nw")
        self.inner.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        # 滚动轮（Windows）
        self.canvas.bind_all(
            "<MouseWheel>", self._on_mousewheel, add="+")

    def _on_mousewheel(self, event):
        # 仅当鼠标在 canvas 区域内才滚动（避免全局劫持）
        x, y = event.x_root, event.y_root
        try:
            x0 = self.canvas.winfo_rootx()
            y0 = self.canvas.winfo_rooty()
            if x0 <= x <= x0 + self.canvas.winfo_width() and \
               y0 <= y <= y0 + self.canvas.winfo_height():
                self.canvas.yview_scroll(int(-event.delta / 120), "units")
        except Exception:
            pass

    def pack(self, **kwargs):
        # 默认 fill/expand；调用方可覆盖（避免重复关键字）
        kwargs.setdefault("fill", "both")
        kwargs.setdefault("expand", True)
        self.canvas.pack(side="left", **kwargs)
        self.scrollbar.pack(side="right", fill="y")

    def grid(self, **kwargs):
        self.canvas.grid(**kwargs)
        self.scrollbar.grid(row=kwargs.get("row"), column=kwargs.get("column", 0) + 1,
                            sticky="ns")

    def destroy(self):
        try:
            self.canvas.unbind_all("<MouseWheel>")
        except Exception:
            pass
        self.canvas.destroy()
        self.scrollbar.destroy()


def make_scrollable(master: Any, height: int = 0) -> ScrollableFrame:
    """便捷创建可滚动框架。"""
    return ScrollableFrame(master, height=height)


# ---------------- 无边框 Section（规格 6.5B 第 5 节：去掉 Labelframe 卡片套卡片） ----------------
def section(master: Any, title: str, padding: int = PAD_MD):
    """创建无边框分节容器：Section 标题 + 内容 Frame。

    返回 ttk.Frame（无边框）。层级由 标题(bold) + 内容间距 + 外部 Separator 产生，
    不再使用 Labelframe 细灰框。
    """
    from tkinter import ttk
    box = ttk.Frame(master)
    # 标题行（bold，Section 字号）
    ttk.Label(box, text=title, font=FONT_SECTION).pack(anchor="w")
    # 内容容器
    body = ttk.Frame(box)
    body.pack(fill="both", expand=True, pady=(PAD_XS, 0))
    # 返回内容容器；调用方在 body 上放控件
    return body


def section_help(master: Any, text: str) -> ttk.Label:
    """辅助说明（secondary，9px）。"""
    from tkinter import ttk
    lb = ttk.Label(master, text=text, font=FONT_HELP, foreground="#7f8c8d")
    return lb


def make_separator(master: Any, pady: int = PAD_MD) -> ttk.Separator:
    """水平分隔线（Section 之间产生层级，替代 Labelframe 边框）。"""
    from tkinter import ttk
    sep = ttk.Separator(master, orient="horizontal")
    sep.pack(fill="x", pady=pady)
    return sep


# ---------------- Notebook 样式（规格 6.5B 第 7 节：现代导航而非原生小标签） ----------------
def style_notebook(style: Any) -> None:
    """配置 Notebook Tab：字体加大 + 上下 padding + 左右 padding 14~18 + 当前 Tab 更明显。

    不手绘导航；只用 ttk 自带 option 增强。
    """
    try:
        style.configure(
            "S65.TNotebook.Tab",
            font=FONT_TAB,
            padding=(PAD_MD + 4, PAD_SM),          # 左右 16 / 上下 8
        )
        style.map(
            "S65.TNotebook.Tab",
            background=[
                ("selected", "#ffffff"),
                ("!selected", "#eef1f4"),
            ],
            foreground=[
                ("selected", "#2c3e50"),
                ("!selected", "#5d6d7e"),
            ],
            expand=[("selected", (1, 1, 1, 0))],
        )
    except Exception:
        pass


# ---------------- 全局字体注入（规格 6.5B 第 3 节：正文视觉提高 10%~15%） ----------------
def apply_global_fonts(style: Any) -> None:
    """把统一字体体系注入 ttk 默认风格（body/label/button/combobox/entry/checkbutton/radio）。

    只设置字体不设颜色；颜色仍由 theme/bootstyle 决定。
    """
    try:
        for tname in ("TLabel", "TButton", "TCheckbutton", "TRadiobutton",
                      "TCombobox", "TEntry", "TSpinbox", "TNotebook"):
            try:
                style.configure(tname, font=FONT_BODY)
            except Exception:
                pass
    except Exception:
        pass


# ---------------- Tree / List 样式 ----------------
def style_tree(style: Any) -> None:
    """Tree 行高与字体（日志/明细类列表）。"""
    try:
        style.configure("S65.Treeview", font=FONT_BODY, rowheight=26)
        style.configure("S65.Treeview.Heading", font=FONT_SECTION_PLAIN)
    except Exception:
        pass
