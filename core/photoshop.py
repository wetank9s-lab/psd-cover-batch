# -*- coding: utf-8 -*-
"""
photoshop.py —— Photoshop COM 安全资源管理层（Stage 1）。

目标（P0-01 / P0-04）：
  本工具只能管理「自己打开 / 自己 Duplicate 出来的文档」。
  绝不遍历 Photoshop 当前全部 Documents 执行 Close，
  绝不关闭无法证明属于本 Session 的用户文档。

设计要点：
  - PhotoshopSession 以「实例」持有 ownership 状态（app_started_by_tool /
    owned_documents / initial_documents），Session 结束即销毁，不再使用
    模块级粘性全局 bool（旧的 _PS_LAUNCHED_BY_US 被移除）。
  - owned_documents 保存真实的 COM Document 对象引用（不依赖 doc.Name 字符串）。
  - cleanup 只遍历 owned_documents 反向关闭；任何 owned close 失败都会
    继续关闭其余 owned docs，绝不因此波及无关文档。
  - 第一版采用最保守 Quit 策略：**默认完全不调用 app.Quit()**。
    Photoshop 由本工具启动时，工作完成后只关闭 owned docs，
    Photoshop 进程保留运行（用户数据安全优先于内存释放）。
    仅在调用方显式传入 quit_if_owned=True 且满足全部 5 个安全条件时才 Quit。
  - com_retry 只对 pywintypes.com_error 重试（且只重试可重试的 HRESULT），
    参数错误 / AttributeError / TypeError 等编程错误立即抛出，绝不吞掉。
"""

import time

import pythoncom
import pywintypes

# 允许的 Photoshop COM ProgID（按优先级尝试）
_DISPATCH_PROGIDS = ("Photoshop.Application",)

# HRESULT 常量（避免对 pywintypes 内部枚举的依赖）
_E_RPC_SERVER_UNAVAILABLE = -2147023174  # 0x800706BA：RPC 服务器不可用
_RPC_E_SERVERCALL_RETRYLATER = -2147417846  # 0x8001010A：服务器忙，稍后重试
_RPC_E_CALL_REJECTED = -2147418111  # 0x80010001：调用被拒绝
_CO_E_SERVER_EXEC_FAILURE = -2146959355  # 0x80080005：服务器执行失败
_CO_E_SERVER_LOCAL_FAILURE = -2146959343  # 0x80080013：服务器本地失败
_CO_E_SERVER_TIMEOUT = -2146959341  # 0x80080015：服务器超时
_E_ACCESSDENIED = -2147024891  # 0x80070005：拒绝访问（常见于 COM busy 另一形态）
_COM_ADMIN_BUSY = -2146959359  # 0x80080001：COM admin busy（CO_E_ADMIN_BUSY 附近）

# 判定为「可重试的 COM 忙/暂不可用」的 HRESULT 集合
_RETRYABLE_HRESULTS = frozenset({
    _E_RPC_SERVER_UNAVAILABLE,
    _RPC_E_SERVERCALL_RETRYLATER,
    _RPC_E_CALL_REJECTED,
    _CO_E_SERVER_EXEC_FAILURE,
    _CO_E_SERVER_LOCAL_FAILURE,
    _CO_E_SERVER_TIMEOUT,
    _E_ACCESSDENIED,
    _COM_ADMIN_BUSY,
})


class PhotoshopSessionError(Exception):
    """Photoshop 会话层的业务异常（COM 重试耗尽 / 无法连接 / 无法启动等）。"""

    def __init__(self, message, *, operation=None, hr=None, cause=None):
        super().__init__(message)
        self.operation = operation
        self.hr = hr
        self.cause = cause


class PhotoshopComError(PhotoshopSessionError):
    """COM 调用重试耗尽后抛出的明确异常，保留原始 com_error 链。"""


def _hr_of(exc):
    """从 pywintypes.com_error 中提取 HRESULT int；无法提取返回 None。"""
    try:
        return int(exc.args[0]) if exc.args else None
    except (TypeError, ValueError, IndexError):
        return None


def is_retryable_com_error(exc):
    """判断 com_error 是否属于「可重试」的忙/暂不可用错误。

    只对明确的 RPC 忙 / 服务器不可用 / 调用被拒等 HRESULT 重试；
    参数错误（E_INVALIDARG）、用户中断等不可重试的错误立即抛出。
    """
    if not isinstance(exc, pywintypes.com_error):
        return False
    hr = _hr_of(exc)
    if hr is None:
        return False
    return hr in _RETRYABLE_HRESULTS


def com_retry(func, *args, operation_name="", retries=5, delay=0.2, backoff=1.6, **kwargs):
    """统一 COM 调用重试：只对可重试的 com_error 重试。

    参数：
      func           —— 要调用的 COM 方法 / 属性 getter / setter 包装。
      operation_name —— 人类可读操作名（日志 / 异常信息用）。
      retries        —— 最大尝试次数（>=1）。
      delay          —— 首次重试前等待秒数（每次按 backoff 放大）。
      backoff        —— 退避倍数。

    行为：
      - 成功直接返回结果；
      - 可重试 com_error：PumpWaitingMessages() + sleep 后重试；
      - 不可重试 com_error / 其他异常：立即抛出，不重试（不吞编程错误）；
      - 重试耗尽：抛 PhotoshopComError（携带 operation_name、hr、原异常链）。
    """
    if retries < 1:
        retries = 1
    last_exc = None
    last_hr = None
    for attempt in range(1, retries + 1):
        try:
            return func(*args, **kwargs)
        except pywintypes.com_error as exc:
            last_exc = exc
            last_hr = _hr_of(exc)
            if not is_retryable_com_error(exc):
                # 不可重试：立即抛出，保持原始异常（不含重试层）
                raise
            if attempt >= retries:
                break
            pythoncom.PumpWaitingMessages()
            time.sleep(delay * (backoff ** (attempt - 1)))
    op = operation_name or getattr(func, "__name__", "com_call")
    raise PhotoshopComError(
        f"Photoshop COM 操作「{op}」在 {retries} 次尝试后仍失败"
        + (f"（HRESULT={last_hr:#x}）" if last_hr is not None else ""),
        operation=op,
        hr=last_hr,
        cause=last_exc,
    ) from last_exc


class PhotoshopSession:
    """一次 Photoshop COM 会话的资源管理上下文。

    用法（with 进入 / 退出，退出自动兜底清理 owned docs）：

        with PhotoshopSession() as ps:
            template = ps.open_document(psd_path)
            dup = ps.duplicate_document(template)
            try:
                ...
            finally:
                ps.close_owned_document(dup)

    实例属性：
      app                 —— 当前连接的 Photoshop COM Application 对象
      app_started_by_tool —— 本次 Session 是否「由本工具启动」了 Photoshop
      owned_documents     —— 本 Session 打开 / Duplicate 出的 Document 引用列表
      initial_documents   —— 连接成功瞬间 Photoshop 已存在的 Documents 快照
                            （仅用于 Quit 安全判断，绝不用于 close）
    """

    def __init__(self, *, retries=5, delay=0.2, backoff=1.6,
                 quit_if_owned=False, display_dialogs=3):
        self._retries = retries
        self._delay = delay
        self._backoff = backoff
        # quit_if_owned=True 时，若满足全部安全条件（见 _should_quit），退出时 Quit。
        # 默认 False：第一版保守策略，即使 Photoshop 由本工具启动也不 Quit。
        self._quit_if_owned = quit_if_owned
        self._display_dialogs = display_dialogs

        self.app = None
        self.app_started_by_tool = False
        self.owned_documents = []
        self.initial_documents = []
        self._closed_in_exit = False  # 防止 __exit__ 重复清理（嵌套调用 close_all_owned）

    # ------------------------------------------------------------------
    # 上下文管理
    # ------------------------------------------------------------------
    def __enter__(self):
        # Session 统一负责 COM 初始化/反初始化，保证 com_retry 的
        # PumpWaitingMessages 在已初始化的 COM 线程中工作。
        self._co_initialize()
        self.attach_or_start()
        return self

    def __exit__(self, exc_type, exc, tb):
        try:
            self.close_all_owned()
        finally:
            self.maybe_quit_owned_app()
            self._co_uninitialize()
        # 不吞异常：让调用方正常处理
        return False

    # ------------------------------------------------------------------
    # 连接 / 启动
    # ------------------------------------------------------------------
    def attach_or_start(self):
        """连接（必要时启动）Photoshop，并记录归属与初始文档快照。

        - 先尝试 GetActiveObject 连接已运行的 Photoshop；
        - 失败（未运行）时用 Dispatch 启动，标记 app_started_by_tool=True；
        - 连接成功瞬间抓取 app.Documents 快照（供 Quit 判断，不用于关闭）。
        """
        app = self._try_get_active()
        if app is None:
            app = com_retry(
                self._dispatch_app,
                operation_name="启动 Photoshop（Dispatch）",
                retries=self._retries, delay=self._delay, backoff=self._backoff,
            )
            self.app_started_by_tool = True
        self.app = app
        try:
            app.DisplayDialogs = self._display_dialogs  # 3=psDisplayNoDialogs
        except pywintypes.com_error:
            # 部分版本 / 只读模式下设置 DisplayDialogs 可能失败，非致命，忽略
            pass
        self._snapshot_initial_documents()

    def _try_get_active(self):
        """尝试连接已运行的 Photoshop；失败返回 None。"""
        import win32com.client
        for progid in _DISPATCH_PROGIDS:
            try:
                app = win32com.client.GetActiveObject(progid)
                if app is not None:
                    return app
            except Exception:
                continue
        return None

    def _dispatch_app(self):
        # 用 win32com.client.Dispatch 启动（与旧实现一致，兼容各 pywin32 版本；
        # 避免依赖 pythoncom.CLSIDFromProgID / CoCreateInstance 的版本差异）。
        import win32com.client
        return win32com.client.Dispatch(_DISPATCH_PROGIDS[0])

    def _snapshot_initial_documents(self):
        """记录连接成功瞬间 Photoshop 已存在的文档列表（Quit 安全判断用）。"""
        try:
            docs = list(self.app.Documents)
        except Exception:
            docs = []
        self.initial_documents = docs

    def _co_initialize(self):
        # 统一用 apartment-threaded 初始化；已初始化时 CoInitialize 会返回
        # S_FALSE（可接受），RPC_E_CHANGED_MODE 则忽略（线程已以其他模式初始化）。
        try:
            pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
        except Exception:
            try:
                pythoncom.CoInitialize()
            except Exception:
                pass

    def _co_uninitialize(self):
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

    # ------------------------------------------------------------------
    # 文档登记 / 打开 / 复制 / 关闭
    # ------------------------------------------------------------------
    def open_document(self, path):
        """打开 PSD；若该文件已在 Photoshop 中打开则复用（不登记 owned），否则登记。

        关键安全点（真实集成发现）：Photoshop 的 app.Open(同路径) 在文件已打开时
        返回的是「同一个 Document 对象」（不是新副本）。若把它登记为 owned，
        Session 退出时会误关用户文档。因此：
          - 打开前文件已在 PS 中（用户文档）→ 直接返回该 doc，**不登记 owned**，
            Session 退出绝不关闭它；
          - 打开前文件不存在 → Open 正常登记 owned，退出时关闭。
        """
        path = str(path)
        # 1) 打开前：在已打开文档中查找同路径（大小写不敏感的绝对路径比对）
        existing = self._find_open_document_by_path(path)
        if existing is not None:
            return existing  # 用户文档：不登记，绝不关闭
        # 2) 未打开：正常打开并登记 owned（打开后 Documents 里出现的就是本工具开的）
        doc = com_retry(
            self.app.Open, path,
            operation_name=f"打开文档 {path}",
            retries=self._retries, delay=self._delay, backoff=self._backoff,
        )
        self.owned_documents.append(doc)
        return doc

    def _find_open_document_by_path(self, path):
        """在 Photoshop 已打开文档中，按「Path + Name」（大小写不敏感）查找同名文档。

        返回 Document 或 None。
        说明：实测 Document.FullPath 属性在本机 Photoshop 不可用（AttributeError），
        但 Path（目录，含结尾分隔符）与 Name（文件名）可用，拼接后比对即可。
        """
        import os
        norm = os.path.normcase(os.path.abspath(path))
        try:
            docs = list(self.app.Documents)
        except Exception:
            return None
        for d in docs:
            try:
                dname = str(d.Name)
            except Exception:
                continue
            try:
                dpath = str(getattr(d, "Path", "")) or ""
            except Exception:
                dpath = ""
            if dpath:
                cand = os.path.normcase(os.path.abspath(os.path.join(dpath, dname)))
            else:
                cand = os.path.normcase(os.path.abspath(dname))
            if cand == norm:
                return d
        return None

    def duplicate_document(self, doc):
        """Duplicate 文档副本，并登记为本 Session owned document。"""
        dup = com_retry(
            doc.Duplicate,
            operation_name="复制文档 Duplicate",
            retries=self._retries, delay=self._delay, backoff=self._backoff,
        )
        self.owned_documents.append(dup)
        return dup

    def close_owned_document(self, doc):
        """关闭一个 owned document（不保存，Close(2)）；成功后从 owned 列表移除。

        安全约束：
          - 只接受本 Session owned_documents 里的对象；
          - 未登记的对象直接忽略（绝不对用户文档执行 Close）；
          - 单个关闭失败抛异常，但由调用方 / close_all_owned 继续清理其余。
        """
        if doc is None:
            return
        if doc not in self.owned_documents:
            return  # 不是本 Session 的文档：不碰
        com_retry(
            doc.Close, 2,
            operation_name="关闭 owned 文档 Close",
            retries=self._retries, delay=self._delay, backoff=self._backoff,
        )
        # 关闭成功后移除
        try:
            self.owned_documents.remove(doc)
        except ValueError:
            pass

    def close_all_owned(self):
        """反向关闭全部 owned documents（兜底）。

        - 只遍历 self.owned_documents，绝不遍历 app.Documents；
        - 单个 close 失败：记录到 errors 并继续关闭其余，随后统一抛出
          PhotoshopSessionError（携带 cause 链），保证「不因一个失败而漏关其余」。
        """
        if self._closed_in_exit:
            return
        self._closed_in_exit = True
        errors = []
        # 反向关闭：后打开的先关（duplicate 先于 template）
        for doc in reversed(list(self.owned_documents)):
            try:
                self.close_owned_document(doc)
            except Exception as exc:
                errors.append((doc, exc))
        if errors:
            first_doc, first_exc = errors[0]
            raise PhotoshopSessionError(
                f"清理 owned 文档时发生 {len(errors)} 个错误（其余 owned 文档已继续关闭）",
                operation="close_all_owned",
                cause=first_exc,
            ) from first_exc

    # ------------------------------------------------------------------
    # Quit 策略（第一版最保守：默认不 Quit）
    # ------------------------------------------------------------------
    def _should_quit(self):
        """是否满足「安全退出 Photoshop」的全部条件。

        仅当调用方显式 quit_if_owned=True 且以下 5 条全部成立时才 Quit：
          1. Photoshop 确认由当前 Session 启动（app_started_by_tool=True）；
          2. 本 Session 所有 owned_documents 已清理完毕；
          3. Photoshop 当前不存在任何非本 Session owned 的 Document；
          4. 不使用模块级全局 bool（归属仅来自本 Session 实例）；
          5. ownership 状态仅属于当前 Session 实例。
        """
        if not self._quit_if_owned:
            return False
        if not self.app_started_by_tool:
            return False
        if self.owned_documents:
            return False
        try:
            current = list(self.app.Documents)
        except Exception:
            return False
        # 当前所有文档都必须属于本 Session 的 initial 集合；若存在任何
        # 非 initial 文档（比如中途别人/用户打开的），一律不 Quit。
        for d in current:
            if d not in self.initial_documents:
                return False
        return True

    def maybe_quit_owned_app(self):
        """在满足全部安全条件时才退出 Photoshop（默认策略下不做任何事）。"""
        if not self._should_quit():
            return
        try:
            com_retry(
                self.app.Quit,
                operation_name="退出 Photoshop Quit",
                retries=self._retries, delay=self._delay, backoff=self._backoff,
            )
        except pywintypes.com_error:
            # Quit 失败不致命：不影响 owned docs 已清理的事实
            pass

    # ------------------------------------------------------------------
    # 便捷属性（只读）
    # ------------------------------------------------------------------
    @property
    def owned_count(self):
        return len(self.owned_documents)
