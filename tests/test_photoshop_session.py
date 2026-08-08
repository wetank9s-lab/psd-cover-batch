# -*- coding: utf-8 -*-
"""
test_photoshop_session.py —— PhotoshopSession / com_retry 的 fake 层单元测试（Stage 1）。

不依赖真实 Photoshop：使用 FakeApp / FakeDocument 模拟 COM 对象，
验证 Session 的 ownership 语义（只关 owned、用户文档不动、异常兜底、
Session 隔离、保守 Quit 策略、COM retry 行为）。
"""

import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import pywintypes  # noqa: E402

from core.photoshop import (  # noqa: E402
    PhotoshopSession,
    PhotoshopComError,
    PhotoshopSessionError,
    com_retry,
    is_retryable_com_error,
    _RPC_E_SERVERCALL_RETRYLATER,
    _E_RPC_SERVER_UNAVAILABLE,
)


# ---------------------------------------------------------------------------
# Fake 对象
# ---------------------------------------------------------------------------
class FakeDocument:
    """最小 Document fake：记录是否被 Close 过，可 Duplicate，带 Path 属性。"""

    def __init__(self, name="doc", path=None):
        self.Name = name
        self._path = path
        self.closed = False
        self.visible = True
        self._app = None  # 由 LinkedFakeApp 注入，Close 时同步移除

    @property
    def Path(self):
        # 模拟真实 Photoshop：Path 返回目录（含结尾分隔符）；未设路径时返回 ""
        if self._path:
            return os.path.dirname(self._path) + os.sep
        return ""

    def Duplicate(self):
        if self._app is not None:
            return self._app.Duplicate(self)
        d = FakeDocument(name=self.Name + "_copy")
        d.closed = False
        return d

    def Close(self, *args):
        if self._app is not None:
            self._app._sync_close(self)
        else:
            self.closed = True


class FakeApp:
    """最小 Application fake：Documents 集合 + Open/Duplicate/Quit。"""

    def __init__(self, docs=None, started_by_tool=False):
        self._docs = list(docs) if docs else []
        self._closed = []
        self.quit_called = False
        self.started_by_tool = started_by_tool
        self.DisplayDialogs = None
        self._next_id = 1

    @property
    def Documents(self):
        return self._docs

    def Open(self, path):
        d = FakeDocument(name=os.path.basename(str(path)))
        self._docs.append(d)
        return d

    def Duplicate(self, doc):
        d = FakeDocument(name=doc.Name + "_copy")
        self._docs.append(d)
        return d

    def Quit(self):
        self.quit_called = True

    # ---- 关闭一个文档时同步从 Documents 集合移除 ----
    def _sync_close(self, doc):
        if doc in self._docs:
            self._docs.remove(doc)
        doc.closed = True


# 把 FakeDocument.Close 接回集合移除（构造一个两对象互通的组合 fake）
class LinkedFakeApp(FakeApp):
    """FakeApp 增强版：document.Close() 会同步从 Documents 移除（更接近真实 COM）。
    同时模拟真实 Photoshop 行为：Open(同路径) 若文件已在集合中则返回同一对象。
    """

    def _doc_by_path(self, path):
        import os as _os
        norm = _os.path.normcase(_os.path.abspath(str(path)))
        for d in self._docs:
            p = getattr(d, "_path", None)
            if p and _os.path.normcase(_os.path.abspath(str(p))) == norm:
                return d
        return None

    def Open(self, path):
        existing = self._doc_by_path(path)
        if existing is not None:
            return existing  # 模拟真实 PS：返回同一对象
        d = FakeDocument(name=os.path.basename(str(path)))
        d._path = str(path)
        d._app = self
        self._docs.append(d)
        return d

    def Duplicate(self, doc):
        d = FakeDocument(name=doc.Name + "_copy")
        d._app = self
        self._docs.append(d)
        return d

    def _sync_close(self, doc):
        if doc in self._docs:
            self._docs.remove(doc)
        doc.closed = True


def _link_doc_close(app):
    """把 app 中所有 FakeDocument 的 Close 接回集合移除。"""
    for d in app._docs:
        d._app = app


# ---------------------------------------------------------------------------
# com_retry 测试
# ---------------------------------------------------------------------------
class TestComRetry:
    def test_retry_success_after_transient_busy(self):
        calls = {"n": 0}

        def flaky():
            calls["n"] += 1
            if calls["n"] < 3:
                raise pywintypes.com_error(
                    _RPC_E_SERVERCALL_RETRYLATER, "busy", None, 0)
            return "ok"

        r = com_retry(flaky, operation_name="flaky", retries=5, delay=0.0)
        assert r == "ok"
        assert calls["n"] == 3

    def test_non_retryable_error_immediately_raises(self):
        def bad():
            raise pywintypes.com_error(-2147024809, "E_INVALIDARG", None, 0)  # 0x80070057

        try:
            com_retry(bad, operation_name="bad", retries=5, delay=0.0)
            raise AssertionError("should have raised")
        except pywintypes.com_error as e:
            assert e.args[0] == -2147024809  # 原异常直接透传，非包装

    def test_retries_exhausted_raises_photoshop_com_error(self):
        def always_busy():
            raise pywintypes.com_error(_RPC_E_SERVERCALL_RETRYLATER, "busy", None, 0)

        try:
            com_retry(always_busy, operation_name="always_busy",
                      retries=3, delay=0.0)
            raise AssertionError("should have raised")
        except PhotoshopComError as e:
            assert "always_busy" in str(e)
            assert e.operation == "always_busy"
            assert e.cause is not None
            assert e.__cause__ is not None  # 保留原异常链

    def test_non_com_exception_not_retried(self):
        calls = {"n": 0}

        def bad():
            calls["n"] += 1
            raise AttributeError("no such attr")

        try:
            com_retry(bad, operation_name="bad", retries=5, delay=0.0)
            raise AssertionError("should have raised")
        except AttributeError:
            pass
        assert calls["n"] == 1  # 编程错误不重试

    def test_retry_uses_backoff_delay_increasing(self):
        import time
        calls = {"n": 0}

        def flaky():
            calls["n"] += 1
            if calls["n"] < 3:
                raise pywintypes.com_error(_RPC_E_SERVERCALL_RETRYLATER, "busy", None, 0)
            return "ok"

        t0 = time.monotonic()
        com_retry(flaky, operation_name="flaky", retries=5,
                  delay=0.05, backoff=2.0)
        elapsed = time.monotonic() - t0
        # 第一次重试等 0.05，第二次等 0.1 —— 总等待约 0.15s
        assert elapsed >= 0.14
        assert calls["n"] == 3

    def test_retry_respects_retries_one(self):
        calls = {"n": 0}

        def always_busy():
            calls["n"] += 1
            raise pywintypes.com_error(_RPC_E_SERVERCALL_RETRYLATER, "busy", None, 0)

        try:
            com_retry(always_busy, operation_name="x", retries=1, delay=0.0)
            raise AssertionError("should have raised")
        except PhotoshopComError:
            pass
        assert calls["n"] == 1  # retries=1 只尝试 1 次


class TestIsRetryable:
    def test_retryable_busy(self):
        assert is_retryable_com_error(
            pywintypes.com_error(_RPC_E_SERVERCALL_RETRYLATER, "busy", None, 0))
        assert is_retryable_com_error(
            pywintypes.com_error(_E_RPC_SERVER_UNAVAILABLE, "unavail", None, 0))

    def test_not_retryable_invalidarg(self):
        assert not is_retryable_com_error(
            pywintypes.com_error(-2147024809, "E_INVALIDARG", None, 0))

    def test_non_com_error_not_retryable(self):
        assert not is_retryable_com_error(ValueError("nope"))
        assert not is_retryable_com_error(RuntimeError("nope"))


# ---------------------------------------------------------------------------
# PhotoshopSession 测试（fake）
# ---------------------------------------------------------------------------
class TestPhotoshopSessionOwnership:
    def _session_with_fake(self, fake_app, quit_if_owned=False):
        """构造一个直接注入 fake_app 的 Session（绕过真实 COM 连接）。"""
        s = PhotoshopSession(quit_if_owned=quit_if_owned, retries=3, delay=0.0)
        s.app = fake_app
        s.app_started_by_tool = fake_app.started_by_tool
        s.initial_documents = list(fake_app.Documents)
        return s

    # 1. Session open 的 document 被登记
    def test_open_document_registered(self):
        app = LinkedFakeApp()
        s = self._session_with_fake(app)
        doc = s.open_document("C:/t/template.psd")
        assert doc in s.owned_documents
        assert s.owned_count == 1

    # 2. Duplicate 被登记
    def test_duplicate_registered(self):
        app = LinkedFakeApp()
        s = self._session_with_fake(app)
        template = s.open_document("C:/t/template.psd")
        dup = s.duplicate_document(template)
        assert dup in s.owned_documents
        assert template in s.owned_documents
        assert s.owned_count == 2

    # 3. close_owned_document 只关闭 owned doc
    def test_close_owned_only(self):
        app = LinkedFakeApp()
        user_doc = FakeDocument(name="user.psd")
        app._docs.append(user_doc)
        s = self._session_with_fake(app)
        template = s.open_document("C:/t/template.psd")
        s.close_owned_document(template)
        assert template.closed is True
        assert user_doc.closed is False  # 用户文档绝不动

    # 4. 用户已有 doc 不会被关闭（Session 不登记 initial docs）
    def test_initial_user_docs_not_owned(self):
        app = LinkedFakeApp()
        user_doc = FakeDocument(name="user.psd")
        app._docs.append(user_doc)
        s = self._session_with_fake(app)
        assert s.owned_count == 0  # initial docs 未进入 owned
        assert user_doc.closed is False

    # 5. __exit__ 只清理 owned docs
    def test_exit_cleans_only_owned(self):
        app = LinkedFakeApp()
        user_doc = FakeDocument(name="user.psd")
        app._docs.append(user_doc)
        s = self._session_with_fake(app)
        template = s.open_document("C:/t/template.psd")
        dup = s.duplicate_document(template)
        s.__exit__(None, None, None)
        assert dup.closed is True
        assert template.closed is True
        assert user_doc.closed is False  # 用户文档不被波及
        assert s.owned_count == 0

    # 6. 一个 owned doc close 失败，不应因此去关闭 unrelated docs
    def test_close_failure_does_not_touch_unrelated(self):
        app = LinkedFakeApp()
        user_doc = FakeDocument(name="user.psd")
        app._docs.append(user_doc)
        s = self._session_with_fake(app)

        class BoomDoc(FakeDocument):
            def Close(self, *a):
                raise pywintypes.com_error(_RPC_E_SERVERCALL_RETRYLATER, "busy", None, 0)

        boom = BoomDoc(name="boom")
        boom._app = app
        app._docs.append(boom)
        # 手动把 boom 塞进 owned（模拟工具打开的文档）
        s.owned_documents.append(boom)

        # 再开一个正常的 owned doc
        good = s.open_document("C:/t/good.psd")

        try:
            s.close_all_owned()
            raise AssertionError("should have raised")
        except PhotoshopSessionError as e:
            assert "1 个错误" in str(e) or "错误" in str(e)

        # 正常 owned doc 仍被关闭；用户文档未被波及
        assert good.closed is True
        assert user_doc.closed is False

    # 7. Session A started_by_tool 状态不会泄漏到 Session B
    def test_started_by_tool_not_leaked_between_sessions(self):
        appA = LinkedFakeApp(started_by_tool=True)
        appB = LinkedFakeApp(started_by_tool=False)
        sA = self._session_with_fake(appA)
        sB = self._session_with_fake(appB)
        assert sA.app_started_by_tool is True
        assert sB.app_started_by_tool is False
        # 改 A 不影响 B
        sA.app_started_by_tool = False
        assert sB.app_started_by_tool is False
        # B 的归属判断不受 A 影响
        assert sB._should_quit() is False

    # 8. Session 清理后 owned list 正确更新
    def test_owned_list_updates_after_close(self):
        app = LinkedFakeApp()
        s = self._session_with_fake(app)
        t = s.open_document("C:/t/t.psd")
        d1 = s.duplicate_document(t)
        d2 = s.duplicate_document(t)
        assert s.owned_count == 3
        s.close_owned_document(d1)
        assert s.owned_count == 2
        assert d1 not in s.owned_documents
        assert d2 in s.owned_documents
        s.close_owned_document(d2)
        assert s.owned_count == 1
        s.close_owned_document(t)
        assert s.owned_count == 0

    # 9. app 原本存在时不 Quit（quit_if_owned=True 也不 Quit，因为不是工具启动）
    def test_no_quit_when_app_preexisted(self):
        app = LinkedFakeApp(started_by_tool=False)
        s = self._session_with_fake(app, quit_if_owned=True)
        doc = s.open_document("C:/t/t.psd")
        s.close_owned_document(doc)
        assert s._should_quit() is False  # 非工具启动 → 不 Quit
        s.maybe_quit_owned_app()
        assert app.quit_called is False

    # 10. 默认保守策略：app 即使由工具启动也不 Quit
    def test_conservative_no_quit_even_if_started_by_tool(self):
        app = LinkedFakeApp(started_by_tool=True)
        s = self._session_with_fake(app)  # quit_if_owned=False（默认）
        doc = s.open_document("C:/t/t.psd")
        s.close_owned_document(doc)
        assert s._should_quit() is False
        s.maybe_quit_owned_app()
        assert app.quit_called is False

    # 11. quit_if_owned=True 且全部条件满足时 Quit
    def test_quit_when_all_conditions_met(self):
        app = LinkedFakeApp(started_by_tool=True)
        s = self._session_with_fake(app, quit_if_owned=True)
        doc = s.open_document("C:/t/t.psd")
        s.close_owned_document(doc)
        # 当前 Documents 只剩 initial 快照里的内容（空集合）
        assert s._should_quit() is True
        s.maybe_quit_owned_app()
        assert app.quit_called is True

    # 12. quit_if_owned=True 但中途出现非 initial 文档 → 不 Quit
    def test_no_quit_when_user_doc_appears(self):
        app = LinkedFakeApp(started_by_tool=True)
        s = self._session_with_fake(app, quit_if_owned=True)
        doc = s.open_document("C:/t/t.psd")
        s.close_owned_document(doc)
        # 用户/其他进程在 Session 期间打开了一个新文档
        app._docs.append(FakeDocument(name="user_late.psd"))
        assert s._should_quit() is False
        s.maybe_quit_owned_app()
        assert app.quit_called is False

    # 13. close_owned_document 对未登记对象直接忽略
    def test_close_unregistered_ignored(self):
        app = LinkedFakeApp()
        s = self._session_with_fake(app)
        foreign = FakeDocument(name="foreign.psd")
        s.close_owned_document(foreign)  # 未登记 → 不碰
        assert foreign.closed is False

    # 14. __exit__ 兜底：即使中途异常，owned docs 也会被清理
    def test_exit_finally_cleans_owned_on_exception(self):
        app = LinkedFakeApp()
        user_doc = FakeDocument(name="user.psd")
        app._docs.append(user_doc)
        s = self._session_with_fake(app)
        template = s.open_document("C:/t/t.psd")
        dup = s.duplicate_document(template)
        try:
            raise RuntimeError("row failed")
        except RuntimeError:
            s.__exit__(RuntimeError, None, None)
        assert dup.closed is True
        assert template.closed is True
        assert user_doc.closed is False

    # 15. 非 owned 文档调用 close_all_owned 不动（含初始用户文档）
    def test_close_all_owned_leaves_initial_user_docs(self):
        app = LinkedFakeApp()
        user_doc = FakeDocument(name="user.psd")
        app._docs.append(user_doc)
        s = self._session_with_fake(app)
        assert s.initial_documents == [user_doc]
        s.close_all_owned()
        assert user_doc.closed is False

    # 16. 同路径文件已在 PS 打开 → open_document 返回同一对象且不登记 owned
    def test_open_same_path_user_doc_not_owned(self):
        app = LinkedFakeApp()
        user_path = "C:/Users/Administrator/Desktop/template.psd"
        user_doc = app.Open(user_path)  # 模拟用户先打开
        s = self._session_with_fake(app)
        assert s.owned_count == 0
        doc = s.open_document(user_path)
        assert doc is user_doc  # 复用同一对象
        assert s.owned_count == 0  # 不登记 owned
        s.close_all_owned()
        assert user_doc.closed is False  # 用户文档绝不被 Session 关闭

    # 17. 同路径已打开时不重复 Open（调用计数不变）
    def test_open_same_path_does_not_reopen(self):
        app = LinkedFakeApp()
        user_path = "C:/t/template.psd"
        app.Open(user_path)
        s = self._session_with_fake(app)
        doc = s.open_document(user_path)
        assert doc is app._docs[0]
        assert len(app._docs) == 1  # 没有新增副本
