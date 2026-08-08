# -*- coding: utf-8 -*-
"""Stage 5：统一 Renderer（Preview / Batch 共用 render_one）单元测试。

不依赖真实 Photoshop：用 FakeDoc / FakeLayer / FakeSession 模拟 COM，
验证 renderer 的语义：
  - RowResult / BatchResult / RowStatus 结构
  - render_one 唯一单行入口（Duplicate→改→导出→finally Close）
  - 文字写入失败 -> FAILED（不静默继续、不导出）
  - 文字 read-back 不一致 -> TextVerificationError -> FAILED
  - Logo read-back 失败 -> 阻止导出
  - 导出文件验证（不存在/0 字节 -> ExportVerificationError）
  - 部分输出失败（PSD 成功 PNG 失败）-> FAILED 且保留已生成文件
  - finally 关 duplicate（各种失败路径）
  - cancel_event 取消
  - Preview 与 Batch 内容一致（仅 base/filename 不同）
  - 单行失败继续下一行（run_batch 汇总）
  - 复用 Stage 4.5 output core（分组目录、碰撞映射）
"""
import os
import sys
import threading
import time

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import pytest  # noqa: E402

from core.excel_data import load_excel_dataset  # noqa: E402
from core.output_paths import (  # noqa: E402
    build_group_folder_map, resolve_output_directory,
)
from core.renderer import (  # noqa: E402
    RowStatus, RowResult, BatchResult,
    RenderError, TextWriteError, TextVerificationError,
    LogoRenderError, ExportVerificationError,
    apply_text_fields, apply_logo_visibility, export_document,
    render_one, run_batch,
)
from core.layer_index import LayerRef  # noqa: E402
from core.logo_mapping import LogoMapping  # noqa: E402
import openpyxl  # noqa: E402


# ---------------------------------------------------------------------------
# Fake 层：FakeLayer / FakeDoc / FakeSession / FakeDispatch
# ---------------------------------------------------------------------------
class _LayersCollection:
    """COM 风格 Layers 集合（1-based 访问）：Count + __call__(i) + __getitem__。

    resolve_layer 用 container.Layers(i)（COM 实时调用）访问，这里两种形式
    都支持，让测试真正走 core.layer_index.resolve_layer 的完整链路。
    """

    def __init__(self, items):
        self._items = list(items)

    @property
    def Count(self):
        return len(self._items)

    def __call__(self, i):
        return self._items[i - 1]          # 1-based（越界抛 IndexError，与 COM 一致）

    def __getitem__(self, i):
        return self._items[i - 1]          # 1-based（resolve_layer 的 fallback 路径）


class FakeLayer:
    """最小图层 fake：支持 TextItem 读写（可配置失败）、Visible、Parent、子组。"""

    def __init__(self, name, layer_id, is_text=False, text=None, parent=None,
                 children=None):
        self.Name = name
        self.id = layer_id
        self.is_text = is_text
        self._text = text or ""
        self._visible = True
        self.Parent = parent
        self._children = list(children or [])   # 组：直接子图层（COM Layers 集合）
        self.fail_write = False       # 写 Contents 时抛异常
        self.fail_read = False        # 读 Contents 时抛异常
        self.fail_visible = False     # 写 Visible 时抛异常
        self.stubborn_visible = False # 写 Visible 不生效（read-back 保持旧值）
        self._text_item = None        # 测试覆盖点：替换 TextItem 行为
        self.contents_written = []    # 记录写入历史（供断言）

    @property
    def Layers(self):
        return _LayersCollection(self._children)

    @property
    def Visible(self):
        return self._visible

    @Visible.setter
    def Visible(self, v):
        if self.fail_visible:
            raise RuntimeError("visible write failed")
        if not self.stubborn_visible:
            self._visible = bool(v)

    @property
    def TextItem(self):
        if not self.is_text:
            raise RuntimeError("not a text layer")   # 模拟 COM 非文字层访问抛异常
        if self._text_item is not None:
            return self._text_item                  # 测试可覆盖（模拟写入未生效等）
        class _TI:
            def __init__(self, layer):
                self._layer = layer
            @property
            def Contents(self):
                if self._layer.fail_read:
                    raise RuntimeError("read failed")
                return self._layer._text
            @Contents.setter
            def Contents(self, v):
                if self._layer.fail_write:
                    raise RuntimeError("write failed")
                self._layer._text = v
                self._layer.contents_written.append(v)
        return _TI(self)


def _mk_ref(name, layer_id, is_text=False, group="文字"):
    return LayerRef(id=layer_id, name=name,
                    display_path=group + " > " + name,
                    index_path=tuple(int(x) for x in layer_id.split("/")),
                    is_group=False, is_text=is_text,
                    name_path=(group, name))


class FakeDoc:
    """最小 Document fake：图层树 + Duplicate/Close/SaveAs（写真实文件）。"""

    def __init__(self, layers=None, name="doc"):
        # 注意：顶层也必须走 COM 风格 Layers 集合
        self._layers = _LayersCollection(layers or [])
        self.Name = name
        self.closed = False
        self.close_calls = 0
        self.saved = []                # (path, fmt)
        self._app = None

    @property
    def Layers(self):
        return self._layers

    @property
    def top_layers(self):
        """顶层 list[FakeLayer]（组）：子类/测试构造嵌套结构用。"""
        return list(getattr(self._layers, "_items", []))

    @property
    def layers(self):
        """扁平 list[FakeLayer]（含组内子层）：测试便利：直接改某个层的 fail 开关。"""
        return self._raw_layers()

    def _raw_layers(self):
        """返回当前树的扁平 list[FakeLayer]（含组内子层）。"""
        out = []

        def walk(items):
            for l in items:
                out.append(l)
                walk(getattr(l, "_children", []))
        walk(getattr(self._layers, "_items", []))
        return out

    def _copy_layer(self, l):
        """深拷贝单个图层（含子组）。"""
        kids = [self._copy_layer(c) for c in getattr(l, "_children", [])]
        nl = FakeLayer(l.Name, l.id, is_text=l.is_text, text=l._text,
                       parent=None, children=kids)
        nl.fail_write = l.fail_write
        nl.fail_read = l.fail_read
        nl.fail_visible = l.fail_visible
        nl.stubborn_visible = l.stubborn_visible
        nl._visible = l._visible
        nl.contents_written = list(l.contents_written)
        # 复制 TextItem 覆盖点（浅拷贝 + 重新绑定到新层）
        if l._text_item is not None:
            import copy
            nl._text_item = copy.copy(l._text_item)
            if hasattr(nl._text_item, "_layer"):
                nl._text_item._layer = nl
        # 子层 Parent 回指
        for c in nl._children:
            c.Parent = nl
        return nl

    def Duplicate(self):
        new_layers = [self._copy_layer(l) for l in self._layers._items]
        d = type(self)(layers=new_layers, name=self.Name + "_copy")
        d._app = self._app
        return d

    def Close(self, *args):
        self.closed = True
        self.close_calls += 1

    def SaveAs(self, path, opt, as_copy=True):
        path = str(path)
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "wb") as f:
            f.write(b"PS" if str(path).endswith(".psd") else b"PNGDATA")
        self.saved.append((path, getattr(opt, "_fmt", None)))
        return None


class FakeApp:
    def __init__(self):
        self.ActiveDocument = None


class FakeSession:
    """最小 PhotoshopSession fake：Duplicate 登记 owned、Close 同步。"""

    def __init__(self, template):
        self.app = FakeApp()
        self.template = template
        template._app = self
        self.owned = []
        self.closed = []

    def duplicate_document(self, doc):
        dup = doc.Duplicate()
        dup._app = self
        self.owned.append(dup)
        return dup

    def close_owned_document(self, doc):
        if doc in self.owned:
            doc.Close()
            self.owned.remove(doc)
            self.closed.append(doc)


class FakeDispatch:
    """SaveOptions fake（记录 fmt，SaveAs 时无需真实 COM）。"""

    def __init__(self, progid):
        self.progid = progid
        self._fmt = "PNG" if "PNG" in progid else ("JPG" if "JPG" in progid else "PSD")


# ---------------------------------------------------------------------------
# fixture：Excel + template + config
# ---------------------------------------------------------------------------
def _make_xlsx(tmp_path, rows, name="data.xlsx"):
    p = tmp_path / name
    wb = openpyxl.Workbook()
    ws = wb.active
    for r in rows:
        ws.append(r)
    wb.save(p)
    return str(p)


def _template():
    """模板：文字组(姓名/电话/销售顾问) + 门店Logo组(康乐家电/诚信电器) + 品牌组(七方Logo)。

    层级：
      [0] 文字        (组)
      [1] 门店Logo    (组)
      [2] 品牌        (组)
    index_path：0/0 姓名, 0/1 电话, 0/2 销售顾问, 1/0 康乐家电, 1/1 诚信电器, 2/0 七方Logo
    """
    text_grp = FakeLayer("文字", "0", children=[
        FakeLayer("姓名", "0/0", is_text=True, text="模板姓名"),
        FakeLayer("电话", "0/1", is_text=True, text="000"),
        FakeLayer("销售顾问", "0/2", is_text=True, text="顾问"),
    ])
    store_grp = FakeLayer("门店Logo", "1", children=[
        FakeLayer("康乐家电", "1/0", is_text=False),
        FakeLayer("诚信电器", "1/1", is_text=False),
    ])
    brand_grp = FakeLayer("品牌", "2", children=[
        FakeLayer("七方Logo", "2/0", is_text=False),
    ])
    for g in (text_grp, store_grp, brand_grp):
        for c in g._children:
            c.Parent = g
    return FakeDoc(layers=[text_grp, store_grp, brand_grp])


def _logo_mapping():
    return LogoMapping(
        store_logo_map={
            "康乐家电": _mk_ref("康乐家电", "1/0", group="门店Logo"),
            "诚信电器": _mk_ref("诚信电器", "1/1", group="门店Logo"),
        },
        brand_logo_refs=[_mk_ref("七方Logo", "2/0", group="品牌")],
        logo_selection_refs=[],
    )


def _text_map():
    return {
        "姓名": _mk_ref("姓名", "0/0", is_text=True),
        "电话": _mk_ref("电话", "0/1", is_text=True),
        "销售顾问": _mk_ref("销售顾问", "0/2", is_text=True),
    }


def _config(fmt="PNG", also_png=False, group_enabled=False, group_column=None):
    cfg = {
        "fmt": fmt,
        "also_png": also_png,
        "group_output_enabled": group_enabled,
        "group_output_column": group_column,
        "text_map": _text_map(),
    }
    return cfg


class _Row:
    """最小 ExcelRow 替身（鸭子类型：excel_row/store/name/phone/role/values）。"""
    def __init__(self, excel_row=2, store="康乐家电", name="张三", phone="138",
                 role="王经理", values=None):
        self.excel_row = excel_row
        self.store = store
        self.name = name
        self.phone = phone
        self.role = role
        # 分组列 0 即门店（与 store 保持一致，供 build_group_folder_map 使用）
        self.values = tuple(values) if values is not None else (store, name, phone, role or "")


def _render(ps, doc, row, cfg, base_dir, folder_map=None, index=1, preview=False,
            cancel=None, com_dispatch=FakeDispatch, log=None, logo_mapping=None,
            layer_index=None):
    return render_one(
        ps_session=ps, template_doc=doc, row=row, config=cfg,
        layer_index=layer_index, logo_mapping=logo_mapping or _logo_mapping(),
        output_context={"base_dir": base_dir, "folder_map": folder_map},
        index=index, preview=preview, cancel_event=cancel,
        com_dispatch=com_dispatch, log=log)


def _by_id(doc, layer_id):
    """扁平找层（按 LayerRef.id）。"""
    for l in doc.layers:
        if l.id == layer_id:
            return l
    raise AssertionError(f"layer {layer_id} not found")


# ---------------------------------------------------------------------------
# RowResult / BatchResult 结构
# ---------------------------------------------------------------------------
def test_row_result_status_semantics():
    r = RowResult(excel_row=2, status=RowStatus.SUCCESS)
    assert r.success and not r.failed and not r.skipped and not r.cancelled
    assert RowStatus.FAILED.value == "FAILED"
    assert RowStatus.SKIPPED.value == "SKIPPED"
    assert RowStatus.CANCELLED.value == "CANCELLED"
    assert len(RowStatus) == 4


def test_batch_result_summary():
    br = BatchResult(total=3, success=2, failed=1, skipped=0, cancelled=False,
                     rows=[], duration_seconds=0.5)
    assert br.total == 3 and br.success == 2 and br.failed == 1


# ---------------------------------------------------------------------------
# render_one：成功路径（Duplicate→改→导出→Close）
# ---------------------------------------------------------------------------
def test_render_one_success_png(tmp_path):
    ps = FakeSession(_template())
    row = _Row()
    out = str(tmp_path / "out")
    res = _render(ps, ps.template, row, _config("PNG"), out, index=1)
    assert res.success, res.errors
    assert res.status == RowStatus.SUCCESS
    assert res.excel_row == 2
    # 导出文件存在且非空
    assert len(res.output_paths) == 1
    assert res.output_paths[0] == os.path.join(out, "001_康乐家电_张三.png")
    assert os.path.exists(res.output_paths[0])
    assert os.path.getsize(res.output_paths[0]) > 0
    # duplicate 已关闭
    assert len(ps.closed) == 1
    assert ps.closed[0].closed
    assert ps.owned == []


def test_render_one_psd_also_png(tmp_path):
    ps = FakeSession(_template())
    row = _Row()
    out = str(tmp_path / "out")
    res = _render(ps, ps.template, row, _config("PSD", also_png=True), out, index=1)
    assert res.success, res.errors
    assert sorted(os.path.basename(p) for p in res.output_paths) == ["001_康乐家电_张三.png", "001_康乐家电_张三.psd"]
    # 两个文件在同一目录
    assert os.path.dirname(res.output_paths[0]) == os.path.dirname(res.output_paths[1])


def test_render_one_grouped_dir(tmp_path):
    """复用 Stage 4.5：分组列 -> 子目录。"""
    ps = FakeSession(_template())
    rows = [_Row(excel_row=2, store="康乐家电"), _Row(excel_row=3, store="诚信电器")]
    folder_map = build_group_folder_map(rows, 0)
    out = str(tmp_path / "out")
    res = _render(ps, ps.template, rows[0], _config("PNG", group_enabled=True, group_column=0),
                  out, folder_map=folder_map, index=1)
    assert res.success, res.errors
    assert res.output_paths[0] == os.path.join(out, "康乐家电", "001_康乐家电_张三.png")


# ---------------------------------------------------------------------------
# 文字：写入失败 / read-back 不一致 -> FAILED，不导出
# ---------------------------------------------------------------------------
def test_text_write_failure_blocks_export(tmp_path):
    ps = FakeSession(_template())
    # 让 电话 层写入失败
    _by_id(ps.template, "0/1").fail_write = True
    row = _Row()
    out = str(tmp_path / "out")
    res = _render(ps, ps.template, row, _config("PNG"), out, index=1)
    assert res.failed
    assert any("电话写入失败" in e for e in res.errors)
    assert res.output_paths == []           # 不导出错误图
    assert len(ps.closed) == 1              # duplicate 仍关闭


def test_text_readback_mismatch_blocks_export(tmp_path):
    ps = FakeSession(_template())
    # 让 姓名 层写入后回读仍是旧值（模拟 PS 写入未生效）：setter 不回写 _text
    orig = _by_id(ps.template, "0/0")   # 姓名
    class _TI2:
        def __init__(self, layer):
            self._layer = layer
        @property
        def Contents(self):
            return self._layer._text
        @Contents.setter
        def Contents(self, v):
            self._layer.contents_written.append(v)
            # 故意不回写 _text -> read-back 仍 "模板姓名" != "张三"
    orig._text_item = _TI2(orig)
    row = _Row()
    out = str(tmp_path / "out")
    res = _render(ps, ps.template, row, _config("PNG"), out, index=1)
    assert res.failed
    assert any("回读不一致" in e for e in res.errors)
    assert res.output_paths == []


def test_text_skip_when_role_none(tmp_path):
    """销售顾问为空 -> 跳过该字段（不失败）。"""
    ps = FakeSession(_template())
    row = _Row(role=None)
    out = str(tmp_path / "out")
    res = _render(ps, ps.template, row, _config("PNG"), out, index=1)
    assert res.success, res.errors
    assert any("销售顾问为空" in w for w in res.warnings)


# ---------------------------------------------------------------------------
# Logo：read-back 失败 -> 阻止导出
# ---------------------------------------------------------------------------
def test_logo_readback_failure_blocks_export(tmp_path):
    ps = FakeSession(_template())
    # 让 诚信电器 Logo 层 Visible 写不生效（模拟 PS read-back 仍是旧值 True）
    # 本行用康乐家电 -> 诚信电器应为 False；但写不进 -> read-back 仍 True -> 校验失败
    for l in ps.template.layers:
        if l.id == "1/1":
            l.stubborn_visible = True      # 写 Visible 不生效（read-back 保持旧值）
    row = _Row()
    out = str(tmp_path / "out")
    res = _render(ps, ps.template, row, _config("PNG"), out, index=1)
    assert res.failed
    assert any("Logo" in e and "read-back" in e for e in res.errors)
    assert res.output_paths == []


# ---------------------------------------------------------------------------
# 导出文件验证：不存在 / 0 字节
# ---------------------------------------------------------------------------
def test_export_verification_missing_file(tmp_path):
    """SaveAs 假装成功但文件没生成 -> ExportVerificationError。"""
    class _NoWriteDoc(FakeDoc):
        def SaveAs(self, path, opt, as_copy=True):
            self.saved.append((str(path), None))   # 不写文件
    ps = FakeSession(_NoWriteDoc(layers=_template().top_layers))
    row = _Row()
    out = str(tmp_path / "out")
    res = _render(ps, ps.template, row, _config("PNG"), out, index=1)
    assert res.failed
    assert any("文件不存在" in e for e in res.errors)


def test_export_verification_zero_byte(tmp_path):
    class _ZeroDoc(FakeDoc):
        def SaveAs(self, path, opt, as_copy=True):
            path = str(path)
            os.makedirs(os.path.dirname(path), exist_ok=True)
            with open(path, "wb") as f:
                f.write(b"")          # 0 字节
            self.saved.append((path, None))
    ps = FakeSession(_ZeroDoc(layers=_template().top_layers))
    row = _Row()
    out = str(tmp_path / "out")
    res = _render(ps, ps.template, row, _config("PNG"), out, index=1)
    assert res.failed
    assert any("0 字节" in e for e in res.errors)


# ---------------------------------------------------------------------------
# 部分输出失败：PSD 成功 PNG 失败 -> FAILED 且保留已生成文件
# ---------------------------------------------------------------------------
def test_partial_export_failure_keeps_successful_files(tmp_path):
    class _PartialDoc(FakeDoc):
        def SaveAs(self, path, opt, as_copy=True):
            path = str(path)
            os.makedirs(os.path.dirname(path), exist_ok=True)
            if str(path).endswith(".png"):
                raise RuntimeError("PNG export failed")   # PNG 失败
            with open(path, "wb") as f:
                f.write(b"PS")
            self.saved.append((path, None))
    ps = FakeSession(_PartialDoc(layers=_template().top_layers))
    row = _Row()
    out = str(tmp_path / "out")
    res = _render(ps, ps.template, row, _config("PSD", also_png=True), out, index=1)
    assert res.failed
    assert any("PNG" in e and "失败" in e for e in res.errors)
    # PSD 已成功生成 -> 保留在 output_paths
    assert len(res.output_paths) == 1
    assert res.output_paths[0].endswith(".psd")
    assert os.path.exists(res.output_paths[0])
    assert len(ps.closed) == 1


# ---------------------------------------------------------------------------
# finally 关 duplicate：各种失败路径都关闭
# ---------------------------------------------------------------------------
def test_duplicate_closed_on_logo_failure(tmp_path):
    ps = FakeSession(_template())
    # 无门店映射 -> prepare 抛错
    row = _Row(store="不存在门店")
    out = str(tmp_path / "out")
    res = _render(ps, ps.template, row, _config("PNG"), out, index=1)
    assert res.failed
    assert len(ps.closed) == 1
    assert ps.owned == []


def test_duplicate_closed_on_export_failure(tmp_path):
    class _FailDoc(FakeDoc):
        def SaveAs(self, path, opt, as_copy=True):
            raise RuntimeError("SaveAs boom")
    ps = FakeSession(_FailDoc(layers=_template().top_layers))
    row = _Row()
    out = str(tmp_path / "out")
    res = _render(ps, ps.template, row, _config("PNG"), out, index=1)
    assert res.failed
    assert len(ps.closed) == 1
    assert ps.owned == []


# ---------------------------------------------------------------------------
# cancel_event：取消时不再继续
# ---------------------------------------------------------------------------
class _Cancelled:
    def __init__(self, cancelled=False):
        self._c = cancelled
    def is_set(self):
        return self._c


def test_render_one_cancel_event_noop():
    """cancel_event 基础接口：render_one 不主动检查（由 run_batch 控制），传 None 也可。"""
    ps = FakeSession(_template())
    out = os.path.join(os.getcwd(), "_ct")
    res = _render(ps, ps.template, _Row(), _config("PNG"), out, cancel=_Cancelled(True))
    # render_one 本身不因 cancel 跳过（cancellation 在 run_batch 层），单行仍成功
    assert res.success


def test_run_batch_cancel_stops(tmp_path):
    ps = FakeSession(_template())
    rows = [_Row(excel_row=2), _Row(excel_row=3), _Row(excel_row=4)]
    out = str(tmp_path / "out")

    class _CancelAfterOne:
        def __init__(self):
            self._n = 0
        def is_set(self):
            self._n += 1
            return self._n > 1   # 第一行后取消
    br = run_batch(ps_session=ps, template_doc=ps.template, rows=rows,
                   config=_config("PNG"), layer_index=None,
                   logo_mapping=_logo_mapping(), out_dir=out,
                   cancel_event=_CancelAfterOne(), com_dispatch=FakeDispatch)
    assert br.cancelled
    assert len(br.rows) == 1
    assert br.success == 1


# ---------------------------------------------------------------------------
# Preview 与 Batch 内容一致（仅 base / filename 不同）
# ---------------------------------------------------------------------------
def test_preview_and_batch_same_content_diff_base(tmp_path):
    """同一 ExcelRow：Preview 与 Batch 渲染结果必须一致（文字/Logo），仅 base 与文件名不同。"""
    ps = FakeSession(_template())
    row = _Row()
    out = str(tmp_path / "out")
    # Batch
    res_b = _render(ps, ps.template, row, _config("PNG"), out, index=1, preview=False)
    assert res_b.success
    assert res_b.output_paths[0] == os.path.join(out, "001_康乐家电_张三.png")
    # Preview：base=out/_preview，filename=preview_张三
    res_p = _render(ps, ps.template, row, _config("PNG"), os.path.join(out, "_preview"),
                    index=1, preview=True)
    assert res_p.success
    assert res_p.output_paths[0] == os.path.join(out, "_preview", "preview_张三.png")
    # 业务内容一致：本测试用同一 template+row，文字写入历史应相同
    # （Batch 的 dup 与 Preview 的 dup 都是模板的副本，写入了相同文本）
    assert res_b.store == res_p.store == "康乐家电"
    assert res_b.name == res_p.name == "张三"


def test_preview_grouped_dir(tmp_path):
    """Preview 分组：out/_preview/分组子目录。"""
    ps = FakeSession(_template())
    rows = [_Row(excel_row=2, store="康乐家电"), _Row(excel_row=3, store="诚信电器")]
    folder_map = build_group_folder_map(rows, 0)
    out = str(tmp_path / "out")
    res = _render(ps, ps.template, rows[0], _config("PNG", group_enabled=True, group_column=0),
                  os.path.join(out, "_preview"), folder_map=folder_map, index=1, preview=True)
    assert res.success, res.errors
    assert res.output_paths[0] == os.path.join(out, "_preview", "康乐家电", "preview_张三.png")


# ---------------------------------------------------------------------------
# run_batch：单行失败继续 + 汇总
# ---------------------------------------------------------------------------
def test_run_batch_continues_after_row_failure(tmp_path):
    ps = FakeSession(_template())
    # 第 2 行电话写入失败
    rows = [_Row(excel_row=2), _Row(excel_row=3), _Row(excel_row=4)]
    out = str(tmp_path / "out")

    class _FlakySession(FakeSession):
        def __init__(self, template):
            super().__init__(template)
            self._n = 0
        def duplicate_document(self, doc):
            self._n += 1
            dup = super().duplicate_document(doc)
            if self._n == 2:      # 第 2 行电话层写失败
                for l in dup.layers:
                    if l.id == "0/1":
                        l.fail_write = True
            return dup
    ps = _FlakySession(_template())
    br = run_batch(ps_session=ps, template_doc=ps.template, rows=rows,
                   config=_config("PNG"), layer_index=None,
                   logo_mapping=_logo_mapping(), out_dir=out, com_dispatch=FakeDispatch)
    assert br.total == 3
    assert br.success == 2
    assert br.failed == 1
    assert len(br.rows) == 3
    # 失败行的 errors 含电话写入失败
    failed = [r for r in br.rows if r.failed]
    assert len(failed) == 1
    assert any("电话写入失败" in e for e in failed[0].errors)
    # 所有行 duplicate 都关闭
    assert len(ps.closed) == 3
    assert ps.owned == []


# ---------------------------------------------------------------------------
# 组件级：apply_text_fields / apply_logo_visibility / export_document
# ---------------------------------------------------------------------------
def test_apply_text_fields_success(tmp_path):
    doc = _template()
    row = _Row()
    warns = apply_text_fields(doc, row, _text_map(), excel_row=2)
    assert warns == []
    assert _by_id(doc, "0/0")._text == "张三"
    assert _by_id(doc, "0/1")._text == "138"
    assert _by_id(doc, "0/2")._text == "王经理"


def test_apply_text_fields_write_error():
    doc = _template()
    _by_id(doc, "0/0").fail_write = True
    row = _Row()
    with pytest.raises(TextWriteError) as ei:
        apply_text_fields(doc, row, _text_map(), excel_row=2)
    assert "姓名写入失败" in str(ei.value)


def test_export_document_png(tmp_path):
    doc = _template()
    out = str(tmp_path / "out")
    paths = export_document(doc, out, "001_康乐家电_张三", "PNG",
                            com_dispatch=FakeDispatch, excel_row=2)
    assert paths == [os.path.join(out, "001_康乐家电_张三.png")]
    assert os.path.getsize(paths[0]) > 0


def test_export_document_psd_also_png(tmp_path):
    doc = _template()
    out = str(tmp_path / "out")
    paths = export_document(doc, out, "x", "PSD", also_png=True,
                            com_dispatch=FakeDispatch, excel_row=2)
    assert sorted(os.path.basename(p) for p in paths) == ["x.png", "x.psd"]


def test_save_option_not_cached_across_calls(tmp_path):
    """回归（Stage 5 打包后发现）：SaveOptions 必须每次新建，绝不模块级缓存。

    真实 bug：_save_option 曾把 SaveOptions 缓存到模块级 _SAVE_OPTIONS，
    Preview（主线程）创建并缓存后，Batch（worker 线程）复用同一 CDispatch
    对象传给 SaveAs → `CDispatch can not be converted to a COM VARIANT`。
    本测试用计数 Dispatch 断言每次导出都新建对象（跨调用无复用）。
    """
    from core.renderer import _save_option
    calls = []

    def counting_dispatch(progid):
        obj = FakeDispatch(progid)
        calls.append(progid)
        return obj

    o1 = _save_option("PNG", counting_dispatch)
    o2 = _save_option("PNG", counting_dispatch)
    o3 = _save_option("PSD", counting_dispatch)
    # 每次调用都新建（不缓存）-> PNG 调了 2 次、PSD 1 次
    assert calls == ["Photoshop.PNGSaveOptions", "Photoshop.PNGSaveOptions",
                     "Photoshop.PhotoshopSaveOptions"]
    # 返回对象不是同一个（跨调用无复用）
    assert o1 is not o2
    assert isinstance(o1, FakeDispatch) and isinstance(o3, FakeDispatch)


def test_export_document_missing_file_raises(tmp_path):
    class _NoWrite(FakeDoc):
        def SaveAs(self, path, opt, as_copy=True):
            pass
    with pytest.raises(ExportVerificationError) as ei:
        export_document(_NoWrite(), str(tmp_path), "x", "PNG",
                        com_dispatch=FakeDispatch, excel_row=2)
    assert "文件不存在" in str(ei.value)
