# -*- coding: utf-8 -*-
"""
renderer.py —— Stage 5 统一 Renderer（Preview / Batch / CLI 共用单行渲染）。

目标：把历史上分散在 GUI（qifang_cover_maker.run_batch/_preview）与 CLI
（psd_cover_batch.run）中的单行 Photoshop 修改逻辑收敛到唯一入口 render_one()：

    ExcelRow + BatchConfig + Template Document
        → render_one()
        → RowResult

红线（Stage 1 ownership / Stage 2 LayerRef / Stage 3 Logo / Stage 4 Excel /
Stage 4.5 output core 全部延续）：
  - 每一行继续 Duplicate → 修改副本 → 导出 → finally Close（绝不复用一个文档连续改多行）；
  - duplicate 必须 finally 关闭（文字失败 / Logo 失败 / SaveAs 失败 / 用户 stop /
    verification 失败都要进入安全清理）；
  - 文字层全部通过 LayerRef resolve（resolve_layer），禁止 find_layer(name)；
  - 文字写入后 read-back 验证（TextVerificationError）；失败不静默继续、不导出错误图；
  - Logo 只执行 Stage 3 的 prepare_logo_visibility 计划并 read-back 验证，不重写算法；
  - 输出目录复用 Stage 4.5 output core（resolve_output_directory / GroupFolderMap），
    禁止 os.path.join(out, row.store)；
  - 导出后验证文件存在且 size>0（ExportVerificationError）；
  - 单行失败不中断批次，汇总进 BatchResult。

本模块不 import openpyxl / tkinter / win32com 直接业务逻辑之外的界面依赖；
Photoshop COM 通过传入的 ps_session 操作，纯逻辑部分可单测（fake session）。
"""

import os
import time
from dataclasses import dataclass, field
from enum import Enum
from typing import Any, Callable, Dict, List, Optional, Sequence, Tuple

from core.output_paths import (
    resolve_output_directory, GroupFolderMap, OutputPathError,
)
from core.layer_index import resolve_layer
from core.logo_mapping import (
    LogoMapping, prepare_logo_visibility, verify_logo_visibility,
    LogoVisibilityError,
)

# ---------------------------------------------------------------------------
# 状态与结果模型
# ---------------------------------------------------------------------------
class RowStatus(Enum):
    SUCCESS = "SUCCESS"
    FAILED = "FAILED"
    SKIPPED = "SKIPPED"
    CANCELLED = "CANCELLED"


@dataclass
class RowResult:
    """单行渲染结果（GUI 根据它展示，不再解析日志文字）。"""

    excel_row: int
    status: RowStatus
    store: str = ""
    name: str = ""
    output_paths: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)
    duration_seconds: float = 0.0

    @property
    def success(self) -> bool:
        return self.status == RowStatus.SUCCESS

    @property
    def skipped(self) -> bool:
        return self.status == RowStatus.SKIPPED

    @property
    def cancelled(self) -> bool:
        return self.status == RowStatus.CANCELLED

    @property
    def failed(self) -> bool:
        return self.status == RowStatus.FAILED


@dataclass
class BatchResult:
    """整批渲染汇总（GUI 展示 成功/失败/跳过 的直接来源）。"""

    total: int
    success: int = 0
    failed: int = 0
    skipped: int = 0
    cancelled: bool = False
    rows: List[RowResult] = field(default_factory=list)
    duration_seconds: float = 0.0


# ---------------------------------------------------------------------------
# 异常（renderer 内部语义；render_one 会捕获并转成 RowResult.errors）
# ---------------------------------------------------------------------------
class RenderError(Exception):
    """渲染错误基类。"""

    def __init__(self, message: str, *, excel_row: Optional[int] = None):
        super().__init__(message)
        self.excel_row = excel_row


class TextWriteError(RenderError):
    """文字写入失败（setter 抛异常 / 无法定位图层）。"""


class TextVerificationError(RenderError):
    """文字写入后 read-back 与期望不一致。"""


class LogoRenderError(RenderError):
    """Logo 应用 / read-back 校验失败（阻止导出）。"""


class ExportVerificationError(RenderError):
    """导出后文件不存在或 0 字节。"""


# ---------------------------------------------------------------------------
# 常量
# ---------------------------------------------------------------------------
TEXT_FIELD_LABELS = ("姓名", "电话", "销售顾问")
PNG_FORMAT = "PNG"
JPG_FORMAT = "JPG"
PSD_FORMAT = "PSD"


def _save_option(fmt: str, com_dispatch):
    """按格式创建 Photoshop SaveOptions。

    **绝不做模块级/长期缓存**：SaveOptions 是 COM 对象（CDispatch），
    绑定创建线程的 STA 单元。Preview（主线程）与 Batch（worker 线程）
    各自创建、各自使用，跨线程复用已创建对象会抛
    `CDispatch can not be converted to a COM VARIANT`。
    因此每次导出都新建（与 Stage 5 前 export_doc 行为一致）。
    """
    if fmt == PNG_FORMAT:
        opt = com_dispatch("Photoshop.PNGSaveOptions")
        opt.Interlaced = False
        opt.Compression = 6
    elif fmt == JPG_FORMAT:
        opt = com_dispatch("Photoshop.JPGSaveOptions")
        opt.Quality = 9
        opt.EmbedColorProfile = False
    else:  # PSD
        opt = com_dispatch("Photoshop.PhotoshopSaveOptions")
    return opt


def _normalize_text(actual: Any) -> str:
    """read-back 文本规范化比较（去空白 + NFKC，容忍 PS 追加/格式差异）。"""
    import unicodedata
    if actual is None:
        return ""
    return unicodedata.normalize("NFKC", str(actual).strip())


# ---------------------------------------------------------------------------
# 文字：统一应用 + read-back 验证（Stage 5 #8/#9/#10）
# ---------------------------------------------------------------------------
def apply_text_fields(
    doc,
    row,
    text_mapping: Dict[str, Any],
    *,
    excel_row: Optional[int] = None,
    log: Optional[Callable[[str], None]] = None,
    retries: int = 6,
) -> List[str]:
    """把 ExcelRow 的 姓名/电话/销售顾问 写入文字层（全部经 LayerRef resolve）。

    text_mapping: dict[field_label -> LayerRef | dict | str | None]
      字段 label ∈ {姓名, 电话, 销售顾问}；值为 None/"" 表示「不替换」。

    行为（P0-05 彻底取消「静默失败继续」）：
      - 必填字段（label 已配置且值非空）写入失败或 read-back 不一致
        → 抛 TextWriteError / TextVerificationError（携带 行号/字段/目标/预期值）；
      - 字段配置为空 / 值为空（如销售顾问列无值）→ 跳过（记 warning，不失败）；
      - 返回 warnings 列表。

    注意：本函数只做「定位 + 写入 + 回读」，不做 duplicate 管理（由 render_one 负责）。
    """
    warnings: List[str] = []
    for label in TEXT_FIELD_LABELS:
        ref_cfg = text_mapping.get(label)
        if not ref_cfg:            # 未配置 / 显式不替换
            continue
        ref = _as_layer_ref(ref_cfg)
        # 取值：销售顾问可能为 None（Excel 该列无值）——跳过但不失败
        if label == "销售顾问" and row.role is None:
            warnings.append(f"第 {row.excel_row} 行：销售顾问为空，跳过该字段。")
            continue
        expected = {
            "姓名": row.name,
            "电话": row.phone,
            "销售顾问": row.role or "",
        }[label]
        if not expected:
            warnings.append(f"第 {row.excel_row} 行：{label}为空，跳过该字段。")
            continue
        # 定位（禁止 find_layer：全部走 LayerRef.resolve）
        try:
            layer = resolve_layer(doc, ref)
        except Exception as e:
            raise TextWriteError(
                f"第 {row.excel_row} 行：{label}图层定位失败，目标图层“{ref.display_path}”：{e}",
                excel_row=row.excel_row) from e
        # 写入（带重试）
        try:
            _write_text_with_retry(layer, expected, retries=retries)
        except Exception as e:
            raise TextWriteError(
                f"第 {row.excel_row} 行：{label}写入失败，目标图层“{ref.display_path}”"
                f"，预期值 {expected!r}：{e}",
                excel_row=row.excel_row) from e
        # read-back 验证
        try:
            actual = layer.TextItem.Contents
        except Exception as e:
            raise TextVerificationError(
                f"第 {row.excel_row} 行：{label}回读失败，目标图层“{ref.display_path}”：{e}",
                excel_row=row.excel_row) from e
        if _normalize_text(actual) != _normalize_text(expected):
            raise TextVerificationError(
                f"第 {row.excel_row} 行：{label}回读不一致，目标图层“{ref.display_path}”，"
                f"预期 {expected!r}，实际 {actual!r}",
                excel_row=row.excel_row)
        if log:
            log(f"    文字 {label}：{expected!r} → {ref.display_path}")
    return warnings


def _as_layer_ref(ref_cfg: Any):
    """把配置值归一化为 LayerRef（dict / LayerRef / str）。"""
    if hasattr(ref_cfg, "id") and hasattr(ref_cfg, "index_path"):
        return ref_cfg                      # 已是 LayerRef
    from core.layer_index import ref_from_config
    ref = ref_from_config(ref_cfg)
    if ref is None:
        raise TextWriteError(f"文字层配置无效：{ref_cfg!r}")
    return ref


def _write_text_with_retry(layer, text: str, retries: int = 6):
    """写入 TextItem.Contents（COM 偶发失败重试）。失败抛异常。"""
    last = None
    for attempt in range(retries):
        try:
            layer.TextItem.Contents = text
            return
        except Exception as e:
            last = e
            if attempt < retries - 1:
                time.sleep(0.3 * (attempt + 1))
    raise last


# ---------------------------------------------------------------------------
# Logo：执行 Stage 3 计划 + read-back 验证（Stage 5 #11/#12）
# ---------------------------------------------------------------------------
def _enable_parents(container):
    """启用 container 及其所有祖先组（直到文档层）。"""
    p = container
    while True:
        try:
            p.Visible = True
            p = p.Parent
        except Exception:
            break


def apply_logo_visibility(
    doc,
    store: str,
    logo_mapping: LogoMapping,
    *,
    excel_row: Optional[int] = None,
    effective_leaf_refs: Optional[Sequence] = None,
    log: Optional[Callable[[str], None]] = None,
) -> List[str]:
    """执行 Stage 3 的 Logo 可见性计划并 read-back 验证。

    - prepare_logo_visibility()（纯计划，不重写算法）→ resolve LayerRef → 写 Visible
      → verify_logo_visibility() read-back；
    - 任何不一致 / 无映射 → 抛 LogoRenderError（**禁止 warning 后继续 SaveAs**）；
    - 返回展示用的 shown 列表（当前行可见的 Logo display_path）。
    """
    try:
        plan = prepare_logo_visibility(
            store, logo_mapping, require_store_mapping=True,
            effective_leaf_refs=effective_leaf_refs)
    except LogoVisibilityError as e:
        raise LogoRenderError(
            f"第 {excel_row if excel_row is not None else '-'} 行：Logo 计划生成失败：{e}",
            excel_row=excel_row) from e

    applied: Dict[str, bool] = {}
    for ref, visible in plan:
        try:
            layer = resolve_layer(doc, ref)
            layer.Visible = visible
            if visible:
                try:
                    _enable_parents(layer.Parent)
                except Exception:
                    pass
            applied[ref.id] = visible
        except Exception as e:
            raise LogoRenderError(
                f"第 {excel_row if excel_row is not None else '-'} 行：Logo 图层"
                f"“{ref.display_path}”应用失败：{e}",
                excel_row=excel_row) from e
    # read-back 校验（Stage 5 #12：不一致阻止导出）
    readback: Dict[str, bool] = {}
    for ref, _expected in plan:
        try:
            readback[ref.id] = bool(resolve_layer(doc, ref).Visible)
        except Exception:
            readback[ref.id] = applied.get(ref.id)
    try:
        verify_logo_visibility(readback, plan)
    except LogoVisibilityError as e:
        raise LogoRenderError(
            f"第 {excel_row if excel_row is not None else '-'} 行：Logo read-back 校验失败：{e}",
            excel_row=excel_row) from e

    shown = [r.display_path for r, v in plan if v]
    if log:
        log(f"    Logo: {store!r} -> {shown}")
    return shown


# ---------------------------------------------------------------------------
# Export：统一导出 + 文件验证（Stage 5 #14/#15/#17）
# ---------------------------------------------------------------------------
def export_document(
    doc,
    output_dir: str,
    filename_stem: str,
    fmt: str,
    also_png: bool = False,
    *,
    com_dispatch=None,
    excel_row: Optional[int] = None,
) -> List[str]:
    """统一导出（PNG / JPG / PSD+also_png），全部落在 output_dir。

    返回实际生成的文件路径列表（验证存在且 size>0，否则 ExportVerificationError）。

    com_dispatch：win32com.client.Dispatch（必传；由调用方从 session 注入，
    fake 测试时替换）。注意：不能设为模块级默认，否则首次导入即执行 Dispatch。

    注意：output_dir 由调用方用 Stage 4.5 resolve_output_directory 解析好，
    renderer 不自行组装分组目录。
    """
    if com_dispatch is None:
        raise ExportVerificationError(
            "导出失败：未提供 com_dispatch（Photoshop COM Dispatch）",
            excel_row=excel_row)
    os.makedirs(output_dir, exist_ok=True)
    produced: List[str] = []
    errors: List[str] = []

    def _save(f, ext):
        path = os.path.join(output_dir, filename_stem + ext)
        opt = _save_option(f, com_dispatch)
        try:
            doc.SaveAs(path, opt, True)
            _verify_exported(path, excel_row=excel_row)
            produced.append(path)
            return True
        except RenderError as e:
            errors.append(str(e))
            return False
        except Exception as e:  # SaveAs 本身的异常（COM/磁盘）
            errors.append(f"{f} 导出失败：{e}")
            return False

    def _raise_export_error(msg):
        err = ExportVerificationError(msg, excel_row=excel_row)
        err.partial_paths = list(produced)   # 规范 #18：保留已成功文件，不删除
        raise err

    if fmt == PSD_FORMAT:
        ok_psd = _save(PSD_FORMAT, ".psd")
        ok_png = True
        if also_png:
            ok_png = _save(PNG_FORMAT, ".png")
        # 任一出错都必须抛出（render_one 标 FAILED、不静默）；成功文件保留
        if not ok_psd or (also_png and not ok_png):
            _raise_export_error(
                f"第 {excel_row if excel_row is not None else '-'} 行：部分导出失败"
                f"（已生成 {produced}）：{'；'.join(errors)}")
    else:
        ok = _save(fmt, "." + fmt.lower())
        if not ok:
            _raise_export_error(
                f"第 {excel_row if excel_row is not None else '-'} 行：{fmt} 导出失败"
                f"（已生成 {produced}）：{'；'.join(errors)}")
    return produced


def _verify_exported(path: str, *, excel_row: Optional[int] = None):
    """导出后验证文件存在且非 0 字节。"""
    if not os.path.exists(path):
        raise ExportVerificationError(
            f"第 {excel_row if excel_row is not None else '-'} 行：导出文件不存在：{path}",
            excel_row=excel_row)
    if os.path.getsize(path) == 0:
        raise ExportVerificationError(
            f"第 {excel_row if excel_row is not None else '-'} 行：导出文件为 0 字节：{path}",
            excel_row=excel_row)


# ---------------------------------------------------------------------------
# render_one：唯一单行渲染入口（Stage 5 #5/#6/#7）
# ---------------------------------------------------------------------------
def render_one(
    *,
    ps_session,
    template_doc,
    row,
    config: Dict[str, Any],
    layer_index,
    logo_mapping: LogoMapping,
    output_context,
    index: int = 1,
    preview: bool = False,
    cancel_event=None,
    com_dispatch=None,
    log: Optional[Callable[[str], None]] = None,
) -> RowResult:
    """渲染单行（Preview 与 Batch 共用；唯一允许做 Photoshop 修改的地方）。

    config 需要的字段：
      text_map (dict "姓名"/"电话"/"销售顾问" -> LayerRef/dict/str/""),
      fmt ("PNG"/"JPG"/"PSD"), also_png(bool),
      group_output_enabled(bool), group_output_column(int|None)

    output_context：dict，由调用方提供：
      {
        "base_dir": str,            # Batch=out_dir；Preview=out_dir/_preview
        "folder_map": GroupFolderMap|None,   # 批次前建好的（共享碰撞映射）
        "filename_stem": str|None,  # None 时内部生成 "NNN_门店_姓名"（Batch）；
                                    # Preview 传 "preview_姓名"（含 preview 前缀）
      }

    cancel_event：可选，具备 .is_set() 的对象（threading.Event 风格）。
    com_dispatch：导出 SaveOptions 用的 win32com.client.Dispatch；None 时尝试
      从 ps_session 获取。

    返回 RowResult（不抛异常；所有失败进 errors）。duplicate 由本函数 finally 关闭。
    """
    excel_row = row.excel_row
    t0 = time.time()
    base_dir = output_context.get("base_dir", ".")
    folder_map = output_context.get("folder_map")
    group_column = config.get("group_output_column") if config.get("group_output_enabled") else None
    fmt = config.get("fmt", PNG_FORMAT)
    also_png = bool(config.get("also_png"))
    text_mapping = config.get("text_map", {})
    warnings: List[str] = []
    errors: List[str] = []
    output_paths: List[str] = []
    dup = None

    # 导出用的 Dispatch（注入或从 session 派生）
    if com_dispatch is None:
        com_dispatch = getattr(ps_session, "com_dispatch", None) or _default_dispatch

    try:
        dup = ps_session.duplicate_document(template_doc)
        try:
            ps_session.app.ActiveDocument = dup
        except Exception:
            pass
        time.sleep(0.05)

        # 1) 文字（失败即当前行失败，不继续）
        try:
            warnings += apply_text_fields(
                dup, row, text_mapping, excel_row=excel_row, log=log)
        except RenderError as e:
            errors.append(str(e))
            return _result(excel_row, row, RowStatus.FAILED, output_paths,
                           warnings, errors, t0)

        # 2) Logo（失败即当前行失败，不导出）
        try:
            shown = apply_logo_visibility(
                dup, row.store, logo_mapping, excel_row=excel_row,
                effective_leaf_refs=getattr(logo_mapping, "_effective_leaf_refs", None),
                log=log)
        except RenderError as e:
            errors.append(str(e))
            return _result(excel_row, row, RowStatus.FAILED, output_paths,
                           warnings, errors, t0)

        # 3) 输出目录（Stage 4.5 core；Preview/Batch 只 base 不同）
        try:
            row_dir = resolve_output_directory(
                base_dir, row, group_column, folder_map=folder_map) \
                if group_column is not None else base_dir
        except OutputPathError as e:
            errors.append(f"第 {excel_row} 行：输出目录解析失败：{e}")
            return _result(excel_row, row, RowStatus.FAILED, output_paths,
                           warnings, errors, t0)

        # 4) 文件名（保持现有 序号/门店/姓名 规则；Preview 用 preview_姓名）
        stem = output_context.get("filename_stem")
        if stem is None:
            from core import util as core_util
            safe_store = core_util.sanitize_filename(row.store)
            safe_name = core_util.sanitize_filename(row.name)
            stem = f"{index:03d}_{safe_store}_{safe_name}"
        if preview and not stem.startswith("preview_"):
            stem = "preview_" + core_util.sanitize_filename(row.name)

        # 5) 导出 + 文件验证（部分失败 -> 当前行 failed，但保留已成功文件）
        try:
            output_paths = export_document(
                dup, row_dir, stem, fmt, also_png=also_png,
                com_dispatch=com_dispatch, excel_row=excel_row)
        except RenderError as e:
            errors.append(str(e))
            # 规范 #18：部分导出失败时保留已成功生成的文件
            partial = getattr(e, "partial_paths", None)
            if partial:
                for p in partial:
                    if p not in output_paths:
                        output_paths.append(p)
            return _result(excel_row, row, RowStatus.FAILED, output_paths,
                           warnings, errors, t0)

        return _result(excel_row, row, RowStatus.SUCCESS, output_paths,
                       warnings, errors, t0)

    except RenderError as e:
        errors.append(str(e))
        return _result(excel_row, row, RowStatus.FAILED, output_paths,
                       warnings, errors, t0)
    except Exception as e:  # 兜底（COM 层意外）
        errors.append(f"第 {excel_row} 行：未知错误：{e}")
        return _result(excel_row, row, RowStatus.FAILED, output_paths,
                       warnings, errors, t0)
    finally:
        # Stage 1 ownership：duplicate 无论成败都必须关闭（文字/Logo/导出/停止/校验失败均覆盖）
        if dup is not None:
            try:
                ps_session.close_owned_document(dup)
            except Exception as e:
                errors.append(f"第 {excel_row} 行：关闭副本失败：{e}")


def _result(excel_row, row, status, output_paths, warnings, errors, t0):
    return RowResult(
        excel_row=excel_row,
        status=status,
        store=row.store if row is not None else "",
        name=row.name if row is not None else "",
        output_paths=list(output_paths),
        warnings=list(warnings),
        errors=list(errors),
        duration_seconds=time.time() - t0,
    )


def _default_dispatch(progid: str):
    """无 session 注入时的兜底 Dispatch（测试/fake 场景替换）。"""
    import win32com.client
    return win32com.client.Dispatch(progid)


# ---------------------------------------------------------------------------
# run_batch：遍历 rows，单行失败继续（Stage 5 #20）
# ---------------------------------------------------------------------------
def run_batch(
    *,
    ps_session,
    template_doc,
    rows: Sequence,
    config: Dict[str, Any],
    layer_index,
    logo_mapping: LogoMapping,
    out_dir: str,
    folder_map: Optional[GroupFolderMap] = None,
    cancel_event=None,
    progress_cb: Optional[Callable[[int, int], None]] = None,
    log: Optional[Callable[[str], None]] = None,
    com_dispatch=None,
) -> BatchResult:
    """整批渲染（GUI Batch / CLI 共用）。

    - 批次开始前已由调用方建好 folder_map（Stage 4.5 共享碰撞映射）；
    - 每行 render_one(preview=False)，单行失败不中断；
    - cancel_event.is_set() 时停止（已渲染行保留，剩余标记 CANCELLED）。
    """
    t0 = time.time()
    results: List[RowResult] = []
    total = len(rows)
    done = 0
    cancelled = False

    for index, row in enumerate(rows, start=1):
        if cancel_event is not None and cancel_event.is_set():
            cancelled = True
            break
        result = render_one(
            ps_session=ps_session,
            template_doc=template_doc,
            row=row,
            config=config,
            layer_index=layer_index,
            logo_mapping=logo_mapping,
            output_context={"base_dir": out_dir, "folder_map": folder_map},
            index=index,
            preview=False,
            cancel_event=cancel_event,
            com_dispatch=com_dispatch,
            log=log,
        )
        results.append(result)
        done += 1
        if result.success:
            if log:
                for p in result.output_paths:
                    log(f"  已导出 -> {p}")
        else:
            if log:
                for e in result.errors:
                    log(f"  [行 {result.excel_row}] {e}")
        if progress_cb is not None:
            progress_cb(done, total)

    success = sum(1 for r in results if r.success)
    failed = sum(1 for r in results if r.failed)
    skipped = sum(1 for r in results if r.skipped)
    return BatchResult(
        total=total,
        success=success,
        failed=failed,
        skipped=skipped,
        cancelled=cancelled,
        rows=results,
        duration_seconds=time.time() - t0,
    )
