#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
psd_cover_batch.py —— CLI 薄调用层（Stage 7 重构）。

职责边界（Stage 7 目标结构）：
    CLI
     ├── argparse（参数解析）
     ├── 路径参数 / 参数转换
     ├── 调用 core（唯一业务实现）
     ├── stdout / stderr 输出
     └── exit code

本文件**不包含**任何业务算法：
  - Excel 解析       -> core.excel_data.load_excel_dataset
  - 图层遍历         -> core.layer_index.collect_layer_index
  - LayerRef resolve -> core.layer_index.resolve_layer / rebind_layer_reference
  - Logo 匹配        -> core.logo_mapping.match_store_logo（评分制，唯一事实来源）
  - Logo 可见性      -> core.logo_mapping.prepare_logo_visibility
  - 分组目录         -> core.output_paths.build_group_folder_map / resolve_output_directory
  - 单行渲染 / 导出  -> core.renderer.run_batch / render_one / export_document
  - 文件名清洗       -> core.util.sanitize_filename

禁止事项（红线延续）：
  - 禁止 import openpyxl / 直接 iter_rows / 访问 Workbook
  - 禁止 registry[layer.Name] = layer 式覆盖（同名图层绝不互相覆盖）
  - 禁止 runtime 按名字猜 Logo（"logo" in name / fuzzy）
  - 禁止 os.path.join(out, store) 自行组装分组目录
  - 禁止 Document.Duplicate / TextItem.Contents / Visible = / SaveAs 出现在本文件

exit code 规则（Stage 7 正式建立）：
  0 = 全部成功
  1 = 参数 / 配置错误（含 Excel 加载失败 / Logo 配置无效 / 分组配置错误）
  2 = Batch 完成但存在失败行（部分失败）
  3 = 内部异常（未预期错误，含 Photoshop COM 层故障）

用法示例：
  # 1) 查看 PSD 图层（无需 Excel）
  python psd_cover_batch.py --psd 8081.psd --inspect

  # 2) 查看 PSD 图层 + Excel 门店 Logo 建议
  python psd_cover_batch.py --psd 8081.psd --xlsx 门店销售0803销售顾问.xlsx --inspect

  # 3) 只合成第 1 行数据做测试
  python psd_cover_batch.py --psd 8081.psd --xlsx 门店销售0803销售顾问.xlsx --row 1 --out ./out

  # 4) 批量：每一行数据导出一张 PNG
  python psd_cover_batch.py --psd 8081.psd --xlsx 门店销售0803销售顾问.xlsx --out ./out

  # 5) 按 A 列分组输出 + 显式品牌 Logo
  python psd_cover_batch.py --psd 8081.psd --xlsx 门店销售0803销售顾问.xlsx \
      --out ./out --group-output-column A --brand-logo 七方logo
"""

import os
import sys
import argparse
import threading

# ---- core 唯一业务实现（CLI 不再自建第二套）----
from core.photoshop import PhotoshopSession
from core.layer_index import collect_layer_index, rebind_layer_reference
from core.logo_mapping import (
    LogoMapping, LogoValidationError,
    match_store_logo, validate_logo_mapping,
    resolve_effective_logo_layers, suggest_brand_logos,
    EXACT, AUTO, AMBIGUOUS,
)
from core.excel_data import (
    ExcelDataError, load_excel_dataset, excel_column_to_index,
    index_to_excel_column,
)
from core.output_paths import (
    OutputPathError, assert_group_column_valid, build_group_folder_map,
)
from core.renderer import run_batch as renderer_run_batch

# ============================================================
# 文字图层常量（按你的模板修改这里；变更后可用 --inspect 确认）
# ============================================================
NAME_LAYER = "姓 名 拷贝"      # 姓名文字图层
PHONE_LAYER = "电话 拷贝"       # 电话文字图层
TITLE_LAYER = "销售顾问 拷贝"   # 职位 / 头衔文字图层

# Excel 列索引（0 开始）：A=门店, B=姓名, C=职位, D=电话
STORE_COL, NAME_COL, TITLE_COL, PHONE_COL = 0, 1, 2, 3

# ---- exit code（Stage 7 正式规则）----
EXIT_OK = 0            # 全部成功
EXIT_CONFIG = 1        # 参数 / 配置错误
EXIT_PARTIAL = 2       # Batch 完成但存在失败行
EXIT_INTERNAL = 3      # 内部异常


# ============================================================
# inspect：查看 PSD 图层 + 可选 Excel Logo 建议（薄层，走 collect_layer_index）
# ============================================================
def _layer_tree(index):
    """把 LayerIndex 渲染成缩进树（[GROUP]/[T] 标签 + display_path；同名附加 id）。"""
    # 按 index_path 排序保证深度优先稳定序；前缀 = 父 id + "/"
    by_path = sorted(index.layers, key=lambda r: [int(x) for x in r.index_path])
    # 同名（name）图层计数：同名时必须附加 id 区分（规格六：同名必须都显示）
    from collections import Counter
    name_cnt = Counter(r.name for r in index.layers)
    lines = []
    for r in by_path:
        if r.is_group:
            tag = "GROUP"
        elif r.is_text:
            tag = "T"
        else:
            tag = "LAYER"
        indent = "  " * len(r.index_path)
        suffix = f"  [id={r.id}]" if name_cnt[r.name] > 1 else ""
        lines.append(f"{indent}[{tag}] {r.display_path}{suffix}")
    return lines


def inspect(psd_path, xlsx_path):
    """PSD 图层树 +（可选）Excel 门店 Logo 建议。Excel 缺失时 PSD 检查仍工作。"""
    with PhotoshopSession() as ps:
        ps.app.DisplayDialogs = 2
        doc = ps.open_document(psd_path)
        print(f"文档尺寸: {int(doc.Width)} x {int(doc.Height)}")
        print("\n图层树:")
        index = collect_layer_index(doc)
        for line in _layer_tree(index):
            print(line)

        # Excel 可选：仅在提供时用于 stores + Logo 建议
        if not xlsx_path:
            print("\n（未提供 --xlsx，跳过门店 Logo 建议；仅显示图层树）")
            return
        if not os.path.exists(xlsx_path):
            print(f"\n[Excel 错误] 文件不存在：{xlsx_path}（仅显示图层树）", file=sys.stderr)
            return
        try:
            ds = load_excel_dataset(
                xlsx_path, has_header=True,
                col_store=STORE_COL, col_name=NAME_COL, col_phone=PHONE_COL,
                col_role=TITLE_COL,
            )
        except ExcelDataError as e:
            print(f"\n[Excel 错误] {e}（仅显示图层树）", file=sys.stderr)
            return
        stores = ds.stores

        # Stage 3：inspect 也走 match_store_logo / suggest_brand_logos（首推逻辑）
        leaf_refs = [r for r in index.layers if not r.is_group]
        print(f"\n匹配到的门店 Logo 图层（match_store_logo 评分匹配，共 {len(stores)} 个门店）:")
        ambiguous = []
        store_map = {}
        for s in sorted(stores):
            mr = match_store_logo(s, leaf_refs)
            if mr.status in (EXACT, AUTO) and mr.best is not None:
                store_map[s] = mr.best
                print(f"  {s!r:12s} -> {mr.best.display_path}")
            elif mr.status == AMBIGUOUS:
                cand = " / ".join(r.display_path for r in mr.hits)
                print(f"  {s!r:12s} -> （歧义：{cand}）")
                ambiguous.append((s, mr))
            else:
                print(f"  {s!r:12s} -> （无匹配）")
        if ambiguous:
            print("\n[提示] 以下门店存在多个 Logo 候选，请明确配置（不自动取第一个）:", file=sys.stderr)
            for s, mr in ambiguous:
                cand = "、".join(r.display_path for r in mr.hits)
                print(f"  门店“{s}”候选：{cand}", file=sys.stderr)
        # 品牌建议：排除已被门店映射的图层（传 store_map 供 suggest 排除）
        brand_logos = [r.display_path for r in suggest_brand_logos(leaf_refs, store_map)]
        print("\n品牌 Logo 图层建议（名字含 logo 且未被门店使用，仅建议）:")
        print("  " + (", ".join(brand_logos) if brand_logos else "(无)"))


# ============================================================
# 参数转换（薄层；业务判定一律走 core）
# ============================================================
def parse_group_output_column(raw):
    """--group-output-column：列字母（A/D/AA）或 0-based 数字 -> int | None。"""
    if raw is None:
        return None
    s = str(raw).strip()
    if not s:
        return None
    if s.isdigit():
        return int(s)
    try:
        return excel_column_to_index(s.upper())
    except ValueError as e:
        raise argparse.ArgumentTypeError(f"--group-output-column 无效：{e}")


def _resolve_brand_refs(index, names):
    """把 --brand-logo 的每个 name resolve 成唯一 LayerRef（多命中/缺失明确报错）。"""
    refs = []
    for nm in names or []:
        matches = index.find_matching(nm)
        if len(matches) == 1:
            refs.append(matches[0])
        elif len(matches) > 1:
            cand = "、".join(r.display_path for r in matches)
            raise LogoValidationError(
                f"品牌 Logo“{nm}”存在多个同名图层，请明确配置（不自动取第一个）：{cand}",
                code="BRAND_AMBIGUOUS")
        else:
            raise LogoValidationError(
                f"品牌 Logo“{nm}”不存在（MISSING），请检查 --brand-logo 名称。",
                code="BRAND_MISSING")
    return refs


# ============================================================
# run：Batch 薄调用层（唯一业务实现来自 core.renderer）
# ============================================================
def run(psd_path, xlsx_path, out_dir, row=None, brand_logos=None,
        group_output_column=None):
    """返回 exit code（0/1/2/3）。"""
    os.makedirs(out_dir, exist_ok=True)
    with PhotoshopSession() as ps:
        ps.app.DisplayDialogs = 2
        doc = ps.open_document(psd_path)
        ps.app.ActiveDocument = doc

        # ---- 1) 图层索引（唯一身份；不丢同名）----
        index = collect_layer_index(doc)

        # ---- 2) Excel 数据（统一入口，CLI 不碰 openpyxl）----
        try:
            ds = load_excel_dataset(
                xlsx_path, has_header=True,
                col_store=STORE_COL, col_name=NAME_COL, col_phone=PHONE_COL,
                col_role=TITLE_COL,
            )
        except ExcelDataError as e:
            print(f"[Excel 错误] {e}", file=sys.stderr)
            return EXIT_CONFIG
        data = ds.valid_rows
        stores = ds.stores
        print(f"Excel 读取完成：{len(data)} 行有效数据，跳过 {len(ds.skipped_rows)} 行"
              f"（sheet：{ds.sheet_name}）。")

        # ---- 3) 分组目录（Preflight + 完整 folder map，与 GUI 同一核心）----
        folder_map = None
        if group_output_column is not None:
            try:
                assert_group_column_valid(ds.max_columns, True, group_output_column)
                folder_map = build_group_folder_map(data, group_output_column)
            except OutputPathError as e:
                print(f"[分组配置错误] {e}", file=sys.stderr)
                return EXIT_CONFIG
            print(f"分组字段：{index_to_excel_column(group_output_column)}，"
                  f"预计创建 {folder_map.distinct_folder_count} 个目录"
                  + (f"，空值 {folder_map.empty_group_count} 条 → {folder_map.fallback}"
                     if folder_map.empty_group_count else "")
                  + (f"，名称冲突 {folder_map.collision_count} 组，已自动区分"
                     if folder_map.collision_count else ""))

        # ---- 4) Logo 配置（启动前 recommendation；运行时只消费确定性映射）----
        leaf_refs = [r for r in index.layers if not r.is_group]
        store_logo_map = {}
        ambiguous_stores = []
        for s in stores:
            mr = match_store_logo(s, leaf_refs)
            if mr.status in (EXACT, AUTO) and mr.best is not None:
                store_logo_map[s] = mr.best
            elif mr.status == AMBIGUOUS:
                store_logo_map[s] = None
                ambiguous_stores.append((s, mr))
            else:
                store_logo_map[s] = None
        if ambiguous_stores:
            print("\n[Logo 配置错误] 以下门店存在多个 Logo 候选，请明确配置"
                  "（不自动取第一个）：", file=sys.stderr)
            for s, mr in ambiguous_stores:
                cand = "、".join(r.display_path for r in mr.hits)
                print(f"  门店“{s}”候选：{cand}", file=sys.stderr)
            return EXIT_CONFIG

        # 品牌 Logo：显式 --brand-logo（resolve 成 LayerRef）；否则 suggest_brand_logos（仅建议）
        if brand_logos:
            try:
                brand_refs = _resolve_brand_refs(index, brand_logos)
            except LogoValidationError as e:
                print(f"[Logo 配置错误] {e}", file=sys.stderr)
                return EXIT_CONFIG
        else:
            brand_refs = suggest_brand_logos(leaf_refs, store_logo_map)

        logo_map = LogoMapping(
            store_logo_map=store_logo_map,
            brand_logo_refs=brand_refs,
            logo_selection_refs=[r for r in leaf_refs],
        )

        # 运行前校验（含 store/brand 冲突；失败 -> exit 1）
        try:
            eff = resolve_effective_logo_layers(index, logo_map.logo_selection_refs)
            validate_logo_mapping(logo_map, effective_leaf_refs=eff,
                                  allow_duplicate_store_targets=True)
        except LogoValidationError as e:
            print(f"[Logo 配置错误] {e}", file=sys.stderr)
            return EXIT_CONFIG
        mapped = sum(1 for v in store_logo_map.values() if v)
        unmapped = sum(1 for v in store_logo_map.values() if v is None)
        print(f"门店 Logo 映射（{mapped} 个已匹配，{unmapped} 个未映射）:")
        for s in stores:
            v = store_logo_map.get(s)
            print(f"    门店 {s!r:12s} -> {v.display_path if v else '（未映射）'}")
        print(f"品牌 Logo 图层（始终显示）{len(brand_refs)} 个: "
              f"{[r.display_path for r in brand_refs]}")

        # ---- 5) 文字层：Legacy name 常量 -> LayerRef（rebind 唯一命中；歧义/缺失不替换）----
        text_refs = {}
        for key, legacy in (("姓名", NAME_LAYER), ("电话", PHONE_LAYER),
                            ("销售顾问", TITLE_LAYER)):
            status, ref = rebind_layer_reference(index, legacy)
            text_refs[key] = ref
            if status in ("AMBIGUOUS", "MISSING") or ref is None:
                print(f"  [警告] 文字层 {legacy!r} 无法唯一解析（{status}），该字段将不替换",
                      file=sys.stderr)

        # ---- 6) 选中行 + 渲染（唯一业务入口 renderer_run_batch）----
        todo = [row - 1] if row else list(range(len(data)))
        selected = [data[i] for i in todo if 0 <= i < len(data)]
        for i in todo:
            if i < 0 or i >= len(data):
                print(f"[跳过] 行号越界: {i + 1}", file=sys.stderr)

        br = renderer_run_batch(
            ps_session=ps,
            template_doc=doc,
            rows=selected,
            config={
                "fmt": "PNG",
                "also_png": False,
                "text_map": text_refs,
                "group_output_enabled": group_output_column is not None,
                "group_output_column": group_output_column,
            },
            layer_index=index,
            logo_mapping=logo_map,
            out_dir=out_dir,
            folder_map=folder_map,
            cancel_event=threading.Event(),   # CLI 无 UI 停止按钮；预留接口
            log=print,
            com_dispatch=_default_dispatch,
        )

        # ---- 7) BatchResult 汇总（信息来自 BatchResult/RowResult，不解析 log）----
        _print_summary(br, out_dir)
        return batch_exit_code(br)


def batch_exit_code(br):
    """根据 BatchResult 决定 exit code（纯函数，供测试锁定）。

    规则（Stage 7 规格十三）：
      0 = 全部成功；2 = Batch 完成但有失败行 / 被取消（部分失败）。
    配置类错误（1）与内部异常（3）在 run/main 层判定。
    """
    if br.failed > 0 or br.cancelled:
        return EXIT_PARTIAL
    return EXIT_OK


def _print_summary(br, out_dir):
    """把 BatchResult 渲染成 stdout 汇总（失败明细来自 RowResult.errors）。"""
    print(f"\n批量完成")
    print(f"总数：{br.total}")
    print(f"成功：{br.success}")
    print(f"失败：{br.failed}")
    print(f"跳过：{br.skipped}")
    print(f"耗时：{br.duration_seconds:.1f}s")
    print(f"输出：{os.path.abspath(out_dir)}")
    if br.cancelled:
        print("用户已停止")
    failed_rows = [r for r in br.rows if r.failed]
    if failed_rows:
        print("\n失败行：")
        for r in failed_rows:
            for e in r.errors:
                print(f"  - Excel {r.excel_row}：{r.store or r.name} — {e}")


def _default_dispatch(progid: str):
    """Photoshop COM Dispatch（薄层唯一 COM 接触点；业务调用方注入使用）。"""
    import win32com.client
    return win32com.client.Dispatch(progid)


# ============================================================
# main：argparse + 参数转换 + 分发（纯薄层）
# ============================================================
def build_parser():
    ap = argparse.ArgumentParser(description="PSD + Excel 批量生成封面 PNG（CLI 薄调用层）")
    ap.add_argument("--psd", required=True, help="PSD 模板路径")
    ap.add_argument("--xlsx", default=None,
                    help="Excel 数据路径（--inspect 时可选：仅用于门店 Logo 建议；"
                         "run 时必填）")
    ap.add_argument("--out", default="./out", help="PNG 输出目录")
    ap.add_argument("--row", type=int, default=0, help="只合成指定行（1 开始），0 表示全部")
    ap.add_argument("--group-output-column", default=None,
                    help="按 Excel 该列的值创建输出子文件夹（支持列字母如 AA 或 0-based 数字如 26）；"
                         "不指定则不分组")
    ap.add_argument("--brand-logo", action="append", default=[],
                    help="品牌 Logo 图层名（每张封面都显示）。可重复指定；不指定则自动识别名字含 logo 的图层")
    ap.add_argument("--inspect", action="store_true", help="只打印图层树后退出")
    return ap


def main(argv=None):
    args = build_parser().parse_args(argv)

    # --xlsx 缺失时：inspect 仍可工作；run 需要 Excel
    if not args.xlsx and not args.inspect:
        print("[参数错误] run 模式必须提供 --xlsx（inspect 模式可省略，仅查 PSD 图层树）",
              file=sys.stderr)
        return EXIT_CONFIG

    # PSD 文件存在性预检（参数错误 -> exit 1，而不是 Photoshop COM 报内部错误）
    if not os.path.exists(args.psd):
        print(f"[参数错误] PSD 文件不存在：{args.psd}", file=sys.stderr)
        return EXIT_CONFIG

    # 路径绝对化（薄层职责：Photoshop COM / SaveAs 需要 Windows 绝对路径）
    psd_path = os.path.abspath(args.psd)
    xlsx_path = os.path.abspath(args.xlsx) if args.xlsx else None
    out_dir = os.path.abspath(args.out)

    if args.inspect:
        try:
            inspect(psd_path, xlsx_path)
        except Exception as e:  # PSD 打开失败 / COM 层故障等
            print(f"[内部错误] {e}", file=sys.stderr)
            return EXIT_INTERNAL
        return EXIT_OK

    # --group-output-column -> int | None
    try:
        group_col = parse_group_output_column(args.group_output_column)
    except argparse.ArgumentTypeError as e:
        print(f"[参数错误] {e}", file=sys.stderr)
        return EXIT_CONFIG

    try:
        return run(psd_path, xlsx_path, out_dir, row=args.row or None,
                   brand_logos=args.brand_logo or None,
                   group_output_column=group_col)
    except Exception as e:  # 未预期内部异常（含 Photoshop COM 层故障）
        print(f"[内部错误] {e}", file=sys.stderr)
        return EXIT_INTERNAL


if __name__ == "__main__":
    # 延迟导入，便于在非 Windows 上给出友好报错
    import win32com.client  # noqa: F401
    sys.exit(main())
