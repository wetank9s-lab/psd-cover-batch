#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
psd_cover_batch.py

用 Excel 表格里的数据，批量替换 PSD 模板中的文字图层（姓名 / 电话 / 销售顾问），
并按「门店」自动切换对应的门店 Logo 图层，最后导出成一张张 PNG 封面图。

适用于：抖音 / 视频号等渠道的「活动海报 / 视频封面」批量制作。
典型场景：一张 PSD 模板 + N 个门店销售顾问，每人一张带自己姓名、电话、门店 Logo 的封面。

依赖：
  - Windows + 已安装 Adobe Photoshop（通过 COM 自动化驱动，无需手动打开界面操作）
  - Python 3.8+ ，安装依赖： pip install -r requirements.txt

用法示例：
  # 1) 先看 PSD 里有哪些图层（确认图层名，方便修改下面的映射）
  python psd_cover_batch.py --psd 8081.psd --xlsx 门店销售0803销售顾问.xlsx --inspect

  # 2) 先合成第 1 行数据做测试
  python psd_cover_batch.py --psd 8081.psd --xlsx 门店销售0803销售顾问.xlsx --row 1 --out ./out

  # 3) 批量：每一行数据导出一张 PNG
  python psd_cover_batch.py --psd 8081.psd --xlsx 门店销售0803销售顾问.xlsx --out ./out
"""

import os
import sys
import time
import argparse
import pythoncom
import pywintypes
import openpyxl

# Stage 0：引入 core 纯函数（行为等价），供测试复用
from core import util as core_util

# ============================================================
# PSD 图层名映射（按你的模板修改这里）
# 变更图层名后，用 --inspect 确认即可。
# ============================================================
NAME_LAYER = "姓 名 拷贝"      # 姓名文字图层（注意原模板里名字可能带空格 / 拷贝后缀）
PHONE_LAYER = "电话 拷贝"       # 电话文字图层
TITLE_LAYER = "销售顾问 拷贝"   # 职位 / 头衔文字图层（如「销售顾问」）

# Excel 列索引（0 开始）：A=门店, B=姓名, C=职位, D=电话
STORE_COL, NAME_COL, TITLE_COL, PHONE_COL = 0, 1, 2, 3

# 输出的门店 Logo 图层判定：图层名与 Excel 门店名「模糊包含」匹配（互含即可，
# 例如 Excel「康乐」可匹配 PSD「康乐电器」）。品牌 Logo（名字含 logo）始终显示。


# ---------- COM 调用重试（解决 Photoshop 「应用程序正忙」）----------
def com_call(func, *a, retries=15, **k):
    last = None
    for _ in range(retries):
        try:
            return func(*a, **k)
        except pywintypes.com_error as e:
            last = e
            pythoncom.PumpWaitingMessages()
            time.sleep(0.4)
    raise last


def collect_layers(layers, registry, depth=0, parent=""):
    """递归收集所有图层 name -> 对象，并记录是否为组。"""
    for layer in layers:
        registry[layer.Name] = layer
        try:
            if layer.Layers.Count > 0:
                registry["__group__:" + layer.Name] = True
                collect_layers(layer.Layers, registry, depth + 1, layer.Name)
        except Exception:
            pass


def set_text(registry, layer_name, text):
    if layer_name in registry:
        try:
            registry[layer_name].TextItem.Contents = text
            return True
        except Exception as e:
            print(f"  [警告] 设置文字层 {layer_name!r} 失败: {e}")
            return False
    print(f"  [警告] 未找到文字层 {layer_name!r}，已跳过")
    return False


def sanitize(name):
    """清洗为 Windows 可用文件名片段（由 core.util 提供统一实现）。"""
    return core_util.sanitize_filename(name)


def inspect(psd_path, xlsx_path):
    pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
    ps = com_call(win32com.client.Dispatch, "Photoshop.Application")
    try:
        ps.DisplayDialogs = 2
    except Exception:
        pass
    doc = com_call(ps.Open, psd_path)
    print(f"文档尺寸: {int(doc.Width)} x {int(doc.Height)}")
    print("\n图层树:")
    registry = {}

    def walk(layers, depth=0):
        for layer in layers:
            is_text = False
            try:
                _ = layer.TextItem
                is_text = True
            except Exception:
                pass
            is_group = False
            try:
                if layer.Layers.Count > 0:
                    is_group = True
            except Exception:
                pass
            tag = "GRP" if is_group else ("TXT" if is_text else "LAY")
            print("  " * depth + f"[{tag}] {layer.Name!r}")
            if is_group:
                walk(layer.Layers, depth + 1)

    walk(doc.Layers)

    # 库存门店名
    stores = set()
    if xlsx_path and os.path.exists(xlsx_path):
        wb = openpyxl.load_workbook(xlsx_path, data_only=True)
        for r in list(wb.active.iter_rows(values_only=True))[1:]:
            if r and r[STORE_COL] is not None:
                stores.add(str(r[STORE_COL]).strip())
    brand_logos = [l.Name for l in doc.Layers
                   if "logo" in l.Name.lower()
                   and l.Name.strip() not in stores
                   and "__group__:" + l.Name not in registry]
    brand_set = set(brand_logos)
    store_map = {}
    for l in doc.Layers:
        if "__group__:" + l.Name in registry or l.Name in brand_set:
            continue
        for s in stores:
            if core_util.fuzzy_contains(l.Name, s):
                store_map.setdefault(s, []).append(l.Name)
                break
    print("\n匹配到的门店 Logo 图层（与 Excel 门店名「模糊包含」匹配）:")
    if store_map:
        for s in sorted(store_map):
            print(f"  {s!r:12s} -> {store_map[s]}")
    else:
        print("  (无，请检查门店名是否与图层名存在包含关系)")
    print("\n品牌 Logo 图层（名字含 logo，每张封面都显示）:")
    print("  " + (", ".join(brand_logos) if brand_logos else "(无)"))

    com_call(doc.Close, 2)
    pythoncom.CoUninitialize()


def run(psd_path, xlsx_path, out_dir, row=None, brand_logos=None):
    os.makedirs(out_dir, exist_ok=True)
    pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
    ps = com_call(win32com.client.Dispatch, "Photoshop.Application")
    try:
        ps.DisplayDialogs = 2
    except Exception:
        pass
    doc = com_call(ps.Open, psd_path)
    ps.ActiveDocument = doc

    registry = {}
    collect_layers(doc.Layers, registry)

    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    rows = list(wb.active.iter_rows(values_only=True))
    header, data = rows[0], rows[1:]
    stores = set(str(r[STORE_COL]).strip() for r in data if r and r[STORE_COL])
    # 品牌 Logo：名字含 "logo"（且不正好等于门店名）的图层 —— 每张封面都显示，
    # 例如「七方logo」「七方logo 拷贝」。可用 --brand-logo 显式指定覆盖自动识别。
    auto_brand = [l.Name for l in doc.Layers
                  if "logo" in l.Name.lower()
                  and l.Name.strip() not in stores
                  and "__group__:" + l.Name not in registry]
    if brand_logos:
        brand_logos = [nm for nm in brand_logos if nm in registry and "__group__:" + nm not in registry]
    else:
        brand_logos = auto_brand
    brand_set = set(brand_logos)

    # 门店 Logo：与 Excel 门店名「模糊包含」匹配（互含即可，不要求完全一致）。
    # 例如 Excel 的「康乐」可匹配 PSD 的「康乐电器」，「九兴」可匹配「九兴电器」。
    store_map = {}      # Excel 门店名 -> 匹配到的 PSD 图层名列表
    store_logos = []    # 去重后的所有门店 Logo 图层名
    for l in doc.Layers:
        if "__group__:" + l.Name in registry:
            continue
        if l.Name in brand_set:
            continue
        for s in stores:
            if core_util.fuzzy_contains(l.Name, s):
                store_map.setdefault(s, []).append(l.Name)
                if l.Name not in store_logos:
                    store_logos.append(l.Name)
                break
    print(f"检测到门店 Logo 图层 {len(store_logos)} 个: {store_logos}")
    for s in sorted(store_map):
        print(f"    门店 {s!r:12s} -> {store_map[s]}")
    print(f"品牌 Logo 图层（始终显示）{len(brand_logos)} 个: {brand_logos}")

    todo = [row - 1] if row else list(range(len(data)))
    exported = 0
    for idx in todo:
        if idx < 0 or idx >= len(data):
            print(f"[跳过] 行号越界: {idx + 1}")
            continue
        r = data[idx]
        if not r or r[NAME_COL] is None:
            print(f"[跳过] 第 {idx + 1} 行数据不完整")
            continue
        store = str(r[STORE_COL]).strip() if r[STORE_COL] is not None else ""
        name = str(r[NAME_COL]).strip()
        title = str(r[TITLE_COL]).strip() if r[TITLE_COL] is not None else ""
        phone = r[PHONE_COL]
        if isinstance(phone, float):
            phone = int(phone)
        phone = str(phone).strip()

        print(f"\n[{idx + 1}/{len(data)}] 门店={store!r} 姓名={name!r} 电话={phone!r}")
        set_text(registry, NAME_LAYER, name)
        set_text(registry, PHONE_LAYER, phone)
        set_text(registry, TITLE_LAYER, title)

        target_layers = store_map.get(store, [])
        for nm in store_logos:
            registry[nm].Visible = (nm in target_layers)
        for nm in brand_logos:
            registry[nm].Visible = True
        shown = target_layers if target_layers else "（无匹配，门店 Logo 全部隐藏）"
        print(f"  门店 Logo: 显示 {store!r} -> {shown}；品牌 Logo: 全部显示")

        fname = f"{idx + 1:03d}_{sanitize(store)}_{sanitize(name)}.png"
        out_path = os.path.join(out_dir, fname)
        opts = win32com.client.Dispatch("Photoshop.PNGSaveOptions")
        opts.Interlaced = False
        com_call(doc.SaveAs, out_path, opts, True)  # asCopy=True，不改动原 PSD
        print(f"  已导出 -> {out_path}")
        exported += 1

    com_call(doc.Close, 2)
    pythoncom.CoUninitialize()
    print(f"\n完成：共导出 {exported} 张 PNG 到 {out_dir}")


def main():
    ap = argparse.ArgumentParser(description="PSD + Excel 批量生成封面 PNG")
    ap.add_argument("--psd", required=True, help="PSD 模板路径")
    ap.add_argument("--xlsx", required=True, help="Excel 数据路径")
    ap.add_argument("--out", default="./out", help="PNG 输出目录")
    ap.add_argument("--row", type=int, default=0, help="只合成指定行（1 开始），0 表示全部")
    ap.add_argument("--brand-logo", action="append", default=[],
                    help="品牌 Logo 图层名（每张封面都显示，如 七方logo）。可重复指定；不指定则自动识别名字含 logo 的图层")
    ap.add_argument("--inspect", action="store_true", help="只打印图层树后退出")
    args = ap.parse_args()

    if args.inspect:
        inspect(args.psd, args.xlsx)
        return
    run(args.psd, args.xlsx, args.out, row=args.row or None, brand_logos=args.brand_logo or None)


if __name__ == "__main__":
    import win32com.client  # 延迟导入，便于在非 Windows 上给出友好报错
    main()
