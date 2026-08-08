# -*- coding: utf-8 -*-
"""Stage 4 BLOCKED B/C 真实验证：GUI 同会话 col_store B->E + has_header 切换。

用真实 tk App 实例（不 mainloop），注入 fake layer_index / layer_labels，
验证 _start 前 _ensure_dataset_fresh 自动重解析并重建门店映射区。
"""
import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import openpyxl
import tkinter as tk

import qifang_cover_maker as g
from core.excel_data import load_excel_dataset


def make_xlsx(p, rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    for row in rows:
        ws.append(row)
    wb.save(p)


def fake_ref(name, id_):
    from core.layer_index import LayerRef
    return LayerRef(id=id_, name=name, index_path=(0,), display_path=name,
                    is_group=False, is_text=False, name_path=())


def main():
    tmp = tempfile.mkdtemp(prefix="stage4_blk_")
    xlsx = os.path.join(tmp, "t.xlsx")
    make_xlsx(xlsx, [
        ["编号", "门店旧", "姓名", "电话", "门店新"],
        [1, "A门店", "张三", 13800000001, "X门店"],
        [2, "B门店", "李四", 13800000002, "Y门店"],
    ])

    root = tk.Tk()
    root.withdraw()
    app = g.App(root)
    try:
        app.xlsx_var.set(xlsx)
        app.psd_var.set("dummy.psd")

        # ---- B: col_store B -> E ----
        # 1) 首次：col_store=B，模拟 _load 的 Excel 段
        app.col_store_var.set("B")
        app.col_name_var.set("C")
        app.col_phone_var.set("D")
        app.col_role_var.set("（不替换）")
        assert app._load_excel_data(), "首次加载失败"
        print(f"[B1] 首次 col_store=B -> stores={app.excel_stores}")
        assert app.excel_stores == ["A门店", "B门店"], app.excel_stores
        assert app.excel_dataset.valid_rows[0].store == "A门店"

        # 2) 注入 fake PSD 层（模拟已点加载后的 layer_index）
        l1, l2 = fake_ref("X门店", "r-x"), fake_ref("Y门店", "r-y")
        app.layer_index = type("LI", (), {
            "layers": [l1, l2],
            "find_matching": lambda self, nm: [l for l in [l1, l2] if l.name == nm],
        })()
        app._ref_by_id = {"r-x": l1, "r-y": l2}
        app.layer_labels = {"r-x": "X门店", "r-y": "Y门店"}
        app._label_to_ref = lambda lb: {"X门店": l1, "Y门店": l2}.get(lb)
        app.logo_checks = {}
        # 预填 map_combos：模拟真实 _load 完成后「门店->Logo 映射区」已构建的状态。
        # （_ensure_dataset_fresh 的守卫 `if self.layer_index is not None and self.map_combos:`
        #   依赖此非空；真实 GUI 中 _load 一定构建过 map_combos，因此此守卫在真实场景成立。）
        app.map_combos = {
            "A门店": tk.StringVar(value="（无）"),
            "B门店": tk.StringVar(value="（无）"),
        }
        # _selected_logo_refs / _effective_logo_layers 依赖 logo_checks 与 layer_index
        app.all_psd_layers = ["X门店", "Y门店"]
        app.all_psd_is_group = {}
        app.all_psd_parent = {}
        app.all_text_layers = []
        app.logo_label_to_ref = {"X门店": l1, "Y门店": l2}
        app.text_label_to_ref = {}
        app.brand_checks = {}
        app.brand_widgets = {}
        app.map_combo_widgets = {}
        # fake _effective_logo_layers
        app._effective_logo_layers = lambda: ["X门店", "Y门店"]
        app._parent_name_of = lambda ref: ""
        app._parent_name_of_str = lambda s: ""
        # _rebuild_logo_lists 需要真实控件：用真实 UI 已创建的 logo_inner/map_inner/brand_inner

        # 3) 用户改门店列为 E -> 触发 _start 前自动重解析
        app.col_store_var.set("E")
        app._ensure_dataset_fresh(trigger="开始")
        print(f"[B2] 改 col_store=E 后 stores={app.excel_stores}")
        assert app.excel_stores == ["X门店", "Y门店"], app.excel_stores
        assert app._ds_key[2] == 4
        # Logo 映射区同步重建：map_combos 的 key 是新门店
        print(f"[B3] 重建后 map_combos keys={list(app.map_combos.keys())}")
        assert set(app.map_combos.keys()) == {"X门店", "Y门店"}, list(app.map_combos.keys())

        # 4) 配置未变 -> 不重复解析
        app._ensure_dataset_fresh(trigger="开始")
        assert app.excel_stores == ["X门店", "Y门店"]

        # ---- C: has_header True -> False ----
        app.header_var.set(False)
        app._ensure_dataset_fresh(trigger="预览")
        # 无表头：首行(原表头行)变成数据 -> name=col_name=C 列="姓名"，store=col_store=E 列="门店新"（B 部分已把门店列改为 E）
        print(f"[C1] has_header=False 后 valid_rows[0]="
              f"{app.excel_dataset.valid_rows[0].name!r} "
              f"store={app.excel_dataset.valid_rows[0].store!r} "
              f"stores={app.excel_stores[:4]}...")
        assert app.excel_dataset.valid_rows[0].excel_row == 1
        assert app.excel_dataset.valid_rows[0].name == "姓名"
        assert app.excel_dataset.valid_rows[0].store == "门店新"

        print("\nB/C 真实验证全部通过（col_store B->E 自动重建 + has_header 切换行定义）")
    finally:
        root.destroy()


if __name__ == "__main__":
    main()
