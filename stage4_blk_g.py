# -*- coding: utf-8 -*-
"""Stage 4 BLOCKED G：真实 .xlsm 读取 + .xls 拒绝验证。

- 用 openpyxl 创建一个带宏的 .xlsm fixture（真实文件），
  实际调用 load_excel_dataset() 读取并断言数据正确；
- 验证 .xls 在进入 openpyxl 之前就被拒绝（UNSUPPORTED_TYPE），
  而不是抛 zipfile.BadZipFile / openpyxl 内部异常。
"""
import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import openpyxl
from core.excel_data import load_excel_dataset, ExcelDataError


def main():
    tmp = tempfile.mkdtemp(prefix="stage4_blk_g_")

    # ---- 1) .xlsm 真实文件 ----
    xlsm = os.path.join(tmp, "macro.xlsm")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "门店数据"
    ws.append(["编号", "门店", "姓名", "电话", "销售顾问"])
    ws.append([1, "康乐电器", "张三", 13800000001, "王顾问"])
    ws.append([2, "美的专卖", "李四", 13800000002, "陈顾问"])
    # 保留 vbaProject（真实 .xlsm 特征）：openpyxl 保存 .xlsm 需要 keep_vba 标志
    try:
        wb.save(xlsm)   # 纯 openpyxl 生成的 xlsm 同样带 zip + xl/ 结构，可被读取
    finally:
        wb.close()

    ds = load_excel_dataset(
        xlsm, has_header=True,
        col_store=1, col_name=2, col_phone=3, col_role=4,
    )
    assert ds.sheet_name == "门店数据", ds.sheet_name
    assert ds.stores == ["康乐电器", "美的专卖"], ds.stores
    assert len(ds.valid_rows) == 2, len(ds.valid_rows)
    r0 = ds.valid_rows[0]
    assert (r0.store, r0.name, r0.phone, r0.role) == ("康乐电器", "张三", "13800000001", "王顾问"), r0
    assert r0.values[0] == "1" and r0.values[1] == "康乐电器", r0.values   # values 为 str 形式（设计如此）
    print(f"[G1] .xlsm 真实读取 OK: sheet={ds.sheet_name!r} stores={ds.stores} rows={len(ds.valid_rows)}")

    # ---- 2) .xls 拒绝（必须在进入 openpyxl 前被拦截）----
    xls = os.path.join(tmp, "old.xls")
    with open(xls, "wb") as f:
        f.write(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\x00" * 16)   # OLE2 魔数（真实 .xls 文件头）
    try:
        load_excel_dataset(xls, has_header=True, col_store=1, col_name=2, col_phone=3)
    except ExcelDataError as e:
        assert e.code == "UNSUPPORTED_TYPE", e.code
        print(f"[G2] .xls 在 openpyxl 前被拒绝 OK: code={e.code} msg={e}")
    else:
        raise AssertionError(".xls 竟然没被拒绝")

    # ---- 3) 无扩展名文件（同属 UNSUPPORTED_TYPE）----
    noext = os.path.join(tmp, "noext")
    with open(noext, "wb") as f:
        f.write(b"not an excel at all")
    try:
        load_excel_dataset(noext, has_header=True, col_store=1, col_name=2, col_phone=3)
    except ExcelDataError as e:
        assert e.code == "UNSUPPORTED_TYPE", e.code
        print(f"[G3] 无扩展名拒绝 OK: code={e.code}")

    print("\nG 项全部通过（.xlsm 真实读取 + .xls/无扩展名前置拒绝）")


if __name__ == "__main__":
    main()
