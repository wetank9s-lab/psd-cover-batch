# -*- coding: utf-8 -*-
"""core.excel_data 的 Stage 4 测试：统一 Excel 数据管线。

覆盖（对应 Stage 4 spec）：
  - 列名工具 A..Z / AA / AB...（0 起 <-> Excel 列名双向）
  - load_excel_dataset 统一入口：
      * has_header True/False
      * 门店在 B / Z / AA 列（不再固定 A 列，P1-01）
      * 电话规范化（int / integer-float / string / 空格 / 前导零保留）
      * 公式单元格 data_only 后为 None -> 视为空，跳过原因含「公式」提示
      * .xls 拒绝（提示使用 .xlsx/.xlsm，不引入 xlrd）
      * sheet 选择 = workbook.active，Dataset 记录 sheet_name
      * stores = strip、保序、去重
      * 物理行号准确（有表头第 2 行 = excel_row 2；无表头第 1 行 = 1）
      * 跳过原因保存（SkippedRow.reason）
      * 字段列冲突 preflight（姓名 != 电话 等）
      * 空 Excel -> NO_VALID_DATA
      * 统一异常 ExcelDataError（FILE_NOT_FOUND / UNSUPPORTED_TYPE /
        WORKBOOK_BROKEN / COL_OUT_OF_RANGE / NO_VALID_DATA / COL_CONFLICT）
      * 开放工作簿不入 renderer（Dataset 建立后只操作 ExcelRow）
  - _load / _preview / run_batch 一致性（同一 stores / 第一条有效行 / 字段值 / excel_row）
  - has_header 默认 True
"""
import os
import sys

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import openpyxl  # noqa: E402  (测试环境固定有 openpyxl)

from core.excel_data import (  # noqa: E402
    ExcelRow,
    ExcelDataset,
    SkippedRow,
    ExcelDataError,
    load_excel_dataset,
    index_to_excel_column,
    excel_column_to_index,
    format_phone,
    parse_row,
    parse_rows,
)


# ---------------------------------------------------------------------------
# 列名工具
# ---------------------------------------------------------------------------
def test_column_roundtrip_single():
    assert index_to_excel_column(0) == "A"
    assert index_to_excel_column(1) == "B"
    assert index_to_excel_column(25) == "Z"
    assert excel_column_to_index("A") == 0
    assert excel_column_to_index("Z") == 25


def test_column_roundtrip_double():
    assert index_to_excel_column(26) == "AA"
    assert index_to_excel_column(27) == "AB"
    assert index_to_excel_column(51) == "AZ"
    assert index_to_excel_column(52) == "BA"
    assert excel_column_to_index("AA") == 26
    assert excel_column_to_index("AB") == 27
    assert excel_column_to_index("AZ") == 51
    assert excel_column_to_index("BA") == 52


def test_column_roundtrip_triple():
    assert index_to_excel_column(701) == "ZZ"
    assert index_to_excel_column(702) == "AAA"
    assert excel_column_to_index("ZZ") == 701
    assert excel_column_to_index("AAA") == 702


def test_column_case_insensitive():
    assert excel_column_to_index("ab") == 27
    assert excel_column_to_index("aA") == 26


def test_column_invalid():
    with pytest.raises(ValueError):
        excel_column_to_index("")
    with pytest.raises(ValueError):
        excel_column_to_index("A1")
    with pytest.raises(ValueError):
        index_to_excel_column(-1)


def test_column_index_of_large_value():
    # 16384 列（XFD，Excel 上限）往返
    assert excel_column_to_index("XFD") == 16383
    assert index_to_excel_column(16383) == "XFD"


# ---------------------------------------------------------------------------
# 电话规范化（含前导零保留）
# ---------------------------------------------------------------------------
def test_phone_int():
    assert format_phone(13800000001) == "13800000001"


def test_phone_float_integer():
    assert format_phone(13800000001.0) == "13800000001"


def test_phone_float_non_integer():
    assert format_phone(123.45) == "123.45"


def test_phone_string_strip():
    assert format_phone("  18090082479  ") == "18090082479"


def test_phone_string_leading_zero_preserved():
    # spec：保留字符串前导零（'00123456' 不得变成 '123456'）
    assert format_phone("00123456") == "00123456"
    assert format_phone(" 00123456 ") == "00123456"


def test_phone_none():
    assert format_phone(None) == ""


def test_phone_bool():
    assert format_phone(True) == "1"
    assert format_phone(False) == "0"


# ---------------------------------------------------------------------------
# parse_row / parse_rows（Stage 0 兼容 + values 保留）
# ---------------------------------------------------------------------------
def test_parse_row_values_preserved():
    r = parse_row(("易田电器", "王兵", "销售顾问", 13800000001, "X"),
                  excel_row=2, col_store=0, col_name=1, col_phone=3, col_role=2)
    assert r is not None
    assert r.values == ("易田电器", "王兵", "销售顾问", "13800000001", "X")


def test_parse_rows_skipped_reason_saved():
    valid, skipped = parse_rows(
        [("门店", "姓名", "职位", "电话"),
         ("易田电器", "王兵", "x", 13800000001),
         ("易田电器", "", "x", 13800000002)], has_header=True)
    assert len(valid) == 1
    assert isinstance(skipped[0], SkippedRow)
    assert skipped[0].excel_row == 3
    assert "姓名" in skipped[0].reason


def test_parse_rows_formula_none_phone_reason():
    valid, skipped = parse_rows(
        [("门店", "姓名", "职位", "电话"),
         ("易田电器", "王兵", "x", None)], has_header=True)
    assert valid == []
    assert skipped[0].excel_row == 2
    assert "公式" in skipped[0].reason


# ---------------------------------------------------------------------------
# fixture 帮助函数
# ---------------------------------------------------------------------------
def _make_xlsx(tmp_path, rows, name="data.xlsx"):
    p = tmp_path / name
    wb = openpyxl.Workbook()
    ws = wb.active
    for row in rows:
        ws.append(row)
    wb.save(p)
    return str(p)


# ---------------------------------------------------------------------------
# load_excel_dataset：has_header / 列位置 / sheet / stores
# ---------------------------------------------------------------------------
def test_load_header_true_basic(tmp_path):
    p = _make_xlsx(tmp_path, [
        ["门店", "姓名", "职位", "电话"],
        ["易田电器", "王兵", "销售顾问", 13800000001],
        ["康乐", "张三", "销售顾问", 13800000002],
    ])
    ds = load_excel_dataset(p)   # 默认 has_header=True, A=门店 B=姓名 D=电话
    assert ds.sheet_name == "Sheet"
    assert ds.headers == ["门店", "姓名", "职位", "电话"]
    assert len(ds.valid_rows) == 2
    assert ds.valid_rows[0].excel_row == 2   # 有表头：数据从第 2 行起
    assert ds.valid_rows[1].excel_row == 3
    assert ds.stores == ["易田电器", "康乐"]
    assert ds.max_columns == 4


def test_load_header_false(tmp_path):
    p = _make_xlsx(tmp_path, [
        ["王兵", "易田电器", 13800000001],
        ["张三", "康乐", 13800000002],
    ])
    ds = load_excel_dataset(p, has_header=False, col_store=1, col_name=0, col_phone=2)
    assert ds.headers == []
    assert ds.valid_rows[0].excel_row == 1   # 无表头：第 1 行即数据
    assert ds.valid_rows[1].excel_row == 2
    assert [r.store for r in ds.valid_rows] == ["易田电器", "康乐"]


def test_load_header_default_true(tmp_path):
    # spec：has_header 默认 True
    p = _make_xlsx(tmp_path, [
        ["门店", "姓名", "职位", "电话"],
        ["易田电器", "王兵", "销售顾问", 13800000001],
    ])
    ds = load_excel_dataset(p)
    assert ds.valid_rows[0].excel_row == 2
    assert ds.valid_rows[0].name == "王兵"


def test_store_in_column_b(tmp_path):
    # P1-01：门店在 B 列（col_store=1），A 列是编号
    p = _make_xlsx(tmp_path, [
        ["编号", "门店", "姓名", "电话"],
        [1, "易田电器", "王兵", 13800000001],
        [2, "康乐", "张三", 13800000002],
    ])
    ds = load_excel_dataset(p, col_store=1, col_name=2, col_phone=3)
    assert [r.store for r in ds.valid_rows] == ["易田电器", "康乐"]
    assert ds.stores == ["易田电器", "康乐"]
    # 第一条有效行字段与物理行号
    r0 = ds.valid_rows[0]
    assert r0.store == "易田电器"
    assert r0.name == "王兵"
    assert r0.phone == "13800000001"
    assert r0.excel_row == 2
    assert r0.values[0] == "1"   # A 列是编号


def test_store_in_column_z_and_aa(tmp_path):
    # spec：门店可在 Z（25）/ AA（26）列
    rows = []
    header = [f"C{i}" for i in range(27)]
    header[25] = "门店"
    header[26] = "姓名"
    rows.append(header)
    data = [f"x{i}" for i in range(27)]
    data[25] = "AA门店"
    data[26] = "AA姓名"
    rows.append(data)
    p = _make_xlsx(tmp_path, rows)
    ds = load_excel_dataset(p, col_store=25, col_name=26, col_phone=3, require_phone=False)
    assert ds.valid_rows[0].store == "AA门店"
    assert ds.valid_rows[0].name == "AA姓名"
    assert ds.max_columns == 27
    assert ds.headers[25] == "门店"
    assert ds.headers[26] == "姓名"


def test_store_in_column_aa(tmp_path):
    # 验收第 6 项：store 独立位于 AA 列（col_store=26）。
    # 注意：Z 列放不同数据（不可与 AA 相同），防止误取。
    header = [f"C{i}" for i in range(27)]      # A..Z, AA = 26 列
    header[25] = "Z门店"                        # Z 列放干扰数据
    header[26] = "门店"                         # AA 列才是门店
    rows = [header]
    for i, (z_val, aa_val) in enumerate([("Z甲", "A门店"), ("Z乙", "B门店")], start=1):
        data = [f"x{i}-{j}" for j in range(27)]
        data[0] = i                            # A 列编号（与门店不同）
        data[25] = z_val                       # Z 列：干扰值，绝不等于 AA
        data[26] = aa_val                      # AA 列：真实门店
        rows.append(data)
    p = _make_xlsx(tmp_path, rows)
    ds = load_excel_dataset(p, col_store=26, col_name=1, col_phone=3,
                            require_phone=False)
    # store 来自 AA 列（26），不是 Z（25）
    assert ds.stores == ["A门店", "B门店"]
    assert [r.store for r in ds.valid_rows] == ["A门店", "B门店"]
    assert [r.values[25] for r in ds.valid_rows] == ["Z甲", "Z乙"]   # Z 列是干扰值
    assert [r.values[26] for r in ds.valid_rows] == ["A门店", "B门店"]  # AA 列真实来源


def test_stores_strip_order_dedup(tmp_path):
    p = _make_xlsx(tmp_path, [
        ["门店", "姓名", "职位", "电话"],
        ["  易田电器  ", "王兵", "销售顾问", 13800000001],
        ["康乐", "张三", "销售顾问", 13800000002],
        ["易田电器", "李四", "销售顾问", 13800000003],
    ])
    ds = load_excel_dataset(p)
    assert ds.stores == ["易田电器", "康乐"]   # strip、保序、去重
    assert ds.valid_rows[0].store == "易田电器"  # 已 strip


def test_skipped_row_physical_numbers(tmp_path):
    p = _make_xlsx(tmp_path, [
        ["门店", "姓名", "职位", "电话"],
        ["易田电器", "王兵", "销售顾问", 13800000001],
        ["易田电器", "", "销售顾问", 13800000002],     # 姓名为空 -> 跳过
        ["易田电器", "李四", "销售顾问", None],          # 电话为空 -> 跳过
        ["易田电器", "赵六", "销售顾问", 13800000004],
    ])
    ds = load_excel_dataset(p)
    assert [r.excel_row for r in ds.valid_rows] == [2, 5]
    assert [s.excel_row for s in ds.skipped_rows] == [3, 4]
    assert all(s.reason for s in ds.skipped_rows)      # 跳过原因必须保存
    assert "姓名" in ds.skipped_rows[0].reason
    assert "电话" in ds.skipped_rows[1].reason


def test_formula_cell_none_skipped(tmp_path):
    # spec：公式单元格 data_only=True 后为 None -> 视为空，跳过原因含「公式」提示
    p = _make_xlsx(tmp_path, [
        ["门店", "姓名", "职位", "电话"],
        ["易田电器", "张三", "销售顾问", "=A2"],   # 公式未计算缓存
    ])
    with pytest.raises(ExcelDataError) as ei:
        load_excel_dataset(p)
    assert ei.value.code == "NO_VALID_DATA"
    assert "公式" in str(ei.value)


# ---------------------------------------------------------------------------
# 异常分类
# ---------------------------------------------------------------------------
def test_error_file_not_found(tmp_path):
    with pytest.raises(ExcelDataError) as ei:
        load_excel_dataset(str(tmp_path / "nope.xlsx"))
    assert ei.value.code == "FILE_NOT_FOUND"


def test_error_xls_rejected(tmp_path):
    p = tmp_path / "old.xls"
    p.write_bytes(b"fake")
    with pytest.raises(ExcelDataError) as ei:
        load_excel_dataset(str(p))
    assert ei.value.code == "UNSUPPORTED_TYPE"
    assert ".xlsx" in str(ei.value) and ".xlsm" in str(ei.value)


def test_error_workbook_broken(tmp_path):
    p = tmp_path / "broken.xlsx"
    p.write_bytes(b"this is not a real xlsx zip")
    with pytest.raises(ExcelDataError) as ei:
        load_excel_dataset(str(p))
    assert ei.value.code == "WORKBOOK_BROKEN"


def test_error_column_out_of_range(tmp_path):
    p = _make_xlsx(tmp_path, [
        ["门店", "姓名"],
        ["易田电器", "王兵"],
    ])
    with pytest.raises(ExcelDataError) as ei:
        load_excel_dataset(p, col_store=0, col_name=1, col_phone=3)  # D 越界
    assert ei.value.code == "COL_OUT_OF_RANGE"
    assert "D" in str(ei.value)


def test_error_no_valid_data(tmp_path):
    p = _make_xlsx(tmp_path, [["门店", "姓名", "职位", "电话"]])  # 只有表头
    with pytest.raises(ExcelDataError) as ei:
        load_excel_dataset(p)
    assert ei.value.code == "NO_VALID_DATA"
    assert "没有可生成的有效数据" in str(ei.value)


def test_error_column_conflict(tmp_path):
    # spec：姓名和电话不能使用同一 Excel 列（先于列越界检查触发）
    p = _make_xlsx(tmp_path, [["门店", "姓名", "职位", "电话"]])
    with pytest.raises(ExcelDataError) as ei:
        load_excel_dataset(p, col_store=0, col_name=1, col_phone=1)   # 姓名==电话==B
    assert ei.value.code == "COL_CONFLICT"
    assert "姓名和电话" in str(ei.value)


def test_error_column_conflict_store_name(tmp_path):
    p = _make_xlsx(tmp_path, [["门店", "姓名", "职位", "电话"]])
    with pytest.raises(ExcelDataError) as ei:
        load_excel_dataset(p, col_store=0, col_name=0)
    assert ei.value.code == "COL_CONFLICT"


def test_xlsm_supported(tmp_path):
    # 验收第 34 项：.xlsm 由 pytest 独立真实创建 + 读取（非 suffix 检查，不依赖仓库固定文件）。
    # 内容对齐 stage4_blk_g.py：sheet / stores / valid_rows / values 全量断言。
    p = tmp_path / "macro.xlsm"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "门店数据"
    ws.append(["编号", "门店", "姓名", "电话", "销售顾问"])
    ws.append([1, "康乐电器", "张三", 13800000001, "王顾问"])
    ws.append([2, "美的专卖", "李四", 13800000002, "陈顾问"])
    wb.save(p)
    wb.close()
    ds = load_excel_dataset(str(p), has_header=True,
                            col_store=1, col_name=2, col_phone=3, col_role=4)
    assert ds.sheet_name == "门店数据"
    assert ds.stores == ["康乐电器", "美的专卖"]
    assert len(ds.valid_rows) == 2
    r0 = ds.valid_rows[0]
    assert (r0.store, r0.name, r0.phone, r0.role) == \
        ("康乐电器", "张三", "13800000001", "王顾问")
    assert r0.values[0] == "1" and r0.values[1] == "康乐电器"   # values 为 str 形式


# ---------------------------------------------------------------------------
# 工作簿生命周期：Dataset 建立后不持有工作簿（开放工作簿不入 renderer）
# ---------------------------------------------------------------------------
def test_dataset_does_not_hold_workbook(tmp_path):
    import gc
    p = _make_xlsx(tmp_path, [
        ["门店", "姓名", "职位", "电话"],
        ["易田电器", "王兵", "销售顾问", 13800000001],
    ])
    ds = load_excel_dataset(p)
    # ExcelRow 是纯数据，不含 openpyxl 对象
    assert not any(hasattr(r, "_workbook") for r in ds.valid_rows)
    assert not hasattr(ds, "_workbook")
    # 原始行也是纯 tuple
    assert all(isinstance(r, tuple) for r in ds.rows)


# ---------------------------------------------------------------------------
# 一致性证明：同一文件，_load / _preview / run_batch 读取相同的
# stores / 第一条有效行 / 字段值 / excel_row
# ---------------------------------------------------------------------------
def test_consistency_across_entries(tmp_path):
    p = _make_xlsx(tmp_path, [
        ["编号", "门店", "姓名", "职位", "电话"],
        [1, "易田电器", "王兵", "销售顾问", 13800000001],
        [2, "康乐", "张三", "销售顾问", 13800000002],
        [3, "易田电器", "李四", "销售顾问", 13800000003],
    ])
    kw = dict(col_store=1, col_name=2, col_phone=4, col_role=3)
    # 同一次调用（模拟三个入口分别加载，结果必须一致）
    ds_a = load_excel_dataset(p, has_header=True, **kw)
    ds_b = load_excel_dataset(p, has_header=True, **kw)
    ds_c = load_excel_dataset(p, has_header=True, **kw)
    # 1) stores 一致
    assert ds_a.stores == ds_b.stores == ds_c.stores == ["易田电器", "康乐"]
    # 2) 第一条有效行一致
    for ds in (ds_a, ds_b, ds_c):
        r0 = ds.valid_rows[0]
        assert r0.store == "易田电器"
        assert r0.name == "王兵"
        assert r0.phone == "13800000001"
        assert r0.role == "销售顾问"
        assert r0.excel_row == 2
        assert r0.values == ("1", "易田电器", "王兵", "销售顾问", "13800000001")
    # 3) 全部有效行 excel_row 一致
    assert [r.excel_row for r in ds_a.valid_rows] == [2, 3, 4]


def test_consistency_header_false(tmp_path):
    p = _make_xlsx(tmp_path, [
        ["王兵", "易田电器", 13800000001],
        ["张三", "康乐", 13800000002],
    ])
    ds1 = load_excel_dataset(p, has_header=False, col_store=1, col_name=0, col_phone=2)
    ds2 = load_excel_dataset(p, has_header=False, col_store=1, col_name=0, col_phone=2)
    assert ds1.stores == ds2.stores == ["易田电器", "康乐"]
    assert ds1.valid_rows[0].excel_row == ds2.valid_rows[0].excel_row == 1


# ---------------------------------------------------------------------------
# 电话在 dataset 中的规范化
# ---------------------------------------------------------------------------
def test_dataset_phone_normalization(tmp_path):
    p = _make_xlsx(tmp_path, [
        ["门店", "姓名", "职位", "电话"],
        ["易田电器", "王兵", "销售顾问", 13800000001.0],      # float 整数
        ["康乐", "张三", "销售顾问", " 13800000002 "],        # 带空格 str
        ["红旗", "李四", "销售顾问", "00123456"],             # 前导零保留
    ])
    ds = load_excel_dataset(p)
    assert ds.valid_rows[0].phone == "13800000001"
    assert ds.valid_rows[1].phone == "13800000002"
    assert ds.valid_rows[2].phone == "00123456"   # 前导零保留


# ---------------------------------------------------------------------------
# 角色列可选 / 不需要 phone 的场景
# ---------------------------------------------------------------------------
def test_dataset_role_none(tmp_path):
    p = _make_xlsx(tmp_path, [
        ["门店", "姓名", "职位", "电话"],
        ["易田电器", "王兵", "销售顾问", 13800000001],
    ])
    ds = load_excel_dataset(p, col_role=None)
    assert ds.valid_rows[0].role is None


def test_dataset_require_phone_false(tmp_path):
    p = _make_xlsx(tmp_path, [
        ["门店", "姓名", "职位", "电话"],
        ["易田电器", "王兵", "销售顾问", None],
    ])
    ds = load_excel_dataset(p, require_phone=False)
    assert ds.valid_rows[0].phone == ""
    assert ds.valid_rows[0].name == "王兵"


# ---------------------------------------------------------------------------
# Stage 4 验收补测（BLOCKED 清单）
# ---------------------------------------------------------------------------
# 场景 14：门店为空（require_store=True）-> skipped
def test_store_blank_with_require_store_skipped(tmp_path):
    p = _make_xlsx(tmp_path, [
        ["门店", "姓名", "职位", "电话"],
        [None, "王兵", "销售顾问", 13800000001],     # 门店空
        ["易田电器", "张三", "销售顾问", 13800000002],
    ])
    ds = load_excel_dataset(p, require_store=True)
    assert [r.excel_row for r in ds.valid_rows] == [3]
    assert [s.excel_row for s in ds.skipped_rows] == [2]
    assert "门店" in ds.skipped_rows[0].reason


# 场景 14b：require_store=False 显式关闭 —— 门店空仍算有效（仅 name/phone 必填）
def test_store_blank_not_required_when_disabled(tmp_path):
    p = _make_xlsx(tmp_path, [
        ["门店", "姓名", "职位", "电话"],
        [None, "王兵", "销售顾问", 13800000001],
    ])
    ds = load_excel_dataset(p, require_store=False)
    assert len(ds.valid_rows) == 1
    assert ds.valid_rows[0].store == ""
    assert ds.stores == []       # 空门店不进入 stores 列表


# 场景 36：unsupported extension rejected（.txt / 无扩展名）
def test_error_unsupported_extension_txt(tmp_path):
    p = tmp_path / "data.txt"
    p.write_text("not excel", encoding="utf-8")
    with pytest.raises(ExcelDataError) as ei:
        load_excel_dataset(str(p))
    assert ei.value.code == "UNSUPPORTED_TYPE"


def test_error_no_extension(tmp_path):
    p = tmp_path / "noext"
    p.write_bytes(b"x")
    with pytest.raises(ExcelDataError) as ei:
        load_excel_dataset(str(p))
    assert ei.value.code == "UNSUPPORTED_TYPE"


# 场景 40：blank store 不进入 stores（保序去重且跳过空）
def test_blank_store_not_included(tmp_path):
    p = _make_xlsx(tmp_path, [
        ["门店", "姓名", "职位", "电话"],
        ["易田电器", "王兵", "销售顾问", 13800000001],
        ["", "张三", "销售顾问", 13800000002],        # 空字符串门店（默认 require_store -> 跳过）
        ["   ", "李四", "销售顾问", 13800000003],      # 空白门店（默认 require_store -> 跳过）
        ["康乐", "赵六", "销售顾问", 13800000004],
    ])
    ds = load_excel_dataset(p)   # 默认 require_store=True
    assert ds.stores == ["易田电器", "康乐"]           # 空白门店被排除（且被跳过）
    assert [r.store for r in ds.valid_rows] == ["易田电器", "康乐"]
    assert [s.excel_row for s in ds.skipped_rows] == [3, 4]


# 场景 18：第一数据行无效、第二数据有效 -> valid_rows[0] 是第二数据行
def test_first_data_row_invalid_second_valid(tmp_path):
    p = _make_xlsx(tmp_path, [
        ["门店", "姓名", "职位", "电话"],
        ["易田电器", "", "销售顾问", 13800000001],     # 姓名为空 -> 无效
        ["康乐", "张三", "销售顾问", 13800000002],     # 第一条有效
    ])
    ds = load_excel_dataset(p)
    assert len(ds.valid_rows) == 1
    assert ds.valid_rows[0].excel_row == 3
    assert ds.valid_rows[0].store == "康乐"
    assert ds.valid_rows[0].name == "张三"
    assert [s.excel_row for s in ds.skipped_rows] == [2]


# ---------------------------------------------------------------------------
# BLOCKED B：字段映射改变后 Dataset 不会陈旧（store B -> E）
# 与 _ensure_dataset_fresh 的 GUI 行为（见 test_gui_* 无头集成测试）
# ---------------------------------------------------------------------------
def test_store_column_change_rebuilds_dataset(tmp_path):
    """同会话内 col_store 从 B 改为 E 后重新解析 -> stores 用新列（核心逻辑）。"""
    p = _make_xlsx(tmp_path, [
        ["编号", "门店旧", "姓名", "电话", "门店新"],
        [1, "A门店", "张三", 13800000001, "X门店"],
        [2, "B门店", "李四", 13800000002, "Y门店"],
    ])
    # 第一次：col_store=B(1)
    ds1 = load_excel_dataset(p, col_store=1, col_name=2, col_phone=3)
    assert ds1.stores == ["A门店", "B门店"]
    # 第二次（同会话，新配置）：col_store=E(4)
    ds2 = load_excel_dataset(p, col_store=4, col_name=2, col_phone=3)
    assert ds2.stores == ["X门店", "Y门店"]
    # 第一条有效行也从 E 列取门店
    assert ds2.valid_rows[0].store == "X门店"
    assert ds2.valid_rows[0].name == "张三"


def test_has_header_change_rebuilds_dataset(tmp_path):
    """同会话内 has_header 从 True 切 False -> 行定义变化（首行被当作数据）。"""
    p = _make_xlsx(tmp_path, [
        ["门店", "姓名", "职位", "电话"],
        ["易田电器", "王兵", "销售顾问", 13800000001],
    ])
    ds_true = load_excel_dataset(p, has_header=True)
    assert ds_true.valid_rows[0].excel_row == 2
    assert ds_true.valid_rows[0].name == "王兵"
    # 同会话切 has_header=False：表头行变成数据行（name=表头值）
    ds_false = load_excel_dataset(p, has_header=False)
    assert ds_false.valid_rows[0].excel_row == 1
    assert ds_false.valid_rows[0].name == "姓名"     # 表头行被视为数据
    assert ds_false.valid_rows[1].name == "王兵"


# ---------------------------------------------------------------------------
# BLOCKED B/C：GUI 无头集成 —— 模拟 _load -> 改列 -> _start/_preview 前自动重解析
# ---------------------------------------------------------------------------
def _gui_app_instance():
    """创建真实 tk App 实例（不 mainloop），返回 (root, app)。"""
    import tkinter as tk
    import qifang_cover_maker as g
    root = tk.Tk()
    root.withdraw()
    app = g.App(root)
    return root, app


def test_gui_store_column_change_rebuilds_mapping(tmp_path):
    """GUI 同会话：col_store B -> E 后，_ensure_dataset_fresh 自动重解析 + 重建映射区。

    证明 Logo 页 store list 与 Batch 使用同一新列（不会出现旧列表 + 新列错配）。
    """
    import qifang_cover_maker as g
    p = _make_xlsx(tmp_path, [
        ["编号", "门店旧", "姓名", "电话", "门店新"],
        [1, "A门店", "张三", 13800000001, "X门店"],
        [2, "B门店", "李四", 13800000002, "Y门店"],
    ])
    app = g.App.__new__(g.App)  # 不跑 __init__（避免全量 UI），手工构造最小状态
    try:
        # 最小 GUI 状态：Excel 段所需控件 + Logo 映射所需
        app.xlsx_var = type("V", (), {"get": lambda self: p, "set": lambda self, v: None})()
        app.header_var = type("V", (), {"get": lambda self: True, "set": lambda self, v: None})()
        app.col_store_var = type("V", (), {"get": lambda self: "B", "set": lambda self, v: None})()
        app.col_name_var = type("V", (), {"get": lambda self: "C", "set": lambda self, v: None})()
        app.col_phone_var = type("V", (), {"get": lambda self: "D", "set": lambda self, v: None})()
        app.col_role_var = type("V", (), {"get": lambda self: "（不替换）", "set": lambda self, v: None})()
        app.excel_dataset = None
        app._ds_key = None
        app.excel_stores = []
        app.excel_headers = []
        app.layer_index = None
        app.map_combos = {}
        app._column_labels = []
        app.col_store_cb = None
        app.col_name_cb = None
        app.col_phone_cb = None
        app.col_role_cb = None
        # _log_gui 用打印替代（避免依赖真实 tk 日志控件）
        def _fake_log(self, msg):
            pass
        app._log_gui = _fake_log.__get__(app)

        # 1) 首次加载（col_store=B）
        assert app._load_excel_data(p)
        assert app.excel_stores == ["A门店", "B门店"]
        assert app._ds_key[2] == 1   # col_store 索引 1

        # 2) 用户改门店列为 E
        app.col_store_var.get = lambda: "E"
        # 此时 _ds_key 仍为旧值 -> _ensure_dataset_fresh 应自动重解析
        assert app._ensure_dataset_fresh(trigger="开始")
        assert app.excel_stores == ["X门店", "Y门店"]
        assert app._ds_key[2] == 4   # col_store 索引 4
        # 3) 配置未变 -> 不重复解析（仍一致）
        assert app._ensure_dataset_fresh(trigger="开始")
        assert app.excel_stores == ["X门店", "Y门店"]
    finally:
        pass

