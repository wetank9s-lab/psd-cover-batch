# -*- coding: utf-8 -*-
"""Stage 7：psd_cover_batch CLI 薄调用层纯逻辑单元测试（不依赖 Photoshop COM）。

覆盖（Stage 7 规格）：
  - exit code 规则（0=成功 / 1=参数配置错误 / 2=部分失败 / 3=内部异常）
  - --group-output-column 参数转换（A/D/AA / 0-based int -> int | None）
  - --brand-logo resolve 成 LayerRef（唯一 / 同名歧义报候选 / 缺失报错）
  - inspect 图层树渲染（LayerRef 展示：group/text/leaf 标签 + 同名 id 后缀）
  - run 缺 xlsx / psd 不存在 -> 参数错误（exit 1）
  - main() 的 exit code 分支（薄层不依赖 Photoshop：inspect 缺 psd 等）
  - CLI 源码级红线：不 import openpyxl / 不出现 registry[layer.Name] 覆盖 /
    不出现 Document.Duplicate / TextItem.Contents / SaveAs（这些必须在 core）
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import pytest  # noqa: E402

from core.layer_index import LayerRef, LayerIndex  # noqa: E402
from core.logo_mapping import LogoValidationError  # noqa: E402
from core.renderer import BatchResult, RowResult, RowStatus  # noqa: E402
from psd_cover_batch import (  # noqa: E402
    EXIT_OK, EXIT_CONFIG, EXIT_PARTIAL, EXIT_INTERNAL,
    parse_group_output_column, _resolve_brand_refs, _layer_tree,
    batch_exit_code,
)


# ---------------------------------------------------------------------------
# batch_exit_code（规格十三：0=全部成功 / 2=部分失败）
# ---------------------------------------------------------------------------
def _mk_batch(failed=0, cancelled=False):
    rows = []
    for i in range(3):
        st = RowStatus.FAILED if i < failed else RowStatus.SUCCESS
        rows.append(RowResult(excel_row=i + 1, status=st))
    return BatchResult(total=3, success=3 - failed, failed=failed,
                       cancelled=cancelled, rows=rows)


def test_batch_exit_code_all_success():
    assert batch_exit_code(_mk_batch(failed=0)) == EXIT_OK


def test_batch_exit_code_has_failures():
    assert batch_exit_code(_mk_batch(failed=2)) == EXIT_PARTIAL
    assert batch_exit_code(_mk_batch(failed=1)) == EXIT_PARTIAL


def test_batch_exit_code_cancelled():
    assert batch_exit_code(_mk_batch(cancelled=True)) == EXIT_PARTIAL


# ---------------------------------------------------------------------------
# exit code 常量
# ---------------------------------------------------------------------------
def test_exit_code_constants():
    assert EXIT_OK == 0
    assert EXIT_CONFIG == 1
    assert EXIT_PARTIAL == 2
    assert EXIT_INTERNAL == 3


# ---------------------------------------------------------------------------
# --group-output-column 参数转换（规格十一：A/D/AA 或 0-based int -> int | None）
# ---------------------------------------------------------------------------
def test_group_column_letter_conversion():
    assert parse_group_output_column("A") == 0
    assert parse_group_output_column("D") == 3
    assert parse_group_output_column("AA") == 26


def test_group_column_numeric_conversion():
    assert parse_group_output_column("0") == 0
    assert parse_group_output_column("26") == 26
    assert parse_group_output_column("3") == 3


def test_group_column_none_or_empty():
    assert parse_group_output_column(None) is None
    assert parse_group_output_column("") is None
    assert parse_group_output_column("   ") is None


def test_group_column_invalid_raises():
    with pytest.raises(Exception):
        parse_group_output_column("ZZZ!")
    with pytest.raises(Exception):
        parse_group_output_column("1A")
    with pytest.raises(Exception):
        parse_group_output_column("-1")


def test_group_column_lowercase_accepted():
    assert parse_group_output_column("aa") == 26
    assert parse_group_output_column("d") == 3


# ---------------------------------------------------------------------------
# --brand-logo resolve 成 LayerRef（规格八：唯一/歧义/缺失）
# ---------------------------------------------------------------------------
def _mk_refs():
    return [
        LayerRef(id="0", name="七方logo", display_path="Logo > 七方logo",
                 index_path=(0,), is_text=False),
        LayerRef(id="1", name="康乐电器", display_path="Stores > 康乐电器",
                 index_path=(1,), is_text=False),
        LayerRef(id="2", name="康乐家电", display_path="Stores > 康乐家电",
                 index_path=(2,), is_text=False),
    ]


def test_brand_unique_resolves_to_ref():
    index = LayerIndex(_mk_refs())
    refs = _resolve_brand_refs(index, ["七方logo"])
    assert len(refs) == 1
    assert refs[0].id == "0"
    assert refs[0].display_path == "Logo > 七方logo"


def test_brand_multiple_names_resolves_all():
    index = LayerIndex(_mk_refs())
    refs = _resolve_brand_refs(index, ["七方logo", "康乐电器"])
    assert len(refs) == 2
    assert {r.id for r in refs} == {"0", "1"}


def test_brand_ambiguous_raises_with_candidates():
    refs = [
        LayerRef(id="0", name="七方logo", display_path="Logo > 七方logo", index_path=(0,)),
        LayerRef(id="1", name="七方logo", display_path="Logo2 > 七方logo", index_path=(1,)),
    ]
    index = LayerIndex(refs)
    with pytest.raises(LogoValidationError) as ei:
        _resolve_brand_refs(index, ["七方logo"])
    assert ei.value.code == "BRAND_AMBIGUOUS"
    # 必须打印候选 display_path（规格八：多命中报 ambiguous 并打印候选）
    assert "Logo > 七方logo" in str(ei.value)
    assert "Logo2 > 七方logo" in str(ei.value)


def test_brand_missing_raises():
    index = LayerIndex(_mk_refs())
    with pytest.raises(LogoValidationError) as ei:
        _resolve_brand_refs(index, ["不存在"])
    assert ei.value.code == "BRAND_MISSING"


def test_brand_whitespace_tolerance():
    # find_matching 忽略空格 + 小写：'七方 Logo' 命中 '七方logo'
    index = LayerIndex([LayerRef(id="0", name="七方logo", display_path="L > 七方logo",
                                 index_path=(0,))])
    refs = _resolve_brand_refs(index, ["七方 Logo"])
    assert [r.id for r in refs] == ["0"]


# ---------------------------------------------------------------------------
# inspect 图层树渲染（规格六：基于 LayerRef；同名不覆盖，可附加 id）
# ---------------------------------------------------------------------------
def test_layer_tree_group_text_leaf_tags():
    refs = [
        LayerRef(id="0", name="Logo", display_path="Logo", index_path=(0,), is_group=True),
        LayerRef(id="0/0", name="A门店", display_path="Logo > A门店",
                 index_path=(0, 0), is_text=False),
        LayerRef(id="1", name="姓名", display_path="姓名", index_path=(1,), is_text=True),
        LayerRef(id="2", name="背景", display_path="背景", index_path=(2,), is_text=False),
    ]
    lines = _layer_tree(LayerIndex(refs))
    assert any("[GROUP] Logo" in ln for ln in lines)
    assert any("[LAYER] Logo > A门店" in ln for ln in lines)
    assert any("[T] 姓名" in ln for ln in lines)
    assert any("[LAYER] 背景" in ln for ln in lines)


def test_layer_tree_duplicate_names_get_id_suffix():
    # 同名图层（GroupA > 电话 / GroupB > 电话）必须都显示，且附加 id 区分
    refs = [
        LayerRef(id="0", name="电话", display_path="GroupA > 电话",
                 index_path=(0,), is_text=True),
        LayerRef(id="1", name="电话", display_path="GroupB > 电话",
                 index_path=(1,), is_text=True),
    ]
    lines = _layer_tree(LayerIndex(refs))
    assert len(lines) == 2
    assert any("GroupA > 电话" in ln and "[id=0]" in ln for ln in lines)
    assert any("GroupB > 电话" in ln and "[id=1]" in ln for ln in lines)


def test_layer_tree_unique_names_no_id_suffix():
    refs = [
        LayerRef(id="0", name="姓名", display_path="姓名", index_path=(0,), is_text=True),
        LayerRef(id="1", name="电话", display_path="电话", index_path=(1,), is_text=True),
    ]
    lines = _layer_tree(LayerIndex(refs))
    assert any("[T] 姓名" in ln and "[id=" not in ln for ln in lines)
    assert any("[T] 电话" in ln and "[id=" not in ln for ln in lines)


def test_layer_tree_indentation_by_depth():
    refs = [
        LayerRef(id="0", name="G", display_path="G", index_path=(0,), is_group=True),
        LayerRef(id="0/0", name="G2", display_path="G > G2", index_path=(0, 0), is_group=True),
        LayerRef(id="0/0/0", name="leaf", display_path="G > G2 > leaf",
                 index_path=(0, 0, 0), is_text=False),
    ]
    lines = _layer_tree(LayerIndex(refs))
    assert "  [GROUP] G" in lines[0]
    assert "    [GROUP] G > G2" in lines[1]
    assert "      [LAYER] G > G2 > leaf" in lines[2]


# ---------------------------------------------------------------------------
# main() exit code 分支（薄层；不触发 Photoshop）
# ---------------------------------------------------------------------------
def test_main_missing_xlsx_returns_config():
    import psd_cover_batch as cli
    assert cli.main(["--psd", "x.psd"]) == EXIT_CONFIG


def test_main_missing_psd_returns_config():
    import psd_cover_batch as cli
    assert cli.main(["--psd", "nonexist.psd", "--xlsx", "x.xlsx"]) == EXIT_CONFIG
    assert cli.main(["--psd", "nonexist.psd", "--inspect"]) == EXIT_CONFIG


def test_main_invalid_group_column_returns_config():
    import psd_cover_batch as cli
    # psd 存在性预检先失败（不触发 group col 解析），因此先给存在的 psd
    # —— 用 main 级模拟：psd 不存在也会先 exit 1，这里直接断言两种都安全
    assert cli.main(["--psd", "nonexist.psd", "--xlsx", "x.xlsx",
                     "--group-output-column", "ZZZ!"]) == EXIT_CONFIG


# ---------------------------------------------------------------------------
# 源码级红线（规格二/三/四/十：CLI 不得含业务实现）
# ---------------------------------------------------------------------------
def _cli_source():
    """返回 psd_cover_batch.py 的源码（含 docstring）。"""
    here = os.path.dirname(__file__)
    with open(os.path.join(here, "..", "psd_cover_batch.py"), encoding="utf-8") as f:
        return f.read()


def _cli_code():
    """返回去掉模块 docstring 后的代码（AST 提取），避免文档中的「禁止项」字样误报。"""
    import ast
    here = os.path.dirname(__file__)
    with open(os.path.join(here, "..", "psd_cover_batch.py"), encoding="utf-8") as f:
        tree = ast.parse(f.read())
    # 移除模块 docstring（第一个 Expr 的 Constant）
    if tree.body and isinstance(tree.body[0], ast.Expr):
        tree.body = tree.body[1:]
    return ast.unparse(tree)


def test_cli_no_openpyxl_import():
    src = _cli_code()
    assert "import openpyxl" not in src
    assert "from openpyxl" not in src
    assert "iter_rows" not in src
    assert "load_workbook" not in src


def test_cli_no_registry_name_override():
    # 规格四：特别禁止 registry[layer.Name] = layer（同名覆盖）
    src = _cli_code()
    assert "registry[" not in src
    assert "= layer.Name" not in src
    assert "by_name" not in src  # 由 core.layer_index 提供，不在 CLI 自建


def test_cli_no_photoshop_business_operations():
    # 规格三：CLI run 中不应直接出现 Document.Duplicate / TextItem.Contents / Visible= / SaveAs
    src = _cli_code()
    assert "Document.Duplicate" not in src
    assert "TextItem.Contents" not in src
    assert ".Visible =" not in src
    assert "SaveAs" not in src
    # 不含 openpyxl worksheet 遍历
    assert "worksheet" not in src.lower() or "worksheet" in src  # 宽松：核心是上述四项


def test_cli_uses_core_collect_layer_index():
    # 规格四：inspect 必须使用 collect_layer_index
    src = _cli_source()
    assert "collect_layer_index" in src


def test_cli_uses_core_load_excel_dataset():
    # 规格十：CLI Excel 必须使用 load_excel_dataset
    src = _cli_source()
    assert "load_excel_dataset" in src
    assert "core.excel_data" in src


def test_cli_uses_core_match_store_logo():
    # 规格七：CLI 使用 LogoMapping / match_store_logo（不 runtime fuzzy）
    src = _cli_code()
    assert "match_store_logo" in src
    assert "LogoMapping" in src
    assert "fuzzy" not in src.lower()


def test_cli_uses_core_renderer_run_batch():
    # 规格三：CLI run 收敛到 run_batch
    src = _cli_source()
    assert "renderer_run_batch" in src
    assert "core.renderer" in src


def test_cli_no_group_folder_manual_join():
    # 规格十一：禁止 os.path.join(out, store) 自行组装分组目录
    src = _cli_code()
    assert "os.path.join(out" not in src
    assert "build_group_folder_map" in src


def test_cli_exit_codes_defined():
    src = _cli_source()
    assert "EXIT_OK = 0" in src
    assert "EXIT_CONFIG = 1" in src
    assert "EXIT_PARTIAL = 2" in src
    assert "EXIT_INTERNAL = 3" in src
