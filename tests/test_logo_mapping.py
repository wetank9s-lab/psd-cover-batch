# -*- coding: utf-8 -*-
"""
test_logo_mapping.py —— Stage 3 Logo 模型与映射逻辑纯 Python 测试。

覆盖（对应 Stage 3 需求第二十九节）：
  selection / expansion    1-6
  store mapping            7-12
  brand                   13-16
  matching                17-23
  visibility plan         24-28
  old config              29-33

fake 模型复用 tests/test_layer_index.py（1-based Layers、build_doc）。
"""
import pytest

from core.layer_index import (
    LayerRef, collect_layer_index, LayerIndex,
)
from core import logo_mapping as lm
from tests.test_layer_index import build_doc


# ---------------------------------------------------------------------------
# 通用 fixture：最小 Logo PSD（与 Stage 3 第三十节一致）
#   Brand
#   └─ 七方Logo
#   Stores
#   ├─ 康乐电器
#   ├─ 九兴电器
#   ├─ 诚信电器
#   └─ 康乐家电
# ---------------------------------------------------------------------------
def make_doc():
    return build_doc("d", {
        "Brand": {"七方Logo": "text"},
        "Stores": {"康乐电器": "text", "九兴电器": "text",
                   "诚信电器": "text", "康乐家电": "text"},
    })


@pytest.fixture
def idx():
    return collect_layer_index(make_doc())


@pytest.fixture
def all_leaves(idx):
    return [r for r in idx.layers if not r.is_group]


# ===========================================================================
# 1. selection / expansion（第 1-6 项）
# ===========================================================================
def test_leaf_selection_keeps_itself(idx):
    r = idx.get("1/0")          # 康乐电器（叶子）
    eff = lm.resolve_effective_logo_layers(idx, [r])
    assert [x.id for x in eff] == ["1/0"]
    assert all(not x.is_group for x in eff)


def test_group_selection_expands_leaves(idx):
    stores = idx.get("1")       # Stores 组
    eff = lm.resolve_effective_logo_layers(idx, [stores])
    assert [x.name for x in eff] == ["康乐电器", "九兴电器", "诚信电器", "康乐家电"]
    assert all(not x.is_group for x in eff)
    assert stores.id not in [x.id for x in eff]   # 组本身不进结果


def test_nested_group_recursive_expand():
    doc = build_doc("d", {"Outer": {"Inner": {"a": "text", "b": "text"}, "c": "text"}})
    i = collect_layer_index(doc)
    outer = i.get("0")
    eff = lm.resolve_effective_logo_layers(i, [outer])
    assert [x.name for x in eff] == ["a", "b", "c"]


def test_group_plus_child_no_duplicate(idx):
    stores = idx.get("1")
    jiuxing = idx.get("1/1")
    eff = lm.resolve_effective_logo_layers(idx, [stores, jiuxing])
    ids = [x.id for x in eff]
    assert ids.count("1/1") == 1                 # 按 id 去重
    assert ids == ["1/0", "1/1", "1/2", "1/3"]


def test_same_name_leaf_not_dedup_by_name():
    # 两个同名叶子在不同组：选择两个组后，同名叶子必须都保留（按 id 去重）
    doc = build_doc("d", {
        "G1": {"电话": "text", "Logo": "text"},
        "G2": {"电话": "text", "Logo": "text"},
    })
    i = collect_layer_index(doc)
    eff = lm.resolve_effective_logo_layers(i, [i.get("0"), i.get("1")])
    names = [x.name for x in eff]
    assert names.count("电话") == 2
    assert names.count("Logo") == 2
    assert len({x.id for x in eff}) == 4


def test_output_order_stable(idx):
    # 先选子再选组：顺序 = 传入顺序展开
    jiuxing = idx.get("1/1")
    stores = idx.get("1")
    eff = lm.resolve_effective_logo_layers(idx, [jiuxing, stores])
    assert [x.id for x in eff] == ["1/1", "1/0", "1/2", "1/3"]
    # 再跑一次：稳定
    eff2 = lm.resolve_effective_logo_layers(idx, [jiuxing, stores])
    assert [x.id for x in eff2] == [x.id for x in eff]


# ===========================================================================
# 2. store mapping（第 7-12 项）
# ===========================================================================
def test_store_mapping_value_must_be_leaf(idx):
    store = idx.get("1/0")
    assert store.is_group is False


def test_store_mapping_group_invalid(idx, all_leaves):
    stores = idx.get("1")
    m = lm.LogoMapping(store_logo_map={"康乐": stores}, brand_logo_refs=[],
                       logo_selection_refs=[stores])
    with pytest.raises(lm.LogoValidationError) as ei:
        lm.validate_logo_mapping(m, effective_leaf_refs=all_leaves)
    assert ei.value.code == "STORE_GROUP"


def test_store_mapping_not_in_effective_invalid(idx, all_leaves):
    # 映射到 Stores 组外的叶子（Brand>七方Logo），但 effective 只含 Stores 组
    qifang = idx.get("0/0")
    m = lm.LogoMapping(store_logo_map={"康乐": qifang}, brand_logo_refs=[],
                       logo_selection_refs=[idx.get("1")])
    eff = lm.resolve_effective_logo_layers(idx, [idx.get("1")])
    with pytest.raises(lm.LogoValidationError) as ei:
        lm.validate_logo_mapping(m, effective_leaf_refs=eff)
    assert ei.value.code == "STORE_NOT_IN_EFFECTIVE"


def test_store_missing_mapping(idx, all_leaves):
    m = lm.LogoMapping(store_logo_map={"康乐": None}, brand_logo_refs=[],
                       logo_selection_refs=[idx.get("1")])
    with pytest.raises(lm.LogoValidationError) as ei:
        lm.validate_logo_mapping(m, effective_leaf_refs=all_leaves)
    assert ei.value.code == "STORE_UNMAPPED"


def test_store_valid_mapping(idx, all_leaves):
    m = lm.LogoMapping(store_logo_map={"康乐": idx.get("1/0")},
                       brand_logo_refs=[idx.get("0/0")],
                       logo_selection_refs=[idx.get("1"), idx.get("0")])
    eff = lm.resolve_effective_logo_layers(idx, [idx.get("1"), idx.get("0")])
    assert lm.validate_logo_mapping(m, effective_leaf_refs=eff) is None


def test_two_stores_share_same_leaf_allowed(idx, all_leaves):
    leaf = idx.get("1/0")
    m = lm.LogoMapping(store_logo_map={"康乐": leaf, "康乐旗舰店": leaf},
                       brand_logo_refs=[],
                       logo_selection_refs=[idx.get("1")])
    eff = lm.resolve_effective_logo_layers(idx, [idx.get("1")])
    assert lm.validate_logo_mapping(m, effective_leaf_refs=eff) is None
    # 可选拒绝
    with pytest.raises(lm.LogoValidationError) as ei:
        lm.validate_logo_mapping(m, effective_leaf_refs=eff,
                                 allow_duplicate_store_targets=False)
    assert ei.value.code == "DUPLICATE_STORE_TARGET"


# ===========================================================================
# 3. brand（第 13-16 项）
# ===========================================================================
def test_brand_must_be_leaf(idx, all_leaves):
    stores = idx.get("1")
    m = lm.LogoMapping(store_logo_map={}, brand_logo_refs=[stores],
                       logo_selection_refs=[stores])
    with pytest.raises(lm.LogoValidationError) as ei:
        lm.validate_logo_mapping(m, effective_leaf_refs=all_leaves)
    assert ei.value.code == "BRAND_GROUP"


def test_brand_store_conflict(idx, all_leaves):
    # 同一 ref 同时映射为 store 与 brand，且该 ref 在 effective 内 -> CONFLICT
    qifang = idx.get("0/0")
    m = lm.LogoMapping(store_logo_map={"康乐": qifang},
                       brand_logo_refs=[qifang],
                       logo_selection_refs=[idx.get("0")])
    eff = lm.resolve_effective_logo_layers(idx, [idx.get("0")])
    with pytest.raises(lm.LogoValidationError) as ei:
        lm.validate_logo_mapping(m, effective_leaf_refs=eff)
    assert ei.value.code == "STORE_BRAND_CONFLICT"


def test_multiple_brand_logos(idx, all_leaves):
    doc = build_doc("d", {"Brand": {"七方Logo": "text", "集团Logo": "text"},
                          "Stores": {"康乐电器": "text"}})
    i = collect_layer_index(doc)
    m = lm.LogoMapping(store_logo_map={"康乐": i.get("1/0")},
                       brand_logo_refs=[i.get("0/0"), i.get("0/1")],
                       logo_selection_refs=[i.get("0"), i.get("1")])
    eff = lm.resolve_effective_logo_layers(i, [i.get("0"), i.get("1")])
    assert lm.validate_logo_mapping(m, effective_leaf_refs=eff) is None


def test_brand_always_visible(idx, all_leaves):
    m = lm.LogoMapping(store_logo_map={"九兴": idx.get("1/1")},
                       brand_logo_refs=[idx.get("0/0")],
                       logo_selection_refs=[idx.get("1"), idx.get("0")])
    plan = lm.prepare_logo_visibility("九兴", m, require_store_mapping=True)
    brand_entries = [(r, v) for r, v in plan if r.name == "七方Logo"]
    assert len(brand_entries) == 1 and brand_entries[0][1] is True


# ===========================================================================
# 4. matching（第 17-23 项）
# ===========================================================================
def test_kanle_to_kangle_dianqi():
    # 康乐 -> 康乐电器（不再被 is_logo_candidate 过滤）：90 分唯一
    # 使用「无康乐家电」的 fixture，避免与歧义用例混淆
    doc = build_doc("d", {"Brand": {"七方Logo": "text"},
                          "Stores": {"康乐电器": "text", "九兴电器": "text"}})
    i = collect_layer_index(doc)
    leaves = [r for r in i.layers if not r.is_group]
    m = lm.match_store_logo("康乐", leaves)
    assert m.status in (lm.AUTO, lm.EXACT)
    assert m.best is not None and m.best.name == "康乐电器"


def test_jiuxing_to_jiuxing_dianqi(all_leaves):
    m = lm.match_store_logo("九兴", all_leaves)
    assert m.status == lm.AUTO and m.best.name == "九兴电器"


def test_exact_preferred(all_leaves):
    m = lm.match_store_logo("康乐电器", all_leaves)
    assert m.status == lm.EXACT and m.score == 100
    assert m.best.name == "康乐电器"


def test_ambiguous_returns_ambiguous(all_leaves):
    # 康乐 -> 康乐电器 / 康乐家电 同分 90 -> AMBIGUOUS，best=None，绝不自动选
    m = lm.match_store_logo("康乐", all_leaves)
    assert m.status == lm.AMBIGUOUS
    assert m.best is None
    assert {r.name for r in m.hits} == {"康乐电器", "康乐家电"}


def test_unrelated_no_match(all_leaves):
    m = lm.match_store_logo("东山", all_leaves)
    assert m.status == lm.NO_MATCH and m.best is None


def test_parent_logo_is_bonus_not_filter(all_leaves):
    # 「logo」出现在父路径/图层名只能是加分项，不能过滤其他候选：
    # 七方Logo 的 display_path 含 logo(+10)、name 含 logo(+5) -> 105 唯一
    m = lm.match_store_logo("七方", all_leaves)
    assert m.status == lm.AUTO
    assert m.best.name == "七方Logo"
    assert m.score == 105
    # 同时证明：不含 logo 的「康乐电器」等候选仍参与了匹配（未被过滤）
    m2 = lm.match_store_logo("康乐", all_leaves)
    assert {r.name for r in m2.hits} == {"康乐电器", "康乐家电"}


def test_same_name_candidates_not_dedup_by_name():
    # 两个同名「Logo」叶子在不同组：候选必须都保留（不按 name 去重）
    doc = build_doc("d", {"G1": {"Logo": "text"}, "G2": {"Logo": "text"}})
    i = collect_layer_index(doc)
    leaves = [r for r in i.layers if not r.is_group]
    assert len(leaves) == 2
    m = lm.match_store_logo("Logo", leaves)
    assert len(m.hits) == 2
    assert {r.id for r in m.hits} == {"0/0", "1/0"}
    assert m.status == lm.AMBIGUOUS


# ===========================================================================
# 5. visibility plan（第 24-28 项）
# ===========================================================================
def test_target_store_true_others_false(idx):
    m = lm.LogoMapping(store_logo_map={"康乐": idx.get("1/0"),
                                       "九兴": idx.get("1/1"),
                                       "诚信": idx.get("1/2"),
                                       "康乐家电": idx.get("1/3")},
                       brand_logo_refs=[idx.get("0/0")],
                       logo_selection_refs=[idx.get("1"), idx.get("0")])
    plan = lm.prepare_logo_visibility("九兴", m, require_store_mapping=True)
    visible = {r.name: v for r, v in plan}
    assert visible["九兴电器"] is True
    assert visible["康乐电器"] is False
    assert visible["诚信电器"] is False
    assert visible["康乐家电"] is False


def test_other_stores_false(idx):
    m = lm.LogoMapping(store_logo_map={"康乐": idx.get("1/0"),
                                       "九兴": idx.get("1/1")},
                       brand_logo_refs=[],
                       logo_selection_refs=[idx.get("1")])
    plan = lm.prepare_logo_visibility("康乐", m, require_store_mapping=True)
    visible = {r.name: v for r, v in plan}
    assert visible["康乐电器"] is True
    assert visible["九兴电器"] is False


def test_brands_true(idx):
    m = lm.LogoMapping(store_logo_map={"九兴": idx.get("1/1")},
                       brand_logo_refs=[idx.get("0/0")],
                       logo_selection_refs=[idx.get("1"), idx.get("0")])
    plan = lm.prepare_logo_visibility("九兴", m, require_store_mapping=True)
    visible = {r.name: v for r, v in plan}
    assert visible["七方Logo"] is True


def test_unmapped_store_error(idx):
    m = lm.LogoMapping(store_logo_map={"九兴": idx.get("1/1")},
                       brand_logo_refs=[], logo_selection_refs=[idx.get("1")])
    with pytest.raises(lm.LogoVisibilityError) as ei:
        lm.prepare_logo_visibility("康乐", m, require_store_mapping=True)
    assert "没有有效 Logo 映射" in str(ei.value)


def test_conflict_error(idx):
    # 同一 ref 在 store_map 与 brand 中：先 prepare 时按 store 用（不冲突检测），
    # 冲突检测属于 validate 阶段；这里验证运行时若配置已被污染，verify 能发现
    qifang = idx.get("0/0")
    m = lm.LogoMapping(store_logo_map={"康乐": qifang},
                       brand_logo_refs=[qifang],
                       logo_selection_refs=[idx.get("0")])
    with pytest.raises(lm.LogoValidationError) as ei:
        lm.validate_logo_mapping(m, effective_leaf_refs=[qifang])
    assert ei.value.code == "STORE_BRAND_CONFLICT"


def test_verify_logo_visibility_mismatch():
    ref = LayerRef(id="1/0", name="康乐电器", display_path="Stores > 康乐电器",
                   index_path=(1, 0), is_text=True)
    plan = [(ref, True)]
    # 一致：通过
    assert lm.verify_logo_visibility({"1/0": True}, plan) is True
    # 不一致：抛错
    with pytest.raises(lm.LogoVisibilityError) as ei:
        lm.verify_logo_visibility({"1/0": False}, plan)
    assert ei.value.code == "READBACK_MISMATCH"


def test_unmapped_candidate_leaf_hidden_when_effective_given(idx):
    # effective 中包含「未被任何门店映射」的候选叶子（如康乐家电）：
    # 当前行必须把它隐藏，绝不保留模板原状态（对应 Case 1「其他 store hidden」）。
    m = lm.LogoMapping(store_logo_map={"九兴": idx.get("1/1")},
                       brand_logo_refs=[idx.get("0/0")],
                       logo_selection_refs=[idx.get("1"), idx.get("0")])
    eff = lm.resolve_effective_logo_layers(idx, [idx.get("1"), idx.get("0")])
    plan = lm.prepare_logo_visibility("九兴", m, require_store_mapping=True,
                                      effective_leaf_refs=eff)
    visible = {r.name: v for r, v in plan}
    # 九兴电器 True；康乐电器/诚信电器/康乐家电（未映射候选）全部 False；品牌 True
    assert visible["九兴电器"] is True
    assert visible["康乐电器"] is False
    assert visible["诚信电器"] is False
    assert visible["康乐家电"] is False
    assert visible["七方Logo"] is True


# ===========================================================================
# 6. old config（第 29-33 项）
# ===========================================================================
def test_old_unique_name_migrated(idx):
    mig, reps = lm.migrate_old_config(["Stores"], {"康乐": "康乐电器"}, idx)
    assert mig.logo_selection_refs[0].display_path == "Stores"
    assert mig.store_logo_map["康乐"].id == "1/0"
    store_reports = [r for r in reps if "store" in r]
    assert store_reports[0]["status"] == lm.MIGRATED


def test_old_duplicate_name_ambiguous(idx):
    # 门店名「康乐」在 PSD 有两个同名候选？构造：两个组里都有「康乐电器」
    doc = build_doc("d", {
        "Stores": {"康乐电器": "text"},
        "Stores2": {"康乐电器": "text"},
    })
    i = collect_layer_index(doc)
    mig, reps = lm.migrate_old_config(["Stores"], {"康乐": "康乐电器"}, i)
    assert mig.store_logo_map["康乐"] is None          # 歧义 -> 不自动选
    store_reports = [r for r in reps if "store" in r]
    assert store_reports[0]["status"] == lm.AMBIGUOUS


def test_old_missing_name_missing(idx):
    mig, reps = lm.migrate_old_config(["不存在组"], {"康乐": "不存在层"}, idx)
    assert mig.logo_selection_refs == []
    assert mig.store_logo_map["康乐"] is None
    statuses = {r["status"] for r in reps}
    assert lm.MISSING in statuses


def test_old_group_selection_migrated_as_group_ref(idx):
    # 旧 selection 是组名（"Stores"）-> 迁移为 selected group ref（可再展开）
    mig, reps = lm.migrate_old_config(["Stores"], {}, idx)
    assert len(mig.logo_selection_refs) == 1
    sel = mig.logo_selection_refs[0]
    assert sel.is_group is True
    eff = lm.resolve_effective_logo_layers(idx, [sel])
    assert len(eff) == 4


def test_old_store_map_no_auto_pick_ambiguous(idx):
    # 旧 store map 指向歧义同名 -> 不自动选第一个（保持未映射）
    doc = build_doc("d", {
        "G1": {"Logo": "text"},
        "G2": {"Logo": "text"},
    })
    i = collect_layer_index(doc)
    mig, reps = lm.migrate_old_config(["G1", "G2"], {"门店A": "Logo"}, i)
    assert mig.store_logo_map["门店A"] is None
    store_reports = [r for r in reps if "store" in r]
    assert store_reports[0]["status"] == lm.AMBIGUOUS


# ===========================================================================
# 附加：品牌建议（第二十一节）
# ===========================================================================
def test_suggest_brand_logos(idx):
    store_map = {"康乐": idx.get("1/0"), "九兴": idx.get("1/1")}
    leaves = [r for r in idx.layers if not r.is_group]
    brands = lm.suggest_brand_logos(leaves, store_map)
    assert [r.name for r in brands] == ["七方Logo"]   # 只有未映射且含 logo 的叶子


def test_suggest_brand_excludes_mapped(idx):
    # 七方Logo 被映射为门店后，不再被推荐为品牌
    store_map = {"七方": idx.get("0/0")}
    leaves = [r for r in idx.layers if not r.is_group]
    brands = lm.suggest_brand_logos(leaves, store_map)
    assert brands == []


# ===========================================================================
# Stage 3 补充：首次加载推荐 + 品牌人工指定（对应补充验收第四节 8 项）
# ===========================================================================
# 结构：普通素材（无 logo 关键字）
#   └─ 康乐电器（无 logo 关键字，父路径也无 logo 关键字）
# 门店 Excel 名 = 「康乐」（非「康乐电器」）
# ---------------------------------------------------------------------------
def make_plain_doc(leaf_name="康乐电器", group_name="普通素材"):
    """无任何 logo 关键字的 PSD：普通组 > 门店名叶子。"""
    return build_doc("d", {group_name: {leaf_name: "text"}})


def _plain_idx(leaf_name="康乐电器", group_name="普通素材"):
    return collect_layer_index(make_plain_doc(leaf_name, group_name))


# --- 1. first load: 康乐 -> 普通素材 > 康乐电器（无 logo keyword）仍 AUTO ---
def test_first_load_recommends_kangle_dianqi_without_logo_keyword():
    i = _plain_idx()
    stores = ["康乐"]
    recommend = lm.recommend_logo_selection(i.layers, stores)
    # 推荐勾选必须包含康乐电器（不能为空）
    assert any(r.name == "康乐电器" for r in recommend)
    eff = lm.resolve_effective_logo_layers(i, recommend)
    assert any(r.name == "康乐电器" for r in eff)
    mr = lm.match_store_logo("康乐", eff)
    assert mr.status == lm.AUTO
    assert mr.best is not None and mr.best.name == "康乐电器"


# --- 2. first load: 九兴 -> 普通组 > 九兴电器（无 logo keyword）仍 AUTO ---
def test_first_load_recommends_jiuxing_dianqi_without_logo_keyword():
    i = _plain_idx(leaf_name="九兴电器", group_name="普通组")
    recommend = lm.recommend_logo_selection(i.layers, ["九兴"])
    assert any(r.name == "九兴电器" for r in recommend)
    eff = lm.resolve_effective_logo_layers(i, recommend)
    mr = lm.match_store_logo("九兴", eff)
    assert mr.status == lm.AUTO
    assert mr.best is not None and mr.best.name == "九兴电器"


# --- 3. brand manual assignment: selected leaf 可人工成为 brand ---
def test_brand_manual_assignment():
    i = _plain_idx()
    leaf = [r for r in i.layers if r.name == "康乐电器"][0]
    m = lm.LogoMapping(store_logo_map={}, brand_logo_refs=[leaf],
                       logo_selection_refs=[leaf])
    # 校验通过（brand 叶子合法）
    lm.validate_logo_mapping(m, effective_leaf_refs=[leaf])
    plan = lm.prepare_logo_visibility("任何门店", m, effective_leaf_refs=[leaf])
    # brand 叶子始终 visible（即使门店无映射，require_store_mapping 默认 False 时不抛）
    visible = [r for r, v in plan if v]
    assert any(r.id == leaf.id for r in visible)


# --- 4. brand manual removal: 取消后不再属于 brand ---
def test_brand_manual_removal():
    i = _plain_idx()
    leaf = [r for r in i.layers if r.name == "康乐电器"][0]
    m = lm.LogoMapping(store_logo_map={}, brand_logo_refs=[],   # 取消 = 空
                       logo_selection_refs=[leaf])
    assert m.brand_logo_refs == []
    plan = lm.prepare_logo_visibility("X", m, effective_leaf_refs=[leaf])
    # 无 brand、无 store 映射：leaf 作为未分类候选被隐藏（selected != brand 证明）
    visible = [r for r, v in plan if v]
    assert all(r.id != leaf.id for r in visible)


# --- 5. brand/store conflict: 同 LayerRef 不能同时 brand + store ---
# 注意：与上面旧版 test_brand_store_conflict（182 行，Brand/Stores fixture）断言等价，
# 但使用独立的 plain fixture（普通素材>康乐电器），避免同名 def 覆盖旧测试导致收集数减少。
def test_brand_store_conflict_plain():
    i = _plain_idx()
    leaf = [r for r in i.layers if r.name == "康乐电器"][0]
    m = lm.LogoMapping(store_logo_map={"康乐": leaf}, brand_logo_refs=[leaf],
                       logo_selection_refs=[leaf])
    with pytest.raises(lm.LogoValidationError) as ei:
        lm.validate_logo_mapping(m, effective_leaf_refs=[leaf])
    assert ei.value.code == "STORE_BRAND_CONFLICT"


# --- 6. brand config roundtrip: serialize -> save-shaped -> reload，LayerRef 保持 ---
def test_brand_config_roundtrip():
    from core.layer_index import serialize_ref, ref_from_config
    i = _plain_idx()
    leaf = [r for r in i.layers if r.name == "康乐电器"][0]
    saved = [serialize_ref(leaf)]          # 保存形态（dict）
    # 重新加载：ref_from_config -> rebind 到新 index（同一 index 模拟重启后重载）
    i2 = _plain_idx()
    m = lm.LogoMapping(store_logo_map={}, brand_logo_refs=[], logo_selection_refs=[])
    raw = [ref_from_config(v) for v in saved]
    m2 = lm.LogoMapping(store_logo_map={},
                        brand_logo_refs=[r for r in raw if r is not None],
                        logo_selection_refs=[])
    # LayerRef 身份保持（id 一致）
    assert [r.id for r in m2.brand_logo_refs] == [leaf.id]
    assert [r.name for r in m2.brand_logo_refs] == ["康乐电器"]


# --- 7. selected-but-unclassified leaf: runtime 隐藏，selected != brand ---
def test_selected_but_unclassified_hidden():
    i = _plain_idx(leaf_name="九兴电器", group_name="普通组")
    jiuxing = [r for r in i.layers if r.name == "九兴电器"][0]
    # 另造一个「被选中但未被映射未品牌」的叶子
    doc = build_doc("d", {"普通组": {"九兴电器": "text", "九兴备用": "text"}})
    i2 = collect_layer_index(doc)
    nine = i2.get("0/0")
    spare = i2.get("0/1")
    m = lm.LogoMapping(store_logo_map={"九兴": nine}, brand_logo_refs=[],
                       logo_selection_refs=[nine, spare])   # spare 被选中但未分类
    eff = [nine, spare]
    plan = lm.prepare_logo_visibility("九兴", m, require_store_mapping=True,
                                      effective_leaf_refs=eff)
    vis = {r.id: v for r, v in plan}
    assert vis[nine.id] is True      # store target 显示
    assert vis[spare.id] is False    # 未分类候选隐藏（证明 selected != brand）


# --- 8. 8081 类似无 logo keyword brand: 人工标记后 runtime 始终 visible ---
def test_no_logo_keyword_brand_always_visible():
    # 模拟 8081：无 logo 关键字的叶子（电话 拷贝 / 圣大）
    doc = build_doc("d", {
        "电话信息": {"电话 拷贝": "text", "姓 名 拷贝": "text"},
        "背景": {"圣大": "text", "销售顾问 拷贝": "text"},
    })
    i = collect_layer_index(doc)
    shengda = [r for r in i.layers if r.name == "圣大"][0]
    tel = [r for r in i.layers if r.name == "电话 拷贝"][0]
    # 人工把圣大标记为 brand（无 logo keyword）
    m = lm.LogoMapping(store_logo_map={"圣大家电": tel}, brand_logo_refs=[shengda],
                       logo_selection_refs=[tel, shengda])
    eff = [tel, shengda]
    lm.validate_logo_mapping(m, effective_leaf_refs=eff)
    plan = lm.prepare_logo_visibility("圣大家电", m, require_store_mapping=True,
                                      effective_leaf_refs=eff)
    vis = {r.id: v for r, v in plan}
    assert vis[tel.id] is True       # store logo 显示
    assert vis[shengda.id] is True   # 人工 brand 始终显示（无 logo keyword 也成立）
