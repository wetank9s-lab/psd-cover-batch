# -*- coding: utf-8 -*-
"""
test_layer_index.py —— Stage 2 LayerRef / LayerIndex 纯 Python 测试（不依赖 Photoshop）。

fake 模型刻意模拟 Photoshop COM 语义：
  - container.Layers 是 1-based 索引（Layers[i], i 从 1 起），
    collect 时记录 index_path 为 0-based（i-1），
    访问 COM 时再 +1 —— 与 core/layer_index.py 的约定完全一致。
"""
import pytest

from core.layer_index import (
    LayerRef,
    LayerIndex,
    LayerResolutionError,
    collect_layer_index,
    resolve_layer,
    rebind_layer_reference,
    serialize_ref,
    ref_from_config,
    unique_display_labels,
    VALID, MIGRATED, AMBIGUOUS, MISSING,
)


# ---------------------------------------------------------------------------
# fake 模型（1-based Layers，TextItem 仅文字层有）
# ---------------------------------------------------------------------------
class FakeTextItem:
    def __init__(self, contents=""):
        self.Contents = contents


class FakeLayer:
    _seq = 0

    def __init__(self, name, children=None, text=False):
        FakeLayer._seq += 1
        self._id = FakeLayer._seq
        self.Name = name
        self.Visible = True
        self._children = list(children or [])
        self._text = FakeTextItem(name + "-text") if text else None
        self._doc = None
        for c in self._children:
            c._doc = None

    @property
    def TextItem(self):
        if self._text is None:
            raise AttributeError("not a text layer")
        return self._text

    @property
    def Layers(self):
        return _LayerCollection(self._children)

    @property
    def Parent(self):
        # 简化：fake 里不维护真实父子链（真实 COM 才有 Parent）
        raise AttributeError("Parent not tracked in fake")

    def __repr__(self):
        return f"<FakeLayer {self.Name!r}>"


class _LayerCollection:
    """模拟 COM 的 Layers 集合：1-based 索引，支持迭代、Count。"""

    def __init__(self, children):
        self._children = list(children)

    def __len__(self):
        return len(self._children)

    @property
    def Count(self):
        return len(self._children)

    def __getitem__(self, i):
        # COM 1-based：i 从 1 起；支持 slice 仅用于测试便利
        if isinstance(i, slice):
            return self._children[i]
        if i < 1 or i > len(self._children):
            raise IndexError(f"index {i} out of range (1..{len(self._children)})")
        return self._children[i - 1]

    def __iter__(self):
        return iter(self._children)

    def add(self, layer):
        self._children.append(layer)
        return layer


class FakeDoc:
    """模拟 Document：只有 Layers 根集合 + Name。"""

    def __init__(self, name, root_layers):
        self.Name = name
        self._root = _LayerCollection(root_layers)

    @property
    def Layers(self):
        return self._root


# ---------------------------------------------------------------------------
# 便捷构造：从嵌套 dict 建文档
# ---------------------------------------------------------------------------
def build_doc(name, tree):
    """tree 可用 dict（key 必须唯一）或 list[(name, sub)]（允许同名，如 [("电话","text"),("电话","text")]）。
    sub 为 "text" 表示文字层，dict/list 表示组，None 表示普通叶子。"""
    def mk(node):
        layers = []
        items = node.items() if isinstance(node, dict) else node
        for entry in items:
            nm, sub = entry
            if sub == "text":
                layers.append(FakeLayer(nm, text=True))
            elif isinstance(sub, (dict, list)):
                layers.append(FakeLayer(nm, children=mk(sub)))
            else:
                layers.append(FakeLayer(nm))
        return layers
    return FakeDoc(name, mk(tree))


def paths_of(index):
    return [(r.id, r.name, r.display_path) for r in index.layers]


# ---------------------------------------------------------------------------
# 1. 单层 PSD
# ---------------------------------------------------------------------------
def test_single_layer():
    doc = build_doc("d", {"电话": "text"})
    idx = collect_layer_index(doc)
    assert len(idx) == 1
    r = idx.layers[0]
    assert r.id == "0"
    assert r.name == "电话"
    assert r.display_path == "电话"
    assert r.index_path == (0,)
    assert r.is_text is True
    assert r.is_group is False
    # resolve
    layer = idx.resolve(doc, r)
    assert layer is not None and layer.Name == "电话"


# ---------------------------------------------------------------------------
# 2. 嵌套 group
# ---------------------------------------------------------------------------
def test_nested_group():
    doc = build_doc("d", {"Header": {"电话": "text"}})
    idx = collect_layer_index(doc)
    assert len(idx) == 2
    header = idx.get("0")
    phone = idx.get("0/0")
    assert header.is_group is True and header.name == "Header"
    assert phone.is_group is False and phone.is_text is True
    assert phone.display_path == "Header > 电话"
    assert phone.index_path == (0, 0)


# ---------------------------------------------------------------------------
# 3. 两个不同 group 中同名 layer
# ---------------------------------------------------------------------------
def test_same_name_in_two_groups():
    doc = build_doc("d", {"Header": {"电话": "text"}, "Footer": {"电话": "text"}})
    idx = collect_layer_index(doc)
    assert len(idx) == 4
    phones = idx.find_by_name("电话")
    assert len(phones) == 2
    ids = {p.id for p in phones}
    assert ids == {"0/0", "1/0"}
    assert phones[0].display_path != phones[1].display_path
    # 各自 resolve 到正确对象
    assert idx.resolve(doc, phones[0]) is doc.Layers[1].Layers[1]
    assert idx.resolve(doc, phones[1]) is doc.Layers[2].Layers[1]


# ---------------------------------------------------------------------------
# 4. 同一个 group 中两个同名 layer
# ---------------------------------------------------------------------------
def test_same_name_within_same_group():
    doc = build_doc("d", [("Root", [("电话", "text"), ("电话", "text")])])
    idx = collect_layer_index(doc)
    assert len(idx) == 3
    phones = idx.find_by_name("电话")
    assert len(phones) == 2
    assert {p.id for p in phones} == {"0/0", "0/1"}
    # 名字相同但 id 不同
    assert phones[0].id != phones[1].id
    # 独立 resolve
    assert idx.resolve(doc, phones[0]) is doc.Layers[1].Layers[1]
    assert idx.resolve(doc, phones[1]) is doc.Layers[1].Layers[2]


# ---------------------------------------------------------------------------
# 5. 同名 group
# ---------------------------------------------------------------------------
def test_same_name_groups():
    doc = build_doc("d", [("G", {"电话": "text"}), ("G", {"Logo": {}})])
    idx = collect_layer_index(doc)
    groups = idx.find_by_name("G")
    assert len(groups) == 2
    assert {g.id for g in groups} == {"0", "1"}
    assert groups[0].is_group and groups[1].is_group
    # 两个组内的子层 id 不同
    assert idx.get("0/0").name == "电话"
    assert idx.get("1/0").name == "Logo"


# ---------------------------------------------------------------------------
# 6. display_path 正确
# ---------------------------------------------------------------------------
def test_display_path():
    doc = build_doc("d", {"文字": {"Header": {"电话": "text"}}})
    idx = collect_layer_index(doc)
    phone = idx.get("0/0/0")
    assert phone.display_path == "文字 > Header > 电话"


# ---------------------------------------------------------------------------
# 7. index_path 唯一
# ---------------------------------------------------------------------------
def test_index_path_unique():
    doc = build_doc("d", {"A": {"x": {}, "y": {}}, "B": {"x": {}}})
    idx = collect_layer_index(doc)
    ips = [r.index_path for r in idx.layers]
    assert len(ips) == len(set(ips))  # 无重复


# ---------------------------------------------------------------------------
# 8. by_name 返回 list 而不是覆盖
# ---------------------------------------------------------------------------
def test_by_name_is_list():
    doc = build_doc("d", {"电话": "text", "Header": {"电话": "text"}})
    idx = collect_layer_index(doc)
    assert isinstance(idx.by_name["电话"], list)
    assert len(idx.by_name["电话"]) == 2
    assert idx.find_by_name("电话") == idx.by_name["电话"]


# ---------------------------------------------------------------------------
# 9. resolve 正确 layer
# ---------------------------------------------------------------------------
def test_resolve_correct_layer():
    doc = build_doc("d", {"电话": "text", "Logo": {}})
    idx = collect_layer_index(doc)
    ref0 = idx.get("0")
    ref1 = idx.get("1")
    assert idx.resolve(doc, ref0) is doc.Layers[1]
    assert idx.resolve(doc, ref1) is doc.Layers[2]


# ---------------------------------------------------------------------------
# 10. index_path 不存在返回明确错误
# ---------------------------------------------------------------------------
def test_resolve_missing_index_path():
    doc = build_doc("d", {"电话": "text"})
    idx = collect_layer_index(doc)
    ref = LayerRef(id="9", name="电话", display_path="电话", index_path=(9,))
    with pytest.raises(LayerResolutionError) as ei:
        idx.resolve(doc, ref)
    assert "越界" in str(ei.value) or "index_path" in str(ei.value)


# ---------------------------------------------------------------------------
# 11. index_path 位置存在但 name 已变化 -> stale
# ---------------------------------------------------------------------------
def test_resolve_stale_name():
    doc = build_doc("d", {"电话": "text"})
    idx = collect_layer_index(doc)
    # 引用位置 (0) 现在叫 "电话"，但 ref 记录 name="姓名"（结构已变）
    stale = LayerRef(id="0", name="姓名", display_path="姓名", index_path=(0,))
    with pytest.raises(LayerResolutionError) as ei:
        idx.resolve(doc, stale)
    assert "不匹配" in str(ei.value)


# ---------------------------------------------------------------------------
# 12. unique old-name migration（MIGRATED）
# ---------------------------------------------------------------------------
def test_rebind_unique_old_name():
    doc = build_doc("d", {"Header": {"电话": "text"}})
    idx = collect_layer_index(doc)
    status, ref = rebind_layer_reference(idx, "电话")  # 旧配置：纯 name
    assert status == MIGRATED
    assert ref is not None
    assert ref.id == "0/0"
    assert ref.name == "电话"


# ---------------------------------------------------------------------------
# 13. duplicate old-name -> ambiguous
# ---------------------------------------------------------------------------
def test_rebind_ambiguous_old_name():
    doc = build_doc("d", {"Header": {"电话": "text"}, "Footer": {"电话": "text"}})
    idx = collect_layer_index(doc)
    status, ref = rebind_layer_reference(idx, "电话")
    assert status == AMBIGUOUS
    assert ref is None


# ---------------------------------------------------------------------------
# 14. missing old-name -> missing
# ---------------------------------------------------------------------------
def test_rebind_missing_old_name():
    doc = build_doc("d", {"电话": "text"})
    idx = collect_layer_index(doc)
    status, ref = rebind_layer_reference(idx, "不存在的层")
    assert status == MISSING
    assert ref is None


# ---------------------------------------------------------------------------
# 15. group/leaf is_group 正确
# ---------------------------------------------------------------------------
def test_is_group_flag():
    doc = build_doc("d", {"G": {"leaf": {}}, "plain": {}})
    idx = collect_layer_index(doc)
    assert idx.get("0").is_group is True
    assert idx.get("0/0").is_group is False
    assert idx.get("1").is_group is False


# ---------------------------------------------------------------------------
# 16. text layer is_text 正确
# ---------------------------------------------------------------------------
def test_is_text_flag():
    doc = build_doc("d", {"姓名": "text", "Logo": {}})
    idx = collect_layer_index(doc)
    assert idx.get("0").is_text is True
    assert idx.get("1").is_text is False


# ---------------------------------------------------------------------------
# 17. 路径中含中文
# ---------------------------------------------------------------------------
def test_chinese_paths():
    doc = build_doc("d", {"文字层组": {"姓名": "text"}})
    idx = collect_layer_index(doc)
    r = idx.get("0/0")
    assert r.display_path == "文字层组 > 姓名"
    assert r.name == "姓名"
    assert idx.resolve(doc, r) is doc.Layers[1].Layers[1]


# ---------------------------------------------------------------------------
# 18. 路径中含相同父组名
# ---------------------------------------------------------------------------
def test_same_parent_group_name():
    doc = build_doc("d", {"G": {"G": {"电话": "text"}}})
    idx = collect_layer_index(doc)
    # 两层都叫 G：display_path 会重复，但 id 唯一
    inner = idx.get("0/0/0")
    assert inner.display_path == "G > G > 电话"
    assert inner.index_path == (0, 0, 0)
    assert idx.resolve(doc, inner) is doc.Layers[1].Layers[1].Layers[1]


# ---------------------------------------------------------------------------
# 19. duplicate Document 可以通过同 LayerRef resolve 到对应 duplicate layer
# ---------------------------------------------------------------------------
def test_resolve_on_duplicate_doc():
    doc = build_doc("d", {"Header": {"电话": "text"}, "Footer": {"电话": "text"}})
    idx = collect_layer_index(doc)
    # 模拟 Duplicate：新建同构文档（深度复制图层树）
    def clone_tree(layers):
        out = []
        for i in range(1, layers.Count + 1):
            L = layers[i]
            if L._text is not None:
                out.append(FakeLayer(L.Name, text=True))
            elif L.Layers.Count > 0:
                out.append(FakeLayer(L.Name, children=clone_tree(L.Layers)))
            else:
                out.append(FakeLayer(L.Name))
        return out
    dup = FakeDoc(doc.Name + "-copy", clone_tree(doc.Layers))
    # 同一 LayerRef 在 duplicate 上 resolve 到对应位置
    phones = idx.find_by_name("电话")
    h = phones[0]
    assert h.display_path == "Header > 电话"
    resolved_in_dup = idx.resolve(dup, h)
    assert resolved_in_dup is dup.Layers[1].Layers[1]
    assert resolved_in_dup is not doc.Layers[1].Layers[1]  # 不是 template 的对象


# ---------------------------------------------------------------------------
# 20. 不允许 name fallback 自动选第一个
# ---------------------------------------------------------------------------
def test_no_name_fallback():
    doc = build_doc("d", {"Header": {"电话": "text"}, "Footer": {"电话": "text"}})
    idx = collect_layer_index(doc)
    # index_path 失效（(5,5) 不存在）时 resolve 必须抛错，绝不能找 name=电话 第一个
    bad = LayerRef(id="5/5", name="电话", display_path="x > 电话", index_path=(5, 5))
    with pytest.raises(LayerResolutionError):
        idx.resolve(doc, bad)


# ---------------------------------------------------------------------------
# 附加：序列化 round-trip / unique labels / find_matching
# ---------------------------------------------------------------------------
def test_serialize_roundtrip():
    doc = build_doc("d", {"Header": {"电话": "text"}})
    idx = collect_layer_index(doc)
    ref = idx.get("0/0")
    d = serialize_ref(ref)
    assert d["layer_id"] == "0/0"
    assert d["name"] == "电话"
    assert d["display_path"] == "Header > 电话"
    back = ref_from_config(d)
    assert back.id == ref.id and back.index_path == ref.index_path
    assert back.name == ref.name


def test_ref_from_config_old_string():
    ref = ref_from_config("电话")
    assert ref is not None
    assert ref.id == ""          # 未绑定
    assert ref.index_path == ()
    assert ref.name == "电话"


def test_unique_labels_same_display():
    doc = build_doc("d", [("G", [("电话", "text"), ("电话", "text")])])
    idx = collect_layer_index(doc)
    labels = idx.labels()
    # 两个同名电话层 display_path 都是 "G > 电话"，必须附加 id 肉眼可区分
    phones = [labels[r.id] for r in idx.find_by_name("电话")]
    assert len(set(phones)) == 2
    assert all("id=" in v for v in phones)


def test_find_matching_ignores_spaces():
    doc = build_doc("d", {"Header": {"姓 名": "text"}})
    idx = collect_layer_index(doc)
    m = idx.find_matching("姓名")   # 忽略空格
    assert len(m) == 1
    assert m[0].id == "0/0"
