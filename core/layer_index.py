# -*- coding: utf-8 -*-
"""
layer_index.py —— PSD 图层唯一身份层（Stage 2）。

核心目标：
  两个图层即使名称完全相同，也必须能被独立识别、独立选择、独立保存配置、
  独立重新定位。

设计决策（全项目固定，测试锁定）：
  - index_path 使用 **0-based tuple**。例如 (0,) 表示文档第 1 个图层，
    (0, 2) 表示文档第 1 个图层的第 3 个子图层。
    只有访问 Photoshop COM collection 时才转换成 COM 的 1-based（下标 +1）。
  - LayerRef.id 是 index_path 的 "/" 连接字符串（如 "0/2/1"），
    作为配置序列化主键（唯一、可排序、可 JSON 序列化）。
  - layer.Name 只用于：搜索 / 自动推荐 / 旧配置迁移 / resolve 时校验。
    绝不用于 identity / 去重 / 覆盖。
  - by_name 永远是 dict[str, list[LayerRef]]，绝不单值覆盖（否则同名再次覆盖）。
  - resolve 失败绝不 fallback 到「第一个同名层」——宁可抛 LayerResolutionError，
    也不静默改错对象。
"""

from dataclasses import dataclass, field

# ---------------------------------------------------------------------------
# 常量：rebind / 迁移状态
# ---------------------------------------------------------------------------
VALID = "VALID"          # index_path 存在且 name 特征一致
MIGRATED = "MIGRATED"    # index_path 失效，但 name 全局唯一，返回唯一候选
AMBIGUOUS = "AMBIGUOUS"  # name 出现多个，不自动选
MISSING = "MISSING"      # 完全找不到

# ---------------------------------------------------------------------------
# 数据模型
# ---------------------------------------------------------------------------
@dataclass(frozen=True)
class LayerRef:
    """一个 PSD 图层的可序列化身份描述（不含任何 COM 对象）。"""
    id: str                    # index_path 的 "/" 连接，如 "0/2/1"；空串表示纯 name（未绑定）
    name: str                  # 图层名（只用于展示/搜索/校验，不承担 identity）
    display_path: str          # 从根到该层的完整路径，如 "文字 > Header > 电话"
    index_path: tuple = ()     # 0-based tuple，如 (0, 2, 1)
    is_group: bool = False     # 是否为图层组（LayerSet）
    is_text: bool = False      # 是否为文字图层（有 TextItem）

    def __post_init__(self):
        if not self.id:
            object.__setattr__(self, "id", "/".join(str(x) for x in self.index_path) or "")


class LayerResolutionError(Exception):
    """LayerRef 无法精确定位（index_path 失效 / 位置 name 不匹配 / 中间层非组）。"""

    def __init__(self, message, *, layer_id=None, name=None, index_path=None):
        super().__init__(message)
        self.layer_id = layer_id
        self.name = name
        self.index_path = index_path


# ---------------------------------------------------------------------------
# COM / fake 统一的图层访问辅助（Layers 集合 1-based，记录 0-based）
# ---------------------------------------------------------------------------
def _layers_of(container):
    """返回容器（Document / LayerSet）的直接子图层列表。

    遍历用 COM 1-based 下标（container.Layers[i], i 从 1 起），
    但 collect 时记录的 index_path 是 0-based（i-1）。
    """
    try:
        n = container.Layers.Count
    except Exception:
        return []
    out = []
    for i in range(1, n + 1):
        try:
            out.append(container.Layers[i])
        except Exception:
            continue
    return out


def _is_group(layer):
    """图层是否为组（LayerSet）：是否有 >0 个子图层。"""
    try:
        return layer.Layers.Count > 0
    except Exception:
        return False


def _is_text(layer):
    """图层是否为文字图层：能否取得 TextItem（COM 非文字层访问会抛异常）。"""
    try:
        t = getattr(layer, "TextItem", None)
        return t is not None
    except Exception:
        return False


def _name_matches(a, b):
    """名字一致性校验（resolve 用）：忽略空格 + 小写（延续工具历史「忽略空格差异」行为）。"""
    na = (a or "").replace(" ", "").lower()
    nb = (b or "").replace(" ", "").lower()
    return na == nb and na != ""


def _id_of(path):
    return "/".join(str(x) for x in path)


# ---------------------------------------------------------------------------
# 收集
# ---------------------------------------------------------------------------
def collect_layer_index(document, max_depth=64):
    """遍历 document（COM Document 或 fake），返回 LayerIndex（不丢任何同名图层）。"""
    refs = []

    def walk(container, path, name_path):
        if len(path) >= max_depth:
            return
        for i, layer in enumerate(_layers_of(container), start=1):
            idx = i - 1                      # COM 1-based -> 内部 0-based
            cur_path = path + (idx,)
            cur_names = name_path + (layer.Name,)
            is_group = _is_group(layer)
            is_text = _is_text(layer)
            refs.append(LayerRef(
                id=_id_of(cur_path),
                name=layer.Name,
                display_path=" > ".join(cur_names),
                index_path=cur_path,
                is_group=is_group,
                is_text=is_text,
            ))
            if is_group:
                walk(layer, cur_path, cur_names)

    walk(document, (), ())
    return LayerIndex(refs)


# ---------------------------------------------------------------------------
# 精确定位
# ---------------------------------------------------------------------------
def resolve_layer(document, layer_ref):
    """按 index_path 在 document 中精确定位 COM Layer。

    - 按 0-based index_path 逐层进入（访问 COM 时下标 +1）；
    - 最后一步校验该位置 layer.Name 与 LayerRef.name 一致；
    - 任一步骤失败（越界 / 中间层非组 / name 不匹配）抛 LayerResolutionError；
    - **绝不** fallback 到 find_layer(name) 取第一个同名层。
    """
    if layer_ref is None:
        raise LayerResolutionError("layer_ref 为 None", name=None)
    path = layer_ref.index_path
    if not path:
        raise LayerResolutionError(
            f"LayerRef 没有 index_path（纯 name 引用 {layer_ref.name!r}，需先 rebind）",
            layer_id=layer_ref.id, name=layer_ref.name, index_path=path)

    container = document
    for pos, idx in enumerate(path):
        try:
            layer = container.Layers[idx + 1]   # 0-based -> COM 1-based
        except Exception as exc:
            raise LayerResolutionError(
                f"index_path 第 {pos + 1} 层索引 {idx} 越界（共 "
                f"{_count_layers(container)} 个子层）",
                layer_id=layer_ref.id, name=layer_ref.name, index_path=path) from exc
        if pos == len(path) - 1:
            # 最后一层：校验 name
            if not _name_matches(layer.Name, layer_ref.name):
                raise LayerResolutionError(
                    f"index_path {path} 处图层名不匹配：期望 {layer_ref.name!r}，"
                    f"实际 {layer.Name!r}（PSD 结构可能已变化 → STALE）",
                    layer_id=layer_ref.id, name=layer_ref.name, index_path=path)
            return layer
        # 中间层：继续进入其子层（若为非组，下一次 container.Layers 会抛异常 -> 越界错误）
        container = layer
    raise LayerResolutionError("index_path 为空", layer_id=layer_ref.id,
                               name=layer_ref.name, index_path=path)


def _count_layers(container):
    try:
        return container.Layers.Count
    except Exception:
        return 0


# ---------------------------------------------------------------------------
# 迁移 / 重绑定（旧 name-only 配置 -> LayerRef）
# ---------------------------------------------------------------------------
def rebind_layer_reference(index, ref_or_name):
    """把旧引用（LayerRef / dict / str name）绑定到当前 LayerIndex。

    返回 (status, LayerRef | None)：
      VALID     —— index_path 存在且 name 特征一致，直接有效；
      MIGRATED  —— index_path 失效（或纯 name），但 name 全局唯一，返回唯一候选；
      AMBIGUOUS —— name 出现多个，不自动选（返回 None）；
      MISSING   —— 完全找不到（返回 None）。
    """
    ref = ref_from_config(ref_or_name) if not isinstance(ref_or_name, LayerRef) else ref_or_name
    if ref is None:
        return (MISSING, None)

    # 1) 有 id 且 id 在当前 index 中
    if ref.id and ref.id in index.by_id:
        cand = index.by_id[ref.id]
        if _name_matches(cand.name, ref.name):
            return (VALID, cand)
        # id 存在但 name 变了：结构可能变化 -> 按旧 name 找候选
        return _rebase_by_name(index, ref.name)

    # 2) 无 id 或 id 失效：按 name
    return _rebase_by_name(index, ref.name)


def _rebase_by_name(index, name):
    matches = index.find_matching(name)
    if len(matches) == 1:
        return (MIGRATED, matches[0])
    if len(matches) > 1:
        return (AMBIGUOUS, None)
    return (MISSING, None)


# ---------------------------------------------------------------------------
# 序列化（配置存取）
# ---------------------------------------------------------------------------
def serialize_ref(ref):
    """LayerRef -> JSON 可序列化 dict（不含任何 COM 对象）。"""
    return {
        "layer_id": ref.id,
        "name": ref.name,
        "display_path": ref.display_path,
        "is_group": ref.is_group,
        "is_text": ref.is_text,
    }


def ref_from_config(v):
    """接受新格式 dict 或旧格式 str（name-only），返回 LayerRef 或 None。

    - dict：{"layer_id": "0/2/1", "name": "电话", "display_path": "...", ...}
    - str ：旧配置，纯图层名（id 为空，index_path 为空，需要 rebind）
    """
    if isinstance(v, dict):
        iid = v.get("layer_id")
        name = v.get("name")
        if not name:
            return None
        ip = ()
        if iid:
            try:
                ip = tuple(int(x) for x in str(iid).split("/"))
            except (TypeError, ValueError):
                ip = ()
        return LayerRef(
            id=str(iid) if iid else _id_of(ip),
            name=str(name),
            display_path=str(v.get("display_path", name)),
            index_path=ip,
            is_group=bool(v.get("is_group", False)),
            is_text=bool(v.get("is_text", False)),
        )
    if isinstance(v, str) and v.strip():
        name = v.strip()
        return LayerRef(id="", name=name, display_path=name, index_path=())
    return None


# ---------------------------------------------------------------------------
# 展示
# ---------------------------------------------------------------------------
def unique_display_labels(refs):
    """返回 {ref.id: label}，label 默认 = display_path；display_path 重复时附加 id 后缀，
    保证同名图层在 GUI 中肉眼可区分且 label 唯一。"""
    from collections import Counter
    labels = {r.id: r.display_path for r in refs}
    cnt = Counter(labels.values())
    for rid, lab in labels.items():
        if cnt[lab] > 1:
            labels[rid] = f"{lab} [id={rid}]"
    return labels


# ---------------------------------------------------------------------------
# LayerIndex
# ---------------------------------------------------------------------------
class LayerIndex:
    """PSD 全部图层的索引：不丢同名、by_name 为 list、提供精确 resolve。"""

    def __init__(self, layers):
        self.layers = list(layers)
        self.by_id = {}
        self.by_name = {}
        for ref in self.layers:
            if ref.id:
                self.by_id[ref.id] = ref
            self.by_name.setdefault(ref.name, []).append(ref)

    # ---- 查询 ----
    def find_by_name(self, name):
        """精确 name 查找（不忽略空格），返回 list（可能多个）。"""
        return list(self.by_name.get(name, []))

    def find_matching(self, name):
        """忽略空格 + 小写匹配（与 resolve 校验一致），返回 list。"""
        return [r for r in self.layers if _name_matches(r.name, name)]

    def get(self, layer_id):
        return self.by_id.get(layer_id)

    def labels(self):
        return unique_display_labels(self.layers)

    # ---- 定位 ----
    def resolve(self, document, layer_ref):
        return resolve_layer(document, layer_ref)

    def __len__(self):
        return len(self.layers)

    def __iter__(self):
        return iter(self.layers)

    def __repr__(self):
        return f"<LayerIndex {len(self.layers)} layers>"
