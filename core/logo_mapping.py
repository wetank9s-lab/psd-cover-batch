# -*- coding: utf-8 -*-
"""
logo_mapping.py —— Stage 3 Logo 模型与映射逻辑（纯 core，不依赖 COM / Tk）。

把过去混在一起的三个概念彻底分开：

  1. selected_logo_refs        用户 GUI 左侧勾选的对象（可能是组，也可能是叶子）
  2. effective_logo_leaf_refs  selected 递归展开后的最终可操作叶子
  3. store_logo_map / brand_logo_refs  运行时唯一可信的显隐来源

本模块的全部函数都是纯函数（只处理 LayerRef / LayerIndex 等数据对象），
不触碰 Photoshop COM，不触碰 Tk。

设计红线（Stage 2 LayerRef 红线在 Logo 层的延伸）：
  - 运行时绝不再用 name heuristic（"logo" in name / fuzzy）决定业务角色；
  - name heuristic 只允许用于「首次加载的自动推荐」；
  - 同名图层按 LayerRef.id 区分，绝不按 name 去重；
  - ambiguous 绝不自动选第一个；
  - store_logo_map 的 value 必须是叶子（is_group=False）；
  - brand_logo_refs 同样必须是叶子；
  - 同一 LayerRef 不能同时是 STORE_LOGO 与 BRAND_LOGO（冲突必须被 Preflight 拒绝）。
"""

from dataclasses import dataclass, field
from enum import Enum

from core.util import normalize_name

# ---------------------------------------------------------------------------
# 角色枚举
# ---------------------------------------------------------------------------
class LogoRole(Enum):
    STORE = "store"
    BRAND = "brand"
    IGNORE = "ignore"


# 匹配状态常量（与 core.layer_index 的 rebind 状态风格一致）
EXACT = "EXACT"            # 归一化完全相等（100 分）
AUTO = "AUTO"              # 唯一高分自动匹配（90/85/70 分）
AMBIGUOUS = "AMBIGUOUS"    # 多个候选同分 / 差异不足以唯一确定
NO_MATCH = "NO_MATCH"      # 无任何候选命中

# 用户确认状态
USER_CONFIRMED = "USER_CONFIRMED"
AUTO_MATCHED = "AUTO_MATCHED"
UNRESOLVED = "UNRESOLVED"

# 旧配置迁移状态（与 layer_index 一致）
VALID = "VALID"
MIGRATED = "MIGRATED"
MISSING = "MISSING"


# ---------------------------------------------------------------------------
# 数据模型
# ---------------------------------------------------------------------------
@dataclass
class LogoMatchResult:
    """match_store_logo() 的返回：一个门店的自动匹配结果。"""
    status: str                  # EXACT / AUTO / AMBIGUOUS / NO_MATCH
    best: "LayerRef | None"      # 唯一最优候选（EXACT/AUTO 时非 None）
    score: "int | None"          # best 的得分
    candidates: list              # 参与匹配的全部候选 LayerRef（保持传入顺序）
    hits: list = field(default_factory=list)   # 命中候选（score > 0）


@dataclass
class LogoMapping:
    """一次完整 Logo 配置（运行时唯一可信来源）。

    store_logo_map : dict[str, LayerRef | None]  —— 门店名 -> 叶子 LayerRef（None 表示未映射）
    brand_logo_refs: list[LayerRef]              —— 品牌固定 Logo 叶子列表
    logo_selection_refs: list[LayerRef]          —— 用户 GUI 勾选（可能含组），仅保存用
    """
    store_logo_map: dict = field(default_factory=dict)
    brand_logo_refs: list = field(default_factory=list)
    logo_selection_refs: list = field(default_factory=list)


# ---------------------------------------------------------------------------
# 1. 组展开：selected -> effective leaves
# ---------------------------------------------------------------------------
def resolve_effective_logo_layers(layer_index, selected_refs):
    """把用户勾选（可能含组）递归展开为最终可操作叶子 LayerRef 列表。

    规则：
      - 叶子 selection 保留自身；
      - 组 selection 递归展开其下所有叶子（组本身不进入结果）；
      - 多个 selection 展开后按 LayerRef.id 去重（绝不按 name 去重）；
      - 顺序稳定：按 selected_refs 传入顺序 + LayerIndex 内顺序展开。

    返回：list[LayerRef]，全部 is_group=False。
    """
    if layer_index is None:
        return []
    seen = set()
    out = []
    layer_index = _as_index(layer_index)

    def add_ref(ref):
        if ref is None:
            return
        if ref.id in seen:
            return
        seen.add(ref.id)
        out.append(ref)

    for ref in selected_refs or []:
        if ref is None:
            continue
        r = layer_index.get(ref.id) or ref
        if r.is_group:
            _walk_leaves(layer_index, r, add_ref)
        else:
            add_ref(r)
    return out


def _walk_leaves(index, group_ref, add_ref):
    """收集 group 下的所有叶子（递归），按 LayerIndex 顺序。"""
    for child in _children_of(index, group_ref):
        if child.is_group:
            _walk_leaves(index, child, add_ref)
        else:
            add_ref(child)


def _children_of(index, group_ref):
    """返回 group_ref 的直接子 LayerRef（按 LayerIndex 顺序；找不到则空）。"""
    prefix = group_ref.id + "/"
    kids = [r for r in index.layers if r.id.startswith(prefix)
            and len(r.index_path) == len(group_ref.index_path) + 1]
    return kids


def _as_index(layer_index):
    """兼容传入 LayerIndex 或 LayerIndex.layers（list）。"""
    if hasattr(layer_index, "layers"):
        return layer_index
    return _ListIndex(layer_index)


class _ListIndex:
    """极简包装：把 list[LayerRef] 包成带 get()/layers 的对象。"""
    def __init__(self, layers):
        self.layers = list(layers)
        self._by_id = {r.id: r for r in self.layers}

    def get(self, rid):
        return self._by_id.get(rid)


# ---------------------------------------------------------------------------
# 2. 自动匹配（评分制，确定性强、可解释）
# ---------------------------------------------------------------------------
def match_store_logo(store, candidates):
    """为门店在候选叶子 LayerRef 中自动匹配最合适的 Logo。

    评分规则（与 Stage 0 pick_best_logo 保持一致，纯加分、无过滤）：
      normalized exact               100
      layer startswith store          90
      store startswith layer          85
      mutual contains                 70
      parent/display path contains logo +10 辅助（加分项，不要求）
      layer name contains logo         +5 辅助（加分项，不要求）

    关键变化（解决「康乐 -> 康乐电器」P0）：
      candidates 是「effective leaf 全集」，**不做** is_logo_candidate 预过滤；
      name 是否含 logo / 父路径是否含 logo 只是加分项，绝不过滤。

    返回 LogoMatchResult：
      EXACT —— 归一化完全相等（100 分且唯一）
      AUTO  —— 唯一高分（90/85/70 组合后的唯一最优）
      AMBIGUOUS —— 最高分并列 / 分数差不足唯一确定
      NO_MATCH —— 无任何候选命中
    """
    result = LogoMatchResult(status=NO_MATCH, best=None, score=None,
                             candidates=list(candidates))
    target = normalize_name(store)
    if not target or not candidates:
        return result

    best_score = 0
    best_ref = None
    tied = False
    for ref in candidates:
        if ref is None or ref.is_group:
            continue
        n = normalize_name(ref.name)
        if not n:
            continue
        score = _score_name(target, n)
        if score <= 0:
            continue
        score += _score_aux(ref, target, n)
        result.hits.append(ref)
        if score > best_score:
            best_score = score
            best_ref = ref
            tied = False
        elif score == best_score:
            tied = True

    if best_ref is None:
        result.status = NO_MATCH
        return result
    result.score = best_score
    result.best = best_ref
    if best_score == 100 and not tied:
        result.status = EXACT
    elif tied:
        result.status = AMBIGUOUS
        result.best = None      # 歧义绝不自动选第一个
        result.score = None
    else:
        result.status = AUTO
    return result


def _score_name(target, n):
    """基础评分：仅基于归一化后的图层名 n 与门店名 target。"""
    if n == target:
        return 100
    if n.startswith(target):
        return 90
    if target.startswith(n):
        return 85
    if target in n or n in target:
        return 70
    return 0


def _score_aux(ref, target, n):
    """辅助加分：父路径含 logo（+10）、图层名含 logo（+5）。只加分，绝不过滤。"""
    bonus = 0
    dp = (ref.display_path or "").lower()
    if "logo" in dp:
        bonus += 10
    if "logo" in n:
        bonus += 5
    return bonus


# ---------------------------------------------------------------------------
# 3. Preflight 校验
# ---------------------------------------------------------------------------
class LogoValidationError(Exception):
    """Logo 配置校验失败（含冲突 / 无效映射 / 未解决歧义）。"""

    def __init__(self, message, *, code=None):
        super().__init__(message)
        self.code = code


def validate_logo_mapping(logo_mapping, effective_leaf_refs=None, store_names=None,
                          *, allow_duplicate_store_targets=True):
    """校验 LogoMapping 是否可安全运行。

    至少检查：
      - store target 存在（LayerRef 非 None）
      - store target 是叶子（is_group=False）
      - store target 在 effective selected leaves 中（若提供 effective 集合）
      - brand refs 存在、是叶子
      - store / brand 无冲突（同一 LayerRef 不能同时是两个角色）
      - ambiguous / missing 未解决（store 有值但映射为 None 或未命中）
      - 两个 store 映射到同一 leaf：默认允许（业务可共用 Logo），
        allow_duplicate_store_targets=False 时视为错误

    返回：None（通过）或抛 LogoValidationError（不通过，附 code）。
    """
    if logo_mapping is None:
        raise LogoValidationError("Logo 配置为空", code="EMPTY")

    effective = {r.id for r in (effective_leaf_refs or [])}
    store_map = logo_mapping.store_logo_map or {}
    brand = logo_mapping.brand_logo_refs or []

    # ---- brand 校验 ----
    brand_ids = set()
    for br in brand:
        if br is None:
            raise LogoValidationError("品牌 Logo 存在空引用", code="BRAND_NULL")
        if br.is_group:
            raise LogoValidationError(
                f"品牌 Logo 必须是叶子图层：{br.display_path!r}（is_group=True）",
                code="BRAND_GROUP")
        brand_ids.add(br.id)

    # ---- store 校验 ----
    store_target_ids = {}
    for s, ref in (store_map or {}).items():
        if store_names is not None and s not in store_names:
            # 不在当前 Excel 门店集合：跳过（不构成运行时错误）
            continue
        if ref is None or not ref.id:
            raise LogoValidationError(
                f"门店“{s}”没有有效 Logo 映射。", code="STORE_UNMAPPED")
        if ref.is_group:
            raise LogoValidationError(
                f"门店“{s}”的 Logo 映射必须是叶子图层，不能是组：{ref.display_path!r}",
                code="STORE_GROUP")
        if effective and ref.id not in effective:
            raise LogoValidationError(
                f"门店“{s}”的 Logo 映射 {ref.display_path!r} 不在当前勾选的 Logo 范围中",
                code="STORE_NOT_IN_EFFECTIVE")
        # 记录 id -> 门店（用于重复 target 检测）
        store_target_ids.setdefault(ref.id, []).append(s)

    # ---- store / brand 冲突 ----
    for rid, ref in [(r.id, r) for r in (store_map or {}).values() if r is not None]:
        if rid in brand_ids:
            raise LogoValidationError(
                f"Logo 冲突：图层 {ref.display_path!r} 同时被用作门店 Logo 与品牌 Logo。"
                f"请修正映射。", code="STORE_BRAND_CONFLICT")

    # ---- 重复 target：默认允许，可选拒绝 ----
    if not allow_duplicate_store_targets:
        for rid, ss in store_target_ids.items():
            if len(ss) > 1:
                raise LogoValidationError(
                    f"多个门店共用同一 Logo（{rid}）：{', '.join(ss)}",
                    code="DUPLICATE_STORE_TARGET")

    return None


# ---------------------------------------------------------------------------
# 4. 运行时可见性计划（纯数据：不触碰 COM）
# ---------------------------------------------------------------------------
def prepare_logo_visibility(store, logo_mapping, *, require_store_mapping=False,
                            effective_leaf_refs=None):
    """为当前行生成 Logo 可见性计划。

    返回：list[(LayerRef, visible)] —— 每个门店叶子 + 品牌叶子的显隐决定。
    这是「纯计划」，不触碰 COM；由调用方 resolve 并写 Visible。

    规则（确定性）：
      - 当前门店映射的叶子 visible=True；
      - 其余被映射为门店 Logo 的叶子 visible=False；
      - 所有品牌叶子 visible=True；
      - 若提供 effective_leaf_refs（用户勾选展开的候选叶子全集）：
        其中「未被任何门店映射、也非品牌」的叶子 visible=False
        （候选 store Logo 但未被当前门店使用 -> 隐藏，绝不保留模板原状态）。

    若 require_store_mapping=True 且当前门店无映射（或映射无效）：
      抛 LogoVisibilityError，绝不静默保留模板原 Logo。
    """
    if logo_mapping is None:
        raise LogoVisibilityError("Logo 配置为空，无法生成可见性计划")
    if require_store_mapping and not store:
        raise LogoVisibilityError("门店为空，且门店 Logo 为必需。")

    store_map = logo_mapping.store_logo_map or {}
    brand_refs = logo_mapping.brand_logo_refs or []
    brand_ids = {r.id for r in brand_refs if r is not None}

    target = store_map.get(store)
    if require_store_mapping and (target is None or not target.id):
        raise LogoVisibilityError(f"门店“{store}”没有有效 Logo 映射。")

    plan = []
    seen_ids = set()
    # 1) 所有门店映射叶子：当前门店 True，其余 False
    for s, ref in store_map.items():
        if ref is None or not ref.id or ref.id in seen_ids:
            continue
        seen_ids.add(ref.id)
        plan.append((ref, (s == store)))
    # 2) 品牌叶子：始终 True
    for br in brand_refs:
        if br is None or br.id in seen_ids:
            continue
        seen_ids.add(br.id)
        plan.append((br, True))
    # 3) effective 中未被映射/品牌的候选叶子：当前行未使用 -> False
    #    （候选 store Logo 但未被任何门店映射 / 或映射给了其他门店但当前门店未用）
    if effective_leaf_refs is not None:
        for ref in effective_leaf_refs:
            if ref is None or ref.id in seen_ids:
                continue
            seen_ids.add(ref.id)
            plan.append((ref, False))
    return plan


class LogoVisibilityError(Exception):
    """运行时 Logo 可见性错误（无映射 / 冲突等），绝不能静默成功。"""

    def __init__(self, message, *, code=None, store=None):
        super().__init__(message)
        self.code = code
        self.store = store


def verify_logo_visibility(actual_visible_map, plan):
    """写 Visible 后 read-back 校验（纯函数）。

    actual_visible_map: dict[ref_id -> bool]（read-back 结果）
    plan: prepare_logo_visibility 的返回值

    不一致抛 LogoVisibilityError（不能静默认为成功）。
    """
    for ref, expected in plan:
        got = actual_visible_map.get(ref.id)
        if got is None:
            raise LogoVisibilityError(
                f"图层 {ref.display_path!r} 未回读可见性（read-back 缺失）",
                code="READBACK_MISSING")
        if bool(got) != bool(expected):
            raise LogoVisibilityError(
                f"图层 {ref.display_path!r} 可见性回读不一致：期望 {expected}，实际 {got}",
                code="READBACK_MISMATCH")
    return True


# ---------------------------------------------------------------------------
# 5. 旧配置迁移（name-only -> LayerRef）
# ---------------------------------------------------------------------------
def migrate_old_config(old_logo_layers, old_store_logo_map, layer_index):
    """把旧配置（logo_layers 字符串列表 / store_logo_map 字符串 dict）迁移为新模型。

    返回 (LogoMapping, reports)：
      LogoMapping —— 迁移后的配置；
      reports     —— 逐项迁移说明（status: MIGRATED / AMBIGUOUS / MISSING / VALID）。

    规则（绝不自动选第一个）：
      - 旧 selection name -> LayerIndex rebind：唯一命中 MIGRATED；
        多同名 AMBIGUOUS（保持为未解析）；不存在 MISSING；
      - 旧 store logo name 同理：唯一命中 -> 叶子；歧义/缺失 -> 保持 None（未映射）。
    """
    reports = []
    selected = []
    seen = set()
    for item in old_logo_layers or []:
        name = _name_of(item)
        if not name:
            continue
        status, ref = _rebase(layer_index, name)
        if ref is not None and ref.id not in seen:
            seen.add(ref.id)
            selected.append(ref)
        reports.append({"item": name, "status": status})

    store_map = {}
    for s, v in (old_store_logo_map or {}).items():
        name = _name_of(v)
        if not name:
            store_map[s] = None
            reports.append({"store": s, "status": MISSING})
            continue
        status, ref = _rebase(layer_index, name)
        store_map[s] = ref
        reports.append({"store": s, "status": status, "ref": ref})

    mapping = LogoMapping(
        store_logo_map=store_map,
        brand_logo_refs=[],          # 旧配置没有独立 brand 字段：留待 heuristic 推荐
        logo_selection_refs=selected,
    )
    return mapping, reports


def _rebase(layer_index, name):
    """在 LayerIndex 中按 name 唯一命中 -> (MIGRATED, ref)；歧义/缺失 -> (AMBIGUOUS/MISSING, None)。"""
    if layer_index is None:
        return MISSING, None
    matches = [r for r in layer_index.layers if _norm_eq(r.name, name)]
    if len(matches) == 1:
        return MIGRATED, matches[0]
    if len(matches) > 1:
        return AMBIGUOUS, None
    return MISSING, None


def _norm_eq(a, b):
    return bool(a) and bool(b) and normalize_name(a) == normalize_name(b)


def _name_of(v):
    """从旧配置值（str name / LayerRef dict / LayerRef）中取 name。"""
    if v is None:
        return ""
    if isinstance(v, str):
        return v.strip()
    if isinstance(v, dict):
        return str(v.get("name", "")).strip()
    return str(getattr(v, "name", "")).strip()


# ---------------------------------------------------------------------------
# 6. 品牌 Logo 首次推荐（仅加载时用，运行时绝不调用）
# ---------------------------------------------------------------------------
def suggest_brand_logos(effective_leaf_refs, store_map):
    """首次加载时推荐品牌 Logo（heuristic，只用于建议，不用于运行）。

    规则：
      - name 含 "logo"（归一化）
      - 且未被任何 store mapping 使用
      - 且是叶子

    返回 list[LayerRef]。运行时永远不调用本函数。
    """
    mapped_ids = set()
    for ref in (store_map or {}).values():
        if ref is not None and ref.id:
            mapped_ids.add(ref.id)
    out = []
    for ref in effective_leaf_refs or []:
        if ref.is_group:
            continue
        if ref.id in mapped_ids:
            continue
        if "logo" in normalize_name(ref.name):
            out.append(ref)
    return out


# ---------------------------------------------------------------------------
# 7. 首次加载 Logo selection 推荐（修复 P1-02：candidate discovery 不挡 fuzzy match）
# ---------------------------------------------------------------------------
def recommend_logo_selection(all_refs, stores, *, is_logo_heuristic=None):
    """首次加载时推荐「勾选集合」（含组）—— 修复 P1-02 的核心。

    背景：旧逻辑只按 name/父路径含 "logo" 推荐勾选；对「普通素材 > 康乐电器」
    （门店 Excel 名=康乐，图层/父路径均无 logo 关键字）这类结构，目标叶子
    不在推荐勾选内 -> effective 为空 -> match_store_logo 即使不预过滤也看不到
    目标（NO_MATCH）。matcher 修了，candidate discovery 没修，链路仍断。

    方案 B（用户指定）：推荐 = logo heuristic 勾选 ∪ 与任一门店有唯一有效
    match score 的叶子。既保证门店匹配目标一定进入 effective，又不把整个 PSD
    全部默认勾上（左侧不会全选）。

    参数：
      all_refs        —— LayerIndex.layers（含组与叶子）
      stores          —— Excel 门店名列表
      is_logo_heuristic —— 可选：自定 logo 启发式（默认 name/父路径含 logo）。
                           仅用于「品牌/素材类」推荐，不决定门店匹配资格。

    返回 list[LayerRef]：推荐勾选集合（叶子 + 组，可传给 resolve_effective_logo_layers）。
    运行时永远不调用本函数。
    """
    if is_logo_heuristic is None:
        def is_logo_heuristic(ref):
            if ref.is_group:
                return False
            if "logo" in normalize_name(ref.name):
                return True
            dp = (ref.display_path or "").lower()
            # 父路径含 logo（组名等）
            parts = [p.strip() for p in dp.split(">")]
            if len(parts) >= 2 and any("logo" in p for p in parts[:-1]):
                return True
            return False

    out = []
    seen_ids = set()

    def add(ref):
        if ref is not None and ref.id not in seen_ids:
            seen_ids.add(ref.id)
            out.append(ref)

    # 1) 名称含 logo 的叶子（品牌/素材，保持旧推荐行为）
    for ref in all_refs or []:
        if not ref.is_group and is_logo_heuristic(ref):
            add(ref)

    # 2) 与任一门店有「唯一有效 match score」的叶子（修复 P1-02）
    #    —— 只纳入唯一高分命中者；AMBIGUOUS 不自动选（不加入，避免默认全勾歧义层）
    leaf_refs = [r for r in all_refs or [] if not r.is_group]
    for s in stores or []:
        mr = match_store_logo(s, leaf_refs)
        if mr.status in (EXACT, AUTO) and mr.best is not None:
            add(mr.best)
        # AMBIGUOUS / NO_MATCH：不推荐（保持（无），用户自行决定）

    # 3) 若叶子命中者所在组名/路径与门店强相关，也可以把组本身加入推荐
    #    （组展开会让用户看到该组下全部叶子，属于「推荐勾选」，不是运行时规则）
    #    这里保持最小：只推荐叶子本身，不自动推荐组（避免左侧意外全勾）。

    return out
