"""Excel 定义名称（名称管理器）解析与作用域解析。

工作簿级名称存于 ``wb.defined_names``，工作表级名称存于
``ws.defined_names``（保存时带 ``localSheetId``）。公式中出现的名称按
Excel 语义解析作用域：

- 无限定的名称（``单价``）：先查公式所在工作表的工作表级名称，再查
  工作簿级名称；
- 跨表限定名称（``计算!单价``）：限定表必须存在，先查该表工作表级名称，
  再回退工作簿级名称。

名称体（``attr_text``）可能是：单元格 / 区域（可含多个区域，以顶层逗号
分隔）、嵌套名称、常量、公式、``#REF!`` 或指向外部工作簿。可指向
单元格 / 区域的名称会被展开为依赖边；其余状态（常量 / 公式）只记录解析
结果；无法解析或指向不存在工作表的名称产生审计问题，且不加入影响面。
"""
from __future__ import annotations

from dataclasses import dataclass, field

from .formula_parser import (
    RawName,
    RawRef,
    tokenize_formula,
)

# 名称嵌套解析的最大深度，防止异常名称链导致的栈膨胀
MAX_NAME_DEPTH = 50

# 解析状态
RESOLVED = "resolved"        # 已展开为一个或多个单元格 / 区域
CONSTANT = "constant"        # 数字 / 字符串 / 布尔等常量名称，不产生边
FORMULA = "formula"          # 公式表达式名称（其中显式引用仍展开为边）
UNDEFINED = "undefined"      # 名称在对应作用域中不存在
INVALID = "invalid"          # 名称存在但目标无效（#REF! / 名称环 / 空定义）
EXTERNAL = "external"        # 指向外部工作簿，不参与本工作簿影响面
NAME_INITIAL = "__initial__"  # 内部：尚未确定状态

STATUS_LABELS = {
    RESOLVED: "已解析",
    CONSTANT: "常量",
    FORMULA: "公式",
    UNDEFINED: "未定义",
    INVALID: "目标无效",
    EXTERNAL: "外部引用",
}

# 仅记录但不参与解析的内置保留名称（打印区域 / 筛选区域等）
_RESERVED_PREFIX = "_xlnm."


@dataclass(frozen=True)
class NameDef:
    """一条定义名称的静态信息。"""

    name: str
    scope: str                 # "workbook" 或工作表名
    local: bool                # 是否工作表级
    attr_text: str | None
    hidden: bool = False
    comment: str | None = None


@dataclass(frozen=True)
class NameIssue:
    """名称解析过程中产生的、可定位的审计线索。"""

    status: str                # UNDEFINED / INVALID / EXTERNAL
    reason: str                # 供 issue message 使用的中文说明
    ref_text: str | None = None       # 出问题的引用片段
    target_sheet: str | None = None   # 不存在的目标工作表名（若有）


@dataclass
class NameResolution:
    """单次公式内名称解析的完整结果。"""

    name: str
    qualifier: str | None          # 跨表限定的表名（None = 按公式所在表）
    text: str                      # 公式中的原文
    status: str = NAME_INITIAL
    scope: str | None = None       # 实际命中的定义作用域
    definition: str | None = None  # 命中名称的 attr_text
    targets: list[RawRef] = field(default_factory=list)
    # 每个展开目标附带的证据：(引用片段, 展开经过的名称链)
    traces: list[tuple[str, tuple[str, ...]]] = field(default_factory=list)
    issues: list[NameIssue] = field(default_factory=list)
    chain: tuple[str, ...] = ()    # 第一条目标的展开链（展示用）


def _split_top_level(body: str) -> list[str]:
    """按顶层逗号切分（忽略括号 / 字符串内的逗号）。"""
    parts: list[str] = []
    depth = 0
    start = 0
    in_str = False
    for i, ch in enumerate(body):
        if ch == '"':
            in_str = not in_str
        elif not in_str:
            if ch == "(":
                depth += 1
            elif ch == ")" and depth:
                depth -= 1
            elif ch == "," and depth == 0:
                parts.append(body[start:i])
                start = i + 1
    parts.append(body[start:])
    return [p.strip() for p in parts if p.strip()]


def _classify_segment(segment: str) -> str:
    """分类名称体片段（已排除外部 / 错误前缀）：reference/name/constant/formula。

    - 仅由引用 / 名称与顶层逗号组成（可多区域联合）→ reference；
    - 仅一个名称 token 完整覆盖片段 → name（嵌套名称）；
    - 含引用 / 名称但还有运算符、函数调用等 → formula；
    - 其余（数字 / 字符串 / 错误表达式）→ constant。

    token 扫描会跳过数字 / 运算符等，故不能只看 token 数；这里额外检查
    片段原文是否出现括号（函数调用）或括号外的算术 / 比较运算符。
    """
    s = segment.strip()
    if not s:
        return CONSTANT
    tokens = tokenize_formula(s)
    if not tokens:
        return CONSTANT
    depth = 0
    in_str = False
    paren_seen = False
    for ch in s:
        if ch == '"':
            in_str = not in_str
        elif not in_str:
            if ch == "(":
                paren_seen = True  # 多区域联合不含括号，出现即函数调用
                depth += 1
            elif ch == ")":
                depth = max(0, depth - 1)
            elif ch in "+-*/^&<>=%" and depth == 0:
                return FORMULA
    if paren_seen:
        return FORMULA
    if len(tokens) == 1:
        return "reference" if tokens[0].kind == "ref" else "name"
    # 多个 token 且无运算符 / 括号：多区域联合（顶层逗号已由切分保证）
    return "reference" if any(t.kind == "ref" for t in tokens) else "name"


def build_name_index(wb_formula) -> dict[str, list[NameDef]]:
    """收集工作簿中全部定义名称，按名称大写形式索引（保留同名不同作用域）。"""
    index: dict[str, list[NameDef]] = {}

    def add(name: str, scope: str, local: bool, dn) -> None:
        if name.startswith(_RESERVED_PREFIX):
            return
        attr = dn.attr_text
        if isinstance(attr, str) and attr.strip().lower() == "none":
            attr = None
        index.setdefault(name.upper(), []).append(
            NameDef(name=name, scope=scope, local=local, attr_text=attr,
                    hidden=bool(getattr(dn, "hidden", False)),
                    comment=getattr(dn, "comment", None))
        )

    for dn in wb_formula.defined_names.values():
        add(dn.name, "workbook", False, dn)
    for ws in wb_formula.worksheets:
        for dn in ws.defined_names.values():
            add(dn.name, ws.title, True, dn)
    return index


class NameResolver:
    """按作用域解析公式中的定义名称，展开为 RawRef 目标。"""

    def __init__(self, wb_formula, sheet_names: set[str]):
        self.index = build_name_index(wb_formula)
        self.sheet_names = sheet_names

    # ------------------------------------------------------------ 查询
    def inventory(self) -> list[NameDef]:
        """全部定义名称（去重后，供导出）。"""
        seen: set[tuple[str, str]] = set()
        out: list[NameDef] = []
        for defs in self.index.values():
            for d in defs:
                key = (d.name, d.scope)
                if key in seen:
                    continue
                seen.add(key)
                out.append(d)
        out.sort(key=lambda d: (d.scope != "workbook", d.scope, d.name))
        return out

    def resolve(self, raw: RawName, context_sheet: str) -> NameResolution:
        """解析公式中的一个名称。context_sheet 为公式所在工作表。"""
        res = NameResolution(name=raw.name, qualifier=raw.qualifier, text=raw.text)
        qualifier = raw.qualifier
        if qualifier is not None and qualifier not in self.sheet_names:
            # 跨表限定名称指向不存在的工作表：不解析、不入图
            res.status = UNDEFINED
            res.issues.append(NameIssue(
                status=UNDEFINED,
                reason=f"限定工作表 {qualifier!r} 不存在，无法解析名称 {raw.name!r}",
                ref_text=raw.text, target_sheet=qualifier))
            return res

        defn = self._lookup(raw.name, qualifier, context_sheet)
        if defn is None:
            res.status = UNDEFINED
            scope_hint = f"工作表「{qualifier}」及" if qualifier else ""
            res.issues.append(NameIssue(
                status=UNDEFINED,
                reason=f"{scope_hint}工作簿名称管理器中均未定义名称 {raw.name!r}",
                ref_text=raw.text, target_sheet=qualifier))
            return res

        res.scope = defn.scope
        res.definition = defn.attr_text
        self._expand_definition(defn, context_sheet, res, set())
        return res

    def _lookup(
        self, name: str, qualifier: str | None, context_sheet: str
    ) -> NameDef | None:
        """按作用域查找名称：限定表的本地名 > 工作簿名；无限定时先本表本地。"""
        defs = self.index.get(name.upper())
        if not defs:
            return None
        if qualifier is not None:
            local = next((d for d in defs if d.local and d.scope == qualifier), None)
            return local or next((d for d in defs if not d.local), None)
        local = next((d for d in defs if d.local and d.scope == context_sheet), None)
        return local or next((d for d in defs if not d.local), None)

    @staticmethod
    def _display_chain_name(defn: NameDef) -> str:
        if defn.local:
            return f"{defn.scope}!{defn.name}"
        return defn.name

    # ------------------------------------------------------------ 展开
    def _expand_definition(
        self,
        defn: NameDef,
        context_sheet: str,
        res: NameResolution,
        visited: set[tuple[str, str]],
    ) -> None:
        """把一条名称定义展开进 res（迭代工作清单，避免递归环）。

        工作表级名称体内未限定的引用相对其所属工作表；工作簿级名称体内
        未限定的引用相对**使用该名称的公式所在工作表**（Excel 语义）。
        每个工作项携带独立名称链，保证钻石型引用各自留证、不互相串链。
        """
        # (定义, 体内未限定引用默认表, 深度, 名称链, 路径已访问定义, 父表达式标记)
        # parent_formula=True 表示该定义是从上层「公式表达式」片段中经嵌套
        # 名称进入的——其引用即使被展开，整次解析的状态仍为 formula。
        # 路径集合用不可变 frozenset：钻石型（同一片段多次引用同一名称）的
        # 兄弟分支各自派生，互不污染；只有真的沿同一条链回到祖先才判环。
        root_seen: frozenset = frozenset(visited)
        # 链（chain）语义：**已处理祖先**的展示名路径，不含当前工作项。
        # 弹出当前节点并通过环检测后，cur_chain 才包含当前节点。
        work: list[tuple[NameDef, str, int, tuple[str, ...], frozenset, bool]] = [
            (defn, context_sheet, 0, (), root_seen, False)
        ]
        while work:
            cur, owner_ctx, depth, chain, seen, parent_formula = work.pop()
            key = (cur.name, cur.scope)
            cur_name = self._display_chain_name(cur)
            if key in seen:
                if cur_name in chain:
                    # 回到链上祖先：从该祖先（含）到链尾，再闭合回该祖先
                    idx = chain.index(cur_name)
                    cycle_path = chain[idx:] + (cur_name,)
                else:
                    cycle_path = chain + (cur_name,)
                res.status = INVALID
                res.issues.append(NameIssue(
                    status=INVALID,
                    reason=f"名称 {cur.name!r} 的定义存在循环引用链："
                           f"{' -> '.join(cycle_path)}",
                    ref_text=cur.attr_text))
                continue
            child_seen = seen | {key}
            cur_chain = chain + (cur_name,)
            body = cur.attr_text
            if body is None:
                res.status = INVALID
                res.issues.append(NameIssue(
                    status=INVALID,
                    reason=f"名称 {cur.name!r} 没有定义内容（空名称）",
                    ref_text=None))
                continue
            body = body.strip()
            if body.startswith("="):
                body = body[1:].strip()
            # 名称体内未限定引用的默认工作表
            default_sheet = cur.scope if cur.local else owner_ctx
            # 该定义自身段产生的正常状态：纯引用段仍可能处于父表达式内
            ref_status = FORMULA if parent_formula else RESOLVED
            # 本次定义体内是否出现过可展开为边的段（否则常量段才有资格定状态）
            expandable_seen = False

            def mark_ok(status: str) -> None:
                """写入正常（非错误）状态：错误一旦产生不被覆盖；
                FORMULA 优先级高于 RESOLVED（表达式名称即使含直接引用也按公式记）。"""
                if res.status in (UNDEFINED, INVALID, EXTERNAL):
                    return
                if res.status == FORMULA and status == RESOLVED:
                    return
                res.status = status

            for seg in _split_top_level(body):
                if self._is_external(seg):
                    if res.status not in (UNDEFINED, INVALID):
                        res.status = EXTERNAL
                    res.issues.append(NameIssue(
                        status=EXTERNAL,
                        reason=f"名称 {cur.name!r} 指向外部工作簿（{seg}），"
                               "外部引用不参与影响面分析",
                        ref_text=seg))
                    continue
                if seg.startswith("#"):
                    res.status = INVALID
                    res.issues.append(NameIssue(
                        status=INVALID,
                        reason=f"名称 {cur.name!r} 的目标已失效（{seg}）",
                        ref_text=seg))
                    continue
                kind = _classify_segment(seg)
                if kind == CONSTANT:
                    continue  # 常量段延后处理，见函数末尾
                if kind == FORMULA:
                    expandable_seen = True
                    # 表达式片段：展开其中明确出现的引用；嵌套名称进入时
                    # 携带 parent_formula=True。
                    mark_ok(FORMULA)
                    for ref in self._refs_of(seg, default_sheet):
                        self._add_target(res, ref, seg, cur_chain)
                    if depth >= MAX_NAME_DEPTH:
                        res.status = INVALID
                        res.issues.append(NameIssue(
                            status=INVALID,
                            reason=f"名称嵌套深度超过 {MAX_NAME_DEPTH} 层",
                            ref_text=cur.attr_text))
                        continue
                    for rn in self._nested_names_of(seg):
                        nd = self._lookup(rn.name, rn.qualifier, default_sheet)
                        if nd is None:
                            res.status = UNDEFINED
                            res.issues.append(NameIssue(
                                status=UNDEFINED,
                                reason=f"名称 {cur.name!r} 的表达式中使用了"
                                       f"未定义的名称 {rn.name!r}",
                                ref_text=rn.text,
                                target_sheet=rn.qualifier))
                            continue
                        work.append((nd, default_sheet, depth + 1,
                                     cur_chain, child_seen, True))
                    continue
                if kind == "name":
                    expandable_seen = True
                    rn = self._sole_name(seg)
                    nd = self._lookup(rn.name, rn.qualifier, default_sheet)
                    if nd is None:
                        res.status = UNDEFINED
                        q = rn.qualifier
                        reason = f"名称 {cur.name!r} 引用的名称 {rn.name!r} 未定义"
                        if q is not None and q not in self.sheet_names:
                            reason += f"（限定工作表 {q!r} 不存在）"
                        res.issues.append(NameIssue(
                            status=UNDEFINED, reason=reason,
                            ref_text=rn.text, target_sheet=q))
                        continue
                    if depth >= MAX_NAME_DEPTH:
                        res.status = INVALID
                        res.issues.append(NameIssue(
                            status=INVALID,
                            reason=f"名称嵌套深度超过 {MAX_NAME_DEPTH} 层",
                            ref_text=cur.attr_text))
                        continue
                    work.append((nd, default_sheet, depth + 1,
                                 cur_chain, child_seen, parent_formula))
                    continue
                # kind == "reference"：纯引用段（可能是多区域之一）
                expandable_seen = True
                mark_ok(ref_status)
                for ref in self._refs_of(seg, default_sheet):
                    self._add_target(res, ref, seg, cur_chain)
            # 整个定义体只有常量 / 运算符（无引用、无嵌套名称）
            if not expandable_seen and not parent_formula and \
                    res.status not in (UNDEFINED, INVALID, EXTERNAL,
                                       RESOLVED, FORMULA):
                res.status = CONSTANT
        if res.status == NAME_INITIAL:
            res.status = CONSTANT
        if not res.chain and res.traces:
            res.chain = res.traces[0][1]

    @staticmethod
    def _is_external(seg: str) -> bool:
        return seg.startswith("[") or "]!" in seg

    def _add_target(self, res: NameResolution, ref: RawRef, seg: str,
                    chain: tuple[str, ...]) -> None:
        if ref in res.targets:
            return
        res.targets.append(ref)
        res.traces.append((seg, chain))

    @staticmethod
    def _refs_of(segment: str, default_sheet: str) -> list[RawRef]:
        out: list[RawRef] = []
        for t in tokenize_formula(segment):
            if t.kind != "ref":
                continue
            r = t.ref
            if r.sheet is None:
                out.append(RawRef(default_sheet, r.kind, r.col1, r.row1,
                                  r.col2, r.row2, r.text))
            else:
                out.append(r)
        return out

    @staticmethod
    def _nested_names_of(segment: str) -> list[RawName]:
        return [t.name for t in tokenize_formula(segment) if t.kind == "name"]

    @staticmethod
    def _sole_name(segment: str) -> RawName:
        return next(t.name for t in tokenize_formula(segment) if t.kind == "name")
