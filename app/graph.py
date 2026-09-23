"""单元格依赖图：构图（含名称管理器展开）、循环引用检测、影响面/依赖链查询。"""
from __future__ import annotations

from collections import deque
from dataclasses import dataclass, field

from openpyxl.utils import get_column_letter

from .formula_parser import RawRef, tokenize_formula
from .names import (
    EXTERNAL as NAME_EXTERNAL,
    INVALID as NAME_INVALID,
    NameDef,
    NameIssue,
    NameResolver,
    UNDEFINED as NAME_UNDEFINED,
)

# 单个区域最多展开的单元格数，防止 SUM(A:XFD) 之类打爆内存
MAX_RANGE_CELLS = 5000
# 影响面查询的默认上限
MAX_IMPACT_NODES = 5000
# 依赖健康审计：扇出（直接引用单元格数）告警的默认阈值
DEFAULT_FANOUT_THRESHOLD = 20
# 单条审计问题附带的证据边上限，防止超大区域把响应撑爆
MAX_AUDIT_EVIDENCE = 100
# 关键节点：上游子图节点数上限
MAX_KEYNODE_NODES = 5000
# 关键节点：路径数按指数增长，超过该值即截断并标记饱和
PATH_COUNT_CAP = 10**12
# 关键节点：每个节点附带的受影响下游单元格上限（接口 / 导出分别取值）
MAX_KEYNODE_AFFECTED = 100
MAX_KEYNODE_AFFECTED_EXPORT = 1000

# 审计问题级别 -> 排序权重（越严重越靠前）
SEVERITY_WEIGHT = {"error": 0, "warning": 1, "info": 2}
AUDIT_LEVELS = ("error", "warning", "info")


def node_id(sheet: str, coord: str) -> str:
    return f"{sheet}!{coord}"


def split_node(node: str) -> tuple[str, str]:
    """工作表名本身可含 '!'，坐标部分不会，故从右侧切分。"""
    sheet, coord = node.rsplit("!", 1)
    return sheet, coord


def display_id(sheet: str, coord: str) -> str:
    """展示用 id：含特殊字符的工作表名加引号。"""
    if sheet.replace("_", "").isalnum():
        return f"{sheet}!{coord}"
    return f"'{sheet}'!{coord}"


@dataclass
class SheetInfo:
    name: str
    max_row: int
    max_col: int
    formula_count: int = 0


@dataclass(frozen=True)
class MissingSheetRef:
    """跨表引用指向不存在的工作表（构图时直接跳过，只在此留证）。"""

    source: str        # 公式单元格内部 id
    formula: str       # 公式原文
    sheet: str         # 被引用但不存在的工作表名
    ref_text: str      # 引用原文，如 NoSuch!B2
    via_name: str | None = None  # 若经由定义名称展开，则为该名称的展示名


@dataclass(frozen=True)
class OutOfRangeRef:
    """显式区域 / 单元格引用超出目标工作表的已用范围。"""

    source: str        # 公式单元格内部 id
    formula: str       # 公式原文
    sheet: str         # 目标工作表名
    ref_text: str      # 引用原文，如 Data!A99:A101
    coord: str         # 第一个越界坐标（展示用），如 A99
    via_name: str | None = None  # 若经由定义名称展开，则为该名称的展示名


@dataclass(frozen=True)
class NameUsage:
    """公式中一次定义名称的解析留痕（用于导出 / 前端展示 / 审计定位）。"""

    source: str                 # 公式单元格内部 id
    name: str                   # 名称（不含工作表限定）
    qualifier: str | None       # 跨表限定的工作表名
    ref_text: str               # 公式中的原文
    status: str                 # names 模块的解析状态常量
    scope: str | None           # 实际命中的定义作用域
    definition: str | None      # 命中名称的定义原文
    chain: tuple[str, ...]      # 展开经过的名称链
    target_count: int           # 展开并加入依赖图的目标数
    issues: tuple[NameIssue, ...] = ()  # 解析中产生的问题（未定义 / 无效 / 外部）


@dataclass
class WorkbookGraph:
    sheets: list[SheetInfo] = field(default_factory=list)
    formulas: dict[str, str] = field(default_factory=dict)      # node -> 公式文本
    values: dict[str, str] = field(default_factory=dict)        # node -> 缓存计算值
    precedents: dict[str, set[str]] = field(default_factory=dict)  # node -> 它引用的节点
    dependents: dict[str, set[str]] = field(default_factory=dict)  # node -> 引用它的节点
    cycles: list[list[str]] = field(default_factory=list)
    cycle_nodes: set[str] = field(default_factory=set)
    edge_count: int = 0
    missing_sheet_refs: list[MissingSheetRef] = field(default_factory=list)
    out_of_range_refs: list[OutOfRangeRef] = field(default_factory=list)
    # node -> 该公式引用的定义名称留痕（仅导出/展示用）
    name_usages: dict[str, list[NameUsage]] = field(default_factory=dict)
    name_inventory: list[NameDef] = field(default_factory=list)
    name_resolver: NameResolver | None = field(default=None, repr=False)
    # node -> 展示 id 的缓存（关键节点 / 审计等大批量展示场景复用）
    _display_cache: dict[str, str] = field(default_factory=dict, repr=False)

    def display(self, node: str) -> str:
        cached = self._display_cache.get(node)
        if cached is None:
            cached = display_id(*split_node(node))
            self._display_cache[node] = cached
        return cached

    # ------------------------------------------------------------------ 构图
    @classmethod
    def from_workbooks(cls, wb_formula, wb_values) -> "WorkbookGraph":
        g = cls()
        bounds: dict[str, tuple[int, int]] = {}
        for ws in wb_formula.worksheets:
            bounds[ws.title] = (ws.max_row or 1, ws.max_column or 1)
            g.sheets.append(SheetInfo(ws.title, ws.max_row or 1, ws.max_column or 1))

        ws_values = {ws.title: ws for ws in wb_values.worksheets}
        sheet_names = set(bounds)
        g.name_resolver = NameResolver(wb_formula, sheet_names)
        g.name_inventory = g.name_resolver.inventory()

        for ws in wb_formula.worksheets:
            wsv = ws_values.get(ws.title)
            info = next(s for s in g.sheets if s.name == ws.title)
            for row in ws.iter_rows():
                for cell in row:
                    formula = cell.value
                    if not isinstance(formula, str) or not formula.startswith("="):
                        continue
                    coord = cell.coordinate
                    node = node_id(ws.title, coord)
                    g.formulas[node] = formula
                    info.formula_count += 1
                    if wsv is not None:
                        v = wsv[coord].value
                        if v is not None:
                            g.values[node] = str(v)
                    for tok in tokenize_formula(formula):
                        if tok.kind == "ref":
                            g._add_reference(node, formula, tok.ref, ws.title,
                                             bounds, ws_values, via_name=None)
                        else:
                            g._add_name(node, formula, tok.name, ws.title,
                                        bounds, ws_values)
        g._detect_cycles()
        return g

    def _add_reference(
        self,
        node: str,
        formula: str,
        ref: RawRef,
        owner_sheet: str,
        bounds: dict[str, tuple[int, int]],
        ws_values: dict,
        via_name: str | None,
    ) -> int:
        """展开一条直接 / 经名称得到的单元格引用并加入依赖图。

        目标工作表不存在或引用完全越界时只留审计证据，不入图；经名称展开
        的引用同样遵循该规则（无效目标不进入影响面）。返回落在已用范围内
        的目标引用片段数（可能为 0）。
        """
        target_sheet = ref.sheet if ref.sheet is not None else owner_sheet
        if target_sheet not in bounds:
            self.missing_sheet_refs.append(
                MissingSheetRef(node, formula, target_sheet, ref.text, via_name)
            )
            return 0  # 引用不存在的工作表，跳过
        wst = ws_values.get(target_sheet)
        coords, out_coord = self._expand(ref, bounds[target_sheet])
        if out_coord is not None:
            self.out_of_range_refs.append(
                OutOfRangeRef(node, formula, target_sheet, ref.text, out_coord,
                              via_name)
            )
        for tcoord in coords:
            tnode = node_id(target_sheet, tcoord)
            self.precedents.setdefault(node, set()).add(tnode)
            self.dependents.setdefault(tnode, set()).add(node)
            self.edge_count += 1
            # 叶子单元格也记录值，便于前端展示
            if tnode not in self.formulas and tnode not in self.values and wst is not None:
                v = wst[tcoord].value
                if v is not None:
                    self.values[tnode] = str(v)
        return len(coords)

    def _add_name(
        self,
        node: str,
        formula: str,
        raw_name,
        owner_sheet: str,
        bounds: dict[str, tuple[int, int]],
        ws_values: dict,
    ) -> None:
        """解析公式中的定义名称：展开其目标为依赖边，或留下审计线索。"""
        assert self.name_resolver is not None
        res = self.name_resolver.resolve(raw_name, owner_sheet)
        added = 0
        # 名称链展示名：工作表级名称带表前缀，与限定写法一致
        for ref in res.targets:
            target_sheet = ref.sheet or owner_sheet
            if target_sheet not in bounds:
                # 名称体内指向不存在工作表的引用
                self.missing_sheet_refs.append(
                    MissingSheetRef(node, formula, target_sheet, ref.text,
                                    res.text)
                )
                continue
            added += self._add_reference(node, formula, ref, owner_sheet,
                                         bounds, ws_values, via_name=res.text)
        self.name_usages.setdefault(node, []).append(NameUsage(
            source=node,
            name=res.name,
            qualifier=res.qualifier,
            ref_text=res.text,
            status=res.status,
            scope=res.scope,
            definition=res.definition,
            chain=tuple(res.chain),
            target_count=added,
            issues=tuple(res.issues),
        ))

    @staticmethod
    def _expand(
        ref: RawRef, bounds: tuple[int, int]
    ) -> tuple[list[str], str | None]:
        """把 RawRef 展开为 (坐标列表, 第一个越界坐标)。

        显式区域（A1:B10）按原样展开；整行/整列区域按工作表已用范围截断
        （语义上即覆盖全表，越界属正常，不报问题）；显式单元格 / 区域
        超出已用范围时截断到范围内，并返回越界证据坐标。
        超过 MAX_RANGE_CELLS 时等比收缩，防止内存被打爆。
        """
        max_row, max_col = bounds
        out_coord: str | None = None
        if ref.kind == "cell":
            if ref.row1 <= max_row and ref.col1 <= max_col:
                return [f"{get_column_letter(ref.col1)}{ref.row1}"], None
            return [], f"{get_column_letter(ref.col1)}{ref.row1}"
        if ref.kind == "range":
            r1, r2, c1, c2 = ref.row1, ref.row2, ref.col1, ref.col2
            if r2 > max_row or c2 > max_col:
                out_coord = f"{get_column_letter(c2)}{r2}"
                r2, c2 = min(r2, max_row), min(c2, max_col)
            # 整个区域都在已用范围之外
            if r1 > r2 or c1 > c2:
                return [], out_coord
        elif ref.kind == "cols":
            r1, r2 = 1, max_row
            c1, c2 = ref.col1, min(ref.col2, max_col)
        else:  # rows
            r1, r2 = ref.row1, ref.row2
            c1, c2 = 1, max_col
        while (r2 - r1 + 1) * (c2 - c1 + 1) > MAX_RANGE_CELLS and r2 > r1:
            r2 -= 1
        while (r2 - r1 + 1) * (c2 - c1 + 1) > MAX_RANGE_CELLS and c2 > c1:
            c2 -= 1
        coords = [
            f"{get_column_letter(c)}{r}"
            for r in range(r1, r2 + 1)
            for c in range(c1, c2 + 1)
        ]
        return coords, out_coord

    # -------------------------------------------------------------- 环检测
    def _detect_cycles(self) -> None:
        """Tarjan 强连通分量（迭代实现）。大小>1 的分量或自环即循环引用。"""
        adj = self.precedents
        index_of: dict[str, int] = {}
        lowlink: dict[str, int] = {}
        on_stack: set[str] = set()
        stack: list[str] = []
        counter = 0

        for root in list(adj):
            if root in index_of:
                continue
            work = [(root, iter(adj.get(root, ())))]
            index_of[root] = lowlink[root] = counter
            counter += 1
            stack.append(root)
            on_stack.add(root)
            while work:
                node, it = work[-1]
                advanced = False
                for nxt in it:
                    if nxt not in adj and nxt not in index_of:
                        continue  # 非公式叶子节点不参与环
                    if nxt not in index_of:
                        index_of[nxt] = lowlink[nxt] = counter
                        counter += 1
                        stack.append(nxt)
                        on_stack.add(nxt)
                        work.append((nxt, iter(adj.get(nxt, ()))))
                        advanced = True
                        break
                    elif nxt in on_stack:
                        lowlink[node] = min(lowlink[node], index_of[nxt])
                if advanced:
                    continue
                work.pop()
                if work:
                    parent = work[-1][0]
                    lowlink[parent] = min(lowlink[parent], lowlink[node])
                if lowlink[node] == index_of[node]:
                    scc = []
                    while True:
                        w = stack.pop()
                        on_stack.discard(w)
                        scc.append(w)
                        if w == node:
                            break
                    is_cycle = len(scc) > 1 or any(
                        n in adj.get(n, ()) for n in scc
                    )
                    if is_cycle:
                        self.cycles.append(sorted(scc))
                        self.cycle_nodes.update(scc)

    # -------------------------------------------------------------- 查询
    def neighbors(self, node: str, direction: str) -> set[str]:
        adj = self.dependents if direction == "downstream" else self.precedents
        return adj.get(node, set())

    def impact(
        self,
        node: str,
        direction: str = "downstream",
        max_depth: int = 50,
        max_nodes: int = MAX_IMPACT_NODES,
    ) -> dict:
        """BFS 求某单元格的影响面（下游）或依赖链（上游）。

        返回每个可达节点的最短路径；被环截断的节点会标记 in_cycle。
        """
        adj = self.dependents if direction == "downstream" else self.precedents
        depth_of = {node: 0}
        parent = {node: None}
        truncated = False
        queue = deque([node])
        while queue:
            cur = queue.popleft()
            d = depth_of[cur]
            if d >= max_depth:
                continue
            for nxt in sorted(adj.get(cur, ())):
                if nxt in depth_of:
                    continue
                if len(depth_of) >= max_nodes:
                    truncated = True
                    break
                depth_of[nxt] = d + 1
                parent[nxt] = cur
                queue.append(nxt)
            if truncated:
                break

        def path_to(n: str) -> list[str]:
            path = []
            while n is not None:
                path.append(n)
                n = parent[n]
            return list(reversed(path))

        items = [
            {
                "cell": n,
                "sheet": split_node(n)[0],
                "coord": split_node(n)[1],
                "display": display_id(*split_node(n)),
                "depth": depth_of[n],
                "formula": self.formulas.get(n),
                "value": self.values.get(n),
                "in_cycle": n in self.cycle_nodes,
                "path": [display_id(*split_node(p)) for p in path_to(n)],
                "path_raw": path_to(n),
            }
            for n in depth_of
            if n != node
        ]
        items.sort(key=lambda x: (x["depth"], x["cell"]))
        direct = sorted(adj.get(node, set()))
        return {
            "root": node,
            "direction": direction,
            "direct": [
                {
                    "cell": n,
                    "display": display_id(*split_node(n)),
                    "formula": self.formulas.get(n),
                    "value": self.values.get(n),
                    "in_cycle": n in self.cycle_nodes,
                }
                for n in direct
            ],
            "impacted": items,
            "truncated": truncated,
            "total": len(items),
        }

    # ---------------------------------------------------------- 关键节点
    def _node_component_of(self, nodes: set[str]) -> dict[str, str]:
        """把节点映射到分量 id：处于循环 SCC 中的节点映射到同一个分量
        （即把强连通分量压缩为单节点），其余节点各自独立。"""
        comp_of: dict[str, str] = {}
        for cyc in self.cycles:
            members = [n for n in cyc if n in nodes]
            if not members:
                continue
            # 取排序后的首个成员作为分量代表 / id
            cid = min(members)
            for n in members:
                comp_of[n] = cid
        for n in nodes:
            comp_of.setdefault(n, n)
        return comp_of

    @staticmethod
    def _dag_walk(
        nodes: set[str],
        edges: dict[str, set[str]],
        sources: set[str],
    ) -> list[str]:
        """压缩 DAG 上从 sources 出发的 Kahn 拓扑遍历，返回拓扑序。"""
        indeg: dict[str, int] = {c: 0 for c in nodes}
        for c in nodes:
            for t in edges.get(c, ()):
                indeg[t] += 1
        queue = deque(sorted(sources))
        order: list[str] = []
        seen: set[str] = set()
        while queue:
            c = queue.popleft()
            if c in seen:
                continue
            seen.add(c)
            order.append(c)
            for t in sorted(edges.get(c, ())):
                indeg[t] -= 1
                if indeg[t] <= 0:
                    queue.append(t)
        return order

    @staticmethod
    def _bfs_parent(
        starts: set[str],
        nodes: set[str],
        neighbors: dict[str, set[str]],
    ) -> dict[str, str | None]:
        """子图内多源 BFS，返回每个可达节点到最近起点的父指针（起点为 None）。

        父指针构成 BFS 森林；证据路径在 enrich 阶段沿指针回溯，代价随返回行数
        增长而不是随子图规模增长。
        """
        parent: dict[str, str | None] = {s: None for s in starts if s in nodes}
        queue = deque(parent)
        while queue:
            cur = queue.popleft()
            for nxt in neighbors.get(cur, ()):
                if nxt in nodes and nxt not in parent:
                    parent[nxt] = cur
                    queue.append(nxt)
        return parent

    @staticmethod
    def _reconstruct(
        member: str,
        psrc: dict[str, str | None],
        ptgt: dict[str, str | None],
        target: str,
    ) -> list[str]:
        """沿两遍多源 BFS 的父指针还原一条 源 -> member -> 目标 的最短证据路径。

        父指针构成 BFS 森林，沿指针回溯不会成环；member 须同时可达于源集与目标。
        """
        # 前半：member 沿 psrc 退到源点，再反转
        path: list[str] = [member]
        cur = psrc[member]
        while cur is not None:
            path.append(cur)
            cur = psrc[cur]
        path.reverse()
        # 后半：member 沿 ptgt 走到目标
        cur = member
        while cur != target:
            cur = ptgt[cur]
            path.append(cur)
        return path

    def key_nodes(
        self,
        target: str,
        max_nodes: int = MAX_KEYNODE_NODES,
        sheet: str | None = None,
        sort: str = "coverage",
        min_coverage: int = 1,
        limit: int | None = None,
        offset: int = 0,
        evidence_limit: int = MAX_KEYNODE_AFFECTED,
        enrich: bool = True,
    ) -> dict:
        """在目标单元格的上游依赖子图中寻找近似最小割点（关键节点）。

        算法步骤（详见 `_keynode_context`）：上游子图 BFS -> 循环 SCC 压缩为
        单节点 -> 压缩 DAG 上路径数 DP（coverage = in_paths × out_paths）->
        位图闭包求受影响下游 -> 为返回页中的行还原证据路径。

        sheet / sort / min_coverage 为服务端筛选排序；limit=None 表示不分页
        （导出场景，仍对每行的受影响下游清单按 MAX_KEYNODE_AFFECTED_EXPORT
        上限截断）。enrich=False 时只返回轻量计数字段，不展开下游清单与证据。
        """
        ctx = self._keynode_context(target, max_nodes=max_nodes)
        rows = ctx["rows"]
        if sheet is not None:
            rows = [r for r in rows if r["sheet"] == sheet]
        if min_coverage > 1:
            rows = [r for r in rows if r["coverage"] >= min_coverage]
        if sort == "affected":
            rows.sort(key=lambda r: (-r["affected_downstream_count"], r["cell"]))
        elif sort == "cell":
            rows.sort(key=lambda r: r["cell"])
        # sort == "coverage" 时 rows 已在 context 中按覆盖数降序

        total = len(rows)
        page = rows if limit is None else rows[offset : offset + limit]
        if enrich:
            ev_limit = (MAX_KEYNODE_AFFECTED_EXPORT if limit is None
                        else evidence_limit)
            for r in page:
                self._keynode_enrich(ctx, r, ev_limit)
        for i, r in enumerate(rows, 1):
            r["rank"] = i
        return {
            "root": target,
            "root_display": display_id(*split_node(target)),
            "upstream_total": ctx["upstream_total"],
            "component_total": ctx["component_total"],
            "source_total": ctx["source_total"],
            "collapsed_scc_count": ctx["collapsed_scc_count"],
            "path_count_cap": PATH_COUNT_CAP,
            "truncated": ctx["truncated"],
            "filters": {"sheet": sheet, "sort": sort, "min_coverage": min_coverage},
            "total": total,
            "limit": limit,
            "offset": offset,
            "key_nodes": page,
        }

    def _keynode_context(self, target: str, max_nodes: int) -> dict:
        """构造关键节点分析的全部共享中间结果与轻量行（仅计数字段）。

        - 沿 precedents 反向 BFS 得到目标的全部上游依赖（含目标自身）；
        - 将循环强连通分量压缩为单个节点，得到 DAG；
        - DAG 上两遍拓扑路径数 DP：in_paths（源->分量）、out_paths（分量->目标），
          coverage = in_paths × out_paths，即移除该分量后被切断的源->目标路径数，
          路径数超过 PATH_COUNT_CAP 时截断并标记 saturated；
        - 整数位图闭包求受影响下游（C 层 bit_count 计数，清单延迟到 enrich）；
        - 两遍多源 BFS 父指针用于在 enrich 阶段按行还原证据路径（计数字段
          与路径还原解耦，避免对全部节点做昂贵的路径回溯）。
        """
        # 1) 上游依赖子图（沿 precedents）
        sub: set[str] = {target}
        truncated = False
        queue = deque([target])
        while queue:
            cur = queue.popleft()
            for nxt in sorted(self.precedents.get(cur, ())):
                if nxt in sub:
                    continue
                if len(sub) >= max_nodes:
                    truncated = True
                    break
                sub.add(nxt)
                queue.append(nxt)
            if truncated:
                break

        # 2) SCC 压缩
        comp_of = self._node_component_of(sub)
        sorted_comps = sorted(set(comp_of.values()))
        comps = set(sorted_comps)
        comp_index = {c: i for i, c in enumerate(sorted_comps)}
        # 数据流方向的压缩边：prec -> dep（去重自环/平行边），同时记后继位图
        cedges: dict[str, set[str]] = {}
        rev_edges: dict[str, set[str]] = {}
        succ_mask: dict[str, int] = {c: 0 for c in comps}
        for dep in sub:
            cd = comp_of[dep]
            for prec in self.precedents.get(dep, ()):
                if prec not in sub:
                    continue
                cp = comp_of[prec]
                if cp == cd:
                    continue
                if cd not in cedges.setdefault(cp, set()):
                    cedges[cp].add(cd)
                    succ_mask[cp] |= 1 << comp_index[cd]
                rev_edges.setdefault(cd, set()).add(cp)

        ctarget = comp_of[target]
        sources = {c for c in comps if not rev_edges.get(c)}

        members_of: dict[str, list[str]] = {}
        for n in sub:
            members_of.setdefault(comp_of[n], []).append(n)
        for ms in members_of.values():
            ms.sort()

        # 3) DAG 路径数 DP（正向 in_paths / 反向 out_paths）
        order = self._dag_walk(comps, cedges, sources)
        in_paths: dict[str, int] = {}
        saturated: set[str] = set()
        for c in order:
            if c in sources:
                in_paths[c] = 1
                continue
            total = 0
            for p in rev_edges.get(c, ()):
                total += in_paths.get(p, 0)
                if total > PATH_COUNT_CAP:
                    total = PATH_COUNT_CAP
                    saturated.add(c)
                    break
            in_paths[c] = total
        # 反向：从目标分量出发，沿 rev_edges 做拓扑式传播
        out_paths: dict[str, int] = {c: 0 for c in comps}
        out_paths[ctarget] = 1
        for c in reversed(order):
            if c == ctarget:
                continue
            total = 0
            for t in cedges.get(c, ()):
                total += out_paths.get(t, 0)
                if total > PATH_COUNT_CAP:
                    total = PATH_COUNT_CAP
                    saturated.add(c)
                    break
            out_paths[c] = total

        # 4) 证据路径：以成员为必经点的最短 源->成员->目标 单元格路径。
        #    两遍多源 BFS 的父指针构成 BFS 树，沿指针回溯即最短路径。
        src_members: set[str] = set()
        for c in sources:
            if c == ctarget:
                continue
            src_members.update(members_of[c])
        psrc = self._bfs_parent(src_members, sub, self.dependents)
        ptgt = self._bfs_parent({target}, sub, self.precedents)
        # 子图内 node -> 展示 id 一次性预算，供全部行 / 路径直接查表
        sub_display = {n: self.display(n) for n in sub}

        # 5) 受影响下游：移除该分量后，仍在子图内、且不再能到达目标的下游分量
        #    （在压缩 DAG 上等价于该分量沿数据流的后代集合，目标分量自身除外）。
        #    整数位图闭包（逆拓扑 DP）：desc[c] = succ(c) ∪ ⋃desc[t]。
        #    分量计数用 C 层 bit_count（长链 5000 节点也是微秒级）；多成员循环
        #    分量再按成员数加权修正。清单仅逐位取出至多 limit 个成员，
        #    避免在极端长链上做 O(V²) 的位图展开。
        desc_mask: dict[str, int] = {c: 0 for c in comps}
        for c in reversed(order):
            mask = succ_mask.get(c, 0)
            for t in cedges.get(c, ()):
                mask |= desc_mask[t]
            desc_mask[c] = mask
        target_bit = 1 << comp_index[ctarget]

        # 多成员 SCC（真实表格中极少且很小）：按成员数加权的后代单元格计数
        heavy = {c: len(members_of[c]) for c in comps
                 if len(members_of[c]) > 1 and c != ctarget}

        def affected_cell_count(c: str) -> int:
            mask = desc_mask[c] & ~target_bit
            total = mask.bit_count()
            for hc, w in heavy.items():
                if mask & (1 << comp_index[hc]):
                    total += w - 1
            return total

        rows: list[dict] = []
        for c in comps:
            if c == ctarget:
                continue  # 目标自身不计入割点
            coverage = in_paths.get(c, 0) * out_paths.get(c, 0)
            if coverage <= 0:
                continue
            if coverage > PATH_COUNT_CAP:
                coverage = PATH_COUNT_CAP
                saturated.add(c)
            members = members_of[c]
            rep = members[0]
            rep_sheet, rep_coord = split_node(rep)
            rows.append({
                "_comp": c,
                "cell": rep,
                "sheet": rep_sheet,
                "coord": rep_coord,
                "display": sub_display[rep],
                "formula": self.formulas.get(rep),
                "value": self.values.get(rep),
                "in_cycle": rep in self.cycle_nodes,
                "is_component": len(members) > 1,
                "component_members": [
                    {"cell": n, "display": sub_display[n],
                     "in_cycle": n in self.cycle_nodes}
                    for n in members
                ],
                "coverage": coverage,
                "in_paths": in_paths.get(c, 0),
                "out_paths": out_paths.get(c, 0),
                "saturated": c in saturated,
                "affected_downstream_count": affected_cell_count(c),
            })

        rows.sort(key=lambda r: (-r["coverage"], r["cell"]))
        return {
            "target": target,
            "sub": sub,
            "comps": comps,
            "members_of": members_of,
            "sorted_comps": sorted_comps,
            "ctarget": ctarget,
            "sources": sources,
            "desc_mask": desc_mask,
            "target_bit": target_bit,
            "psrc": psrc,
            "ptgt": ptgt,
            "sub_display": sub_display,
            "rows": rows,
            "upstream_total": len(sub) - 1,
            "component_total": len(comps),
            "source_total": len(sources - {ctarget}),
            "collapsed_scc_count": sum(
                1 for ms in members_of.values() if len(ms) > 1),
            "truncated": truncated,
        }

    def _keynode_enrich(self, ctx: dict, row: dict, limit: int) -> None:
        """为一行补上受影响下游清单与一条源->节点->目标的证据路径。

        每行代价与（证据路径长度 + 下游清单条数）成正比，故只在返回页调用，
        避免在长链 / 大子图上对全部节点做昂贵的路径还原。
        """
        c = row.pop("_comp")
        target = ctx["target"]
        members_of = ctx["members_of"]
        sorted_comps = ctx["sorted_comps"]
        sub_display = ctx["sub_display"]
        cycle_nodes = self.cycle_nodes

        mask = ctx["desc_mask"][c] & ~ctx["target_bit"]
        out: list[str] = []
        truncated = False
        while mask:
            lb = mask & -mask
            ms = members_of[sorted_comps[lb.bit_length() - 1]]
            if len(out) + len(ms) > limit:
                out.extend(ms[: max(0, limit - len(out))])
                truncated = True
                break
            out.extend(ms)
            mask ^= lb
        row["affected_downstream"] = [
            {"cell": n, "display": sub_display[n],
             "in_cycle": n in cycle_nodes}
            for n in out
        ]
        if truncated:
            row["affected_truncated"] = True

        members = members_of[c]
        psrc, ptgt = ctx["psrc"], ctx["ptgt"]
        ev_member = next((n for n in members
                          if n in psrc and n in ptgt), members[0])
        raw = self._reconstruct(ev_member, psrc, ptgt, target)
        row["evidence_path"] = [sub_display[p] for p in raw]
        row["evidence_path_raw"] = raw

    # ------------------------------------------------------ 依赖健康审计
    def _edge(self, src: str, dst: str | None = None,
              dst_display: str | None = None, ref: str | None = None) -> dict:
        e: dict = {
            "from": src,
            "from_display": display_id(*split_node(src)) if src else None,
        }
        if dst is not None:
            e["to"] = dst
        if dst_display is not None:
            e["to_display"] = dst_display
        if ref is not None:
            e["ref"] = ref
        return e

    def audit(
        self,
        levels: tuple[str, ...] = AUDIT_LEVELS,
        sheet: str | None = None,
        fanout_threshold: int = DEFAULT_FANOUT_THRESHOLD,
        evidence_limit: int | None = MAX_AUDIT_EVIDENCE,
    ) -> dict:
        """扫描依赖图，生成依赖健康问题清单。

        问题类型：
          missing_sheet   error   跨表公式（含经定义名称展开）指向不存在的工作表
          undefined_name  error   公式使用了名称管理器中不存在（或限定表不存在）的名称
          invalid_name    error   定义名称目标失效（#REF!）、定义为空或存在名称循环
          empty_ref       warning 被公式引用却为空的单元格
          out_of_range    warning 引用超出目标工作表的已用范围
          external_name   warning 定义名称指向外部工作簿，不参与影响面
          high_fanout     warning 公式直接引用的单元格数超过阈值
          orphan_formula  info    未被任何公式引用的公式单元格
        sheet 筛选按问题所属工作表（缺失工作表类按公式所在工作表）。
        """
        issues: list[dict] = []

        def add(issue: dict) -> None:
            if sheet is not None and issue["sheet"] != sheet:
                return
            if issue["level"] not in levels:
                return
            ev = issue["evidence"]
            if evidence_limit is not None and len(ev) > evidence_limit:
                issue["evidence"] = ev[:evidence_limit]
                issue["evidence_truncated"] = True
                issue["evidence_total"] = len(ev)
            issues.append(issue)

        # 1) 跨表引用指向不存在的工作表（同一公式 / 工作表 / 引用 / 名称去重）
        seen: set[tuple] = set()
        for m in self.missing_sheet_refs:
            key = (m.source, m.sheet, m.ref_text, m.via_name)
            if key in seen:
                continue
            seen.add(key)
            via = f"（经定义名称 {m.via_name} 展开）" if m.via_name else ""
            add({
                "code": "missing_sheet",
                "level": "error",
                "sheet": split_node(m.source)[0],
                "cell": m.source,
                "cell_display": display_id(*split_node(m.source)),
                "message": f"引用了不存在的工作表 {m.sheet!r}（{m.ref_text}）{via}",
                "formula": m.formula,
                "target_sheet": m.sheet,
                "target_display": m.ref_text,
                "via_name": m.via_name,
                "evidence": [self._edge(m.source, None, m.ref_text, m.ref_text)],
            })

        # 2) 显式引用超出目标工作表的已用范围
        seen.clear()
        for o in self.out_of_range_refs:
            key = (o.source, o.sheet, o.ref_text, o.via_name)
            if key in seen:
                continue
            seen.add(key)
            tinfo = next((s for s in self.sheets if s.name == o.sheet), None)
            target_display = display_id(o.sheet, o.coord)
            via = f"（经定义名称 {o.via_name} 展开）" if o.via_name else ""
            add({
                "code": "out_of_range",
                "level": "warning",
                "sheet": split_node(o.source)[0],
                "cell": o.source,
                "cell_display": display_id(*split_node(o.source)),
                "message": "引用 {ref} 超出工作表「{s}」已用范围"
                           "（{r} 行 × {c} 列）{via}".format(
                    ref=o.ref_text, s=o.sheet,
                    r=tinfo.max_row if tinfo else "?",
                    c=tinfo.max_col if tinfo else "?", via=via),
                "formula": o.formula,
                "target_sheet": o.sheet,
                "target_display": target_display,
                "via_name": o.via_name,
                "evidence": [
                    self._edge(o.source, node_id(o.sheet, o.coord),
                               target_display, o.ref_text)
                ],
            })

        # 3) 定义名称解析问题（未定义 / 目标无效 / 外部引用）
        #    同一公式中同一名称的同一原因只报一次，证据保留名称与定义。
        seen_name: set[tuple] = set()
        level_of = {
            NAME_UNDEFINED: "error",
            NAME_INVALID: "error",
            NAME_EXTERNAL: "warning",
        }
        code_of = {
            NAME_UNDEFINED: "undefined_name",
            NAME_INVALID: "invalid_name",
            NAME_EXTERNAL: "external_name",
        }
        for node in sorted(self.name_usages):
            for usage in self.name_usages[node]:
                for issue in usage.issues:
                    level = level_of.get(issue.status)
                    if level is None:
                        continue
                    key = (node, usage.ref_text, issue.status, issue.ref_text)
                    if key in seen_name:
                        continue
                    seen_name.add(key)
                    scope_label = ("工作簿级" if usage.scope == "workbook"
                                   else f"工作表「{usage.scope}」级") if usage.scope else None
                    qualifier = usage.qualifier
                    name_display = (display_id(qualifier, usage.name)
                                    if qualifier else usage.name)
                    target_display = issue.ref_text or name_display
                    add({
                        "code": code_of[issue.status],
                        "level": level,
                        "sheet": split_node(node)[0],
                        "cell": node,
                        "cell_display": display_id(*split_node(node)),
                        "message": f"定义名称 {name_display}：{issue.reason}",
                        "formula": self.formulas.get(node),
                        "name": usage.name,
                        "name_qualifier": qualifier,
                        "name_display": name_display,
                        "name_scope": scope_label,
                        "name_definition": usage.definition,
                        "name_status": usage.status,
                        "name_chain": list(usage.chain),
                        "target_sheet": issue.target_sheet,
                        "target_display": target_display,
                        "evidence": [self._edge(node, None, target_display,
                                               usage.ref_text)],
                    })

        # 4) 被公式引用却为空的单元格（无公式、无值；按目标单元格聚合证据边）
        for target in sorted(self.dependents):
            if target in self.formulas or target in self.values:
                continue
            sources = sorted(self.dependents[target])
            s_sheet, s_coord = split_node(target)
            add({
                "code": "empty_ref",
                "level": "warning",
                "sheet": s_sheet,
                "cell": target,
                "cell_display": display_id(s_sheet, s_coord),
                "message": f"单元格 {display_id(s_sheet, s_coord)} 为空，"
                           f"但被 {len(sources)} 个公式引用",
                "formula": None,
                "referenced_by": sources,
                "evidence": [self._edge(s, target) for s in sources],
            })

        # 5) 未被任何公式引用的公式单元格（孤立公式）
        for node in sorted(self.formulas):
            if self.dependents.get(node):
                continue
            n_sheet, n_coord = split_node(node)
            add({
                "code": "orphan_formula",
                "level": "info",
                "sheet": n_sheet,
                "cell": node,
                "cell_display": display_id(n_sheet, n_coord),
                "message": "该公式未被任何其他公式引用，其结果不参与其他计算",
                "formula": self.formulas[node],
                "evidence": [],
            })

        # 6) 扇出过大的公式节点（直接依赖单元格数 > 阈值）
        for node in sorted(self.precedents):
            fanout = len(self.precedents[node])
            if fanout <= fanout_threshold:
                continue
            f_sheet, f_coord = split_node(node)
            add({
                "code": "high_fanout",
                "level": "warning",
                "sheet": f_sheet,
                "cell": node,
                "cell_display": display_id(f_sheet, f_coord),
                "message": f"直接引用 {fanout} 个单元格，超过扇出阈值 "
                           f"{fanout_threshold}",
                "formula": self.formulas[node],
                "fanout": fanout,
                "threshold": fanout_threshold,
                "evidence": [self._edge(node, t)
                             for t in sorted(self.precedents[node])],
            })

        issues.sort(key=lambda x: (
            SEVERITY_WEIGHT[x["level"]], x["code"], x["sheet"], x["cell"]
        ))
        counts_by_level = {lv: 0 for lv in AUDIT_LEVELS}
        counts_by_code: dict[str, int] = {}
        for it in issues:
            counts_by_level[it["level"]] += 1
            counts_by_code[it["code"]] = counts_by_code.get(it["code"], 0) + 1
        return {
            "issues": issues,
            "total": len(issues),
            "counts_by_level": counts_by_level,
            "counts_by_code": counts_by_code,
            "fanout_threshold": fanout_threshold,
            "filters": {
                "levels": list(levels),
                "sheet": sheet,
            },
        }

    # -------------------------------------------------------------- 汇总
    @staticmethod
    def _name_display(usage: NameUsage) -> str:
        if usage.qualifier:
            return display_id(usage.qualifier, usage.name)
        return usage.name

    def _name_usage_json(self, usage: NameUsage) -> dict:
        """单个名称使用留痕的 JSON 形态（含解析状态与证据）。"""
        issues = [
            {
                "status": i.status,
                "reason": i.reason,
                "ref": i.ref_text,
                "target_sheet": i.target_sheet,
            }
            for i in usage.issues
        ]
        return {
            "name": usage.name,
            "qualifier": usage.qualifier,
            "name_display": self._name_display(usage),
            "ref_text": usage.ref_text,
            "status": usage.status,
            "scope": "工作簿级" if usage.scope == "workbook"
                     else (f"工作表「{usage.scope}」级" if usage.scope else None),
            "scope_id": usage.scope,
            "definition": usage.definition,
            "chain": list(usage.chain),
            "target_count": usage.target_count,
            "issues": issues,
        }

    def _names_of(self, node: str) -> list[dict]:
        return [self._name_usage_json(u)
                for u in self.name_usages.get(node, ())]

    def _names_summary(self) -> dict:
        local = sum(1 for d in self.name_inventory if d.local)
        broken = 0
        for usages in self.name_usages.values():
            for u in usages:
                if u.status in (NAME_UNDEFINED, NAME_INVALID):
                    broken += 1
        return {
            "defined_name_count": len(self.name_inventory),
            "workbook_scoped_count": len(self.name_inventory) - local,
            "sheet_scoped_count": local,
            "name_usage_count": sum(len(v) for v in self.name_usages.values()),
            "name_issue_usage_count": broken,
        }

    def summary(self) -> dict:
        return {
            "sheets": [
                {
                    "name": s.name,
                    "max_row": s.max_row,
                    "max_col": s.max_col,
                    "formula_count": s.formula_count,
                }
                for s in self.sheets
            ],
            "formula_count": len(self.formulas),
            "edge_count": self.edge_count,
            "cycle_count": len(self.cycles),
            "cycles": [
                [display_id(*split_node(n)) for n in cyc] for cyc in self.cycles
            ],
            **self._names_summary(),
        }

    def cell_list(
        self,
        sheet: str | None = None,
        query: str = "",
        formulas_only: bool = True,
        limit: int = 1000,
        offset: int = 0,
    ) -> dict:
        nodes = set(self.formulas)
        if not formulas_only:
            nodes |= set(self.dependents)  # 被引用的非公式单元格
        rows = []
        q = query.lower()
        for n in sorted(nodes):
            s, coord = split_node(n)
            if sheet and s != sheet:
                continue
            disp = display_id(s, coord)
            names = self._names_of(n)
            haystack = disp.lower() + " " + (self.formulas.get(n) or "").lower()
            if names:
                haystack += " " + " ".join(
                    (x["name_display"] + " " + (x["definition"] or "")).lower()
                    for x in names)
            if q and q not in haystack:
                continue
            rows.append(
                {
                    "cell": n,
                    "sheet": s,
                    "coord": coord,
                    "display": disp,
                    "formula": self.formulas.get(n),
                    "value": self.values.get(n),
                    "in_cycle": n in self.cycle_nodes,
                    "dependent_count": len(self.dependents.get(n, ())),
                    "precedent_count": len(self.precedents.get(n, ())),
                    "names": names,
                }
            )
        return {"total": len(rows), "cells": rows[offset : offset + limit]}

    def name_inventory_json(self) -> list[dict]:
        """名称管理器清单（含静态作用域与动态解析状态汇总），供导出。"""
        # 汇总每个定义名称被公式使用的状态
        usage_stats: dict[tuple[str, str], dict[str, int]] = {}
        for usages in self.name_usages.values():
            for u in usages:
                key = (u.name, u.scope or "")
                st = usage_stats.setdefault(key, {})
                st[u.status] = st.get(u.status, 0) + 1
        out = []
        for d in self.name_inventory:
            stats = usage_stats.get((d.name, d.scope), {})
            out.append({
                "name": d.name,
                "scope": "workbook" if not d.local else "sheet",
                "scope_sheet": None if not d.local else d.scope,
                "definition": d.attr_text,
                "hidden": d.hidden,
                "comment": d.comment,
                "usage_status_counts": stats,
                "used_in_formulas": sum(stats.values()),
            })
        return out

    def name_usages_json(self) -> list[dict]:
        """全部公式中的名称使用留痕（名称、状态、定义、链、目标数、问题）。"""
        rows: list[dict] = []
        for node in sorted(self.name_usages):
            for u in self.name_usages[node]:
                row = {
                    "cell": node,
                    "cell_display": display_id(*split_node(node)),
                    **self._name_usage_json(u),
                }
                rows.append(row)
        return rows

    def export(self, filename: str) -> dict:
        """完整分析结果，供导出 JSON。"""
        edges = [
            {"from": src, "to": dst}
            for src, dsts in sorted(self.dependents.items())
            for dst in sorted(dsts)
        ]
        return {
            "filename": filename,
            "summary": self.summary(),
            "defined_names": self.name_inventory_json(),
            "name_usages": self.name_usages_json(),
            "formulas": {
                display_id(*split_node(n)): {
                    "formula": f,
                    "value": self.values.get(n),
                    "in_cycle": n in self.cycle_nodes,
                    "names": self._names_of(n),
                }
                for n, f in sorted(self.formulas.items())
            },
            "edges": [
                {"from": display_id(*split_node(e["from"])), "to": display_id(*split_node(e["to"]))}
                for e in edges
            ],
            "cycles": [
                [display_id(*split_node(n)) for n in cyc] for cyc in self.cycles
            ],
            "audit": self.audit(evidence_limit=None),
        }
