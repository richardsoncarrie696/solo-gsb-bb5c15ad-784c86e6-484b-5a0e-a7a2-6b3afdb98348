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
# 关键节点分析：枚举「输入源 -> 目标」依赖路径的默认 / 上限条数（近似割点用）
DEFAULT_KEY_PATHS = 200
MAX_KEY_PATHS_QUERY = 1000
# 路径枚举时每层保留的部分路径上限与总扩展次数上限（防组合爆炸打爆 CPU）
MAX_KEY_FRONTIER = 20000
MAX_KEY_EXTENSIONS = 1_000_000
# 关键节点接口默认返回行数（导出不受限）；每行受影响下游示例数
DEFAULT_KEY_NODE_LIMIT = 500
KEY_NODE_DOWNSTREAM_SAMPLE = 100

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

    # ------------------------------------------------------ 关键节点分析
    @staticmethod
    def _scc_on(nodes: set[str], adj: dict[str, set[str]]) -> list[list[str]]:
        """在给定节点子集与邻接表上跑迭代版 Tarjan，返回强连通分量列表。

        指向子集外的边一律忽略（等价于先取诱导子图）。非公式叶子不在 adj
        的键中，预先作为独立起点登记，避免被漏掉而成为单节点 SCC。
        """
        index_of: dict[str, int] = {}
        lowlink: dict[str, int] = {}
        on_stack: set[str] = set()
        stack: list[str] = []
        sccs: list[list[str]] = []
        counter = 0

        for root in sorted(nodes):
            if root in index_of:
                continue
            work = [(root, iter(sorted(adj.get(root, set()) & nodes)))]
            index_of[root] = lowlink[root] = counter
            counter += 1
            stack.append(root)
            on_stack.add(root)
            while work:
                node, it = work[-1]
                advanced = False
                for nxt in it:
                    if nxt not in nodes:
                        continue
                    if nxt not in index_of:
                        index_of[nxt] = lowlink[nxt] = counter
                        counter += 1
                        stack.append(nxt)
                        on_stack.add(nxt)
                        work.append((nxt, iter(sorted(adj.get(nxt, set()) & nodes))))
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
                    sccs.append(scc)
        return sccs

    def key_nodes(
        self,
        node: str,
        sheet: str | None = None,
        sort: str = "coverage",
        min_coverage: int = 1,
        limit: int | None = DEFAULT_KEY_NODE_LIMIT,
        max_paths: int = DEFAULT_KEY_PATHS,
    ) -> dict:
        """在上游依赖子图中计算关键节点（近似最小割点分析）。

        步骤：
          1. 反向遍历求目标单元格的全部上游可达节点，得到上游诱导子图；
          2. Tarjan 求强连通分量并把每个循环分量压缩为单节点，压缩图为 DAG；
          3. DAG 路径计数 DP：in_paths[v]（源 -> v）与 out_paths[v]
             （v -> 目标），二者乘积即经过 v 的独立依赖路径总数（精确值，
             Python 大整数），覆盖率 100% 即必经节点（精确割点）；
          4. 位集 DP 求每个分量可达的受影响下游；广度枚举前 max_paths 条
             完整路径作为证据样本，未覆盖到的节点用 DP 父链拼出一条经它
             到目标的证据路径兜底。

        「近似」体现在以经过节点的源→目标路径数衡量割点重要性（而非直接求
        全局最小点割集，后者为 NP 难）；单节点必经关系则是精确的。
        """
        # 1) 上游可达诱导子图（目标自身含在 reach 内）
        reach: set[str] = {node}
        stack_u = [node]
        upstream_truncated = False
        while stack_u:
            cur = stack_u.pop()
            for pre in self.precedents.get(cur, ()):  # set 迭代，规模有界
                if pre in reach:
                    continue
                if len(reach) >= MAX_IMPACT_NODES:
                    upstream_truncated = True
                    break
                reach.add(pre)
                stack_u.append(pre)
            if upstream_truncated:
                break

        # 2) 子图上的 SCC（与全局环检测结果一致，但限制在子图范围内）
        sccs = self._scc_on(reach, self.precedents)
        comp_of: dict[str, int] = {}
        members: list[list[str]] = []
        for scc in sccs:
            ci = len(members)
            ms = sorted(scc)
            members.append(ms)
            for m in ms:
                comp_of[m] = ci
        # 大小>1 或存在指向自身分量的边（自环）即循环分量
        cyclic = [
            len(ms) > 1
            or any(t in reach and comp_of[t] == ci
                   for t in self.precedents.get(ms[0], ()))
            for ci, ms in enumerate(members)
        ]
        target_ci = comp_of[node]

        # 3) 压缩 DAG：pred[ci] 为直接上游分量，cons[ci] 为直接下游分量。
        #    非公式叶子不是 precedents 的键，边的收集要覆盖 reach 中全部节点。
        pred: list[set[int]] = [set() for _ in members]
        cons: list[set[int]] = [set() for _ in members]
        for u in reach:
            cu = comp_of[u]
            for v in self.precedents.get(u, ()):
                if v not in reach:
                    continue
                cv = comp_of[v]
                if cv != cu:
                    pred[cu].add(cv)
                    cons[cv].add(cu)
        source_cis = [ci for ci in range(len(members))
                      if not pred[ci] and ci != target_ci]

        # 4) DAG 路径计数（精确值，Python 大整数）：
        #    in_paths[ci]  = 任一输入源 -> ci 的路径数（源为 1）
        #    out_paths[ci] = ci -> 目标的路径数（目标为 1；上游子图内每个节点
        #                    都可达目标，故该值恒 >= 1）
        #    经过 ci 的完整依赖路径数 = in_paths[ci] * out_paths[ci]
        ncomp_t = len(members)
        topo: list[int] = []  # 源 -> 目标 的拓扑序
        indeg = [len(pred[ci]) for ci in range(ncomp_t)]
        q = deque(ci for ci in range(ncomp_t) if indeg[ci] == 0)
        while q:
            cur = q.popleft()
            topo.append(cur)
            for child in sorted(cons[cur]):
                indeg[child] -= 1
                if indeg[child] == 0:
                    q.append(child)

        in_paths = [0] * ncomp_t
        for ci in topo:
            if ci == target_ci:
                # 目标无上游时路径数为 0（没有「源 -> 目标」的路径）；
                # 有上游时由前驱的 in_paths 之和给出
                in_paths[ci] = (sum(in_paths[p] for p in pred[ci])
                                if pred[ci] else 0)
            elif not pred[ci]:
                in_paths[ci] = 1
            else:
                in_paths[ci] = sum(in_paths[p] for p in pred[ci])
        out_paths = [0] * ncomp_t
        for ci in reversed(topo):
            out_paths[ci] = 1 if ci == target_ci else sum(
                out_paths[c] for c in cons[ci])
        total_paths = in_paths[target_ci]
        coverage = [in_paths[ci] * out_paths[ci] for ci in range(ncomp_t)]

        # 深度：距目标的最长链层数（目标 0）
        depth_of: dict[int, int] = {ci: 0 for ci in range(ncomp_t)}
        for ci in reversed(topo):
            if cons[ci]:
                depth_of[ci] = 1 + max(depth_of[c] for c in cons[ci])

        # 受影响下游：DAG 位集 DP，一次求出每个分量可达的下游分量集合
        desc_bits = [0] * ncomp_t
        for ci in reversed(topo):
            bits = 0
            for child in cons[ci]:
                bits |= 1 << child
                bits |= desc_bits[child]
            desc_bits[ci] = bits
        member_count = [len(ms) for ms in members]

        def downstream_of(ci: int) -> tuple[list[str], int]:
            """返回 (受影响下游单元格（已排序）, 总数)。

            受影响下游 = 子图内该分量可达的全部分量（含目标自身——割点被移除
            后正是目标的值无法算出；不含该分量自身）。
            """
            bits = desc_bits[ci]
            total = sum(member_count[cj] for cj in range(ncomp_t)
                        if bits & (1 << cj))
            out: list[str] = []
            b = bits
            while b:
                low = b & -b
                cj = low.bit_length() - 1
                out.extend(members[cj])
                b -= low
            out.sort()
            return out, total

        # 5) 证据路径：先按层广度枚举至多 max_paths 条完整路径，取每个分量
        #    第一次出现的路径作为多样化证据；未被枚举覆盖的分量（路径组合
        #    爆炸时必然存在）用 DP 父链直接拼出一条经过它的路径兜底。
        evidence: dict[int, list[int]] = {}
        sample_paths: list[list[int]] = []
        frontier: list[list[int]] = [[s] for s in source_cis]
        extensions = 0
        path_truncated = False
        while frontier:
            nxt_frontier: list[list[int]] = []
            stopped = False
            for p in frontier:
                for child in sorted(cons[p[-1]]):
                    if child in p:
                        continue
                    extensions += 1
                    np = p + [child]
                    if child == target_ci:
                        if len(sample_paths) < max_paths:
                            sample_paths.append(np)
                            for cj in np:
                                evidence.setdefault(cj, np)
                        if len(sample_paths) >= max_paths:
                            stopped = True
                            break
                    else:
                        nxt_frontier.append(np)
                if stopped:
                    break
                if extensions > MAX_KEY_EXTENSIONS:
                    path_truncated = True
                    stopped = True
                    break
            if stopped or extensions > MAX_KEY_EXTENSIONS:
                if len(sample_paths) >= max_paths:
                    path_truncated = True
                break
            if len(nxt_frontier) > MAX_KEY_FRONTIER:
                path_truncated = True
            frontier = nxt_frontier[:MAX_KEY_FRONTIER]
        if total_paths > len(sample_paths):
            path_truncated = True

        # 兜底父链：任取一条通向源（反向）与通向目标（正向）的链
        chain_to_source: dict[int, int | None] = {}
        for ci in topo:
            ps = pred[ci]
            chain_to_source[ci] = (max(ps, key=lambda p: in_paths[p])
                                   if ps else None)
        chain_to_target: dict[int, int | None] = {}
        for ci in reversed(topo):
            cs = cons[ci]
            chain_to_target[ci] = (min(cs, key=lambda c: out_paths[c])
                                   if cs else None)

        def evidence_for(ci: int) -> list[int]:
            if ci in evidence:
                return evidence[ci]
            prefix: list[int] = []
            cur: int | None = ci
            seen_p = set()
            while cur is not None and cur not in seen_p:
                seen_p.add(cur)
                prefix.append(cur)
                cur = chain_to_source[cur]
            prefix.reverse()
            suffix: list[int] = []
            cur = chain_to_target[ci]
            seen_s = {ci}
            while cur is not None and cur not in seen_s:
                seen_s.add(cur)
                suffix.append(cur)
                cur = chain_to_target[cur]
            return prefix + suffix

        def row_for(ci: int, with_detail: bool) -> dict | None:
            ms = members[ci]
            is_scc = cyclic[ci] or len(ms) > 1
            node_sheet = split_node(ms[0])[0]
            row: dict = {
                "cell": None if is_scc else ms[0],
                "type": "scc" if is_scc else "cell",
                "node_id": f"scc:{ci}",
                "sheet": node_sheet if is_scc else split_node(ms[0])[0],
                "coord": None if is_scc else split_node(ms[0])[1],
                "display": ("🔁 循环分量（" + "、".join(
                    display_id(*split_node(m)) for m in ms[:5])
                    + ("…" if len(ms) > 5 else "") + "）") if is_scc
                    else display_id(*split_node(ms[0])),
                "members": ms if is_scc else [ms[0]],
                "member_displays": [display_id(*split_node(m)) for m in ms],
                "in_cycle": bool(is_scc),
                "formula": None if is_scc else self.formulas.get(ms[0]),
                "value": None if is_scc else self.values.get(ms[0]),
                "depth": depth_of.get(ci, 0),
                "covered_paths": coverage[ci],
                "covered_paths_text": str(coverage[ci]),
                "coverage_ratio": (coverage[ci] / total_paths)
                                  if total_paths else 0.0,
                "cuts_all_paths": bool(total_paths and coverage[ci] == total_paths),
            }
            if with_detail:
                downstream, aff_total = (
                    downstream_of(ci) if coverage[ci] else ([], 0))
                ev = evidence_for(ci) if coverage[ci] else None
                row.update({
                    "affected_downstream_count": aff_total,
                    "affected_downstream": downstream[:KEY_NODE_DOWNSTREAM_SAMPLE],
                    "affected_downstream_display": [
                        display_id(*split_node(d))
                        for d in downstream[:KEY_NODE_DOWNSTREAM_SAMPLE]
                    ],
                    "affected_downstream_truncated":
                        aff_total > KEY_NODE_DOWNSTREAM_SAMPLE,
                    "evidence_path": None if ev is None else [
                        display_id(*split_node(members[ci2][0])) for ci2 in ev
                    ],
                    "evidence_path_raw": None if ev is None else [
                        members[ci2][0] for ci2 in ev
                    ],
                    "evidence_components": None if ev is None else [
                        {
                            "node_id": f"scc:{ci2}",
                            "type": "scc" if (cyclic[ci2] or
                                              len(members[ci2]) > 1) else "cell",
                            "displays": [display_id(*split_node(m))
                                         for m in members[ci2]],
                        }
                        for ci2 in ev
                    ],
                })
            return row

        # 候选：目标自身以外、被至少一条路径覆盖的分量
        candidates = [
            ci for ci in range(len(members))
            if ci != target_ci and coverage[ci] >= min_coverage
        ]
        if sheet is not None:
            candidates = [ci for ci in candidates
                          if any(split_node(m)[0] == sheet for m in members[ci])]
        if sort == "depth":
            candidates.sort(key=lambda ci: (-depth_of.get(ci, 0),
                                            -coverage[ci], members[ci][0]))
        elif sort == "affected":
            def aff_count(ci: int) -> int:
                return sum(member_count[cj] for cj in range(ncomp_t)
                           if desc_bits[ci] & (1 << cj))
            candidates.sort(key=lambda ci: (-aff_count(ci),
                                            -coverage[ci], members[ci][0]))
        else:
            candidates.sort(key=lambda ci: (-coverage[ci],
                                            -depth_of.get(ci, 0), members[ci][0]))
        total_nodes = len(candidates)
        sliced = candidates if limit is None else candidates[:limit]
        rows = [r for ci in sliced if (r := row_for(ci, with_detail=True))]

        return {
            "root": node,
            "upstream_node_count": len(reach) - 1,
            "scc_count": sum(1 for ci in range(len(members))
                             if cyclic[ci] or len(members[ci]) > 1),
            "source_count": len(source_cis),
            "total_paths": total_paths,
            "total_paths_text": str(total_paths),
            "sampled_paths": len(sample_paths),
            "path_truncated": path_truncated,
            "upstream_truncated": upstream_truncated,
            "max_paths": max_paths,
            "total": total_nodes,
            "returned": len(rows),
            "filters": {"sheet": sheet, "sort": sort,
                        "min_coverage": min_coverage},
            "key_nodes": rows,
            # 目标若落在循环分量中，给出提示所需信息
            "root_in_cycle": node in self.cycle_nodes,
            "root_component_members": members[target_ci]
                                      if (cyclic[target_ci]
                                          or len(members[target_ci]) > 1)
                                      else [node],
        }

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
