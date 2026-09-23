"""定义名称（名称管理器）解析、作用域、影响面与审计测试。

运行：python -m pytest tests/test_names.py -v
"""
import io

from openpyxl import Workbook, load_workbook
from openpyxl.workbook.defined_name import DefinedName

from app.formula_parser import parse_formula, parse_names
from app.graph import WorkbookGraph


def _load(wb: Workbook) -> WorkbookGraph:
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    wbf = load_workbook(buf, data_only=False)
    buf.seek(0)
    wbv = load_workbook(buf, data_only=True)
    return WorkbookGraph.from_workbooks(wbf, wbv)


def _global(wb, name, value):
    wb.defined_names[name] = DefinedName(name, attr_text=value)


def _local(ws, name, value):
    ws.defined_names[name] = DefinedName(name, attr_text=value,
                                         localSheetId=ws.parent.sheetnames.index(ws.title))


# ------------------------------------------------------------ 词法：名称 token
def test_parser_bare_name():
    names = parse_names("=单价*数量")
    assert [n.name for n in names] == ["单价", "数量"]
    assert all(n.qualifier is None for n in names)


def test_parser_qualified_name():
    n = parse_names("=计算!单价*2")[0]
    assert n.name == "单价" and n.qualifier == "计算"
    n = parse_names("='月报 2026'!税率")[0]
    assert n.name == "税率" and n.qualifier == "月报 2026"


def test_parser_name_not_cell_or_function():
    # 函数名带括号不算名称；TRUE/FALSE 为保留字；普通引用仍是引用
    assert [n.name for n in parse_names("=LOG10(单价)+SUM(B1:B2)")] == ["单价"]
    assert parse_names("=TRUE+FALSE") == []
    refs = parse_formula("=单价+A1")
    assert [(r.col1, r.row1) for r in refs] == [(1, 1)]


def test_parser_scientific_notation_still_not_name():
    assert parse_names("=1E2+2E3") == []
    assert [n.name for n in parse_names("=2E3+税率")] == ["税率"]


def test_parser_cell_like_name_with_underscore():
    # A1_NAME 是合法定义名称，不应被拆成单元格 A1 + 文本
    n = parse_names("=A1_NAME*2")
    assert len(n) == 1 and n[0].name == "A1_NAME"
    assert parse_formula("=A1_NAME*2") == []
    # 限定表 + 名称
    n = parse_names("=S1!A1_NAME")[0]
    assert n.name == "A1_NAME" and n.qualifier == "S1"


def test_parser_structured_reference_ignored():
    assert parse_names("=SUM(表1[数量])") == []
    assert parse_names("=表1[[#汇总],[数量]]") == []


# ------------------------------------------------------------ 作用域解析
def _scope_wb():
    wb = Workbook()
    p = wb.active
    p.title = "参数"
    p["B2"] = 10
    p["B3"] = 5
    calc = wb.create_sheet("计算")
    summ = wb.create_sheet("汇总")
    _global(wb, "单价", "参数!$B$2")
    _local(calc, "单价", "参数!$B$3")   # 同名工作表级遮蔽工作簿级
    calc["B1"] = "=单价"               # -> 本地 单价 -> 参数!B3
    summ["B1"] = "=单价"               # -> 全局 单价 -> 参数!B2
    summ["B2"] = "=计算!单价"          # 限定到计算表 -> 其本地名 -> 参数!B3
    return wb


def test_same_name_sheet_scope_shadows_workbook_scope():
    g = _load(_scope_wb())
    # 计算!B1 经本地名称指向 参数!B3
    assert "计算!B1" in g.dependents["参数!B3"]
    assert "计算!B1" not in g.dependents.get("参数!B2", set())
    # 汇总!B1 无本地名，回退工作簿级，指向 参数!B2
    assert "汇总!B1" in g.dependents["参数!B2"]
    # 汇总!B2 用跨表限定，命中计算表的本地名
    assert "汇总!B2" in g.dependents["参数!B3"]


def test_name_usage_records_scope_and_status():
    g = _load(_scope_wb())
    u_calc = next(u for u in g.name_usages["计算!B1"] if u.name == "单价")
    assert u_calc.status == "resolved" and u_calc.scope == "计算"
    u_summ = next(u for u in g.name_usages["汇总!B1"] if u.name == "单价")
    assert u_summ.scope == "workbook"
    inv = {(d.name, d.scope) for d in g.name_inventory}
    assert ("单价", "workbook") in inv and ("单价", "计算") in inv


# ------------------------------------------------------------ 名称展开为依赖边
def test_workbook_scoped_cell_and_range_names_expand_to_edges():
    wb = Workbook()
    d = wb.active
    d.title = "数据"
    for i in range(1, 6):
        d.cell(row=i, column=2, value=i)
    c = wb.create_sheet("计算")
    c["B1"] = "=单个"
    c["B2"] = "=SUM(区域名)"
    _global(wb, "单个", "数据!$B$1")
    _global(wb, "区域名", "数据!$B$1:$B$5")
    g = _load(wb)
    assert g.dependents["数据!B1"] == {"计算!B1", "计算!B2"}
    assert "数据!B5" in g.dependents
    assert g.dependents["数据!B5"] == {"计算!B2"}
    u = next(u for u in g.name_usages["计算!B2"] if u.name == "区域名")
    assert u.status == "resolved" and u.target_count == 5


def test_nested_and_multi_area_names():
    wb = Workbook()
    d = wb.active
    d.title = "数据"
    d["A1"] = 1; d["A2"] = 2; d["C1"] = 9
    c = wb.create_sheet("计算")
    c["B1"] = "=链名"
    c["B2"] = "=多区"
    _global(wb, "链名", "中转")
    _global(wb, "中转", "数据!$A$1:$A$2")
    _global(wb, "多区", "数据!$A$1:$A$2,数据!$C$1")
    g = _load(wb)
    assert "计算!B1" in g.dependents["数据!A1"]
    assert "计算!B2" in g.dependents["数据!A1"]
    assert "计算!B2" in g.dependents["数据!C1"]
    chain = next(u for u in g.name_usages["计算!B1"] if u.name == "链名").chain
    assert tuple(chain) == ("链名", "中转")


def test_constant_and_formula_names_produce_no_edges_but_keep_status():
    wb = Workbook()
    d = wb.active
    d.title = "数据"
    d["A1"] = 1
    c = wb.create_sheet("计算")
    c["B1"] = "=税率"                    # 纯常量
    c["B2"] = "=表达式名"                # 公式表达式（含引用）
    c["B3"] = "=纯函数名"                # 函数调用
    _global(wb, "税率", "0.13")
    _global(wb, "表达式名", "数据!$A$1+1")
    _global(wb, "纯函数名", "SUM(数据!$A$1)")
    g = _load(wb)
    u1 = g.name_usages["计算!B1"][0]
    assert u1.status == "constant" and u1.target_count == 0
    assert "计算!B1" not in g.dependents.get("数据!A1", set())
    u2 = g.name_usages["计算!B2"][0]
    assert u2.status == "formula" and u2.target_count == 1
    assert "计算!B2" in g.dependents["数据!A1"]
    u3 = g.name_usages["计算!B3"][0]
    assert u3.status == "formula" and u3.target_count == 1


def test_name_target_feeds_impact_analysis():
    wb = Workbook()
    d = wb.active
    d.title = "数据"
    d["B1"] = 100
    c = wb.create_sheet("计算")
    c["B1"] = "=源名称*2"
    c["B2"] = "=B1+1"
    _global(wb, "源名称", "数据!$B$1")
    g = _load(wb)
    res = g.impact("数据!B1", direction="downstream")
    cells = {i["cell"] for i in res["impacted"]}
    assert {"计算!B1", "计算!B2"} <= cells


def test_name_based_cycle_does_not_crash_and_invalid_targets_excluded():
    wb = Workbook()
    s = wb.active
    s.title = "S"
    s["A1"] = "=环A"
    _global(wb, "环A", "环B")
    _global(wb, "环B", "环A")
    g = _load(wb)
    u = g.name_usages["S!A1"][0]
    assert u.status == "invalid" and u.target_count == 0
    assert g.edge_count == 0


# ------------------------------------------------------------ 审计：名称问题
def test_undefined_name_is_error_and_localizable():
    wb = Workbook()
    c = wb.active
    c.title = "计算"
    c["B1"] = "=幽灵名称+1"
    g = _load(wb)
    res = g.audit()
    issue = next(i for i in res["issues"] if i["code"] == "undefined_name")
    assert issue["level"] == "error"
    assert issue["cell"] == "计算!B1"
    assert "幽灵名称" in issue["message"]
    assert issue["evidence"][0]["from"] == "计算!B1"
    assert issue["evidence"][0]["to_display"] == "幽灵名称"
    assert "to" not in issue["evidence"][0]


def test_undefined_qualified_name_missing_sheet():
    wb = Workbook()
    c = wb.active
    c.title = "计算"
    c["B1"] = "=不存在的表!某名称"
    g = _load(wb)
    issue = next(i for i in g.audit()["issues"] if i["code"] == "undefined_name")
    assert issue["level"] == "error"
    assert issue["target_sheet"] == "不存在的表"
    assert "不存在的表" in issue["message"]
    # 无效名称目标不产生任何边
    assert g.edge_count == 0


def test_invalid_name_ref_error_and_empty_definition():
    wb = Workbook()
    c = wb.active
    c.title = "计算"
    c["B1"] = "=已删名称"
    c["B2"] = "=空名称"
    _global(wb, "已删名称", "#REF!")
    wb.defined_names["空名称"] = DefinedName("空名称")  # attr_text=None
    g = _load(wb)
    codes = {i["code"] for i in g.audit()["issues"]}
    assert "invalid_name" in codes
    for u in g.name_usages["计算!B1"] + g.name_usages["计算!B2"]:
        assert u.status == "invalid" and u.target_count == 0
    assert g.edge_count == 0


def test_name_pointing_to_missing_sheet_is_missing_sheet_error():
    wb = Workbook()
    c = wb.active
    c.title = "计算"
    c["B1"] = "=坏表名称+1"
    _global(wb, "坏表名称", "幽灵表!$A$1")
    g = _load(wb)
    issue = next(i for i in g.audit()["issues"] if i["code"] == "missing_sheet")
    assert issue["level"] == "error"
    assert issue["target_sheet"] == "幽灵表"
    assert issue["via_name"] == "坏表名称"
    assert "坏表名称" in issue["message"]
    # 无效目标不在影响面
    assert "幽灵表!A1" not in g.dependents


def test_external_name_is_warning_not_error():
    wb = Workbook()
    c = wb.active
    c.title = "计算"
    c["B1"] = "=外部名"
    _global(wb, "外部名", "[1]预算.xlsx!Sheet1!$A$1")
    g = _load(wb)
    res = g.audit()
    ext = next(i for i in res["issues"] if i["code"] == "external_name")
    assert ext["level"] == "warning"
    errors = {i["code"] for i in res["issues"] if i["level"] == "error"}
    assert errors == set()
    assert g.edge_count == 0


def test_name_issues_dedup_per_formula():
    wb = Workbook()
    c = wb.active
    c.title = "计算"
    c["B1"] = "=幽灵+幽灵"
    g = _load(wb)
    issues = [i for i in g.audit()["issues"] if i["code"] == "undefined_name"]
    assert len(issues) == 1


def test_out_of_range_through_name_recorded_with_via_name():
    wb = Workbook()
    d = wb.active
    d.title = "数据"
    d["A1"] = 1
    c = wb.create_sheet("计算")
    c["B1"] = "=越界名"
    _global(wb, "越界名", "数据!$A$99")
    g = _load(wb)
    issue = next(i for i in g.audit()["issues"] if i["code"] == "out_of_range")
    assert issue["via_name"] == "越界名"
    assert "越界名" in issue["message"]
    # 越界坐标不入图
    assert "数据!A99" not in g.dependents


def test_named_range_counts_toward_fanout_and_empty_audit():
    wb = Workbook()
    d = wb.active
    d.title = "数据"
    for i in range(1, 25):
        d.cell(row=i, column=1, value=i)
    d["B25"] = "标记"  # 使第 25 行落在已用范围内，A25 本身留空
    c = wb.create_sheet("计算")
    c["B1"] = "=SUM(二十五区)"
    _global(wb, "二十五区", "数据!$A$1:$A$25")
    g = _load(wb)
    res = g.audit(fanout_threshold=20)
    assert "high_fanout" in res["counts_by_code"]
    # 名称展开的空单元格同样触发 empty_ref
    empties = {i["cell"] for i in res["issues"] if i["code"] == "empty_ref"}
    assert "数据!A25" in empties


# ------------------------------------------------------------ 导出保留名称
def test_full_export_keeps_names_status_and_evidence():
    wb = Workbook()
    d = wb.active
    d.title = "数据"
    d["A1"] = 1
    c = wb.create_sheet("计算")
    c["B1"] = "=好名称"
    c["B2"] = "=坏名称"
    _global(wb, "好名称", "数据!$A$1")
    _global(wb, "坏名称", "#REF!")
    data = _load(wb).export("t.xlsx")
    names = {n["name"]: n for n in data["defined_names"]}
    assert set(names) == {"好名称", "坏名称"}
    assert names["好名称"]["scope"] == "workbook"
    usages = {u["ref_text"]: u for u in data["name_usages"]}
    assert usages["好名称"]["status"] == "resolved"
    assert usages["好名称"]["chain"] == ["好名称"]
    assert usages["坏名称"]["status"] == "invalid"
    assert usages["坏名称"]["issues"] and usages["坏名称"]["issues"][0]["status"] == "invalid"
    # 公式段也带名称留痕
    f = data["formulas"]["计算!B1"]
    assert f["names"][0]["name"] == "好名称"
    codes = {i["code"] for i in data["audit"]["issues"]}
    assert "invalid_name" in codes


def test_summary_reports_name_counts():
    wb = Workbook()
    c = wb.active
    c.title = "计算"
    c["B1"] = "=单价"
    _global(wb, "单价", "0.1")
    _local(c, "本地名", "0.2")
    s = _load(wb).summary()
    assert s["defined_name_count"] == 2
    assert s["workbook_scoped_count"] == 1
    assert s["sheet_scoped_count"] == 1
    assert s["name_usage_count"] == 1


# ------------------------------------------------------------ 现有能力不回归
def test_direct_refs_still_work_alongside_names():
    wb = Workbook()
    d = wb.active
    d.title = "数据"
    d["A1"] = 1; d["A2"] = 2
    c = wb.create_sheet("计算")
    c["B1"] = "=数据!A1+命名"
    c["B2"] = "=B1+数据!A2"
    _global(wb, "命名", "数据!$A$2")
    g = _load(wb)
    # 直接引用与名称展开共存于同一公式
    assert g.precedents["计算!B1"] == {"数据!A1", "数据!A2"}
    res = g.impact("数据!A2", direction="downstream")
    assert {"计算!B1", "计算!B2"} <= {i["cell"] for i in res["impacted"]}


def test_named_edges_participate_in_cell_cycle_detection():
    # A1 -> 名称 -> B1 -> A1：经名称建立的单元格环也应被 SCC 检出
    wb = Workbook()
    s = wb.active
    s.title = "循环"
    s["A1"] = "=指B"
    s["B1"] = "=A1+1"
    _global(wb, "指B", "循环!$B$1")
    g = _load(wb)
    flat = {n for cyc in g.cycles for n in cyc}
    assert {"循环!A1", "循环!B1"} <= flat


def test_reserved_xlnm_names_are_ignored():
    # 打印区域 / 自动筛选产生的 _xlnm.* 内置名称不应进清单、不应报未定义
    wb = Workbook()
    ws = wb.active
    ws.title = "S"
    ws["A1"] = 1
    ws.print_area = "A1:A1"
    ws.auto_filter.ref = "A1:A1"
    ws["B1"] = "=A1+1"
    g = _load(wb)
    assert g.name_inventory == []
    codes = {i["code"] for i in g.audit()["issues"]}
    assert "undefined_name" not in codes


def test_whole_column_name_not_out_of_range_and_expands():
    wb = Workbook()
    d = wb.active
    d.title = "数据"
    d["A1"] = 1; d["A2"] = 2
    c = wb.create_sheet("计算")
    c["B1"] = "=SUM(整列名)"
    _global(wb, "整列名", "数据!$A:$A")
    g = _load(wb)
    assert g.out_of_range_refs == []
    assert g.precedents["计算!B1"] == {"数据!A1", "数据!A2"}


def test_quoted_space_sheet_name_in_definition():
    wb = Workbook()
    d = wb.active
    d.title = "月报 2026"
    d["B2"] = 8
    c = wb.create_sheet("计算")
    c["B1"] = "=空格表名"
    _global(wb, "空格表名", "'月报 2026'!$B$2")
    g = _load(wb)
    assert "计算!B1" in g.dependents["月报 2026!B2"]
    assert g.name_usages["计算!B1"][0].status == "resolved"


def test_leading_equals_constant_definition():
    wb = Workbook()
    c = wb.active
    c.title = "计算"
    c["B1"] = "=带等号常量"
    _global(wb, "带等号常量", "=0.2")
    g = _load(wb)
    u = g.name_usages["计算!B1"][0]
    assert u.status == "constant" and u.target_count == 0
    assert g.edge_count == 0


def test_three_node_name_cycle_reports_full_chain_once():
    wb = Workbook()
    s = wb.active
    s.title = "S"
    s["A1"] = "=环甲"
    _global(wb, "环甲", "环乙")
    _global(wb, "环乙", "环丙")
    _global(wb, "环丙", "环甲")
    g = _load(wb)
    res = g.audit()
    inv = [i for i in res["issues"] if i["code"] == "invalid_name"]
    assert len(inv) == 1
    assert "环甲 -> 环乙 -> 环丙 -> 环甲" in inv[0]["message"]


def test_diamond_name_reference_not_falsely_flagged_as_cycle():
    # 同一公式片段两次引用同一名称（基础+基础）是钻石型，不是名称环
    wb = Workbook()
    d = wb.active
    d.title = "数据"
    d["A1"] = 1
    c = wb.create_sheet("计算")
    c["B1"] = "=钻石"
    _global(wb, "基础", "数据!$A$1")
    _global(wb, "钻石", "基础+基础")
    g = _load(wb)
    u = g.name_usages["计算!B1"][0]
    assert u.status == "formula" and u.target_count == 1
    assert not u.issues
