"""依赖健康审计的核心逻辑测试。运行：python -m pytest tests/ -v"""
import io

from openpyxl import Workbook, load_workbook

from app.graph import DEFAULT_FANOUT_THRESHOLD, WorkbookGraph


def load_graph(wb: Workbook) -> WorkbookGraph:
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    wbf = load_workbook(buf, data_only=False)
    buf.seek(0)
    wbv = load_workbook(buf, data_only=True)
    return WorkbookGraph.from_workbooks(wbf, wbv)


def build_audit_workbook() -> WorkbookGraph:
    """覆盖全部审计问题类型的工作簿。"""
    wb = Workbook()
    p = wb.active
    p.title = "参数"
    p["A1"] = "标签"
    p["B1"] = 10
    p["A2"] = "空行"                 # B2 落在已用范围内但为空
    d = wb.create_sheet("数据")
    for i in range(1, 51):
        d.cell(row=i, column=1, value=i)
    c = wb.create_sheet("计算")
    c["B1"] = "=参数!B1*2"           # 正常公式，被 B2 引用
    c["B2"] = "=B1+1"                # 正常
    c["B3"] = "=幽灵表!A1"           # 缺失工作表
    c["B4"] = "=参数!B2+1"           # 引用空单元格
    c["B5"] = "=数据!A999"           # 超出已用范围
    c["B6"] = "=SUM(数据!A1:A50)"    # 50 个直接依赖，高扇出
    c["B7"] = "=1+2"                 # 孤立公式
    return load_graph(wb)


def _by_code(res):
    return {i["code"]: i for i in res["issues"]}


def test_audit_finds_all_issue_types():
    g = build_audit_workbook()
    res = g.audit(fanout_threshold=20)
    by = _by_code(res)
    assert set(by) == {
        "missing_sheet", "empty_ref", "out_of_range",
        "high_fanout", "orphan_formula",
    }
    assert by["missing_sheet"]["level"] == "error"
    assert by["missing_sheet"]["cell"] == "计算!B3"
    assert by["missing_sheet"]["target_sheet"] == "幽灵表"
    assert by["missing_sheet"]["evidence"][0]["from"] == "计算!B3"
    assert by["missing_sheet"]["evidence"][0]["to_display"] == "幽灵表!A1"

    assert by["empty_ref"]["level"] == "warning"
    assert by["empty_ref"]["cell"] == "参数!B2"
    assert by["empty_ref"]["evidence"][0]["from"] == "计算!B4"
    assert by["empty_ref"]["evidence"][0]["to"] == "参数!B2"

    assert by["out_of_range"]["cell"] == "计算!B5"
    assert by["out_of_range"]["target_sheet"] == "数据"
    assert by["out_of_range"]["evidence"][0]["to"] == "数据!A999"

    assert by["high_fanout"]["cell"] == "计算!B6"
    assert by["high_fanout"]["fanout"] == 50
    assert by["high_fanout"]["threshold"] == 20
    assert len(by["high_fanout"]["evidence"]) == 50

    assert by["orphan_formula"]["cell"] == "计算!B7"
    assert by["orphan_formula"]["level"] == "info"
    assert by["orphan_formula"]["evidence"] == []


def test_audit_counts_consistent_with_issues():
    g = build_audit_workbook()
    res = g.audit(fanout_threshold=20)
    assert res["total"] == len(res["issues"])
    assert sum(res["counts_by_level"].values()) == res["total"]
    assert sum(res["counts_by_code"].values()) == res["total"]
    assert res["counts_by_level"]["error"] == 1


def test_audit_clean_workbook_has_no_issues():
    wb = Workbook()
    ws = wb.active
    ws.title = "S1"
    ws["A1"] = 1
    ws["A2"] = 2
    ws["B1"] = "=A1+A2"
    ws["B2"] = "=B1*2"
    ws["B3"] = "=B2+1"                # 末端公式：合法但无人引用，属孤立公式
    g = load_graph(wb)
    res = g.audit()
    codes = set(res["counts_by_code"])
    assert codes <= {"orphan_formula"}
    assert res["counts_by_code"].get("orphan_formula") == 1
    assert res["counts_by_level"]["error"] == 0
    assert res["counts_by_level"]["warning"] == 0


def test_audit_level_filter():
    g = build_audit_workbook()
    res = g.audit(levels=("error",), fanout_threshold=20)
    assert {i["code"] for i in res["issues"]} == {"missing_sheet"}
    res = g.audit(levels=("warning",), fanout_threshold=20)
    assert set(res["counts_by_code"]) == {"empty_ref", "out_of_range", "high_fanout"}
    res = g.audit(levels=("info",), fanout_threshold=20)
    assert all(i["code"] == "orphan_formula" for i in res["issues"])


def test_audit_sheet_filter_uses_formula_owner_for_missing_sheet():
    g = build_audit_workbook()
    res = g.audit(sheet="参数", fanout_threshold=20)
    # 空引用问题归属于目标单元格所在工作表
    assert {i["code"] for i in res["issues"]} == {"empty_ref"}
    res = g.audit(sheet="计算", fanout_threshold=20)
    codes = set(res["counts_by_code"])
    assert "missing_sheet" in codes and "out_of_range" in codes
    assert "empty_ref" not in codes


def test_audit_fanout_threshold():
    g = build_audit_workbook()
    # 阈值提高到 50（> 而非 >=），高扇出问题消失
    res = g.audit(fanout_threshold=50)
    assert "high_fanout" not in res["counts_by_code"]
    res = g.audit(fanout_threshold=49)
    assert "high_fanout" in res["counts_by_code"]
    # 默认阈值
    assert g.audit()["fanout_threshold"] == DEFAULT_FANOUT_THRESHOLD


def test_audit_issues_sorted_by_severity():
    g = build_audit_workbook()
    levels = [i["level"] for i in g.audit(fanout_threshold=20)["issues"]]
    weights = {"error": 0, "warning": 1, "info": 2}
    assert levels == sorted(levels, key=lambda lv: weights[lv])


def test_audit_evidence_truncation():
    wb = Workbook()
    ws = wb.active
    ws.title = "S1"
    for i in range(1, 201):
        ws.cell(row=i, column=1, value=i)
    ws["C1"] = "=SUM(A1:A200)"          # 200 条证据边
    g = load_graph(wb)
    res = g.audit(fanout_threshold=10)
    issue = next(i for i in res["issues"] if i["code"] == "high_fanout")
    assert len(issue["evidence"]) == 100
    assert issue["evidence_truncated"] is True
    assert issue["evidence_total"] == 200
    # 导出模式不截断
    full = g.audit(fanout_threshold=10, evidence_limit=None)
    issue = next(i for i in full["issues"] if i["code"] == "high_fanout")
    assert len(issue["evidence"]) == 200
    assert "evidence_truncated" not in issue


def test_out_of_range_edge_not_in_graph_but_recorded():
    g = build_audit_workbook()
    # 越界坐标不入依赖图，避免污染影响面分析
    assert "数据!A999" not in g.dependents
    o = g.out_of_range_refs[0]
    assert o.coord == "A999" and o.sheet == "数据"
    # 同公式重复越界引用只报一次
    assert len(g.out_of_range_refs) == 1


def test_missing_sheet_refs_recorded():
    g = build_audit_workbook()
    assert len(g.missing_sheet_refs) == 1
    m = g.missing_sheet_refs[0]
    assert m.source == "计算!B3" and m.sheet == "幽灵表" and m.ref_text == "幽灵表!A1"


def test_whole_column_range_is_not_out_of_range():
    # 整列引用语义上覆盖全表，即使工作表很空也不算越界
    wb = Workbook()
    ws = wb.active
    ws.title = "S1"
    ws["A1"] = 1
    ws["B1"] = "=SUM(A:A)"
    g = load_graph(wb)
    assert g.out_of_range_refs == []
    res = g.audit()
    assert "out_of_range" not in res["counts_by_code"]
