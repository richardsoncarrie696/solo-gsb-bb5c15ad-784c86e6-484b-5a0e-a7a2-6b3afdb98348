# 电子表格公式影响分析（Spreadsheet Formula Impact Analysis）

一个可自托管的 Web 应用：上传 Excel（`.xlsx`）文件后，系统解析工作表中的公式，
将**跨工作表引用**、**矩形区域引用**与**名称管理器定义名称**展开为单元格级别的
依赖图，自动检测**循环引用**（含名称循环），并可针对任意单元格列出其**所有直接 /
间接受影响单元格及完整影响路径**。

- 后端：FastAPI + openpyxl
- 前端：原生 HTML / CSS / JavaScript 单页（无需构建，无需 Node）
- 部署：Dockerfile + docker-compose 一键启动

## 功能一览

- 📤 拖拽 / 点选上传 `.xlsx`（上限 25MB，内存中保留最近 5 个文件）
- 🔗 解析单元格引用、矩形区域（`A1:B10`）、整列（`B:B`）、整行（`2:5`）、
  绝对引用（`$A$1`）、跨表引用（`参数!B2`）、带空格的工作表名（`'月报 2026'!B2`）
- 🏷 **名称管理器支持**：解析公式中的定义名称并展开其指向的单元格 / 区域为依赖边
  - 同时识别**工作簿级**与**工作表级**名称；无限定名称按「先本表本地、再工作簿」
    解析，跨表限定名称（`计算!单价`）按「先限定表本地、再工作簿」解析，正确处理同名遮蔽
  - 支持多区域（`A1:A3,C1:C2`）、整行 / 整列、嵌套名称链、公式表达式名称、常量名称
  - **未定义名称、限定到不存在工作表的名称、`#REF!` 失效名称、名称循环**生成可定位的
    审计问题；指向不存在工作表 / 越界的名称目标**不加入影响面**
  - 常量 / 外部工作簿名称不产生依赖边，仅保留解析状态与证据
- 🔁 Tarjan 强连通分量算法检测循环引用（含自环），并单独检测名称定义循环，
  页面顶部高亮展示
- 🔍 选定单元格后展示：
  - **下游影响**：哪些单元格直接 / 间接引用了它（改了它会影响谁）
  - **上游依赖**：该单元格直接 / 间接依赖谁（它的计算需要谁）
  - 每个受影响单元格的**层级（深度）**和一条**最短影响路径**
- 🗂 可按工作表筛选、按单元格 / 公式文本 / 名称名 / 名称定义搜索、切换是否只看公式单元格
- 📄 查看每个单元格的公式原文、Excel 缓存计算值，及其引用的定义名称与解析状态
- 🩺 **依赖健康审计**：扫描依赖图并按规则定级，支持按级别 / 工作表筛选
  - `error` —— 跨表公式或定义名称指向**不存在的工作表**、公式使用了**未定义名称**、
    定义名称目标**失效（`#REF!` / 空定义 / 名称循环）**
  - `warning` —— 被公式引用却**为空的单元格**、引用**超出目标工作表已用范围**、
    定义名称**指向外部工作簿**、**扇出（直接依赖数）超过阈值**的公式节点
  - `info` —— **未被任何公式引用的孤立公式**
  - 每条问题给出级别、单元格位置、问题说明、涉及名称的作用域 / 定义 / 解析链与
    **证据边**，并附带按级别 / 类型的汇总计数；页面提供审计面板与「导出审计 JSON」
- 💾 一键导出完整分析 JSON（工作表统计、**名称清单与名称使用留痕**、公式清单、
  依赖边、循环环、依赖健康审计）

## 一键启动（Docker Compose）

```bash
docker compose up --build -d
```

然后浏览器打开：

> **http://localhost:8000**

停止：

```bash
docker compose down
```

如需修改端口，编辑 `docker-compose.yml` 中的 `"8000:8000"`（左为主机端口）。

容器启动后可直接验证健康状态与审计接口：

```bash
curl http://localhost:8000/api/health
# 首页下载示例文件上传后，用返回的 workbook_id 查看审计结果
curl -F "file=@samples/sample_impact.xlsx" http://localhost:8000/api/upload
curl "http://localhost:8000/api/workbooks/$WID/audit?level=warning"
```

> 上传的工作簿仅保存在容器内存中（最近 5 个），容器重启即清空；
> 镜像构建时会预生成示例文件 `samples/sample_impact.xlsx`。

## 本地开发（不使用 Docker）

需要 Python 3.11+：

```bash
pip install -r requirements.txt
python scripts/make_sample.py          # 生成示例文件（可选，仓库已附带）
uvicorn app.main:app --host 0.0.0.0 --port 8000 --reload
```

运行测试：

```bash
pip install pytest httpx
python -m pytest tests/ -v
```

## 端口与 API

服务监听 **8000** 端口，全部接口：

| 方法 | 路径 | 说明 |
|---|---|---|
| `GET` | `/` | 前端单页 |
| `POST` | `/api/upload` | 上传 xlsx（multipart 字段名 `file`），返回工作簿 id、统计与循环信息 |
| `GET` | `/api/workbooks/{id}` | 工作簿统计摘要 |
| `GET` | `/api/workbooks/{id}/cells?sheet=&q=&formulas_only=` | 单元格列表（支持工作表筛选 / 搜索） |
| `GET` | `/api/workbooks/{id}/impact?cell=参数!B2&direction=downstream` | 影响面分析；`direction` 取 `downstream`（默认，受影响）或 `upstream`（依赖链） |
| `GET` | `/api/workbooks/{id}/audit?level=&sheet=&fanout_threshold=` | 依赖健康审计；支持按级别（`error`/`warning`/`info`，可重复）、工作表、扇出阈值筛选 |
| `GET` | `/api/workbooks/{id}/audit/export` | 导出依赖健康审计 JSON（证据边不截断） |
| `GET` | `/api/workbooks/{id}/export` | 导出完整分析 JSON（含 `audit` 段） |
| `GET` | `/samples/sample_impact.xlsx` | 下载示例文件 |
| `GET` | `/api/health` | 健康检查 |

### 接口示例

```bash
# 上传并取得 workbook_id
curl -F "file=@samples/sample_impact.xlsx" http://localhost:8000/api/upload

# 全量审计（默认扇出阈值 20）
curl "http://localhost:8000/api/workbooks/$WID/audit"

# 只看错误与警告、限定「审计示例」工作表、扇出阈值改为 5
curl "http://localhost:8000/api/workbooks/$WID/audit?level=error&level=warning&sheet=审计示例&fanout_threshold=5"

# 下载审计 JSON（证据边完整，不做 100 条截断）
curl -OJ "http://localhost:8000/api/workbooks/$WID/audit/export"
```

错误响应均为 4xx：工作簿不存在 / 已被淘汰返回 `404`；
`level` 非法、`fanout_threshold` 不是 1–10000 的整数返回 `400`
（响应体形如 `{"detail": "fanout_threshold 必须是 1-10000 的整数"}`）。

### 审计 JSON 结构

```json
{
  "workbook_id": "ab12cd34ef56",
  "issues": [
    {
      "code": "undefined_name",
      "level": "error",
      "sheet": "名称",
      "cell": "名称!B9",
      "cell_display": "名称!B9",
      "message": "定义名称 不存在的名称：工作簿名称管理器中均未定义名称 '不存在的名称'",
      "formula": "=不存在的名称+1",
      "name": "不存在的名称",
      "name_qualifier": null,
      "name_display": "不存在的名称",
      "name_scope": null,
      "name_definition": null,
      "name_status": "undefined",
      "name_chain": [],
      "target_sheet": null,
      "target_display": "不存在的名称",
      "evidence": [
        { "from": "名称!B9", "from_display": "名称!B9",
          "to_display": "不存在的名称", "ref": "不存在的名称" }
      ]
    },
    {
      "code": "missing_sheet",
      "level": "error",
      "sheet": "名称",
      "cell": "名称!B11",
      "message": "引用了不存在的工作表 '不存在的表'（不存在的表!$A$1）（经定义名称 幽灵表名称 展开）",
      "via_name": "幽灵表名称",
      "target_sheet": "不存在的表",
      "target_display": "不存在的表!$A$1",
      "evidence": [
        { "from": "名称!B11", "to_display": "不存在的表!$A$1",
          "ref": "不存在的表!$A$1" }
      ]
    },
    {
      "code": "empty_ref",
      "level": "warning",
      "sheet": "参数",
      "cell": "参数!B8",
      "cell_display": "参数!B8",
      "message": "单元格 参数!B8 为空，但被 2 个公式引用",
      "formula": null,
      "referenced_by": ["汇总!B7", "审计示例!B3"],
      "evidence": [
        { "from": "汇总!B7", "from_display": "汇总!B7",
          "to": "参数!B8", "to_display": "参数!B8" }
      ]
    },
    {
      "code": "high_fanout",
      "level": "warning",
      "cell": "审计示例!B4", "fanout": 200, "threshold": 20,
      "message": "直接引用 200 个单元格，超过扇出阈值 20",
      "evidence": [ { "from": "审计示例!B4", "to": "数据!B2" } ]
    }
  ],
  "total": 41,
  "counts_by_level": { "error": 6, "warning": 7, "info": 28 },
  "counts_by_code": { "missing_sheet": 2, "undefined_name": 2, "invalid_name": 2,
                      "external_name": 1, "empty_ref": 2, "out_of_range": 2,
                      "high_fanout": 2, "orphan_formula": 28 },
  "fanout_threshold": 20,
  "filters": { "levels": ["error", "warning", "info"], "sheet": null }
}
```

问题类型（`code`）与级别：

| code | 级别 | 含义 |
|---|---|---|
| `missing_sheet` | error | 跨表公式或定义名称指向不存在的工作表（证据边只含 `to_display` 文本；经名称展开时带 `via_name`） |
| `undefined_name` | error | 公式使用了名称管理器中不存在的名称，或跨表限定的工作表不存在；问题带 `name*` 字段 |
| `invalid_name` | error | 定义名称存在但目标无效：`#REF!`、空定义、名称间循环引用；目标不加入依赖图 |
| `external_name` | warning | 定义名称指向外部工作簿（`[1]预算.xlsx!...`），不参与本工作簿影响面 |
| `empty_ref` | warning | 目标单元格无公式也无值，按目标单元格聚合所有引用它的公式（含经名称展开）为证据边 |
| `out_of_range` | warning | 显式单元格 / 区域引用（含经名称展开）超出目标工作表已用范围（整行 / 整列引用不判越界） |
| `high_fanout` | warning | 公式直接依赖的去重单元格数 **> 阈值**（默认 20，可用 `fanout_threshold` 调整；名称展开目标同样计入） |
| `orphan_formula` | info | 公式结果未被任何其他公式引用 |

名称解析状态（`name_status`，出现在名称使用留痕与名称类问题中）：

| status | 含义 |
|---|---|
| `resolved` | 已展开为一个或多个具体单元格 / 区域并加入依赖图 |
| `formula` | 公式表达式名称（其中显式引用仍展开为边） |
| `constant` | 数字 / 字符串等常量名称，不产生依赖边 |
| `undefined` | 名称在对应作用域未定义（对应 `undefined_name` 问题） |
| `invalid` | 目标失效 / 名称循环（对应 `invalid_name` 问题），无效目标不入图 |
| `external` | 指向外部工作簿（对应 `external_name` 问题），不入图 |

> 接口返回的单条问题最多附带 100 条证据边（超出时带 `evidence_truncated` /
> `evidence_total`）；`/audit/export` 与完整导出不截断。

### 导出 JSON 结构

完整导出 `/api/workbooks/{id}/export`：

```json
{
  "filename": "sample_impact.xlsx",
  "summary": { "sheets": [...], "formula_count": 41, "edge_count": 459,
               "defined_name_count": 16, "workbook_scoped_count": 14,
               "sheet_scoped_count": 2, "name_usage_count": 17,
               "cycle_count": 2, "cycles": [["循环!A1","循环!B1"],["循环!C1"]] },
  "defined_names": [
    { "name": "单价", "scope": "workbook", "scope_sheet": null,
      "definition": "参数!$B$2", "hidden": false, "comment": null,
      "usage_status_counts": { "resolved": 2 }, "used_in_formulas": 2 }
  ],
  "name_usages": [
    { "cell": "名称!B1", "name": "单价", "qualifier": null,
      "name_display": "单价", "ref_text": "单价", "status": "resolved",
      "scope": "工作簿级", "scope_id": "workbook",
      "definition": "参数!$B$2", "chain": ["单价"],
      "target_count": 1, "issues": [] }
  ],
  "formulas": {
    "计算!B2": { "formula": "=B2*参数!B4", "value": null, "in_cycle": false,
                 "names": [ ...同 name_usages 单项... ] }
  },
  "edges": [ { "from": "参数!B2", "to": "计算!B2" } ],
  "cycles": [ ["循环!A1", "循环!B1"], ["循环!C1"] ],
  "audit": { "issues": [ ... ], "total": 41, "counts_by_level": { ... } }
}
```

审计导出 `/api/workbooks/{id}/audit/export` 在审计结果外另含
`defined_names`（名称清单）与 `name_usages`（完整名称使用留痕，含解析问题证据）。

> 说明：`value` 取自 Excel 保存时的缓存计算结果。openpyxl 不执行公式计算，
> 由程序新建、从未在 Excel 中打开保存的文件其缓存值为 `null`，属正常现象。

## 示例文件

`samples/sample_impact.xlsx`（由 `scripts/make_sample.py` 生成，也可在首页直接下载）
包含 8 个工作表，覆盖所有分析场景：

| 工作表 | 内容 |
|---|---|
| `参数` | 纯输入项（单价、数量、折扣率、税率、月数） |
| `计算` | 跨表引用链：单价×数量 → 折扣 → 税额 → 月合计 → 年合计；含工作表级名称 `系数` |
| `汇总` | 跨表矩形区域 `SUM(计算!B2:B5)`、整列引用 `SUM(参数!B:B)` |
| `月报 2026` | 带空格的工作表名（公式里写作 `='月报 2026'!...`） |
| `循环` | 故意制造的循环引用：`A1↔B1` 互引、`C1=C1*2` 自环、`D1` 为环的下游 |
| `数据` | 200 行数据表，供审计中的高扇出公式 `SUM(数据!B2:B201)` 引用 |
| `审计示例` | 覆盖直接引用类审计问题：缺失工作表、越界引用、空单元格引用、高扇出、孤立公式与一条正常公式 |
| `名称` | 覆盖名称管理器：工作簿级 / 工作表级同名遮蔽、跨表限定、区域 / 嵌套 / 公式 / 常量名称，
  以及未定义名称、`#REF!`、指向不存在工作表、外部工作簿、名称循环等审计问题 |

推荐体验路径：

1. 首页下载并上传 `sample_impact.xlsx`，顶部可看到红色循环引用提示、定义名称数与审计问题计数；
2. 左侧选择 `参数!B2`（单价），右侧显示受影响单元格（含经名称展开的 `名称!B2` 等）与逐层路径；
3. 切换到「上游依赖」，选择 `'月报 2026'!B3`，可看到完整回溯链直到 `参数!B2`；
4. 用工作表筛选定位 `循环!A1`，观察循环标记与影响面的截断行为；
5. 在「名称」表查看同名作用域遮蔽（`名称!B3` 用本地 `折扣`、`名称!B4` 跨表引用
   `计算!系数`），以及未定义 / 失效 / 外部名称的状态标记；
6. 点击右上角「🩺 依赖健康审计」，在审计面板按级别 / 工作表筛选（试试限定到「名称」表），
   或将扇出阈值调小查看高扇出节点，点击证据边中的单元格可跳转到其影响面；
7. 点击「导出分析 JSON」或「导出审计 JSON」获得机器可读的完整结果（含名称清单与使用留痕）。

## 实现要点

- `app/formula_parser.py`：词法式正则扫描公式，区分字符串字面量 / `#REF!` 错误 /
  函数名 / 结构化引用（`表[列]`），避免把 `LOG10`、`"A1"`、`1E2` 误判为单元格；
  输出未展开的 `RawRef` 与定义名称 `RawName`（含跨表限定）。
- `app/names.py`：名称管理器解析。
  - 从 `wb.defined_names`（工作簿级）与各 `ws.defined_names`（工作表级）收集名称，
    忽略 `_xlnm.*` 内置保留名，同名不同作用域并存；
  - 按 Excel 语义解析作用域（本表本地 → 工作簿；限定表本地 → 工作簿），
    顶层逗号切分多区域，区分纯引用 / 嵌套名称 / 公式表达式 / 常量 / `#REF!` / 外部引用；
  - 名称体内未限定引用：工作表级名称相对其所属表，工作簿级名称相对使用公式所在表；
  - 迭代工作清单展开嵌套名称链（上限 `MAX_NAME_DEPTH=50`），用不可变路径集合检测
    名称循环且不误判钻石型引用；返回解析状态、目标 `RawRef`、名称链与可定位问题。
- `app/graph.py`：
  - 加载两次工作簿（`data_only=False` 取公式、`data_only=True` 取缓存值）；
  - 区域展开为具体单元格，整行 / 整列按工作表已用范围截断，单个区域上限 5000 格
    （`MAX_RANGE_CELLS`，防 `SUM(A:XFD)` 之类打爆内存）；
  - 名称目标与直接引用走同一套展开 / 越界 / 缺失工作表逻辑，无效目标一律不入图；
  - 迭代版 Tarjan SCC 检测单元格循环引用（经名称建立的边同样参与）；
  - BFS 求影响面，记录最短路径、层级与环标记，默认最大深度 50 / 5000 节点；
  - 构图时记录不入图的异常引用与名称使用留痕，`audit()` 汇总八类问题并按级别 /
    工作表筛选，证据边超过 100 条时截断；导出保留名称、解析状态与完整证据。
- `app/main.py`：FastAPI 路由与内存 LRU 式工作簿缓存（不落盘，重启即清空）。

## 目录结构

```
.
├── app/
│   ├── main.py              # FastAPI 入口与 API
│   ├── formula_parser.py    # 公式引用 / 定义名称词法解析
│   ├── names.py             # 名称管理器：作用域解析、展开、名称环检测
│   ├── graph.py             # 依赖图 / 环检测 / 影响面 BFS / 审计
│   └── static/index.html    # 前端单页
├── samples/sample_impact.xlsx
├── scripts/make_sample.py   # 示例文件生成脚本
├── tests/                   # pytest 用例（78 个）
├── Dockerfile
├── docker-compose.yml
└── requirements.txt
```
