#!/usr/bin/env python3
"""Generate a standalone Chapter 5 fact-rule pairing HTML page."""

from __future__ import annotations

import argparse
import csv
import html
import json
from datetime import datetime
from pathlib import Path
from typing import Any


DESIGN_ROWS = [
    "项目流程",
    "环境保护",
    "景观保护",
    "用地类型",
    "建筑高度与建筑密度",
]
CONSTRUCTION_ROWS = [
    "施工过程对环境的影响（包括污染排放、地质灾害、水土流失、生态环境）",
    "施工过程对地下不明文物的扰动",
]
OPERATION_ROWS = [
    "污染排放与生态环境影响",
    "整体建筑风貌与视觉景观影响",
]
IMPACT_FACTORS = [
    "有较大益处",
    "有较小益处",
    "正面影响可忽略",
    "没有改变",
    "负面影响可忽略",
    "有较小负面影响",
    "有较大负面影响",
]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="生成第五章事实与条文人工匹配 HTML 页面")
    parser.add_argument("--工作目录", "--workspace", dest="work_dir", required=True)
    parser.add_argument("--输出目录", dest="output_dir", default="")
    parser.add_argument("--文件名", dest="filename", default="chapter5_fact_rule_pairing.html")
    parser.add_argument("--项目名称", dest="project_name", default="")
    return parser.parse_args()


def read_jsonl(path: Path) -> list[dict[str, Any]]:
    if not path.exists():
        return []
    rows: list[dict[str, Any]] = []
    for line in path.read_text(encoding="utf-8", errors="ignore").splitlines():
        if line.strip():
            rows.append(json.loads(line))
    return rows


def first_text(row: dict[str, Any], keys: list[str], default: str = "") -> str:
    for key in keys:
        value = row.get(key)
        if value not in (None, ""):
            return str(value)
    return default


def normalize_facts(rows: list[dict[str, Any]]) -> list[dict[str, str]]:
    facts = []
    for index, row in enumerate(rows, start=1):
        fact_id = first_text(row, ["fact_id", "id"], f"F{index:04d}")
        field_name = first_text(row, ["field_name", "fact_type"], "项目事实")
        value = first_text(row, ["value", "fact_text", "text"])
        if not value:
            continue
        facts.append(
            {
                "id": fact_id,
                "title": field_name,
                "text": value,
                "source": first_text(row, ["source_file", "source_title"]),
                "location": first_text(row, ["source_location", "location"]),
                "evidence": first_text(row, ["source_evidence_id", "evidence_id"]),
                "notes": first_text(row, ["notes", "use_mode", "confidence"]),
            }
        )
    return facts


def normalize_rules(rows: list[dict[str, Any]]) -> list[dict[str, str]]:
    rules = []
    for index, row in enumerate(rows, start=1):
        rule_id = first_text(row, ["requirement_id", "rule_id", "id"], f"R{index:04d}")
        text = first_text(row, ["requirement_text", "rule_text", "text", "value"])
        if not text:
            continue
        source_title = first_text(row, ["requirement_source_title", "source_title", "source_file"])
        control = first_text(row, ["control_object", "applies_to", "field_name"], "管控条文")
        rules.append(
            {
                "id": rule_id,
                "title": control,
                "text": text,
                "source": source_title,
                "location": first_text(row, ["requirement_location", "source_location", "location"]),
                "evidence": first_text(row, ["requirement_source_evidence_id", "source_evidence_id", "evidence_id"]),
                "control_value": first_text(row, ["control_value"]),
                "notes": first_text(row, ["notes", "applies_to"]),
            }
        )
    return rules


def infer_project_name(work_dir: Path, explicit: str) -> str:
    if explicit:
        return explicit
    state_path = work_dir / "module_state.json"
    if state_path.exists():
        try:
            state = json.loads(state_path.read_text(encoding="utf-8"))
            return str(state.get("project_name") or "")
        except json.JSONDecodeError:
            return ""
    return ""


def write_seed_csv(path: Path, facts: list[dict[str, str]], rules: list[dict[str, str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["type", "id", "title", "text", "source", "location", "evidence", "notes"])
        for row in facts:
            writer.writerow(["fact", row["id"], row["title"], row["text"], row["source"], row["location"], row["evidence"], row["notes"]])
        for row in rules:
            writer.writerow(["rule", row["id"], row["title"], row["text"], row["source"], row["location"], row["evidence"], row["notes"]])


def build_html(payload: dict[str, Any]) -> str:
    data = json.dumps(payload, ensure_ascii=False)
    project = html.escape(payload.get("project_name") or "未命名项目")
    return f"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>第五章事实与条文人工匹配</title>
  <style>
    :root {{
      --bg: #f7f8f5;
      --panel: #ffffff;
      --panel-soft: #eef3ef;
      --text: #1f2926;
      --muted: #66736e;
      --line: #d8dfd9;
      --accent: #236b5b;
      --accent-soft: #dcebe5;
      --warn: #9d5a26;
      --warn-soft: #f7eadf;
      --shadow: 0 12px 28px rgba(22, 42, 37, 0.08);
      --radius: 8px;
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      background: var(--bg);
      color: var(--text);
      font-family: -apple-system, BlinkMacSystemFont, "PingFang SC", "Microsoft YaHei", sans-serif;
      font-size: 14px;
      line-height: 1.55;
    }}
    header {{
      position: sticky;
      top: 0;
      z-index: 4;
      display: flex;
      justify-content: space-between;
      gap: 16px;
      padding: 16px 22px;
      border-bottom: 1px solid var(--line);
      background: rgba(255, 255, 255, 0.92);
      backdrop-filter: blur(10px);
    }}
    h1 {{ margin: 0; font-size: 18px; font-weight: 700; letter-spacing: 0; }}
    .project {{ margin-top: 2px; color: var(--muted); font-size: 13px; }}
    .toolbar {{ display: flex; align-items: center; gap: 8px; flex-wrap: wrap; justify-content: flex-end; }}
    button, input, select, textarea {{ font: inherit; }}
    button {{
      min-height: 34px;
      border: 1px solid var(--line);
      border-radius: var(--radius);
      background: #fff;
      color: var(--text);
      padding: 6px 10px;
      cursor: pointer;
    }}
    button.primary {{ background: var(--accent); border-color: var(--accent); color: #fff; }}
    button:disabled {{ opacity: 0.45; cursor: not-allowed; }}
    main {{ display: grid; grid-template-columns: 1fr 1fr 420px; gap: 14px; padding: 14px; }}
    .panel {{
      min-height: calc(100vh - 96px);
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: var(--radius);
      box-shadow: var(--shadow);
      overflow: hidden;
      display: flex;
      flex-direction: column;
    }}
    .panel-head {{ padding: 12px; border-bottom: 1px solid var(--line); background: var(--panel-soft); }}
    .panel-title {{ display: flex; align-items: center; justify-content: space-between; gap: 10px; font-weight: 700; }}
    .count {{ color: var(--muted); font-weight: 500; font-size: 12px; }}
    .search {{ margin-top: 10px; width: 100%; border: 1px solid var(--line); border-radius: var(--radius); padding: 7px 9px; }}
    .list {{ padding: 10px; overflow: auto; display: grid; gap: 8px; }}
    .item {{
      border: 1px solid var(--line);
      border-radius: var(--radius);
      padding: 10px;
      background: #fff;
      cursor: pointer;
    }}
    .item.selected {{ border-color: var(--accent); background: var(--accent-soft); }}
    .item.used:not(.selected) {{ background: #f8faf8; }}
    .item-id {{ color: var(--accent); font-weight: 700; font-size: 12px; }}
    .item-title {{ margin-top: 2px; font-weight: 700; }}
    .item-text {{ margin-top: 6px; white-space: pre-wrap; }}
    .meta {{ margin-top: 8px; color: var(--muted); font-size: 12px; }}
    .form {{ padding: 12px; display: grid; gap: 10px; border-bottom: 1px solid var(--line); }}
    label {{ display: grid; gap: 5px; color: var(--muted); font-size: 12px; }}
    select, textarea {{
      width: 100%;
      border: 1px solid var(--line);
      border-radius: var(--radius);
      background: #fff;
      color: var(--text);
      padding: 7px 9px;
    }}
    textarea {{ min-height: 72px; resize: vertical; }}
    .selection {{ display: grid; grid-template-columns: 1fr 1fr; gap: 8px; }}
    .slot {{ border: 1px dashed var(--line); border-radius: var(--radius); padding: 8px; background: #fafbf9; min-height: 56px; }}
    .slot strong {{ display: block; color: var(--text); font-size: 12px; }}
    .pairs {{ padding: 10px; overflow: auto; display: grid; gap: 8px; }}
    .pair {{ border: 1px solid var(--line); border-radius: var(--radius); padding: 10px; background: #fff; }}
    .pair-head {{ display: flex; justify-content: space-between; gap: 10px; }}
    .pair-title {{ font-weight: 700; }}
    .pair-actions {{ display: flex; gap: 6px; }}
    .danger {{ color: var(--warn); border-color: #d9b08e; background: var(--warn-soft); }}
    .empty {{ color: var(--muted); padding: 18px; text-align: center; border: 1px dashed var(--line); border-radius: var(--radius); }}
    @media (max-width: 1180px) {{ main {{ grid-template-columns: 1fr; }} .panel {{ min-height: auto; max-height: none; }} }}
  </style>
</head>
<body>
  <header>
    <div>
      <h1>第五章事实与条文人工匹配</h1>
      <div class="project">{project}</div>
    </div>
    <div class="toolbar">
      <button id="exportJson" class="primary">导出 JSON</button>
      <button id="exportCsv">导出 CSV</button>
      <button id="clearPairs" class="danger">清空配对</button>
    </div>
  </header>
  <main>
    <section class="panel">
      <div class="panel-head">
        <div class="panel-title">项目事实 <span id="factCount" class="count"></span></div>
        <input id="factSearch" class="search" placeholder="筛选事实" />
      </div>
      <div id="factList" class="list"></div>
    </section>
    <section class="panel">
      <div class="panel-head">
        <div class="panel-title">规章条文 <span id="ruleCount" class="count"></span></div>
        <input id="ruleSearch" class="search" placeholder="筛选条文" />
      </div>
      <div id="ruleList" class="list"></div>
    </section>
    <section class="panel">
      <div class="form">
        <div class="selection">
          <div id="selectedFact" class="slot"><strong>已选事实</strong><span>未选择</span></div>
          <div id="selectedRule" class="slot"><strong>已选条文</strong><span>未选择</span></div>
        </div>
        <label>对应正文表
          <select id="targetTable">
            <option>项目设计合规性评估</option>
            <option>项目建设期影响评估</option>
            <option>项目运营期影响评估</option>
          </select>
        </label>
        <label>固定影响内容
          <select id="impactContent"></select>
        </label>
        <label>人工判读
          <select id="humanJudgment">
            <option value="">暂不填写</option>
            <option>符合</option>
            <option>不符合</option>
            <option>调整后符合</option>
            <option>不构成判断</option>
            <option>需补充条文或事实</option>
          </select>
        </label>
        <label>影响程度
          <select id="impactFactor">
            <option value="">暂不填写</option>
            {''.join(f'<option>{html.escape(value)}</option>' for value in IMPACT_FACTORS)}
          </select>
        </label>
        <label>备注
          <textarea id="pairNotes" placeholder="可填写计算口径、专业边界或需要补证的点"></textarea>
        </label>
        <button id="addPair" class="primary" disabled>加入配对记录</button>
      </div>
      <div id="pairList" class="pairs"></div>
    </section>
  </main>
  <script id="seedData" type="application/json">{data.replace("</", "<\\/")}</script>
  <script>
    const seed = JSON.parse(document.getElementById("seedData").textContent);
    const storageKey = "chapter5_fact_rule_pairs:" + (seed.workspace || seed.project_name || "default");
    let selectedFact = null;
    let selectedRule = null;
    let pairs = JSON.parse(localStorage.getItem(storageKey) || "[]");

    const rowsByTable = {{
      "项目设计合规性评估": {json.dumps(DESIGN_ROWS, ensure_ascii=False)},
      "项目建设期影响评估": {json.dumps(CONSTRUCTION_ROWS, ensure_ascii=False)},
      "项目运营期影响评估": {json.dumps(OPERATION_ROWS, ensure_ascii=False)}
    }};

    const els = {{
      factList: document.getElementById("factList"),
      ruleList: document.getElementById("ruleList"),
      pairList: document.getElementById("pairList"),
      factSearch: document.getElementById("factSearch"),
      ruleSearch: document.getElementById("ruleSearch"),
      selectedFact: document.getElementById("selectedFact"),
      selectedRule: document.getElementById("selectedRule"),
      addPair: document.getElementById("addPair"),
      targetTable: document.getElementById("targetTable"),
      impactContent: document.getElementById("impactContent"),
      humanJudgment: document.getElementById("humanJudgment"),
      impactFactor: document.getElementById("impactFactor"),
      pairNotes: document.getElementById("pairNotes"),
      factCount: document.getElementById("factCount"),
      ruleCount: document.getElementById("ruleCount")
    }};

    function esc(value) {{
      return String(value || "").replace(/[&<>"']/g, c => ({{"&":"&amp;","<":"&lt;",">":"&gt;","\\"":"&quot;","'":"&#39;"}}[c]));
    }}
    function itemMatches(item, query) {{
      const haystack = [item.id, item.title, item.text, item.source, item.location, item.evidence, item.notes].join(" ").toLowerCase();
      return haystack.includes(query.trim().toLowerCase());
    }}
    function isUsed(id, kind) {{
      return pairs.some(pair => kind === "fact" ? pair.fact_id === id : pair.requirement_id === id);
    }}
    function itemHtml(item, kind) {{
      const selected = (kind === "fact" && selectedFact?.id === item.id) || (kind === "rule" && selectedRule?.id === item.id);
      const used = isUsed(item.id, kind);
      return `<article class="item ${{selected ? "selected" : ""}} ${{used ? "used" : ""}}" data-kind="${{kind}}" data-id="${{esc(item.id)}}">
        <div class="item-id">${{esc(item.id)}}</div>
        <div class="item-title">${{esc(item.title)}}</div>
        <div class="item-text">${{esc(item.text)}}</div>
        <div class="meta">${{esc([item.source, item.location, item.evidence].filter(Boolean).join(" | "))}}</div>
      </article>`;
    }}
    function renderItems() {{
      const factQuery = els.factSearch.value;
      const ruleQuery = els.ruleSearch.value;
      const facts = seed.facts.filter(item => itemMatches(item, factQuery));
      const rules = seed.rules.filter(item => itemMatches(item, ruleQuery));
      els.factCount.textContent = `${{facts.length}}/${{seed.facts.length}}`;
      els.ruleCount.textContent = `${{rules.length}}/${{seed.rules.length}}`;
      els.factList.innerHTML = facts.map(item => itemHtml(item, "fact")).join("") || `<div class="empty">没有项目事实</div>`;
      els.ruleList.innerHTML = rules.map(item => itemHtml(item, "rule")).join("") || `<div class="empty">没有规章条文</div>`;
    }}
    function renderSelection() {{
      els.selectedFact.innerHTML = `<strong>已选事实</strong><span>${{selectedFact ? esc(selectedFact.id + " " + selectedFact.title) : "未选择"}}</span>`;
      els.selectedRule.innerHTML = `<strong>已选条文</strong><span>${{selectedRule ? esc(selectedRule.id + " " + selectedRule.title) : "未选择"}}</span>`;
      els.addPair.disabled = !(selectedFact && selectedRule);
    }}
    function updateImpactRows() {{
      const rows = rowsByTable[els.targetTable.value] || [];
      els.impactContent.innerHTML = rows.map(row => `<option>${{esc(row)}}</option>`).join("");
    }}
    function savePairs() {{
      localStorage.setItem(storageKey, JSON.stringify(pairs));
    }}
    function renderPairs() {{
      els.pairList.innerHTML = pairs.map((pair, index) => `<article class="pair">
        <div class="pair-head">
          <div class="pair-title">${{esc(pair.pair_id)}} | ${{esc(pair.target_table)}} | ${{esc(pair.impact_content)}}</div>
          <div class="pair-actions"><button data-remove="${{index}}" class="danger">删除</button></div>
        </div>
        <div class="meta">事实：${{esc(pair.fact_id)}} | 条文：${{esc(pair.requirement_id)}} | 判读：${{esc(pair.human_judgment || "未填")}} | 影响程度：${{esc(pair.impact_factor || "未填")}}</div>
        <div class="item-text">${{esc(pair.notes || "")}}</div>
      </article>`).join("") || `<div class="empty">尚未建立配对记录</div>`;
    }}
    function addPair() {{
      if (!(selectedFact && selectedRule)) return;
      const pair = {{
        pair_id: `P${{String(pairs.length + 1).padStart(4, "0")}}`,
        fact_id: selectedFact.id,
        requirement_id: selectedRule.id,
        target_table: els.targetTable.value,
        impact_content: els.impactContent.value,
        human_judgment: els.humanJudgment.value,
        impact_factor: els.impactFactor.value,
        notes: els.pairNotes.value.trim()
      }};
      pairs.push(pair);
      els.pairNotes.value = "";
      savePairs();
      renderItems();
      renderPairs();
    }}
    function exportPayload() {{
      return {{
        schema: "chapter5_fact_rule_pairs.v1",
        project_name: seed.project_name,
        workspace: seed.workspace,
        generated_at: new Date().toISOString(),
        pairs
      }};
    }}
    function download(filename, text, type) {{
      const blob = new Blob([text], {{ type }});
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = filename;
      a.click();
      URL.revokeObjectURL(url);
    }}
    function toCsv() {{
      const header = ["pair_id","fact_id","requirement_id","target_table","impact_content","human_judgment","impact_factor","notes"];
      const lines = [header.join(",")];
      for (const pair of pairs) {{
        lines.push(header.map(key => `"${{String(pair[key] || "").replaceAll('"', '""')}}"`).join(","));
      }}
      return "\\ufeff" + lines.join("\\n");
    }}
    document.addEventListener("click", event => {{
      const item = event.target.closest(".item");
      if (item) {{
        const list = item.dataset.kind === "fact" ? seed.facts : seed.rules;
        const found = list.find(row => row.id === item.dataset.id);
        if (item.dataset.kind === "fact") selectedFact = found;
        if (item.dataset.kind === "rule") selectedRule = found;
        renderSelection();
        renderItems();
      }}
      const remove = event.target.dataset.remove;
      if (remove !== undefined) {{
        pairs.splice(Number(remove), 1);
        savePairs();
        renderItems();
        renderPairs();
      }}
    }});
    els.factSearch.addEventListener("input", renderItems);
    els.ruleSearch.addEventListener("input", renderItems);
    els.targetTable.addEventListener("change", updateImpactRows);
    els.addPair.addEventListener("click", addPair);
    document.getElementById("exportJson").addEventListener("click", () => download("chapter5_fact_rule_pairs.json", JSON.stringify(exportPayload(), null, 2), "application/json"));
    document.getElementById("exportCsv").addEventListener("click", () => download("chapter5_fact_rule_pairs.csv", toCsv(), "text/csv;charset=utf-8"));
    document.getElementById("clearPairs").addEventListener("click", () => {{
      if (!pairs.length || confirm("清空全部配对记录？")) {{
        pairs = [];
        savePairs();
        renderItems();
        renderPairs();
      }}
    }});
    updateImpactRows();
    renderSelection();
    renderItems();
    renderPairs();
  </script>
</body>
</html>
"""


def main() -> None:
    args = parse_args()
    work_dir = Path(args.work_dir).expanduser().resolve()
    output_dir = Path(args.output_dir).expanduser().resolve() if args.output_dir else work_dir / "human_input"
    output_dir.mkdir(parents=True, exist_ok=True)

    facts = normalize_facts(read_jsonl(work_dir / "facts" / "project_facts.jsonl"))
    rules = normalize_rules(read_jsonl(work_dir / "facts" / "requirement_facts.jsonl"))
    project_name = infer_project_name(work_dir, args.project_name)
    generated_at = datetime.now().isoformat(timespec="seconds")
    payload = {
        "schema": "chapter5_fact_rule_pairing_seed.v1",
        "project_name": project_name,
        "workspace": str(work_dir),
        "generated_at": generated_at,
        "facts": facts,
        "rules": rules,
    }

    seed_path = output_dir / "chapter5_fact_rule_pairing_seed.json"
    seed_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    write_seed_csv(output_dir / "chapter5_fact_rule_pairing_seed.csv", facts, rules)

    html_path = output_dir / args.filename
    html_path.write_text(build_html(payload), encoding="utf-8")
    prompt_dir = work_dir / "next_prompts"
    prompt_dir.mkdir(parents=True, exist_ok=True)
    (prompt_dir / "next_prompt_analysis.md").write_text(
        "\n".join(
            [
                "# 下一阶段启动提示",
                "",
                "## 阶段",
                "analysis",
                "",
                "## 项目路径",
                f"`{work_dir}`",
                "",
                "## 前置条件",
                "- 用户已完成 `human_input/chapter5_fact_rule_pairing.html` 中的事实与条文配对。",
                "- 配对结果已保存为 `human_input/chapter5_fact_rule_pairs.json`、`chapter5_fact_rule_pairs.csv` 或 `chapter5_pairing_notes.md`。",
                "",
                "## 必读规则文件",
                "- `references/04-分析判断模块.md`",
                "- `references/19-第五章事实条文匹配页面与影响因子规则.md`",
                "- `references/05-文评成稿样本与文章架构.md`",
                "- `references/11-固定正文结构与固定内容.md`",
                "- `references/12-文物对象概述与价值评估规则.md`",
                "- `references/13-空间关系写作规则.md`",
                "",
                "## 只读输入",
                "- `facts/project_facts.jsonl`",
                "- `facts/heritage_facts.jsonl`",
                "- `facts/requirement_facts.jsonl`",
                "- `facts/quote_candidates.jsonl`",
                "- `facts/source_coverage.jsonl`",
                "- `evidence/evidence_register.jsonl`",
                "- `processing_output/external_sources.jsonl`",
                "- `human_input/chapter5_fact_rule_pairs.json` 或同等配对输入",
                "",
                "## 必写输出",
                "- `analysis/chapter5_fact_rule_pairs.jsonl`",
                "- `analysis/chapter5_design_compliance.jsonl`",
                "- `analysis/chapter5_construction_impacts.jsonl`",
                "- `analysis/chapter5_operation_impacts.jsonl`",
                "- `analysis/impact_matrix.jsonl`",
                "- `analysis/mitigation_matrix.jsonl`",
                "- `analysis/risk_flags.jsonl`",
                "- `next_prompts/next_prompt_report_assembly.md`",
            ]
        )
        + "\n",
        encoding="utf-8",
    )
    print(html_path)


if __name__ == "__main__":
    main()
