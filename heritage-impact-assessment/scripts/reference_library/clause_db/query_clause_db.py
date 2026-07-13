#!/usr/bin/env python3
"""Query the SQLite clause library and emit requirement_facts JSONL."""

from __future__ import annotations

import argparse
import json
import os
import re
import sqlite3
from pathlib import Path
from typing import Any


REFERENCE_LIBRARY_ROOT = Path(
    os.environ.get("HERITAGE_REFERENCE_LIBRARY", "/Users/drevan01/Desktop/文物影响评估与保护方案资料库")
)
LIB_ROOT = REFERENCE_LIBRARY_ROOT / "01_法规政策与标准/法规条文库"
DB_PATH = LIB_ROOT / "法规条文库.sqlite"
GENERIC_IMPACT_FACTORS = {"施工期", "运营期", "环境监测", "水环境", "水质监测", "空气质量"}


def connect(db_path: Path) -> sqlite3.Connection:
    conn = sqlite3.connect(str(db_path))
    conn.row_factory = sqlite3.Row
    return conn


def normalize_terms(values: Any) -> set[str]:
    if not values:
        return set()
    if isinstance(values, str):
        return {p.strip() for p in re.split(r"[、,，;；\s]+", values) if p.strip()}
    if isinstance(values, list):
        out: set[str] = set()
        for value in values:
            out |= normalize_terms(value)
        return out
    return {str(values)}


def load_profile(path: Path | None) -> dict[str, Any]:
    if path is None:
        return {}
    return json.loads(path.read_text(encoding="utf-8"))


def profile_terms(profile: dict[str, Any]) -> dict[str, set[str]]:
    project_names = normalize_terms(profile.get("project_name"))
    heritage_names = normalize_terms(profile.get("heritage_name"))
    keywords = normalize_terms(profile.get("keywords")) - project_names - heritage_names - GENERIC_IMPACT_FACTORS
    return {
        "zone": normalize_terms(profile.get("zones")),
        "topic": normalize_terms(profile.get("topics")),
        "impact_factor": normalize_terms(profile.get("impact_factors")),
        "project_type": normalize_terms(profile.get("project_type")) | normalize_terms(profile.get("project_types")),
        "heritage_type": normalize_terms(profile.get("heritage_type")) | normalize_terms(profile.get("heritage_types")),
        "keyword": keywords,
    }


def fts_query_from_keywords(keywords: set[str]) -> str:
    safe_terms = []
    for word in sorted(keywords):
        if not word or len(word) < 3 or len(word) > 40:
            continue
        cleaned = re.sub(r'["\']', " ", word).strip()
        if cleaned:
            safe_terms.append(f'"{cleaned}"')
    return " OR ".join(safe_terms[:20])


def add_score(scores: dict[str, dict[str, Any]], clause_id: str, points: int, reason: str) -> None:
    item = scores.setdefault(clause_id, {"score": 0, "reasons": []})
    item["score"] += points
    item["reasons"].append(reason)


def collect_scores(conn: sqlite3.Connection, terms: dict[str, set[str]]) -> dict[str, dict[str, Any]]:
    scores: dict[str, dict[str, Any]] = {}
    weights = {
        "zone": 7,
        "topic": 5,
        "impact_factor": 5,
        "project_type": 4,
        "heritage_type": 4,
        "keyword": 2,
    }

    for tag_type, values in terms.items():
        if not values:
            continue
        placeholders = ",".join("?" for _ in values)
        sql = f"""
            SELECT clause_id, tag_value
            FROM clause_tags
            WHERE tag_type = ? AND tag_value IN ({placeholders})
        """
        for row in conn.execute(sql, [tag_type, *sorted(values)]):
            add_score(scores, row["clause_id"], weights[tag_type], f"{tag_type}:{row['tag_value']}")

    searchable_impact_terms = terms["impact_factor"] - GENERIC_IMPACT_FACTORS
    query = fts_query_from_keywords(terms["keyword"] | searchable_impact_terms)
    if query:
        try:
            for row in conn.execute(
                """
                SELECT clause_id, bm25(clauses_fts) AS rank
                FROM clauses_fts
                WHERE clauses_fts MATCH ?
                LIMIT 200
                """,
                (query,),
            ):
                add_score(scores, row["clause_id"], 2, "fts")
        except sqlite3.OperationalError:
            pass

    short_terms = sorted(term for term in (terms["keyword"] | searchable_impact_terms) if 1 < len(term) <= 40)
    if short_terms:
        placeholders = ",".join("?" for _ in short_terms)
        sql = f"""
            SELECT clause_id, term
            FROM clause_search_terms
            WHERE term IN ({placeholders})
        """
        for row in conn.execute(sql, short_terms):
            add_score(scores, row["clause_id"], 2, f"search:{row['term']}")

    return scores


def clause_tag_values(conn: sqlite3.Connection, clause_id: str, tag_type: str) -> set[str]:
    return {
        row["tag_value"]
        for row in conn.execute(
            "SELECT tag_value FROM clause_tags WHERE clause_id = ? AND tag_type = ?",
            (clause_id, tag_type),
        )
    }


def allowed_zone_tags(profile_zones: set[str]) -> set[str]:
    allowed: set[str] = set()
    if not profile_zones:
        return allowed

    allowed.add("zone:all")
    has_buffer = "zone:buffer" in profile_zones or any(z.startswith("zone:buffer_class_") for z in profile_zones)
    has_heritage_area = "zone:heritage_area" in profile_zones or any(
        z.startswith("zone:heritage_area_class_") for z in profile_zones
    )

    if has_buffer:
        allowed.add("zone:buffer")
        allowed.update(z for z in profile_zones if z.startswith("zone:buffer_class_"))
    if has_heritage_area:
        allowed.add("zone:heritage_area")
        allowed.update(z for z in profile_zones if z.startswith("zone:heritage_area_class_"))

    # Keep non-YSD/general zone tags if a future library uses them explicitly.
    allowed.update(z for z in profile_zones if not z.startswith("zone:buffer") and not z.startswith("zone:heritage_area"))
    return allowed


def passes_zone_gate(conn: sqlite3.Connection, clause_id: str, profile_zones: set[str]) -> tuple[bool, str]:
    if not profile_zones:
        return True, ""
    clause_zones = clause_tag_values(conn, clause_id, "zone")
    if not clause_zones:
        return True, ""

    allowed = allowed_zone_tags(profile_zones)
    clause_specific_zones = clause_zones - {"zone:all"}
    allowed_specific_zones = allowed - {"zone:all"}
    if clause_specific_zones:
        if clause_specific_zones & allowed_specific_zones:
            return True, ""
        return False, "zone_gate_excluded:" + ",".join(sorted(clause_zones))

    if clause_zones & allowed:
        return True, ""

    return False, "zone_gate_excluded:" + ",".join(sorted(clause_zones))


def passes_tag_gate(
    conn: sqlite3.Connection,
    clause_id: str,
    tag_type: str,
    profile_values: set[str],
    *,
    require_when_clause_tagged: bool,
) -> tuple[bool, str]:
    if not profile_values:
        return True, ""
    clause_values = clause_tag_values(conn, clause_id, tag_type)
    if not clause_values:
        return True, ""
    if clause_values & profile_values:
        return True, ""
    if require_when_clause_tagged:
        return False, f"{tag_type}_gate_excluded:" + ",".join(sorted(clause_values))
    return True, ""


def passes_category_gates(conn: sqlite3.Connection, clause_id: str, terms: dict[str, set[str]]) -> tuple[bool, list[str]]:
    reasons: list[str] = []
    checks = [
        passes_zone_gate(conn, clause_id, terms["zone"]),
        passes_tag_gate(conn, clause_id, "project_type", terms["project_type"], require_when_clause_tagged=True),
        passes_tag_gate(conn, clause_id, "heritage_type", terms["heritage_type"], require_when_clause_tagged=True),
        passes_tag_gate(conn, clause_id, "impact_factor", terms["impact_factor"], require_when_clause_tagged=True),
    ]
    for passed, reason in checks:
        if not passed:
            reasons.append(reason)
    return not reasons, reasons


def substantive_match(clause: dict[str, Any], reasons: list[str]) -> bool:
    source_type = clause.get("source_type", "")
    granularity = clause.get("granularity", "")

    if source_type == "protection_plan":
        return any(
            reason.startswith(("zone:", "topic:", "impact_factor:", "project_type:", "keyword:"))
            for reason in reasons
        )

    if source_type == "standard" or granularity == "source_level":
        matched_impacts = {
            reason.split(":", 1)[1]
            for reason in reasons
            if reason.startswith("impact_factor:") and ":" in reason
        }
        specific_impacts = matched_impacts - GENERIC_IMPACT_FACTORS
        return bool(specific_impacts or any(reason.startswith(("keyword:", "search:")) for reason in reasons))

    return True


def load_clause(conn: sqlite3.Connection, clause_id: str) -> dict[str, Any]:
    row = conn.execute("SELECT * FROM clauses WHERE clause_id = ?", (clause_id,)).fetchone()
    if row is None:
        raise KeyError(clause_id)
    return dict(row)


def to_requirement_fact(clause: dict[str, Any], score: int, reasons: list[str]) -> dict[str, Any]:
    notes = [
        f"sqlite_score={score}",
        f"granularity={clause.get('granularity', '')}",
        f"source_status={clause.get('source_status', '')}",
    ]
    if reasons:
        notes.append("matched=" + " | ".join(reasons[:20]))
    if clause.get("notes"):
        notes.append(str(clause["notes"]))
    return {
        "requirement_id": clause["clause_id"],
        "requirement_source_title": clause.get("source_title", ""),
        "requirement_source_evidence_id": clause.get("source_id", ""),
        "requirement_location": clause.get("clause_location", ""),
        "requirement_text": clause.get("clause_text", ""),
        "control_object": clause.get("control_object", ""),
        "control_value": clause.get("control_value", ""),
        "applies_to": clause.get("applies_to", ""),
        "related_project_fact_ids": [],
        "notes": "；".join(notes),
    }


def write_jsonl(path: Path, rows: list[dict[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as f:
        for row in rows:
            f.write(json.dumps(row, ensure_ascii=False, sort_keys=True) + "\n")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--db", type=Path, default=DB_PATH)
    parser.add_argument("--project-profile", type=Path)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--min-score", type=int, default=1)
    parser.add_argument("--limit", type=int, default=80)
    args = parser.parse_args()

    profile = load_profile(args.project_profile)
    terms = profile_terms(profile)
    with connect(args.db) as conn:
        scores = collect_scores(conn, terms)
        selected = []
        for clause_id, payload in scores.items():
            score = payload["score"]
            if score < args.min_score:
                continue
            allowed, gate_reasons = passes_category_gates(conn, clause_id, terms)
            if not allowed:
                continue
            clause = load_clause(conn, clause_id)
            if not substantive_match(clause, payload["reasons"]):
                continue
            if clause.get("source_status") == "replaced":
                score -= 20
            if score >= args.min_score:
                selected.append((score, payload["reasons"], clause))

        selected.sort(key=lambda item: (-item[0], item[2].get("source_type", ""), item[2]["clause_id"]))
        rows = [to_requirement_fact(clause, score, reasons) for score, reasons, clause in selected[: args.limit]]

    write_jsonl(args.output, rows)
    print(json.dumps({"selected": len(rows), "output": str(args.output)}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
