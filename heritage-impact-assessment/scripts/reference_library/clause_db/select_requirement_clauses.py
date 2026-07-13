#!/usr/bin/env python3
"""Select candidate clauses and emit requirement_facts.jsonl."""

from __future__ import annotations

import argparse
import json
import os
import re
from pathlib import Path
from typing import Any


REFERENCE_LIBRARY_ROOT = Path(
    os.environ.get("HERITAGE_REFERENCE_LIBRARY", "/Users/drevan01/Desktop/文物影响评估与保护方案资料库")
)
LIB_ROOT = REFERENCE_LIBRARY_ROOT / "01_法规政策与标准/法规条文库"
CLAUSE_LIBRARY = LIB_ROOT / "clause_library.jsonl"


def read_jsonl(path: Path) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    if not path.exists():
        return rows
    with path.open("r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line:
                rows.append(json.loads(line))
    return rows


def write_jsonl(path: Path, rows: list[dict[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as f:
        for row in rows:
            f.write(json.dumps(row, ensure_ascii=False, sort_keys=True) + "\n")


def load_profile(path: Path | None) -> dict[str, Any]:
    if path is None:
        return {}
    return json.loads(path.read_text(encoding="utf-8"))


def normalize_terms(values: Any) -> set[str]:
    if not values:
        return set()
    if isinstance(values, str):
        parts = re.split(r"[、,，;；\s]+", values)
        return {p.strip() for p in parts if p.strip()}
    if isinstance(values, list):
        out: set[str] = set()
        for value in values:
            out |= normalize_terms(value)
        return out
    return {str(values)}


def profile_terms(profile: dict[str, Any]) -> dict[str, set[str]]:
    return {
        "zones": normalize_terms(profile.get("zones")),
        "topics": normalize_terms(profile.get("topics")),
        "impact_factors": normalize_terms(profile.get("impact_factors")),
        "project_types": normalize_terms(profile.get("project_type")) | normalize_terms(profile.get("project_types")),
        "heritage_types": normalize_terms(profile.get("heritage_type")) | normalize_terms(profile.get("heritage_types")),
        "keywords": normalize_terms(profile.get("keywords"))
        | normalize_terms(profile.get("project_name"))
        | normalize_terms(profile.get("heritage_name")),
    }


def score_clause(clause: dict[str, Any], terms: dict[str, set[str]]) -> tuple[int, list[str]]:
    score = 0
    reasons: list[str] = []

    weights = {
        "zones": 7,
        "topics": 5,
        "impact_factors": 5,
        "project_types": 4,
        "heritage_types": 4,
    }
    for field, weight in weights.items():
        matches = terms[field] & normalize_terms(clause.get(field))
        if matches:
            score += weight * len(matches)
            reasons.append(f"{field}:{','.join(sorted(matches))}")

    keyword_text = "\n".join(
        [
            clause.get("source_title", ""),
            clause.get("clause_text", ""),
            " ".join(clause.get("trigger_keywords") or []),
            clause.get("applies_to", ""),
        ]
    )
    keyword_hits = sorted(k for k in terms["keywords"] if k and k in keyword_text)
    if keyword_hits:
        score += 2 * len(keyword_hits)
        reasons.append("keywords:" + ",".join(keyword_hits[:10]))

    if clause.get("source_status") == "replaced":
        score -= 20
        reasons.append("status:replaced")

    return score, reasons


def to_requirement_fact(clause: dict[str, Any], score: int, reasons: list[str]) -> dict[str, Any]:
    notes = [
        f"clause_library_score={score}",
        f"granularity={clause.get('granularity', '')}",
        f"source_status={clause.get('source_status', '')}",
    ]
    if reasons:
        notes.append("matched=" + " | ".join(reasons))
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


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--project-profile", type=Path, help="JSON project profile.")
    parser.add_argument("--output", type=Path, required=True, help="Output requirement_facts.jsonl path.")
    parser.add_argument("--min-score", type=int, default=1)
    parser.add_argument("--limit", type=int, default=80)
    args = parser.parse_args()

    profile = load_profile(args.project_profile)
    terms = profile_terms(profile)
    clauses = read_jsonl(CLAUSE_LIBRARY)

    selected: list[tuple[int, list[str], dict[str, Any]]] = []
    for clause in clauses:
        score, reasons = score_clause(clause, terms)
        if score >= args.min_score:
            selected.append((score, reasons, clause))

    selected.sort(key=lambda item: (-item[0], item[2].get("source_type", ""), item[2]["clause_id"]))
    rows = [to_requirement_fact(clause, score, reasons) for score, reasons, clause in selected[: args.limit]]
    write_jsonl(args.output, rows)
    print(json.dumps({"selected": len(rows), "output": str(args.output)}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
