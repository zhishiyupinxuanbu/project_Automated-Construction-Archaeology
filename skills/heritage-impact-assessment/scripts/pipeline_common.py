from __future__ import annotations

import csv
import datetime as dt
import json
from pathlib import Path
from typing import Any


def now_iso() -> str:
    return dt.datetime.now(dt.timezone.utc).astimezone().isoformat(timespec="seconds")


def read_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def write_json(path: Path, data: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def read_jsonl(path: Path) -> list[dict[str, Any]]:
    if not path.exists():
        return []
    rows = []
    for line in path.read_text(encoding="utf-8").splitlines():
        if line.strip():
            rows.append(json.loads(line))
    return rows


def append_jsonl(path: Path, row: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a", encoding="utf-8") as f:
        f.write(json.dumps(row, ensure_ascii=False) + "\n")


def write_jsonl(path: Path, rows: list[dict[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as f:
        for row in rows:
            f.write(json.dumps(row, ensure_ascii=False) + "\n")


def read_csv(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8", newline="") as f:
        return list(csv.DictReader(f))


def update_module_state(work_dir: Path, stage: str) -> None:
    state_path = work_dir / "module_state.json"
    state = read_json(state_path) if state_path.exists() else {}
    completed = state.setdefault("completed_stages", [])
    if stage not in completed:
        completed.append(stage)
    state["current_stage"] = stage
    state["blocked"] = False
    state["updated_at"] = now_iso()
    write_json(state_path, state)


def first_text(rows: list[dict[str, Any]], work_dir: Path, keywords: list[str]) -> str:
    for row in rows:
        source = row.get("source_file", "")
        haystack = source + " " + row.get("text_path", "")
        if any(keyword in haystack for keyword in keywords):
            path = work_dir / row["text_path"]
            if path.exists():
                text = path.read_text(encoding="utf-8", errors="ignore").strip()
                if text:
                    return text
    for row in rows:
        path = work_dir / row.get("text_path", "")
        if path.exists():
            text = path.read_text(encoding="utf-8", errors="ignore").strip()
            if text:
                return text
    return ""


def truncate(text: str, limit: int = 500) -> str:
    text = " ".join(text.split())
    return text[:limit] + ("..." if len(text) > limit else "")
