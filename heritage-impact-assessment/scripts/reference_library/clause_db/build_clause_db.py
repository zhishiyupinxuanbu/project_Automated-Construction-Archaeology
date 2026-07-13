#!/usr/bin/env python3
"""Build SQLite + FTS database from the clause library JSONL files."""

from __future__ import annotations

import json
import os
import re
import sqlite3
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


REFERENCE_LIBRARY_ROOT = Path(
    os.environ.get("HERITAGE_REFERENCE_LIBRARY", "/Users/drevan01/Desktop/文物影响评估与保护方案资料库")
)
LIB_ROOT = REFERENCE_LIBRARY_ROOT / "01_法规政策与标准/法规条文库"
DB_PATH = LIB_ROOT / "法规条文库.sqlite"
SOURCES_JSONL = LIB_ROOT / "sources.jsonl"
CLAUSES_JSONL = LIB_ROOT / "clause_library.jsonl"
EXTRA_SOURCE_JSONLS = [
    LIB_ROOT / "reference_report_sources.jsonl",
]
EXTRA_CLAUSE_JSONLS = [
    LIB_ROOT / "reference_report_atomic_clauses.jsonl",
]


TAG_FIELDS = {
    "zones": "zone",
    "topics": "topic",
    "impact_factors": "impact_factor",
    "project_types": "project_type",
    "heritage_types": "heritage_type",
    "trigger_keywords": "keyword",
}


def now_iso() -> str:
    return datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds")


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


def read_many_jsonl(paths: list[Path]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for path in paths:
        rows.extend(read_jsonl(path))
    return rows


def json_text(value: Any) -> str:
    return json.dumps(value if value is not None else [], ensure_ascii=False, sort_keys=True)


def connect(db_path: Path) -> sqlite3.Connection:
    conn = sqlite3.connect(str(db_path))
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    return conn


def create_schema(conn: sqlite3.Connection) -> None:
    conn.executescript(
        """
        DROP TABLE IF EXISTS build_meta;
        DROP TABLE IF EXISTS source_files;
        DROP TABLE IF EXISTS clause_search_terms;
        DROP TABLE IF EXISTS clause_tags;
        DROP TABLE IF EXISTS clauses;
        DROP TABLE IF EXISTS sources;
        DROP TABLE IF EXISTS clauses_fts;

        CREATE TABLE build_meta (
          key TEXT PRIMARY KEY,
          value TEXT NOT NULL
        );

        CREATE TABLE sources (
          source_id TEXT PRIMARY KEY,
          source_title TEXT NOT NULL,
          source_type TEXT,
          source_code TEXT,
          source_version TEXT,
          source_status TEXT,
          source_url TEXT,
          source_page TEXT,
          source_tier TEXT,
          local_path TEXT,
          text_path TEXT,
          content_sha256 TEXT,
          source_note TEXT,
          replaces_json TEXT,
          ocr_required INTEGER DEFAULT 0,
          text_chars INTEGER DEFAULT 0,
          raw_json TEXT NOT NULL
        );

        CREATE TABLE clauses (
          clause_id TEXT PRIMARY KEY,
          source_id TEXT NOT NULL,
          source_title TEXT NOT NULL,
          source_type TEXT,
          source_code TEXT,
          source_version TEXT,
          source_status TEXT,
          granularity TEXT NOT NULL,
          clause_location TEXT,
          clause_text TEXT NOT NULL,
          control_object TEXT,
          control_value TEXT,
          control_unit TEXT,
          applies_to TEXT,
          local_path TEXT,
          text_path TEXT,
          source_url TEXT,
          notes TEXT,
          raw_json TEXT NOT NULL,
          FOREIGN KEY (source_id) REFERENCES sources(source_id) ON DELETE CASCADE
        );

        CREATE TABLE clause_tags (
          clause_id TEXT NOT NULL,
          tag_type TEXT NOT NULL,
          tag_value TEXT NOT NULL,
          PRIMARY KEY (clause_id, tag_type, tag_value),
          FOREIGN KEY (clause_id) REFERENCES clauses(clause_id) ON DELETE CASCADE
        );

        CREATE TABLE clause_search_terms (
          clause_id TEXT NOT NULL,
          term TEXT NOT NULL,
          PRIMARY KEY (clause_id, term),
          FOREIGN KEY (clause_id) REFERENCES clauses(clause_id) ON DELETE CASCADE
        );

        CREATE TABLE source_files (
          source_id TEXT NOT NULL,
          file_role TEXT NOT NULL,
          file_path TEXT NOT NULL,
          PRIMARY KEY (source_id, file_role, file_path),
          FOREIGN KEY (source_id) REFERENCES sources(source_id) ON DELETE CASCADE
        );

        CREATE VIRTUAL TABLE clauses_fts USING fts5(
          clause_id UNINDEXED,
          source_title,
          clause_location,
          clause_text,
          control_object,
          applies_to,
          tags,
          tokenize='trigram'
        );

        CREATE INDEX idx_clause_tags_type_value ON clause_tags(tag_type, tag_value);
        CREATE INDEX idx_clause_search_terms_term ON clause_search_terms(term);
        CREATE INDEX idx_clauses_source_id ON clauses(source_id);
        CREATE INDEX idx_clauses_granularity ON clauses(granularity);
        CREATE INDEX idx_clauses_status ON clauses(source_status);
        """
    )


def search_terms(text: str, seed_terms: list[str]) -> set[str]:
    terms = {term.strip() for term in seed_terms if term and term.strip()}
    for match in re.findall(r"\d+(?:\.\d+)?\s*(?:mm/s|km/h|cm|m)", text, flags=re.I):
        terms.add(re.sub(r"\s+", "", match))
    compact = re.sub(r"\s+", "", text)
    compact = re.sub(r"[^\w\u4e00-\u9fff]+", "", compact)
    cjk_runs = re.findall(r"[\u4e00-\u9fff]{2,}", compact)
    for run in cjk_runs:
        for size in (2, 3, 4):
            for i in range(0, max(0, len(run) - size + 1)):
                terms.add(run[i : i + size])
    return {term for term in terms if 1 < len(term) <= 40}


def insert_sources(conn: sqlite3.Connection, sources: list[dict[str, Any]]) -> None:
    for row in sources:
        conn.execute(
            """
            INSERT OR REPLACE INTO sources (
              source_id, source_title, source_type, source_code, source_version,
              source_status, source_url, source_page, source_tier, local_path,
              text_path, content_sha256, source_note, replaces_json,
              ocr_required, text_chars, raw_json
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                row.get("source_id", ""),
                row.get("source_title", ""),
                row.get("source_type", ""),
                row.get("source_code", ""),
                row.get("source_version", ""),
                row.get("source_status", ""),
                row.get("source_url", ""),
                row.get("source_page", ""),
                row.get("source_tier", ""),
                row.get("local_path", ""),
                row.get("text_path", ""),
                row.get("content_sha256", ""),
                row.get("source_note", ""),
                json_text(row.get("replaces", [])),
                1 if row.get("ocr_required") else 0,
                int(row.get("text_chars") or 0),
                json.dumps(row, ensure_ascii=False, sort_keys=True),
            ),
        )
        if row.get("local_path"):
            conn.execute(
                "INSERT OR IGNORE INTO source_files (source_id, file_role, file_path) VALUES (?, ?, ?)",
                (row.get("source_id", ""), "source", row["local_path"]),
            )
        if row.get("text_path"):
            conn.execute(
                "INSERT OR IGNORE INTO source_files (source_id, file_role, file_path) VALUES (?, ?, ?)",
                (row.get("source_id", ""), "text", row["text_path"]),
            )


def ensure_sources_for_clauses(conn: sqlite3.Connection, clauses: list[dict[str, Any]]) -> None:
    existing = {row["source_id"] for row in conn.execute("SELECT source_id FROM sources")}
    for row in clauses:
        source_id = row.get("source_id", "")
        if not source_id or source_id in existing:
            continue
        conn.execute(
            """
            INSERT INTO sources (
              source_id, source_title, source_type, source_code, source_version,
              source_status, raw_json
            ) VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                source_id,
                row.get("source_title", ""),
                row.get("source_type", ""),
                row.get("source_code", ""),
                row.get("source_version", ""),
                row.get("source_status", ""),
                json.dumps({"generated_from_clause": row}, ensure_ascii=False, sort_keys=True),
            ),
        )
        existing.add(source_id)


def insert_clauses(conn: sqlite3.Connection, clauses: list[dict[str, Any]]) -> None:
    for row in clauses:
        conn.execute(
            """
            INSERT OR REPLACE INTO clauses (
              clause_id, source_id, source_title, source_type, source_code,
              source_version, source_status, granularity, clause_location,
              clause_text, control_object, control_value, control_unit,
              applies_to, local_path, text_path, source_url, notes, raw_json
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                row.get("clause_id", ""),
                row.get("source_id", ""),
                row.get("source_title", ""),
                row.get("source_type", ""),
                row.get("source_code", ""),
                row.get("source_version", ""),
                row.get("source_status", ""),
                row.get("granularity", ""),
                row.get("clause_location", ""),
                row.get("clause_text", ""),
                row.get("control_object", ""),
                row.get("control_value", ""),
                row.get("control_unit", ""),
                row.get("applies_to", ""),
                row.get("local_path", ""),
                row.get("text_path", ""),
                row.get("source_url", ""),
                row.get("notes", ""),
                json.dumps(row, ensure_ascii=False, sort_keys=True),
            ),
        )

        tag_values: list[str] = []
        for field, tag_type in TAG_FIELDS.items():
            for value in row.get(field) or []:
                if value:
                    tag_values.append(str(value))
                    conn.execute(
                        "INSERT OR IGNORE INTO clause_tags (clause_id, tag_type, tag_value) VALUES (?, ?, ?)",
                        (row.get("clause_id", ""), tag_type, str(value)),
                    )

        searchable_text = "\n".join(
            [
                row.get("source_title", ""),
                row.get("clause_location", ""),
                row.get("clause_text", ""),
                row.get("control_object", ""),
                row.get("control_value", ""),
                row.get("control_unit", ""),
                row.get("applies_to", ""),
                row.get("notes", ""),
            ]
        )
        seed_terms = [
            *tag_values,
            row.get("control_object", ""),
            row.get("control_value", ""),
            row.get("control_unit", ""),
        ]
        for term in search_terms(searchable_text, seed_terms):
            conn.execute(
                "INSERT OR IGNORE INTO clause_search_terms (clause_id, term) VALUES (?, ?)",
                (row.get("clause_id", ""), term),
            )

        conn.execute(
            """
            INSERT INTO clauses_fts (
              clause_id, source_title, clause_location, clause_text,
              control_object, applies_to, tags
            ) VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                row.get("clause_id", ""),
                row.get("source_title", ""),
                row.get("clause_location", ""),
                row.get("clause_text", ""),
                row.get("control_object", ""),
                row.get("applies_to", ""),
                " ".join(tag_values),
            ),
        )


def main() -> None:
    sources = read_many_jsonl([SOURCES_JSONL, *EXTRA_SOURCE_JSONLS])
    clauses = read_many_jsonl([CLAUSES_JSONL, *EXTRA_CLAUSE_JSONLS])
    if not clauses:
        raise SystemExit(f"No clauses found: {CLAUSES_JSONL}")

    if DB_PATH.exists():
        DB_PATH.unlink()
    for suffix in ["-wal", "-shm"]:
        sidecar = Path(str(DB_PATH) + suffix)
        if sidecar.exists():
            sidecar.unlink()

    with connect(DB_PATH) as conn:
        create_schema(conn)
        insert_sources(conn, sources)
        ensure_sources_for_clauses(conn, clauses)
        insert_clauses(conn, clauses)
        conn.execute("INSERT INTO build_meta (key, value) VALUES (?, ?)", ("generated_at", now_iso()))
        conn.execute("INSERT INTO build_meta (key, value) VALUES (?, ?)", ("source_count", str(len(sources))))
        conn.execute("INSERT INTO build_meta (key, value) VALUES (?, ?)", ("clause_count", str(len(clauses))))
        conn.commit()
        conn.execute("INSERT INTO clauses_fts(clauses_fts) VALUES('optimize')")

    print(
        json.dumps(
            {
                "db": str(DB_PATH),
                "source_count": len(sources),
                "clause_count": len(clauses),
                "generated_at": now_iso(),
            },
            ensure_ascii=False,
            indent=2,
        )
    )


if __name__ == "__main__":
    main()
