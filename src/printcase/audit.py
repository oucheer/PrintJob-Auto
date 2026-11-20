from __future__ import annotations

import csv
import json
import sqlite3
from datetime import datetime
from pathlib import Path
from typing import Iterable

from .models import CaseResult


class AuditStore:
    def __init__(self, database_path: Path) -> None:
        self.database_path = database_path
        self.database_path.parent.mkdir(parents=True, exist_ok=True)
        with self._connect() as con:
            con.execute("""
                CREATE TABLE IF NOT EXISTS print_results (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    run_id TEXT NOT NULL,
                    case_id TEXT NOT NULL,
                    source_file TEXT,
                    printer TEXT,
                    status TEXT NOT NULL,
                    message TEXT,
                    job_id INTEGER,
                    started_at TEXT,
                    ended_at TEXT,
                    elapsed_seconds REAL,
                    settings_json TEXT
                )
            """)
            con.execute("CREATE INDEX IF NOT EXISTS idx_print_results_run ON print_results(run_id)")

    def _connect(self) -> sqlite3.Connection:
        con = sqlite3.connect(self.database_path, timeout=10)
        con.execute("PRAGMA journal_mode=WAL")
        return con

    def save(self, run_id: str, result: CaseResult) -> None:
        with self._connect() as con:
            con.execute("""
                INSERT INTO print_results (
                    run_id, case_id, source_file, printer, status, message, job_id,
                    started_at, ended_at, elapsed_seconds, settings_json
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                run_id, result.case_id, result.source_file, result.printer,
                result.status.value, result.message, result.job_id,
                result.started_at.isoformat(timespec="seconds") if result.started_at else None,
                result.ended_at.isoformat(timespec="seconds") if result.ended_at else None,
                result.elapsed_seconds, json.dumps(result.settings, ensure_ascii=False),
            ))

    def recent(self, limit: int = 500) -> list[dict]:
        """按时间倒序返回最近的历史执行记录（供界面历史页展示）。"""
        with self._connect() as con:
            cursor = con.execute("""
                SELECT started_at, run_id, case_id, status, printer, job_id,
                       elapsed_seconds, source_file, message
                FROM print_results
                ORDER BY id DESC
                LIMIT ?
            """, (limit,))
            columns = [c[0] for c in cursor.description]
            return [dict(zip(columns, row)) for row in cursor.fetchall()]

    def recent(self, limit: int = 500) -> list[dict]:
        """按时间倒序返回最近的历史执行记录（供界面历史页展示）。"""
        with self._connect() as con:
            cursor = con.execute("""
                SELECT started_at, run_id, case_id, status, printer, job_id,
                       elapsed_seconds, source_file, message
                FROM print_results
                ORDER BY id DESC
                LIMIT ?
            """, (limit,))
            columns = [c[0] for c in cursor.description]
            return [dict(zip(columns, row)) for row in cursor.fetchall()]

    def export(self, run_id: str, output_path: Path, results: Iterable[CaseResult]) -> None:
        rows = [r.to_dict() for r in results]
        if output_path.suffix.lower() == ".json":
            output_path.write_text(json.dumps({"run_id": run_id, "exported_at": datetime.now().isoformat(timespec="seconds"), "results": rows}, ensure_ascii=False, indent=2), encoding="utf-8")
            return
        with output_path.open("w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=["case_id", "status", "message", "printer", "job_id", "started_at", "ended_at", "elapsed_seconds", "source_file", "settings"])
            writer.writeheader()
            for row in rows:
                row["settings"] = json.dumps(row["settings"], ensure_ascii=False)
                writer.writerow(row)
