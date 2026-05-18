"""補助金情報のSQLite永続化（companies.db に subsidies テーブルを追加）"""

from __future__ import annotations

import sqlite3
from datetime import datetime, timedelta
from pathlib import Path
from typing import Iterable

import pandas as pd


DB_PATH = Path(__file__).resolve().parent.parent / "companies.db"


_SCHEMA = """
CREATE TABLE IF NOT EXISTS subsidies (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    title TEXT NOT NULL,
    summary TEXT,
    source_name TEXT,
    source_category TEXT,
    url TEXT UNIQUE,
    domain TEXT,
    matched_keyword TEXT,
    collected_at TEXT,
    first_seen_at TEXT DEFAULT (datetime('now', 'localtime')),
    last_seen_at TEXT
)
"""


def _conn() -> sqlite3.Connection:
    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    return conn


def init_db() -> None:
    conn = _conn()
    conn.execute(_SCHEMA)
    conn.commit()
    conn.close()


def upsert_subsidies(items: Iterable[dict]) -> tuple[int, int]:
    """URL ベースで upsert。新規件数と更新件数を返す。"""
    init_db()
    conn = _conn()
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    new_count = 0
    update_count = 0
    for item in items:
        existing = conn.execute(
            "SELECT id FROM subsidies WHERE url = ?", (item["url"],)
        ).fetchone()
        if existing:
            conn.execute(
                """
                UPDATE subsidies
                SET title = ?, summary = ?, source_name = ?, source_category = ?,
                    domain = ?, matched_keyword = ?, collected_at = ?, last_seen_at = ?
                WHERE url = ?
                """,
                (
                    item["title"], item["summary"], item["source_name"],
                    item["source_category"], item["domain"], item["matched_keyword"],
                    item["collected_at"], now, item["url"],
                ),
            )
            update_count += 1
        else:
            conn.execute(
                """
                INSERT INTO subsidies
                (title, summary, source_name, source_category, url, domain,
                 matched_keyword, collected_at, first_seen_at, last_seen_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    item["title"], item["summary"], item["source_name"],
                    item["source_category"], item["url"], item["domain"],
                    item["matched_keyword"], item["collected_at"], now, now,
                ),
            )
            new_count += 1
    conn.commit()
    conn.close()
    return new_count, update_count


def get_subsidies_since(days: int = 7) -> pd.DataFrame:
    """直近 N 日に新たに見つかった補助金を返す"""
    init_db()
    conn = _conn()
    threshold = (datetime.now() - timedelta(days=days)).strftime("%Y-%m-%d %H:%M:%S")
    df = pd.read_sql_query(
        "SELECT * FROM subsidies WHERE first_seen_at >= ? ORDER BY first_seen_at DESC",
        conn,
        params=(threshold,),
    )
    conn.close()
    return df


def get_all_subsidies() -> pd.DataFrame:
    init_db()
    conn = _conn()
    df = pd.read_sql_query(
        "SELECT * FROM subsidies ORDER BY first_seen_at DESC", conn
    )
    conn.close()
    return df
