# -*- coding: utf-8 -*-
"""Schema-level tests for the match-group tables (SQLite backend)."""
import importlib
import sqlite3

import pytest


@pytest.fixture()
def fresh_db(tmp_path, monkeypatch):
    monkeypatch.setenv("PBK_DB", str(tmp_path / "g.db"))
    monkeypatch.delenv("DATABASE_URL", raising=False)
    from app import db as dbmod
    importlib.reload(dbmod)          # re-read PBK_DB / USE_PG
    dbmod.init_db()
    yield dbmod
    importlib.reload(dbmod)          # restore for other test modules


def test_group_tables_exist_and_link(fresh_db):
    db = fresh_db
    with db.get_conn() as conn:
        conn.execute("INSERT INTO purchase_lines(qty) VALUES (10)")
        conn.execute("INSERT INTO sale_lines(qty) VALUES (10)")
        cur = conn.execute("INSERT INTO match_groups(note) VALUES ('hi')")
        gid = cur.lastrowid
        conn.execute(
            "INSERT INTO match_group_purchases(group_id, purchase_line_id) VALUES (?,?)",
            (gid, 1))
        conn.execute(
            "INSERT INTO match_group_sales(group_id, sale_line_id) VALUES (?,?)",
            (gid, 1))
        row = conn.execute(
            "SELECT note FROM match_groups WHERE id=?", (gid,)).fetchone()
    assert row["note"] == "hi"


def test_purchase_line_cannot_join_two_groups(fresh_db):
    db = fresh_db
    with db.get_conn() as conn:
        conn.execute("INSERT INTO purchase_lines(qty) VALUES (5)")
        for note in ("g1", "g2"):
            conn.execute("INSERT INTO match_groups(note) VALUES (?)", (note,))
        conn.execute(
            "INSERT INTO match_group_purchases(group_id, purchase_line_id) VALUES (1,1)")
        with pytest.raises(sqlite3.IntegrityError):
            conn.execute(
                "INSERT INTO match_group_purchases(group_id, purchase_line_id) VALUES (2,1)")


def test_deleting_group_cascades_membership(fresh_db):
    db = fresh_db
    with db.get_conn() as conn:
        conn.execute("INSERT INTO purchase_lines(qty) VALUES (5)")
        conn.execute("INSERT INTO match_groups(note) VALUES ('x')")
        conn.execute(
            "INSERT INTO match_group_purchases(group_id, purchase_line_id) VALUES (1,1)")
        conn.execute("DELETE FROM match_groups WHERE id=1")
        left = conn.execute("SELECT COUNT(*) c FROM match_group_purchases").fetchone()["c"]
    assert left == 0
