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


def _make_legacy_matches(db):
    """Recreate the pre-migration `matches` table with two 1:1 rows."""
    with db.get_conn() as conn:
        for q in (10, 20):
            conn.execute("INSERT INTO purchase_lines(qty) VALUES (?)", (q,))
            conn.execute("INSERT INTO sale_lines(qty) VALUES (?)", (q,))
        conn.execute(
            "CREATE TABLE matches ("
            " id INTEGER PRIMARY KEY AUTOINCREMENT,"
            " purchase_line_id INTEGER, sale_line_id INTEGER, note TEXT DEFAULT '')")
        conn.execute("INSERT INTO matches(purchase_line_id, sale_line_id, note) "
                     "VALUES (1,1,'a'),(2,2,'b')")


def test_migration_converts_matches_to_groups(fresh_db):
    db = fresh_db
    _make_legacy_matches(db)
    with db.get_conn() as conn:
        db._migrate_matches_to_groups(conn)
    with db.get_conn() as conn:
        groups = conn.execute("SELECT COUNT(*) c FROM match_groups").fetchone()["c"]
        mp = conn.execute("SELECT COUNT(*) c FROM match_group_purchases").fetchone()["c"]
        ms = conn.execute("SELECT COUNT(*) c FROM match_group_sales").fetchone()["c"]
        still = conn.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name='matches'"
        ).fetchone()
    assert (groups, mp, ms) == (2, 2, 2)
    assert still is None          # matches table dropped


def test_migration_is_idempotent_and_safe_without_matches(fresh_db):
    db = fresh_db
    # no `matches` table at all → no-op, no error
    with db.get_conn() as conn:
        db._migrate_matches_to_groups(conn)
    # now create + migrate twice
    _make_legacy_matches(db)
    with db.get_conn() as conn:
        db._migrate_matches_to_groups(conn)
    with db.get_conn() as conn:
        db._migrate_matches_to_groups(conn)         # second call: matches gone → no-op
        groups = conn.execute("SELECT COUNT(*) c FROM match_groups").fetchone()["c"]
    assert groups == 2
