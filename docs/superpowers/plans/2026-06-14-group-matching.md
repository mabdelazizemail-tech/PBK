# Group Matching Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Allow linking many purchase invoices to one sale invoice (and vice-versa, and manual N:M), replacing the strict 1:1 `matches` table with a match-group model, plus conservative auto-discovery of groups whose quantities sum up.

**Architecture:** A `match_groups` table + two membership tables (`match_group_purchases`, `match_group_sales`) generalize the old 1:1 pair (a 1:1 match becomes a group of one purchase + one sale). A one-time idempotent migration converts legacy `matches` rows. The matching engine runs three phases: lock near-exact 1:1 pairs, discover N:1/1:N groups by subset-sum, then greedy approximate 1:1. Excel export emits each group as a stacked block; column `SUBTOTAL` totals stay exact. The UI switches the unmatched tables to multi-select checkboxes with a live reconciliation bar.

**Tech Stack:** Python 3 / FastAPI, dual SQLite+Postgres (`app/db.py` shim), openpyxl, vanilla JS + dialog/CSS frontend, pytest.

**Spec:** [docs/superpowers/specs/2026-06-14-group-matching-stock-card-design.md](../specs/2026-06-14-group-matching-stock-card-design.md)

**Conventions to respect:**
- Tests run on SQLite only (root `conftest.py` pins `PBK_DB`). `tests/test_api.py` shares **one DB across the module** — never assume an empty DB; use unique items/quantities per test.
- The PG shim translates `?`→`%s` and appends `RETURNING id` only for tables in `db._ID_TABLES`. Membership tables have **no** `id` column, so they are NOT added to `_ID_TABLES`.
- Keep the existing element IDs `#tbl-unmatched-p`, `#tbl-unmatched-s`, `#tbl-matches`, `#btn-auto-match` so `tests/ui_smoke.py` stays valid.

---

### Task 1: Group tables in the schema

**Files:**
- Modify: `app/db.py` (the `SCHEMA` string and `_ID_TABLES`)
- Test: `tests/test_db_groups.py` (create)

- [ ] **Step 1: Write the failing test**

```python
# tests/test_db_groups.py
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
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_db_groups.py -v`
Expected: FAIL — `no such table: match_groups`.

- [ ] **Step 3: Implement the schema change**

In `app/db.py`, **remove** the `CREATE TABLE IF NOT EXISTS matches (...)` block from the `SCHEMA` string and append the three new tables so `SCHEMA` ends like this:

```python
CREATE TABLE IF NOT EXISTS sale_lines (
    {_LINE_COLUMNS},
    internal_ref TEXT DEFAULT '',
    note TEXT DEFAULT '',
    UNIQUE(eta_uuid, eta_line_index)
);
CREATE TABLE IF NOT EXISTS match_groups (
    id {_PK},
    note TEXT DEFAULT '',
    created_at TEXT DEFAULT {_NOW}
);
CREATE TABLE IF NOT EXISTS match_group_purchases (
    group_id INTEGER NOT NULL REFERENCES match_groups(id) ON DELETE CASCADE,
    purchase_line_id INTEGER NOT NULL UNIQUE REFERENCES purchase_lines(id) ON DELETE CASCADE,
    PRIMARY KEY (group_id, purchase_line_id)
);
CREATE TABLE IF NOT EXISTS match_group_sales (
    group_id INTEGER NOT NULL REFERENCES match_groups(id) ON DELETE CASCADE,
    sale_line_id INTEGER NOT NULL UNIQUE REFERENCES sale_lines(id) ON DELETE CASCADE,
    PRIMARY KEY (group_id, sale_line_id)
);
"""
```

Update the ID-tables set (was `{"purchase_lines", "sale_lines", "matches"}`):

```python
_ID_TABLES = {"purchase_lines", "sale_lines", "match_groups"}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_db_groups.py -v`
Expected: PASS (3 passed).

- [ ] **Step 5: Commit**

```bash
git add app/db.py tests/test_db_groups.py
git commit -m "feat(db): add match_groups + membership tables, drop matches from schema"
```

---

### Task 2: Idempotent migration from legacy `matches`

**Files:**
- Modify: `app/db.py` (add `_table_exists`, `_migrate_matches_to_groups`; call from `init_db`)
- Test: `tests/test_db_groups.py` (append)

- [ ] **Step 1: Write the failing test**

```python
# append to tests/test_db_groups.py
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
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_db_groups.py -k migration -v`
Expected: FAIL — `module 'app.db' has no attribute '_migrate_matches_to_groups'`.

- [ ] **Step 3: Implement the migration**

In `app/db.py`, add these functions above `init_db` and call the migration inside `init_db`:

```python
def _table_exists(conn, name: str) -> bool:
    if USE_PG:
        row = conn.execute("SELECT to_regclass(?) AS t", (name,)).fetchone()
        return bool(row and row["t"])
    row = conn.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name=?", (name,)
    ).fetchone()
    return row is not None


def _migrate_matches_to_groups(conn):
    """One-time, idempotent conversion of legacy 1:1 `matches` rows into groups.

    Checks table existence FIRST (never a bare SELECT on a maybe-missing table —
    a failed statement aborts the open Postgres transaction that also created
    the schema).
    """
    if not _table_exists(conn, "matches"):
        return
    already = conn.execute("SELECT COUNT(*) AS c FROM match_groups").fetchone()["c"]
    if already and already > 0:
        conn.execute("DROP TABLE matches")      # groups already populated elsewhere
        return
    rows = conn.execute(
        "SELECT purchase_line_id, sale_line_id, note FROM matches").fetchall()
    for r in rows:
        cur = conn.execute("INSERT INTO match_groups(note) VALUES (?)", (r["note"] or "",))
        gid = cur.lastrowid
        conn.execute(
            "INSERT INTO match_group_purchases(group_id, purchase_line_id) VALUES (?,?)",
            (gid, r["purchase_line_id"]))
        conn.execute(
            "INSERT INTO match_group_sales(group_id, sale_line_id) VALUES (?,?)",
            (gid, r["sale_line_id"]))
    conn.execute("DROP TABLE matches")
```

Change `init_db`:

```python
def init_db():
    with get_conn() as conn:
        conn.executescript(SCHEMA)
        _migrate_matches_to_groups(conn)
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_db_groups.py -v`
Expected: PASS (5 passed).

- [ ] **Step 5: Commit**

```bash
git add app/db.py tests/test_db_groups.py
git commit -m "feat(db): idempotent migration of legacy matches into groups"
```

---

### Task 3: Group-aware matching engine

**Files:**
- Modify: `app/matching.py` (new constants + helpers; rewrite `auto_match`; expose `_parse_date` reuse)
- Test: `tests/test_matching.py` (update `TestAutoMatch` shape + add group tests)

- [ ] **Step 1: Write the failing tests**

Replace the body of `test_pairs_exact_quantities_correctly` (it used the old scalar shape) and append a new `TestGroupDiscovery` class:

```python
# tests/test_matching.py — replace the assertions in test_pairs_exact_quantities_correctly:
    def test_pairs_exact_quantities_correctly(self):
        purchases = [
            line(1, "2026-01-04", "سيفتى باك", "نموذج 2", 24000, 9),
            line(2, "2026-01-04", "سيفتى باك", "نموذج 2", 12000, 9),
        ]
        sales = [
            line(10, "2026-01-18", "الفشاوي", "نموذج 2", 12000, 10.5),
            line(11, "2026-01-18", "الفشاوي", "نموذج 2", 24000, 10.5),
        ]
        result = auto_match(purchases, sales)
        pairs = {(m["purchase_ids"][0], m["sale_ids"][0]) for m in result}
        assert (1, 11) in pairs
        assert (2, 10) in pairs


class TestGroupDiscovery:
    def test_two_purchases_match_one_sale(self):
        # flagship case: 28,634 + 916 == 29,550
        purchases = [
            line(1, "2026-06-05", "سيفتى باك", "نموذج 2", 28634, 9.95),
            line(2, "2026-06-06", "سيفتى باك", "نموذج 2", 916, 9.95),
        ]
        sales = [line(10, "2026-05-20", "الفيشاوي", "نموذج 2", 29550, 12)]
        result = auto_match(purchases, sales)
        groups = [m for m in result if len(m["purchase_ids"]) > 1 or len(m["sale_ids"]) > 1]
        assert len(groups) == 1
        g = groups[0]
        assert set(g["purchase_ids"]) == {1, 2}
        assert g["sale_ids"] == [10]
        assert g["qty_diff"] == pytest.approx(0)

    def test_one_purchase_matches_two_sales(self):
        purchases = [line(1, "2026-03-01", "سيفتى باك", "نموذج 5", 5000, 9)]
        sales = [
            line(10, "2026-03-03", "الفشاوي", "نموذج 5", 3000, 11),
            line(11, "2026-03-09", "براميدز", "نموذج 5", 2000, 11),
        ]
        result = auto_match(purchases, sales)
        groups = [m for m in result if len(m["sale_ids"]) > 1]
        assert len(groups) == 1
        assert groups[0]["purchase_ids"] == [1]
        assert set(groups[0]["sale_ids"]) == {10, 11}

    def test_does_not_group_across_different_items(self):
        purchases = [
            line(1, "2026-03-01", "سيفتى باك", "نموذج 1", 700, 9),
            line(2, "2026-03-01", "سيفتى باك", "نموذج 2", 300, 9),
        ]
        sales = [line(10, "2026-03-02", "الفشاوي", "نموذج 1", 1000, 11)]
        result = auto_match(purchases, sales)
        assert all(len(m["purchase_ids"]) == 1 for m in result)  # no cross-item group

    def test_does_not_group_outside_date_window(self):
        purchases = [
            line(1, "2026-01-01", "سيفتى باك", "نموذج 9", 600, 9),
            line(2, "2026-12-01", "سيفتى باك", "نموذج 9", 400, 9),   # ~11 months later
        ]
        sales = [line(10, "2026-01-03", "الفشاوي", "نموذج 9", 1000, 11)]
        result = auto_match(purchases, sales)
        assert all(len(m["purchase_ids"]) <= 1 for m in result)
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_matching.py -v`
Expected: FAIL — `KeyError: 'purchase_ids'` / group tests fail.

- [ ] **Step 3: Rewrite `auto_match` with three phases + group discovery**

In `app/matching.py`, add `from itertools import combinations` at the top, and **replace** the existing `auto_match` function (keep `normalize_ar`, `similarity`, `_parse_date`, `score_pair` unchanged) with:

```python
# --- group auto-discovery tunables -----------------------------------------
GROUP_MAX = 4            # max lines on the "many" side of an auto group
GROUP_DATE_WINDOW = 60   # days between a line and the anchor it groups under
GROUP_QTY_TOL = 0.02     # relative tolerance for a group's quantity sum
GROUP_CAND_CAP = 12      # cap candidates per item bucket before subset search
EXACT_REL = 0.005        # |qty| relative diff treated as an exact 1:1 lock


def _pair_suggestion(p: dict, s: dict) -> dict:
    return {
        "purchase_ids": [p["id"]], "sale_ids": [s["id"]],
        "score": round(score_pair(p, s), 1),
        "qty_diff": (p.get("qty") or 0) - (s.get("qty") or 0),
    }


def _is_exact_pair(p: dict, s: dict) -> bool:
    pq, sq = abs(p.get("qty") or 0), abs(s.get("qty") or 0)
    if pq == 0 or sq == 0:
        return False
    rel = abs(pq - sq) / max(pq, sq)
    return rel <= EXACT_REL and similarity(p.get("item"), s.get("item")) >= 0.5


def _greedy_pairs(purchases, sales, used_p, used_s, result, accept):
    cands = []
    for p in purchases:
        if p["id"] in used_p:
            continue
        for s in sales:
            if s["id"] in used_s:
                continue
            sc = score_pair(p, s)
            if accept(p, s, sc):
                cands.append((sc, p, s))
    cands.sort(key=lambda t: -t[0])
    for sc, p, s in cands:
        if p["id"] in used_p or s["id"] in used_s:
            continue
        used_p.add(p["id"])
        used_s.add(s["id"])
        result.append(_pair_suggestion(p, s))


def _bucket_by_item(lines):
    buckets = {}
    for ln in lines:
        buckets.setdefault(normalize_ar(ln.get("item")), []).append(ln)
    return buckets


def _date_distance(a: dict, b: dict) -> int:
    da, db_ = _parse_date(a.get("invoice_date")), _parse_date(b.get("invoice_date"))
    if da is None or db_ is None:
        return 10 ** 6
    return abs((da - db_).days)


def _find_subset(anchor, pool, used):
    """Smallest subset (size 2..GROUP_MAX) of `pool` whose |qty| sums ≈ anchor |qty|."""
    target = abs(anchor.get("qty") or 0)
    if target <= 0:
        return None
    cands = [x for x in pool
             if x["id"] not in used and _date_distance(anchor, x) <= GROUP_DATE_WINDOW]
    cands.sort(key=lambda x: _date_distance(anchor, x))
    cands = cands[:GROUP_CAND_CAP]
    tol = max(GROUP_QTY_TOL * target, 1.0)
    for size in range(2, GROUP_MAX + 1):
        for combo in combinations(cands, size):
            total = sum(abs(x.get("qty") or 0) for x in combo)
            if abs(total - target) <= tol:
                return list(combo)
    return None


def _group_score(ps, ss) -> float:
    pq = sum(abs(p.get("qty") or 0) for p in ps)
    sq = sum(abs(s.get("qty") or 0) for s in ss)
    if max(pq, sq) == 0:
        return 0.0
    return round(max(0.0, 100 * (1 - abs(pq - sq) / max(pq, sq))), 1)


def _discover_groups(rem_p, rem_s, used_p, used_s):
    out = []
    p_buckets, s_buckets = _bucket_by_item(rem_p), _bucket_by_item(rem_s)
    for key, sales_in in s_buckets.items():            # N purchases : 1 sale
        pool = p_buckets.get(key, [])
        for s in sales_in:
            if s["id"] in used_s:
                continue
            subset = _find_subset(s, pool, used_p)
            if subset:
                used_s.add(s["id"])
                used_p.update(p["id"] for p in subset)
                pq = sum(p.get("qty") or 0 for p in subset)
                out.append({"purchase_ids": [p["id"] for p in subset],
                            "sale_ids": [s["id"]],
                            "score": _group_score(subset, [s]),
                            "qty_diff": pq - (s.get("qty") or 0)})
    for key, purch_in in p_buckets.items():            # 1 purchase : N sales
        pool = s_buckets.get(key, [])
        for p in purch_in:
            if p["id"] in used_p:
                continue
            subset = _find_subset(p, pool, used_s)
            if subset:
                used_p.add(p["id"])
                used_s.update(s["id"] for s in subset)
                sq = sum(s.get("qty") or 0 for s in subset)
                out.append({"purchase_ids": [p["id"]],
                            "sale_ids": [s["id"] for s in subset],
                            "score": _group_score([p], subset),
                            "qty_diff": (p.get("qty") or 0) - sq})
    return out


def auto_match(purchases: list[dict], sales: list[dict]) -> list[dict]:
    """Suggest matches. Each suggestion is {purchase_ids, sale_ids, score, qty_diff}.

    Phase A locks near-exact 1:1 pairs, phase B discovers N:1/1:N groups whose
    quantities sum tightly, phase C greedily pairs the approximate leftovers.
    """
    used_p, used_s, result = set(), set(), []
    _greedy_pairs(purchases, sales, used_p, used_s, result,
                  accept=lambda p, s, sc: _is_exact_pair(p, s))
    rem_p = [p for p in purchases if p["id"] not in used_p]
    rem_s = [s for s in sales if s["id"] not in used_s]
    result += _discover_groups(rem_p, rem_s, used_p, used_s)
    _greedy_pairs(purchases, sales, used_p, used_s, result,
                  accept=lambda p, s, sc: sc >= MATCH_THRESHOLD)
    return result
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_matching.py -v`
Expected: PASS (all, including `TestGroupDiscovery`).

- [ ] **Step 5: Commit**

```bash
git add app/matching.py tests/test_matching.py
git commit -m "feat(matching): three-phase auto-match with N:1/1:N group discovery"
```

---

### Task 4: Group CRUD + line/dashboard queries

**Files:**
- Modify: `app/main.py` (`MatchCreate`, `list_matches`, `create_match`, `delete_match`, `_unmatched`, `list_lines`, `dashboard`)
- Test: `tests/test_api.py` (update `TestMatching` 1:1 tests; add group tests)

- [ ] **Step 1: Write the failing tests**

In `tests/test_api.py`, **replace** `test_manual_match_and_unlink` and `test_line_cannot_match_twice` to use the array shape, and add group tests:

```python
    def test_manual_match_and_unlink(self, client):
        p, s = _purchase(client), _sale(client)
        r = client.post("/api/matches",
                        json={"purchase_ids": [p["id"]], "sale_ids": [s["id"]]})
        assert r.status_code == 200, r.text
        gid = r.json()["id"]
        groups = client.get("/api/matches").json()
        g = [x for x in groups if x["id"] == gid][0]
        assert [pp["id"] for pp in g["purchases"]] == [p["id"]]
        assert g["qty_diff"] == pytest.approx(0)
        assert client.delete(f"/api/matches/{gid}").status_code == 200

    def test_line_cannot_match_twice(self, client):
        p, s1, s2 = _purchase(client), _sale(client), _sale(client)
        assert client.post("/api/matches", json={
            "purchase_ids": [p["id"]], "sale_ids": [s1["id"]]}).status_code == 200
        r = client.post("/api/matches", json={
            "purchase_ids": [p["id"]], "sale_ids": [s2["id"]]})
        assert r.status_code == 409

    def test_group_two_purchases_one_sale(self, client):
        p1 = _purchase(client, qty=28634, item="بلوك صنف أ")
        p2 = _purchase(client, qty=916, item="بلوك صنف أ")
        s = _sale(client, qty=29550, item="بلوك صنف أ")
        r = client.post("/api/matches", json={
            "purchase_ids": [p1["id"], p2["id"]], "sale_ids": [s["id"]]})
        assert r.status_code == 200, r.text
        gid = r.json()["id"]
        g = [x for x in client.get("/api/matches").json() if x["id"] == gid][0]
        assert {pp["id"] for pp in g["purchases"]} == {p1["id"], p2["id"]}
        assert g["purchase_qty"] == pytest.approx(29550)
        assert g["qty_diff"] == pytest.approx(0)

    def test_match_requires_both_sides(self, client):
        p = _purchase(client)
        r = client.post("/api/matches", json={"purchase_ids": [p["id"]], "sale_ids": []})
        assert r.status_code == 400

    def test_unmatched_filter_excludes_grouped_lines(self, client):
        p = _purchase(client, item="فلتر صنف ب")
        s = _sale(client, item="فلتر صنف ب")
        client.post("/api/matches",
                    json={"purchase_ids": [p["id"]], "sale_ids": [s["id"]]})
        un = client.get("/api/lines", params={"kind": "purchase", "matched": "false"}).json()
        assert all(row["id"] != p["id"] for row in un)
        allp = client.get("/api/lines", params={"kind": "purchase"}).json()
        assert [row for row in allp if row["id"] == p["id"]][0]["group_id"] is not None
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_api.py::TestMatching -v`
Expected: FAIL — `create_match` still expects scalar `purchase_id`; `group_id` column missing.

- [ ] **Step 3: Rewrite the endpoints**

In `app/main.py`, replace the `MatchCreate` model and the `list_matches` / `create_match` / `delete_match` functions, and update `_unmatched`, `list_lines`, and the match-related lines of `dashboard`:

```python
class MatchCreate(BaseModel):
    purchase_ids: list[int] = []
    sale_ids: list[int] = []
    note: str = ""


@app.get("/api/matches")
def list_matches():
    with db.get_conn() as conn:
        groups = conn.execute("SELECT id, note FROM match_groups ORDER BY id").fetchall()
        out = []
        for g in groups:
            ps = [dict(r) for r in conn.execute(
                "SELECT p.* FROM match_group_purchases mp "
                "JOIN purchase_lines p ON p.id=mp.purchase_line_id "
                "WHERE mp.group_id=? ORDER BY p.invoice_date IS NULL, p.invoice_date, p.id",
                (g["id"],))]
            ss = [dict(r) for r in conn.execute(
                "SELECT s.* FROM match_group_sales ms "
                "JOIN sale_lines s ON s.id=ms.sale_line_id "
                "WHERE ms.group_id=? ORDER BY s.invoice_date IS NULL, s.invoice_date, s.id",
                (g["id"],))]
            pq = sum(p["qty"] or 0 for p in ps)
            sq = sum(s["qty"] or 0 for s in ss)
            out.append({"id": g["id"], "note": g["note"], "purchases": ps, "sales": ss,
                        "purchase_qty": pq, "sale_qty": sq, "qty_diff": pq - sq})
    return out


def _create_group(conn, purchase_ids, sale_ids, note=""):
    """Insert a group + memberships. Raises HTTPException on validation failure."""
    if not purchase_ids or not sale_ids:
        raise HTTPException(400, "لازم سطر مشتريات وسطر مبيعات على الأقل في كل مطابقة")
    for pid in purchase_ids:
        if not conn.execute("SELECT 1 FROM purchase_lines WHERE id=?", (pid,)).fetchone():
            raise HTTPException(404, f"سطر المشتريات {pid} غير موجود")
        if conn.execute("SELECT 1 FROM match_group_purchases WHERE purchase_line_id=?",
                        (pid,)).fetchone():
            raise HTTPException(409, "أحد سطور المشتريات مرتبط بمطابقة أخرى")
    for sid in sale_ids:
        if not conn.execute("SELECT 1 FROM sale_lines WHERE id=?", (sid,)).fetchone():
            raise HTTPException(404, f"سطر المبيعات {sid} غير موجود")
        if conn.execute("SELECT 1 FROM match_group_sales WHERE sale_line_id=?",
                        (sid,)).fetchone():
            raise HTTPException(409, "أحد سطور المبيعات مرتبط بمطابقة أخرى")
    cur = conn.execute("INSERT INTO match_groups(note) VALUES (?)", (note,))
    gid = cur.lastrowid
    for pid in purchase_ids:
        conn.execute("INSERT INTO match_group_purchases(group_id, purchase_line_id) "
                     "VALUES (?,?)", (gid, pid))
    for sid in sale_ids:
        conn.execute("INSERT INTO match_group_sales(group_id, sale_line_id) "
                     "VALUES (?,?)", (gid, sid))
    return gid


@app.post("/api/matches")
def create_match(body: MatchCreate):
    with db.get_conn() as conn:
        gid = _create_group(conn, body.purchase_ids, body.sale_ids, body.note)
    return {"id": gid}


@app.delete("/api/matches/{group_id}")
def delete_match(group_id: int):
    with db.get_conn() as conn:
        cur = conn.execute("DELETE FROM match_groups WHERE id=?", (group_id,))
    if cur.rowcount == 0:
        raise HTTPException(404, "المطابقة غير موجودة")
    return {"deleted": group_id}
```

Replace `_unmatched`:

```python
def _unmatched(conn, kind):
    if kind == "purchase":
        rows = conn.execute(
            "SELECT t.* FROM purchase_lines t "
            "LEFT JOIN match_group_purchases m ON m.purchase_line_id=t.id "
            "WHERE m.purchase_line_id IS NULL").fetchall()
    else:
        rows = conn.execute(
            "SELECT t.* FROM sale_lines t "
            "LEFT JOIN match_group_sales m ON m.sale_line_id=t.id "
            "WHERE m.sale_line_id IS NULL").fetchall()
    return [dict(r) for r in rows]
```

Replace `list_lines` (the join/select and matched filter) so it returns `group_id`:

```python
@app.get("/api/lines")
def list_lines(kind: str, matched: bool | None = None, q: str = "", party: str = ""):
    table = _table(kind)
    if kind == "purchase":
        join = "LEFT JOIN match_group_purchases m ON m.purchase_line_id = t.id"
        keycol = "m.purchase_line_id"
    else:
        join = "LEFT JOIN match_group_sales m ON m.sale_line_id = t.id"
        keycol = "m.sale_line_id"
    sql = f"SELECT t.*, m.group_id AS group_id FROM {table} t {join} WHERE 1=1"
    params: list = []
    if matched is True:
        sql += f" AND {keycol} IS NOT NULL"
    elif matched is False:
        sql += f" AND {keycol} IS NULL"
    if q:
        sql += " AND (t.item LIKE ? OR t.invoice_no LIKE ? OR t.party LIKE ?)"
        params += [f"%{q}%"] * 3
    if party:
        sql += " AND t.party LIKE ?"
        params.append(f"%{party}%")
    sql += " ORDER BY t.invoice_date IS NULL, t.invoice_date, t.id"
    with db.get_conn() as conn:
        rows = conn.execute(sql, params).fetchall()
    return [_row_to_dict(r) for r in rows]
```

In `dashboard`, replace the three match-related queries:

```python
        unmatched_p = conn.execute(
            "SELECT COUNT(*) c FROM purchase_lines t "
            "LEFT JOIN match_group_purchases m ON m.purchase_line_id=t.id "
            "WHERE m.purchase_line_id IS NULL").fetchone()["c"]
        unmatched_s = conn.execute(
            "SELECT COUNT(*) c FROM sale_lines t "
            "LEFT JOIN match_group_sales m ON m.sale_line_id=t.id "
            "WHERE m.sale_line_id IS NULL").fetchone()["c"]
        mismatch = conn.execute(
            "SELECT COUNT(*) c FROM ("
            "  SELECT g.id,"
            "   COALESCE((SELECT SUM(p.qty) FROM match_group_purchases mp "
            "     JOIN purchase_lines p ON p.id=mp.purchase_line_id "
            "     WHERE mp.group_id=g.id),0) AS pq,"
            "   COALESCE((SELECT SUM(s.qty) FROM match_group_sales ms "
            "     JOIN sale_lines s ON s.id=ms.sale_line_id "
            "     WHERE ms.group_id=g.id),0) AS sq"
            "  FROM match_groups g) x WHERE ABS(pq - sq) > 0.001").fetchone()["c"]
        matches_count = conn.execute("SELECT COUNT(*) c FROM match_groups").fetchone()["c"]
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_api.py -v`
Expected: PASS (the auto/accept test stays red until Task 5 — run only `-k "TestMatching and not auto"` if isolating; otherwise expect `test_auto_match_suggests_and_accepts` to fail here).

Run: `python -m pytest "tests/test_api.py::TestMatching::test_manual_match_and_unlink" "tests/test_api.py::TestMatching::test_line_cannot_match_twice" "tests/test_api.py::TestMatching::test_group_two_purchases_one_sale" "tests/test_api.py::TestMatching::test_match_requires_both_sides" "tests/test_api.py::TestMatching::test_unmatched_filter_excludes_grouped_lines" -v`
Expected: PASS (5 passed).

- [ ] **Step 5: Commit**

```bash
git add app/main.py tests/test_api.py
git commit -m "feat(api): group CRUD, group-aware line/dashboard queries"
```

---

### Task 5: Auto-discover + accept endpoints

**Files:**
- Modify: `app/main.py` (`auto_match_endpoint`, `AcceptBody`/`AcceptGroup`, `accept_matches`)
- Test: `tests/test_api.py` (update `test_auto_match_suggests_and_accepts`; add group-accept test)

- [ ] **Step 1: Write the failing tests**

Replace `test_auto_match_suggests_and_accepts` and add a group-discovery acceptance test:

```python
    def test_auto_match_suggests_and_accepts(self, client):
        p = _purchase(client, qty=7777, price=9, item="صنف فريد للاختبار")
        s = _sale(client, qty=7777, price=11, item="صنف فريد للاختبار")
        sugg = client.post("/api/matches/auto").json()
        mine = [m for m in sugg if p["id"] in m["purchase_ids"]]
        assert mine and mine[0]["sale_ids"] == [s["id"]]
        r = client.post("/api/matches/accept", json={"groups": [
            {"purchase_ids": [p["id"]], "sale_ids": [s["id"]]}]})
        assert r.status_code == 200
        assert r.json()["created"] == 1

    def test_auto_discovers_and_accepts_group(self, client):
        p1 = _purchase(client, qty=8000, item="صنف تجميعي")
        p2 = _purchase(client, qty=2000, item="صنف تجميعي")
        s = _sale(client, qty=10000, item="صنف تجميعي")
        sugg = client.post("/api/matches/auto").json()
        grp = [m for m in sugg if set(m["purchase_ids"]) == {p1["id"], p2["id"]}]
        assert grp and grp[0]["sale_ids"] == [s["id"]]
        r = client.post("/api/matches/accept", json={"groups": [
            {"purchase_ids": [p1["id"], p2["id"]], "sale_ids": [s["id"]]}]})
        assert r.json()["created"] == 1
        assert [m for m in client.get("/api/matches").json()
                if {pp["id"] for pp in m["purchases"]} == {p1["id"], p2["id"]}]
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_api.py -k auto -v`
Expected: FAIL — suggestions lack `purchase_ids`; accept still expects `pairs`.

- [ ] **Step 3: Update the endpoints**

In `app/main.py`, replace `auto_match_endpoint`, the accept models, and `accept_matches`:

```python
@app.post("/api/matches/auto")
def auto_match_endpoint():
    with db.get_conn() as conn:
        purchases = _unmatched(conn, "purchase")
        sales = _unmatched(conn, "sale")
    for r in purchases + sales:
        r["party"] = r.get("party") or ""
    suggestions = auto_match(purchases, sales)
    p_by_id = {p["id"]: p for p in purchases}
    s_by_id = {s["id"]: s for s in sales}
    for m in suggestions:
        m["purchases"] = [p_by_id[i] for i in m["purchase_ids"]]
        m["sales"] = [s_by_id[i] for i in m["sale_ids"]]
    return suggestions


class AcceptGroup(BaseModel):
    purchase_ids: list[int] = []
    sale_ids: list[int] = []
    note: str = ""


class AcceptBody(BaseModel):
    groups: list[AcceptGroup]


@app.post("/api/matches/accept")
def accept_matches(body: AcceptBody):
    created = 0
    with db.get_conn() as conn:
        for grp in body.groups:
            if not grp.purchase_ids or not grp.sale_ids:
                continue
            taken = any(conn.execute(
                "SELECT 1 FROM match_group_purchases WHERE purchase_line_id=?",
                (pid,)).fetchone() for pid in grp.purchase_ids) or any(conn.execute(
                "SELECT 1 FROM match_group_sales WHERE sale_line_id=?",
                (sid,)).fetchone() for sid in grp.sale_ids)
            if taken:
                continue
            cur = conn.execute("INSERT INTO match_groups(note) VALUES (?)", (grp.note,))
            gid = cur.lastrowid
            for pid in grp.purchase_ids:
                conn.execute("INSERT INTO match_group_purchases(group_id, purchase_line_id) "
                             "VALUES (?,?)", (gid, pid))
            for sid in grp.sale_ids:
                conn.execute("INSERT INTO match_group_sales(group_id, sale_line_id) "
                             "VALUES (?,?)", (gid, sid))
            created += 1
    return {"created": created}
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_api.py -v`
Expected: PASS (whole module green).

- [ ] **Step 5: Commit**

```bash
git add app/main.py tests/test_api.py
git commit -m "feat(api): group-aware auto-discovery and accept endpoints"
```

---

### Task 6: Excel export as group blocks

**Files:**
- Modify: `app/excel_io.py` (`export_workbook` signature + row building)
- Test: `tests/test_excel_io.py` (update fixture/real-file calls; add multi-line group total test)

- [ ] **Step 1: Update existing export tests + add a group test**

In `tests/test_excel_io.py`, change the two `export_workbook(...)` calls to pass **groups** (lists of id-lists), and add a totals test:

```python
# in small_export fixture: replace the export call
        export_workbook(purchases, sales, [([1], [11])], path, vat_rate=0.14)

# in test_real_file_round_trip_totals: replace the pairs/export lines
        groups = [([pi + 1], [1000 + si + 1]) for pi, si in imported["matches"]]
        export_workbook(imported["purchases"], imported["sales"], groups, path, vat_rate=0.14)

# add a new test in TestExport:
    def test_multiline_group_keeps_subtotals(self, tmp_path):
        purchases = [
            _line(1, "2026-06-05", "100", "سيفتى باك", "نموذج 2", 28634, 9.95, 0),
            _line(2, "2026-06-06", "101", "سيفتى باك", "نموذج 2", 916, 9.95, 0),
        ]
        sales = [_line(10, "2026-05-20", "200", "الفيشاوي", "نموذج 2", 29550, 12, 0)]
        path = str(tmp_path / "grp.xlsx")
        export_workbook(purchases, sales, [([1, 2], [10])], path, vat_rate=0.14)
        wb = load_workbook(path)
        ws = wb.active
        # two stacked rows: purchases in E6,E7; the single sale in N6
        assert ws["E6"].value == 28634
        assert ws["E7"].value == 916
        assert ws["N6"].value == 29550
        assert ws["N7"].value is None
        # column SUBTOTALs cover the whole block and reconcile
        assert ws["E4"].value == "=SUBTOTAL(9,E6:E7)"
        assert ws["N4"].value == "=SUBTOTAL(9,N6:N7)"
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_excel_io.py -v`
Expected: FAIL — `export_workbook` still treats arg as scalar pairs (`p_by_id[pid]` with a list key → TypeError/KeyError).

- [ ] **Step 3: Rewrite the row-building part of `export_workbook`**

In `app/excel_io.py`, rename the third parameter and replace the block that builds `rows` (everything from `p_by_id = ...` down to the `rows += [(None, s) ...]` line). Keep the rest of the function (worksheet setup, formulas, the `for i, (p, s) in enumerate(rows)` writer loop) unchanged.

```python
def export_workbook(purchases: list[dict], sales: list[dict],
                    groups: list[tuple], path: str, vat_rate: float = 0.14):
    p_by_id = {p["id"]: p for p in purchases}
    s_by_id = {s["id"]: s for s in sales}
    matched_p, matched_s = set(), set()
    blocks: list[tuple] = []        # each = (list[purchase dict], list[sale dict])
    for p_ids, s_ids in groups:
        ps = [p_by_id[i] for i in p_ids if i in p_by_id]
        ss = [s_by_id[i] for i in s_ids if i in s_by_id]
        if not ps and not ss:
            continue
        matched_p.update(p["id"] for p in ps)
        matched_s.update(s["id"] for s in ss)
        blocks.append((ps, ss))
    blocks += [([p], []) for p in purchases if p["id"] not in matched_p]
    blocks += [([], [s]) for s in sales if s["id"] not in matched_s]

    rows: list[tuple] = []          # each = (purchase dict | None, sale dict | None)
    for ps, ss in blocks:
        for i in range(max(len(ps), len(ss))):
            rows.append((ps[i] if i < len(ps) else None,
                         ss[i] if i < len(ss) else None))
```

(The existing code after this point — `wb = Workbook()`, the SUBTOTAL/formula setup using `last = 5 + max(len(rows), 1)`, and the writer loop — needs no change.)

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_excel_io.py -v`
Expected: PASS (all, including `test_multiline_group_keeps_subtotals`).

- [ ] **Step 5: Commit**

```bash
git add app/excel_io.py tests/test_excel_io.py
git commit -m "feat(excel): export match groups as stacked blocks (totals stay exact)"
```

---

### Task 7: Wire export/import to groups

**Files:**
- Modify: `app/main.py` (`export_excel`, `import_excel`)
- Test: `tests/test_api.py` (add import→group test)

- [ ] **Step 1: Write the failing test**

```python
# in tests/test_api.py, add to TestSettingsAndExport (or a new class):
    def test_export_still_returns_workbook_with_group(self, client):
        # create a 2:1 group, then export — must still produce a valid xlsx
        p1 = _purchase(client, qty=600, item="تصدير صنف")
        p2 = _purchase(client, qty=400, item="تصدير صنف")
        s = _sale(client, qty=1000, item="تصدير صنف")
        client.post("/api/matches", json={
            "purchase_ids": [p1["id"], p2["id"]], "sale_ids": [s["id"]]})
        r = client.get("/api/export/excel")
        assert r.status_code == 200
        assert r.content[:2] == b"PK"
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_api.py -k export_still -v`
Expected: FAIL — `export_excel` still reads `FROM matches`.

- [ ] **Step 3: Update `export_excel` and `import_excel`**

In `app/main.py`, replace the data-gathering block of `export_excel`:

```python
@app.get("/api/export/excel")
def export_excel():
    with db.get_conn() as conn:
        purchases = [dict(r) for r in conn.execute("SELECT * FROM purchase_lines")]
        sales = [dict(r) for r in conn.execute("SELECT * FROM sale_lines")]
        group_rows = conn.execute(
            "SELECT g.id, (SELECT MIN(p.invoice_date) FROM match_group_purchases mp "
            "   JOIN purchase_lines p ON p.id=mp.purchase_line_id WHERE mp.group_id=g.id) d "
            "FROM match_groups g ORDER BY d IS NULL, d, g.id").fetchall()
        groups = []
        for g in group_rows:
            p_ids = [r["purchase_line_id"] for r in conn.execute(
                "SELECT purchase_line_id FROM match_group_purchases WHERE group_id=? "
                "ORDER BY purchase_line_id", (g["id"],))]
            s_ids = [r["sale_line_id"] for r in conn.execute(
                "SELECT sale_line_id FROM match_group_sales WHERE group_id=? "
                "ORDER BY sale_line_id", (g["id"],))]
            groups.append((p_ids, s_ids))
    exports = Path(db.DB_PATH).parent / "exports"
    exports.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y-%m-%d_%H%M")
    out = exports / f"matching_{stamp}.xlsx"
    export_workbook(purchases, sales, groups, str(out), vat_rate=db.vat_rate())
    arabic_name = f"مطابقة المشتريات والمبيعات {stamp}.xlsx"
    quoted = urllib.parse.quote(arabic_name)
    return FileResponse(
        str(out),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition":
                 f"attachment; filename=matching_{stamp}.xlsx; filename*=UTF-8''{quoted}"})
```

In `import_excel`, replace the replace-mode delete and the match-insert loop:

```python
        if mode == "replace":
            conn.execute("DELETE FROM match_groups")   # cascades to memberships
            conn.execute("DELETE FROM purchase_lines")
            conn.execute("DELETE FROM sale_lines")
```

```python
        for pi, si in data["matches"]:
            cur = conn.execute("INSERT INTO match_groups(note) VALUES ('')")
            gid = cur.lastrowid
            conn.execute("INSERT INTO match_group_purchases(group_id, purchase_line_id) "
                         "VALUES (?,?)", (gid, p_ids[pi]))
            conn.execute("INSERT INTO match_group_sales(group_id, sale_line_id) "
                         "VALUES (?,?)", (gid, s_ids[si]))
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_api.py -v`
Expected: PASS (whole module).

- [ ] **Step 5: Commit**

```bash
git add app/main.py tests/test_api.py
git commit -m "feat(api): export/import use match groups (1:1 group per imported pair)"
```

---

### Task 8: Matching screen markup (multi-select + live bar)

**Files:**
- Modify: `static/index.html` (the `#tab-matching` section)

- [ ] **Step 1: Update the unmatched tables and add the reconciliation bar**

In `static/index.html`, inside `#tab-matching`, replace the `.match-grid` block so the two unmatched tables drop the radio-only header cell, keep their IDs, and add a live totals bar + a header checkbox column. Replace the existing `<div class="match-grid"> ... </div>`:

```html
    <div class="match-grid">
      <div class="card">
        <div class="card-head"><h3>مشتريات غير مرتبطة <em class="count" id="cnt-unmatched-p"></em></h3></div>
        <div class="table-wrap tall"><table class="tbl selectable" id="tbl-unmatched-p">
          <thead><tr><th></th><th>التاريخ</th><th>رقم</th><th>المورد</th><th>الصنف</th><th>الكمية</th></tr></thead>
          <tbody></tbody>
        </table></div>
      </div>
      <div class="match-link-col">
        <button class="btn primary round" id="btn-manual-link" title="ربط المحدد">
          <svg viewBox="0 0 24 24"><path d="M3.9 12c0-1.71 1.39-3.1 3.1-3.1h4V7H7c-2.76 0-5 2.24-5 5s2.24 5 5 5h4v-1.9H7c-1.71 0-3.1-1.39-3.1-3.1zM8 13h8v-2H8v2zm9-6h-4v1.9h4c1.71 0 3.1 1.39 3.1 3.1s-1.39 3.1-3.1 3.1h-4V17h4c2.76 0 5-2.24 5-5s-2.24-5-5-5z"/></svg>
          ربط
        </button>
      </div>
      <div class="card">
        <div class="card-head"><h3>مبيعات غير مرتبطة <em class="count" id="cnt-unmatched-s"></em></h3></div>
        <div class="table-wrap tall"><table class="tbl selectable" id="tbl-unmatched-s">
          <thead><tr><th></th><th>التاريخ</th><th>رقم</th><th>المشتري</th><th>الصنف</th><th>الكمية</th></tr></thead>
          <tbody></tbody>
        </table></div>
      </div>
    </div>

    <div class="recon-bar" id="recon-bar" hidden>
      <span class="recon-item">كمية الشراء المحددة: <b id="recon-p">0</b></span>
      <span class="recon-item">كمية البيع المحددة: <b id="recon-s">0</b></span>
      <span class="recon-item">الفرق: <b id="recon-diff" class="recon-diff">0</b></span>
    </div>
```

- [ ] **Step 2: Verify markup loads (manual)**

Run: `python -c "import pathlib,re; html=pathlib.Path('static/index.html').read_text(encoding='utf-8'); assert 'recon-bar' in html and 'tbl-unmatched-p' in html; print('markup OK')"`
Expected: `markup OK`

- [ ] **Step 3: Commit**

```bash
git add static/index.html
git commit -m "feat(ui): matching screen reconciliation bar + multi-select tables"
```

---

### Task 9: Matching screen behaviour (multi-select, live totals, groups)

**Files:**
- Modify: `static/app.js` (the «المطابقة» section: state, `renderUnmatched`, selection handler, link button, `renderMatches`, `renderSuggestions`, `acceptPairs`; plus the `renderLines` pill in the lines section)

- [ ] **Step 1: Fix the lines-table "مرتبط" pill**

`list_lines` now returns `group_id` instead of `match_id`. In `renderLines` (the «سطور المشتريات/المبيعات» section, NOT the matching section), change the linked pill:

```javascript
      <td>${r.group_id ? '<span class="pill linked">مرتبط</span>' : ""}</td>
```

(was `${r.match_id ? ...}`).

- [ ] **Step 2: Replace the matching block in `app.js`**

Replace everything from `let suggestions = [];` down to the end of `acceptPairs` with the group-aware version:

```javascript
/* ـــــــــــــــــــــ المطابقة ـــــــــــــــــــــ */
let suggestions = [];
const selP = new Set();
const selS = new Set();

async function loadMatching() {
  try {
    const [up, us, groups] = await Promise.all([
      api("/api/lines?kind=purchase&matched=false"),
      api("/api/lines?kind=sale&matched=false"),
      api("/api/matches"),
    ]);
    selP.clear(); selS.clear();
    renderUnmatched("#tbl-unmatched-p", up, "p");
    renderUnmatched("#tbl-unmatched-s", us, "s");
    $("#cnt-unmatched-p").textContent = `(${up.length})`;
    $("#cnt-unmatched-s").textContent = `(${us.length})`;
    renderRecon();
    renderMatches(groups);
  } catch (e) { toast(e.message, true); }
}

function renderUnmatched(sel, rows, side) {
  const set = side === "p" ? selP : selS;
  const tbody = $(sel + " tbody");
  tbody.innerHTML = rows.map((r) => `
    <tr data-select="${side}:${r.id}" data-qty="${r.qty || 0}" class="${set.has(r.id) ? "selected" : ""}">
      <td><input type="checkbox" ${set.has(r.id) ? "checked" : ""}></td>
      <td class="num">${esc(r.invoice_date || "—")}</td>
      <td>${esc(r.invoice_no)}</td>
      <td>${esc(r.party)}</td>
      <td>${esc(r.item)}</td>
      <td class="num">${qty(r.qty)}</td>
    </tr>`).join("") || `<tr><td colspan="6" class="empty">لا يوجد</td></tr>`;
}

function sumSelected(sel, set) {
  let t = 0;
  $$(sel + " tbody tr[data-select]").forEach((tr) => {
    const [, id] = tr.dataset.select.split(":");
    if (set.has(+id)) t += parseFloat(tr.dataset.qty) || 0;
  });
  return t;
}

function renderRecon() {
  const bar = $("#recon-bar");
  const any = selP.size || selS.size;
  bar.hidden = !any;
  if (!any) return;
  const p = sumSelected("#tbl-unmatched-p", selP);
  const s = sumSelected("#tbl-unmatched-s", selS);
  $("#recon-p").textContent = qty(p);
  $("#recon-s").textContent = qty(s);
  const diff = p - s;
  const el = $("#recon-diff");
  el.textContent = qty(diff);
  el.classList.toggle("ok", Math.abs(diff) < 1e-9);
  el.classList.toggle("bad", Math.abs(diff) >= 1e-9);
}

document.addEventListener("click", (e) => {
  const tr = e.target.closest("tr[data-select]");
  if (!tr) return;
  const [side, id] = tr.dataset.select.split(":");
  const set = side === "p" ? selP : selS;
  const n = +id;
  if (set.has(n)) set.delete(n); else set.add(n);
  tr.classList.toggle("selected", set.has(n));
  const box = tr.querySelector("input[type=checkbox]");
  if (box) box.checked = set.has(n);
  renderRecon();
});

$("#btn-manual-link").addEventListener("click", () => {
  if (!selP.size || !selS.size)
    return toast("اختر سطراً واحداً على الأقل من المشتريات ومن المبيعات", true);
  api("/api/matches", { method: "POST", json: {
    purchase_ids: [...selP], sale_ids: [...selS],
  } })
    .then(() => { toast("تم الربط ✓"); loadMatching(); loadDashboard(); })
    .catch((err) => toast(err.message, true));
});

// Purchase side = 5 cells (incl. item); sale side = 4 cells (item shown once,
// on the purchase side) → 5 + 4 + qty_diff + actions = 11 columns, matching the
// existing #tbl-matches header. Do NOT add an item cell to saleCells.
function purchaseCells(p) {
  return `<td class="num">${esc(p ? p.invoice_date || "—" : "")}</td>
    <td>${esc(p ? p.invoice_no : "")}</td>
    <td>${esc(p ? p.party : "")}</td>
    <td>${esc(p ? p.item : "")}</td>
    <td class="num">${p ? qty(p.qty) : ""}</td>`;
}

function saleCells(s) {
  return `<td class="num">${esc(s ? s.invoice_date || "—" : "")}</td>
    <td>${esc(s ? s.invoice_no : "")}</td>
    <td>${esc(s ? s.party : "")}</td>
    <td class="num">${s ? qty(s.qty) : ""}</td>`;
}

function renderMatches(groups) {
  $("#cnt-matches").textContent = `(${groups.length})`;
  const html = groups.map((g) => {
    const h = Math.max(g.purchases.length, g.sales.length, 1);
    let block = "";
    for (let i = 0; i < h; i++) {
      const p = g.purchases[i] || null;
      const s = g.sales[i] || null;
      const first = i === 0;
      block += `<tr class="grp${first ? " grp-first" : ""}">
        ${purchaseCells(p)}${saleCells(s)}
        ${first ? `<td class="num" rowspan="${h}">${g.qty_diff ? signedCell(g.qty_diff, qty) : '<span class="num">0</span>'}</td>
        <td class="actions" rowspan="${h}"><button class="btn danger tiny" data-unlink="${g.id}">فك الربط</button></td>` : ""}
      </tr>`;
    }
    return block;
  }).join("");
  $("#tbl-matches tbody").innerHTML = html ||
    `<tr><td colspan="11" class="empty">لا توجد مطابقات بعد</td></tr>`;
}

document.addEventListener("click", (e) => {
  const u = e.target.closest("[data-unlink]");
  if (u) {
    api(`/api/matches/${u.dataset.unlink}`, { method: "DELETE" })
      .then(() => { toast("تم فك الربط"); loadMatching(); loadDashboard(); })
      .catch((err) => toast(err.message, true));
  }
});

$("#btn-auto-match").addEventListener("click", async () => {
  const btn = $("#btn-auto-match");
  btn.disabled = true;
  try {
    suggestions = await api("/api/matches/auto", { method: "POST" });
    renderSuggestions();
    if (!suggestions.length) toast("لا توجد مقترحات — كل السطور المتشابهة مرتبطة بالفعل");
  } catch (e) { toast(e.message, true); }
  btn.disabled = false;
});

function partyList(lines) {
  return lines.map((l) => `${esc(l.invoice_no)} · ${esc(l.party)}`).join("<br>");
}

function renderSuggestions() {
  $("#suggestions-empty").hidden = suggestions.length > 0;
  $("#suggestions-wrap").hidden = suggestions.length === 0;
  $("#btn-accept-all").hidden = suggestions.length === 0;
  $("#tbl-suggestions tbody").innerHTML = suggestions.map((m, i) => {
    const pq = m.purchases.reduce((a, p) => a + (p.qty || 0), 0);
    const sq = m.sales.reduce((a, s) => a + (s.qty || 0), 0);
    const multi = m.purchases.length > 1 || m.sales.length > 1;
    return `<tr>
      <td><span class="pill ${m.score >= 85 ? "score-hi" : "score-md"}">${m.score}%</span>${multi ? ' <span class="pill linked">مجموعة</span>' : ""}</td>
      <td>${partyList(m.purchases)}</td>
      <td>${esc(m.purchases[0] ? m.purchases[0].item : "")}</td>
      <td class="num">${qty(pq)}</td>
      <td>${partyList(m.sales)}</td>
      <td class="num">${qty(sq)}</td>
      <td>${m.qty_diff ? signedCell(m.qty_diff, qty) : "0"}</td>
      <td class="actions">
        <button class="btn tiny primary" data-accept="${i}">اعتماد</button>
        <button class="btn tiny ghost" data-dismiss="${i}">تجاهل</button>
      </td>
    </tr>`;
  }).join("");
}

document.addEventListener("click", (e) => {
  const acc = e.target.closest("[data-accept]");
  if (acc) acceptGroups([suggestions[+acc.dataset.accept]]);
  const dis = e.target.closest("[data-dismiss]");
  if (dis) { suggestions.splice(+dis.dataset.dismiss, 1); renderSuggestions(); }
});

$("#btn-accept-all").addEventListener("click", () => acceptGroups(suggestions));

function acceptGroups(list) {
  const groups = list.map((m) => ({ purchase_ids: m.purchase_ids, sale_ids: m.sale_ids }));
  api("/api/matches/accept", { method: "POST", json: { groups } })
    .then((r) => {
      toast(`تم اعتماد ${r.created} مطابقة ✓`);
      suggestions = suggestions.filter((m) => !list.includes(m));
      renderSuggestions();
      loadMatching();
      loadDashboard();
    })
    .catch((err) => toast(err.message, true));
}
```

- [ ] **Step 3: Sanity-check the JS parses**

Run: `node --check static/app.js`
Expected: no output (exit 0). If `node` is unavailable, skip — the UI smoke test in Task 11 covers it.

- [ ] **Step 4: Commit**

```bash
git add static/app.js
git commit -m "feat(ui): multi-select linking, live reconciliation, group rendering"
```

---

### Task 10: Styles for the reconciliation bar and group blocks

**Files:**
- Modify: `static/style.css` (append)

- [ ] **Step 1: Append styles**

Add to the end of `static/style.css`:

```css
/* ـــــ شريط المصالحة ـــــ */
.recon-bar {
  display: flex; gap: 24px; align-items: center; flex-wrap: wrap;
  margin: 12px 0; padding: 12px 18px;
  background: var(--card, #fff); border: 1px solid #e5e7eb; border-radius: 12px;
  box-shadow: 0 1px 2px rgba(0,0,0,.04);
}
.recon-item { color: #6b7280; font-size: 14px; }
.recon-item b { color: #111827; font-variant-numeric: tabular-nums; }
.recon-diff.ok { color: #15803d; }
.recon-diff.bad { color: #b91c1c; }

/* ـــــ بلوكات المجموعات في جدول المطابقات ـــــ */
#tbl-matches tr.grp-first > td { border-top: 2px solid #d1d5db; }
#tbl-matches tr.grp:not(.grp-first) > td { border-top: 0; }
```

- [ ] **Step 2: Verify CSS file still loads**

Run: `python -c "import pathlib; css=pathlib.Path('static/style.css').read_text(encoding='utf-8'); assert '.recon-bar' in css; print('css OK')"`
Expected: `css OK`

- [ ] **Step 3: Commit**

```bash
git add static/style.css
git commit -m "style(ui): reconciliation bar and match-group block borders"
```

---

### Task 11: Full verification

**Files:** none (runs the suite + smoke + export audit)

- [ ] **Step 1: Run the whole pytest suite**

Run: `python -m pytest -q`
Expected: all tests pass (the original 48 + the new group/migration tests). Investigate any failure before continuing.

- [ ] **Step 2: Launch the app and load the real workbook**

Run (PowerShell, from repo root, a separate terminal):
```
$env:PBK_DB="data/pbk.db"; python -m uvicorn app.main:app --port 8077
```
Then in another shell import the bundled workbook so the smoke test has data:
```
curl.exe -s -F "file=@مطابقة المشتريات والمبيعات.xlsx" -F "mode=replace" http://127.0.0.1:8077/api/import/excel
```

- [ ] **Step 3: Run the UI smoke test**

Run: `python tests/ui_smoke.py`
Expected: `FAILURES: none`. The matches table must still show **52 rows** (52 imported 1:1 groups) and `matches_count` 52 — confirming the group rendering is backward-compatible.

- [ ] **Step 4: Run the export formula audit**

Run: `python tests/verify_export.py`
Expected: all `PASS`, `problems: 0` — group export keeps `J2`/`R2`/`SUBTOTAL` totals exact.

- [ ] **Step 5: Manual group check (browser)**

Open `http://127.0.0.1:8077`, go to المطابقة, tick two purchase rows of the same item whose quantities sum to a sale row, confirm the reconciliation bar shows الفرق = 0 (green), click ربط, and confirm the new group appears in المطابقات الحالية with one فك الربط button spanning both rows.

- [ ] **Step 6: Commit any fixes**

```bash
git add -A
git commit -m "test: full-suite + UI smoke + export audit green for group matching"
```

---

## Self-Review

**Spec coverage:** schema+migration (Tasks 1-2) ✓; group CRUD/validation (Task 4) ✓; auto-discovery N:1/1:N (Tasks 3,5) ✓; line/dashboard group-awareness (Task 4) ✓; Excel block export with exact totals (Tasks 6-7) ✓; multi-select + live reconciliation UI (Tasks 8-10) ✓; tests incl. flagship 28,634+916 case (Tasks 3-7) ✓; import 1:1-group behaviour (Task 7) ✓. The documented Excel re-import limitation needs no code.

**Type consistency:** suggestion dicts use `purchase_ids`/`sale_ids`/`score`/`qty_diff` everywhere (matching.py → auto endpoint → app.js → accept). Group dicts from `/api/matches` use `purchases`/`sales`/`purchase_qty`/`sale_qty`/`qty_diff`/`id`/`note` (list_matches → app.js renderMatches). `export_workbook` third arg is `groups: list[(p_ids, s_ids)]` (excel_io ← export_excel ← test fixtures). Line rows expose `group_id` (list_lines → app.js pill).

**Placeholder scan:** none — every step shows full code or an exact command.
