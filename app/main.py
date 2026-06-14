# -*- coding: utf-8 -*-
"""FastAPI app: Arabic purchases↔sales matching tool with ETA integration."""
from __future__ import annotations

import json
import tempfile
import urllib.parse
from datetime import date, datetime
from pathlib import Path

from fastapi import FastAPI, HTTPException, UploadFile, File, Form
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field

from . import db
from .eta_client import ETAClient, ETAError
from .excel_io import export_stock_card, export_workbook, import_workbook
from .matching import auto_match
from .stock import build_item_card, build_overview

app = FastAPI(title="مطابقة المشتريات والمبيعات")
db.init_db()

_TABLES = {"purchase": "purchase_lines", "sale": "sale_lines"}


def _table(kind: str) -> str:
    if kind not in _TABLES:
        raise HTTPException(400, "kind يجب أن يكون purchase أو sale")
    return _TABLES[kind]


def _row_to_dict(row) -> dict:
    return dict(row) if row is not None else None


# --------------------------------------------------------------------------- lines
class LineCreate(BaseModel):
    kind: str
    invoice_date: str | None = None
    invoice_no: str = ""
    party: str = ""
    item: str = ""
    qty: float
    unit_price: float = 0
    vat: float | None = None
    doc_type: str = "i"
    internal_ref: str = ""
    note: str = ""


class LineUpdate(BaseModel):
    invoice_date: str | None = None
    invoice_no: str | None = None
    party: str | None = None
    item: str | None = None
    qty: float | None = None
    unit_price: float | None = None
    vat: float | None = None
    doc_type: str | None = None
    internal_ref: str | None = None
    note: str | None = None


@app.post("/api/lines")
def create_line(body: LineCreate):
    table = _table(body.kind)
    vat = body.vat if body.vat is not None else body.qty * body.unit_price * db.vat_rate()
    cols = {"invoice_date": body.invoice_date, "invoice_no": body.invoice_no,
            "party": body.party, "item": body.item, "qty": body.qty,
            "unit_price": body.unit_price, "vat": vat, "doc_type": body.doc_type,
            "source": "manual"}
    if body.kind == "sale":
        cols.update(internal_ref=body.internal_ref, note=body.note)
    with db.get_conn() as conn:
        cur = conn.execute(
            f"INSERT INTO {table} ({','.join(cols)}) VALUES ({','.join('?' * len(cols))})",
            list(cols.values()))
        row = conn.execute(f"SELECT * FROM {table} WHERE id=?", (cur.lastrowid,)).fetchone()
    return _row_to_dict(row)


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


@app.put("/api/lines/{kind}/{line_id}")
def update_line(kind: str, line_id: int, body: LineUpdate):
    table = _table(kind)
    changes = {k: v for k, v in body.model_dump(exclude_unset=True).items()
               if not (kind == "purchase" and k in ("internal_ref", "note"))}
    with db.get_conn() as conn:
        row = conn.execute(f"SELECT * FROM {table} WHERE id=?", (line_id,)).fetchone()
        if not row:
            raise HTTPException(404, "السطر غير موجود")
        # manual rows keep VAT at the configured rate unless explicitly overridden
        if "vat" not in changes and row["source"] == "manual" and \
                ("qty" in changes or "unit_price" in changes):
            qty = changes.get("qty", row["qty"])
            price = changes.get("unit_price", row["unit_price"])
            old_default = abs(row["vat"] - row["qty"] * row["unit_price"] * db.vat_rate()) <= 0.005
            if old_default:
                changes["vat"] = qty * price * db.vat_rate()
        if changes:
            sets = ", ".join(f"{k}=?" for k in changes)
            conn.execute(f"UPDATE {table} SET {sets} WHERE id=?",
                         [*changes.values(), line_id])
        row = conn.execute(f"SELECT * FROM {table} WHERE id=?", (line_id,)).fetchone()
    return _row_to_dict(row)


@app.delete("/api/lines/{kind}/{line_id}")
def delete_line(kind: str, line_id: int):
    table = _table(kind)
    with db.get_conn() as conn:
        cur = conn.execute(f"DELETE FROM {table} WHERE id=?", (line_id,))
    if cur.rowcount == 0:
        raise HTTPException(404, "السطر غير موجود")
    return {"deleted": line_id}


# --------------------------------------------------------------------------- matches
class MatchCreate(BaseModel):
    purchase_ids: list[int] = []
    sale_ids: list[int] = []
    note: str = ""


class AcceptBody(BaseModel):
    pairs: list[MatchCreate]


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
        m["purchase"] = p_by_id[m["purchase_id"]]
        m["sale"] = s_by_id[m["sale_id"]]
    return suggestions


@app.post("/api/matches/accept")
def accept_matches(body: AcceptBody):
    created = 0
    with db.get_conn() as conn:
        for pair in body.pairs:
            cur = conn.execute(
                "INSERT INTO matches(purchase_line_id, sale_line_id, note) VALUES (?,?,?) "
                "ON CONFLICT DO NOTHING",
                (pair.purchase_id, pair.sale_id, pair.note))
            created += cur.rowcount or 0
    return {"created": created}


# --------------------------------------------------------------------------- dashboard
@app.get("/api/dashboard")
def dashboard():
    with db.get_conn() as conn:
        p = conn.execute(
            "SELECT COUNT(*) c, COALESCE(SUM(qty*unit_price),0) net, "
            "COALESCE(SUM(vat),0) vat FROM purchase_lines").fetchone()
        s = conn.execute(
            "SELECT COUNT(*) c, COALESCE(SUM(qty*unit_price),0) net, "
            "COALESCE(SUM(vat),0) vat FROM sale_lines").fetchone()
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
    purchases = {"count": p["c"], "net": p["net"], "vat": p["vat"],
                 "total": p["net"] + p["vat"]}
    sales = {"count": s["c"], "net": s["net"], "vat": s["vat"],
             "total": s["net"] + s["vat"]}
    return {
        "purchases": purchases, "sales": sales,
        "vat_position": sales["vat"] - purchases["vat"],   # J2 = Q4-H4
        "gross_diff": sales["total"] - purchases["total"],  # R2 = R4-I4
        "net_margin": sales["net"] - purchases["net"],
        "unmatched_purchases": unmatched_p, "unmatched_sales": unmatched_s,
        "qty_mismatch_count": mismatch, "matches_count": matches_count,
    }


# --------------------------------------------------------------------------- stock card
def _all_lines(conn):
    purchases = [dict(r) for r in conn.execute("SELECT * FROM purchase_lines")]
    sales = [dict(r) for r in conn.execute("SELECT * FROM sale_lines")]
    return purchases, sales


@app.get("/api/stock-card")
def stock_card_overview():
    with db.get_conn() as conn:
        purchases, sales = _all_lines(conn)
    return build_overview(purchases, sales)


@app.get("/api/stock-card/item")
def stock_card_item(key: str):
    with db.get_conn() as conn:
        purchases, sales = _all_lines(conn)
    return build_item_card(purchases, sales, key)


@app.get("/api/stock-card/export")
def stock_card_export(key: str = ""):
    with db.get_conn() as conn:
        purchases, sales = _all_lines(conn)
    overview = build_overview(purchases, sales)
    if key:
        cards = [build_item_card(purchases, sales, key)]
        summary = None
    else:
        cards = [build_item_card(purchases, sales, o["item_key"]) for o in overview]
        summary = overview
    stamp = datetime.now().strftime("%Y-%m-%d_%H%M")
    out = _export_dir() / f"stock_card_{stamp}.xlsx"
    export_stock_card(summary, cards, str(out))
    arabic_name = f"كارت صنف {stamp}.xlsx"
    quoted = urllib.parse.quote(arabic_name)
    return FileResponse(
        str(out),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition":
                 f"attachment; filename=stock_card_{stamp}.xlsx; filename*=UTF-8''{quoted}"})


# --------------------------------------------------------------------------- excel
@app.post("/api/import/excel")
def import_excel(file: UploadFile = File(...), mode: str = Form("replace")):
    suffix = Path(file.filename or "upload.xlsx").suffix or ".xlsx"
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(file.file.read())
        tmp_path = tmp.name
    try:
        data = import_workbook(tmp_path)
    except Exception as e:
        raise HTTPException(400, f"تعذر قراءة الملف: {e}")
    finally:
        Path(tmp_path).unlink(missing_ok=True)

    with db.get_conn() as conn:
        if mode == "replace":
            conn.execute("DELETE FROM matches")
            conn.execute("DELETE FROM purchase_lines")
            conn.execute("DELETE FROM sale_lines")
        p_ids, s_ids = [], []
        for p in data["purchases"]:
            cur = conn.execute(
                "INSERT INTO purchase_lines(invoice_date,invoice_no,party,item,qty,"
                "unit_price,vat,doc_type,source) VALUES (?,?,?,?,?,?,?,?, 'excel')",
                (p["invoice_date"], p["invoice_no"], p["party"], p["item"],
                 p["qty"], p["unit_price"], p["vat"], p["doc_type"]))
            p_ids.append(cur.lastrowid)
        for s in data["sales"]:
            cur = conn.execute(
                "INSERT INTO sale_lines(invoice_date,invoice_no,party,item,qty,"
                "unit_price,vat,doc_type,source,internal_ref,note) "
                "VALUES (?,?,?,?,?,?,?,?, 'excel',?,?)",
                (s["invoice_date"], s["invoice_no"], s["party"], s["item"],
                 s["qty"], s["unit_price"], s["vat"], s["doc_type"],
                 s["internal_ref"], s["note"]))
            s_ids.append(cur.lastrowid)
        for pi, si in data["matches"]:
            conn.execute(
                "INSERT INTO matches(purchase_line_id, sale_line_id) VALUES (?,?)",
                (p_ids[pi], s_ids[si]))
    return {"purchases": len(p_ids), "sales": len(s_ids),
            "matches": len(data["matches"]), "warnings": data["warnings"]}


def _export_dir() -> Path:
    """Where generated workbooks are written.

    Locally (SQLite) we keep an archive next to the database. On a serverless /
    Postgres deployment the project filesystem is read-only, so we use the
    system temp dir (writable, ephemeral) — the file is streamed back in the
    same request, so persistence isn't needed there.
    """
    base = (Path(tempfile.gettempdir()) / "pbk_exports"
            if db.USE_PG else Path(db.DB_PATH).parent / "exports")
    base.mkdir(parents=True, exist_ok=True)
    return base


@app.get("/api/export/excel")
def export_excel():
    with db.get_conn() as conn:
        purchases = [dict(r) for r in conn.execute("SELECT * FROM purchase_lines")]
        sales = [dict(r) for r in conn.execute("SELECT * FROM sale_lines")]
        pairs = [(r["purchase_line_id"], r["sale_line_id"]) for r in conn.execute(
            "SELECT m.purchase_line_id, m.sale_line_id FROM matches m "
            "JOIN purchase_lines p ON p.id=m.purchase_line_id "
            "ORDER BY p.invoice_date IS NULL, p.invoice_date, m.id")]
    stamp = datetime.now().strftime("%Y-%m-%d_%H%M")
    out = _export_dir() / f"matching_{stamp}.xlsx"
    export_workbook(purchases, sales, pairs, str(out), vat_rate=db.vat_rate())
    arabic_name = f"مطابقة المشتريات والمبيعات {stamp}.xlsx"
    quoted = urllib.parse.quote(arabic_name)
    return FileResponse(
        str(out),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition":
                 f"attachment; filename=matching_{stamp}.xlsx; filename*=UTF-8''{quoted}"})


# --------------------------------------------------------------------------- settings
class SettingsBody(BaseModel):
    eta_env: str = Field(pattern="^(prod|preprod)$")
    eta_client_id: str = ""
    eta_client_secret: str = ""   # empty string = keep existing
    vat_rate: float = 0.14


@app.get("/api/settings")
def get_settings():
    return {
        "eta_env": db.get_setting("eta_env", "preprod"),
        "eta_client_id": db.get_setting("eta_client_id"),
        "has_secret": bool(db.get_setting("eta_client_secret")),
        "vat_rate": db.vat_rate(),
    }


def _clean_cred(value: str) -> str:
    """Trim whitespace and any stray BOM from a pasted credential — the usual
    cause of a spurious ETA 'login failed' after copy-paste."""
    return (value or "").strip().encode("utf-8").decode("utf-8-sig").strip()


@app.put("/api/settings")
def put_settings(body: SettingsBody):
    db.set_setting("eta_env", body.eta_env.strip())
    db.set_setting("eta_client_id", _clean_cred(body.eta_client_id))
    secret = _clean_cred(body.eta_client_secret)
    if secret:
        db.set_setting("eta_client_secret", secret)
    db.set_setting("vat_rate", str(body.vat_rate))
    return get_settings()


# --------------------------------------------------------------------------- ETA sync
def _eta_client() -> ETAClient:
    cid = _clean_cred(db.get_setting("eta_client_id"))
    secret = _clean_cred(db.get_setting("eta_client_secret"))
    if not cid or not secret:
        raise HTTPException(400, "أدخل بيانات الاتصال بمنظومة الفواتير أولاً من صفحة الإعدادات")
    return ETAClient(db.get_setting("eta_env", "preprod").strip(), cid, secret)


@app.post("/api/eta/test")
def eta_test():
    client = _eta_client()
    try:
        client.get_token()
    except ETAError as e:
        raise HTTPException(e.status_code or 502, str(e))
    return {"ok": True, "message": "تم الاتصال بمنظومة الفواتير الإلكترونية بنجاح"}


class SyncBody(BaseModel):
    date_from: date
    date_to: date
    directions: list[str] = ["Received", "Sent"]
    refresh: bool = False


# ETA sync runs as short, resumable chunks driven by client polling so it works
# on stateless serverless platforms (no long-lived background thread). All job
# state lives in the DB (a reserved settings row), surviving across invocations:
#   POST /api/eta/sync       logs in, searches ETA, and queues documents to fetch
#   POST /api/eta/sync/step  fetches the next few documents' lines, then persists
#                            progress; the browser calls it in a loop until done
#   GET  /api/eta/sync/status  returns the job state (read-only)
_SYNC_KEY = "__sync_state__"
_SYNC_CHUNK = 4   # documents fetched per /step call (each throttled ~2 s by ETA)
# kept server-side only — never returned to the browser
_SYNC_PRIVATE = {"queue", "token", "token_expiry", "last_request_at"}


def _default_sync() -> dict:
    return {"running": False, "log": [], "error": None,
            "started_at": None, "finished_at": None,
            "stats": {"documents": 0, "lines": 0},
            "total": 0, "queue": [], "last_request_at": 0.0,
            "token": None, "token_expiry": 0.0}


def _load_sync() -> dict:
    raw = db.get_setting(_SYNC_KEY, "")
    if raw:
        try:
            return json.loads(raw)
        except (ValueError, TypeError):
            pass
    return _default_sync()


def _save_sync(state: dict):
    db.set_setting(_SYNC_KEY, json.dumps(state, ensure_ascii=False, default=str))


def _public_sync(state: dict) -> dict:
    """The slice the UI needs — without the work queue or bearer token."""
    return {k: v for k, v in state.items() if k not in _SYNC_PRIVATE}


def _stamp(msg: str) -> str:
    return f"[{datetime.now().strftime('%H:%M:%S')}] {msg}"


def _seed_client(state: dict) -> ETAClient:
    """Build a client and restore throttle clock + cached token from job state."""
    client = _eta_client()
    client._last_request_at = state.get("last_request_at", 0.0)
    if state.get("token"):
        client._token = state["token"]
        client._token_expiry = state.get("token_expiry", 0.0)
    return client


def _upsert_eta_line(conn, table: str, ln: dict):
    extra_cols = ", internal_ref, note" if table == "sale_lines" else ""
    extra_vals = ", '', ''" if table == "sale_lines" else ""
    conn.execute(
        f"INSERT INTO {table}(invoice_date,invoice_no,party,item,qty,"
        f"unit_price,vat,doc_type,source,eta_uuid,eta_line_index{extra_cols}) "
        f"VALUES (?,?,?,?,?,?,?,?, 'eta', ?, ?{extra_vals}) "
        f"ON CONFLICT(eta_uuid, eta_line_index) DO UPDATE SET "
        f"invoice_date=excluded.invoice_date, invoice_no=excluded.invoice_no,"
        f"party=excluded.party, item=excluded.item, qty=excluded.qty,"
        f"unit_price=excluded.unit_price, vat=excluded.vat,"
        f"doc_type=excluded.doc_type",
        (ln["invoice_date"], ln["invoice_no"], ln["party"], ln["item"],
         ln["qty"], ln["unit_price"], ln["vat"], ln["doc_type"],
         ln["eta_uuid"], ln["eta_line_index"]))


@app.post("/api/eta/sync")
def eta_sync(body: SyncBody):
    client = _eta_client()  # validates credentials exist
    if _load_sync().get("running"):
        raise HTTPException(409, "هناك مزامنة قيد التنفيذ بالفعل")
    state = _default_sync()
    state["running"] = True
    state["started_at"] = datetime.now().isoformat(timespec="seconds")
    log = state["log"]
    try:
        log.append(_stamp("جارٍ تسجيل الدخول إلى منظومة الفواتير…"))
        client.get_token()
        log.append(_stamp("تم تسجيل الدخول بنجاح"))
        with db.get_conn() as conn:
            known = {r["eta_uuid"] for r in conn.execute(
                "SELECT eta_uuid FROM purchase_lines WHERE eta_uuid IS NOT NULL "
                "UNION SELECT eta_uuid FROM sale_lines WHERE eta_uuid IS NOT NULL")}
        for direction in body.directions:
            label = "المشتريات (الواردة)" if direction == "Received" else "المبيعات (الصادرة)"
            log.append(_stamp(
                f"جارٍ البحث عن مستندات {label} من {body.date_from} إلى {body.date_to}…"))
            summaries = list(client.search_documents(
                body.date_from, body.date_to, direction=direction,
                on_progress=lambda a, b, n: log.append(_stamp(f"  نافذة {a} → {b}: {n} مستند"))))
            log.append(_stamp(f"إجمالي مستندات {label}: {len(summaries)}"))
            for summary in summaries:
                if body.refresh or summary.get("uuid") not in known:
                    state["queue"].append({"direction": direction, "summary": summary})
        state["total"] = len(state["queue"])
        state["last_request_at"] = client._last_request_at
        state["token"] = client._token
        state["token_expiry"] = client._token_expiry
        if not state["queue"]:
            state["running"] = False
            state["finished_at"] = datetime.now().isoformat(timespec="seconds")
            log.append(_stamp("لا توجد مستندات جديدة للمزامنة ✓"))
    except Exception as e:  # noqa: BLE001 — surface everything to the UI log
        state["running"] = False
        state["error"] = str(e)
        state["finished_at"] = datetime.now().isoformat(timespec="seconds")
        log.append(_stamp(f"خطأ: {e}"))
    state["log"] = log[-80:]
    _save_sync(state)
    return {"started": True, "total": state["total"], "running": state["running"]}


@app.post("/api/eta/sync/step")
def eta_sync_step():
    state = _load_sync()
    if not state.get("running"):
        return _public_sync(state)
    queue = state.get("queue", [])
    log = state.setdefault("log", [])
    if not queue:
        state["running"] = False
        state["finished_at"] = datetime.now().isoformat(timespec="seconds")
        log.append(_stamp("اكتملت المزامنة ✓"))
        state["log"] = log[-80:]
        _save_sync(state)
        return _public_sync(state)
    try:
        client = _seed_client(state)
        for _ in range(_SYNC_CHUNK):
            if not queue:
                break
            item = queue.pop(0)
            table = "purchase_lines" if item["direction"] == "Received" else "sale_lines"
            lines = client.get_document_lines(item["summary"], item["direction"])
            with db.get_conn() as conn:
                for ln in lines:
                    _upsert_eta_line(conn, table, ln)
            state["stats"]["documents"] += 1
            state["stats"]["lines"] += len(lines)
        state["last_request_at"] = client._last_request_at
        state["token"] = client._token
        state["token_expiry"] = client._token_expiry
        log.append(_stamp(f"تمت معالجة {state['stats']['documents']} من {state['total']}"))
        if not queue:
            state["running"] = False
            state["finished_at"] = datetime.now().isoformat(timespec="seconds")
            log.append(_stamp("اكتملت المزامنة ✓"))
    except Exception as e:  # noqa: BLE001
        state["running"] = False
        state["error"] = str(e)
        state["finished_at"] = datetime.now().isoformat(timespec="seconds")
        log.append(_stamp(f"خطأ: {e}"))
    state["queue"] = queue
    state["log"] = log[-80:]
    _save_sync(state)
    return _public_sync(state)


@app.get("/api/eta/sync/status")
def sync_status():
    return _public_sync(_load_sync())


# --------------------------------------------------------------------------- static UI
_static = Path(__file__).resolve().parent.parent / "static"
if _static.is_dir():
    app.mount("/", StaticFiles(directory=str(_static), html=True), name="static")
