# -*- coding: utf-8 -*-
"""FastAPI app: Arabic purchases↔sales matching tool with ETA integration."""
from __future__ import annotations

import tempfile
import threading
import urllib.parse
from datetime import date, datetime
from pathlib import Path

from fastapi import FastAPI, HTTPException, UploadFile, File, Form
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field

from . import db
from .eta_client import ETAClient, ETAError
from .excel_io import export_workbook, import_workbook
from .matching import auto_match

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
    fk = "purchase_line_id" if kind == "purchase" else "sale_line_id"
    sql = (f"SELECT t.*, m.id AS match_id FROM {table} t "
           f"LEFT JOIN matches m ON m.{fk} = t.id WHERE 1=1")
    params: list = []
    if matched is True:
        sql += " AND m.id IS NOT NULL"
    elif matched is False:
        sql += " AND m.id IS NULL"
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
    purchase_id: int
    sale_id: int
    note: str = ""


class AcceptBody(BaseModel):
    pairs: list[MatchCreate]


@app.get("/api/matches")
def list_matches():
    sql = """
    SELECT m.id, m.note,
           p.id AS p_id, p.invoice_date AS p_date, p.invoice_no AS p_no,
           p.party AS supplier, p.item AS p_item, p.qty AS p_qty,
           p.unit_price AS p_price, p.vat AS p_vat, p.source AS p_source,
           s.id AS s_id, s.invoice_date AS s_date, s.invoice_no AS s_no,
           s.party AS customer, s.item AS s_item, s.qty AS s_qty,
           s.unit_price AS s_price, s.vat AS s_vat, s.source AS s_source,
           s.note AS s_note
    FROM matches m
    JOIN purchase_lines p ON p.id = m.purchase_line_id
    JOIN sale_lines s ON s.id = m.sale_line_id
    ORDER BY p.invoice_date IS NULL, p.invoice_date, m.id
    """
    with db.get_conn() as conn:
        rows = conn.execute(sql).fetchall()
    out = []
    for r in rows:
        d = dict(r)
        d["qty_diff"] = (r["p_qty"] or 0) - (r["s_qty"] or 0)
        out.append(d)
    return out


@app.post("/api/matches")
def create_match(body: MatchCreate):
    with db.get_conn() as conn:
        p = conn.execute("SELECT id FROM purchase_lines WHERE id=?",
                         (body.purchase_id,)).fetchone()
        s = conn.execute("SELECT id FROM sale_lines WHERE id=?",
                         (body.sale_id,)).fetchone()
        if not p or not s:
            raise HTTPException(404, "سطر المشتريات أو المبيعات غير موجود")
        cur = conn.execute(
            "INSERT INTO matches(purchase_line_id, sale_line_id, note) VALUES (?,?,?) "
            "ON CONFLICT DO NOTHING",
            (body.purchase_id, body.sale_id, body.note))
        if cur.rowcount == 0:
            raise HTTPException(409, "أحد السطرين مرتبط بالفعل بمطابقة أخرى")
    return {"id": cur.lastrowid}


@app.delete("/api/matches/{match_id}")
def delete_match(match_id: int):
    with db.get_conn() as conn:
        cur = conn.execute("DELETE FROM matches WHERE id=?", (match_id,))
    if cur.rowcount == 0:
        raise HTTPException(404, "المطابقة غير موجودة")
    return {"deleted": match_id}


def _unmatched(conn, kind):
    table, fk = _TABLES[kind], "purchase_line_id" if kind == "purchase" else "sale_line_id"
    rows = conn.execute(
        f"SELECT t.* FROM {table} t LEFT JOIN matches m ON m.{fk}=t.id "
        f"WHERE m.id IS NULL").fetchall()
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
            "SELECT COUNT(*) c FROM purchase_lines t LEFT JOIN matches m "
            "ON m.purchase_line_id=t.id WHERE m.id IS NULL").fetchone()["c"]
        unmatched_s = conn.execute(
            "SELECT COUNT(*) c FROM sale_lines t LEFT JOIN matches m "
            "ON m.sale_line_id=t.id WHERE m.id IS NULL").fetchone()["c"]
        mismatch = conn.execute(
            "SELECT COUNT(*) c FROM matches m "
            "JOIN purchase_lines p ON p.id=m.purchase_line_id "
            "JOIN sale_lines s ON s.id=m.sale_line_id "
            "WHERE ABS(p.qty - s.qty) > 0.001").fetchone()["c"]
        matches_count = conn.execute("SELECT COUNT(*) c FROM matches").fetchone()["c"]
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


@app.get("/api/export/excel")
def export_excel():
    with db.get_conn() as conn:
        purchases = [dict(r) for r in conn.execute("SELECT * FROM purchase_lines")]
        sales = [dict(r) for r in conn.execute("SELECT * FROM sale_lines")]
        pairs = [(r["purchase_line_id"], r["sale_line_id"]) for r in conn.execute(
            "SELECT m.purchase_line_id, m.sale_line_id FROM matches m "
            "JOIN purchase_lines p ON p.id=m.purchase_line_id "
            "ORDER BY p.invoice_date IS NULL, p.invoice_date, m.id")]
    exports = Path(db.DB_PATH).parent / "exports"
    exports.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y-%m-%d_%H%M")
    out = exports / f"matching_{stamp}.xlsx"
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


@app.put("/api/settings")
def put_settings(body: SettingsBody):
    db.set_setting("eta_env", body.eta_env)
    db.set_setting("eta_client_id", body.eta_client_id)
    if body.eta_client_secret:
        db.set_setting("eta_client_secret", body.eta_client_secret)
    db.set_setting("vat_rate", str(body.vat_rate))
    return get_settings()


# --------------------------------------------------------------------------- ETA sync
def _eta_client() -> ETAClient:
    cid = db.get_setting("eta_client_id")
    secret = db.get_setting("eta_client_secret")
    if not cid or not secret:
        raise HTTPException(400, "أدخل بيانات الاتصال بمنظومة الفواتير أولاً من صفحة الإعدادات")
    return ETAClient(db.get_setting("eta_env", "preprod"), cid, secret)


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


_sync_state = {"running": False, "log": [], "error": None,
               "started_at": None, "finished_at": None,
               "stats": {"documents": 0, "lines": 0}}
_sync_lock = threading.Lock()


def _sync_log(msg: str):
    _sync_state["log"].append(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")


def _run_sync(body: SyncBody):
    try:
        client = ETAClient(db.get_setting("eta_env", "preprod"),
                           db.get_setting("eta_client_id"),
                           db.get_setting("eta_client_secret"))
        _sync_log("جارٍ تسجيل الدخول إلى منظومة الفواتير…")
        client.get_token()
        _sync_log("تم تسجيل الدخول بنجاح")
        with db.get_conn() as conn:
            known = {r["eta_uuid"] for r in conn.execute(
                "SELECT eta_uuid FROM purchase_lines WHERE eta_uuid IS NOT NULL "
                "UNION SELECT eta_uuid FROM sale_lines WHERE eta_uuid IS NOT NULL")}
        for direction in body.directions:
            label = "المشتريات (الواردة)" if direction == "Received" else "المبيعات (الصادرة)"
            _sync_log(f"جارٍ البحث عن مستندات {label} من {body.date_from} إلى {body.date_to}…")
            summaries = list(client.search_documents(
                body.date_from, body.date_to, direction=direction,
                on_progress=lambda a, b, n: _sync_log(f"  نافذة {a} → {b}: {n} مستند")))
            _sync_log(f"إجمالي مستندات {label}: {len(summaries)}")
            table = "purchase_lines" if direction == "Received" else "sale_lines"
            for i, summary in enumerate(summaries, 1):
                uuid = summary.get("uuid")
                if not body.refresh and uuid in known:
                    continue
                lines = client.get_document_lines(summary, direction)
                with db.get_conn() as conn:
                    for ln in lines:
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
                _sync_state["stats"]["documents"] += 1
                _sync_state["stats"]["lines"] += len(lines)
                if i % 10 == 0:
                    _sync_log(f"  تمت معالجة {i} من {len(summaries)}")
        _sync_log("اكتملت المزامنة ✓")
    except (ETAError, Exception) as e:  # noqa: BLE001 — surface everything to the UI log
        _sync_state["error"] = str(e)
        _sync_log(f"خطأ: {e}")
    finally:
        _sync_state["running"] = False
        _sync_state["finished_at"] = datetime.now().isoformat(timespec="seconds")


@app.post("/api/eta/sync")
def eta_sync(body: SyncBody):
    _eta_client()  # validates credentials exist
    with _sync_lock:
        if _sync_state["running"]:
            raise HTTPException(409, "هناك مزامنة قيد التنفيذ بالفعل")
        _sync_state.update(running=True, log=[], error=None,
                           started_at=datetime.now().isoformat(timespec="seconds"),
                           finished_at=None, stats={"documents": 0, "lines": 0})
    threading.Thread(target=_run_sync, args=(body,), daemon=True).start()
    return {"started": True}


@app.get("/api/eta/sync/status")
def sync_status():
    return _sync_state


# --------------------------------------------------------------------------- static UI
_static = Path(__file__).resolve().parent.parent / "static"
if _static.is_dir():
    app.mount("/", StaticFiles(directory=str(_static), html=True), name="static")
