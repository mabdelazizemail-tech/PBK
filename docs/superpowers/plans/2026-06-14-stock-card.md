# Stock Card (كارت صنف) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** A read-only "كارت صنف" report: for each item (grouped by normalized name), a date-ordered ledger of incoming (purchases) and outgoing (sales) movements with a running balance, viewable on screen and exportable to Excel.

**Architecture:** A pure read model in a new `app/stock.py` aggregates purchase/sale line dicts: it buckets lines by `normalize_ar(item)`, computes per-movement stock effect `e = +qty` (purchase) / `-qty` (sale), routes positive effects to «وارد» and negative to «منصرف», and accumulates a running balance — so credit notes (stored negative) land in the opposite column automatically. Three GET endpoints expose an overview, a single item card, and an Excel export. A new «كارت الصنف» tab renders the overview list with drill-down. This feature is fully independent of matching.

**Tech Stack:** Python 3 / FastAPI, openpyxl, vanilla JS frontend, pytest.

**Spec:** [docs/superpowers/specs/2026-06-14-group-matching-stock-card-design.md](../specs/2026-06-14-group-matching-stock-card-design.md)

**Conventions to respect:**
- Tests run on SQLite (root `conftest.py`). `tests/test_api.py` shares **one DB across the module** — use a unique item name per test and filter results to it.
- Reuse `normalize_ar` and `_parse_date` from `app/matching.py`; do not re-implement normalization.
- This plan can be implemented independently of (and in either order with) the group-matching plan.

---

### Task 1: Stock read model

**Files:**
- Create: `app/stock.py`
- Test: `tests/test_stock.py` (create)

- [ ] **Step 1: Write the failing tests**

```python
# tests/test_stock.py
# -*- coding: utf-8 -*-
"""Tests for the stock-card read model."""
import pytest
from app.stock import build_overview, build_item_card, _key


def p(id, item, qty, date="2026-01-10", no="P", party="مورد"):
    return {"id": id, "item": item, "qty": qty, "invoice_date": date,
            "invoice_no": no, "party": party, "source": "manual", "doc_type": "i", "note": ""}


def s(id, item, qty, date="2026-01-20", no="S", party="عميل", doc_type="i", note=""):
    return {"id": id, "item": item, "qty": qty, "invoice_date": date,
            "invoice_no": no, "party": party, "source": "manual",
            "doc_type": doc_type, "note": note}


class TestOverview:
    def test_groups_spelling_variants_but_keeps_distinct_models(self):
        purchases = [p(1, "نموذج 2", 100), p(2, "نموذج  2", 50), p(3, "نموذج 1", 70)]
        ov = build_overview(purchases, [])
        keys = {o["item_key"] for o in ov}
        assert _key("نموذج 2") in keys
        assert _key("نموذج 1") in keys
        n2 = [o for o in ov if o["item_key"] == _key("نموذج 2")][0]
        assert n2["total_in"] == pytest.approx(150)        # 100 + 50 merged
        assert len({o["item_key"] for o in ov}) == 2       # model 1 stays separate

    def test_balance_in_minus_out(self):
        purchases = [p(1, "صنف", 100), p(2, "صنف", 50)]
        sales = [s(10, "صنف", 30)]
        ov = build_overview(purchases, sales)[0]
        assert ov["total_in"] == pytest.approx(150)
        assert ov["total_out"] == pytest.approx(30)
        assert ov["balance"] == pytest.approx(120)
        assert ov["movements_count"] == 3

    def test_skips_blank_items(self):
        assert build_overview([p(1, "", 10)], []) == []


class TestItemCard:
    def test_running_balance_is_chronological(self):
        purchases = [p(1, "صنف", 100, "2026-01-05"), p(2, "صنف", 40, "2026-01-20")]
        sales = [s(10, "صنف", 30, "2026-01-10")]
        card = build_item_card(purchases, sales, _key("صنف"))
        balances = [r["balance"] for r in card["rows"]]
        assert balances == pytest.approx([100, 70, 110])   # +100, -30, +40
        assert card["balance"] == pytest.approx(110)

    def test_sale_credit_note_returns_to_stock(self):
        # customer return: sale qty stored negative → counts as وارد, raises balance
        card = build_item_card([p(1, "صنف", 100)],
                               [s(10, "صنف", -20, doc_type="c", note="اشعار دائن")],
                               _key("صنف"))
        credit = [r for r in card["rows"] if r["invoice_no"] == "S"][0]
        assert credit["kind"] == "in"
        assert credit["qty_in"] == pytest.approx(20)
        assert card["balance"] == pytest.approx(120)

    def test_purchase_credit_note_leaves_stock(self):
        # return to supplier: purchase qty negative → counts as منصرف, lowers balance
        card = build_item_card([p(1, "صنف", 100), p(2, "صنف", -30, no="PC")],
                               [], _key("صنف"))
        credit = [r for r in card["rows"] if r["invoice_no"] == "PC"][0]
        assert credit["kind"] == "out"
        assert credit["qty_out"] == pytest.approx(30)
        assert card["balance"] == pytest.approx(70)

    def test_label_is_most_common_spelling(self):
        purchases = [p(1, "نموذج 2", 1), p(2, "نموذج 2", 1), p(3, "نموذج  2", 1)]
        card = build_item_card(purchases, [], _key("نموذج 2"))
        assert card["item_label"] == "نموذج 2"
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_stock.py -v`
Expected: FAIL — `No module named 'app.stock'`.

- [ ] **Step 3: Implement `app/stock.py`**

```python
# -*- coding: utf-8 -*-
"""Stock-card read model: per-item in/out ledger with running balance.

Movement stock effect ``e`` is ``+qty`` for a purchase and ``-qty`` for a sale.
Positive effects are «وارد», negative are «منصرف», so credit notes (stored with a
negative qty) automatically land in the opposite column. Items are grouped by the
shared Arabic normalization so spelling variants merge while distinct models stay
apart.
"""
from __future__ import annotations

from collections import Counter
from datetime import date

from .matching import _parse_date, normalize_ar


def _key(item) -> str:
    return normalize_ar(item)


def _label_for(items) -> str:
    counts = Counter(i for i in items if i)
    if not counts:
        return ""
    return max(counts.items(), key=lambda kv: (kv[1], len(kv[0])))[0]


def _sort_key(line):
    d = _parse_date(line.get("invoice_date"))
    return (d is None, d or date.min, line.get("id") or 0)


def _move_label(kind: str, qty: float) -> str:
    if kind == "purchase":
        return "إشعار دائن مشتريات" if qty < 0 else "شراء"
    return "إشعار دائن مبيعات" if qty < 0 else "بيع"


def build_overview(purchases: list[dict], sales: list[dict]) -> list[dict]:
    buckets: dict[str, dict] = {}

    def add(line, kind):
        k = _key(line.get("item"))
        if not k:
            return
        b = buckets.setdefault(
            k, {"labels": [], "in": 0.0, "out": 0.0, "bal": 0.0, "n": 0})
        qty = line.get("qty") or 0
        e = qty if kind == "purchase" else -qty
        if e > 0:
            b["in"] += e
        else:
            b["out"] += -e
        b["bal"] += e
        b["n"] += 1
        b["labels"].append(line.get("item"))

    for ln in purchases:
        add(ln, "purchase")
    for ln in sales:
        add(ln, "sale")

    out = [{"item_key": k, "item_label": _label_for(b["labels"]),
            "total_in": b["in"], "total_out": b["out"],
            "balance": b["bal"], "movements_count": b["n"]}
           for k, b in buckets.items()]
    out.sort(key=lambda x: x["item_label"])
    return out


def build_item_card(purchases: list[dict], sales: list[dict], key: str) -> dict:
    moves = [("purchase", ln) for ln in purchases if _key(ln.get("item")) == key]
    moves += [("sale", ln) for ln in sales if _key(ln.get("item")) == key]
    moves.sort(key=lambda kv: _sort_key(kv[1]))

    rows, balance, total_in, total_out, labels = [], 0.0, 0.0, 0.0, []
    for kind, ln in moves:
        labels.append(ln.get("item"))
        qty = ln.get("qty") or 0
        e = qty if kind == "purchase" else -qty
        qin = e if e > 0 else 0
        qout = -e if e < 0 else 0
        balance += e
        total_in += qin
        total_out += qout
        rows.append({
            "date": ln.get("invoice_date"), "kind": "in" if e > 0 else "out",
            "label": _move_label(kind, qty), "invoice_no": ln.get("invoice_no") or "",
            "party": ln.get("party") or "", "qty_in": qin, "qty_out": qout,
            "balance": balance, "source": ln.get("source") or "",
            "doc_type": ln.get("doc_type") or "i", "note": ln.get("note") or "",
        })
    return {"item_key": key, "item_label": _label_for(labels), "rows": rows,
            "total_in": total_in, "total_out": total_out, "balance": balance}
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_stock.py -v`
Expected: PASS (all).

- [ ] **Step 5: Commit**

```bash
git add app/stock.py tests/test_stock.py
git commit -m "feat(stock): per-item in/out ledger read model with signed balance"
```

---

### Task 2: Stock-card API endpoints

**Files:**
- Modify: `app/main.py` (import + two endpoints)
- Test: `tests/test_api.py` (add stock endpoint tests)

- [ ] **Step 1: Write the failing tests**

```python
# add a new class to tests/test_api.py
class TestStockCard:
    def test_overview_and_item(self, client):
        _purchase(client, qty=500, item="مخزون صنف أ", date="2026-01-05")
        _purchase(client, qty=300, item="مخزون صنف أ", date="2026-01-09")
        _sale(client, qty=200, item="مخزون صنف أ", date="2026-01-12")
        ov = client.get("/api/stock-card").json()
        mine = [o for o in ov if o["item_label"] == "مخزون صنف أ"]
        assert mine and mine[0]["total_in"] == pytest.approx(800)
        assert mine[0]["total_out"] == pytest.approx(200)
        assert mine[0]["balance"] == pytest.approx(600)
        key = mine[0]["item_key"]
        card = client.get("/api/stock-card/item", params={"key": key}).json()
        assert card["balance"] == pytest.approx(600)
        assert [r["balance"] for r in card["rows"]] == pytest.approx([500, 800, 600])
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_api.py -k Stock -v`
Expected: FAIL — 404, endpoints not defined.

- [ ] **Step 3: Add the endpoints**

In `app/main.py`, add the import near the other app imports:

```python
from .stock import build_item_card, build_overview
```

And add the endpoints (e.g. just before the `# ---- excel` section):

```python
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
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_api.py -k Stock -v`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add app/main.py tests/test_api.py
git commit -m "feat(api): stock-card overview and item endpoints"
```

---

### Task 3: Excel export for the stock card

**Files:**
- Modify: `app/excel_io.py` (add `import re`, `_safe_sheet_name`, `export_stock_card`)
- Test: `tests/test_excel_io.py` (add stock export test)

- [ ] **Step 1: Write the failing test**

```python
# add to tests/test_excel_io.py
from app.excel_io import export_stock_card  # add to the existing import line


class TestStockExport:
    def test_single_card_sheet(self, tmp_path):
        card = {
            "item_key": "صنف", "item_label": "نموذج 2",
            "rows": [
                {"date": "2026-01-05", "kind": "in", "label": "شراء",
                 "invoice_no": "P1", "party": "مورد", "qty_in": 100, "qty_out": 0,
                 "balance": 100, "source": "manual", "doc_type": "i", "note": ""},
                {"date": "2026-01-12", "kind": "out", "label": "بيع",
                 "invoice_no": "S1", "party": "عميل", "qty_in": 0, "qty_out": 30,
                 "balance": 70, "source": "manual", "doc_type": "i", "note": ""},
            ],
            "total_in": 100, "total_out": 30, "balance": 70,
        }
        path = str(tmp_path / "card.xlsx")
        export_stock_card(None, [card], path)
        wb = load_workbook(path)
        ws = wb.active
        assert ws.sheet_view.rightToLeft is True
        # header row 3, first movement row 4
        assert [ws.cell(row=3, column=c).value for c in range(1, 8)] == [
            "التاريخ", "الحركة", "رقم الفاتورة", "الطرف", "وارد", "منصرف", "الرصيد"]
        assert ws.cell(row=4, column=5).value == 100      # وارد
        assert ws.cell(row=5, column=6).value == 30       # منصرف
        assert ws.cell(row=6, column=7).value == 70       # totals balance

    def test_overview_plus_per_item_sheets(self, tmp_path):
        overview = [{"item_key": "k", "item_label": "نموذج 2",
                     "total_in": 100, "total_out": 30, "balance": 70,
                     "movements_count": 2}]
        card = {"item_key": "k", "item_label": "نموذج 2", "rows": [],
                "total_in": 100, "total_out": 30, "balance": 70}
        path = str(tmp_path / "all.xlsx")
        export_stock_card(overview, [card], path)
        wb = load_workbook(path)
        assert "نظرة عامة" in wb.sheetnames
        assert any(name != "نظرة عامة" for name in wb.sheetnames)
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_excel_io.py -k Stock -v`
Expected: FAIL — `cannot import name 'export_stock_card'`.

- [ ] **Step 3: Implement the export**

In `app/excel_io.py`, add `import re` at the top with the other imports, then append:

```python
def _safe_sheet_name(name, used: set) -> str:
    s = re.sub(r"[\[\]:*?/\\]", " ", str(name or "صنف")).strip()[:28] or "صنف"
    base, i, candidate = s, 1, s
    while candidate in used:
        i += 1
        candidate = f"{base[:25]} {i}"
    used.add(candidate)
    return candidate


def export_stock_card(overview, cards: list[dict], path: str):
    """Write a stock-card workbook. If `overview` is given, the first sheet is a
    summary of all items; then one sheet per card."""
    wb = Workbook()
    used: set = set()
    header_fill = PatternFill("solid", start_color="D9E1F2")
    thin = Side(style="thin", color="9CA3AF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def header(ws, titles, row=1):
        for c, t in enumerate(titles, start=1):
            cell = ws.cell(row=row, column=c, value=t)
            cell.font = Font(bold=True)
            cell.fill = header_fill
            cell.border = border
            cell.alignment = Alignment(horizontal="center", vertical="center")

    first = True
    if overview is not None:
        ws = wb.active
        ws.title = _safe_sheet_name("نظرة عامة", used)
        ws.sheet_view.rightToLeft = True
        header(ws, ["الصنف", "وارد", "منصرف", "الرصيد"])
        for i, o in enumerate(overview, start=2):
            ws.cell(row=i, column=1, value=o["item_label"])
            for col, k in ((2, "total_in"), (3, "total_out"), (4, "balance")):
                ws.cell(row=i, column=col, value=o[k]).number_format = NUM_FMT
        ws.column_dimensions["A"].width = 28
        for col in ("B", "C", "D"):
            ws.column_dimensions[col].width = 14
        first = False

    for card in cards:
        ws = wb.active if first else wb.create_sheet()
        ws.title = _safe_sheet_name(card["item_label"], used)
        first = False
        ws.sheet_view.rightToLeft = True
        ws.cell(row=1, column=1,
                value=f"كارت صنف: {card['item_label']}").font = Font(bold=True, size=13)
        header(ws, ["التاريخ", "الحركة", "رقم الفاتورة", "الطرف", "وارد", "منصرف", "الرصيد"],
               row=3)
        r = 4
        for row in card["rows"]:
            ws.cell(row=r, column=1, value=row["date"] or "")
            ws.cell(row=r, column=2, value=row["label"])
            ws.cell(row=r, column=3, value=str(row["invoice_no"]))
            ws.cell(row=r, column=4, value=row["party"])
            ws.cell(row=r, column=5, value=row["qty_in"] or None).number_format = NUM_FMT
            ws.cell(row=r, column=6, value=row["qty_out"] or None).number_format = NUM_FMT
            ws.cell(row=r, column=7, value=row["balance"]).number_format = NUM_FMT
            r += 1
        ws.cell(row=r, column=4, value="الإجمالي").font = Font(bold=True)
        for col, k in ((5, "total_in"), (6, "total_out"), (7, "balance")):
            ws.cell(row=r, column=col, value=card[k]).number_format = NUM_FMT
        for col, w in {"A": 12, "B": 18, "C": 13, "D": 22,
                       "E": 11, "F": 11, "G": 13}.items():
            ws.column_dimensions[col].width = w

    wb.save(path)
    return path
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_excel_io.py -k Stock -v`
Expected: PASS (2 passed).

- [ ] **Step 5: Commit**

```bash
git add app/excel_io.py tests/test_excel_io.py
git commit -m "feat(excel): stock-card workbook export (overview + per-item sheets)"
```

---

### Task 4: Stock-card export endpoint

**Files:**
- Modify: `app/main.py` (import `export_stock_card`; add export endpoint)
- Test: `tests/test_api.py` (add export test)

- [ ] **Step 1: Write the failing test**

```python
# add to TestStockCard in tests/test_api.py
    def test_stock_card_export_returns_workbook(self, client):
        _purchase(client, qty=10, item="تصدير مخزون")
        _sale(client, qty=4, item="تصدير مخزون")
        r = client.get("/api/stock-card/export")
        assert r.status_code == 200
        assert r.content[:2] == b"PK"
        # single-item export by key
        ov = client.get("/api/stock-card").json()
        key = [o for o in ov if o["item_label"] == "تصدير مخزون"][0]["item_key"]
        r2 = client.get("/api/stock-card/export", params={"key": key})
        assert r2.status_code == 200 and r2.content[:2] == b"PK"
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_api.py -k export_returns_workbook -v`
Expected: FAIL — 404.

- [ ] **Step 3: Add the endpoint**

In `app/main.py`, extend the excel import to include `export_stock_card`:

```python
from .excel_io import export_stock_card, export_workbook, import_workbook
```

Add after `stock_card_item`:

```python
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
    exports = Path(db.DB_PATH).parent / "exports"
    exports.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y-%m-%d_%H%M")
    out = exports / f"stock_card_{stamp}.xlsx"
    export_stock_card(summary, cards, str(out))
    arabic_name = f"كارت صنف {stamp}.xlsx"
    quoted = urllib.parse.quote(arabic_name)
    return FileResponse(
        str(out),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition":
                 f"attachment; filename=stock_card_{stamp}.xlsx; filename*=UTF-8''{quoted}"})
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_api.py -k Stock -v`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add app/main.py tests/test_api.py
git commit -m "feat(api): stock-card Excel export endpoint"
```

---

### Task 5: Stock-card tab markup + nav

**Files:**
- Modify: `static/index.html` (rail button + new `#tab-stockcard` section)

- [ ] **Step 1: Add the rail nav button**

In `static/index.html`, add this button inside `.rail-nav` right after the «المبيعات» (`data-tab="sales"`) button:

```html
    <button class="rail-btn" data-tab="stockcard">
      <svg viewBox="0 0 24 24"><path d="M20 2H8c-1.1 0-2 .9-2 2v12c0 1.1.9 2 2 2h12c1.1 0 2-.9 2-2V4c0-1.1-.9-2-2-2zm0 14H8V4h12v12zM4 6H2v14c0 1.1.9 2 2 2h14v-2H4V6zm6 3h8v2h-8zm0 3h5v2h-5zm0-6h8v2h-8z"/></svg>
      <span>كارت الصنف</span>
    </button>
```

- [ ] **Step 2: Add the tab section**

Add this section after the `#tab-sales` section closes (`</section>`) and before `#tab-eta`:

```html
  <!-- ـــــــــــــــــــــ كارت الصنف ـــــــــــــــــــــ -->
  <section class="tab" id="tab-stockcard" hidden>
    <div class="stockcard-grid">
      <div class="card">
        <div class="card-head">
          <h3>الأصناف <em class="count" id="cnt-stock-items"></em></h3>
          <div class="row-actions">
            <input type="search" id="q-stock" class="input" placeholder="بحث عن صنف…">
            <button class="btn ghost" id="btn-stock-export-all">تصدير كل الأصناف</button>
          </div>
        </div>
        <div class="table-wrap tall"><table class="tbl selectable" id="tbl-stock-items">
          <thead><tr><th>الصنف</th><th>وارد</th><th>منصرف</th><th>الرصيد</th></tr></thead>
          <tbody></tbody>
        </table></div>
      </div>
      <div class="card">
        <div class="card-head">
          <h3 id="stock-card-title">اختر صنفاً لعرض كارت الصنف</h3>
          <button class="btn primary" id="btn-stock-export-one" hidden>تصدير Excel</button>
        </div>
        <div class="table-wrap tall"><table class="tbl" id="tbl-stock-card">
          <thead><tr><th>التاريخ</th><th>الحركة</th><th>رقم الفاتورة</th><th>الطرف</th><th>وارد</th><th>منصرف</th><th>الرصيد</th></tr></thead>
          <tbody></tbody>
        </table></div>
        <div class="stock-totals" id="stock-totals" hidden></div>
      </div>
    </div>
  </section>
```

- [ ] **Step 3: Verify markup**

Run: `python -c "import pathlib; h=pathlib.Path('static/index.html').read_text(encoding='utf-8'); assert 'tab-stockcard' in h and 'tbl-stock-items' in h; print('markup OK')"`
Expected: `markup OK`

- [ ] **Step 4: Commit**

```bash
git add static/index.html
git commit -m "feat(ui): stock-card tab markup and nav button"
```

---

### Task 6: Stock-card tab behaviour

**Files:**
- Modify: `static/app.js` (`TITLES`, `showTab` dispatch, new stock-card section)

- [ ] **Step 1: Register the tab**

In `static/app.js`, add to the `TITLES` object:

```javascript
  stockcard: "كارت الصنف",
```

And in `showTab`, add the dispatch line alongside the others:

```javascript
  if (name === "stockcard") loadStockCard();
```

- [ ] **Step 2: Add the stock-card section**

Append to `static/app.js` (before the final `loadDashboard();` bootstrap line, or after the import/export section):

```javascript
/* ـــــــــــــــــــــ كارت الصنف ـــــــــــــــــــــ */
let stockItems = [];
let stockSelKey = null;

async function loadStockCard() {
  try {
    stockItems = await api("/api/stock-card");
    renderStockItems();
  } catch (e) { toast(e.message, true); }
}

function renderStockItems() {
  const q = ($("#q-stock").value || "").trim();
  const rows = q ? stockItems.filter((it) => (it.item_label || "").includes(q)) : stockItems;
  $("#cnt-stock-items").textContent = `(${rows.length})`;
  $("#tbl-stock-items tbody").innerHTML = rows.map((it) => `
    <tr data-stock="${esc(it.item_key)}" class="${it.item_key === stockSelKey ? "selected" : ""}">
      <td>${esc(it.item_label)}</td>
      <td class="num">${qty(it.total_in)}</td>
      <td class="num">${qty(it.total_out)}</td>
      <td class="num"><b>${qty(it.balance)}</b></td>
    </tr>`).join("") ||
    `<tr><td colspan="4" class="empty">لا توجد أصناف — استورد بيانات أو زامن من ETA</td></tr>`;
}

$("#q-stock").addEventListener("input", () => renderStockItems());

document.addEventListener("click", (e) => {
  const tr = e.target.closest("tr[data-stock]");
  if (!tr) return;
  stockSelKey = tr.dataset.stock;
  renderStockItems();
  loadStockCardDetail(stockSelKey);
});

async function loadStockCardDetail(key) {
  try {
    const card = await api(`/api/stock-card/item?key=${encodeURIComponent(key)}`);
    $("#stock-card-title").textContent = `كارت صنف: ${card.item_label}`;
    $("#btn-stock-export-one").hidden = false;
    $("#tbl-stock-card tbody").innerHTML = card.rows.map((r) => `
      <tr>
        <td class="num">${esc(r.date || "—")}</td>
        <td>${esc(r.label)} ${r.doc_type === "c" ? '<span class="pill credit">دائن</span>' : ""}</td>
        <td>${esc(r.invoice_no)}</td>
        <td>${esc(r.party)}</td>
        <td class="num">${r.qty_in ? qty(r.qty_in) : ""}</td>
        <td class="num">${r.qty_out ? qty(r.qty_out) : ""}</td>
        <td class="num"><b>${qty(r.balance)}</b></td>
      </tr>`).join("") || `<tr><td colspan="7" class="empty">لا توجد حركات لهذا الصنف</td></tr>`;
    const tot = $("#stock-totals");
    tot.hidden = false;
    tot.innerHTML = `إجمالي الوارد: <b>${qty(card.total_in)}</b> · ` +
      `إجمالي المنصرف: <b>${qty(card.total_out)}</b> · ` +
      `الرصيد الحالي: <b>${qty(card.balance)}</b>`;
  } catch (e) { toast(e.message, true); }
}

$("#btn-stock-export-one").addEventListener("click", () => {
  if (!stockSelKey) return;
  window.location.href = `/api/stock-card/export?key=${encodeURIComponent(stockSelKey)}`;
  toast("جارٍ تجهيز كارت الصنف…");
});
$("#btn-stock-export-all").addEventListener("click", () => {
  window.location.href = "/api/stock-card/export";
  toast("جارٍ تجهيز ملف الأصناف…");
});
```

- [ ] **Step 3: Sanity-check the JS parses**

Run: `node --check static/app.js`
Expected: exit 0 (no output). If `node` is unavailable, the UI check in Task 8 covers it.

- [ ] **Step 4: Commit**

```bash
git add static/app.js
git commit -m "feat(ui): stock-card tab — overview list, drill-down ledger, export"
```

---

### Task 7: Stock-card styles

**Files:**
- Modify: `static/style.css` (append)

- [ ] **Step 1: Append styles**

Add to the end of `static/style.css`:

```css
/* ـــــ كارت الصنف ـــــ */
.stockcard-grid {
  display: grid; grid-template-columns: 380px 1fr; gap: 18px; align-items: start;
}
.stock-totals {
  margin-top: 12px; padding: 10px 14px; border-radius: 10px;
  background: #f9fafb; color: #374151; font-size: 14px;
}
.stock-totals b { color: #111827; font-variant-numeric: tabular-nums; }
@media (max-width: 1100px) { .stockcard-grid { grid-template-columns: 1fr; } }
```

- [ ] **Step 2: Verify**

Run: `python -c "import pathlib; c=pathlib.Path('static/style.css').read_text(encoding='utf-8'); assert '.stockcard-grid' in c; print('css OK')"`
Expected: `css OK`

- [ ] **Step 3: Commit**

```bash
git add static/style.css
git commit -m "style(ui): stock-card two-pane layout"
```

---

### Task 8: Verification

**Files:** none

- [ ] **Step 1: Run the pytest suite**

Run: `python -m pytest -q`
Expected: all pass, including `tests/test_stock.py` and the new `TestStockCard` / `TestStockExport`.

- [ ] **Step 2: Launch with real data**

Run (PowerShell, repo root):
```
$env:PBK_DB="data/pbk.db"; python -m uvicorn app.main:app --port 8077
```
In another shell, ensure data is loaded:
```
curl.exe -s -F "file=@مطابقة المشتريات والمبيعات.xlsx" -F "mode=replace" http://127.0.0.1:8077/api/import/excel
```

- [ ] **Step 3: Confirm the existing UI smoke still passes**

Run: `python tests/ui_smoke.py`
Expected: `FAILURES: none` — the new tab must not break existing tabs.

- [ ] **Step 4: Manual stock-card check (browser)**

Open `http://127.0.0.1:8077`, click «كارت الصنف». Confirm: the items list shows balances (وارد − منصرف); clicking «نموذج 2» loads a date-ordered ledger whose final الرصيد equals وارد − منصرف; a credit-note row appears in the opposite column (e.g. an اشعار دائن sale shows under وارد); «تصدير Excel» downloads a file that opens in Excel with a RTL «كارت صنف: نموذج 2» sheet.

- [ ] **Step 5: Sanity-check API balance vs overview**

Run: `python -c "import requests; ov=requests.get('http://127.0.0.1:8077/api/stock-card').json(); print('items:', len(ov)); print('sample:', ov[0] if ov else 'none')"`
Expected: a non-empty list; each item's `balance == total_in - total_out`.

- [ ] **Step 6: Commit any fixes**

```bash
git add -A
git commit -m "test: stock-card suite + UI smoke green"
```

---

## Self-Review

**Spec coverage:** normalized-name grouping (Task 1 `_key`/overview) ✓; signed in/out ledger + running balance with credit-note flips (Task 1, tests cover both purchase & sale credit notes) ✓; overview + item endpoints (Task 2) ✓; Excel export per-item and all-items (Tasks 3-4) ✓; new tab with list + drill-down + export buttons (Tasks 5-7) ✓; representative label = most common spelling (Task 1, tested) ✓; verification incl. existing UI smoke unaffected (Task 8) ✓.

**Type consistency:** `build_overview` items expose `item_key`/`item_label`/`total_in`/`total_out`/`balance`/`movements_count` (stock.py → /api/stock-card → app.js renderStockItems → excel overview sheet). `build_item_card` returns `item_label`/`rows`/`total_in`/`total_out`/`balance`, each row `date`/`kind`/`label`/`invoice_no`/`party`/`qty_in`/`qty_out`/`balance`/`source`/`doc_type`/`note` (stock.py → /api/stock-card/item → app.js loadStockCardDetail → excel card sheet). `export_stock_card(overview_or_None, cards, path)` signature matches both call sites (endpoint + tests).

**Placeholder scan:** none — every code step is complete; every run step has an exact command and expected result.
