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
