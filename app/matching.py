# -*- coding: utf-8 -*-
"""Arabic text normalization and purchase↔sale auto-matching engine."""
from __future__ import annotations

import re
from datetime import date, datetime
from difflib import SequenceMatcher
from itertools import combinations

MATCH_THRESHOLD = 55
_DIACRITICS = re.compile(r"[ً-ْٰـ]")  # tashkeel + dagger alef + tatweel
_ALEF = re.compile(r"[أإآٱ]")  # أ إ آ ٱ
_PUNCT = re.compile(r"[^\w\s؀-ۿ]", re.UNICODE)
_WS = re.compile(r"\s+")


def normalize_ar(text) -> str:
    if not text:
        return ""
    s = str(text)
    s = _DIACRITICS.sub("", s)
    s = _ALEF.sub("ا", s)          # → ا
    s = s.replace("ى", "ي")   # ى → ي
    s = s.replace("ة", "ه")   # ة → ه
    s = s.replace("ئ", "ي")   # ئ → ي
    s = s.replace("ؤ", "و")   # ؤ → و
    s = _PUNCT.sub(" ", s)
    s = _WS.sub(" ", s).strip().lower()
    return s


def similarity(a, b) -> float:
    na, nb = normalize_ar(a), normalize_ar(b)
    if not na or not nb:
        return 0.0
    if na == nb:
        return 1.0
    return SequenceMatcher(None, na, nb).ratio()


def _parse_date(value):
    if value in (None, ""):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    s = str(value).strip()[:10]
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def score_pair(purchase: dict, sale: dict) -> float:
    pq, sq = purchase.get("qty") or 0, sale.get("qty") or 0
    score = 0.0
    if pq and sq:
        rel = abs(abs(pq) - abs(sq)) / max(abs(pq), abs(sq))
        if rel < 1e-9:
            score += 50
        elif rel <= 0.02:
            score += 35
        elif rel <= 0.10:
            score += 15
        elif rel <= 0.25:
            score += 5
        else:
            return 0.0  # quantities too far apart — not the same goods
    score += 30 * similarity(purchase.get("item"), sale.get("item"))
    pd, sd = _parse_date(purchase.get("invoice_date")), _parse_date(sale.get("invoice_date"))
    if pd and sd:
        delta = abs((sd - pd).days)
        if delta <= 60:
            score += 15 * (1 - delta / 60)
    if (pq >= 0) == (sq >= 0):  # same polarity (regular vs credit note)
        score += 5
    return min(score, 100.0)


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
