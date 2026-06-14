# -*- coding: utf-8 -*-
"""Tests for the Arabic normalization and auto-matching engine."""
import pytest
from app.matching import normalize_ar, similarity, score_pair, auto_match


def line(id, date, party, item, qty, price, vat=0.0, doc_type="i"):
    return {
        "id": id, "invoice_date": date, "invoice_no": str(id), "party": party,
        "item": item, "qty": qty, "unit_price": price, "vat": vat, "doc_type": doc_type,
    }


class TestNormalizeAr:
    def test_unifies_ya_and_alef_maqsura(self):
        assert normalize_ar("الفيشاوى") == normalize_ar("الفيشاوي")

    def test_unifies_hamza_forms(self):
        assert normalize_ar("أحمد") == normalize_ar("احمد")
        assert normalize_ar("الإجمالي") == normalize_ar("الاجمالي")

    def test_unifies_ta_marbuta(self):
        assert normalize_ar("شركة") == normalize_ar("شركه")

    def test_collapses_whitespace_and_strips(self):
        assert normalize_ar("سيفتى  باك ") == normalize_ar("سيفتي باك")

    def test_removes_tatweel_and_diacritics(self):
        assert normalize_ar("مـــحـمَّد") == normalize_ar("محمد")

    def test_handles_none_and_empty(self):
        assert normalize_ar(None) == ""
        assert normalize_ar("") == ""


class TestSimilarity:
    def test_identical_after_normalization_is_1(self):
        assert similarity("نموذج 2", "نموذج 2") == 1.0

    def test_close_spellings_score_high(self):
        assert similarity("الفشاوي", "الفيشاوى") > 0.8

    def test_unrelated_score_low(self):
        assert similarity("نموذج 2", "ورق فلوت فاخر") < 0.5


class TestScorePair:
    def test_exact_qty_same_item_near_dates_scores_high(self):
        p = line(1, "2026-01-04", "سيفتى باك", "نموذج 2", 24000, 9)
        s = line(2, "2026-01-18", "الفشاوي", "نموذج 2", 24000, 10.5)
        assert score_pair(p, s) >= 85

    def test_qty_far_off_disqualifies(self):
        p = line(1, "2026-01-04", "سيفتى باك", "نموذج 2", 24000, 9)
        s = line(2, "2026-01-18", "الفشاوي", "نموذج 2", 10000, 10.5)
        assert score_pair(p, s) == 0

    def test_small_qty_diff_still_matches(self):
        # real case from the sheet: 2058 purchased vs 2085 sold
        p = line(1, "2026-01-26", "سيفتى باك", "نموذج 2", 2058, 9)
        s = line(2, "2026-01-26", "براميدز", "نموذج 2", 2085, 10.5)
        assert 0 < score_pair(p, s) < 100

    def test_sale_before_purchase_not_disqualified(self):
        # real case: purchase 2026-06-05 was sold 2026-05-20
        p = line(1, "2026-06-05", "سيفتى باك", "نموذج 2", 28634, 9.95)
        s = line(2, "2026-05-20", "الفيشاوي", "نموذج 2", 28634, 12)
        assert score_pair(p, s) >= 85


class TestAutoMatch:
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

    def test_each_line_used_at_most_once(self):
        purchases = [
            line(1, "2026-01-04", "سيفتى باك", "نموذج 2", 24000, 9),
            line(2, "2026-01-05", "سيفتى باك", "نموذج 2", 24000, 9),
        ]
        sales = [line(10, "2026-01-18", "الفشاوي", "نموذج 2", 24000, 10.5)]
        result = auto_match(purchases, sales)
        assert len(result) == 1

    def test_no_suggestion_below_threshold(self):
        purchases = [line(1, "2026-01-04", "سيفتى باك", "عسل", 3600, 17.5)]
        sales = [line(10, "2026-01-18", "الفشاوي", "كرتون شورتينج", 99999, 16)]
        assert auto_match(purchases, sales) == []

    def test_suggestion_includes_score_and_qty_diff(self):
        purchases = [line(1, "2026-01-26", "سيفتى باك", "نموذج 2", 2058, 9)]
        sales = [line(10, "2026-01-26", "براميدز", "نموذج 2", 2085, 10.5)]
        result = auto_match(purchases, sales)
        assert len(result) == 1
        m = result[0]
        assert m["qty_diff"] == pytest.approx(-27)
        assert 0 < m["score"] <= 100

    def test_empty_inputs(self):
        assert auto_match([], []) == []


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
