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
