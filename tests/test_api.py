# -*- coding: utf-8 -*-
"""API integration tests over a temporary database."""
import os
import tempfile

import pytest

_tmpdb = os.path.join(tempfile.mkdtemp(), "test_pbk.db")
os.environ["PBK_DB"] = _tmpdb

from fastapi.testclient import TestClient  # noqa: E402
from app.main import app  # noqa: E402


@pytest.fixture()
def client():
    return TestClient(app)


def _purchase(client, qty=24000, price=9.0, item="نموذج 2", date="2026-01-04"):
    r = client.post("/api/lines", json={
        "kind": "purchase", "invoice_date": date, "invoice_no": "14426",
        "party": "سيفتى باك", "item": item, "qty": qty, "unit_price": price,
    })
    assert r.status_code == 200, r.text
    return r.json()


def _sale(client, qty=24000, price=10.5, item="نموذج 2", date="2026-01-18", vat=0.0):
    r = client.post("/api/lines", json={
        "kind": "sale", "invoice_date": date, "invoice_no": "19-2026",
        "party": "الفشاوي", "item": item, "qty": qty, "unit_price": price, "vat": vat,
    })
    assert r.status_code == 200, r.text
    return r.json()


class TestLinesAndDashboard:
    def test_create_line_defaults_vat_to_14_percent(self, client):
        p = _purchase(client)
        assert p["vat"] == pytest.approx(24000 * 9 * 0.14)

    def test_explicit_vat_respected(self, client):
        s = _sale(client, vat=0.0)
        assert s["vat"] == 0.0

    def test_dashboard_reflects_lines(self, client):
        r = client.get("/api/dashboard")
        assert r.status_code == 200
        d = r.json()
        assert d["purchases"]["count"] >= 1
        assert d["sales"]["count"] >= 1
        assert d["vat_position"] == pytest.approx(
            d["sales"]["vat"] - d["purchases"]["vat"])

    def test_update_and_delete_line(self, client):
        p = _purchase(client, qty=100, price=10)
        pid = p["id"]
        r = client.put(f"/api/lines/purchase/{pid}", json={"qty": 150})
        assert r.status_code == 200
        assert r.json()["qty"] == 150
        r = client.delete(f"/api/lines/purchase/{pid}")
        assert r.status_code == 200
        r = client.get("/api/lines", params={"kind": "purchase"})
        assert all(l["id"] != pid for l in r.json())


class TestMatching:
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


class TestSettingsAndExport:
    def test_settings_roundtrip_masks_secret(self, client):
        r = client.put("/api/settings", json={
            "eta_env": "preprod", "eta_client_id": "cid", "eta_client_secret": "s3cret",
            "vat_rate": 0.14})
        assert r.status_code == 200
        got = client.get("/api/settings").json()
        assert got["eta_client_id"] == "cid"
        assert "s3cret" not in str(got)
        assert got["has_secret"] is True

    def test_credentials_trimmed_on_save(self, client):
        from app import db
        client.put("/api/settings", json={
            "eta_env": "prod", "eta_client_id": "  cid-123  ",
            "eta_client_secret": "  sec-456\n", "vat_rate": 0.14})
        assert db.get_setting("eta_client_id") == "cid-123"
        assert db.get_setting("eta_client_secret") == "sec-456"

    def test_export_returns_workbook(self, client):
        r = client.get("/api/export/excel")
        assert r.status_code == 200
        assert r.headers["content-type"].startswith(
            "application/vnd.openxmlformats-officedocument.spreadsheetml")
        assert len(r.content) > 1000

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


class TestSyncStateMachine:
    """The resumable, DB-backed ETA sync — exercised without any ETA network."""

    def test_status_default_shape(self, client):
        from app import main
        main._save_sync(main._default_sync())
        st = client.get("/api/eta/sync/status").json()
        assert st["running"] is False
        assert st["stats"] == {"documents": 0, "lines": 0}
        # internal fields are never exposed to the browser
        assert not ({"queue", "token", "last_request_at"} & set(st))

    def test_step_is_noop_when_idle(self, client):
        from app import main
        main._save_sync(main._default_sync())
        st = client.post("/api/eta/sync/step").json()  # returns at once, no ETA call
        assert st["running"] is False

    def test_public_state_never_leaks_token_or_queue(self, client):
        from app import main
        state = main._default_sync()
        state.update(queue=[{"direction": "Received", "summary": {"uuid": "x"}}],
                     token="SECRET-BEARER", last_request_at=123.0)
        main._save_sync(state)
        st = client.get("/api/eta/sync/status").json()
        assert "SECRET-BEARER" not in str(st)
        assert "queue" not in st and "token" not in st

    def test_sync_conflicts_when_already_running(self, client):
        from app import db, main
        db.set_setting("eta_client_id", "cid")
        db.set_setting("eta_client_secret", "secret")
        running = main._default_sync()
        running["running"] = True
        main._save_sync(running)
        r = client.post("/api/eta/sync", json={
            "date_from": "2026-01-01", "date_to": "2026-01-10",
            "directions": ["Received"]})
        assert r.status_code == 409
        main._save_sync(main._default_sync())  # reset for any later test

    def test_sync_requires_credentials(self, client):
        from app import db, main
        db.set_setting("eta_client_id", "")
        db.set_setting("eta_client_secret", "")
        main._save_sync(main._default_sync())
        r = client.post("/api/eta/sync", json={
            "date_from": "2026-01-01", "date_to": "2026-01-10",
            "directions": ["Received"]})
        assert r.status_code == 400


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
