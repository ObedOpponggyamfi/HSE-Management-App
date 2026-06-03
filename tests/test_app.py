"""Smoke + behaviour tests for the HSE Management App.

Uses a throwaway SQLite database (set before importing the app) and the dummy
data in data/ (CI runs generate_dummy_data.py first).
"""
import os
import tempfile

# Point the app at a temp DB and disable browser/reloader BEFORE importing it.
_TMPDB = os.path.join(tempfile.gettempdir(), "hse_ci_test.sqlite")
os.environ["HSE_DATABASE_URI"] = f"sqlite:///{_TMPDB}"
os.environ.setdefault("HSE_NO_BROWSER", "1")
os.environ.setdefault("HSE_NO_RELOAD", "1")
if os.path.exists(_TMPDB):
    os.remove(_TMPDB)

import pytest  # noqa: E402

import app as A  # noqa: E402

ALL = {"year": "All", "month": "All", "dept": "All", "area": "All"}


@pytest.fixture()
def client():
    A.app.config["WTF_CSRF_ENABLED"] = False
    A.app.config["TESTING"] = True
    return A.app.test_client()


def login(c, username="admin", password="admin123"):
    return c.post("/login", data={"username": username, "password": password})


def test_unauthenticated_redirects(client):
    assert client.get("/").status_code == 302


def test_all_pages_load_for_admin(client):
    login(client)
    for path in ["/", "/rates", "/incidents", "/events", "/actions", "/training", "/compliance",
                 "/environmental", "/contractors", "/registers", "/alerts",
                 "/report", "/data", "/admin/users", "/admin/audit"]:
        assert client.get(path).status_code == 200, path


def test_rolling_rates_shape():
    r = A.store.rolling_rates(ALL)
    assert set(("labels", "trifr", "ltifr", "aifr", "latest")).issubset(r)
    assert len(r["labels"]) == len(r["trifr"]) == len(r["aifr"])


def test_dashboard_has_twelve_kpis():
    assert len(A.store.kpis(ALL)) == 12


def test_capture_event_persists(client):
    login(client)
    before = A.store.df("events").shape[0]
    resp = client.post("/events/new", data={
        "category": "Near Miss", "date": "2026-05-20", "area": "Nkran Open Pit",
        "severity": "2", "description": "ci automated test report"})
    assert resp.status_code == 302
    assert A.store.df("events").shape[0] == before + 1


def test_role_gating_blocks_worker(client):
    login(client, "worker", "worker123")
    assert client.get("/incidents/new").status_code == 302   # blocked -> redirect
    assert client.get("/events/new").status_code == 200       # allowed


def test_import_audit_recorded():
    from models import ImportRun
    with A.app.app_context():
        assert ImportRun.query.count() >= 1


def test_trifr_is_non_negative_number():
    cards = {c["label"]: c["value"] for c in A.store.kpis(ALL)}
    assert isinstance(cards["TRIFR"], (int, float)) and cards["TRIFR"] >= 0
