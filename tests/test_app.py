from __future__ import annotations

import datetime as dt
import importlib
import os

import pytest

import config as C
from core import DataStore
from generate_dummy_data import generate
from hse_dashboard.database import make_engine
from hse_dashboard.importers import import_spreadsheets

os.environ.setdefault("HSE_SKIP_GLOBAL_APP", "1")
os.environ.setdefault("HSE_NO_BROWSER", "1")
os.environ.setdefault("HSE_NO_RELOAD", "1")


def _write_workbooks(folder, anchor=dt.date(2026, 5, 31)):
    data = generate(anchor=anchor)
    folder.mkdir(parents=True, exist_ok=True)
    for key, spec in C.DATASETS.items():
        data[key].to_excel(folder / spec["file"], sheet_name=spec["sheet"], index=False)
    return data


@pytest.fixture()
def client(tmp_path, monkeypatch):
    data_dir = tmp_path / "data"
    db_path = tmp_path / "hse.sqlite"
    _write_workbooks(data_dir)
    import_spreadsheets("site-a", data_dir, make_engine(str(db_path)), include_workbooks=False)
    store = DataStore(data_dir=str(data_dir), db_path=str(db_path), site_id="site-a", auto_import=False)

    monkeypatch.setenv("HSE_SKIP_GLOBAL_APP", "1")
    app_module = importlib.import_module("app")
    flask_app = app_module.create_app(store=store)
    flask_app.config["TESTING"] = True
    return flask_app.test_client()


def test_public_pages_load(client):
    for path in [
        "/",
        "/incidents",
        "/actions",
        "/compliance",
        "/environmental",
        "/contractors",
        "/registers",
        "/data",
    ]:
        response = client.get(path)
        assert response.status_code == 200, path


def test_filter_scope_is_page_specific(client):
    environmental = client.get("/environmental?year=2026&month=May&dept=Mining")
    assert environmental.status_code == 200
    assert b'name="year"' in environmental.data
    assert b'name="month"' in environmental.data
    assert b'name="dept"' not in environmental.data

    data_page = client.get("/data")
    assert data_page.status_code == 200
    assert b"filterbar" not in data_page.data


def test_refresh_blocks_external_redirects(client):
    response = client.post("/refresh", data={"next": "https://example.com/phish"})
    assert response.status_code == 302
    assert response.headers["Location"] == "/"


def test_tables_have_client_enhancement_hooks(client):
    response = client.get("/incidents?year=2026&month=May")
    assert response.status_code == 200
    assert b'class="data-table sortable"' in response.data
    assert b"tables.js" in response.data
