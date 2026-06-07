from __future__ import annotations

import datetime as dt
import importlib
import os

import pandas as pd

import config as C
from core import DataStore
from generate_dummy_data import generate
from hse_dashboard.database import make_engine
from hse_dashboard.importers import import_spreadsheets, import_workbook


def write_standard_workbooks(folder, anchor=dt.date(2026, 5, 31)):
    data = generate(anchor=anchor)
    folder.mkdir(parents=True, exist_ok=True)
    for key, spec in C.DATASETS.items():
        data[key].to_excel(folder / spec["file"], sheet_name=spec["sheet"], index=False)
    return data


def test_standard_import_and_derivations(tmp_path):
    data_dir = tmp_path / "data"
    write_standard_workbooks(data_dir)
    db_path = tmp_path / "hse.sqlite"
    engine = make_engine(str(db_path))

    results = import_spreadsheets("site-a", data_dir, engine, include_workbooks=False)
    assert len(results) == len(C.DATASETS)
    assert all(r.rows_accepted > 0 for r in results)
    assert all(r.rows_rejected == 0 for r in results)

    store = DataStore(data_dir=str(data_dir), db_path=str(db_path), site_id="site-a", auto_import=False)
    all_filters = {"year": "All", "month": "All", "dept": "All", "area": "All"}
    may_filters = {"year": "2026", "month": "May", "dept": "All", "area": "All"}

    assert store.kpis(all_filters)[0]["value"] >= store.kpis(may_filters)[0]["value"]
    assert store.monthly_trend(may_filters)["labels"] == ["May-26"]
    assert store.actions_by_status(all_filters) != store.actions_by_status(may_filters)
    assert {"Recordable", "LTI", "CAR_Overdue"}.issubset(store.df("incidents").columns)


def test_rich_incident_workbook_adapter_rejects_duplicates_and_ignores_date_only_rows(tmp_path):
    workbook = tmp_path / "incidents.xlsx"
    df = pd.DataFrame([
        {
            "CASE ID": "IR-001",
            "DATE OF INCIDENT": dt.date(2026, 1, 3),
            "BUSINESS DEPARTMENT / UNIT": "Mining",
            "BUSINESS PARTNER": "AUMS (Mining Contractor)",
            "LOCATION": "Nkran Open Pit",
            "INCIDENT TYPE": "Lost Time Injury",
            "STATUS": "Closed",
            "RISK RATING (Actual Consequence)": "High",
        },
        {
            "CASE ID": "IR-001",
            "DATE OF INCIDENT": dt.date(2026, 1, 4),
            "LOCATION": "Nkran Open Pit",
            "INCIDENT TYPE": "First Aid",
        },
        {"DATE OF INCIDENT": dt.date(2026, 1, 5)},
        {"DATE OF INCIDENT": dt.date(2026, 1, 6), "INCIDENT TYPE": "Other"},
    ])
    with pd.ExcelWriter(workbook) as writer:
        df.to_excel(writer, sheet_name="DataBase", index=False, startrow=1)

    result = import_workbook("site-a", str(workbook), "asanko_incidents_v1",
                             engine=make_engine(str(tmp_path / "hse.sqlite")))
    assert result.rows_accepted == 1
    assert result.rows_rejected == 2
    assert any(e["message"] == "duplicate CASE ID in workbook" for e in result.errors)
    assert any(e["message"] == "ignored date-only rows" and e["rows"] == 1 for e in result.errors)


def test_flask_routes_refresh_redirect_and_filter_visibility(tmp_path, monkeypatch):
    data_dir = tmp_path / "data"
    write_standard_workbooks(data_dir)
    db_path = tmp_path / "hse.sqlite"
    engine = make_engine(str(db_path))
    import_spreadsheets("site-a", data_dir, engine, include_workbooks=False)
    store = DataStore(data_dir=str(data_dir), db_path=str(db_path), site_id="site-a", auto_import=False)

    monkeypatch.setenv("HSE_SKIP_GLOBAL_APP", "1")
    app_module = importlib.import_module("app")
    flask_app = app_module.create_app(store=store)
    client = flask_app.test_client()

    env_response = client.get("/environmental?dept=Mining")
    assert env_response.status_code == 200
    assert b'name="year"' in env_response.data
    assert b'name="dept"' not in env_response.data

    data_response = client.get("/data")
    assert data_response.status_code == 200
    assert b"Canonical store" in data_response.data
    assert b"filterbar" not in data_response.data

    refresh = client.post("/refresh", data={"next": "https://example.com"})
    assert refresh.status_code == 302
    assert refresh.headers["Location"] == "/"


def test_workbook_generator_uses_shared_config():
    import generate_hse_dashboard as generator

    settings = {name: value for name, _label, value, _fmt in generator.SETTINGS}
    assert generator.COMPANY_NAME == C.COMPANY_NAME
    assert generator.INCIDENT_TYPES == C.INCIDENT_TYPES
    assert settings["TARGET_TRIFR"] == C.TARGET_TRIFR
    assert settings["TARGET_TRAINING"] == C.TARGET_TRAINING
    assert settings["PM10_LIMIT"] == C.PM10_LIMIT
