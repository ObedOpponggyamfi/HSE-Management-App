#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Database bootstrap & write helpers.

  - import_excel_to_db()  : load every data/*.xlsx into a plain DB table
  - seed_database()       : first-run setup (import data, create users, sample events)
  - insert_incident/insert_action/update_action_status : used by the capture forms
  - next_ref()            : generate the next INC-/CAR-/EVT- reference

Bulk datasets are stored as plain tables (pandas <-> SQL); Users/Events/AuditLog
are ORM-managed (models.py).
"""
import datetime as dt
import os

import numpy as np
import pandas as pd
from sqlalchemy import inspect, text

import config as C
from extensions import db
from models import AuditLog, Event, User


# ---------------------------------------------------------------------------
# Excel  ->  DB
# ---------------------------------------------------------------------------
def import_excel_to_db():
    """Load each configured Excel file into a same-named DB table (replace)."""
    counts = {}
    for key, spec in C.DATASETS.items():
        path = os.path.join(C.DATA_DIR, spec["file"])
        if not os.path.exists(path):
            continue
        df = pd.read_excel(path, sheet_name=spec["sheet"])
        df.to_sql(key, db.engine, if_exists="replace", index=False)
        counts[key] = len(df)
    return counts


def table_exists(name):
    return name in set(inspect(db.engine).get_table_names())


def table_count(name):
    if not table_exists(name):
        return 0
    return db.session.execute(text(f'SELECT COUNT(*) FROM "{name}"')).scalar() or 0


# ---------------------------------------------------------------------------
# first-run seeding
# ---------------------------------------------------------------------------
def seed_database(force_import=False):
    """Create ORM tables, import Excel (if empty), create users + sample events."""
    db.create_all()
    result = {"imported": {}, "users": 0, "events": 0}

    if force_import or table_count("incidents") == 0:
        result["imported"] = import_excel_to_db()

    for spec in C.SEED_USERS:
        if not User.query.filter_by(username=spec["username"]).first():
            u = User(username=spec["username"], name=spec["name"],
                     email=spec["email"], role=spec["role"], active=True)
            u.set_password(spec["password"])
            db.session.add(u)
            result["users"] += 1

    if Event.query.count() == 0:
        result["events"] = _seed_events()

    db.session.commit()
    return result


def _seed_events(n=45):
    """A handful of individually-reported Near Miss / Hazard / Observation records."""
    rng = np.random.default_rng(C.RNG_SEED + 7)
    today = dt.date.today()
    cats = C.EVENT_CATEGORIES
    descs = {
        "Near Miss": ["Dropped object near walkway", "Vehicle reversing without spotter",
                      "Slippery floor in plant", "Near contact with mobile equipment",
                      "Hose under pressure came loose"],
        "Hazard": ["Damaged guardrail", "Exposed cabling", "Poor lighting in store",
                   "Spill not bunded", "Missing fire extinguisher signage"],
        "Observation": ["Good use of PPE by crew", "Housekeeping needs attention",
                        "Tool left on platform", "Positive isolation observed",
                        "Smoking outside designated area"],
    }
    made = 0
    for i in range(1, n + 1):
        cat = rng.choice(cats, p=[0.5, 0.3, 0.2])
        area = rng.choice(C.AREAS)
        d = today - dt.timedelta(days=int(rng.integers(1, 150)))
        sev = int(rng.choice([1, 2, 3], p=[0.55, 0.32, 0.13]))
        status = rng.choice(C.EVENT_STATUS, p=[0.45, 0.2, 0.35])
        db.session.add(Event(
            ref=f"EVT-{i:04d}", category=str(cat), date=d, area=area,
            department=C.AREA_DEPT[area], severity=sev,
            description=str(rng.choice(descs[str(cat)])),
            reported_by=str(rng.choice(C.OWNERS)), status=str(status)))
        made += 1
    return made


# ---------------------------------------------------------------------------
# references
# ---------------------------------------------------------------------------
def next_ref(table, col, prefix, width=4):
    """Next sequential reference, e.g. INC-0042, for a given table/column."""
    values = []
    try:
        if table_exists(table):
            values = db.session.execute(text(f'SELECT "{col}" FROM "{table}"')).scalars().all()
    except Exception:
        values = []
    mx = 0
    for v in values:
        if isinstance(v, str) and v.startswith(prefix + "-"):
            try:
                mx = max(mx, int(v.split("-")[-1]))
            except ValueError:
                pass
    return f"{prefix}-{mx + 1:0{width}d}"


# ---------------------------------------------------------------------------
# writes used by the capture forms
# ---------------------------------------------------------------------------
def insert_incident(row: dict):
    pd.DataFrame([row]).to_sql("incidents", db.engine, if_exists="append", index=False)


def insert_action(row: dict):
    pd.DataFrame([row]).to_sql("actions", db.engine, if_exists="append", index=False)


def update_action_status(action_id: str, status: str):
    db.session.execute(text('UPDATE actions SET "Status"=:s WHERE "Action_ID"=:a'),
                       {"s": status, "a": action_id})
    db.session.commit()


def log_audit(username, action, entity, ref="", detail=""):
    db.session.add(AuditLog(username=username, action=action, entity=entity,
                            ref=ref, detail=detail))
    db.session.commit()
