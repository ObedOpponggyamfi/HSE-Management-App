#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Database bootstrap & validated import pipeline.

  - import_all()        : validate + load every data/*.xlsx into the DB, recording
                          an ImportRun (rows seen/accepted/rejected + errors) per source
  - import_rich_incidents(): map a real Asanko incident workbook (DataBase sheet) into
                          the incident register
  - seed_database()     : first-run setup (import data, create users, sample events)
  - insert_incident / insert_action / update_action_status : capture-form writes
  - next_ref()          : next INC-/CAR-/EVT- reference

Bulk datasets are plain tables (pandas <-> SQL); Users/Events/AuditLog/ImportRun
are ORM-managed (models.py).
"""
import datetime as dt
import json
import os

import numpy as np
import pandas as pd
from sqlalchemy import inspect, text

import config as C
from extensions import db
from models import Event, ImportRun, Investigation, User

# Minimum required (non-empty) columns for a standard row to be accepted.
REQUIRED = {
    "incidents": ["ID", "Date"], "activity": ["Period", "Area"],
    "actions": ["Action_ID"], "compliance": ["Item"], "environmental": ["Period"],
    "permits": ["Permit"], "audits": ["Audit"], "equipment": ["Asset"],
    "competency": ["Person", "Competency"],
    "tailings_inspections": ["TSF", "Date"], "piezometers": ["TSF", "Date"],
}


# ---------------------------------------------------------------------------
# small coercion helpers (shared with the rich workbook importer)
# ---------------------------------------------------------------------------
def _clean(value, default=""):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return default
    return str(value).strip()


def _date(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    value = pd.to_datetime(value, errors="coerce")
    return None if pd.isna(value) else value.date()


def _int(value, default=0):
    try:
        if value is None or pd.isna(value):
            return default
        return int(round(float(value)))
    except (TypeError, ValueError):
        return default


def _severity_from_risk(value):
    t = _clean(value).lower()
    if not t:
        return 1
    if any(k in t for k in ("critical", "catastrophic", "extreme", "5")):
        return 5
    if any(k in t for k in ("high", "major", "4")):
        return 4
    if any(k in t for k in ("medium", "moderate", "3")):
        return 3
    if any(k in t for k in ("low", "minor", "2")):
        return 2
    return 1


def _norm_incident_type(raw):
    t = _clean(raw).lower()
    if "lost time" in t or t in ("lti",):
        return "Lost Time Injury"
    if "restricted" in t or "rwc" in t:
        return "Restricted Work"
    if "medical" in t or t == "mtc":
        return "Medical Treatment"
    if "first aid" in t or "fac" in t:
        return "First Aid"
    if "property" in t or "damage" in t:
        return "Property Damage"
    if "environ" in t or "spill" in t:
        return "Environmental"
    return "Other"


def _record_run(dataset, source_file, profile, seen, accepted, rejected, errors):
    run = ImportRun(dataset=dataset, source_file=os.path.basename(source_file),
                    profile=profile, rows_seen=seen, rows_accepted=accepted,
                    rows_rejected=rejected, errors=json.dumps(errors[:25], default=str))
    db.session.add(run)
    db.session.commit()
    return {"dataset": dataset, "source_file": os.path.basename(source_file),
            "profile": profile, "rows_seen": seen, "rows_accepted": accepted,
            "rows_rejected": rejected, "import_id": run.id}


# ---------------------------------------------------------------------------
# standard dataset import (validated)
# ---------------------------------------------------------------------------
def _import_standard(dataset, spec):
    path = os.path.join(C.DATA_DIR, spec["file"])
    if not os.path.exists(path):
        return _record_run(dataset, spec["file"], "standard", 0, 0, 0, [{"row": None, "message": "file not found"}])
    df = pd.read_excel(path, sheet_name=spec["sheet"])
    seen = len(df)
    errors, keep = [], pd.Series(True, index=df.index)
    for col in REQUIRED.get(dataset, []):
        if col not in df.columns:
            keep &= False
            errors.append({"row": None, "message": f"missing column '{col}'"})
            continue
        ok = df[col].notna() & (df[col].astype(str).str.strip() != "")
        if dataset in ("incidents",) and col == "Date":
            ok &= pd.to_datetime(df[col], errors="coerce").notna()
        for idx in df.index[~ok]:
            errors.append({"row": int(idx) + 2, "message": f"empty/invalid '{col}'"})
        keep &= ok
    clean = df[keep].copy()
    clean.to_sql(dataset, db.engine, if_exists="replace", index=False)
    return _record_run(dataset, spec["file"], "standard", seen, len(clean), seen - len(clean), errors)


# ---------------------------------------------------------------------------
# rich Asanko incident workbook import
# ---------------------------------------------------------------------------
def import_rich_incidents(profile, spec):
    path = os.path.join(C.DATA_DIR, spec["file"])
    if not os.path.exists(path):
        return None
    df = pd.read_excel(path, sheet_name=spec["sheet"], header=spec.get("header", 1))
    seen = len(df)
    rows, errors, seen_ids, skipped, rejected = [], [], set(), 0, 0
    for offset, r in df.dropna(how="all").iterrows():
        row_no = int(offset) + spec.get("header", 1) + 2
        case_id = _clean(r.get("CASE ID"))
        when = _date(r.get("DATE OF INCIDENT"))
        if not case_id or when is None:
            skipped += 1            # blank / pre-formatted template / incomplete rows: skip, don't reject
            continue
        if case_id in seen_ids:
            rejected += 1           # a genuine conflict
            errors.append({"row": row_no, "message": f"duplicate CASE ID {case_id}"})
            continue
        seen_ids.add(case_id)
        risk = r.get("RISK RATING (Actual Consequence)") or r.get("RISK RATING (Potential Consequence)")
        rows.append({
            "ID": case_id, "Date": when,
            "Area": _clean(r.get("LOCATION")) or _clean(r.get("SPECIFIC LOCATION")),
            "Department": _clean(r.get("BUSINESS DEPARTMENT / UNIT")),
            "Company": _clean(r.get("BUSINESS PARTNER")) or "Owner",
            "Type": _norm_incident_type(r.get("INCIDENT TYPE")),
            "Class": _clean(r.get("INJURY OUTCOME")) or _clean(r.get("TYPE OF INJURY")),
            "Severity": _severity_from_risk(risk),
            "Status": _clean(r.get("STATUS"), "Closed"),
            "Reported": when, "CAR_Due": None,
            "Owner": _clean(r.get("LEAD INVESTIGATOR")) or "Unassigned",
        })
    if rows:
        pd.DataFrame(rows).to_sql("incidents", db.engine, if_exists="replace", index=False)
    if skipped:
        errors.append({"row": None, "message": f"{skipped} incomplete/blank template rows skipped"})
    return _record_run(spec["dataset"], spec["file"], profile, seen, len(rows), rejected, errors)


# ---------------------------------------------------------------------------
# orchestration
# ---------------------------------------------------------------------------
def import_all(profiles=True):
    results = []
    for dataset, spec in C.DATASETS.items():
        results.append(_import_standard(dataset, spec))
    if profiles:
        for profile, spec in C.WORKBOOK_PROFILES.items():
            res = import_rich_incidents(profile, spec)
            if res:                       # rich workbook overrides the standard incidents
                results.append(res)
    return results


def latest_import_runs():
    """dataset -> latest ImportRun (for the Data page)."""
    out = {}
    for run in ImportRun.query.order_by(ImportRun.id.asc()).all():
        out[run.dataset] = run
    return out


def table_count(name):
    if name not in set(inspect(db.engine).get_table_names()):
        return 0
    return db.session.execute(text(f'SELECT COUNT(*) FROM "{name}"')).scalar() or 0


# ---------------------------------------------------------------------------
# first-run seeding
# ---------------------------------------------------------------------------
def seed_database(force_import=False):
    db.create_all()
    result = {"imported": [], "users": 0, "events": 0}
    if force_import or table_count("incidents") == 0:
        result["imported"] = import_all()
    for spec in C.SEED_USERS:
        if not User.query.filter_by(username=spec["username"]).first():
            u = User(username=spec["username"], name=spec["name"], email=spec["email"],
                     role=spec["role"], active=True)
            u.set_password(spec["password"])
            db.session.add(u)
            result["users"] += 1
    if Event.query.count() == 0:
        result["events"] = _seed_events()
    if Investigation.query.count() == 0:
        result["investigations"] = _seed_investigations()
    db.session.commit()
    return result


def _seed_events(n=45):
    rng = np.random.default_rng(C.RNG_SEED + 7)
    today = dt.date.today()
    descs = {
        "Near Miss": ["Dropped object near walkway", "Vehicle reversing without spotter",
                      "Slippery floor in plant", "Near contact with mobile equipment"],
        "Hazard": ["Damaged guardrail", "Exposed cabling", "Poor lighting in store",
                   "Spill not bunded", "Missing fire extinguisher signage"],
        "Observation": ["Good use of PPE by crew", "Housekeeping needs attention",
                        "Tool left on platform", "Positive isolation observed"],
    }
    for i in range(1, n + 1):
        cat = str(rng.choice(C.EVENT_CATEGORIES, p=[0.5, 0.3, 0.2]))
        area = str(rng.choice(C.AREAS))
        db.session.add(Event(
            ref=f"EVT-{i:04d}", category=cat,
            date=today - dt.timedelta(days=int(rng.integers(1, 150))), area=area,
            department=C.AREA_DEPT[area], severity=int(rng.choice([1, 2, 3], p=[0.55, 0.32, 0.13])),
            description=str(rng.choice(descs[cat])), reported_by=str(rng.choice(C.OWNERS)),
            status=str(rng.choice(C.EVENT_STATUS, p=[0.45, 0.2, 0.35]))))
    return n


def _seed_investigations(n=6):
    rng = np.random.default_rng(C.RNG_SEED + 11)
    try:
        df = pd.read_sql('SELECT "ID","Type","Area" FROM incidents', db.engine)
    except Exception:
        df = pd.DataFrame()
    if df.empty:
        return 0
    pool = df[df["Type"].isin(C.RECORDABLE_TYPES)] if "Type" in df else df
    if pool.empty:
        pool = df
    picks = pool.head(40).sample(min(n, len(pool)), random_state=11)
    methods = ["5-Whys", "ICAM", "Fishbone"]
    whys = ["Task deviated from the safe procedure", "Procedure skipped under time pressure",
            "Supervision gap during shift change", "Refresher training overdue",
            "Inadequate planning / risk assessment of the task"]
    made = 0
    for i, (_, r) in enumerate(picks.iterrows(), start=1):
        db.session.add(Investigation(
            ref=f"INV-{i:04d}", incident_id=str(r["ID"]), hipo=bool(rng.random() < 0.4),
            method=str(rng.choice(methods)),
            immediate_cause=f"Immediate cause linked to {r.get('Type', 'event')} at {r.get('Area', 'site')}",
            root_cause=whys[-1], why1=whys[0], why2=whys[1], why3=whys[2], why4=whys[3], why5=whys[4],
            status=str(rng.choice(["Completed", "In Progress", "Open"], p=[0.5, 0.3, 0.2])),
            investigator=str(rng.choice(C.OWNERS)), created_by="seed"))
        made += 1
    return made


# ---------------------------------------------------------------------------
# references + capture-form writes
# ---------------------------------------------------------------------------
def next_ref(table, col, prefix, width=4):
    values = []
    try:
        if table in set(inspect(db.engine).get_table_names()):
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


def insert_incident(row):
    pd.DataFrame([row]).to_sql("incidents", db.engine, if_exists="append", index=False)


def insert_action(row):
    pd.DataFrame([row]).to_sql("actions", db.engine, if_exists="append", index=False)


def update_action_status(action_id, status):
    db.session.execute(text('UPDATE actions SET "Status"=:s WHERE "Action_ID"=:a'),
                       {"s": status, "a": action_id})
    db.session.commit()


def log_audit(username, action, entity, ref="", detail=""):
    from models import AuditLog
    db.session.add(AuditLog(username=username, action=action, entity=entity, ref=ref, detail=detail))
    db.session.commit()
