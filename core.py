#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
core.py -- canonical-store analytics engine.

The web app now imports spreadsheets into a local SQLite store first, then reads
dashboard data from that canonical store. Refresh is an audited import, not a
silent re-read of workbooks into process memory.
"""
from __future__ import annotations

import datetime as dt
import json
import os
from functools import partial

import pandas as pd
from sqlalchemy import desc, select
from sqlalchemy.orm import Session

import config as C
from hse_dashboard.database import (
    DEFAULT_DB_PATH,
    DEFAULT_SITE_ID,
    has_operational_rows,
    import_version,
    initialize_database,
    make_engine,
)
from hse_dashboard.importers import import_spreadsheets
from hse_dashboard.models import (
    Activity,
    Audit,
    ComplianceItem,
    CorrectiveAction,
    EnvironmentalRecord,
    Equipment,
    ImportRun,
    Incident,
    Permit,
)

TODAY = pd.Timestamp(dt.date.today())


def _today():
    return pd.Timestamp(dt.date.today())


def lower_better(val, target):
    if val <= target:
        return "good"
    if val <= target * 1.15:
        return "warn"
    return "bad"


def higher_better(val, target):
    if val >= target:
        return "good"
    if val >= target * 0.9:
        return "warn"
    return "bad"


def zero_best(val):
    return "good" if val == 0 else "bad"


def days_lti_status(val):
    if val is None:
        return "warn"
    if val >= C.THRESH_LTI_GOOD:
        return "good"
    if val >= C.THRESH_LTI_WARN:
        return "warn"
    return "bad"


class DataStore:
    """Reads canonical data from SQLite and computes dashboard analytics."""

    def __init__(self, data_dir: str = C.DATA_DIR, db_path: str = DEFAULT_DB_PATH,
                 site_id: str = DEFAULT_SITE_ID, auto_import: bool = True):
        self.data_dir = data_dir
        self.db_path = db_path
        self.site_id = site_id
        self.engine = make_engine(db_path)
        initialize_database(self.engine, site_id=site_id)
        self.frames: dict[str, pd.DataFrame] = {}
        self.status: list[dict] = []
        self.loaded_at: dt.datetime | None = None
        self.data_version = "empty"
        self._cache: dict[tuple, object] = {}

        if auto_import:
            with Session(self.engine) as session:
                if not has_operational_rows(session, site_id):
                    import_spreadsheets(site_id=site_id, data_dir=data_dir, engine=self.engine)
        self.load()

    # ----- loading / refreshing -------------------------------------------
    def load(self):
        global TODAY
        TODAY = _today()
        self.frames = self._load_frames()
        self.status = self._load_status()
        with Session(self.engine) as session:
            self.data_version = import_version(session, self.site_id)
        self.loaded_at = dt.datetime.now()
        self._cache.clear()
        return self

    def refresh(self):
        import_spreadsheets(site_id=self.site_id, data_dir=self.data_dir, engine=self.engine)
        return self.load()

    def df(self, key) -> pd.DataFrame:
        return self.frames.get(key, pd.DataFrame())

    def _query_dataframe(self, session: Session, model, mapping: dict[str, str]) -> pd.DataFrame:
        rows = session.execute(select(model).where(model.site_id == self.site_id)).scalars().all()
        records = []
        for obj in rows:
            records.append({column: getattr(obj, attr) for column, attr in mapping.items()})
        return pd.DataFrame(records, columns=list(mapping.keys()))

    def _load_frames(self) -> dict[str, pd.DataFrame]:
        with Session(self.engine) as session:
            frames = {
                "incidents": self._query_dataframe(session, Incident, {
                    "ID": "incident_id", "Date": "date", "Area": "area",
                    "Department": "department", "Company": "company", "Type": "type",
                    "Class": "incident_class", "Severity": "severity", "Status": "status",
                    "Reported": "reported", "CAR_Due": "car_due", "Owner": "owner",
                }),
                "activity": self._query_dataframe(session, Activity, {
                    "Period": "period", "Area": "area", "Department": "department",
                    "Company": "company", "ManHours": "manhours", "NearMisses": "nearmisses",
                    "Hazards": "hazards", "ObsRaised": "obsraised", "ObsClosed": "obsclosed",
                    "InspPlanned": "inspplanned", "InspDone": "inspdone",
                    "InspScore": "inspscore", "Audits": "audits", "Toolbox": "toolbox",
                    "TrainAssigned": "trainassigned", "TrainCompleted": "traincompleted",
                    "PPE": "ppe",
                }),
                "actions": self._query_dataframe(session, CorrectiveAction, {
                    "Action_ID": "action_id", "Source_Incident": "source_incident",
                    "Description": "description", "Raised": "raised", "Due": "due",
                    "Owner": "owner", "Department": "department", "Area": "area",
                    "Priority": "priority", "Status": "status",
                }),
                "compliance": self._query_dataframe(session, ComplianceItem, {
                    "Item": "item", "Regulator": "regulator", "Reference": "reference",
                    "Frequency_Months": "frequency_months",
                    "Last_Completed": "last_completed", "Owner": "owner",
                }),
                "environmental": self._query_dataframe(session, EnvironmentalRecord, {
                    "Period": "period", "Waste_t": "waste_t", "Recycling": "recycling",
                    "Energy_MWh": "energy_mwh", "Water_m3": "water_m3",
                    "Fuel_L": "fuel_l", "PM10": "pm10", "pH": "ph", "WAD_CN": "wad_cn",
                }),
                "permits": self._query_dataframe(session, Permit, {
                    "Permit": "permit", "Authority": "authority", "Holder": "holder",
                    "Issue_Date": "issue_date", "Expiry_Date": "expiry_date",
                }),
                "audits": self._query_dataframe(session, Audit, {
                    "Audit": "audit", "Type": "type", "Date": "date", "Auditor": "auditor",
                    "Score": "score", "Findings": "findings",
                    "Closed_Findings": "closed_findings",
                }),
                "equipment": self._query_dataframe(session, Equipment, {
                    "Asset_ID": "asset_id", "Asset": "asset", "Type": "type",
                    "Location": "location", "Last_Inspection": "last_inspection",
                    "Next_Inspection": "next_inspection",
                }),
            }
        return {key: self._derive(key, df) for key, df in frames.items()}

    def _load_status(self) -> list[dict]:
        source_specs = [
            (key, spec["file"], "standard") for key, spec in C.DATASETS.items()
        ] + [
            (spec["dataset"], spec["file"], profile)
            for profile, spec in C.WORKBOOK_PROFILES.items()
            if os.path.exists(os.path.join(self.data_dir, spec["file"]))
        ]
        with Session(self.engine) as session:
            runs = session.execute(
                select(ImportRun)
                .where(ImportRun.site_id == self.site_id)
                .order_by(desc(ImportRun.id))
            ).scalars().all()
        latest = {}
        for run in runs:
            latest.setdefault((run.source_file, run.profile), run)

        status = []
        for key, filename, profile in source_specs:
            path = os.path.join(self.data_dir, filename)
            run = latest.get((filename, profile))
            errors = json.loads(run.errors_json) if run and run.errors_json else []
            fatal_errors = [e for e in errors if e.get("message") != "ignored date-only rows"]
            status.append({
                "key": key if profile == "standard" else f"{key}:{profile}",
                "file": filename,
                "profile": profile,
                "rows": run.rows_accepted if run else 0,
                "rejected": run.rows_rejected if run else 0,
                "loaded": bool(run and run.rows_accepted > 0 and os.path.exists(path)),
                "modified": dt.datetime.fromtimestamp(os.path.getmtime(path)) if os.path.exists(path) else None,
                "error": "; ".join(e.get("message", "") for e in fatal_errors[:3]) if fatal_errors else None,
                "import_id": run.id if run else None,
                "imported_at": run.finished_at if run else None,
            })
        return status

    # ----- per-dataset derivations ----------------------------------------
    def _derive(self, key, df: pd.DataFrame) -> pd.DataFrame:
        df = df.copy()
        if df.empty:
            return df

        if key == "incidents":
            df["Date"] = pd.to_datetime(df["Date"])
            for c in ("Reported", "CAR_Due"):
                if c in df:
                    df[c] = pd.to_datetime(df[c], errors="coerce")
            df["Year"] = df["Date"].dt.year
            df["Month"] = df["Date"].dt.month
            df["MonthName"] = df["Date"].dt.strftime("%b")
            df["MonthKey"] = df["Date"].dt.strftime("%Y-%m")
            df["Recordable"] = df["Type"].isin(C.RECORDABLE_TYPES).astype(int)
            df["LTI"] = (df["Type"] == "Lost Time Injury").astype(int)
            df["CAR_Overdue"] = ((df["Status"] != "Closed") & df["CAR_Due"].notna()
                                 & (df["CAR_Due"] < TODAY)).astype(int)

        elif key == "activity":
            df["Period"] = pd.to_datetime(df["Period"])
            df["Year"] = df["Period"].dt.year
            df["Month"] = df["Period"].dt.month
            df["MonthName"] = df["Period"].dt.strftime("%b")
            df["MonthKey"] = df["Period"].dt.strftime("%Y-%m")

        elif key == "actions":
            for c in ("Raised", "Due"):
                df[c] = pd.to_datetime(df[c], errors="coerce")
            basis = df["Due"].fillna(df["Raised"])
            df["Year"] = basis.dt.year
            df["Month"] = basis.dt.month
            df["MonthName"] = basis.dt.strftime("%b")
            df["Overdue"] = ((df["Status"] != "Closed") & (df["Due"] < TODAY)).astype(int)
            df["DaysOverdue"] = ((TODAY - df["Due"]).dt.days).where(df["Overdue"] == 1, 0)

        elif key == "compliance":
            df["Last_Completed"] = pd.to_datetime(df["Last_Completed"], errors="coerce")
            df["Due_Date"] = df.apply(
                lambda r: r["Last_Completed"] + pd.DateOffset(months=int(r["Frequency_Months"]))
                if pd.notna(r["Last_Completed"]) else pd.NaT,
                axis=1)
            df["DaysToDue"] = (df["Due_Date"] - TODAY).dt.days
            df["Status"] = df["DaysToDue"].apply(
                lambda d: "Overdue" if pd.notna(d) and d < 0
                else ("Due Soon" if pd.notna(d) and d <= 30 else "Compliant"))

        elif key == "environmental":
            df["Period"] = pd.to_datetime(df["Period"])
            df["Year"] = df["Period"].dt.year
            df["MonthName"] = df["Period"].dt.strftime("%b")
            df["MonthKey"] = df["Period"].dt.strftime("%Y-%m")
            df["PM10_Limit"] = C.PM10_LIMIT
            df["CN_Limit"] = C.WADCN_LIMIT
            df["PM10_OK"] = (df["PM10"] <= C.PM10_LIMIT).astype(int)
            df["pH_OK"] = ((df["pH"] >= C.PH_MIN) & (df["pH"] <= C.PH_MAX)).astype(int)
            df["CN_OK"] = (df["WAD_CN"] <= C.WADCN_LIMIT).astype(int)

        elif key == "permits":
            for c in ("Issue_Date", "Expiry_Date"):
                df[c] = pd.to_datetime(df[c], errors="coerce")
            df["DaysToExpiry"] = (df["Expiry_Date"] - TODAY).dt.days
            df["Status"] = df["DaysToExpiry"].apply(
                lambda d: "Expired" if pd.notna(d) and d < 0
                else ("Expiring" if pd.notna(d) and d <= 60 else "Active"))

        elif key == "audits":
            df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
            df["Status"] = (df["Closed_Findings"] >= df["Findings"]).map({True: "Closed", False: "Open"})

        elif key == "equipment":
            for c in ("Last_Inspection", "Next_Inspection"):
                df[c] = pd.to_datetime(df[c], errors="coerce")
            df["DaysToDue"] = (df["Next_Inspection"] - TODAY).dt.days
            df["Status"] = df["DaysToDue"].apply(
                lambda d: "Overdue" if pd.notna(d) and d < 0
                else ("Due Soon" if pd.notna(d) and d <= 30 else "Compliant"))

        return df

    # ----- filtering / cache helpers --------------------------------------
    @staticmethod
    def _filter(df, f, year=True, month=True, dept=True, area=True):
        if df.empty:
            return df
        mask = pd.Series(True, index=df.index)
        if year and f.get("year", "All") != "All" and "Year" in df:
            mask &= df["Year"] == int(f["year"])
        if month and f.get("month", "All") != "All" and "MonthName" in df:
            mask &= df["MonthName"] == f["month"]
        if dept and f.get("dept", "All") != "All" and "Department" in df:
            mask &= df["Department"] == f["dept"]
        if area and f.get("area", "All") != "All" and "Area" in df:
            mask &= df["Area"] == f["area"]
        return df[mask]

    @staticmethod
    def _filter_key(f) -> tuple:
        return tuple((k, f.get(k, "All")) for k in ("year", "month", "dept", "area"))

    def _cached(self, name: str, f: dict, fn):
        key = (self.data_version, name, self._filter_key(f))
        if key not in self._cache:
            self._cache[key] = fn()
        return self._cache[key]

    def _timeline(self, f=None):
        """Ordered list of (MonthKey, label) across the selected activity history."""
        act = self.df("activity")
        if act.empty:
            return []
        if f:
            act = self._filter(act, f, dept=False, area=False)
        keys = sorted(act["MonthKey"].dropna().unique())
        labels = {k: pd.to_datetime(k + "-01").strftime("%b-%y") for k in keys}
        return [(k, labels[k]) for k in keys]

    # ----- KPIs ------------------------------------------------------------
    def kpis(self, f):
        return self._cached("kpis", f, partial(self._kpis, f))

    def _kpis(self, f):
        inc = self._filter(self.df("incidents"), f)
        act = self._filter(self.df("activity"), f)
        acts = self._filter(self.df("actions"), f)

        manhours = float(act["ManHours"].sum()) if not act.empty else 0.0
        recordable = int(inc["Recordable"].sum()) if not inc.empty else 0
        lti = int(inc["LTI"].sum()) if not inc.empty else 0
        near = int(act["NearMisses"].sum()) if not act.empty else 0

        trifr = (recordable / manhours * C.RATE_BASE) if manhours else 0.0
        ltifr = (lti / manhours * C.RATE_BASE) if manhours else 0.0

        if not act.empty and act["InspDone"].sum() > 0:
            insp = float((act["InspDone"] * act["InspScore"]).sum() / act["InspDone"].sum())
        else:
            insp = 0.0
        if not act.empty and act["TrainAssigned"].sum() > 0:
            train = float(act["TrainCompleted"].sum() / act["TrainAssigned"].sum())
        else:
            train = 0.0

        env = self.env_compliance(f)
        open_actions = int((acts["Status"] != "Closed").sum()) if not acts.empty else 0
        overdue_actions = int(acts["Overdue"].sum()) if not acts.empty else 0

        days_since = self.days_since_last_lti()

        def card(label, value, fmt, status, target=None, sub=""):
            return {"label": label, "value": value, "fmt": fmt,
                    "status": status, "target": target, "sub": sub}

        return [
            card("Total Incidents", len(inc), "int", "neutral", sub="period total"),
            card("Near Misses", near, "int",
                 higher_better(near, C.TARGET_NEARMISS_PERMONTH),
                 C.TARGET_NEARMISS_PERMONTH, "leading indicator"),
            card("TRIFR", round(trifr, 2), "dec", lower_better(trifr, C.TARGET_TRIFR),
                 C.TARGET_TRIFR, f"target <= {C.TARGET_TRIFR}"),
            card("LTIFR", round(ltifr, 2), "dec", lower_better(ltifr, C.TARGET_LTIFR),
                 C.TARGET_LTIFR, f"target <= {C.TARGET_LTIFR}"),
            card("Inspection Score", insp, "pct", higher_better(insp, C.TARGET_INSPECTION),
                 C.TARGET_INSPECTION, f"target {C.TARGET_INSPECTION:.0%}"),
            card("Training Completion", train, "pct", higher_better(train, C.TARGET_TRAINING),
                 C.TARGET_TRAINING, f"target {C.TARGET_TRAINING:.0%}"),
            card("Env Compliance", env, "pct", higher_better(env, C.TARGET_ENV),
                 C.TARGET_ENV, "readings within limit"),
            card("Days Since Last LTI", days_since if days_since is not None else "N/A",
                 "int", days_lti_status(days_since), sub="site-wide"),
            card("Open Actions", open_actions, "int", "neutral", sub="awaiting close-out"),
            card("Overdue Actions", overdue_actions, "int", zero_best(overdue_actions),
                 sub="past due date"),
            card("Recordables", recordable, "int", "neutral", sub="MTC + RWC + LTI"),
            card("Man-Hours", int(manhours), "int", "neutral", sub="exposure (period)"),
        ]

    def days_since_last_lti(self):
        inc = self.df("incidents")
        if inc.empty:
            return None
        ltis = inc.loc[inc["LTI"] == 1, "Date"]
        if ltis.empty:
            return None
        return int((TODAY - ltis.max()).days)

    def env_compliance(self, f):
        env = self._filter(self.df("environmental"), f, dept=False, area=False)
        if env.empty:
            return 0.0
        ok = int(env["PM10_OK"].sum() + env["pH_OK"].sum() + env["CN_OK"].sum())
        return ok / (3 * len(env))

    # ----- chart series ----------------------------------------------------
    def monthly_trend(self, f):
        return self._cached("monthly_trend", f, partial(self._monthly_trend, f))

    def _monthly_trend(self, f):
        inc = self._filter(self.df("incidents"), f)
        act = self._filter(self.df("activity"), f)
        labels, total, rec, near, trifr, insp = [], [], [], [], [], []
        for key, label in self._timeline(f):
            i = inc[inc["MonthKey"] == key] if not inc.empty else inc
            a = act[act["MonthKey"] == key] if not act.empty else act
            mh = float(a["ManHours"].sum()) if not a.empty else 0.0
            r = int(i["Recordable"].sum()) if not i.empty else 0
            labels.append(label)
            total.append(int(len(i)))
            rec.append(r)
            near.append(int(a["NearMisses"].sum()) if not a.empty else 0)
            trifr.append(round(r / mh * C.RATE_BASE, 2) if mh else 0)
            if not a.empty and a["InspDone"].sum() > 0:
                insp.append(round(float((a["InspDone"] * a["InspScore"]).sum()
                                         / a["InspDone"].sum()) * 100, 1))
            else:
                insp.append(0)
        return {"labels": labels, "incidents": total, "recordables": rec,
                "near_misses": near, "trifr": trifr, "inspection": insp}

    def by_type(self, f):
        return self._cached("by_type", f, partial(self._by_type, f))

    def _by_type(self, f):
        inc = self._filter(self.df("incidents"), f)
        counts = {t: 0 for t in C.INCIDENT_TYPES}
        if not inc.empty:
            for t, n in inc["Type"].fillna("Other").replace("", "Other").value_counts().items():
                counts[t] = int(n)
        return {"labels": list(counts.keys()), "data": list(counts.values())}

    def by_location(self, f, top=12):
        return self._cached("by_location", f, partial(self._by_location, f, top))

    def _by_location(self, f, top=12):
        inc = self._filter(self.df("incidents"), f)
        if inc.empty:
            return {"labels": [], "data": []}
        vc = inc["Area"].fillna("Unknown").replace("", "Unknown").value_counts().head(top)
        return {"labels": list(vc.index), "data": [int(x) for x in vc.values]}

    def actions_by_status(self, f):
        return self._cached("actions_by_status", f, partial(self._actions_by_status, f))

    def _actions_by_status(self, f):
        acts = self._filter(self.df("actions"), f)
        counts = {s: 0 for s in C.ACTION_STATUS}
        if not acts.empty:
            for s, n in acts["Status"].value_counts().items():
                counts[s] = int(n)
        return {"labels": list(counts.keys()), "data": list(counts.values())}

    def training_by_dept(self, f):
        return self._cached("training_by_dept", f, partial(self._training_by_dept, f))

    def _training_by_dept(self, f):
        act = self._filter(self.df("activity"), f, dept=False)
        labels, data = [], []
        if not act.empty:
            g = act.groupby("Department")[["TrainCompleted", "TrainAssigned"]].sum()
            for dept, row in g.iterrows():
                labels.append(dept)
                data.append(round(row["TrainCompleted"] / row["TrainAssigned"] * 100, 1)
                            if row["TrainAssigned"] else 0)
        return {"labels": labels, "data": data, "target": round(C.TARGET_TRAINING * 100)}

    # ----- table builders --------------------------------------------------
    def recent_incidents(self, f, n=8):
        inc = self._filter(self.df("incidents"), f)
        if inc.empty:
            return []
        inc = inc.sort_values("Date", ascending=False).head(n)
        return [{"id": r.ID, "date": r.Date.strftime("%d-%b-%Y"), "area": r.Area,
                 "type": r.Type, "severity": int(r.Severity), "status": r.Status}
                for r in inc.itertuples()]

    def incidents_table(self, f):
        inc = self._filter(self.df("incidents"), f).sort_values("Date", ascending=False)
        rows = []
        for r in inc.itertuples():
            rows.append({
                "id": r.ID, "date": r.Date.strftime("%d-%b-%Y"), "area": r.Area,
                "department": r.Department, "company": r.Company, "type": r.Type,
                "klass": r.Class, "severity": int(r.Severity), "status": r.Status,
                "owner": r.Owner,
                "car_due": r.CAR_Due.strftime("%d-%b-%Y") if pd.notna(r.CAR_Due) else "",
                "overdue": int(r.CAR_Overdue),
            })
        return rows

    def actions_table(self, f):
        acts = self._filter(self.df("actions"), f)
        if acts.empty:
            return []
        acts = acts.sort_values(["Overdue", "Due"], ascending=[False, True])
        return [{
            "id": r.Action_ID, "source": r.Source_Incident, "desc": r.Description,
            "raised": r.Raised.strftime("%d-%b-%Y") if pd.notna(r.Raised) else "",
            "due": r.Due.strftime("%d-%b-%Y") if pd.notna(r.Due) else "",
            "owner": r.Owner, "area": r.Area, "priority": r.Priority,
            "status": r.Status, "overdue": int(r.Overdue),
            "days_overdue": int(r.DaysOverdue),
        } for r in acts.itertuples()]

    def compliance_table(self):
        df = self.df("compliance")
        if df.empty:
            return [], {"pct": 0, "counts": {}}
        rows = [{
            "item": r.Item, "regulator": r.Regulator, "reference": r.Reference,
            "frequency": int(r.Frequency_Months),
            "last": r.Last_Completed.strftime("%d-%b-%Y") if pd.notna(r.Last_Completed) else "",
            "due": r.Due_Date.strftime("%d-%b-%Y") if pd.notna(r.Due_Date) else "",
            "days": int(r.DaysToDue) if pd.notna(r.DaysToDue) else "",
            "owner": r.Owner, "status": r.Status,
        } for r in df.itertuples()]
        counts = df["Status"].value_counts().to_dict()
        pct = round(df["Status"].eq("Compliant").mean() * 100, 1)
        return rows, {"pct": pct, "counts": {k: int(v) for k, v in counts.items()}}

    def environmental_view(self, f):
        env = self._filter(self.df("environmental"), f, dept=False, area=False)
        if env.empty:
            return {"rows": [], "chart": {}, "pct": 0}
        env = env.sort_values("Period")
        rows = [{
            "period": r.Period.strftime("%b-%Y"), "waste": r.Waste_t,
            "recycling": round(r.Recycling * 100, 1), "energy": int(r.Energy_MWh),
            "water": int(r.Water_m3), "fuel": int(r.Fuel_L),
            "pm10": r.PM10, "ph": r.pH, "cn": r.WAD_CN,
            "pm10_ok": int(r.PM10_OK), "ph_ok": int(r.pH_OK), "cn_ok": int(r.CN_OK),
        } for r in env.itertuples()]
        chart = {
            "labels": [r["period"] for r in rows],
            "pm10": [r["pm10"] for r in rows], "pm10_limit": C.PM10_LIMIT,
            "cn": [r["cn"] for r in rows], "cn_limit": C.WADCN_LIMIT,
            "ph": [r["ph"] for r in rows], "ph_min": C.PH_MIN, "ph_max": C.PH_MAX,
            "energy": [r["energy"] for r in rows], "water": [r["water"] for r in rows],
        }
        ok = int(env["PM10_OK"].sum() + env["pH_OK"].sum() + env["CN_OK"].sum())
        pct = round(ok / (3 * len(env)) * 100, 1)
        return {"rows": rows, "chart": chart, "pct": pct}

    def contractors_view(self, f):
        inc = self._filter(self.df("incidents"), f)
        act = self._filter(self.df("activity"), f)
        rows = []
        for co in C.COMPANIES:
            mh = float(act.loc[act["Company"] == co, "ManHours"].sum()) if not act.empty else 0.0
            rec = int(inc.loc[inc["Company"] == co, "Recordable"].sum()) if not inc.empty else 0
            lti = int(inc.loc[inc["Company"] == co, "LTI"].sum()) if not inc.empty else 0
            rows.append({
                "company": co, "scope": C.CONTRACTOR_SCOPE.get(co, ""),
                "manhours": int(mh), "recordables": rec, "ltis": lti,
                "trifr": round(rec / mh * C.RATE_BASE, 2) if mh else 0,
                "ltifr": round(lti / mh * C.RATE_BASE, 2) if mh else 0,
            })
        rows = [r for r in rows if r["manhours"] > 0] or rows
        chart = {"labels": [r["company"].split(" (")[0] for r in rows],
                 "trifr": [r["trifr"] for r in rows], "target": C.TARGET_TRIFR}
        return {"rows": rows, "chart": chart}

    def register(self, key):
        df = self.df(key)
        if df.empty:
            return []
        out = []
        for r in df.to_dict("records"):
            row = {}
            for k, v in r.items():
                if isinstance(v, pd.Timestamp):
                    v = v.strftime("%d-%b-%Y")
                elif isinstance(v, float):
                    row[k] = round(v, 3)
                    continue
                row[k] = v
            out.append(row)
        return out

    # ----- status / filters ------------------------------------------------
    def data_status(self):
        total_rows = sum(len(v) for v in self.frames.values())
        return {
            "files": self.status,
            "loaded_at": self.loaded_at.strftime("%d-%b-%Y %H:%M:%S") if self.loaded_at else "-",
            "total_files": sum(1 for s in self.status if s["loaded"]),
            "total_rows": total_rows,
            "data_dir": self.data_dir,
            "db_path": self.db_path,
            "site_id": self.site_id,
            "version": self.data_version,
        }

    def filter_options(self):
        inc = self.df("incidents")
        years = sorted(int(y) for y in inc["Year"].dropna().unique().tolist()) if not inc.empty else []
        areas = sorted({*C.AREAS, *([a for a in inc["Area"].dropna().unique() if a])})
        depts = sorted({*C.DEPARTMENTS, *([d for d in inc["Department"].dropna().unique() if d])})
        return {
            "years": ["All"] + [str(y) for y in years],
            "months": ["All"] + ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                                  "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
            "departments": ["All"] + depts,
            "areas": ["All"] + areas,
        }
