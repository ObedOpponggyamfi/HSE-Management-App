#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
core.py -- the consolidation + analytics engine (database-backed).

DataStore reads each safety dataset from the SQLite database (populated from the
Excel files by importer.py), derives the fields the dashboard needs
(Year/Month, Recordable/LTI flags, overdue/compliance status ...), keeps them in
memory, and recomputes all KPIs / chart series on demand. `refresh()` reloads
from the database, so newly captured records (or a re-import from Excel) surface
immediately.
"""
import datetime as dt
import os

import pandas as pd
from sqlalchemy import inspect

import config as C
from extensions import db

TODAY = pd.Timestamp(dt.date.today())


def _today():
    return pd.Timestamp(dt.date.today())


# ---------------------------------------------------------------------------
# status helpers (return 'good' / 'warn' / 'bad' for traffic lights)
# ---------------------------------------------------------------------------
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
    """Loads & consolidates the safety datasets from the DB; computes analytics."""

    def __init__(self, data_dir: str = C.DATA_DIR):
        self.data_dir = data_dir
        self.frames: dict[str, pd.DataFrame] = {}
        self.status: list[dict] = []
        self.loaded_at: dt.datetime | None = None
        self.load()

    # ----- loading / refreshing (from the database) -----------------------
    def load(self):
        global TODAY
        TODAY = _today()
        self.frames, self.status = {}, []
        try:
            existing = set(inspect(db.engine).get_table_names())
        except Exception:
            existing = set()

        for key in list(C.DATASETS.keys()) + ["events"]:
            entry = {"key": key, "rows": 0, "loaded": False, "error": None}
            try:
                if key in existing:
                    df = pd.read_sql_table(key, db.engine)
                    df = self._derive(key, df)
                    self.frames[key] = df
                    entry.update(loaded=not df.empty, rows=len(df))
                else:
                    self.frames[key] = pd.DataFrame()
                    entry["error"] = "not in database"
            except Exception as exc:
                self.frames[key] = pd.DataFrame()
                entry["error"] = str(exc)
            self.status.append(entry)
        self.loaded_at = dt.datetime.now()
        return self

    def refresh(self):
        return self.load()

    def df(self, key) -> pd.DataFrame:
        return self.frames.get(key, pd.DataFrame())

    # ----- per-dataset derivations ----------------------------------------
    def _derive(self, key, df: pd.DataFrame) -> pd.DataFrame:
        df = df.copy()
        if df.empty:
            return df

        if key == "incidents":
            df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
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
            df["Period"] = pd.to_datetime(df["Period"], errors="coerce")
            df["Year"] = df["Period"].dt.year
            df["Month"] = df["Period"].dt.month
            df["MonthName"] = df["Period"].dt.strftime("%b")
            df["MonthKey"] = df["Period"].dt.strftime("%Y-%m")

        elif key == "actions":
            for c in ("Raised", "Due"):
                df[c] = pd.to_datetime(df[c], errors="coerce")
            df["Overdue"] = ((df["Status"] != "Closed") & (df["Due"] < TODAY)).astype(int)
            df["DaysOverdue"] = ((TODAY - df["Due"]).dt.days).where(df["Overdue"] == 1, 0)

        elif key == "compliance":
            df["Last_Completed"] = pd.to_datetime(df["Last_Completed"], errors="coerce")
            df["Due_Date"] = df.apply(
                lambda r: r["Last_Completed"] + pd.DateOffset(months=int(r["Frequency_Months"])),
                axis=1)
            df["DaysToDue"] = (df["Due_Date"] - TODAY).dt.days
            df["Status"] = df["DaysToDue"].apply(
                lambda d: "Overdue" if d < 0 else ("Due Soon" if d <= 30 else "Compliant"))

        elif key == "environmental":
            df["Period"] = pd.to_datetime(df["Period"], errors="coerce")
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
                lambda d: "Expired" if d < 0 else ("Expiring" if d <= 60 else "Active"))

        elif key == "audits":
            df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
            df["Status"] = (df["Closed_Findings"] >= df["Findings"]).map({True: "Closed", False: "Open"})

        elif key == "equipment":
            for c in ("Last_Inspection", "Next_Inspection"):
                df[c] = pd.to_datetime(df[c], errors="coerce")
            df["DaysToDue"] = (df["Next_Inspection"] - TODAY).dt.days
            df["Status"] = df["DaysToDue"].apply(
                lambda d: "Overdue" if d < 0 else ("Due Soon" if d <= 30 else "Compliant"))

        elif key == "tailings_inspections":
            df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
            df["MonthKey"] = df["Date"].dt.strftime("%b-%y")
            df["DaysSince"] = (TODAY - df["Date"]).dt.days

        elif key == "piezometers":
            df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
            df["MonthKey"] = df["Date"].dt.strftime("%Y-%m")
            df["Exceedance"] = (df["Reading_m"] > df["Threshold_m"]).astype(int)
            df["Status"] = df.apply(
                lambda r: "Exceedance" if r["Reading_m"] > r["Threshold_m"]
                else ("Elevated" if r["Reading_m"] > 0.9 * r["Threshold_m"] else "Normal"), axis=1)

        elif key == "competency":
            for c in ("Completed", "Expiry"):
                df[c] = pd.to_datetime(df[c], errors="coerce")
            df["DaysToExpiry"] = (df["Expiry"] - TODAY).dt.days
            df["Status"] = df["DaysToExpiry"].apply(
                lambda d: "Expired" if pd.isna(d) or d < 0 else ("Expiring" if d <= 60 else "Valid"))

        elif key == "events":
            df = df.rename(columns={
                "ref": "Ref", "category": "Category", "date": "Date", "area": "Area",
                "department": "Department", "severity": "Severity",
                "description": "Description", "reported_by": "Reported_By", "status": "Status"})
            df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
            df["Year"] = df["Date"].dt.year
            df["MonthName"] = df["Date"].dt.strftime("%b")
            df["MonthKey"] = df["Date"].dt.strftime("%Y-%m")

        return df

    # ----- filtering -------------------------------------------------------
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

    def _timeline(self):
        act = self.df("activity")
        if act.empty:
            return []
        keys = sorted(act["MonthKey"].dropna().unique())
        return [(k, pd.to_datetime(k + "-01").strftime("%b-%y")) for k in keys]

    # ----- KPIs ------------------------------------------------------------
    def kpis(self, f):
        inc = self._filter(self.df("incidents"), f)
        act = self._filter(self.df("activity"), f)
        acts = self._filter(self.df("actions"), f, year=False, month=False)
        evt = self._filter(self.df("events"), f)

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
        reports = int(len(evt))

        def card(label, value, fmt, status, target=None, sub=""):
            return {"label": label, "value": value, "fmt": fmt,
                    "status": status, "target": target, "sub": sub}

        return [
            card("Total Incidents", len(inc), "int", "neutral", sub="period total"),
            card("Near Misses", near, "int",
                 higher_better(near, C.TARGET_NEARMISS_PERMONTH),
                 C.TARGET_NEARMISS_PERMONTH, "leading indicator"),
            card("TRIFR", round(trifr, 2), "dec", lower_better(trifr, C.TARGET_TRIFR),
                 C.TARGET_TRIFR, f"target ≤ {C.TARGET_TRIFR}"),
            card("LTIFR", round(ltifr, 2), "dec", lower_better(ltifr, C.TARGET_LTIFR),
                 C.TARGET_LTIFR, f"target ≤ {C.TARGET_LTIFR}"),
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
            card("Event Reports", reports, "int", "neutral", sub="near miss / hazard / obs"),
            card("Man-Hours", int(manhours), "int", "neutral", sub="exposure (period)"),
        ]

    def days_since_last_lti(self):
        inc = self.df("incidents")
        if inc.empty:
            return None
        ltis = inc.loc[inc["LTI"] == 1, "Date"]
        if ltis.empty or ltis.isna().all():
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
        inc = self._filter(self.df("incidents"), f, year=False, month=False)
        act = self._filter(self.df("activity"), f, year=False, month=False)
        labels, total, rec, near, trifr, insp = [], [], [], [], [], []
        for key, label in self._timeline():
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
        inc = self._filter(self.df("incidents"), f)
        counts = {t: 0 for t in C.INCIDENT_TYPES}
        if not inc.empty:
            for t, n in inc["Type"].value_counts().items():
                counts[t] = int(n)
        return {"labels": list(counts.keys()), "data": list(counts.values())}

    def by_location(self, f, top=12):
        inc = self._filter(self.df("incidents"), f, area=False)
        if inc.empty:
            return {"labels": [], "data": []}
        vc = inc["Area"].value_counts().head(top)
        return {"labels": list(vc.index), "data": [int(x) for x in vc.values]}

    def actions_by_status(self, f):
        acts = self._filter(self.df("actions"), f, year=False, month=False)
        counts = {s: 0 for s in C.ACTION_STATUS}
        if not acts.empty:
            for s, n in acts["Status"].value_counts().items():
                counts[s] = int(n)
        return {"labels": list(counts.keys()), "data": list(counts.values())}

    def training_by_dept(self, f):
        act = self._filter(self.df("activity"), f, dept=False, area=False)
        labels, data = [], []
        if not act.empty:
            g = act.groupby("Department")[["TrainCompleted", "TrainAssigned"]].sum()
            for dept, row in g.iterrows():
                labels.append(dept)
                data.append(round(row["TrainCompleted"] / row["TrainAssigned"] * 100, 1)
                            if row["TrainAssigned"] else 0)
        return {"labels": labels, "data": data, "target": round(C.TARGET_TRAINING * 100)}

    def rolling_rates(self, f, window=12):
        """Rolling 12-month TRIFR / LTIFR / AIFR (the standard reporting basis)."""
        inc = self._filter(self.df("incidents"), f, year=False, month=False)
        act = self._filter(self.df("activity"), f, year=False, month=False)
        injury = set(C.RECORDABLE_TYPES) | {"First Aid"}
        rec, lti, ai, mh, labels = [], [], [], [], []
        for key, label in self._timeline():
            i = inc[inc["MonthKey"] == key] if not inc.empty else inc
            a = act[act["MonthKey"] == key] if not act.empty else act
            labels.append(label)
            rec.append(int(i["Recordable"].sum()) if not i.empty else 0)
            lti.append(int(i["LTI"].sum()) if not i.empty else 0)
            ai.append(int(i["Type"].isin(injury).sum()) if not i.empty else 0)
            mh.append(float(a["ManHours"].sum()) if not a.empty else 0.0)

        def roll(arr, i):
            lo = max(0, i - window + 1)
            return sum(arr[lo:i + 1])

        trifr, ltifr, aifr = [], [], []
        for i in range(len(labels)):
            h = roll(mh, i)
            trifr.append(round(roll(rec, i) / h * C.RATE_BASE, 2) if h else 0)
            ltifr.append(round(roll(lti, i) / h * C.RATE_BASE, 2) if h else 0)
            aifr.append(round(roll(ai, i) / h * C.RATE_BASE, 2) if h else 0)
        latest = {"trifr": trifr[-1] if trifr else 0, "ltifr": ltifr[-1] if ltifr else 0,
                  "aifr": aifr[-1] if aifr else 0}
        return {"labels": labels, "trifr": trifr, "ltifr": ltifr, "aifr": aifr,
                "target_trifr": C.TARGET_TRIFR, "target_ltifr": C.TARGET_LTIFR,
                "latest": latest, "window": window}

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
                "overdue": int(r.CAR_Overdue)})
        return rows

    def actions_table(self, f):
        acts = self._filter(self.df("actions"), f, year=False, month=False)
        if acts.empty:
            return []
        acts = acts.sort_values(["Overdue", "Due"], ascending=[False, True])
        return [{
            "id": r.Action_ID, "source": r.Source_Incident, "desc": r.Description,
            "raised": r.Raised.strftime("%d-%b-%Y") if pd.notna(r.Raised) else "",
            "due": r.Due.strftime("%d-%b-%Y") if pd.notna(r.Due) else "",
            "owner": r.Owner, "area": r.Area, "priority": r.Priority,
            "status": r.Status, "overdue": int(r.Overdue),
            "days_overdue": int(r.DaysOverdue)} for r in acts.itertuples()]

    def open_actions_list(self):
        """For the action-update dropdown."""
        acts = self.df("actions")
        if acts.empty:
            return []
        return list(acts.loc[acts["Status"] != "Closed", "Action_ID"])

    def compliance_table(self):
        df = self.df("compliance")
        if df.empty:
            return [], {"pct": 0, "counts": {}}
        rows = [{
            "item": r.Item, "regulator": r.Regulator, "reference": r.Reference,
            "frequency": int(r.Frequency_Months),
            "last": r.Last_Completed.strftime("%d-%b-%Y"),
            "due": r.Due_Date.strftime("%d-%b-%Y"),
            "days": int(r.DaysToDue), "owner": r.Owner, "status": r.Status}
            for r in df.itertuples()]
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
            "pm10_ok": int(r.PM10_OK), "ph_ok": int(r.pH_OK), "cn_ok": int(r.CN_OK)}
            for r in env.itertuples()]
        chart = {
            "labels": [r["period"] for r in rows],
            "pm10": [r["pm10"] for r in rows], "pm10_limit": C.PM10_LIMIT,
            "cn": [r["cn"] for r in rows], "cn_limit": C.WADCN_LIMIT,
            "ph": [r["ph"] for r in rows], "ph_min": C.PH_MIN, "ph_max": C.PH_MAX,
            "energy": [r["energy"] for r in rows], "water": [r["water"] for r in rows]}
        ok = int(env["PM10_OK"].sum() + env["pH_OK"].sum() + env["CN_OK"].sum())
        pct = round(ok / (3 * len(env)) * 100, 1)
        return {"rows": rows, "chart": chart, "pct": pct}

    def contractors_view(self, f):
        inc = self._filter(self.df("incidents"), f, area=False)
        act = self._filter(self.df("activity"), f, area=False)
        rows = []
        for co in C.COMPANIES:
            mh = float(act.loc[act["Company"] == co, "ManHours"].sum()) if not act.empty else 0.0
            rec = int(inc.loc[inc["Company"] == co, "Recordable"].sum()) if not inc.empty else 0
            lti = int(inc.loc[inc["Company"] == co, "LTI"].sum()) if not inc.empty else 0
            rows.append({
                "company": co, "scope": C.CONTRACTOR_SCOPE.get(co, ""),
                "manhours": int(mh), "recordables": rec, "ltis": lti,
                "trifr": round(rec / mh * C.RATE_BASE, 2) if mh else 0,
                "ltifr": round(lti / mh * C.RATE_BASE, 2) if mh else 0})
        rows = [r for r in rows if r["manhours"] > 0] or rows
        chart = {"labels": [r["company"].split(" (")[0] for r in rows],
                 "trifr": [r["trifr"] for r in rows], "target": C.TARGET_TRIFR}
        return {"rows": rows, "chart": chart}

    def events_view(self, f):
        evt = self._filter(self.df("events"), f)
        cats = {c: 0 for c in C.EVENT_CATEGORIES}
        rows = []
        if not evt.empty:
            for c, n in evt["Category"].value_counts().items():
                cats[c] = int(n)
            evt = evt.sort_values("Date", ascending=False)
            for r in evt.itertuples():
                rows.append({
                    "ref": r.Ref, "category": r.Category,
                    "date": r.Date.strftime("%d-%b-%Y") if pd.notna(r.Date) else "",
                    "area": r.Area, "department": r.Department,
                    "severity": int(r.Severity) if pd.notna(r.Severity) else "",
                    "description": r.Description, "reporter": r.Reported_By,
                    "status": r.Status})
        # monthly trend across timeline
        ev_all = self._filter(self.df("events"), f, year=False, month=False)
        labels, series = [], []
        for key, label in self._timeline():
            labels.append(label)
            series.append(int((ev_all["MonthKey"] == key).sum()) if not ev_all.empty else 0)
        chart = {"cats": {"labels": list(cats.keys()), "data": list(cats.values())},
                 "trend": {"labels": labels, "data": series}}
        return {"rows": rows, "counts": cats, "total": len(rows), "chart": chart}

    def competency_view(self, f):
        df = self._filter(self.df("competency"), f, year=False, month=False, area=False)
        if df.empty:
            return {"total": 0, "valid": 0, "expiring": 0, "expired": 0, "pct": 0, "rows": [],
                    "chart": {"dept": {"labels": [], "data": []}, "type": {"labels": [], "data": []}}}
        total = len(df)
        valid = int((df["Status"] == "Valid").sum())
        expiring = int((df["Status"] == "Expiring").sum())
        expired = int((df["Status"] == "Expired").sum())
        dept_labels, dept_data = [], []
        for dept, g in df.groupby("Department"):
            dept_labels.append(dept)
            dept_data.append(round((g["Status"] == "Valid").mean() * 100, 1))
        type_labels, type_data = [], []
        for t, g in df.groupby("Type"):
            type_labels.append(t)
            type_data.append(len(g))
        rows = []
        for r in df.sort_values("DaysToExpiry").itertuples():
            rows.append({
                "person": r.Person, "department": r.Department, "competency": r.Competency,
                "type": r.Type,
                "completed": r.Completed.strftime("%d-%b-%Y") if pd.notna(r.Completed) else "",
                "expiry": r.Expiry.strftime("%d-%b-%Y") if pd.notna(r.Expiry) else "",
                "days": int(r.DaysToExpiry) if pd.notna(r.DaysToExpiry) else "",
                "status": r.Status})
        return {"total": total, "valid": valid, "expiring": expiring, "expired": expired,
                "pct": round(valid / total * 100, 1), "rows": rows,
                "chart": {"dept": {"labels": dept_labels, "data": dept_data},
                          "type": {"labels": type_labels, "data": type_data}}}

    def tailings_view(self):
        insp = self.df("tailings_inspections")
        piez = self.df("piezometers")
        kpi = {"last_status": "—", "days_since": None, "exceedances": 0,
               "min_freeboard": None, "facilities": 0}
        insp_rows, piez_rows = [], []
        chart = {"labels": [], "phreatic": [], "threshold": 0, "freeboard": []}
        if not insp.empty:
            insp = insp.sort_values("Date", ascending=False)
            latest = insp.iloc[0]
            kpi["last_status"] = latest["Status"]
            kpi["days_since"] = int(latest["DaysSince"]) if pd.notna(latest["DaysSince"]) else None
            kpi["min_freeboard"] = round(float(insp["Freeboard_m"].min()), 2)
            kpi["facilities"] = int(insp["TSF"].nunique())
            for r in insp.head(40).itertuples():
                insp_rows.append({"tsf": r.TSF,
                    "date": r.Date.strftime("%d-%b-%Y") if pd.notna(r.Date) else "",
                    "inspector": r.Inspector, "freeboard": r.Freeboard_m,
                    "status": r.Status, "findings": r.Findings})
        if not piez.empty:
            kpi["exceedances"] = int(piez["Exceedance"].sum())
            chart["threshold"] = round(float(piez["Threshold_m"].mean()), 2)
            for k in sorted(piez["MonthKey"].dropna().unique()):
                sub = piez[piez["MonthKey"] == k]
                chart["labels"].append(pd.to_datetime(k + "-01").strftime("%b-%y"))
                chart["phreatic"].append(round(float(sub["Reading_m"].mean()), 2))
            for pid, g in piez.sort_values("Date").groupby("Piezo_ID"):
                last = g.iloc[-1]
                piez_rows.append({"tsf": last["TSF"], "piezo": pid,
                    "date": last["Date"].strftime("%d-%b-%Y") if pd.notna(last["Date"]) else "",
                    "reading": last["Reading_m"], "threshold": last["Threshold_m"],
                    "status": last["Status"]})
        if not insp.empty and chart["labels"]:
            fb = insp.groupby(insp["Date"].dt.strftime("%b-%y"))["Freeboard_m"].min().to_dict()
            chart["freeboard"] = [fb.get(lbl) for lbl in chart["labels"]]
        return {"kpi": kpi, "insp_rows": insp_rows, "piez_rows": piez_rows, "chart": chart}

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
                    v = round(v, 3)
                row[k] = v
            out.append(row)
        return out

    # ----- status / options ------------------------------------------------
    def data_status(self):
        try:
            from importer import latest_import_runs
            runs = latest_import_runs()
        except Exception:
            runs = {}
        files = []
        for s in self.status:
            key = s["key"]
            spec = C.DATASETS.get(key)
            run = runs.get(key)
            files.append({
                "key": key, "file": spec["file"] if spec else "(in-app)",
                "rows": s["rows"], "loaded": s["loaded"], "error": s.get("error"),
                "rejected": run.rows_rejected if run else 0,
                "import_id": run.id if run else None,
                "profile": run.profile if run else None,
                "modified": run.ts if run else None,
            })
        total_rows = sum(len(v) for v in self.frames.values())
        return {
            "files": files,
            "loaded_at": self.loaded_at.strftime("%d-%b-%Y %H:%M:%S") if self.loaded_at else "-",
            "total_files": sum(1 for f in files if f["loaded"]),
            "total_rows": total_rows,
            "data_dir": self.data_dir,
            "db_path": C.DB_PATH,
            "site_id": C.SITE_ID,
        }

    def excel_files(self):
        out = []
        for key, spec in C.DATASETS.items():
            p = os.path.join(self.data_dir, spec["file"])
            exists = os.path.exists(p)
            out.append({"key": key, "file": spec["file"], "exists": exists,
                        "modified": dt.datetime.fromtimestamp(os.path.getmtime(p)) if exists else None})
        return out

    def filter_options(self):
        inc = self.df("incidents")
        act = self.df("activity")
        years = sorted(inc["Year"].dropna().unique().tolist()) if not inc.empty else []
        depts, areas = set(C.DEPARTMENTS), set(C.AREAS)
        for fr in (inc, act):
            if not fr.empty and "Department" in fr:
                depts |= {d for d in fr["Department"].dropna().astype(str) if d.strip()}
            if not fr.empty and "Area" in fr:
                areas |= {a for a in fr["Area"].dropna().astype(str) if a.strip()}
        return {
            "years": ["All"] + [str(int(y)) for y in years],
            "months": ["All"] + ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                                  "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
            "departments": ["All"] + sorted(depts),
            "areas": ["All"] + sorted(areas),
        }
