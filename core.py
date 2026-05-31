#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
core.py -- the consolidation + analytics engine.

DataStore reads every safety Excel file from the data/ folder, derives the
fields the dashboard needs (Year/Month, Recordable/LTI flags, overdue/compliance
status ...), keeps them in memory, and recomputes all KPIs / chart series on
demand. `refresh()` simply re-reads the folder, so dropping new rows into the
Excel files and hitting Refresh in the app surfaces them immediately.
"""
import datetime as dt
import os

import pandas as pd

import config as C

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
    """Loads & consolidates the safety spreadsheets; computes all analytics."""

    def __init__(self, data_dir: str = C.DATA_DIR):
        self.data_dir = data_dir
        self.frames: dict[str, pd.DataFrame] = {}
        self.status: list[dict] = []
        self.loaded_at: dt.datetime | None = None
        self.load()

    # ----- loading / refreshing -------------------------------------------
    def load(self):
        global TODAY
        TODAY = _today()
        self.frames, self.status = {}, []
        for key, spec in C.DATASETS.items():
            path = os.path.join(self.data_dir, spec["file"])
            entry = {"key": key, "file": spec["file"], "rows": 0,
                     "loaded": False, "modified": None, "error": None}
            try:
                if os.path.exists(path):
                    df = pd.read_excel(path, sheet_name=spec["sheet"])
                    df = self._derive(key, df)
                    self.frames[key] = df
                    entry.update(loaded=True, rows=len(df),
                                 modified=dt.datetime.fromtimestamp(os.path.getmtime(path)))
                else:
                    entry["error"] = "file not found"
                    self.frames[key] = pd.DataFrame()
            except Exception as exc:                       # keep app alive on bad file
                entry["error"] = str(exc)
                self.frames[key] = pd.DataFrame()
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
            df["Overdue"] = ((df["Status"] != "Closed") & (df["Due"] < TODAY)).astype(int)
            df["DaysOverdue"] = ((TODAY - df["Due"]).dt.days).where(df["Overdue"] == 1, 0)

        elif key == "compliance":
            df["Last_Completed"] = pd.to_datetime(df["Last_Completed"])
            df["Due_Date"] = df.apply(
                lambda r: r["Last_Completed"] + pd.DateOffset(months=int(r["Frequency_Months"])),
                axis=1)
            df["DaysToDue"] = (df["Due_Date"] - TODAY).dt.days
            df["Status"] = df["DaysToDue"].apply(
                lambda d: "Overdue" if d < 0 else ("Due Soon" if d <= 30 else "Compliant"))

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

    # ----- timeline helper -------------------------------------------------
    def _timeline(self):
        """Ordered list of (MonthKey, label) across the activity history."""
        act = self.df("activity")
        if act.empty:
            return []
        keys = sorted(act["MonthKey"].unique())
        labels = {k: pd.to_datetime(k + "-01").strftime("%b-%y") for k in keys}
        return [(k, labels[k]) for k in keys]

    # ----- KPIs ------------------------------------------------------------
    def kpis(self, f):
        inc = self._filter(self.df("incidents"), f)
        act = self._filter(self.df("activity"), f)
        acts = self._filter(self.df("actions"), f, year=False, month=False)

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
        """Respect dept/area filters but keep the full month timeline."""
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

    # ----- table builders (return list-of-dicts for templates) -------------
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
            "days_overdue": int(r.DaysOverdue),
        } for r in acts.itertuples()]

    def compliance_table(self):
        df = self.df("compliance")
        if df.empty:
            return [], {"pct": 0, "counts": {}}
        rows = [{
            "item": r.Item, "regulator": r.Regulator, "reference": r.Reference,
            "frequency": int(r.Frequency_Months),
            "last": r.Last_Completed.strftime("%d-%b-%Y"),
            "due": r.Due_Date.strftime("%d-%b-%Y"),
            "days": int(r.DaysToDue), "owner": r.Owner, "status": r.Status,
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
                "ltifr": round(lti / mh * C.RATE_BASE, 2) if mh else 0,
            })
        rows = [r for r in rows if r["manhours"] > 0] or rows
        chart = {"labels": [r["company"].split(" (")[0] for r in rows],
                 "trifr": [r["trifr"] for r in rows], "target": C.TARGET_TRIFR}
        return {"rows": rows, "chart": chart}

    def register(self, key):
        """Generic table for permits / audits / equipment."""
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

    # ----- data status -----------------------------------------------------
    def data_status(self):
        total_rows = sum(len(v) for v in self.frames.values())
        return {
            "files": self.status,
            "loaded_at": self.loaded_at.strftime("%d-%b-%Y %H:%M:%S") if self.loaded_at else "-",
            "total_files": sum(1 for s in self.status if s["loaded"]),
            "total_rows": total_rows,
            "data_dir": self.data_dir,
        }

    # ----- filter option lists --------------------------------------------
    def filter_options(self):
        inc = self.df("incidents")
        years = sorted(inc["Year"].unique().tolist()) if not inc.empty else []
        return {
            "years": ["All"] + [str(y) for y in years],
            "months": ["All"] + ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                                  "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
            "departments": ["All"] + C.DEPARTMENTS,
            "areas": ["All"] + C.AREAS,
        }
