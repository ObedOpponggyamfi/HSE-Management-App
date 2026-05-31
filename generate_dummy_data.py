#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Generate DUMMY safety datasets as plain Excel files in the data/ folder.

These simulate the safety spreadsheets that would normally come from different
departments / sources. The Flask app consolidates them on startup and on every
Refresh. Real values & dates only -- NO formulas (the app does the maths).

Run:  python generate_dummy_data.py
"""
import calendar
import datetime as dt
import os

import numpy as np
import pandas as pd

import config as C


def _month_timeline(anchor: dt.date, n: int):
    out, y, m = [], anchor.year, anchor.month
    for _ in range(n):
        out.append({"year": y, "month": m, "first": dt.date(y, m, 1)})
        m -= 1
        if m == 0:
            m, y = 12, y - 1
    return list(reversed(out))


def generate(anchor: dt.date | None = None) -> dict:
    """Build every dummy DataFrame (reproducible via fixed seed)."""
    anchor = anchor or dt.date.today()
    rng = np.random.default_rng(C.RNG_SEED)
    months = _month_timeline(anchor, C.MONTHS_OF_HISTORY)
    today = anchor

    # ---- Incidents ----------------------------------------------------------
    # Recordables (MTC/RWC/LTI) are a small slice of all reported events so the
    # frequency rates land in a realistic band near target.
    type_w = np.array([0.045, 0.42, 0.025, 0.02, 0.16, 0.15, 0.18])  # see INCIDENT_TYPES
    type_w /= type_w.sum()
    rows, iid = [], 1
    for k, mo in enumerate(months):
        lam = 9.0 - 4.0 * (k / max(1, len(months) - 1))           # gently improving
        ndays = calendar.monthrange(mo["year"], mo["month"])[1]
        for _ in range(int(rng.poisson(max(2.0, lam)))):
            area = rng.choice(C.AREAS)
            itype = rng.choice(C.INCIDENT_TYPES, p=type_w)
            if itype == "Lost Time Injury":
                sev = int(rng.choice([3, 4, 5], p=[0.5, 0.35, 0.15]))
            elif itype in ("Restricted Work", "Property Damage", "Environmental"):
                sev = int(rng.choice([2, 3, 4], p=[0.4, 0.45, 0.15]))
            else:
                sev = int(rng.choice([1, 2, 3], p=[0.55, 0.35, 0.10]))
            date = dt.date(mo["year"], mo["month"], int(rng.integers(1, ndays + 1)))
            reported = date + dt.timedelta(days=int(rng.integers(0, 3)))
            age = (today - date).days
            if age > 150:
                status = rng.choice(C.INCIDENT_STATUS, p=[0.02, 0.03, 0.05, 0.90])
            elif age > 60:
                status = rng.choice(C.INCIDENT_STATUS, p=[0.08, 0.12, 0.20, 0.60])
            else:
                status = rng.choice(C.INCIDENT_STATUS, p=[0.30, 0.30, 0.20, 0.20])
            needs_car = itype in C.RECORDABLE_TYPES or itype in ("Environmental", "Property Damage")
            car_due = reported + dt.timedelta(days=int(rng.integers(14, 61))) if needs_car else None
            rows.append({
                "ID": f"INC-{iid:04d}", "Date": date, "Area": area,
                "Department": C.AREA_DEPT[area], "Company": C.AREA_COMPANY[area],
                "Type": itype, "Class": rng.choice(C.INCIDENT_CLASSES), "Severity": sev,
                "Status": status, "Reported": reported, "CAR_Due": car_due,
                "Owner": rng.choice(C.OWNERS),
            })
            iid += 1
    incidents = pd.DataFrame(rows)

    # ---- Activity (per area, per month) ------------------------------------
    base_head = {
        "Nkran Open Pit": 120, "Esaase Open Pit": 110, "Drill & Blast": 40,
        "Haul Roads": 60, "Explosives Magazine": 15, "Crushing Circuit": 35,
        "Processing Plant (CIL)": 90, "Elution & Gold Room": 20,
        "Tailings Storage Facility": 18, "Water Treatment": 12, "Assay Laboratory": 25,
        "HME Workshop": 55, "Fuel Farm": 10, "Power Station": 22,
        "Warehouse & Stores": 24, "Administration": 60, "Accommodation Camp": 30,
        "Security Gatehouse": 45,
    }
    rows = []
    for mo in months:
        for area in C.AREAS:
            hc = base_head[area]
            obs_r = int(rng.poisson(max(2, hc / 12)))
            insp_p = int(max(2, hc / 18))
            tr_assigned = int(max(3, hc / 6))
            rows.append({
                "Period": mo["first"], "Area": area, "Department": C.AREA_DEPT[area],
                "Company": C.AREA_COMPANY[area],
                "ManHours": int(hc * rng.integers(180, 230)),
                "NearMisses": int(rng.poisson(max(1, hc / 22))),
                "Hazards": int(rng.poisson(max(1, hc / 30))),
                "ObsRaised": obs_r, "ObsClosed": int(round(obs_r * rng.uniform(0.75, 0.98))),
                "InspPlanned": insp_p, "InspDone": int(round(insp_p * rng.uniform(0.85, 1.0))),
                "InspScore": round(float(rng.uniform(0.88, 0.995)), 3),
                "Audits": int(rng.integers(0, 2)),
                "Toolbox": int(rng.poisson(max(2, hc / 10))),
                "TrainAssigned": tr_assigned,
                "TrainCompleted": int(round(tr_assigned * rng.uniform(0.9, 1.0))),
                "PPE": round(float(rng.uniform(0.9, 1.0)), 3),
            })
    activity = pd.DataFrame(rows)

    # ---- Corrective actions -------------------------------------------------
    rows, aid = [], 1
    for _, r in incidents[incidents["CAR_Due"].notna()].iterrows():
        closed = r["Status"] == "Closed"
        status = "Closed" if closed else rng.choice(["Open", "In Progress", "Due Soon"],
                                                    p=[0.45, 0.35, 0.20])
        rows.append({
            "Action_ID": f"CAR-{aid:04d}", "Source_Incident": r["ID"],
            "Description": f"Corrective action for {r['Type']} at {r['Area']}",
            "Raised": r["Reported"], "Due": r["CAR_Due"], "Owner": r["Owner"],
            "Department": r["Department"], "Area": r["Area"],
            "Priority": rng.choice(C.PRIORITY, p=[0.2, 0.4, 0.3, 0.1]), "Status": status,
        })
        aid += 1
    for _ in range(28):
        area = rng.choice(C.AREAS)
        raised = today - dt.timedelta(days=int(rng.integers(5, 200)))
        rows.append({
            "Action_ID": f"CAR-{aid:04d}", "Source_Incident": "Inspection/Audit",
            "Description": rng.choice([
                "Install machine guard", "Repair edge protection", "Replace fire extinguisher",
                "Update SOP / JSA", "Bund repair at fuel storage", "Improve signage",
                "Spill kit replenishment", "Lighting upgrade", "Housekeeping campaign"]),
            "Raised": raised, "Due": raised + dt.timedelta(days=int(rng.integers(20, 75))),
            "Owner": rng.choice(C.OWNERS), "Department": C.AREA_DEPT[area], "Area": area,
            "Priority": rng.choice(C.PRIORITY, p=[0.3, 0.4, 0.2, 0.1]),
            "Status": rng.choice(C.ACTION_STATUS, p=[0.20, 0.25, 0.15, 0.40]),
        })
        aid += 1
    actions = pd.DataFrame(rows)

    # ---- Compliance (Ghana regulators) -------------------------------------
    comp_defs = [
        ("Mining Lease / Operating Permit", "Minerals Commission", "Act 703", 12),
        ("Annual Mineral Right Rent", "Minerals Commission", "LI 2176", 12),
        ("Environmental Permit", "EPA Ghana", "LI 1652", 12),
        ("Annual Environmental Mgmt Report (AEMR)", "EPA Ghana", "EA Regs", 12),
        ("Effluent / Discharge Monitoring", "EPA Ghana", "GEPA Std", 3),
        ("Factory Registration & Inspection", "Factories Inspectorate", "Act 328", 12),
        ("Pressure Vessel Certification", "Factories Inspectorate", "Act 328", 12),
        ("ICMC Cyanide Code Certification", "ICMI", "Cyanide Code", 36),
        ("Cyanide Transport Audit", "ICMI", "Cyanide Code", 12),
        ("Radiation Source Licence", "Nuclear Regulatory Authority", "Act 895", 12),
        ("Water Use Permit", "Water Resources Commission", "Act 522", 12),
        ("Fire Safety Certificate", "Ghana National Fire Service", "Act 537", 12),
        ("Explosives Licence (Magazine)", "Minerals Commission", "LI 2177", 12),
    ]
    rows = []
    for item, reg, ref, freq in comp_defs:
        last = today - dt.timedelta(days=int(rng.integers(20, int(freq * 30 + 40))))
        rows.append({"Item": item, "Regulator": reg, "Reference": ref,
                     "Frequency_Months": freq, "Last_Completed": last,
                     "Owner": rng.choice(C.OWNERS)})
    compliance = pd.DataFrame(rows)

    # ---- Environmental monitoring ------------------------------------------
    rows = []
    for k, mo in enumerate(months):
        pm10 = round(float(rng.normal(48, 12)), 1)
        if k in (6, 17):
            pm10 = round(float(rng.uniform(74, 92)), 1)
        ph = round(float(rng.normal(7.6, 0.5)), 2)
        if k == 11:
            ph = round(float(rng.uniform(9.1, 9.6)), 2)
        cn = round(float(rng.normal(28, 9)), 1)
        if k == 20:
            cn = round(float(rng.uniform(52, 66)), 1)
        rows.append({
            "Period": mo["first"], "Waste_t": round(float(rng.uniform(120, 260)), 1),
            "Recycling": round(float(rng.uniform(0.35, 0.62)), 3),
            "Energy_MWh": round(float(rng.uniform(5200, 6800)), 0),
            "Water_m3": round(float(rng.uniform(38000, 52000)), 0),
            "Fuel_L": round(float(rng.uniform(420000, 560000)), 0),
            "PM10": pm10, "pH": ph, "WAD_CN": cn,
        })
    environmental = pd.DataFrame(rows)

    # ---- Supporting registers ----------------------------------------------
    permits = pd.DataFrame([{
        "Permit": p, "Authority": auth, "Holder": rng.choice(C.OWNERS),
        "Issue_Date": today - dt.timedelta(days=int(rng.integers(120, 700))),
        "Expiry_Date": today + dt.timedelta(days=int(d)),
    } for p, auth, d in [
        ("Hot Work Permit - Plant", "HSE Dept", 45),
        ("Confined Space - Elution", "HSE Dept", 120),
        ("Working at Height - Crusher", "HSE Dept", -10),
        ("Excavation Permit - Esaase", "HSE Dept", 200),
        ("Explosives Handling", "Minerals Commission", 300),
        ("Radiation Source Permit", "Nuclear Reg. Authority", 30),
        ("Electrical Isolation - HV", "Engineering", 90),
        ("Lifting Operations Permit", "Engineering", 150)]])

    audits = pd.DataFrame([{
        "Audit": a, "Type": t, "Date": today - dt.timedelta(days=int(rng.integers(10, 330))),
        "Auditor": rng.choice(C.OWNERS), "Score": round(float(rng.uniform(0.72, 0.97)), 3),
        "Findings": f, "Closed_Findings": int(rng.integers(0, f + 1)),
    } for a, t, f in [
        ("ICMC Cyanide Code Audit", "External", 8), ("ISO 45001 Surveillance", "External", 6),
        ("EPA Environmental Audit", "Regulatory", 5), ("Internal HSE System Audit", "Internal", 12),
        ("Contractor HSE Audit - AUMS", "Internal", 9), ("Tailings Facility Audit", "External", 4),
        ("Emergency Preparedness Audit", "Internal", 7), ("Explosives Magazine Audit", "Regulatory", 3)]])

    equipment = pd.DataFrame([{
        "Asset_ID": f"EQP-{i:03d}", "Asset": a, "Type": t, "Location": rng.choice(C.AREAS),
        "Last_Inspection": today - dt.timedelta(days=int(rng.integers(20, 200))),
        "Next_Inspection": today + dt.timedelta(days=int(d)),
    } for i, (a, t, d) in enumerate([
        ("Fire Pump House 1", "Fire", 25), ("Foam System - Fuel Farm", "Fire", -5),
        ("SCBA Set A", "Emergency", 40), ("Overhead Crane - Workshop", "Lifting", 90),
        ("Mobile Crane 50t", "Lifting", 15), ("Gas Detection Network", "Monitoring", 60),
        ("Eyewash Stations (CIL)", "Emergency", -2), ("Emergency Generator", "Power", 120),
        ("Ambulance 1", "Medical", 30), ("Tailings Piezometers", "Geotech", 75),
        ("AED Units", "Medical", 50), ("Spill Response Trailer", "Environmental", 100)], start=1)])

    return {
        "incidents": incidents, "activity": activity, "actions": actions,
        "compliance": compliance, "environmental": environmental,
        "permits": permits, "audits": audits, "equipment": equipment,
    }


def write_files(data: dict):
    os.makedirs(C.DATA_DIR, exist_ok=True)
    written = []
    for key, spec in C.DATASETS.items():
        if key not in data:
            continue
        path = os.path.join(C.DATA_DIR, spec["file"])
        data[key].to_excel(path, sheet_name=spec["sheet"], index=False)
        written.append((spec["file"], len(data[key])))
    return written


def main():
    print("Generating dummy HSE datasets ...")
    data = generate()
    written = write_files(data)
    print(f"Wrote {len(written)} Excel files to: {C.DATA_DIR}")
    for fname, n in written:
        print(f"   - {fname:28s} {n:4d} rows")
    print("Done. Start the app with:  python app.py")


if __name__ == "__main__":
    main()
