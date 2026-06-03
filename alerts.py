#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Alert engine: scans the consolidated data for things needing attention and
(optionally) emails a digest.

Alert sources: overdue corrective actions, permits expiring/expired, compliance
items due/overdue, safety-critical equipment inspections due/overdue, and recent
high-severity open incidents.
"""
import datetime as dt
import smtplib
from email.mime.text import MIMEText

import pandas as pd

import config as C

LEVEL_RANK = {"high": 0, "medium": 1, "low": 2}


def compute_alerts(store):
    today = pd.Timestamp(dt.date.today())
    items = []

    def add(level, category, message, link, ref=""):
        items.append({"level": level, "category": category, "message": message,
                      "link": link, "ref": ref})

    # ---- overdue corrective actions ----
    acts = store.df("actions")
    if not acts.empty:
        for r in acts[acts["Overdue"] == 1].itertuples():
            add("high", "Overdue Action",
                f"{r.Action_ID}: {r.Description} — {int(r.DaysOverdue)} days overdue",
                "/actions", r.Action_ID)

    # ---- permits ----
    permits = store.df("permits")
    if not permits.empty:
        for r in permits.itertuples():
            d = int(r.DaysToExpiry)
            if d < 0:
                add("high", "Permit Expired", f"{r.Permit} expired {abs(d)} days ago", "/registers", r.Permit)
            elif d <= C.ALERT_PERMIT_DAYS:
                add("medium", "Permit Expiring", f"{r.Permit} expires in {d} days", "/registers", r.Permit)

    # ---- compliance ----
    comp = store.df("compliance")
    if not comp.empty:
        for r in comp.itertuples():
            d = int(r.DaysToDue)
            if d < 0:
                add("high", "Compliance Overdue", f"{r.Item} ({r.Regulator}) overdue by {abs(d)} days", "/compliance", r.Item)
            elif d <= C.ALERT_COMPLIANCE_DAYS:
                add("medium", "Compliance Due", f"{r.Item} ({r.Regulator}) due in {d} days", "/compliance", r.Item)

    # ---- equipment ----
    equip = store.df("equipment")
    if not equip.empty:
        for r in equip.itertuples():
            d = int(r.DaysToDue)
            if d < 0:
                add("high", "Equipment Overdue", f"{r.Asset} inspection overdue by {abs(d)} days", "/registers", r.Asset)
            elif d <= C.ALERT_EQUIPMENT_DAYS:
                add("medium", "Equipment Due", f"{r.Asset} inspection due in {d} days", "/registers", r.Asset)

    # ---- competency / licence expiry (critical types only) ----
    comp = store.df("competency")
    if not comp.empty:
        crit = comp[comp["Type"].isin(["Statutory Licence", "Medical"])]
        for r in crit.itertuples():
            d = int(r.DaysToExpiry) if pd.notna(r.DaysToExpiry) else -1
            if r.Status == "Expired":
                add("high", "Competency Expired", f"{r.Person}: {r.Competency} has expired", "/training", r.Person)
            elif r.Status == "Expiring":
                add("medium", "Competency Expiring", f"{r.Person}: {r.Competency} expires in {d} days", "/training", r.Person)

    # ---- recent high-severity open incidents ----
    inc = store.df("incidents")
    if not inc.empty:
        recent = inc[(inc["Severity"] >= C.ALERT_HIGH_SEVERITY)
                     & (inc["Status"] != "Closed")
                     & (inc["Date"] >= today - pd.Timedelta(days=C.ALERT_INCIDENT_RECENT_DAYS))]
        for r in recent.itertuples():
            add("high", "High-Severity Incident",
                f"{r.ID}: {r.Type} at {r.Area} (severity {int(r.Severity)})", "/incidents", r.ID)

    items.sort(key=lambda x: LEVEL_RANK.get(x["level"], 9))
    counts = {"high": 0, "medium": 0, "low": 0}
    by_category = {}
    for it in items:
        counts[it["level"]] = counts.get(it["level"], 0) + 1
        by_category.setdefault(it["category"], []).append(it)
    return {"items": items, "counts": counts, "total": len(items), "by_category": by_category}


def digest_text(alerts):
    lines = [f"HSE ALERT DIGEST — {dt.date.today():%d %b %Y}",
             f"{alerts['total']} open alerts "
             f"(high: {alerts['counts']['high']}, medium: {alerts['counts']['medium']})", ""]
    for cat, lst in alerts["by_category"].items():
        lines.append(f"== {cat} ({len(lst)}) ==")
        lines += [f"  - {it['message']}" for it in lst]
        lines.append("")
    return "\n".join(lines)


def email_configured():
    return bool(C.SMTP_HOST and C.ALERT_RECIPIENTS)


def send_digest(alerts):
    """Send the digest by email if SMTP is configured. Returns (ok, message)."""
    if not email_configured():
        return False, "SMTP not configured — set HSE_SMTP_HOST and HSE_ALERT_RECIPIENTS."
    try:
        msg = MIMEText(digest_text(alerts))
        msg["Subject"] = f"HSE Alert Digest — {alerts['total']} open ({alerts['counts']['high']} high)"
        msg["From"] = C.SMTP_FROM
        msg["To"] = ", ".join(C.ALERT_RECIPIENTS)
        with smtplib.SMTP(C.SMTP_HOST, C.SMTP_PORT, timeout=20) as s:
            if C.SMTP_TLS:
                s.starttls()
            if C.SMTP_USER:
                s.login(C.SMTP_USER, C.SMTP_PASS)
            s.sendmail(C.SMTP_FROM, C.ALERT_RECIPIENTS, msg.as_string())
        return True, f"Digest emailed to {len(C.ALERT_RECIPIENTS)} recipient(s)."
    except Exception as exc:
        return False, f"Email failed: {exc}"
