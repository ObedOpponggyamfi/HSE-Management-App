#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
HSE Management Web App (Flask).

Consolidates the safety Excel files in data/ and serves a live local dashboard.
Edit/append rows in the Excel files, hit "Refresh" in the app, and every KPI,
chart and table updates.

Run:
    python generate_dummy_data.py     # once, to create sample data/ files
    python app.py                      # -> http://127.0.0.1:5000
"""
import os
import webbrowser
from threading import Timer

from flask import (Flask, flash, redirect, render_template, request, url_for)

import config as C
from core import DataStore

app = Flask(__name__)
app.secret_key = "hse-management-app-local"

# Single in-memory consolidation of the Excel data, shared across requests.
store = DataStore()


# ---------------------------------------------------------------------------
# template number-format filters
# ---------------------------------------------------------------------------
@app.template_filter("comma")
def _comma(v):
    try:
        return f"{int(round(float(v))):,}"
    except (ValueError, TypeError):
        return v


@app.template_filter("pct1")
def _pct1(v):
    try:
        return f"{float(v) * 100:.1f}%"
    except (ValueError, TypeError):
        return v


@app.template_filter("dec2")
def _dec2(v):
    try:
        return f"{float(v):.2f}"
    except (ValueError, TypeError):
        return v


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------
def get_filters():
    return {
        "year": request.args.get("year", "All"),
        "month": request.args.get("month", "All"),
        "dept": request.args.get("dept", "All"),
        "area": request.args.get("area", "All"),
    }


@app.context_processor
def inject_globals():
    return {
        "options": store.filter_options(),
        "filters": get_filters(),
        "status": store.data_status(),
        "nav_active": request.endpoint,
        "company": "ASANKO GOLD MINE",
    }


# ---------------------------------------------------------------------------
# routes
# ---------------------------------------------------------------------------
@app.route("/")
def dashboard():
    f = get_filters()
    charts = {
        "trend": store.monthly_trend(f),
        "type": store.by_type(f),
        "location": store.by_location(f),
        "actions": store.actions_by_status(f),
        "training": store.training_by_dept(f),
    }
    return render_template("dashboard.html", kpis=store.kpis(f), charts=charts,
                           recent=store.recent_incidents(f))


@app.route("/incidents")
def incidents():
    f = get_filters()
    return render_template("incidents.html", rows=store.incidents_table(f))


@app.route("/actions")
def actions():
    f = get_filters()
    return render_template("actions.html", rows=store.actions_table(f),
                           summary=store.actions_by_status(f))


@app.route("/compliance")
def compliance():
    rows, summary = store.compliance_table()
    return render_template("compliance.html", rows=rows, summary=summary)


@app.route("/environmental")
def environmental():
    f = get_filters()
    view = store.environmental_view(f)
    return render_template("environmental.html", view=view, charts={"env": view["chart"]})


@app.route("/contractors")
def contractors():
    f = get_filters()
    view = store.contractors_view(f)
    return render_template("contractors.html", view=view, charts={"contractor": view["chart"]})


@app.route("/registers")
def registers():
    return render_template("registers.html",
                           permits=store.register("permits"),
                           audits=store.register("audits"),
                           equipment=store.register("equipment"))


@app.route("/data")
def data_page():
    return render_template("data.html")


@app.route("/refresh", methods=["POST"])
def refresh():
    store.refresh()
    st = store.data_status()
    flash(f"Refreshed {st['total_files']} files / {st['total_rows']:,} rows from {st['data_dir']}.")
    return redirect(request.form.get("next") or url_for("dashboard"))


PORT = int(os.environ.get("PORT", "5050"))   # 5050 by default to avoid common :5000 clashes


def _open_browser():
    webbrowser.open_new(f"http://127.0.0.1:{PORT}/")


if __name__ == "__main__":
    if not os.path.isdir(C.DATA_DIR) or not store.data_status()["total_files"]:
        print("No data found. Run:  python generate_dummy_data.py")
    use_reloader = os.environ.get("HSE_NO_RELOAD") != "1"
    if (os.environ.get("HSE_NO_BROWSER") != "1"
            and os.environ.get("WERKZEUG_RUN_MAIN") != "true"):
        Timer(1.0, _open_browser).start()      # auto-open the browser once
    print(f" * HSE Management App -> http://127.0.0.1:{PORT}/")
    app.run(host="127.0.0.1", port=PORT, debug=True, use_reloader=use_reloader)
