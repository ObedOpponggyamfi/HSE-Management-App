#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
HSE Management Web App (Flask).

Runs local-first from a SQLite canonical store. Spreadsheet refreshes import the
Excel files into that store, then every KPI/chart/table reads from the same
audited data source.
"""
from __future__ import annotations

import logging
import os
import webbrowser
from threading import Lock, Timer
from urllib.parse import urlparse

from flask import (
    Flask,
    current_app,
    flash,
    redirect,
    render_template,
    request,
    url_for,
)

import config as C
from core import DataStore

PORT = int(os.environ.get("PORT", "5050"))


def _store() -> DataStore:
    return current_app.config["DATA_STORE"]


def _safe_next(default_endpoint: str = "dashboard") -> str:
    candidate = request.form.get("next") or request.args.get("next") or url_for(default_endpoint)
    parsed = urlparse(candidate)
    if parsed.scheme or parsed.netloc:
        return url_for(default_endpoint)
    return candidate if candidate.startswith("/") else url_for(default_endpoint)


def get_filters():
    return {
        "year": request.args.get("year", "All"),
        "month": request.args.get("month", "All"),
        "dept": request.args.get("dept", "All"),
        "area": request.args.get("area", "All"),
    }


def register_filters(app: Flask) -> None:
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


def create_app(store: DataStore | None = None) -> Flask:
    app = Flask(__name__)
    app.secret_key = os.environ.get("HSE_SECRET_KEY", "hse-management-app-local-dev")
    app.config["DATA_STORE"] = store or DataStore(
        data_dir=os.environ.get("HSE_DATA_DIR", C.DATA_DIR),
        db_path=os.environ.get("HSE_DB_PATH", os.path.join(C.INSTANCE_DIR, "hse_dashboard.sqlite")),
        site_id=os.environ.get("HSE_SITE_ID", "default"),
    )
    app.config["IMPORT_LOCK"] = Lock()

    logging.basicConfig(
        level=os.environ.get("HSE_LOG_LEVEL", "INFO"),
        format="%(asctime)s %(levelname)s %(name)s %(message)s",
    )
    register_filters(app)

    @app.context_processor
    def inject_globals():
        store = _store()
        return {
            "options": store.filter_options(),
            "filters": get_filters(),
            "status": store.data_status(),
            "nav_active": request.endpoint,
            "company": C.COMPANY_NAME,
            "filter_fields": ("year", "month", "dept", "area"),
        }

    @app.route("/")
    def dashboard():
        store = _store()
        f = get_filters()
        charts = {
            "trend": store.monthly_trend(f),
            "type": store.by_type(f),
            "location": store.by_location(f),
            "actions": store.actions_by_status(f),
            "training": store.training_by_dept(f),
        }
        return render_template(
            "dashboard.html",
            kpis=store.kpis(f),
            charts=charts,
            recent=store.recent_incidents(f),
            filter_fields=("year", "month", "dept", "area"),
        )

    @app.route("/incidents")
    def incidents():
        return render_template(
            "incidents.html",
            rows=_store().incidents_table(get_filters()),
            filter_fields=("year", "month", "dept", "area"),
        )

    @app.route("/actions")
    def actions():
        store = _store()
        f = get_filters()
        return render_template(
            "actions.html",
            rows=store.actions_table(f),
            summary=store.actions_by_status(f),
            filter_fields=("year", "month", "dept", "area"),
        )

    @app.route("/compliance")
    def compliance():
        rows, summary = _store().compliance_table()
        return render_template("compliance.html", rows=rows, summary=summary)

    @app.route("/environmental")
    def environmental():
        f = get_filters()
        view = _store().environmental_view(f)
        return render_template(
            "environmental.html",
            view=view,
            charts={"env": view["chart"]},
            filter_fields=("year", "month"),
        )

    @app.route("/contractors")
    def contractors():
        f = get_filters()
        view = _store().contractors_view(f)
        return render_template(
            "contractors.html",
            view=view,
            charts={"contractor": view["chart"]},
            filter_fields=("year", "month", "dept", "area"),
        )

    @app.route("/registers")
    def registers():
        store = _store()
        return render_template(
            "registers.html",
            permits=store.register("permits"),
            audits=store.register("audits"),
            equipment=store.register("equipment"),
        )

    @app.route("/data")
    def data_page():
        return render_template("data.html")

    @app.route("/refresh", methods=["POST"])
    def refresh():
        lock: Lock = current_app.config["IMPORT_LOCK"]
        if not lock.acquire(blocking=False):
            flash("A refresh is already running. Please try again in a moment.")
            return redirect(_safe_next())
        try:
            current_app.logger.info("refresh.start site=%s", _store().site_id)
            _store().refresh()
            st = _store().data_status()
            flash(
                f"Refreshed {st['total_files']} sources / {st['total_rows']:,} rows "
                f"for site {st['site_id']}."
            )
            current_app.logger.info("refresh.done site=%s rows=%s", st["site_id"], st["total_rows"])
        finally:
            lock.release()
        return redirect(_safe_next())

    return app


def _open_browser():
    webbrowser.open_new(f"http://127.0.0.1:{PORT}/")


app = None if os.environ.get("HSE_SKIP_GLOBAL_APP") == "1" else create_app()


if __name__ == "__main__":
    run_app = app or create_app()
    store = run_app.config["DATA_STORE"]
    if not os.path.isdir(store.data_dir) or not store.data_status()["total_files"]:
        print("No data found. Run:  python generate_dummy_data.py")
    use_reloader = os.environ.get("HSE_NO_RELOAD") != "1"
    debug = os.environ.get("HSE_DEBUG", "0") == "1"
    if (os.environ.get("HSE_NO_BROWSER") != "1"
            and os.environ.get("WERKZEUG_RUN_MAIN") != "true"):
        Timer(1.0, _open_browser).start()
    print(f" * HSE Management App -> http://127.0.0.1:{PORT}/")
    run_app.run(host="127.0.0.1", port=PORT, debug=debug, use_reloader=use_reloader)
