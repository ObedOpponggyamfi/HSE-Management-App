#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
HSE Management Web App (Flask, database-backed).

A local HSE management system: consolidates safety data, lets users log
incidents / near misses / hazards / actions, enforces login + roles, keeps an
audit trail, raises alerts, and exports reports.

Run:
    python app.py                      # -> http://127.0.0.1:5050
First run auto-creates the SQLite DB, imports data/*.xlsx and seeds users.
"""
import datetime as dt
import os
import webbrowser
from functools import wraps
from threading import Timer

from flask import (Flask, abort, flash, redirect, render_template, request,
                   send_file, url_for)
from flask_login import (current_user, login_required, login_user, logout_user)

import alerts as alerts_mod
import config as C
import reports as reports_mod
from core import DataStore
from extensions import db, login_manager
from forms import (ActionForm, EventForm, IncidentForm, InvestigationForm,
                   LoginForm, UserForm)
from importer import (insert_action, insert_incident, log_audit, next_ref,
                      seed_database, update_action_status)

app = Flask(__name__)
app.config.update(
    SECRET_KEY=C.SECRET_KEY,
    SQLALCHEMY_DATABASE_URI=C.SQLALCHEMY_DATABASE_URI,
    SQLALCHEMY_TRACK_MODIFICATIONS=False,
)
db.init_app(app)
login_manager.init_app(app)

# Models must be imported so their tables are registered before create_all().
from models import AuditLog, Event, Investigation, User  # noqa: E402

# First-run bootstrap + initial consolidation (inside an app context).
with app.app_context():
    seed_database()
    store = DataStore()


# ---------------------------------------------------------------------------
# template filters + globals
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


def get_filters():
    return {
        "year": request.args.get("year", "All"),
        "month": request.args.get("month", "All"),
        "dept": request.args.get("dept", "All"),
        "area": request.args.get("area", "All"),
    }


@app.context_processor
def inject_globals():
    try:
        alert_counts = alerts_mod.compute_alerts(store)["counts"]
        alert_total = sum(alert_counts.values())
    except Exception:
        alert_counts, alert_total = {"high": 0, "medium": 0, "low": 0}, 0
    return {
        "options": store.filter_options(),
        "filters": get_filters(),
        "status": store.data_status(),
        "nav_active": request.endpoint,
        "company": "ASANKO GOLD MINE",
        "alert_total": alert_total,
        "alert_counts": alert_counts,
    }


def role_required(role):
    """Allow only users whose role rank >= the given role."""
    def deco(fn):
        @wraps(fn)
        @login_required
        def wrapper(*a, **k):
            if not current_user.can(role):
                flash("You don't have permission to do that.", "warn")
                return redirect(url_for("dashboard"))
            return fn(*a, **k)
        return wrapper
    return deco


# ---------------------------------------------------------------------------
# auth
# ---------------------------------------------------------------------------
@app.route("/login", methods=["GET", "POST"])
def login():
    if current_user.is_authenticated:
        return redirect(url_for("dashboard"))
    form = LoginForm()
    if form.validate_on_submit():
        user = User.query.filter_by(username=form.username.data.strip()).first()
        if user and user.active and user.check_password(form.password.data):
            login_user(user)
            log_audit(user.username, "login", "auth")
            return redirect(request.args.get("next") or url_for("dashboard"))
        flash("Invalid username or password.", "bad")
    return render_template("login.html", form=form)


@app.route("/logout")
@login_required
def logout():
    log_audit(current_user.username, "logout", "auth")
    logout_user()
    flash("You have been signed out.")
    return redirect(url_for("login"))


# ---------------------------------------------------------------------------
# dashboards / views (all require login)
# ---------------------------------------------------------------------------
@app.route("/")
@login_required
def dashboard():
    f = get_filters()
    charts = {
        "trend": store.monthly_trend(f), "type": store.by_type(f),
        "location": store.by_location(f), "actions": store.actions_by_status(f),
        "training": store.training_by_dept(f),
    }
    return render_template("dashboard.html", kpis=store.kpis(f), charts=charts,
                           recent=store.recent_incidents(f))


@app.route("/incidents")
@login_required
def incidents():
    return render_template("incidents.html", rows=store.incidents_table(get_filters()))


@app.route("/actions")
@login_required
def actions():
    f = get_filters()
    return render_template("actions.html", rows=store.actions_table(f),
                           summary=store.actions_by_status(f))


@app.route("/events")
@login_required
def events():
    f = get_filters()
    view = store.events_view(f)
    return render_template("events.html", view=view, charts={"events": view["chart"]})


@app.route("/compliance")
@login_required
def compliance():
    rows, summary = store.compliance_table()
    return render_template("compliance.html", rows=rows, summary=summary)


@app.route("/environmental")
@login_required
def environmental():
    f = get_filters()
    view = store.environmental_view(f)
    return render_template("environmental.html", view=view, charts={"env": view["chart"]})


@app.route("/contractors")
@login_required
def contractors():
    f = get_filters()
    view = store.contractors_view(f)
    return render_template("contractors.html", view=view, charts={"contractor": view["chart"]})


@app.route("/rates")
@login_required
def rates():
    f = get_filters()
    roll = store.rolling_rates(f)
    return render_template("rates.html", roll=roll, charts={"rates": roll})


@app.route("/training")
@login_required
def training():
    f = get_filters()
    view = store.competency_view(f)
    return render_template("training.html", view=view, charts={"comp": view["chart"]})


@app.route("/investigations")
@login_required
def investigations():
    items = Investigation.query.order_by(Investigation.id.desc()).all()
    stats = {"total": len(items),
             "open": sum(1 for i in items if i.status != "Completed"),
             "hipo": sum(1 for i in items if i.hipo),
             "completed": sum(1 for i in items if i.status == "Completed")}
    return render_template("investigations.html", items=items, stats=stats)


@app.route("/investigations/new", methods=["GET", "POST"])
@role_required("hse_officer")
def investigation_new():
    form = InvestigationForm()
    inc = store.df("incidents")
    if not inc.empty:
        ids = list(inc.sort_values("Date", ascending=False)["ID"].astype(str).head(300))
        form.incident_id.choices = [(x, x) for x in ids]
    else:
        form.incident_id.choices = [("", "(no incidents)")]
    if not form.is_submitted():
        form.investigator.data = current_user.name
    if form.validate_on_submit():
        ref = next_ref("investigations", "ref", "INV")
        db.session.add(Investigation(
            ref=ref, incident_id=form.incident_id.data, hipo=(form.hipo.data == "Yes"),
            method=form.method.data, immediate_cause=form.immediate_cause.data,
            root_cause=form.root_cause.data, why1=form.why1.data, why2=form.why2.data,
            why3=form.why3.data, why4=form.why4.data, why5=form.why5.data,
            status=form.status.data, investigator=form.investigator.data or current_user.name,
            created_by=current_user.username))
        db.session.commit()
        log_audit(current_user.username, "create", "investigation", ref, form.method.data)
        flash(f"Investigation {ref} saved.")
        return redirect(url_for("investigations"))
    return render_template("form_page.html", form=form, title="Incident Investigation (RCA)",
                           subtitle="Classify HiPo, capture 5-Whys / ICAM root cause, link to the incident.")


@app.route("/registers")
@login_required
def registers():
    return render_template("registers.html", permits=store.register("permits"),
                           audits=store.register("audits"), equipment=store.register("equipment"))


@app.route("/alerts")
@login_required
def alerts_page():
    return render_template("alerts.html", alerts=alerts_mod.compute_alerts(store),
                           email_ready=alerts_mod.email_configured())


@app.route("/report")
@login_required
def report():
    f = get_filters()
    rows, comp_summary = store.compliance_table()
    charts = {"trend": store.monthly_trend(f), "type": store.by_type(f),
              "location": store.by_location(f)}
    return render_template("report.html", kpis=store.kpis(f), charts=charts,
                           comp=comp_summary, env=store.environmental_view(f),
                           generated=dt.datetime.now().strftime("%d-%b-%Y %H:%M"))


@app.route("/data")
@login_required
def data_page():
    return render_template("data.html", excel_files=store.excel_files())


# ---------------------------------------------------------------------------
# data capture (write) — role-gated
# ---------------------------------------------------------------------------
@app.route("/incidents/new", methods=["GET", "POST"])
@role_required("supervisor")
def incident_new():
    form = IncidentForm()
    if not form.is_submitted():
        form.date.data = dt.date.today()
        form.reported.data = dt.date.today()
        form.owner.data = current_user.name
    if form.validate_on_submit():
        ref = next_ref("incidents", "ID", "INC")
        insert_incident({
            "ID": ref, "Date": form.date.data, "Area": form.area.data,
            "Department": form.department.data, "Company": form.company.data,
            "Type": form.type.data, "Class": form.klass.data,
            "Severity": int(form.severity.data), "Status": form.status.data,
            "Reported": form.reported.data or form.date.data, "CAR_Due": form.car_due.data,
            "Owner": form.owner.data})
        log_audit(current_user.username, "create", "incident", ref, form.type.data)
        store.refresh()
        flash(f"Incident {ref} logged.")
        return redirect(url_for("incidents"))
    return render_template("form_page.html", form=form, title="Log an Incident",
                           subtitle="Recordable & lost-time injuries flow into TRIFR/LTIFR automatically.")


@app.route("/events/new", methods=["GET", "POST"])
@login_required
def event_new():
    form = EventForm()
    if not form.is_submitted():
        form.date.data = dt.date.today()
    if form.validate_on_submit():
        ref = next_ref("events", "ref", "EVT")
        db.session.add(Event(
            ref=ref, category=form.category.data, date=form.date.data, area=form.area.data,
            department=C.AREA_DEPT.get(form.area.data, ""), severity=int(form.severity.data),
            description=form.description.data, reported_by=current_user.name, status="Open"))
        db.session.commit()
        log_audit(current_user.username, "create", "event", ref, form.category.data)
        store.refresh()
        flash(f"Report {ref} submitted. Thank you for reporting.")
        return redirect(url_for("events"))
    return render_template("form_page.html", form=form, title="Report a Near Miss / Hazard",
                           subtitle="Every report helps prevent the next incident.")


@app.route("/actions/new", methods=["GET", "POST"])
@role_required("hse_officer")
def action_new():
    form = ActionForm()
    if form.validate_on_submit():
        ref = next_ref("actions", "Action_ID", "CAR")
        insert_action({
            "Action_ID": ref, "Source_Incident": form.source.data or "Manual",
            "Description": form.description.data, "Raised": dt.date.today(),
            "Due": form.due.data, "Owner": form.owner.data,
            "Department": form.department.data, "Area": form.area.data,
            "Priority": form.priority.data, "Status": "Open"})
        log_audit(current_user.username, "create", "action", ref)
        store.refresh()
        flash(f"Action {ref} created.")
        return redirect(url_for("actions"))
    return render_template("form_page.html", form=form, title="Create Corrective Action",
                           subtitle="Assign an owner and due date; overdue items are flagged automatically.")


@app.route("/actions/update", methods=["POST"])
@role_required("hse_officer")
def action_update():
    action_id = request.form.get("action_id", "")
    status = request.form.get("status", "")
    if action_id and status in C.ACTION_STATUS:
        update_action_status(action_id, status)
        log_audit(current_user.username, "update", "action", action_id, f"status={status}")
        store.refresh()
        flash(f"Action {action_id} set to {status}.")
    return redirect(request.form.get("next") or url_for("actions"))


# ---------------------------------------------------------------------------
# alerts email, exports, import, refresh
# ---------------------------------------------------------------------------
@app.route("/alerts/email", methods=["POST"])
@role_required("hse_officer")
def alerts_email():
    ok, msg = alerts_mod.send_digest(alerts_mod.compute_alerts(store))
    flash(msg, "good" if ok else "warn")
    return redirect(url_for("alerts_page"))


@app.route("/export/<key>.xlsx")
@login_required
def export_table(key):
    if key == "all":
        bio = reports_mod.export_workbook(store.frames)
        name = "hse_all_data.xlsx"
    else:
        if key not in store.frames:
            abort(404)
        bio = reports_mod.export_dataframe(store.df(key), sheet=key)
        name = f"hse_{key}.xlsx"
    return send_file(bio, as_attachment=True, download_name=name,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/import-excel", methods=["POST"])
@role_required("admin")
def import_excel():
    result = seed_database(force_import=True)
    log_audit(current_user.username, "import", "excel", detail=str(result.get("imported", {})))
    store.refresh()
    flash("Re-imported all Excel files from the data/ folder into the database.")
    return redirect(url_for("data_page"))


@app.route("/refresh", methods=["POST"])
@login_required
def refresh():
    store.refresh()
    st = store.data_status()
    flash(f"Reloaded {st['total_rows']:,} rows from the database.")
    return redirect(request.form.get("next") or url_for("dashboard"))


# ---------------------------------------------------------------------------
# admin
# ---------------------------------------------------------------------------
@app.route("/admin/users", methods=["GET", "POST"])
@role_required("admin")
def admin_users():
    form = UserForm()
    if form.validate_on_submit():
        if User.query.filter_by(username=form.username.data.strip()).first():
            flash("That username already exists.", "warn")
        else:
            u = User(username=form.username.data.strip(), name=form.name.data,
                     email=form.email.data, role=form.role.data, active=True)
            u.set_password(form.password.data)
            db.session.add(u)
            db.session.commit()
            log_audit(current_user.username, "create", "user", u.username, u.role)
            flash(f"User {u.username} created ({u.role_label}).")
        return redirect(url_for("admin_users"))
    return render_template("admin_users.html", form=form,
                           users=User.query.order_by(User.username).all())


@app.route("/admin/audit")
@role_required("manager")
def admin_audit():
    logs = AuditLog.query.order_by(AuditLog.ts.desc()).limit(300).all()
    return render_template("admin_audit.html", logs=logs)


# ---------------------------------------------------------------------------
# entry point
# ---------------------------------------------------------------------------
PORT = int(os.environ.get("PORT", "5050"))


def _open_browser():
    webbrowser.open_new(f"http://127.0.0.1:{PORT}/")


if __name__ == "__main__":
    use_reloader = os.environ.get("HSE_NO_RELOAD") != "1"
    if (os.environ.get("HSE_NO_BROWSER") != "1"
            and os.environ.get("WERKZEUG_RUN_MAIN") != "true"):
        Timer(1.0, _open_browser).start()
    print(f" * HSE Management App -> http://127.0.0.1:{PORT}/   (login: admin / admin123)")
    app.run(host="127.0.0.1", port=PORT, debug=True, use_reloader=use_reloader)
