# HSE Management App (Gold Mine)

A local **Python web application** for Health, Safety & Environment management,
tailored for an Asanko-style gold mine in Ghana. It consolidates safety data into
a database, lets your team **capture incidents / near misses / hazards / actions**
in the browser, enforces **login + roles**, keeps an **audit trail**, raises
**alerts**, and produces **reports** — all running on your PC, no cloud required.
It also includes industry-standard modules: rolling 12-month frequency rates, a
training & competency matrix with licence expiry, incident root-cause
investigations, and GISTM tailings monitoring.

---

## Quick start

```bash
pip install -r requirements.txt
python generate_dummy_data.py     # creates sample Excel files in data/ (first run only)
python app.py                      # auto-creates the DB, opens your browser
```

Serves at **http://127.0.0.1:5050/**. **Windows:** just double-click **`run_app.bat`**.

**Sign in** with one of the seeded demo accounts (change these in real use):

| Username | Password | Role |
|---|---|---|
| `admin` | `admin123` | Administrator |
| `hse` | `hse123` | HSE Officer |
| `worker` | `worker123` | Worker |

*Port:* set `PORT=8080` to change it (5050 default avoids the common :5000 clash).

---

## What it does

### 1. Data capture (digitalisation)
- Log **incidents** (flow straight into TRIFR/LTIFR), report **near misses / hazards /
  observations**, and raise **corrective actions** — all from in-app forms.
- New records write to the database and appear on the dashboard immediately.
- Bulk-load is still supported: drop Excel files in `data/` and an admin clicks
  **Import from Excel**. Any dataset can be **exported back to Excel**.

### 2. Login, roles & audit trail
- Authentication with hashed passwords (Flask-Login).
- Role hierarchy **Worker → Supervisor → HSE Officer → Manager → Admin**; each
  action is gated (e.g. workers can report events, HSE officers manage actions,
  admins manage users).
- Every create/update/login is written to an **Audit Trail** (who, what, when).

### 3. Alerts & escalation
- The **Alerts** page (and the bell badge) surface overdue actions, expiring/expired
  permits, compliance due/overdue, equipment inspections due, and recent
  high-severity open incidents.
- Optional **email digest** — set `HSE_SMTP_HOST` and `HSE_ALERT_RECIPIENTS`
  (and `HSE_SMTP_USER`/`HSE_SMTP_PASS`) to email it; otherwise it's previewed in-app.

### 4. Dashboards & reports
- Executive **Dashboard** — 12 KPI cards with traffic-light status (TRIFR, LTIFR,
  near misses, inspection score, training, env compliance, open/overdue actions,
  days since last LTI …) plus charts, all driven by **Year / Month / Department /
  Location** filters.
- **Report** page — print-ready (Print → Save as PDF) board pack, plus one-click
  Excel exports.

### 5. Industry modules
- **Frequency rates** — rolling 12-month TRIFR / LTIFR / AIFR vs targets.
- **Training & competency** — certificate / statutory-licence / medical expiry tracking.
- **Investigations / RCA** — HiPo classification + 5-Whys / ICAM root cause, linked to incidents.
- **Tailings (GISTM)** — dam-safety inspections + piezometer monitoring vs thresholds.

---

## Pages

| Page | What it shows |
|---|---|
| **Dashboard** | KPI scorecard + charts (incident trend, by type, high-risk locations, actions by status, training by department), filterable. |
| **Frequency Rates** | Rolling 12-month TRIFR / LTIFR / AIFR vs targets, with trend chart and traffic-light cards. |
| **Incident Register** | Incident table; overdue CARs highlighted; **+ Log incident**. |
| **Event Reports** | Captured near misses / hazards / observations with category & trend charts; **+ Report event**. |
| **Corrective Actions** | Tracker with inline status change; overdue flagged via `TODAY()`; **+ New action**. |
| **Training & Competency** | Competency matrix with certificate / statutory-licence / medical expiry (Valid / Expiring / Expired); currency by department. |
| **Investigations / RCA** | Incident root-cause workflow: HiPo classification, 5-Whys / ICAM, linked to the register. |
| **Compliance** | Ghana regulatory register (Minerals Commission, EPA Ghana, Factories Inspectorate, ICMC cyanide code, Nuclear Regulatory Authority, Water Resources Commission …) with auto-derived status. |
| **Environmental** | PM10 dust & WAD cyanide vs regulatory limits, energy/water trends, exceedances flagged. |
| **Tailings (GISTM)** | TSF dam-safety inspections, freeboard, and piezometer phreatic readings vs thresholds. |
| **Contractors** | Owner-vs-contractor TRIFR comparison. |
| **Registers** | Permits, audits, safety-critical equipment with status. |
| **Alerts** | Everything needing attention, grouped & escalated; optional email digest. |
| **Report** | Print/PDF board report + Excel export. |
| **Data & Refresh** | DB datasets, Excel import (admin), per-dataset export, refresh. |
| **Admin → Users / Audit Trail** | Manage accounts & roles; view the change log. |

---

## Architecture

```
data/*.xlsx ──import──►  SQLite DB  ──►  DataStore (pandas analytics)  ──►  Flask + Chart.js UI
                          ▲   ▲
              in-app capture   admin re-import / export
```

| File | Role |
|---|---|
| `app.py` | Flask routes, auth, role gating, capture, alerts, reports |
| `core.py` | `DataStore` — reads the DB, derives fields, computes all KPIs/charts |
| `models.py` | ORM models: `User`, `AuditLog`, `Event`, `ImportRun`, `Investigation` |
| `importer.py` | Validated Excel→DB import (audit + rejected rows + real-workbook profile), seeding, write helpers |
| `forms.py` | Flask-WTF forms (login + capture, CSRF-protected) |
| `alerts.py` | Alert computation + email digest |
| `reports.py` | Excel export helpers |
| `config.py` | Targets, limits, roles, SMTP, domain lists |
| `templates/`, `static/` | Jinja2 pages; CSS; vendored Chart.js (offline) |
| `data/` | Sample Excel datasets (edit / replace with your own) |
| `tests/`, `.github/` | pytest suite + GitHub Actions CI |

Built with **Flask, Flask-SQLAlchemy, Flask-Login, Flask-WTF, pandas, openpyxl**;
charts use **Chart.js** vendored locally (works offline). The SQLite DB (`hse.db`)
is created on first run and is **not** committed — it reseeds from `data/`.

### Configuration (environment variables)
`PORT`, `HSE_SECRET_KEY`, `HSE_DATABASE_URI`, `HSE_ADMIN_PASSWORD`,
`HSE_SMTP_HOST` / `HSE_SMTP_PORT` / `HSE_SMTP_USER` / `HSE_SMTP_PASS` /
`HSE_SMTP_FROM` / `HSE_ALERT_RECIPIENTS`.

---

## Bonus: standalone Excel dashboard
`generate_hse_dashboard.py` builds a single, fully **formula-driven** `.xlsx`
workbook (live SUMPRODUCT/COUNTIF KPIs, conditional formatting, charts, data
validation, Excel Tables). Validate it with `python verify_workbook.py`.

---

## Production notes
This ships with Flask's development server for convenience. For real multi-user
deployment, serve with **waitress** (`waitress-serve --port=8080 app:app`) behind
HTTPS, set a strong `HSE_SECRET_KEY`, change the seeded passwords, and back up
`hse.db` (or point `HSE_DATABASE_URI` at PostgreSQL).
