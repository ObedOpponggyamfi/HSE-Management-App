# HSE Management App (Gold Mine)

A local **Python web app** that consolidates safety data from Excel spreadsheets
into a live, refreshable **Health, Safety & Environment dashboard** — tailored for
an Asanko-style gold mine in Ghana.

Drop your safety Excel files into a `data/` folder, run the app, and you get an
executive dashboard with KPIs, charts, filters and registers. Edit the Excel files
and click **Refresh** — every number, traffic-light and chart updates. No database,
no cloud, runs entirely on your PC.

> The repo also ships a separate generator that builds a fully-formula-driven Excel
> dashboard workbook (see [Bonus](#bonus-standalone-excel-dashboard) below).

---

## Quick start

```bash
pip install -r requirements.txt
python generate_dummy_data.py     # creates sample Excel files in data/ (first run only)
python app.py                      # starts the app and opens your browser
```

The app serves at **http://127.0.0.1:5050/** (it auto-opens your browser).
To use a different port: `set PORT=8080 && python app.py` (Windows) /
`PORT=8080 python app.py` (macOS/Linux).

*Why 5050 and not 5000?* Port 5000 is often taken by other local apps (e.g. another
Flask project), so the default avoids the clash. Change it with `PORT` anytime.

---

## How the workflow works

```
   data/*.xlsx  ──read & consolidate──►  in-memory analytics  ──►  web dashboard
        ▲                                      ▲
   you edit / add rows                    click "Refresh"
```

1. Safety data lives as Excel files in **`data/`** — one file per dataset
   (incidents, activity log, corrective actions, compliance, environmental,
   permits, audits, equipment).
2. The app reads and **consolidates** all of them on startup, derives the fields it
   needs (Year/Month, recordable & LTI flags, overdue/compliance status, frequency
   rates …) and computes every KPI and chart.
3. Edit or append rows in the Excel files (keep the column headers unchanged), then
   click **⟳ Refresh** in the app — it re-reads the folder and recomputes everything.
4. The **Data & Refresh** page lists every file, its row count, last-modified time
   and load status.

Targets, rate bases and regulatory limits that drive the traffic-lights live in
**`config.py`** (TRIFR 3.0, LTIFR 0.8, inspection 95 %, training 90 %, PM10 70 µg/m³,
WAD CN 50 mg/L, pH 6–9 …).

---

## Pages

| Page | What it shows |
|---|---|
| **Dashboard** | 12 KPI cards with traffic-light status (TRIFR, LTIFR, near misses, inspection score, training, env compliance, open/overdue actions, days since last LTI …) + charts: incident trend, by type, high-risk locations, actions by status, training by department. All driven by the **Year / Month / Department / Location** filters. |
| **Incident Register** | Full incident table; overdue CARs highlighted; severity & status pills. |
| **Corrective Actions** | Actions tracker with status summary; overdue items flagged via `TODAY()`. |
| **Compliance** | Ghana regulatory register (Minerals Commission, EPA Ghana, Factories Inspectorate, ICMC cyanide code, Nuclear Regulatory Authority, Water Resources Commission …) with auto-derived Compliant / Due Soon / Overdue status and overall %. |
| **Environmental** | PM10 dust, WAD cyanide vs regulatory limits, energy & water trends, monthly table with exceedances flagged. |
| **Contractors** | Owner-vs-contractor TRIFR comparison (man-hours & recordables consolidated from the spine). |
| **Registers** | Permit-to-work, audits and safety-critical equipment registers with status. |
| **Data & Refresh** | File consolidation status + one-click refresh. |

---

## Project layout

```
app.py                  Flask app (routes, filters, refresh)
core.py                 consolidation + analytics engine (DataStore)
config.py               targets, limits, domain lists, data-file registry
generate_dummy_data.py  writes sample Excel datasets into data/
data/                   the Excel files the app consolidates  (edit these)
templates/              Jinja2 pages (dashboard, incidents, … )
static/                 css, vendored Chart.js (offline), chart JS
requirements.txt
```

Built with **Flask + pandas + openpyxl**; charts use **Chart.js** (vendored locally,
so the app works fully offline).

### Using your own data
Replace the files in `data/` with your real spreadsheets, keeping the same file
names, sheet names and column headers (see `config.py → DATASETS` and
`generate_dummy_data.py` for the expected columns). Then click **Refresh**.

---

## Bonus: standalone Excel dashboard

`generate_hse_dashboard.py` builds a single, fully **formula-driven** `.xlsx`
workbook (live SUMPRODUCT/COUNTIF KPIs, conditional-format traffic lights, charts,
data-validation dropdowns, Excel Tables) — useful if you want the dashboard *inside*
Excel rather than the web app. Validate it with `python verify_workbook.py`.

```bash
python generate_hse_dashboard.py   # -> Asanko_HSE_Management_Dashboard.xlsx
```

---

## Notes
- Sample data uses a fixed seed and is anchored to today's date, so overdue /
  days-since-LTI logic is meaningful the moment you run it.
- This is a local development server; for multi-user deployment put it behind a
  production WSGI server (gunicorn/waitress) and a real datastore.
