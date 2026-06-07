# HSE Management App (Gold Mine)

A local-first Python web app that imports HSE spreadsheet data into a canonical
SQLite store, then serves a refreshable Health, Safety & Environment dashboard
for an Asanko-style gold mine in Ghana.

The app remains usable on one device, but the data model includes `site_id`,
import history, and source-file provenance so it can later move toward a
multi-site deployment without rewriting the analytics layer.

## Quick Start

```bash
pip install -r requirements.txt
python generate_dummy_data.py
python -m hse_dashboard.import spreadsheets --site default --path data
python app.py
```

The app serves at `http://127.0.0.1:5050/`. Change the port with `PORT=8080`
or, on Windows PowerShell, `$env:PORT=8080; python app.py`.

On first app start, if the local database has no operational rows, the app will
import the spreadsheets in `data/` automatically. Clicking **Refresh** performs
another audited import into `instance/hse_dashboard.sqlite`.

## Data Workflow

```text
data/*.xlsx -> spreadsheet import adapters -> SQLite canonical store -> dashboard analytics
```

Supported sources:

- The configured workbook set in `config.py -> DATASETS`
- Optional rich incident workbook profile: `data/incidents.xlsx`, sheet
  `DataBase`, header row 2, profile `asanko_incidents_v1`

Import commands:

```bash
python -m hse_dashboard.import spreadsheets --site default --path data
python -m hse_dashboard.import workbook --site default --file data/incidents.xlsx --profile asanko_incidents_v1
```

Each import records an `import_id`, source file, accepted rows, rejected rows,
validation messages, and timestamp. The rich incident adapter ignores date-only
filler rows and rejects duplicate case IDs.

## App Pages

| Page | What it shows |
|---|---|
| Dashboard | KPI cards, trend charts, incident mix, action status, and training completion. |
| Incident Register | Canonical incident table with severity and overdue CAR indicators. |
| Corrective Actions | Action tracker filtered consistently by year/month/department/area. |
| Compliance | Regulatory register with due status and compliance percentage. |
| Environmental | Year/month-filtered environmental monitoring and limits. |
| Contractors | Owner/contractor TRIFR and exposure comparison. |
| Registers | Permits, audits, and safety-critical equipment. |
| Data & Refresh | Source import status, rejected rows, import IDs, data folder, and SQLite path. |

Operational tables include client-side search, sorting, and pagination.

## Configuration

Shared HSE definitions live in `config.py`:

- targets and rate bases (`TARGET_TRIFR`, `TARGET_TRAINING`, etc.)
- regulatory limits (`PM10_LIMIT`, `PH_MIN`, `PH_MAX`, `WADCN_LIMIT`)
- areas, departments, companies, owners, incident/action enumerations
- spreadsheet dataset registry and workbook import profiles

Useful environment variables:

- `HSE_DB_PATH`: SQLite path, default `instance/hse_dashboard.sqlite`
- `HSE_DATA_DIR`: spreadsheet directory, default `data/`
- `HSE_SITE_ID`: site identifier, default `default`
- `HSE_SECRET_KEY`: Flask session key
- `HSE_DEBUG=1`: enable Flask debug mode
- `HSE_NO_BROWSER=1`: prevent auto-opening the browser

## Standalone Excel Dashboard

The Excel workbook remains an independent output, but now shares definitions
from `config.py`.

```bash
python generate_hse_dashboard.py --as-of 2026-06-03 --output Asanko_HSE_Management_Dashboard.xlsx
python generate_hse_dashboard.py --as-of 2026-06-03 --chart-months 36
python verify_workbook.py
```

`--chart-months` reserves the calculation-sheet chart timeline intentionally;
use a larger value when generating a workbook meant to cover a longer reporting
window.

## Development

```bash
python -m pytest
python verify_workbook.py
```

CI runs the same checks on Windows. Runtime artifacts such as SQLite databases,
Python bytecode, Office lock files, and `instance/` are ignored by Git.
