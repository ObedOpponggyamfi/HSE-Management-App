# HSE Management Dashboard (Gold Mine)

An **enterprise-grade, multi-sheet Excel HSE (Health, Safety & Environment) management
dashboard**, generated from a single Python script. It is tailored for an
**Asanko-style gold mine in Ghana** but built to the feature scope of a premium
commercial HSE dashboard.

The workbook is **live**, not a static value dump: every KPI, status light and chart
is a native Excel formula that points back at one auditable data spine. Add a new row
to the data sheets and **every number, traffic-light and chart recalculates
automatically** — no redesign, no re-running the script.

---

## Quick start

```bash
pip install -r requirements.txt
python generate_hse_dashboard.py
```

This produces **`Asanko_HSE_Management_Dashboard.xlsx`**. Open it in Excel and the
formulas calculate on load (the workbook sets `fullCalcOnLoad`). A pre-generated copy
is included in the repo so you can open it immediately.

> Only `openpyxl`, `pandas` and `numpy` are used — no internet-install or exotic
> dependencies.

---

## Core architecture — one data spine, everything else is a formula

| Spine sheet | Role |
|---|---|
| **Incident_Register** | One row per event (ID, Date, Area, Department, Company, Type, Class, Severity 1–5, Status, CAR due, Owner). `Year`, `Month`, `Recordable`, `LTI` and overdue-CAR flags auto-calculate. |
| **Activity_Log** | Per area, per month: man-hours, near misses, hazards, observations, inspections (planned/done/score), audits, toolbox talks, training (assigned/completed), PPE %. |
| **Settings** | Rate bases (1,000,000 mining / 200,000 OSHA), targets (TRIFR 3.0, LTIFR 0.8, inspection 95 %, training 90 %, env 98 %) and thresholds — exposed as **named ranges**. Every formula references these cells, never a literal. |

**You only ever edit `Incident_Register` and `Activity_Log`** (and optionally tune
`Settings` / extend dropdowns on the `Lists` sheet).

### Frequency rates
```
TRIFR = (Fatal + LTI + RWC + MTC) / man-hours × RATE_BASE
LTIFR = (Fatal + LTI)             / man-hours × RATE_BASE
```
(Fatalities are modelled as severity-5 Lost Time Injuries, so they count in both.)

### Interactive filtering
The Dashboard has **Year / Month / Department / Location** data-validation dropdowns.
Choosing a value re-drives every headline KPI (SUMPRODUCT with full **"All"** handling)
and all six charts. Set a filter back to `All` to clear it.

---

## Sheets / modules

- **Dashboard** — executive scorecard cards (formula-driven, conditional-format traffic
  lights vs target): Total Incidents, Near Misses, Inspection Score, Training Completion,
  Environmental Compliance %, Open/Overdue Actions, TRIFR, LTIFR, Days-Since-Last-LTI.
  Charts: multi-year incident trend, incidents-by-type doughnut, incidents-by-location
  ranking, inspection-score trend, leading-vs-lagging combo, actions-by-status.
- **Incident_Register** — the register module: an Excel Table with autofilter and
  conditional formatting on severity and overdue CARs.
- **Near_Miss** — reporting trend and near-miss : recordable ratio.
- **Risk_Matrix** — colour-coded 5×5 likelihood × consequence matrix plus a risk
  register (L, C, score = L×C, control adequacy).
- **Inspections** — completion rate, score trend, target achievement.
- **Training** — completion %, outstanding, by-department compliance.
- **Corrective_Actions** — actions-by-status tracker; overdue flagged via `TODAY()`.
- **Compliance** — Ghana regulatory register (Minerals Commission, EPA Ghana, Factories
  Inspectorate, ICMC cyanide code, Nuclear Regulatory Authority, Water Resources
  Commission …) with auto-derived Compliant / Due Soon / Overdue status.
- **Environmental** — waste, recycling, energy, water, fuel, plus PM10 dust / discharge
  pH / WAD cyanide vs regulatory limit lines and trend indicators.
- **Contractors** — owner-vs-contractor TRIFR comparison (man-hours & recordables live
  from the spine).
- **Permits / Audits / Drills / PPE / Equipment** — supporting registers with status
  formatting.
- **ReadMe** — in-workbook explanation of the architecture.
- **Settings / Lists / Calc** — named ranges, dropdown sources and the filter-aware
  chart engine (`Lists` and `Calc` are hidden helpers).

---

## Native Excel features used

Formulas (SUMPRODUCT / SUMIFS / COUNTIFS / AVERAGEIFS / TODAY / EDATE), conditional
formatting (traffic lights, colour scales, risk-matrix bands), data-validation
dropdowns, openpyxl charts (line / bar / doughnut / combo with secondary axis), named
ranges, Excel Tables (auto-expanding), frozen panes, and applied number/percent formats.

---

## Notes

- The seeded sample data uses a fixed numpy seed for reproducibility. Delete it and type
  your own — the workbook keeps working.
- `openpyxl` writes formulas but does not compute them; Excel calculates everything on
  first open.
