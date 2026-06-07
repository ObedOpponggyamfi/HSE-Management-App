from __future__ import annotations

import datetime as dt
import json
import os
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Iterable

import pandas as pd
from sqlalchemy import delete
from sqlalchemy.engine import Engine
from sqlalchemy.orm import Session

import config as C
from .database import DEFAULT_SITE_ID, initialize_database, make_engine
from .models import (
    Activity,
    Audit,
    ComplianceItem,
    CorrectiveAction,
    EnvironmentalRecord,
    Equipment,
    ImportRun,
    Incident,
    Permit,
    utc_now,
)


@dataclass
class ImportResult:
    import_id: int | None
    dataset: str
    source_file: str
    profile: str
    rows_seen: int = 0
    rows_accepted: int = 0
    rows_rejected: int = 0
    errors: list[dict] = field(default_factory=list)
    imported_at: str | None = None

    def to_dict(self) -> dict:
        return asdict(self)


MODEL_BY_DATASET = {
    "incidents": Incident,
    "activity": Activity,
    "actions": CorrectiveAction,
    "compliance": ComplianceItem,
    "environmental": EnvironmentalRecord,
    "permits": Permit,
    "audits": Audit,
    "equipment": Equipment,
}


def _clean(value, default: str = "") -> str:
    if value is None or pd.isna(value):
        return default
    return str(value).strip()


def _date(value) -> dt.date | None:
    if value is None or pd.isna(value):
        return None
    value = pd.to_datetime(value, errors="coerce")
    if pd.isna(value):
        return None
    return value.date()


def _int(value, default: int = 0) -> int:
    try:
        if value is None or pd.isna(value):
            return default
        return int(round(float(value)))
    except (TypeError, ValueError):
        return default


def _float(value, default: float = 0.0) -> float:
    try:
        if value is None or pd.isna(value):
            return default
        return float(value)
    except (TypeError, ValueError):
        return default


def _severity_from_risk(value) -> int:
    text = _clean(value).lower()
    if not text:
        return 1
    if any(token in text for token in ("critical", "catastrophic", "extreme", "5")):
        return 5
    if any(token in text for token in ("high", "major", "4")):
        return 4
    if any(token in text for token in ("medium", "moderate", "3")):
        return 3
    if any(token in text for token in ("low", "minor", "2")):
        return 2
    return 1


def _source_kwargs(site_id: str, run_id: int, source_file: str, source_row: int | None) -> dict:
    return {
        "site_id": site_id,
        "import_run_id": run_id,
        "source_file": os.path.basename(source_file),
        "source_row": source_row,
    }


def _record_error(errors: list[dict], row: int | None, message: str, **extra) -> None:
    entry = {"row": row, "message": message}
    entry.update(extra)
    errors.append(entry)


def _replace_source_rows(session: Session, model, site_id: str, source_file: str) -> None:
    session.execute(delete(model).where(
        model.site_id == site_id,
        model.source_file == os.path.basename(source_file),
    ))


def _import_run(session: Session, site_id: str, dataset: str, source_file: str,
                profile: str) -> ImportRun:
    run = ImportRun(
        site_id=site_id,
        dataset=dataset,
        source_file=os.path.basename(source_file),
        profile=profile,
        started_at=utc_now(),
    )
    session.add(run)
    session.flush()
    return run


def _finish_run(session: Session, run: ImportRun, result: ImportResult) -> ImportResult:
    now = utc_now()
    run.finished_at = now
    run.rows_seen = result.rows_seen
    run.rows_accepted = result.rows_accepted
    run.rows_rejected = result.rows_rejected
    run.errors_json = json.dumps(result.errors, default=str)
    result.import_id = run.id
    result.imported_at = now.isoformat(timespec="seconds")
    session.commit()
    return result


def _load_excel(path: str | os.PathLike, *, sheet_name: str, header: int = 0) -> pd.DataFrame:
    return pd.read_excel(path, sheet_name=sheet_name, header=header)


def _import_standard_rows(session: Session, site_id: str, dataset: str, source_file: str,
                          df: pd.DataFrame, run: ImportRun) -> ImportResult:
    result = ImportResult(run.id, dataset, os.path.basename(source_file), "standard",
                          rows_seen=len(df))
    model = MODEL_BY_DATASET[dataset]
    _replace_source_rows(session, model, site_id, source_file)

    for offset, row in df.dropna(how="all").iterrows():
        row_no = int(offset) + 2
        try:
            if dataset == "incidents":
                when = _date(row.get("Date"))
                incident_id = _clean(row.get("ID"))
                if not incident_id or when is None:
                    raise ValueError("Incident ID and Date are required")
                obj = Incident(
                    **_source_kwargs(site_id, run.id, source_file, row_no),
                    incident_id=incident_id,
                    date=when,
                    area=_clean(row.get("Area")),
                    department=_clean(row.get("Department")),
                    company=_clean(row.get("Company")),
                    type=_clean(row.get("Type"), "Other"),
                    incident_class=_clean(row.get("Class")),
                    severity=max(1, min(5, _int(row.get("Severity"), 1))),
                    status=_clean(row.get("Status")),
                    reported=_date(row.get("Reported")),
                    car_due=_date(row.get("CAR_Due")),
                    owner=_clean(row.get("Owner")),
                )
            elif dataset == "activity":
                period = _date(row.get("Period"))
                if period is None:
                    raise ValueError("Period is required")
                obj = Activity(
                    **_source_kwargs(site_id, run.id, source_file, row_no),
                    period=period,
                    area=_clean(row.get("Area")),
                    department=_clean(row.get("Department")),
                    company=_clean(row.get("Company")),
                    manhours=_int(row.get("ManHours")),
                    nearmisses=_int(row.get("NearMisses")),
                    hazards=_int(row.get("Hazards")),
                    obsraised=_int(row.get("ObsRaised")),
                    obsclosed=_int(row.get("ObsClosed")),
                    inspplanned=_int(row.get("InspPlanned")),
                    inspdone=_int(row.get("InspDone")),
                    inspscore=_float(row.get("InspScore")),
                    audits=_int(row.get("Audits")),
                    toolbox=_int(row.get("Toolbox")),
                    trainassigned=_int(row.get("TrainAssigned")),
                    traincompleted=_int(row.get("TrainCompleted")),
                    ppe=_float(row.get("PPE")),
                )
            elif dataset == "actions":
                action_id = _clean(row.get("Action_ID"))
                if not action_id:
                    raise ValueError("Action_ID is required")
                obj = CorrectiveAction(
                    **_source_kwargs(site_id, run.id, source_file, row_no),
                    action_id=action_id,
                    source_incident=_clean(row.get("Source_Incident")),
                    description=_clean(row.get("Description")),
                    raised=_date(row.get("Raised")),
                    due=_date(row.get("Due")),
                    owner=_clean(row.get("Owner")),
                    department=_clean(row.get("Department")),
                    area=_clean(row.get("Area")),
                    priority=_clean(row.get("Priority")),
                    status=_clean(row.get("Status")),
                )
            elif dataset == "compliance":
                obj = ComplianceItem(
                    **_source_kwargs(site_id, run.id, source_file, row_no),
                    item=_clean(row.get("Item")),
                    regulator=_clean(row.get("Regulator")),
                    reference=_clean(row.get("Reference")),
                    frequency_months=max(1, _int(row.get("Frequency_Months"), 12)),
                    last_completed=_date(row.get("Last_Completed")),
                    owner=_clean(row.get("Owner")),
                )
            elif dataset == "environmental":
                period = _date(row.get("Period"))
                if period is None:
                    raise ValueError("Period is required")
                obj = EnvironmentalRecord(
                    **_source_kwargs(site_id, run.id, source_file, row_no),
                    period=period,
                    waste_t=_float(row.get("Waste_t")),
                    recycling=_float(row.get("Recycling")),
                    energy_mwh=_float(row.get("Energy_MWh")),
                    water_m3=_float(row.get("Water_m3")),
                    fuel_l=_float(row.get("Fuel_L")),
                    pm10=_float(row.get("PM10")),
                    ph=_float(row.get("pH")),
                    wad_cn=_float(row.get("WAD_CN")),
                )
            elif dataset == "permits":
                obj = Permit(
                    **_source_kwargs(site_id, run.id, source_file, row_no),
                    permit=_clean(row.get("Permit")),
                    authority=_clean(row.get("Authority")),
                    holder=_clean(row.get("Holder")),
                    issue_date=_date(row.get("Issue_Date")),
                    expiry_date=_date(row.get("Expiry_Date")),
                )
            elif dataset == "audits":
                obj = Audit(
                    **_source_kwargs(site_id, run.id, source_file, row_no),
                    audit=_clean(row.get("Audit")),
                    type=_clean(row.get("Type")),
                    date=_date(row.get("Date")),
                    auditor=_clean(row.get("Auditor")),
                    score=_float(row.get("Score")),
                    findings=_int(row.get("Findings")),
                    closed_findings=_int(row.get("Closed_Findings")),
                )
            elif dataset == "equipment":
                obj = Equipment(
                    **_source_kwargs(site_id, run.id, source_file, row_no),
                    asset_id=_clean(row.get("Asset_ID")),
                    asset=_clean(row.get("Asset")),
                    type=_clean(row.get("Type")),
                    location=_clean(row.get("Location")),
                    last_inspection=_date(row.get("Last_Inspection")),
                    next_inspection=_date(row.get("Next_Inspection")),
                )
            else:
                raise ValueError(f"Unsupported dataset: {dataset}")
            session.add(obj)
            result.rows_accepted += 1
        except Exception as exc:
            result.rows_rejected += 1
            _record_error(result.errors, row_no, str(exc))

    return result


def import_standard_dataset(session: Session, site_id: str, dataset: str, path: str,
                            sheet_name: str) -> ImportResult:
    run = _import_run(session, site_id, dataset, path, "standard")
    result = ImportResult(run.id, dataset, os.path.basename(path), "standard")
    try:
        if not os.path.exists(path):
            _record_error(result.errors, None, "file not found")
            return _finish_run(session, run, result)
        df = _load_excel(path, sheet_name=sheet_name)
        result = _import_standard_rows(session, site_id, dataset, path, df, run)
    except Exception as exc:
        _record_error(result.errors, None, str(exc))
        result.rows_rejected = result.rows_seen
    return _finish_run(session, run, result)


def _looks_date_only(row: pd.Series) -> bool:
    meaningful = [
        "CASE ID", "BUSINESS DEPARTMENT / UNIT", "BUSINESS PARTNER",
        "LOCATION", "INCIDENT TYPE", "STATUS", "INCIDENT DESCRIPTION",
        "DETAILS OF INCIDENT",
    ]
    return bool(_date(row.get("DATE OF INCIDENT")) and all(not _clean(row.get(c)) for c in meaningful))


def import_rich_incident_workbook(session: Session, site_id: str, path: str,
                                  profile: str = "asanko_incidents_v1") -> ImportResult:
    spec = C.WORKBOOK_PROFILES[profile]
    run = _import_run(session, site_id, spec["dataset"], path, profile)
    result = ImportResult(run.id, spec["dataset"], os.path.basename(path), profile)
    try:
        if not os.path.exists(path):
            _record_error(result.errors, None, "file not found")
            return _finish_run(session, run, result)

        df = _load_excel(path, sheet_name=spec["sheet"], header=spec.get("header", 1))
        result.rows_seen = len(df)
        _replace_source_rows(session, Incident, site_id, path)

        seen_case_ids: set[str] = set()
        ignored_date_only = 0
        for offset, row in df.dropna(how="all").iterrows():
            row_no = int(offset) + spec.get("header", 1) + 2
            if _looks_date_only(row):
                ignored_date_only += 1
                continue

            case_id = _clean(row.get("CASE ID"))
            incident_date = _date(row.get("DATE OF INCIDENT"))
            if not case_id and incident_date is None:
                continue
            if not case_id or incident_date is None:
                result.rows_rejected += 1
                _record_error(result.errors, row_no, "CASE ID and DATE OF INCIDENT are required")
                continue
            if case_id in seen_case_ids:
                result.rows_rejected += 1
                _record_error(result.errors, row_no, "duplicate CASE ID in workbook", case_id=case_id)
                continue
            seen_case_ids.add(case_id)

            risk_value = row.get("RISK RATING (Actual Consequence)") or row.get(
                "RISK RATING (Potential Consequence)"
            )
            obj = Incident(
                **_source_kwargs(site_id, run.id, path, row_no),
                incident_id=case_id,
                date=incident_date,
                area=_clean(row.get("LOCATION")) or _clean(row.get("SPECIFIC LOCATION")),
                department=_clean(row.get("BUSINESS DEPARTMENT / UNIT")),
                company=_clean(row.get("BUSINESS PARTNER")),
                type=_clean(row.get("INCIDENT TYPE"), "Other"),
                incident_class=_clean(row.get("INJURY OUTCOME")) or _clean(row.get("TYPE OF INJURY")),
                severity=_severity_from_risk(risk_value),
                status=_clean(row.get("STATUS"), "Closed"),
                reported=incident_date,
                car_due=None,
                owner=_clean(row.get("LEAD INVESTIGATOR")),
                description=_clean(row.get("DETAILS OF INCIDENT")) or _clean(row.get("INCIDENT DESCRIPTION")),
                root_cause=_clean(row.get("ROOT CAUSE")),
                risk_actual=_clean(row.get("RISK RATING (Actual Consequence)")),
                risk_potential=_clean(row.get("RISK RATING (Potential Consequence)")),
            )
            session.add(obj)
            result.rows_accepted += 1

        if ignored_date_only:
            _record_error(result.errors, None, "ignored date-only rows", rows=ignored_date_only)
    except Exception as exc:
        _record_error(result.errors, None, str(exc))
        result.rows_rejected = max(result.rows_rejected, result.rows_seen - result.rows_accepted)

    return _finish_run(session, run, result)


def import_spreadsheets(site_id: str = DEFAULT_SITE_ID, data_dir: str | os.PathLike = C.DATA_DIR,
                        engine: Engine | None = None, include_workbooks: bool = True) -> list[ImportResult]:
    engine = engine or make_engine()
    initialize_database(engine, site_id=site_id)
    results: list[ImportResult] = []
    with Session(engine) as session:
        for dataset, spec in C.DATASETS.items():
            path = os.path.join(str(data_dir), spec["file"])
            results.append(import_standard_dataset(session, site_id, dataset, path, spec["sheet"]))

        if include_workbooks:
            for profile, spec in C.WORKBOOK_PROFILES.items():
                path = os.path.join(str(data_dir), spec["file"])
                if os.path.exists(path):
                    results.append(import_rich_incident_workbook(session, site_id, path, profile))
    return results


def import_workbook(site_id: str, path: str, profile: str,
                    engine: Engine | None = None) -> ImportResult:
    engine = engine or make_engine()
    initialize_database(engine, site_id=site_id)
    with Session(engine) as session:
        if profile == "asanko_incidents_v1":
            return import_rich_incident_workbook(session, site_id, path, profile)
        raise ValueError(f"Unsupported workbook profile: {profile}")


def results_to_json(results: ImportResult | Iterable[ImportResult]) -> str:
    if isinstance(results, ImportResult):
        payload = results.to_dict()
    else:
        payload = [r.to_dict() for r in results]
    return json.dumps(payload, indent=2, default=str)
