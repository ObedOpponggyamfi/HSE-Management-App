from __future__ import annotations

import datetime as dt

from sqlalchemy import (
    Date,
    DateTime,
    Float,
    ForeignKey,
    Integer,
    String,
    Text,
    UniqueConstraint,
)
from sqlalchemy.orm import DeclarativeBase, Mapped, mapped_column, relationship


class Base(DeclarativeBase):
    pass


def utc_now() -> dt.datetime:
    return dt.datetime.now(dt.timezone.utc).replace(tzinfo=None)


class Site(Base):
    __tablename__ = "sites"

    id: Mapped[str] = mapped_column(String(64), primary_key=True)
    name: Mapped[str] = mapped_column(String(160), nullable=False)
    created_at: Mapped[dt.datetime] = mapped_column(DateTime, default=utc_now)


class ImportRun(Base):
    __tablename__ = "import_runs"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    site_id: Mapped[str] = mapped_column(ForeignKey("sites.id"), index=True)
    dataset: Mapped[str] = mapped_column(String(64), index=True)
    source_file: Mapped[str] = mapped_column(String(260))
    profile: Mapped[str] = mapped_column(String(80), default="default")
    started_at: Mapped[dt.datetime] = mapped_column(DateTime, default=utc_now)
    finished_at: Mapped[dt.datetime | None] = mapped_column(DateTime, nullable=True)
    rows_seen: Mapped[int] = mapped_column(Integer, default=0)
    rows_accepted: Mapped[int] = mapped_column(Integer, default=0)
    rows_rejected: Mapped[int] = mapped_column(Integer, default=0)
    errors_json: Mapped[str] = mapped_column(Text, default="[]")

    site: Mapped[Site] = relationship()


class SourceMixin:
    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    site_id: Mapped[str] = mapped_column(ForeignKey("sites.id"), index=True)
    import_run_id: Mapped[int | None] = mapped_column(ForeignKey("import_runs.id"), nullable=True)
    source_file: Mapped[str] = mapped_column(String(260), index=True)
    source_row: Mapped[int | None] = mapped_column(Integer, nullable=True)


class Incident(SourceMixin, Base):
    __tablename__ = "incidents"
    __table_args__ = (
        UniqueConstraint("site_id", "source_file", "incident_id", name="uq_incident_source_id"),
    )

    incident_id: Mapped[str] = mapped_column(String(80), index=True)
    date: Mapped[dt.date] = mapped_column(Date, index=True)
    area: Mapped[str] = mapped_column(String(160), default="")
    department: Mapped[str] = mapped_column(String(120), default="")
    company: Mapped[str] = mapped_column(String(160), default="")
    type: Mapped[str] = mapped_column(String(120), default="")
    incident_class: Mapped[str] = mapped_column(String(160), default="")
    severity: Mapped[int] = mapped_column(Integer, default=1)
    status: Mapped[str] = mapped_column(String(80), default="")
    reported: Mapped[dt.date | None] = mapped_column(Date, nullable=True)
    car_due: Mapped[dt.date | None] = mapped_column(Date, nullable=True)
    owner: Mapped[str] = mapped_column(String(120), default="")
    description: Mapped[str] = mapped_column(Text, default="")
    root_cause: Mapped[str] = mapped_column(Text, default="")
    risk_actual: Mapped[str] = mapped_column(String(80), default="")
    risk_potential: Mapped[str] = mapped_column(String(80), default="")


class Activity(SourceMixin, Base):
    __tablename__ = "activity"

    period: Mapped[dt.date] = mapped_column(Date, index=True)
    area: Mapped[str] = mapped_column(String(160), default="")
    department: Mapped[str] = mapped_column(String(120), default="")
    company: Mapped[str] = mapped_column(String(160), default="")
    manhours: Mapped[int] = mapped_column(Integer, default=0)
    nearmisses: Mapped[int] = mapped_column(Integer, default=0)
    hazards: Mapped[int] = mapped_column(Integer, default=0)
    obsraised: Mapped[int] = mapped_column(Integer, default=0)
    obsclosed: Mapped[int] = mapped_column(Integer, default=0)
    inspplanned: Mapped[int] = mapped_column(Integer, default=0)
    inspdone: Mapped[int] = mapped_column(Integer, default=0)
    inspscore: Mapped[float] = mapped_column(Float, default=0.0)
    audits: Mapped[int] = mapped_column(Integer, default=0)
    toolbox: Mapped[int] = mapped_column(Integer, default=0)
    trainassigned: Mapped[int] = mapped_column(Integer, default=0)
    traincompleted: Mapped[int] = mapped_column(Integer, default=0)
    ppe: Mapped[float] = mapped_column(Float, default=0.0)


class CorrectiveAction(SourceMixin, Base):
    __tablename__ = "corrective_actions"

    action_id: Mapped[str] = mapped_column(String(80), index=True)
    source_incident: Mapped[str] = mapped_column(String(80), default="")
    description: Mapped[str] = mapped_column(Text, default="")
    raised: Mapped[dt.date | None] = mapped_column(Date, nullable=True)
    due: Mapped[dt.date | None] = mapped_column(Date, nullable=True)
    owner: Mapped[str] = mapped_column(String(120), default="")
    department: Mapped[str] = mapped_column(String(120), default="")
    area: Mapped[str] = mapped_column(String(160), default="")
    priority: Mapped[str] = mapped_column(String(40), default="")
    status: Mapped[str] = mapped_column(String(80), default="")


class ComplianceItem(SourceMixin, Base):
    __tablename__ = "compliance_items"

    item: Mapped[str] = mapped_column(String(220), default="")
    regulator: Mapped[str] = mapped_column(String(160), default="")
    reference: Mapped[str] = mapped_column(String(120), default="")
    frequency_months: Mapped[int] = mapped_column(Integer, default=12)
    last_completed: Mapped[dt.date | None] = mapped_column(Date, nullable=True)
    owner: Mapped[str] = mapped_column(String(120), default="")


class EnvironmentalRecord(SourceMixin, Base):
    __tablename__ = "environmental_records"

    period: Mapped[dt.date] = mapped_column(Date, index=True)
    waste_t: Mapped[float] = mapped_column(Float, default=0.0)
    recycling: Mapped[float] = mapped_column(Float, default=0.0)
    energy_mwh: Mapped[float] = mapped_column(Float, default=0.0)
    water_m3: Mapped[float] = mapped_column(Float, default=0.0)
    fuel_l: Mapped[float] = mapped_column(Float, default=0.0)
    pm10: Mapped[float] = mapped_column(Float, default=0.0)
    ph: Mapped[float] = mapped_column(Float, default=0.0)
    wad_cn: Mapped[float] = mapped_column(Float, default=0.0)


class Permit(SourceMixin, Base):
    __tablename__ = "permits"

    permit: Mapped[str] = mapped_column(String(220), default="")
    authority: Mapped[str] = mapped_column(String(160), default="")
    holder: Mapped[str] = mapped_column(String(120), default="")
    issue_date: Mapped[dt.date | None] = mapped_column(Date, nullable=True)
    expiry_date: Mapped[dt.date | None] = mapped_column(Date, nullable=True)


class Audit(SourceMixin, Base):
    __tablename__ = "audits"

    audit: Mapped[str] = mapped_column(String(220), default="")
    type: Mapped[str] = mapped_column(String(80), default="")
    date: Mapped[dt.date | None] = mapped_column(Date, nullable=True)
    auditor: Mapped[str] = mapped_column(String(120), default="")
    score: Mapped[float] = mapped_column(Float, default=0.0)
    findings: Mapped[int] = mapped_column(Integer, default=0)
    closed_findings: Mapped[int] = mapped_column(Integer, default=0)


class Equipment(SourceMixin, Base):
    __tablename__ = "equipment"

    asset_id: Mapped[str] = mapped_column(String(80), default="")
    asset: Mapped[str] = mapped_column(String(220), default="")
    type: Mapped[str] = mapped_column(String(80), default="")
    location: Mapped[str] = mapped_column(String(160), default="")
    last_inspection: Mapped[dt.date | None] = mapped_column(Date, nullable=True)
    next_inspection: Mapped[dt.date | None] = mapped_column(Date, nullable=True)
