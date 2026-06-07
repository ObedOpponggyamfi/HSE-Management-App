from __future__ import annotations

import os
from pathlib import Path

from sqlalchemy import create_engine, func, select
from sqlalchemy.engine import Engine
from sqlalchemy.orm import Session, sessionmaker

import config as C
from .models import Base, ImportRun, Site

DEFAULT_DB_PATH = os.path.join(C.INSTANCE_DIR, "hse_dashboard.sqlite")
DEFAULT_SITE_ID = os.environ.get("HSE_SITE_ID", "default")
DEFAULT_SITE_NAME = os.environ.get("HSE_SITE_NAME", C.COMPANY_NAME)


def make_engine(db_path: str | None = None) -> Engine:
    path = db_path or DEFAULT_DB_PATH
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    return create_engine(f"sqlite:///{path}", future=True)


def initialize_database(engine: Engine, site_id: str = DEFAULT_SITE_ID,
                        site_name: str = DEFAULT_SITE_NAME) -> None:
    Base.metadata.create_all(engine)
    with Session(engine) as session:
        if session.get(Site, site_id) is None:
            session.add(Site(id=site_id, name=site_name))
            session.commit()


def session_factory(engine: Engine) -> sessionmaker[Session]:
    return sessionmaker(bind=engine, future=True)


def import_version(session: Session, site_id: str = DEFAULT_SITE_ID) -> str:
    stmt = select(func.max(ImportRun.finished_at)).where(ImportRun.site_id == site_id)
    value = session.execute(stmt).scalar_one_or_none()
    return value.isoformat() if value else "empty"


def has_operational_rows(session: Session, site_id: str = DEFAULT_SITE_ID) -> bool:
    from .models import Incident, Activity

    inc = session.execute(select(func.count()).select_from(Incident)
                          .where(Incident.site_id == site_id)).scalar_one()
    act = session.execute(select(func.count()).select_from(Activity)
                          .where(Activity.site_id == site_id)).scalar_one()
    return bool(inc or act)
