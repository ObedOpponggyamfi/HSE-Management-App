#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ORM models for the app-managed entities:
  - User      : authentication + role-based access
  - AuditLog  : who-changed-what trail (ISO 45001 documented-information control)
  - Event     : individually-captured Near Miss / Hazard / Observation reports

The bulk safety datasets (incidents, activity, compliance ...) live in plain
tables created by the importer and queried with pandas; they don't need ORM
mapping. The entities here are the ones the app creates/edits directly.
"""
import datetime as dt

from flask_login import UserMixin
from werkzeug.security import check_password_hash, generate_password_hash

import config as C
from extensions import db, login_manager


def _utcnow():
    return dt.datetime.now(dt.timezone.utc).replace(tzinfo=None)


class User(UserMixin, db.Model):
    __tablename__ = "users"
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False, index=True)
    name = db.Column(db.String(120))
    email = db.Column(db.String(120))
    password_hash = db.Column(db.String(255))
    role = db.Column(db.String(20), default="worker")
    active = db.Column(db.Boolean, default=True)
    created_at = db.Column(db.DateTime, default=_utcnow)

    def set_password(self, pw):
        self.password_hash = generate_password_hash(pw)

    def check_password(self, pw):
        return check_password_hash(self.password_hash, pw)

    @property
    def is_active(self):
        return bool(self.active)

    def can(self, required_role):
        """True if this user's role rank >= the required role's rank."""
        return C.ROLE_RANK.get(self.role, 0) >= C.ROLE_RANK.get(required_role, 99)

    @property
    def role_label(self):
        return C.ROLE_LABELS.get(self.role, self.role)


@login_manager.user_loader
def load_user(user_id):
    return db.session.get(User, int(user_id))


class AuditLog(db.Model):
    __tablename__ = "audit_log"
    id = db.Column(db.Integer, primary_key=True)
    ts = db.Column(db.DateTime, default=_utcnow, index=True)
    username = db.Column(db.String(80))
    action = db.Column(db.String(40))        # create / update / login / import ...
    entity = db.Column(db.String(40))        # incident / action / event / user ...
    ref = db.Column(db.String(80))
    detail = db.Column(db.Text)


class Event(db.Model):
    __tablename__ = "events"
    id = db.Column(db.Integer, primary_key=True)
    ref = db.Column(db.String(20))
    category = db.Column(db.String(30))      # Near Miss / Hazard / Observation
    date = db.Column(db.Date, index=True)
    area = db.Column(db.String(80))
    department = db.Column(db.String(80))
    severity = db.Column(db.Integer)
    description = db.Column(db.Text)
    reported_by = db.Column(db.String(80))
    status = db.Column(db.String(20), default="Open")
    created_at = db.Column(db.DateTime, default=_utcnow)


class ImportRun(db.Model):
    """Audit record of each spreadsheet import (rows seen / accepted / rejected)."""
    __tablename__ = "import_runs"
    id = db.Column(db.Integer, primary_key=True)
    ts = db.Column(db.DateTime, default=_utcnow, index=True)
    dataset = db.Column(db.String(40), index=True)
    source_file = db.Column(db.String(260))
    profile = db.Column(db.String(60), default="standard")
    rows_seen = db.Column(db.Integer, default=0)
    rows_accepted = db.Column(db.Integer, default=0)
    rows_rejected = db.Column(db.Integer, default=0)
    errors = db.Column(db.Text, default="[]")


class Investigation(db.Model):
    """Root-cause investigation linked to an incident (5-Whys / ICAM / Fishbone)."""
    __tablename__ = "investigations"
    id = db.Column(db.Integer, primary_key=True)
    ref = db.Column(db.String(20))
    incident_id = db.Column(db.String(80), index=True)
    hipo = db.Column(db.Boolean, default=False)          # high-potential incident
    method = db.Column(db.String(40), default="5-Whys")
    immediate_cause = db.Column(db.Text, default="")
    root_cause = db.Column(db.Text, default="")
    why1 = db.Column(db.String(300), default="")
    why2 = db.Column(db.String(300), default="")
    why3 = db.Column(db.String(300), default="")
    why4 = db.Column(db.String(300), default="")
    why5 = db.Column(db.String(300), default="")
    status = db.Column(db.String(20), default="Open")
    investigator = db.Column(db.String(120), default="")
    created_by = db.Column(db.String(80))
    created_at = db.Column(db.DateTime, default=_utcnow)
