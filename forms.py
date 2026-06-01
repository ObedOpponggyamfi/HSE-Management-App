#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Flask-WTF forms (CSRF-protected) for login + data capture."""
from flask_wtf import FlaskForm
from wtforms import (DateField, PasswordField, SelectField, StringField,
                     SubmitField, TextAreaField)
from wtforms.validators import DataRequired, Email, Length, Optional

import config as C

DATE_KW = {"type": "date"}


def _ch(values):
    return [(v, v) for v in values]


class LoginForm(FlaskForm):
    username = StringField("Username", validators=[DataRequired()])
    password = PasswordField("Password", validators=[DataRequired()])
    submit = SubmitField("Sign in")


class IncidentForm(FlaskForm):
    date = DateField("Date of incident", validators=[DataRequired()], render_kw=DATE_KW)
    area = SelectField("Location / Area", choices=_ch(C.AREAS), validators=[DataRequired()])
    department = SelectField("Department", choices=_ch(C.DEPARTMENTS), validators=[DataRequired()])
    company = SelectField("Company", choices=_ch(C.COMPANIES), validators=[DataRequired()])
    type = SelectField("Type", choices=_ch(C.INCIDENT_TYPES), validators=[DataRequired()])
    klass = SelectField("Class", choices=_ch(C.INCIDENT_CLASSES), validators=[DataRequired()])
    severity = SelectField("Severity (1-5)", coerce=int,
                           choices=[(i, str(i)) for i in range(1, 6)], validators=[DataRequired()])
    status = SelectField("Status", choices=_ch(C.INCIDENT_STATUS), validators=[DataRequired()])
    owner = StringField("Owner / Investigator", validators=[DataRequired()])
    reported = DateField("Date reported", validators=[Optional()], render_kw=DATE_KW)
    car_due = DateField("Corrective action due", validators=[Optional()], render_kw=DATE_KW)
    submit = SubmitField("Log incident")


class EventForm(FlaskForm):
    category = SelectField("Category", choices=_ch(C.EVENT_CATEGORIES), validators=[DataRequired()])
    date = DateField("Date", validators=[DataRequired()], render_kw=DATE_KW)
    area = SelectField("Location / Area", choices=_ch(C.AREAS), validators=[DataRequired()])
    severity = SelectField("Potential severity (1-5)", coerce=int,
                           choices=[(i, str(i)) for i in range(1, 6)], validators=[DataRequired()])
    description = TextAreaField("What happened / what was observed",
                                validators=[DataRequired(), Length(min=5)])
    submit = SubmitField("Submit report")


class ActionForm(FlaskForm):
    description = TextAreaField("Action description", validators=[DataRequired()])
    area = SelectField("Location / Area", choices=_ch(C.AREAS), validators=[DataRequired()])
    department = SelectField("Department", choices=_ch(C.DEPARTMENTS), validators=[DataRequired()])
    owner = StringField("Owner", validators=[DataRequired()])
    priority = SelectField("Priority", choices=_ch(C.PRIORITY), validators=[DataRequired()])
    due = DateField("Due date", validators=[DataRequired()], render_kw=DATE_KW)
    source = StringField("Source", default="Manual")
    submit = SubmitField("Create action")


class UserForm(FlaskForm):
    username = StringField("Username", validators=[DataRequired(), Length(min=3)])
    name = StringField("Full name", validators=[DataRequired()])
    email = StringField("Email", validators=[Optional(), Email()])
    role = SelectField("Role", choices=_ch(C.ROLE_ORDER), validators=[DataRequired()])
    password = PasswordField("Password", validators=[DataRequired(), Length(min=5)])
    submit = SubmitField("Create user")
