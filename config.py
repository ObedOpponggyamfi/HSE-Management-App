#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Shared configuration for the HSE Management web app:
  - domain enumerations (areas, departments, companies, incident types ...)
  - targets / rate bases / regulatory limits (used for traffic-light status)
  - file locations

Everything that the dummy-data generator, the consolidation engine and the
Flask app need to agree on lives here in one place.
"""
import os

# --------------------------------------------------------------------------
# Paths
# --------------------------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")          # drop Excel files here

# --------------------------------------------------------------------------
# Targets, rate bases & regulatory limits  (drive the traffic lights)
# --------------------------------------------------------------------------
RATE_BASE = 1_000_000                # mining standard (per 1,000,000 hours)
RATE_BASE_OSHA = 200_000

TARGET_TRIFR = 3.0
TARGET_LTIFR = 0.8
TARGET_INSPECTION = 0.95
TARGET_TRAINING = 0.90
TARGET_ENV = 0.98
TARGET_NEARMISS_PERMONTH = 40
NEARMISS_RATIO_TGT = 10              # near-miss : recordable

THRESH_LTI_GOOD = 30                 # days-since-LTI green threshold
THRESH_LTI_WARN = 7                  # days-since-LTI amber threshold

PM10_LIMIT = 70                      # ug/m3  (EPA Ghana ambient)
PH_MIN = 6.0
PH_MAX = 9.0
WADCN_LIMIT = 50                     # mg/L   (ICMC cyanide code)

# --------------------------------------------------------------------------
# HSE domain enumerations
# --------------------------------------------------------------------------
INCIDENT_TYPES = [
    "Medical Treatment", "First Aid", "Restricted Work", "Lost Time Injury",
    "Property Damage", "Environmental", "Other",
]
RECORDABLE_TYPES = ["Lost Time Injury", "Restricted Work", "Medical Treatment"]
INCIDENT_CLASSES = [
    "Personal Injury", "Occupational Illness", "Fire / Explosion",
    "Chemical / Cyanide Spill", "Vehicle / Mobile Equipment", "Slip / Trip / Fall",
    "Fall from Height", "Caught Between / Struck By", "Electrical",
    "Environmental Release", "Security", "Other",
]
INCIDENT_STATUS = ["Open", "Under Investigation", "Action Pending", "Closed"]
ACTION_STATUS = ["Open", "In Progress", "Due Soon", "Closed"]
PRIORITY = ["Low", "Medium", "High", "Critical"]

# Area -> (Department, dominant Company). Master location list reused everywhere.
AREA_INFO = [
    ("Nkran Open Pit",            "Mining",         "AUMS (Mining Contractor)"),
    ("Esaase Open Pit",           "Mining",         "AUMS (Mining Contractor)"),
    ("Drill & Blast",             "Mining",         "Maxam (Blasting)"),
    ("Haul Roads",                "Mining",         "AUMS (Mining Contractor)"),
    ("Explosives Magazine",       "Mining",         "Maxam (Blasting)"),
    ("Crushing Circuit",          "Processing",     "Owner (Asanko)"),
    ("Processing Plant (CIL)",    "Processing",     "Owner (Asanko)"),
    ("Elution & Gold Room",       "Processing",     "Owner (Asanko)"),
    ("Tailings Storage Facility", "Environment",    "Owner (Asanko)"),
    ("Water Treatment",           "Environment",    "Owner (Asanko)"),
    ("Assay Laboratory",          "Laboratory",     "SGS (Laboratory)"),
    ("HME Workshop",              "Engineering",    "Owner (Asanko)"),
    ("Fuel Farm",                 "Engineering",    "Owner (Asanko)"),
    ("Power Station",             "Engineering",    "Genser (Power)"),
    ("Warehouse & Stores",        "Logistics",      "Owner (Asanko)"),
    ("Administration",            "Administration", "Owner (Asanko)"),
    ("Accommodation Camp",        "Administration", "Catering Co (Camp)"),
    ("Security Gatehouse",        "Security",       "G4S (Security)"),
]
AREAS = [a[0] for a in AREA_INFO]
DEPARTMENTS = sorted({a[1] for a in AREA_INFO})
COMPANIES = ["Owner (Asanko)", "AUMS (Mining Contractor)", "Maxam (Blasting)",
             "SGS (Laboratory)", "Genser (Power)", "Catering Co (Camp)", "G4S (Security)"]
AREA_DEPT = {a[0]: a[1] for a in AREA_INFO}
AREA_COMPANY = {a[0]: a[2] for a in AREA_INFO}

OWNERS = ["K. Mensah", "A. Owusu", "J. Boateng", "E. Asante", "P. Annan",
          "Y. Darko", "S. Addo", "M. Quaye", "F. Agyeman", "D. Tetteh"]

CONTRACTOR_SCOPE = {
    "Owner (Asanko)": "Process plant, admin, environment",
    "AUMS (Mining Contractor)": "Load & haul, drilling, pit operations",
    "Maxam (Blasting)": "Explosives supply & blasting",
    "SGS (Laboratory)": "Assay & sample preparation",
    "Genser (Power)": "Power generation & distribution",
    "Catering Co (Camp)": "Catering & camp services",
    "G4S (Security)": "Site security & access control",
}

# Months of dummy history to seed (trailing window ending this month).
MONTHS_OF_HISTORY = 24
RNG_SEED = 42

# The Excel files the app knows how to consolidate from DATA_DIR.
DATASETS = {
    "incidents":     {"file": "incident_register.xlsx",  "sheet": "Incidents"},
    "activity":      {"file": "activity_log.xlsx",        "sheet": "Activity"},
    "actions":       {"file": "corrective_actions.xlsx",  "sheet": "Actions"},
    "compliance":    {"file": "compliance.xlsx",          "sheet": "Compliance"},
    "environmental": {"file": "environmental.xlsx",       "sheet": "Environmental"},
    "permits":       {"file": "permits.xlsx",             "sheet": "Permits"},
    "audits":        {"file": "audits.xlsx",              "sheet": "Audits"},
    "equipment":     {"file": "equipment.xlsx",           "sheet": "Equipment"},
}
