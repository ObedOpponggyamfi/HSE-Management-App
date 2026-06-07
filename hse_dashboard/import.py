from __future__ import annotations

import argparse
import sys

from .database import DEFAULT_SITE_ID, make_engine
from .importers import import_spreadsheets, import_workbook, results_to_json


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Import HSE spreadsheet data into the local store.")
    parser.add_argument("--db", default=None, help="SQLite database path; defaults to instance/hse_dashboard.sqlite")
    sub = parser.add_subparsers(dest="command", required=True)

    spreadsheets = sub.add_parser("spreadsheets", help="Import the configured spreadsheet set.")
    spreadsheets.add_argument("--site", default=DEFAULT_SITE_ID, help="Site identifier")
    spreadsheets.add_argument("--path", default="data", help="Directory containing the configured workbooks")
    spreadsheets.add_argument("--skip-workbook-profiles", action="store_true",
                              help="Skip optional workbook profiles such as data/incidents.xlsx")

    workbook = sub.add_parser("workbook", help="Import one workbook using a named profile.")
    workbook.add_argument("--site", default=DEFAULT_SITE_ID, help="Site identifier")
    workbook.add_argument("--file", required=True, help="Workbook path")
    workbook.add_argument("--profile", required=True, help="Import profile, e.g. asanko_incidents_v1")

    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    engine = make_engine(args.db)
    if args.command == "spreadsheets":
        results = import_spreadsheets(
            site_id=args.site,
            data_dir=args.path,
            engine=engine,
            include_workbooks=not args.skip_workbook_profiles,
        )
    elif args.command == "workbook":
        results = import_workbook(
            site_id=args.site,
            path=args.file,
            profile=args.profile,
            engine=engine,
        )
    else:
        raise AssertionError(args.command)
    print(results_to_json(results))
    return 0


if __name__ == "__main__":
    sys.exit(main())
