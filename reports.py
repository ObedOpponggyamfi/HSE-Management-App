#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Excel export helpers (one dataset, or a full multi-sheet workbook)."""
import io

import pandas as pd


def export_dataframe(df: pd.DataFrame, sheet="Data") -> io.BytesIO:
    bio = io.BytesIO()
    out = df if (df is not None and not df.empty) else pd.DataFrame({"info": ["no data"]})
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        out.to_excel(writer, index=False, sheet_name=str(sheet)[:31])
    bio.seek(0)
    return bio


def export_workbook(frames: dict) -> io.BytesIO:
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        wrote = False
        for key, df in frames.items():
            if df is None or df.empty:
                continue
            df.to_excel(writer, index=False, sheet_name=str(key)[:31])
            wrote = True
        if not wrote:
            pd.DataFrame({"info": ["no data"]}).to_excel(writer, index=False, sheet_name="info")
    bio.seek(0)
    return bio
