"""
Exporter – converts a list of CaseEntry objects to CSV or XLSX bytes.
"""

from __future__ import annotations

import io
from typing import List

import pandas as pd

from app.services.parser import CaseEntry

COLUMNS = ["Case", "Topic", "Subtopic", "Sub-subtopic", "Further Division"]


def entries_to_dataframe(entries: List[CaseEntry]) -> pd.DataFrame:
    rows = [e.as_dict() for e in entries]
    if not rows:
        return pd.DataFrame(columns=COLUMNS)
    return pd.DataFrame(rows, columns=COLUMNS)


def to_csv(entries: List[CaseEntry]) -> bytes:
    df = entries_to_dataframe(entries)
    return df.to_csv(index=False).encode("utf-8")


def to_xlsx(entries: List[CaseEntry]) -> bytes:
    df = entries_to_dataframe(entries)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Cases")
    return buf.getvalue()
