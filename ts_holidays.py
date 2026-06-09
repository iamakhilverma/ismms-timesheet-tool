#!/usr/bin/env python3
"""Extract holiday dates from a holiday-schedule PDF (e.g. the MSH 2026 schedule).

The schedule lists, per holiday, three dates in order: the actual date, the date
the holiday is *observed*, and the premium pay day. We default to the **observed**
date because that's the day you don't work (e.g. July 4 is a Saturday but observed
Friday 07/03). Returns ISO date strings (YYYY-MM-DD).
"""
import datetime as dt
import re
from pathlib import Path

from pypdf import PdfReader

_DATE = re.compile(r"(\d{2})/(\d{2})/(\d{4})")
_COLS = {"actual": 0, "observed": 1, "payday": 2}


def extract_holidays(pdf_path, which: str = "observed"):
    if which not in _COLS:
        raise ValueError(f"which must be one of {list(_COLS)}")
    text = "".join((pg.extract_text() or "") for pg in PdfReader(str(pdf_path)).pages)
    dates = [f"{y}-{m}-{d}" for (m, d, y) in _DATE.findall(text)]
    if len(dates) % 3 != 0:
        # not clean triples -> fall back to all valid unique dates
        picked = dates
    else:
        idx = _COLS[which]
        picked = [dates[i] for i in range(idx, len(dates), 3)]
    out = []
    for iso in picked:
        try:
            dt.date.fromisoformat(iso)
            out.append(iso)
        except ValueError:
            pass
    return sorted(set(out))
