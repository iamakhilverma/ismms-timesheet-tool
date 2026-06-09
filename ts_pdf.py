#!/usr/bin/env python3
"""PDF fill logic for the Daily Attendance Record (config-driven).

Row layout (verified against the sample sheet):
    row 1 = Sunday, 2 = Monday, ... 6 = Friday, 7 = Saturday
Per-row fields: "ABSENCE CODE_n", "TIME IN_n", "TIME OUT_n", "DAILY HOURS_n"
Period: Month_From/Day_From/Year_From .. Month_To/Day_To/Year_To
Header: "Name", "Life #"   Footer: "Total Hours"
"""
import datetime as dt
import logging
from pathlib import Path

from pypdf import PdfReader, PdfWriter

# The template has a couple of stale object refs; pypdf reads it fine but is noisy.
logging.getLogger("pypdf").setLevel(logging.ERROR)

HERE = Path(__file__).resolve().parent
TEMPLATE = HERE / "template.pdf"


def _fmt_hours(total: float) -> str:
    """37.5 -> '37.5', 30.0 -> '30'."""
    return str(int(total)) if total == int(total) else (f"{total:.1f}").rstrip("0").rstrip(".")


def day_values(day: dt.date, override, cfg):
    """Return (code, time_in, time_out, daily_hours_or_None) for one day.

    `override`: None, or a dict like {"kind": "holiday"} / {"kind": "pto",
    "code": "PTOS"|"PTOU", "paid": bool}. CLI flags supply overrides.
    """
    full = float(cfg["daily_hours"])
    if override:
        if override["kind"] == "holiday":
            return (cfg["holiday_code"], "", "", full)
        if override["kind"] == "pto":
            return (override["code"], "", "", full if override.get("paid", True) else None)
    if day.weekday() <= 4:                       # Mon..Fri
        return ("", cfg["time_in"], cfg["time_out"], full)
    return ("", "", "", None)                    # weekend


def build_fields(week_start: dt.date, overrides: dict, cfg: dict):
    """overrides: {date -> override dict}. Returns (fields, week_end, total)."""
    assert week_start.weekday() == 6, "week_start must be a Sunday"
    week_end = week_start + dt.timedelta(days=6)
    fields = {"Name": cfg["name"], "Life #": cfg["life_num"]}
    total = 0.0
    for i in range(7):
        day = week_start + dt.timedelta(days=i)
        n = i + 1
        code, tin, tout, hours = day_values(day, overrides.get(day), cfg)
        fields[f"Month_{n}"] = str(day.month)
        fields[f"Day_{n}"] = str(day.day)
        fields[f"ABSENCE CODE_{n}"] = code
        fields[f"TIME IN_{n}"] = tin
        fields[f"TIME OUT_{n}"] = tout
        if hours is None:
            fields[f"DAILY HOURS_{n}"] = ""
        else:
            fields[f"DAILY HOURS_{n}"] = cfg["daily_hours"]
            total += hours
    fields["Month_From"], fields["Day_From"], fields["Year_From"] = (
        str(week_start.month), str(week_start.day), str(week_start.year))
    fields["Month_To"], fields["Day_To"], fields["Year_To"] = (
        str(week_end.month), str(week_end.day), str(week_end.year))
    fields["Total Hours"] = _fmt_hours(total)
    return fields, week_end, total


def fill(week_start: dt.date, out_path: Path, overrides: dict, cfg: dict):
    from pypdf.generic import NameObject, TextStringObject
    writer = PdfWriter(clone_from=str(TEMPLATE))
    fields, week_end, total = build_fields(week_start, overrides, cfg)

    # 1. pypdf update: sets new /V AND regenerates the appearance (/AP) on the page
    #    widget annotations -- this is what Preview / Quick Look (PDFKit) render.
    page_annot_ids = set()
    for page in writer.pages:
        writer.update_page_form_field_values(page, fields, auto_regenerate=True)
        for ref in (page.get("/Annots", []) or []):
            try:
                page_annot_ids.add(ref.idnum)
            except Exception:
                pass

    # 2. This template is malformed: the AcroForm /Fields are a SECOND set of objects,
    #    separate from the page annotations. pypdf doesn't touch them, so Acrobat (which
    #    reads the AcroForm values) shows stale data. Force their /V, and clear THEIR
    #    stale /AP so Acrobat rebuilds it from /V (these objects aren't drawn by PDFKit,
    #    so clearing /AP here can't blank Preview).
    acro = writer._root_object.get("/AcroForm")
    for ref in (acro.get_object().get("/Fields", []) if acro else []):
        try:
            obj = ref.get_object()
            name = obj.get("/T")
        except Exception:
            continue
        if name is None or str(name) not in fields:
            continue
        obj[NameObject("/V")] = TextStringObject(fields[str(name)])
        if ref.idnum not in page_annot_ids and "/AP" in obj:
            del obj[NameObject("/AP")]

    # 3. Ask viewers to (re)build any appearances they don't already have.
    try:
        writer.set_need_appearances_writer(True)
    except Exception:
        pass

    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with open(out_path, "wb") as fh:
        writer.write(fh)
    return fields, week_end, total
