"""
Foglio WoW Pivot (analisi week-over-week per data) + ordine case type condiviso.

NB: le tabelle di scoring/deepdive (Weekly Deepdive, tblClass, EXPORT) sono in
`deepdive_sheets.py`; qui resta solo l'analisi temporale WoW Pivot.
"""

import pandas as pd
from ..data_loader import weighted_aht
from .styles import make_formats


# Ordine fisso dei case type (dall'alfabetico del template)
_TEMPLATE_CASE_TYPES = [
    "Account Takeover", "Booking Information", "Booking Research (Rates)",
    "Bulk Update Request", "Connectivity Questions", "Contact Update",
    "Content Update", "Contract Update", "Customer Review Removal", "EVC",
    "Expedia Collect Invoice", "Hotel Closure Outreach", "Hotel Collect Issue",
    "Inventory Closeout", "Known Tool Outage Issue", "Live Site Investigation",
    "New Contract", "Partner Central Access", "Pre-Onboarding", "Promotion",
    "Property Classification", "Property Details", "Property Photos",
    "Property Settings", "Rates & Inventory Changes", "Refund Request",
    "Room Type/Rate Plan", "Special Check-In Instructions", "Specialty Functions",
    "Supplier Initiated Relocation", "Supplier Initiated Traveler Contact",
    "Traveler Outreach",
]


def _case_type_order(df: pd.DataFrame) -> list[str]:
    """
    Ordine dei case type: prima quelli del template (ordine fisso), poi eventuali
    case type presenti nei dati ma non nel template, in coda in ordine alfabetico.
    """
    present = set(df["Case Type"].dropna().unique())
    extra = sorted(present - set(_TEMPLATE_CASE_TYPES))
    return _TEMPLATE_CASE_TYPES + extra


def write_wow_pivot(wb, df: pd.DataFrame) -> None:
    """WoW Pivot: Case Origin (group) × date → aht, volume, std_dev."""
    ws    = wb.add_worksheet("WoW Pivot")
    fmts  = make_formats(wb)

    dates = sorted(df["Date (Range)"].dropna().unique()) if "Date (Range)" in df.columns else []
    channels = ["Phone", "Other"]

    row = 0
    # Titolo
    ws.merge_range(row, 0, row, max(1, len(dates) * 3), "WoW Pivot — Case Origin × Date", fmts["title"])
    row += 2

    # Sub-header date
    ws.write(row, 0, "Case Origin (group)", fmts["header"])
    col = 1
    for d in dates:
        ws.merge_range(row, col, row, col + 2, str(d)[:10], fmts["header"])
        col += 3
    row += 1
    ws.write(row, 0, "", fmts["header"])
    col = 1
    for _ in dates:
        ws.write(row, col,     "aht (min)", fmts["header_wrap"])
        ws.write(row, col + 1, "volume",    fmts["header_wrap"])
        ws.write(row, col + 2, "std_dev",   fmts["header_wrap"])
        col += 3
    row += 1

    for ch in channels:
        sub = df[df["Case Origin (group)"] == ch]
        ws.write(row, 0, ch, fmts["data_left"])
        col = 1
        for d in dates:
            day = sub[sub["Date (Range)"] == d] if "Date (Range)" in sub.columns else pd.DataFrame()
            if day.empty:
                ws.write(row, col, "N/A", fmts["data"])
                ws.write(row, col + 1, 0, fmts["data_int"])
                ws.write(row, col + 2, "N/A", fmts["data"])
            else:
                aht_val = weighted_aht(day)
                vol_val = int(day["cases"].sum())
                std_val = float(day["aht"].std()) if len(day) > 1 else 0.0
                ws.write_number(row, col,     aht_val, fmts["data_num"])
                ws.write_number(row, col + 1, vol_val, fmts["data_int"])
                ws.write_number(row, col + 2, std_val, fmts["data_num"])
            col += 3
        row += 1

    ws.set_column(0, 0, 22)
    ws.set_column(1, 1 + len(dates) * 3, 10)
