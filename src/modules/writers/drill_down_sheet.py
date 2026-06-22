"""
Foglio drill-down "xLori & Costa({abbrev})".
Aggiunge il foglio a un workbook esistente (openpyxl).
"""

from pathlib import Path

import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# Abbreviazioni case type
CASE_TYPE_ABBREVS: dict[str, str] = {
    "Rates & Inventory Changes":           "R&I",
    "Booking Information":                 "B.Info",
    "Room Type/Rate Plan":                 "RT-RP",
    "Partner Central Access":              "PCA",
    "Content Update":                      "CU",
    "EVC":                                 "EVC",
    "Expedia Collect Invoice":             "ECI",
    "Hotel Collect Issue":                 "HCI",
    "Refund Request":                      "RefReq",
    "Supplier Initiated Traveler Contact": "SITC",
    "Supplier Initiated Relocation":       "SIR",
    "Property Settings":                   "PropSet",
    "Booking Research (Rates)":            "BR(R)",
    "Connectivity Questions":              "ConnQ",
    "Property Details":                    "PropDet",
    "Customer Review Removal":             "CRR",
    "Pre-Onboarding":                      "Pre-Onb",
    "Hotel Closure Outreach":              "HCO",
    "Live Site Investigation":             "LSI",
    "Promotion":                           "Promo",
}

_THIN  = Side(style="thin", color="BFBFBF")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_BLUE  = "4472C4"
_LIGHT = "EBF3FB"

# Caratteri vietati nei nomi dei fogli Excel: [ ] : * ? / \
_INVALID_SHEET_CHARS = str.maketrans({c: "-" for c in r"[]:*?/\\"})


def _safe_sheet_name(name: str) -> str:
    """Rende un nome valido come titolo di foglio Excel (no caratteri vietati, max 31)."""
    cleaned = name.translate(_INVALID_SHEET_CHARS).strip()
    return cleaned[:31] or "Sheet"


def _c(ws, row, col, value=None, *, bold=False, bg=None, align="center", fmt=None):
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = Font(bold=bold, name="Calibri", size=10)
    cell.alignment = Alignment(horizontal=align, vertical="center")
    cell.border = _BORDER
    if bg:
        cell.fill = PatternFill("solid", fgColor=bg)
    if fmt:
        cell.number_format = fmt
    return cell


def add_drill_down_sheet(xlsx_path: Path, df_full: pd.DataFrame,
                          case_type: str) -> str:
    """
    Apre il workbook esistente con openpyxl e aggiunge il foglio drill-down.
    Ritorna (nome_foglio, numero_righe).
    """
    abbrev     = CASE_TYPE_ABBREVS.get(case_type, case_type[:6])
    sheet_name = _safe_sheet_name(f"xLori & Costa({abbrev})")

    wb = openpyxl.load_workbook(xlsx_path)

    # Rimuovi il foglio se già esiste (per rigenerare)
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]

    ws = wb.create_sheet(sheet_name)

    # Filtra e ordina
    sub = df_full[df_full["Case Type"] == case_type].copy()
    sub = sub.sort_values("aht", ascending=False).reset_index(drop=True)

    # ── Intestazione ──────────────────────────────────────────────────────────
    headers = ["case_number", "Case Type", "Primary Category", "Employee Name",
               "Case Origin (group)", "aht", "Scenario", "BUCKETING", "Real call AHT"]
    for col_i, h in enumerate(headers, start=1):
        _c(ws, 1, col_i, h, bold=True, bg=_BLUE, align="center")
        ws.cell(row=1, column=col_i).font = Font(
            bold=True, color="FFFFFF", name="Calibri", size=10)

    # ── Dati ─────────────────────────────────────────────────────────────────
    source_cols = {
        "case_number":          "case_number",
        "Case Type":            "Case Type",
        "Primary Category":     "Primary Category",
        "Employee Name":        "Employee Name",
        "Case Origin (group)":  "Case Origin (group)",
        "aht":                  "aht",
    }
    for r_i, row_data in sub.iterrows():
        xl_row = r_i + 2
        bg_row = _LIGHT if r_i % 2 == 0 else None
        for col_i, (_, src_col) in enumerate(source_cols.items(), start=1):
            val = row_data.get(src_col, None)
            if col_i == 6 and val is not None:  # aht: numero
                _c(ws, xl_row, col_i, float(val) if val else None,
                   bg=bg_row, align="center", fmt="0.0")
            elif col_i == 1 and val is not None:  # case_number
                _c(ws, xl_row, col_i, str(int(val)) if val else "",
                   bg=bg_row, align="center")
            else:
                _c(ws, xl_row, col_i, str(val) if val is not None else "",
                   bg=bg_row, align="left" if col_i in (3, 4) else "center")
        # Colonne vuote: Scenario, BUCKETING, Real call AHT
        for col_i in range(7, 10):
            _c(ws, xl_row, col_i, None, bg=bg_row)

    # ── Larghezze ─────────────────────────────────────────────────────────────
    col_widths = [14, 24, 22, 24, 20, 8, 12, 14, 14]
    for col_i, w in enumerate(col_widths, start=1):
        ws.column_dimensions[ws.cell(1, col_i).column_letter].width = w
    ws.row_dimensions[1].height = 16

    wb.save(xlsx_path)
    return sheet_name, len(sub)
