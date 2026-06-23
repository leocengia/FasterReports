"""
Fogli WoW DeepDive: WoW Pivot, WoW Monthly, tblClass, WoW Export.
"""

import pandas as pd
from ..classification import build_tbl_class
from ..data_loader import CHANNEL_CONFIG, weighted_aht
from .styles import make_formats, C_BLUE_HEADER, C_BLUE_LIGHT, C_YELLOW_LIGHT, CLASSIFICATION_COLORS


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
    Garantisce che nessun case type reale venga omesso dai fogli WoW.
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


def write_wow_monthly(wb, df: pd.DataFrame) -> None:
    """WoW Monthly: per ogni canale × case type → metriche."""
    ws   = wb.add_worksheet("WoW Monthly")
    fmts = make_formats(wb)
    _write_monthly_table(ws, fmts, df)


def _write_monthly_table(ws, fmts, df: pd.DataFrame, start_row: int = 0):
    """Scrive le tabelle Phone e Non-live affiancate."""
    headers = ["case_type", "aht (min)", "volume", "impact (%)", "out_of_target (%)", "std_dev", "p90-p10"]

    sections = [
        ("Phone", "Phone", 0, C_YELLOW_LIGHT),
        ("Non-live", "Non-live", 10, C_BLUE_LIGHT),
    ]

    case_types = _case_type_order(df)
    row0 = start_row
    for label, channel, start_col, bg in sections:
        ws.merge_range(row0, start_col, row0, start_col + len(headers) - 1,
                       f"[ {label} ]", fmts["header"])
        for i, h in enumerate(headers):
            ws.write(row0 + 1, start_col + i, h, fmts["header"])

        cfg = CHANNEL_CONFIG[channel]
        target = cfg["target_aht"]
        sub = df[df["channel"] == channel]
        total_min = float((sub["aht"] * sub["cases"]).sum()) if not sub.empty else 1.0

        row = row0 + 2
        for ct in case_types:
            grp = sub[sub["Case Type"] == ct]
            ws.write(row, start_col, ct, fmts["data_left"])
            if grp.empty:
                for c in range(1, len(headers)):
                    ws.write(row, start_col + c, "N/A", fmts["data"])
            else:
                vol     = int(grp["cases"].sum())
                total   = float((grp["aht"] * grp["cases"]).sum())
                aht_v   = total / vol if vol else 0.0
                impact  = total / total_min if total_min else 0.0
                oot     = float((grp["aht"] > target).sum()) / vol if vol else 0.0
                std     = float(grp["aht"].std()) if len(grp) > 1 else 0.0
                p90p10  = float(grp["aht"].quantile(0.9) - grp["aht"].quantile(0.1))
                ws.write_number(row, start_col + 1, aht_v,  fmts["data_num"])
                ws.write_number(row, start_col + 2, vol,    fmts["data_int"])
                ws.write_number(row, start_col + 3, impact, fmts["data_pct"])
                ws.write_number(row, start_col + 4, oot,    fmts["data_pct"])
                ws.write_number(row, start_col + 5, std,    fmts["data_num"])
                ws.write_number(row, start_col + 6, p90p10, fmts["data_num"])
            row += 1

    ws.set_column(0, 0, 30)
    ws.set_column(1, 6, 12)
    ws.set_column(10, 10, 30)
    ws.set_column(11, 16, 12)


def write_tbl_class(wb, df: pd.DataFrame) -> None:
    """tblClass: tabella master di scoring."""
    ws   = wb.add_worksheet("tblClass")
    fmts = make_formats(wb)
    tbl  = build_tbl_class(df)

    if tbl.empty:
        ws.write(0, 0, "Nessun dato", fmts["data"])
        return

    columns = [
        ("channel",          "Channel",         12, "data_left"),
        ("case_type",        "Case Type",        28, "data_left"),
        ("chiave",           "Chiave",           32, "data_left"),
        ("volume",           "Volume",           8,  "data_int"),
        ("avg_aht",          "AHT",              8,  "data_num"),
        ("total_minutes",    "Total Min",        10, "data_num"),
        ("impact_channel",   "Impact",           8,  "data_pct"),
        ("out_of_target",    "OOT %",            8,  "data_pct"),
        ("std_dev",          "StdDev",           8,  "data_num"),
        ("p90p10",           "P90-P10",          8,  "data_num"),
        ("target_aht",       "Target AHT",       10, "data_num"),
        ("gap_vs_target",    "Gap vs Target",    12, "data_num"),
        ("opp_minutes",      "Opp. Min",         10, "data_num"),
        ("std_dev_ratio",    "StdDev Ratio",     10, "data_num"),
        ("p90p10_ratio",     "P90P10 Ratio",     10, "data_num"),
        ("std_dev_rank",     "StdDev Rank",      10, "data_num"),
        ("p90p10_rank",      "P90P10 Rank",      10, "data_num"),
        ("variability_score","Var. Score",       10, "data_num"),
        ("impact_rank",      "Impact Rank",      10, "data_num"),
        ("volume_rank",      "Volume Rank",      10, "data_num"),
        ("classification",   "Classification",   18, "data"),
    ]

    for col_idx, (_, label, width, _fmt) in enumerate(columns):
        ws.write(0, col_idx, label, fmts["header"])
        ws.set_column(col_idx, col_idx, width)

    cls_fmts = {
        "Highly Actionable": fmts["ha"],
        "Actionable":        fmts["act"],
        "Process Driven":    fmts["pd"],
    }

    for row_idx, data_row in tbl.iterrows():
        cls = data_row.get("classification", "Process Driven")
        row_fmt = cls_fmts.get(cls, fmts["data"])
        row_num_fmt = wb.add_format({
            "font_name": "Calibri", "font_size": 10, "border": 1,
            "bg_color": CLASSIFICATION_COLORS.get(cls, "#FFFFFF"),
            "align": "center", "valign": "vcenter", "num_format": "0.0000",
        })
        for col_idx, (field, _, _, fmt_key) in enumerate(columns):
            val = data_row.get(field, "")
            if val is None or (isinstance(val, float) and pd.isna(val)):
                ws.write(row_idx + 1, col_idx, "N/A", row_fmt)
            elif fmt_key in ("data_num", "data_pct"):
                ws.write_number(row_idx + 1, col_idx, float(val), row_num_fmt)
            elif fmt_key == "data_int":
                ws.write_number(row_idx + 1, col_idx, int(val), row_fmt)
            else:
                ws.write(row_idx + 1, col_idx, str(val), row_fmt)

    ws.freeze_panes(1, 0)


def write_wow_export(wb, df: pd.DataFrame) -> None:
    """WoW Export: Phone + Non-live side by side."""
    ws   = wb.add_worksheet("WoW Export")
    fmts = make_formats(wb)

    headers_ch = ["aht (min)", "volume", "impact (%)", "out_of_target (%)", "std_dev", "p90-p10"]

    # Headers
    ws.write(0, 0, "case_type", fmts["header"])
    ws.merge_range(0, 1, 0, len(headers_ch), "Phone", fmts["header"])
    ws.merge_range(0, len(headers_ch) + 2, 0, len(headers_ch) * 2 + 1, "Non-live", fmts["header"])

    ws.write(1, 0, "", fmts["header"])
    for i, h in enumerate(headers_ch):
        ws.write(1, 1 + i, h, fmts["header"])
        ws.write(1, len(headers_ch) + 2 + i, h, fmts["header"])
    ws.write(1, len(headers_ch) + 1, "", fmts["header"])

    phone_sub = df[df["channel"] == "Phone"]
    nl_sub    = df[df["channel"] == "Non-live"]
    phone_total = float((phone_sub["aht"] * phone_sub["cases"]).sum()) if not phone_sub.empty else 1.0
    nl_total    = float((nl_sub["aht"] * nl_sub["cases"]).sum())       if not nl_sub.empty   else 1.0

    for r, ct in enumerate(_case_type_order(df), start=2):
        row_fmt = fmts["gray"] if r % 2 == 0 else fmts["data"]
        row_num = wb.add_format({
            "font_name": "Calibri", "font_size": 10, "border": 1,
            "bg_color": "#F2F2F2" if r % 2 == 0 else "#FFFFFF",
            "align": "center", "num_format": "0.00",
        })
        ws.write(r, 0, ct, fmts["data_left"] if r % 2 != 0 else fmts["gray_left"])

        for sub, total, start_col in [
            (phone_sub, phone_total, 1),
            (nl_sub, nl_total, len(headers_ch) + 2),
        ]:
            grp = sub[sub["Case Type"] == ct]
            if grp.empty:
                for c in range(len(headers_ch)):
                    ws.write(r, start_col + c, "N/A", row_fmt)
            else:
                vol    = int(grp["cases"].sum())
                aht_v  = weighted_aht(grp)
                impact = (aht_v * vol) / total if total else 0.0
                cfg    = CHANNEL_CONFIG["Phone"] if start_col == 1 else CHANNEL_CONFIG["Non-live"]
                oot    = float((grp["aht"] > cfg["target_aht"]).sum()) / vol if vol else 0.0
                std    = float(grp["aht"].std()) if len(grp) > 1 else 0.0
                p90p10 = float(grp["aht"].quantile(0.9) - grp["aht"].quantile(0.1))
                for c, v in enumerate([aht_v, vol, impact, oot, std, p90p10]):
                    ws.write_number(r, start_col + c, v, row_num)

    ws.set_column(0, 0, 30)
    ws.set_column(1, len(headers_ch) * 2 + 2, 12)
    ws.freeze_panes(2, 1)
