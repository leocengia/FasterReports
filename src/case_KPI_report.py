"""
Genera case_KPI_report.xlsx a partire da un export completo dei casi.

Input:  raw_data/<filename>.csv  (separatore ;, decimale ,)
        Colonne attese: Case Type, Primary Category, Case AHT (mins),
                        Distinct Cases, Case Origin (group) o Case Channel,
                        Employee Name (opzionale)
Output: output/case_KPI_report_W<settimana>.xlsx

Uso:
    python src/case_KPI_report.py "raw_data/dataset.csv" --week 24
    python src/case_KPI_report.py "raw_data/dataset.csv" --week 24 --target-phone 19.98 --target-nonlive 18.96
"""

import argparse
import sys
from pathlib import Path

import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference
from openpyxl.drawing.spreadsheet_drawing import TwoCellAnchor, AnchorMarker

# ── Parametri ─────────────────────────────────────────────────────────────────
DEFAULT_TGT_PHONE = 19.98
DEFAULT_TGT_NL    = 18.96
TOP_N_CASETYPE    = 20
TOP_N_PRIMCAT     = 15

NON_LIVE_CHANNELS = {"Contact Us", "Market Management", "Others", "Email"}
PHONE_CHANNELS    = {"Phone"}

# ── Colori ────────────────────────────────────────────────────────────────────
C_HDR_BG  = "000099"   # blu scuro intestazioni
C_HDR_FG  = "FFFFFF"   # bianco testo
C_ROW_ALT = "D9E1F2"   # azzurro chiaro righe alternate
C_ACT_BG  = "FFC000"   # giallo — Actionable
C_HA_BG   = "C00000"   # rosso scuro — Highly Actionable
C_HA_FG   = "FFFFFF"   # bianco su rosso scuro

THIN   = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


# ── Helpers ───────────────────────────────────────────────────────────────────

def detect_col(df, *candidates):
    for name in candidates:
        if name in df.columns:
            return name
    return None


def _cell(ws, row, col, value=None, *, bold=False, size=11,
          align="center", fmt=None, bg=None, fg="000000"):
    c = ws.cell(row=row, column=col, value=value)
    c.font      = Font(bold=bold, name="Calibri", size=size, color=fg)
    c.alignment = Alignment(horizontal=align, vertical="center")
    c.border    = BORDER
    if bg:
        c.fill = PatternFill("solid", fgColor=bg)
    if fmt:
        c.number_format = fmt
    return c


def weighted_aht(group, aht_col, cases_col):
    total = group[cases_col].sum()
    if total == 0:
        return 0.0
    return (group[aht_col] * group[cases_col]).sum() / total


# ── Aggregazioni ──────────────────────────────────────────────────────────────

def aggregate(df, aht_col, cases_col, group_col, target, top_n):
    agg = (
        df.groupby(group_col)
        .apply(lambda g: pd.Series({
            "avg_aht": weighted_aht(g, aht_col, cases_col),
            "volume":  g[cases_col].sum(),
        }), include_groups=False)
        .reset_index()
        .rename(columns={group_col: "label"})
        .sort_values("volume", ascending=False)
        .head(top_n)
        .reset_index(drop=True)
    )
    agg["target"] = target
    return agg


# ── Scrittura foglio Case Type Analysis ───────────────────────────────────────

def write_analysis_sheet(ws, case_types, primary_cats, channel_label="(Tutto)"):

    def _write_table(data, start_col, headers, start_row=1):
        """Scrive intestazioni + dati, restituisce n righe dati."""
        for i, h in enumerate(headers):
            _cell(ws, start_row, start_col + i, h,
                  bold=True, bg=C_HDR_BG, fg=C_HDR_FG)
        for row_i, row in data.iterrows():
            xl_row = start_row + 1 + row_i
            bg = C_ROW_ALT if row_i % 2 == 0 else None
            vals  = list(row)
            fmts  = ["", "0.00", "0", "0.00"]
            aligns = ["left"] + ["center"] * (len(headers) - 1)
            for c_i, (val, fmt, al) in enumerate(zip(vals, fmts, aligns)):
                _cell(ws, xl_row, start_col + c_i, val,
                      bg=bg, fmt=fmt if fmt else None, align=al)
        return len(data)

    # ── A) Left pivot (colonne A=1 – C=3) ────────────────────────────────────
    _cell(ws, 2, 1, "Case Origin (group)", align="left")
    _cell(ws, 2, 2, channel_label, align="center")
    _cell(ws, 3, 1, "case_number", align="left")
    _cell(ws, 3, 2, "(Tutto)", align="center")

    # intestazioni pivot
    _cell(ws, 5, 1, "Etichette di riga", align="left")
    _cell(ws, 5, 2, "aht", align="center")
    _cell(ws, 5, 3, "volume", align="center")

    for i, row in case_types.iterrows():
        xl_row = 6 + i
        bg = C_ROW_ALT if i % 2 == 0 else None
        _cell(ws, xl_row, 1, row["label"],          align="left",   bg=bg)
        _cell(ws, xl_row, 2, round(row["avg_aht"], 2), fmt="0.00", bg=bg)
        _cell(ws, xl_row, 3, int(row["volume"]),    align="center", bg=bg)

    # ── B) Actionable / Highly Actionable labels (colonne L=12 – N=14) ───────
    ws.merge_cells("L2:N2")
    _cell(ws, 2, 12, "Actionable", bold=True, bg=C_ACT_BG, fg="000000")

    ws.merge_cells("L4:N4")
    _cell(ws, 4, 12, "Highly Actionable", bold=True, bg=C_HA_BG, fg=C_HA_FG)

    # ── C) Case Type table (colonne G=7 – J=10) ───────────────────────────────
    n_ct = _write_table(case_types[["label", "avg_aht", "volume", "target"]],
                        start_col=7,
                        headers=["case_type", "avg_aht", "volume", "target"])

    # ── D) Primary Category table (colonne P=16 – S=19) ──────────────────────
    n_pc = 0
    if not primary_cats.empty:
        n_pc = _write_table(primary_cats[["label", "avg_aht", "volume", "target"]],
                            start_col=16,
                            headers=["primary_category", "avg_aht", "volume", "target"])

    # ── E) Larghezze colonne ──────────────────────────────────────────────────
    col_widths = {
        "A": 35.14, "B": 9.43,  "C": 7.71,
        "G": 32.0,  "H": 7.86,  "I": 7.71, "J": 6.29,
        "K": 4.0,
        "L": 13.14,
        "O": 4.43,
        "P": 27.0,  "Q": 7.86,  "R": 7.71, "S": 6.29,
    }
    for col, w in col_widths.items():
        ws.column_dimensions[col].width = w

    # ── F) Altezze righe ──────────────────────────────────────────────────────
    for r in (1, 2, 3, 4, 8, 10):
        ws.row_dimensions[r].height = 15.75

    # ── G) Grafico 1: Avg. AHT by Case Type ───────────────────────────────────
    bc1 = BarChart()
    bc1.type         = "bar"
    bc1.grouping     = "clustered"
    bc1.title        = "Avg. AHT by Case Type"
    bc1.style        = 10
    bc1.x_axis.title = "AHT (mins)"
    bc1.y_axis.title = "Case Type"

    bc1.add_data(Reference(ws, min_col=8, min_row=1, max_row=1 + n_ct),
                 titles_from_data=True)
    bc1.set_categories(Reference(ws, min_col=7, min_row=2, max_row=1 + n_ct))

    a1        = TwoCellAnchor()
    a1._from  = AnchorMarker(col=5,  colOff=438149, row=8,  rowOff=133349)
    a1.to     = AnchorMarker(col=14, colOff=180975, row=34, rowOff=66675)
    bc1.anchor = a1
    ws.add_chart(bc1)

    # ── H) Grafico 2: Avg. AHT by Primary Category ────────────────────────────
    if not primary_cats.empty:
        bc2 = BarChart()
        bc2.type         = "bar"
        bc2.grouping     = "clustered"
        bc2.title        = "Avg. AHT by Primary Category"
        bc2.style        = 10
        bc2.x_axis.title = "AHT (mins)"
        bc2.y_axis.title = "Primary Category"

        bc2.add_data(Reference(ws, min_col=17, min_row=1, max_row=1 + n_pc),
                     titles_from_data=True)
        bc2.set_categories(Reference(ws, min_col=16, min_row=2, max_row=1 + n_pc))

        a2        = TwoCellAnchor()
        a2._from  = AnchorMarker(col=15, colOff=19050,  row=10, rowOff=85725)
        a2.to     = AnchorMarker(col=23, colOff=390525, row=36, rowOff=161925)
        bc2.anchor = a2
        ws.add_chart(bc2)


# ── Scrittura foglio xLori & Costa (outlier AHT) ──────────────────────────────

def write_xloricrosta_sheet(ws, df, aht_col, cas_col, typ_col, cat_col,
                             ch_col, emp_col, num_col,
                             target_phone, target_nonlive):
    # Determina target per-riga in base al canale
    if ch_col:
        df = df.copy()
        df["_target"] = df[ch_col].apply(
            lambda ch: target_phone if ch == "Phone" else target_nonlive
        )
        outliers = df[df[aht_col] > df["_target"]].copy()
    else:
        outliers = df[df[aht_col] > target_phone].copy()

    outliers = outliers.sort_values(aht_col, ascending=False)

    # Costruisce le colonne da scrivere (solo quelle disponibili)
    cols_wanted = []
    if num_col:
        cols_wanted.append((num_col, "case_number", 12.86))
    cols_wanted.append((typ_col, "Case Type",             25.0))
    if cat_col:
        cols_wanted.append((cat_col, "Primary Category",  25.71))
    if emp_col:
        cols_wanted.append((emp_col, "Employee Name",     18.0))
    if ch_col:
        cols_wanted.append((ch_col,  "Case Origin (group)", 18.29))
    cols_wanted.append((aht_col, "aht", 6.0))

    # Intestazioni
    hdr_font = Font(name="Calibri", size=11, bold=True)
    for c_idx, (_, hdr, _) in enumerate(cols_wanted, 1):
        cell = ws.cell(row=1, column=c_idx, value=hdr)
        cell.font = hdr_font

    # Dati
    for r_idx, (_, row_data) in enumerate(outliers.iterrows(), 2):
        for c_idx, (src_col, _, _) in enumerate(cols_wanted, 1):
            val = row_data[src_col]
            ws.cell(row=r_idx, column=c_idx, value=val)

    # Larghezze
    for c_idx, (_, _, w) in enumerate(cols_wanted, 1):
        ws.column_dimensions[get_column_letter(c_idx)].width = w


# ── Scrittura DATASET ─────────────────────────────────────────────────────────

def write_dataset_sheet(ws, df):
    hdr_font = Font(name="Calibri", size=11, bold=False)
    for c_idx, col_name in enumerate(df.columns, 1):
        cell = ws.cell(row=1, column=c_idx, value=col_name)
        cell.font = hdr_font

    for r_idx, row in enumerate(df.itertuples(index=False), 2):
        for c_idx, val in enumerate(row, 1):
            ws.cell(row=r_idx, column=c_idx, value=val)


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Genera case_KPI_report.xlsx")
    parser.add_argument("input",
                        help="Path al CSV raw (export completo casi)")
    parser.add_argument("--week", default="",
                        help="Numero settimana (es. 24)")
    parser.add_argument("--target-phone", type=float, default=DEFAULT_TGT_PHONE,
                        help=f"Target AHT Phone (default {DEFAULT_TGT_PHONE})")
    parser.add_argument("--target-nonlive", type=float, default=DEFAULT_TGT_NL,
                        help=f"Target AHT Non-live (default {DEFAULT_TGT_NL})")
    parser.add_argument("--channel", default=None,
                        help="Filtra per canale: Phone, Non-live, o vuoto per tutti")
    args = parser.parse_args()

    csv_path = Path(args.input)
    if not csv_path.exists():
        print(f"Errore: file non trovato: {csv_path}", file=sys.stderr)
        sys.exit(1)

    print("Lettura dati...")
    df = pd.read_csv(csv_path, sep=";", encoding="utf-8-sig",
                     decimal=",", low_memory=False)

    # Rileva colonne
    aht_col = detect_col(df, "Case AHT (mins)", "DS", "aht", "Case AHT")
    cas_col = detect_col(df, "Distinct Cases",  "DX", "cases")
    typ_col = detect_col(df, "Case Type",       "Z",  "case_type")
    cat_col = detect_col(df, "Primary Category","BV", "primary_category")
    ch_col  = detect_col(df, "Case Origin (group)", "U", "Case Channel", "S")
    emp_col = detect_col(df, "Employee Name",   "EE", "Agent Name")
    num_col = detect_col(df, "case_number",     "Case Number", "A")

    missing = [n for n, c in [("Case AHT",      aht_col),
                               ("Distinct Cases", cas_col),
                               ("Case Type",     typ_col)] if not c]
    if missing:
        print(f"Errore: colonne obbligatorie non trovate: {missing}", file=sys.stderr)
        print(f"Colonne disponibili: {list(df.columns[:20])}...", file=sys.stderr)
        sys.exit(1)

    df[aht_col] = pd.to_numeric(df[aht_col], errors="coerce").fillna(0)
    df[cas_col] = pd.to_numeric(df[cas_col], errors="coerce").fillna(0)
    df = df[df[cas_col] > 0].copy()

    # Mappatura canale se la colonna è "Case Channel"
    if ch_col and ch_col == "Case Channel":
        def map_channel(ch):
            if ch in PHONE_CHANNELS:
                return "Phone"
            elif ch in NON_LIVE_CHANNELS:
                return "Non-live"
            return "Other"
        df["_channel"] = df[ch_col].apply(map_channel)
        ch_col = "_channel"

    # Filtro canale opzionale
    df_analysis = df
    if args.channel and ch_col:
        df_analysis = df[df[ch_col] == args.channel]

    channel_label = args.channel if args.channel else "(Tutto)"

    # Aggregazioni
    print("Calcolo aggregazioni...")
    case_types   = aggregate(df_analysis, aht_col, cas_col, typ_col,
                             args.target_phone, TOP_N_CASETYPE)
    primary_cats = (aggregate(df_analysis, aht_col, cas_col, cat_col,
                              args.target_nonlive, TOP_N_PRIMCAT)
                    if cat_col else pd.DataFrame())

    # Excel output
    wb = openpyxl.Workbook()

    ws_data = wb.active
    ws_data.title = "DATASET"
    print("Scrittura DATASET...")
    write_dataset_sheet(ws_data, df)

    ws_analysis = wb.create_sheet("Case Type Analysis")
    print("Scrittura Case Type Analysis...")
    write_analysis_sheet(ws_analysis, case_types, primary_cats, channel_label)

    ws_outliers = wb.create_sheet("xLori & Costa(B.Info)")
    print("Scrittura xLori & Costa(B.Info)...")
    write_xloricrosta_sheet(
        ws_outliers, df_analysis,
        aht_col, cas_col, typ_col, cat_col, ch_col, emp_col, num_col,
        args.target_phone, args.target_nonlive,
    )

    # Rendi attivo Case Type Analysis
    wb.active = ws_analysis

    week_label = args.week or csv_path.stem
    out_dir    = Path("output")
    out_dir.mkdir(exist_ok=True)
    out_path   = out_dir / f"case_KPI_report_W{week_label}.xlsx"
    wb.save(out_path)

    print(f"\nSalvato: {out_path}")
    print(f"\nRiepilogo:")
    print(f"  Righe raw data:       {len(df)}")
    print(f"  Case Types (top {TOP_N_CASETYPE}):  {len(case_types)}")
    if not primary_cats.empty:
        print(f"  Primary Cat. (top {TOP_N_PRIMCAT}): {len(primary_cats)}")
    outlier_count = len(df_analysis[df_analysis[aht_col] > args.target_phone])
    print(f"  Outlier AHT cases:    {outlier_count}")


if __name__ == "__main__":
    main()
