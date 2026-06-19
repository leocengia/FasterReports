"""
Genera agent_performance_score.xlsx a partire dai dati raw AHT.

Input:  raw_data/<filename>.csv  (separatore ;, decimale ,)
Output: output/agent_performance_score_<settimana>.xlsx

Uso:
    python src/agent_performance_score.py "raw_data/AHT blended WEEK 24.csv" --week 24
"""

import argparse
import sys
from pathlib import Path

import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# ── Parametri configurabili ────────────────────────────────────────────────────
WORKING_DAYS = 22          # giorni lavorativi nel mese (denominatore avg_daily_vol)
TARGETS = {
    "non_live": 20.0,
    "phone":    229 / 12,  # ≈ 19.083  (valore dal template)
    "blended":  234.5 / 12, # ≈ 19.542  (valore dal template)
}

NON_LIVE_CHANNELS = {"Contact Us", "Market Management", "Others", "Email"}
PHONE_CHANNELS    = {"Phone"}


# ── Calcoli ────────────────────────────────────────────────────────────────────

def weighted_aht(group: pd.DataFrame) -> float:
    total = group["cases"].sum()
    if total == 0:
        return 0.0
    return (group["aht"] * group["cases"]).sum() / total


def build_section(df: pd.DataFrame, channels: set | None = None,
                  sort_by_aht: bool = False, days: int = WORKING_DAYS) -> pd.DataFrame:
    """Aggrega per agente. channels=None → tutti i canali (blended)."""
    if channels is not None:
        df = df[df["Case Channel"].isin(channels)]

    agg = (
        df.groupby("Employee Name")
        .apply(lambda g: pd.Series({
            "aht":    weighted_aht(g),
            "volume": g["cases"].sum(),
        }), include_groups=False)
        .reset_index()
        .rename(columns={"Employee Name": "agent"})
    )

    # team_avg = AHT pesato sull'intero gruppo
    total_cases = agg["volume"].sum()
    team_avg = (agg["aht"] * agg["volume"]).sum() / total_cases if total_cases else 0.0
    agg["team_avg"] = team_avg
    agg["avg_daily_vol"] = agg["volume"] / days

    if sort_by_aht:
        agg = agg.sort_values("aht", ascending=False)
    else:
        agg = agg.sort_values("agent")

    return agg.reset_index(drop=True)


# ── Excel output ───────────────────────────────────────────────────────────────

# Palette colori (dal template)
COLOR_HEADER_BG  = "4472C4"   # blu
COLOR_SECTION_NL = "D9E1F2"   # azzurro chiaro   (Non-live)
COLOR_SECTION_PH = "FFF2CC"   # giallo chiaro    (Phone)
COLOR_SECTION_BL = "E2EFDA"   # verde chiaro     (Blended)

THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _cell(ws, row, col, value=None, *, bold=False, bg=None,
          align="center", fmt=None, formula=None):
    c = ws.cell(row=row, column=col, value=value if formula is None else formula)
    c.font = Font(bold=bold, name="Calibri", size=10)
    c.alignment = Alignment(horizontal=align, vertical="center")
    c.border = BORDER
    if bg:
        c.fill = PatternFill("solid", fgColor=bg)
    if fmt:
        c.number_format = fmt
    return c


def write_excel(non_live: pd.DataFrame, phone: pd.DataFrame,
                blended: pd.DataFrame, out_path: Path, week: str,
                days: int = WORKING_DAYS) -> None:

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Agent Performance"

    # ── Intestazione ──────────────────────────────────────────────────────────
    ws.merge_cells("B1:Y1")
    title_cell = ws["B1"]
    title_cell.value = f"Agent Performance Score — Week {week}"
    title_cell.font = Font(bold=True, size=14, name="Calibri")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 22

    # ── Sezioni: colonne di partenza ──────────────────────────────────────────
    # Non-live: B-G  |  Phone: I-N  |  Blended: P-V
    sections = [
        ("Non-live", non_live, 2,  TARGETS["non_live"], COLOR_SECTION_NL,
         ["Agent", "AHT", "Volume", "Team Avg", "Target", "Avg Daily Vol"]),
        ("Phone",    phone,    9,  TARGETS["phone"],    COLOR_SECTION_PH,
         ["Agent", "AHT", "Volume", "Team Avg", "Target", "Avg Daily Vol"]),
        ("Blended",  blended,  16, TARGETS["blended"],  COLOR_SECTION_BL,
         ["Agent", "AHT", "Team Avg", "Target", "Monthly Vol", "Avg Daily Vol"]),
    ]

    for section_label, data, start_col, target, bg_color, col_headers in sections:
        n_cols = len(col_headers)
        end_col = start_col + n_cols - 1

        # Titolo sezione (riga 3)
        ws.merge_cells(
            start_row=3, start_column=start_col,
            end_row=3,   end_column=end_col
        )
        _cell(ws, 3, start_col, section_label, bold=True, bg=COLOR_HEADER_BG,
              align="center")
        ws.cell(row=3, column=start_col).font = Font(
            bold=True, color="FFFFFF", size=11, name="Calibri")

        # Header colonne (riga 4)
        for i, h in enumerate(col_headers):
            _cell(ws, 4, start_col + i, h, bold=True, bg=COLOR_HEADER_BG,
                  align="center")
            ws.cell(row=4, column=start_col + i).font = Font(
                bold=True, color="FFFFFF", size=10, name="Calibri")

        is_blended = section_label == "Blended"

        for row_i, agent_row in data.iterrows():
            xl_row = 5 + row_i

            if is_blended:
                # Blended: Agent, AHT, team_avg, target, monthly_vol, avg_daily_vol
                _cell(ws, xl_row, start_col,     agent_row["agent"],    bg=bg_color, align="left")
                _cell(ws, xl_row, start_col + 1, agent_row["aht"],      bg=bg_color, fmt="0.00")
                _cell(ws, xl_row, start_col + 2, agent_row["team_avg"], bg=bg_color, fmt="0.00")
                _cell(ws, xl_row, start_col + 3, target,                bg=bg_color, fmt="0.00")
                _cell(ws, xl_row, start_col + 4, agent_row["volume"],   bg=bg_color, fmt="0")
                _cell(ws, xl_row, start_col + 5,
                      agent_row["volume"] / days, bg=bg_color, fmt="0.00")
            else:
                # Non-live / Phone: Agent, AHT, Volume, team_avg, target, avg_daily_vol
                _cell(ws, xl_row, start_col,     agent_row["agent"],    bg=bg_color, align="left")
                _cell(ws, xl_row, start_col + 1, agent_row["aht"],      bg=bg_color, fmt="0.00")
                _cell(ws, xl_row, start_col + 2, agent_row["volume"],   bg=bg_color, fmt="0")
                _cell(ws, xl_row, start_col + 3, agent_row["team_avg"], bg=bg_color, fmt="0.00")
                _cell(ws, xl_row, start_col + 4, target,                bg=bg_color, fmt="0.00")
                _cell(ws, xl_row, start_col + 5,
                      agent_row["volume"] / days, bg=bg_color, fmt="0.00")

    # ── Larghezze colonne ─────────────────────────────────────────────────────
    col_widths = {
        "B": 22, "C": 8,  "D": 8,  "E": 10, "F": 8,  "G": 12,  # Non-live
        "I": 22, "J": 8,  "K": 8,  "L": 10, "M": 8,  "N": 12,  # Phone
        "P": 22, "Q": 8,  "R": 10, "S": 8,  "T": 12, "U": 12,  # Blended
    }
    for col_letter, width in col_widths.items():
        ws.column_dimensions[col_letter].width = width

    # ── Altezza righe ─────────────────────────────────────────────────────────
    ws.row_dimensions[3].height = 18
    ws.row_dimensions[4].height = 16
    for r in range(5, 5 + max(len(non_live), len(phone), len(blended))):
        ws.row_dimensions[r].height = 15

    wb.save(out_path)
    print(f"Salvato: {out_path}")


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Genera agent_performance_score.xlsx")
    parser.add_argument("input", help="Path al CSV raw (es. raw_data/AHT blended WEEK 24.csv)")
    parser.add_argument("--week", default="", help="Numero settimana (es. 24)")
    parser.add_argument("--days", type=int, default=WORKING_DAYS,
                        help=f"Giorni lavorativi nel mese (default {WORKING_DAYS})")
    args = parser.parse_args()

    csv_path = Path(args.input)
    if not csv_path.exists():
        print(f"Errore: file non trovato: {csv_path}", file=sys.stderr)
        sys.exit(1)

    days = args.days

    # Lettura
    df = pd.read_csv(csv_path, sep=";", encoding="utf-8-sig", decimal=",")
    df["aht"]   = pd.to_numeric(df["Case AHT (mins)"],  errors="coerce").fillna(0)
    df["cases"] = pd.to_numeric(df["Distinct Cases"],   errors="coerce").fillna(0)
    df = df[df["cases"] > 0]

    # Sezioni
    non_live = build_section(df, channels=NON_LIVE_CHANNELS, sort_by_aht=False, days=days)
    phone    = build_section(df, channels=PHONE_CHANNELS,    sort_by_aht=False, days=days)
    blended  = build_section(df, channels=None,              sort_by_aht=True,  days=days)

    week_label = args.week or csv_path.stem
    out_dir = Path("output")
    out_dir.mkdir(exist_ok=True)
    out_path = out_dir / f"agent_performance_score_W{week_label}.xlsx"

    write_excel(non_live, phone, blended, out_path, week=week_label, days=days)

    # Riepilogo a terminale
    print(f"\nRiepilogo:")
    print(f"  Non-live agenti: {len(non_live)}  | team AHT: {non_live['team_avg'].iloc[0]:.2f} min")
    print(f"  Phone agenti:    {len(phone)}  | team AHT: {phone['team_avg'].iloc[0]:.2f} min")
    print(f"  Blended agenti:  {len(blended)}  | team AHT: {blended['team_avg'].iloc[0]:.2f} min")


if __name__ == "__main__":
    main()
