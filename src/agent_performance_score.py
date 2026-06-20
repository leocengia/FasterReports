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
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, RadarChart, Reference
from openpyxl.drawing.spreadsheet_drawing import TwoCellAnchor, AnchorMarker

# ── Parametri configurabili ────────────────────────────────────────────────────
WORKING_DAYS = 22
TARGETS = {
    "non_live": 20.0,
    "phone":    229 / 12,   # ≈ 19.083
    "blended":  234.5 / 12, # ≈ 19.542
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

THIN   = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _cell(ws, row, col, value=None, *, bold=False, size=11,
          align="center", fmt=None):
    c = ws.cell(row=row, column=col, value=value)
    c.font = Font(bold=bold, name="Calibri", size=size)
    c.alignment = Alignment(horizontal=align, vertical="center")
    c.border = BORDER
    if fmt:
        c.number_format = fmt
    return c


def _add_chart(ws, chart, from_col, from_row, to_col, to_row,
               from_col_off=0, from_row_off=0, to_col_off=0, to_row_off=0):
    """Aggiunge un grafico con anchor a due celle (coordinate 0-indexed)."""
    anchor = TwoCellAnchor()
    anchor._from = AnchorMarker(col=from_col, colOff=from_col_off,
                                row=from_row, rowOff=from_row_off)
    anchor.to    = AnchorMarker(col=to_col,   colOff=to_col_off,
                                row=to_row,   rowOff=to_row_off)
    chart.anchor = anchor
    ws.add_chart(chart)


def write_excel(non_live: pd.DataFrame, phone: pd.DataFrame,
                blended: pd.DataFrame, out_path: Path, week: str,
                days: int = WORKING_DAYS) -> None:

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Agent Performance"

    # ── Layout colonne (identico al template) ─────────────────────────────────
    # Non-live : E(5)–J(10)
    # Phone    : L(12)–Q(17)
    # Blended  : S(19)–X(24)
    sections = [
        ("Non-live", non_live, 5,  TARGETS["non_live"],
         ["Agent", "aht", "volume", "team_avg", "target", "avg_daily_vol"]),
        ("Phone",    phone,    12, TARGETS["phone"],
         ["Agent", "aht", "volume", "team_avg", "target", "avg_daily_vol"]),
        ("Blended",  blended,  19, TARGETS["blended"],
         ["Agent", "aht", "team_avg", "target", "monthly_vol", "avg_daily_vol"]),
    ]

    for section_label, data, start_col, target, col_headers in sections:
        n_cols   = len(col_headers)
        end_col  = start_col + n_cols - 1
        is_blended = section_label == "Blended"
        # colonna volume: G(7) per Non-live/Phone, W(23) per Blended
        vol_col  = start_col + (4 if is_blended else 2)

        # Titolo sezione – riga 3
        ws.merge_cells(start_row=3, start_column=start_col,
                       end_row=3,   end_column=end_col)
        _cell(ws, 3, start_col, section_label, bold=True, size=11)

        # Intestazioni colonne – riga 4
        for i, h in enumerate(col_headers):
            _cell(ws, 4, start_col + i, h, bold=True, size=11)

        # Dati – dalla riga 5 in poi
        for row_i, agent_row in data.iterrows():
            xl_row     = 5 + row_i
            vol_letter = get_column_letter(vol_col)

            if is_blended:
                _cell(ws, xl_row, start_col,     agent_row["agent"],    align="left")
                _cell(ws, xl_row, start_col + 1, agent_row["aht"],      fmt="0.00")
                _cell(ws, xl_row, start_col + 2, agent_row["team_avg"], fmt="0.00")
                _cell(ws, xl_row, start_col + 3, target,                fmt="0.00")
                _cell(ws, xl_row, start_col + 4, agent_row["volume"],   fmt="0")
                _cell(ws, xl_row, start_col + 5,
                      f"={vol_letter}{xl_row}/{days}", fmt="0.00")
            else:
                _cell(ws, xl_row, start_col,     agent_row["agent"],    align="left")
                _cell(ws, xl_row, start_col + 1, agent_row["aht"],      fmt="0.00")
                _cell(ws, xl_row, start_col + 2, agent_row["volume"],   fmt="0")
                _cell(ws, xl_row, start_col + 3, agent_row["team_avg"], fmt="0.00")
                _cell(ws, xl_row, start_col + 4, target,                fmt="0.00")
                _cell(ws, xl_row, start_col + 5,
                      f"={vol_letter}{xl_row}/{days}", fmt="0.00")

    # ── Larghezze colonne (dal template) ──────────────────────────────────────
    col_widths = {
        # Colonne prima dei dati
        "A": 23.0, "B": 17.28, "C": 7.71,
        # Non-live  E F (G default) H (I default) J
        "E": 21.14, "F": 5.57, "H": 9.57, "J": 17.28,
        # Separatore
        "K": 2.71,
        # Phone  L M N (O default) (P default) Q
        "L": 21.14, "M": 5.57, "N": 7.71, "Q": 13.14,
        # Separatore
        "R": 2.85,
        # Blended  S T U (V default) W X
        "S": 21.14, "T": 12.14, "U": 21.14, "W": 12.14, "X": 21.14,
    }
    for col_letter, width in col_widths.items():
        ws.column_dimensions[col_letter].width = width

    # ── Altezze righe (dal template) ──────────────────────────────────────────
    for r in [2, 3, 4]:
        ws.row_dimensions[r].height = 15.75
    n_data = max(len(non_live), len(phone), len(blended))
    if n_data > 0:
        ws.row_dimensions[4 + n_data].height = 15.75

    # ── Grafici ───────────────────────────────────────────────────────────────
    n_nl = len(non_live)
    n_ph = len(phone)
    n_bl = len(blended)

    # Chart 0 – LineChart "Blended Agent KPI insights"
    # Serie: T(20)=aht, U(21)=team_avg, V(22)=target  |  Cat: S(19)
    # Posizione: sotto la tabella (col D row 49 → col Y row 89)
    lc = LineChart()
    lc.title  = "Blended Agent KPI insights"
    lc.style  = 10
    lc.add_data(
        Reference(ws, min_col=20, max_col=22, min_row=4, max_row=4 + n_bl),
        titles_from_data=True,
    )
    lc.set_categories(Reference(ws, min_col=19, min_row=5, max_row=4 + n_bl))
    _add_chart(ws, lc,
               from_col=3,  from_row=48, from_col_off=271894, from_row_off=91785,
               to_col=24,   to_row=88,   to_col_off=34636,    to_row_off=51954)

    # Chart 1 – RadarChart "Agent vs Team AHT [Phone]"
    # (template: usa dati Non-live  F(6)=aht, H(8)=team_avg, I(9)=target, cat E(5))
    # Posizione: destra, metà superiore (col Y row 2 → col AN row 42)
    rc_phone = RadarChart()
    rc_phone.title = "Agent vs Team AHT [Phone]"
    rc_phone.style = 26
    for col_idx in [6, 8, 9]:  # F, H, I → Non-live aht/team_avg/target
        rc_phone.add_data(
            Reference(ws, min_col=col_idx, min_row=4, max_row=4 + n_nl),
            titles_from_data=True,
        )
    rc_phone.set_categories(Reference(ws, min_col=5, min_row=5, max_row=4 + n_nl))
    _add_chart(ws, rc_phone,
               from_col=24, from_row=1,  from_col_off=324969, from_row_off=186016,
               to_col=39,   to_row=41,   to_col_off=190500,   to_row_off=122465)

    # Chart 2 – RadarChart "Agent vs Team AHT [Non-live]"
    # (template: usa dati Phone  M(13)=aht, O(15)=team_avg, P(16)=target, cat L(12))
    # Posizione: destra, metà inferiore (col Y row 43 → col AN row 81)
    rc_nonlive = RadarChart()
    rc_nonlive.title = "Agent vs Team AHT [Non-live]"
    rc_nonlive.style = 26
    for col_idx in [13, 15, 16]:  # M, O, P → Phone aht/team_avg/target
        rc_nonlive.add_data(
            Reference(ws, min_col=col_idx, min_row=4, max_row=4 + n_ph),
            titles_from_data=True,
        )
    rc_nonlive.set_categories(Reference(ws, min_col=12, min_row=5, max_row=4 + n_ph))
    _add_chart(ws, rc_nonlive,
               from_col=24, from_row=42, from_col_off=326570, from_row_off=186417,
               to_col=39,   to_row=80,   to_col_off=204106,   to_row_off=81643)

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
    df = pd.read_csv(csv_path, sep=";", encoding="utf-8-sig", decimal=",")
    df["aht"]   = pd.to_numeric(df["Case AHT (mins)"],  errors="coerce").fillna(0)
    df["cases"] = pd.to_numeric(df["Distinct Cases"],   errors="coerce").fillna(0)
    df = df[df["cases"] > 0]

    non_live = build_section(df, channels=NON_LIVE_CHANNELS, sort_by_aht=False, days=days)
    phone    = build_section(df, channels=PHONE_CHANNELS,    sort_by_aht=False, days=days)
    blended  = build_section(df, channels=None,              sort_by_aht=True,  days=days)

    week_label = args.week or csv_path.stem
    out_dir    = Path("output")
    out_dir.mkdir(exist_ok=True)
    out_path   = out_dir / f"agent_performance_score_W{week_label}.xlsx"

    write_excel(non_live, phone, blended, out_path, week=week_label, days=days)

    print(f"\nRiepilogo:")
    print(f"  Non-live agenti: {len(non_live)}  | team AHT: {non_live['team_avg'].iloc[0]:.2f} min")
    print(f"  Phone agenti:    {len(phone)}  | team AHT: {phone['team_avg'].iloc[0]:.2f} min")
    print(f"  Blended agenti:  {len(blended)}  | team AHT: {blended['team_avg'].iloc[0]:.2f} min")


if __name__ == "__main__":
    main()
