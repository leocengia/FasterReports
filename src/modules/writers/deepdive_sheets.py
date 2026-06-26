"""
Fogli del Case Type Deepdive (allineati al template wow_CaseType_Deepdive):
  - Weekly Deepdive : tabelle aggregate per canale + tabella tblParam editabile
  - tblClass        : scoring per Channel×Case Type; Gap/Variability/Score/Categoria
                      sono FORMULE Excel che leggono tblParam (ricalcolo live) e il
                      color coding è via formattazione condizionale
  - EXPORT          : tabella Phone/Non-live affiancata per copia-incolla su PPT
                      (stessa formattazione del vecchio foglio WoW Export)

I valori delle metriche/rank sono calcolati in Python (compute_deepdive) e scritti
come valori; le colonne derivate dai parametri sono formule live.
"""

import pandas as pd
from xlsxwriter.utility import xl_rowcol_to_cell, xl_col_to_name

from ..data_loader import CHANNEL_CONFIG
from .styles import (
    make_formats, C_YELLOW_LIGHT, C_BLUE_LIGHT, C_RED_LIGHT, C_GREEN_LIGHT,
)
from .wow_sheets import _case_type_order

# ── tblParam: posizione sul foglio Weekly Deepdive ───────────────────────────
_DD_SHEET      = "Weekly Deepdive"
_PARAM_COL0    = 18                 # colonna S
_PARAM_HDR_ROW = 0
_PARAM_ROW     = {"Phone": 1, "Non-live": 2, "All": 3}
_PARAM_FIELDS  = ["channel", "target_aht", "min_vol", "highly", "actionable",
                  "impact_high", "oot_mid", "oot_high", "var_high", "opp_rank_min"]
_PARAM_HEADERS = ["Channel", "Target_AHT", "Min_Volume", "Highly_Threshold",
                  "Actionable_Threshold", "Impact_High", "OOT_Mid", "OOT_High",
                  "Var_High", "Opp_Rank_Min"]


def _param_cell(channel: str, field: str) -> str:
    """Riferimento assoluto a una cella di tblParam sul foglio Weekly Deepdive."""
    r = _PARAM_ROW[channel]
    c = _PARAM_COL0 + _PARAM_FIELDS.index(field)
    return f"'{_DD_SHEET}'!{xl_rowcol_to_cell(r, c, row_abs=True, col_abs=True)}"


# ── Weekly Deepdive ───────────────────────────────────────────────────────────

def write_weekly_deepdive(wb, dd: pd.DataFrame) -> None:
    ws   = wb.add_worksheet(_DD_SHEET)
    fmts = make_formats(wb)
    if dd.empty:
        ws.write(0, 0, "Nessun dato", fmts["data"])
        return
    _write_agg_tables(ws, fmts, dd)
    _write_param_table(wb, ws, dd)


def _write_agg_tables(ws, fmts, dd: pd.DataFrame) -> None:
    headers = ["case_type", "aht (mins)", "volume", "impact (%)",
               "out_of_target (%)", "std_dev", "p90-p10"]
    case_types = _case_type_order(pd.DataFrame({"Case Type": dd["case_type"]}))

    for label, channel, start_col in [("Phone", "Phone", 0), ("Non-live", "Non-live", 10)]:
        sub = dd[dd["channel"] == channel].set_index("case_type")
        ws.merge_range(0, start_col, 0, start_col + len(headers) - 1, f"[ {label} ]", fmts["header"])
        for i, h in enumerate(headers):
            ws.write(1, start_col + i, h, fmts["header"])
        row = 2
        for ct in case_types:
            ws.write(row, start_col, ct, fmts["data_left"])
            if ct in sub.index:
                r = sub.loc[ct]
                ws.write_number(row, start_col + 1, float(r["aht"]),            fmts["data_num"])
                ws.write_number(row, start_col + 2, int(r["volume"]),           fmts["data_int"])
                ws.write_number(row, start_col + 3, float(r["impact_channel"]), fmts["data_pct"])
                ws.write_number(row, start_col + 4, float(r["out_of_target"]),  fmts["data_pct"])
                ws.write_number(row, start_col + 5, float(r["std_dev"]),        fmts["data_num"])
                ws.write_number(row, start_col + 6, float(r["p90p10"]),         fmts["data_num"])
            else:
                for c in range(1, len(headers)):
                    ws.write(row, start_col + c, "N/A", fmts["data"])
            row += 1

    ws.set_column(0,  0,  30)
    ws.set_column(1,  6,  12)
    ws.set_column(10, 10, 30)
    ws.set_column(11, 16, 12)


def _write_param_table(wb, ws, dd: pd.DataFrame) -> None:
    """tblParam editabile (righe Phone/Non-live/All) come Excel Table."""
    base = CHANNEL_CONFIG["Phone"]
    all_target = float(dd[dd["channel"] == "All"]["target_aht"].iloc[0])

    rows = []
    for ch in ["Phone", "Non-live", "All"]:
        cfg    = CHANNEL_CONFIG.get(ch, base)
        target = all_target if ch == "All" else cfg["target_aht"]
        rows.append([ch, target, cfg["min_vol"], cfg["highly"], cfg["actionable"],
                     cfg["impact_high"], cfg["oot_mid"], cfg["oot_high"],
                     cfg["var_high"], cfg["opp_rank_min"]])

    first_row, first_col = _PARAM_HDR_ROW, _PARAM_COL0
    last_row, last_col   = first_row + len(rows), first_col + len(_PARAM_HEADERS) - 1
    ws.add_table(first_row, first_col, last_row, last_col, {
        "name":    "tblParam",
        "columns": [{"header": h} for h in _PARAM_HEADERS],
        "data":    rows,
        "style":   "Table Style Medium 9",
    })
    ws.set_column(first_col, last_col, 14)


# ── tblClass (scoring con formule live) ───────────────────────────────────────

# (field, header, width, kind) — kind "f_*" = formula
_TBLCLASS_COLS = [
    ("channel",           "Channel",        12, "text"),
    ("case_type",         "Case Type",      28, "text"),
    ("chiave",            "Chiave",         30, "text"),
    ("volume",            "Volume",          8, "int"),
    ("aht",               "AHT",             8, "num"),
    ("total_minutes",     "Total Min",      10, "num"),
    ("impact_channel",    "Impact",          9, "pct"),
    ("out_of_target",     "OOT %",           8, "pct"),
    ("std_dev",           "StdDev",          8, "num"),
    ("p90p10",            "P90-P10",         9, "num"),
    ("target_aht",        "Target AHT",     10, "f_num"),
    ("gap_vs_target",     "Gap vs Target",  12, "f_num"),
    ("opp_minutes",       "Opp. Min",       10, "f_num"),
    ("std_dev_ratio",     "StdDev Ratio",   10, "num"),
    ("p90p10_ratio",      "P90P10 Ratio",   10, "num"),
    ("std_dev_rank",      "StdDev Rank",    10, "num"),
    ("p90p10_rank",       "P90P10 Rank",    10, "num"),
    ("variability_score", "Var. Score",     10, "f_num"),
    ("impact_rank",       "Impact Rank",    10, "num"),
    ("volume_rank",       "Volume Rank",    10, "num"),
    ("opportunity_rank",  "Opp. Rank",      10, "num"),
    ("oot_rank",          "OOT Rank",       10, "num"),
    ("score",             "Score",           9, "f_num"),
    ("categoria",         "Categoria",      18, "f_txt"),
]
_COL = {spec[0]: i for i, spec in enumerate(_TBLCLASS_COLS)}


def write_tblclass(wb, dd: pd.DataFrame) -> None:
    ws   = wb.add_worksheet("tblClass")
    fmts = make_formats(wb)

    for c, (_, label, width, _k) in enumerate(_TBLCLASS_COLS):
        ws.write(0, c, label, fmts["header"])
        ws.set_column(c, c, width)

    if dd.empty:
        ws.write(1, 0, "Nessun dato", fmts["data"])
        return

    val_fmt = {
        "int": fmts["data_int"], "num": fmts["data_num"],
        "pct": fmts["data_pct"], "text": fmts["data_left"],
        "f_num": fmts["data_num"], "f_txt": fmts["data"],
    }

    for i, (_, d) in enumerate(dd.iterrows()):
        wr = i + 1
        ch = d["channel"]
        cell = lambda field: xl_rowcol_to_cell(wr, _COL[field])

        # ── colonne a valore ──
        for field, _lbl, _w, kind in _TBLCLASS_COLS:
            if kind.startswith("f_"):
                continue
            c = _COL[field]
            v = d[field]
            if kind == "int":
                ws.write_number(wr, c, int(v), val_fmt[kind])
            elif kind in ("num", "pct"):
                ws.write_number(wr, c, float(v), val_fmt[kind])
            else:
                ws.write(wr, c, str(v), val_fmt[kind])

        # ── colonne a formula (live su tblParam) ──
        ws.write_formula(wr, _COL["target_aht"], f"={_param_cell(ch, 'target_aht')}",
                         val_fmt["f_num"], float(d["target_aht"]))
        ws.write_formula(wr, _COL["gap_vs_target"],
                         f"=MAX(0,{cell('aht')}-{cell('target_aht')})",
                         val_fmt["f_num"], float(d["gap_vs_target"]))
        ws.write_formula(wr, _COL["opp_minutes"],
                         f"={cell('gap_vs_target')}*{cell('volume')}",
                         val_fmt["f_num"], float(d["opp_minutes"]))
        ws.write_formula(wr, _COL["variability_score"],
                         f"=0.4*{cell('std_dev_rank')}+0.6*{cell('p90p10_rank')}",
                         val_fmt["f_num"], float(d["variability_score"]))
        ws.write_formula(wr, _COL["score"],
                         (f"=0.25*{cell('impact_rank')}+0.2*{cell('opportunity_rank')}"
                          f"+0.15*{cell('volume_rank')}+0.2*{cell('oot_rank')}"
                          f"+0.2*{cell('variability_score')}"),
                         val_fmt["f_num"], float(d["score"]))
        ws.write_formula(wr, _COL["categoria"], _categoria_formula(ch, cell),
                         val_fmt["f_txt"], str(d["categoria"]))

    # ── formattazione condizionale per Categoria (color coding live) ──
    n = len(dd)
    cat_letter = xl_col_to_name(_COL["categoria"])
    last_col   = len(_TBLCLASS_COLS) - 1
    rng = f"A2:{xl_rowcol_to_cell(n, last_col)}"
    for cls, color in [("Highly Actionable", C_RED_LIGHT),
                       ("Actionable",        C_YELLOW_LIGHT),
                       ("Process Driven",    C_GREEN_LIGHT)]:
        ws.conditional_format(rng, {
            "type":     "formula",
            "criteria": f'=${cat_letter}2="{cls}"',
            "format":   wb.add_format({"bg_color": color}),
        })

    ws.freeze_panes(1, 0)


def _categoria_formula(ch: str, cell) -> str:
    P = lambda f: _param_cell(ch, f)
    score, impact, gap = cell("score"), cell("impact_channel"), cell("gap_vs_target")
    oot, var, vol, opp = (cell("out_of_target"), cell("variability_score"),
                          cell("volume"), cell("opportunity_rank"))
    return (
        f'=IF({vol}<{P("min_vol")},"Process Driven",'
        f'IF(AND({score}>={P("highly")},{impact}>={P("impact_high")},{gap}>0,'
        f'OR({oot}>={P("oot_high")},{var}>={P("var_high")})),"Highly Actionable",'
        f'IF(OR(AND({score}>={P("actionable")},OR({gap}>0,{oot}>={P("oot_mid")})),'
        f'AND({var}>={P("var_high")},{vol}>=2*{P("min_vol")}),'
        f'AND({opp}>={P("opp_rank_min")},{gap}>0)),"Actionable",'
        f'"Process Driven")))'
    )


# ── EXPORT (per copia-incolla su PPT, formattazione del vecchio WoW Export) ────

def write_export(wb, dd: pd.DataFrame) -> None:
    ws   = wb.add_worksheet("EXPORT")
    fmts = make_formats(wb)
    headers_ch = ["aht (min)", "volume", "impact (%)", "out_of_target (%)", "std_dev", "p90-p10"]

    ws.write(0, 0, "case_type", fmts["header"])
    ws.merge_range(0, 1, 0, len(headers_ch), "Phone", fmts["header"])
    ws.merge_range(0, len(headers_ch) + 2, 0, len(headers_ch) * 2 + 1, "Non-live", fmts["header"])
    ws.write(1, 0, "", fmts["header"])
    for i, h in enumerate(headers_ch):
        ws.write(1, 1 + i, h, fmts["header"])
        ws.write(1, len(headers_ch) + 2 + i, h, fmts["header"])
    ws.write(1, len(headers_ch) + 1, "", fmts["header"])

    if dd.empty:
        ws.set_column(0, 0, 30)
        return

    ph = dd[dd["channel"] == "Phone"].set_index("case_type")
    nl = dd[dd["channel"] == "Non-live"].set_index("case_type")

    for r, ct in enumerate(_case_type_order(pd.DataFrame({"Case Type": dd["case_type"]})), start=2):
        zebra = (r % 2 == 0)
        row_fmt = fmts["gray"] if zebra else fmts["data"]
        row_num = wb.add_format({
            "font_name": "Calibri", "font_size": 10, "border": 1,
            "bg_color": "#F2F2F2" if zebra else "#FFFFFF",
            "align": "center", "num_format": "0.00",
        })
        ws.write(r, 0, ct, fmts["gray_left"] if zebra else fmts["data_left"])

        for src, start_col in [(ph, 1), (nl, len(headers_ch) + 2)]:
            if ct in src.index:
                rr   = src.loc[ct]
                vals = [float(rr["aht"]), int(rr["volume"]), float(rr["impact_channel"]),
                        float(rr["out_of_target"]), float(rr["std_dev"]), float(rr["p90p10"])]
                for c, v in enumerate(vals):
                    ws.write_number(r, start_col + c, v, row_num)
            else:
                for c in range(len(headers_ch)):
                    ws.write(r, start_col + c, "N/A", row_fmt)

    ws.set_column(0, 0, 30)
    ws.set_column(1, len(headers_ch) * 2 + 2, 12)
    ws.freeze_panes(2, 1)
