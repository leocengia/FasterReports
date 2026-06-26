"""
Fogli del Case Type Deepdive (allineati al template wow_CaseType_Deepdive):
  - tblClass        : scoring per Channel×Case Type. Gap/Variability/Score/Categoria
                      sono FORMULE Excel che leggono tblParam (ricalcolo live).
                      NESSUN color coding (foglio di calcolo/parametri).
  - Weekly Deepdive : tabelle Phone/Non-live (stile template) + tblParam editabile;
                      righe colorate per Categoria via formattazione condizionale
                      (lookup live in tblClass).
  - CT DD Graphs    : grafici Ranking per Score + Pareto Opportunity Minutes.
  - EXPORT          : 3 blocchi come manualEXPORT (Phone tutti / Non-live tutti /
                      selezionati per canale raggruppati per Categoria) per PPT.
"""

import pandas as pd
from xlsxwriter.utility import xl_rowcol_to_cell

from ..data_loader import CHANNEL_CONFIG
from .styles import (
    make_formats, C_RED_LIGHT, C_YELLOW_LIGHT, C_GREEN_LIGHT,
    C_EG_BLUE, C_EG_BLUE_LT, C_EG_ORANGE, C_EG_GRAY, CHART_CATEGORY_COLORS,
)
from .wow_sheets import _case_type_order

# ── tblParam: posizione sul foglio Weekly Deepdive ───────────────────────────
_DD_SHEET      = "Weekly Deepdive"
_PARAM_COL0    = 9                 # colonna J
_PARAM_HDR_ROW = 0
_PARAM_ROW     = {"Phone": 1, "Non-live": 2, "All": 3}
_PARAM_FIELDS  = ["channel", "target_aht", "min_vol", "highly", "actionable",
                  "impact_high", "oot_mid", "oot_high", "var_high", "opp_rank_min"]
_PARAM_HEADERS = ["Channel", "Target_AHT", "Min_Volume", "Highly_Threshold",
                  "Actionable_Threshold", "Impact_High", "OOT_Mid", "OOT_High",
                  "Var_High", "Opp_Rank_Min"]

_METRICS = ["aht", "volume", "impact_channel", "out_of_target", "std_dev", "p90p10"]
_METRIC_HEADERS = ["aht (mins)", "volume", "impact (%)", "out_of_target (%)", "std_dev", "p90-p10"]
_CLS_COLORS = [("Highly Actionable", C_RED_LIGHT), ("Actionable", C_YELLOW_LIGHT),
               ("Process Driven", C_GREEN_LIGHT)]


def _param_cell(channel: str, field: str) -> str:
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

    sect = wb.add_format({"bold": True, "font_size": 11, "font_color": "white",
                          "bg_color": C_EG_BLUE, "align": "center", "valign": "vcenter", "border": 1})
    hdr  = wb.add_format({"bold": True, "font_size": 11, "bg_color": C_EG_BLUE_LT,
                          "align": "center", "valign": "vcenter", "border": 1})
    f_txt = wb.add_format({"font_size": 10, "border": 1})
    f_num = wb.add_format({"font_size": 10, "border": 1, "num_format": "0.00", "align": "center"})
    f_int = wb.add_format({"font_size": 10, "border": 1, "num_format": "0", "align": "center"})
    f_pct = wb.add_format({"font_size": 10, "border": 1, "num_format": "0.00%", "align": "center"})
    f_na  = wb.add_format({"font_size": 10, "border": 1, "align": "center"})

    case_types = _case_type_order(pd.DataFrame({"Case Type": dd["case_type"]}))

    r = 0
    blocks = []
    for label, channel in [("[ Phone ]", "Phone"), ("[ Non-live ]", "Non-live")]:
        ws.merge_range(r, 0, r, 6, label, sect)
        r += 1
        ws.write(r, 0, "case_type", hdr)
        for i, h in enumerate(_METRIC_HEADERS):
            ws.write(r, 1 + i, h, hdr)
        r += 1
        first = r
        sub = dd[dd["channel"] == channel].set_index("case_type")
        for ct in case_types:
            ws.write(r, 0, ct, f_txt)
            if ct in sub.index:
                row = sub.loc[ct]
                ws.write_number(r, 1, float(row["aht"]),            f_num)
                ws.write_number(r, 2, int(row["volume"]),           f_int)
                ws.write_number(r, 3, float(row["impact_channel"]), f_pct)
                ws.write_number(r, 4, float(row["out_of_target"]),  f_pct)
                ws.write_number(r, 5, float(row["std_dev"]),        f_num)
                ws.write_number(r, 6, float(row["p90p10"]),         f_num)
            else:
                for c in range(1, 7):
                    ws.write(r, c, "N/A", f_na)
            r += 1
        blocks.append((channel, first, r - 1))
        r += 1  # riga vuota tra i blocchi

    # larghezze (dal template)
    for col, w in [(0, 33), (1, 14.5), (2, 12.3), (3, 15), (4, 21.3), (5, 12.6), (6, 13)]:
        ws.set_column(col, col, w)

    _write_param_table(wb, ws, dd)

    # color coding per Categoria (lookup live in tblClass tramite la colonna Chiave)
    for channel, first, last in blocks:
        rng = f"A{first + 1}:G{last + 1}"
        for cls, color in _CLS_COLORS:
            ws.conditional_format(rng, {
                "type": "formula",
                "criteria": (f'=INDEX(tblClass!$X:$X,'
                             f'MATCH("{channel}|"&$A{first + 1},tblClass!$C:$C,0))="{cls}"'),
                "format": wb.add_format({"bg_color": color}),
            })


def _write_param_table(wb, ws, dd: pd.DataFrame) -> None:
    base = CHANNEL_CONFIG["Phone"]
    all_target = float(dd[dd["channel"] == "All"]["target_aht"].iloc[0])
    rows = []
    for ch in ["Phone", "Non-live", "All"]:
        cfg    = CHANNEL_CONFIG.get(ch, base)
        target = all_target if ch == "All" else cfg["target_aht"]
        rows.append([ch, target, cfg["min_vol"], cfg["highly"], cfg["actionable"],
                     cfg["impact_high"], cfg["oot_mid"], cfg["oot_high"],
                     cfg["var_high"], cfg["opp_rank_min"]])
    first_col = _PARAM_COL0
    last_col  = first_col + len(_PARAM_HEADERS) - 1
    ws.add_table(_PARAM_HDR_ROW, first_col, _PARAM_HDR_ROW + len(rows), last_col, {
        "name":    "tblParam",
        "columns": [{"header": h} for h in _PARAM_HEADERS],
        "data":    rows,
        "style":   "Table Style Medium 2",
    })
    ws.set_column(first_col, last_col, 13)


# ── tblClass (scoring con formule live, senza color coding) ──────────────────

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

    val_fmt = {"int": fmts["data_int"], "num": fmts["data_num"], "pct": fmts["data_pct"],
               "text": fmts["data_left"], "f_num": fmts["data_num"], "f_txt": fmts["data"]}

    for i, (_, d) in enumerate(dd.iterrows()):
        wr = i + 1
        ch = d["channel"]
        cell = lambda field: xl_rowcol_to_cell(wr, _COL[field])

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


# ── CT DD Graphs ──────────────────────────────────────────────────────────────

def write_ct_dd_graphs(wb, dd: pd.DataFrame) -> None:
    ws   = wb.add_worksheet("CT DD Graphs")
    fmts = make_formats(wb)
    if dd.empty:
        ws.write(0, 0, "Nessun dato", fmts["data"])
        return
    allr = dd[dd["channel"] == "All"].copy()

    # ── Ranking per Score (top 15) ──
    rank = allr.sort_values("score", ascending=False).head(15).reset_index(drop=True)
    ws.write(0, 0, "Case Type", fmts["header"])
    ws.write(0, 1, "Score", fmts["header"])
    ws.write(0, 2, "Categoria", fmts["header"])
    for i, row in rank.iterrows():
        ws.write(i + 1, 0, str(row["case_type"]), fmts["data_left"])
        ws.write_number(i + 1, 1, float(row["score"]), fmts["data_num"])
        ws.write(i + 1, 2, str(row["categoria"]), fmts["data"])
    n1 = len(rank)
    ch1 = wb.add_chart({"type": "bar"})
    ch1.add_series({
        "name":       "Score",
        "categories": ["CT DD Graphs", 1, 0, n1, 0],
        "values":     ["CT DD Graphs", 1, 1, n1, 1],
        "points":     [{"fill": {"color": CHART_CATEGORY_COLORS.get(c, C_EG_GRAY)}}
                       for c in rank["categoria"]],
        "data_labels": {"value": True, "num_format": "0.00"},
    })
    ch1.set_title({"name": "Case Type — Ranking per Score"})
    ch1.set_x_axis({"name": "Score"})
    ch1.set_legend({"none": True})
    ch1.set_size({"width": 620, "height": 440})
    ws.insert_chart("E1", ch1)

    # ── Pareto Opportunity Minutes (top 15) ──
    base = 26
    par = allr.sort_values("opp_minutes", ascending=False).head(15).reset_index(drop=True)
    tot = float(par["opp_minutes"].sum())
    par["cum_pct"] = par["opp_minutes"].cumsum() / tot if tot else 0.0
    ws.write(base, 0, "Case Type", fmts["header"])
    ws.write(base, 1, "Opp. Min", fmts["header"])
    ws.write(base, 2, "Cum %", fmts["header"])
    for i, row in par.iterrows():
        ws.write(base + 1 + i, 0, str(row["case_type"]), fmts["data_left"])
        ws.write_number(base + 1 + i, 1, float(row["opp_minutes"]), fmts["data_num"])
        ws.write_number(base + 1 + i, 2, float(row["cum_pct"]), fmts["data_pct"])
    n2 = len(par)
    col = wb.add_chart({"type": "column"})
    col.add_series({
        "name":       "Opportunity Minutes",
        "categories": ["CT DD Graphs", base + 1, 0, base + n2, 0],
        "values":     ["CT DD Graphs", base + 1, 1, base + n2, 1],
        "fill":       {"color": C_EG_BLUE},
    })
    line = wb.add_chart({"type": "line"})
    line.add_series({
        "name":       "% cumulata",
        "categories": ["CT DD Graphs", base + 1, 0, base + n2, 0],
        "values":     ["CT DD Graphs", base + 1, 2, base + n2, 2],
        "line":       {"color": C_EG_ORANGE, "width": 2.25},
        "y2_axis":    True,
    })
    col.combine(line)
    col.set_title({"name": "Pareto — Opportunity Minutes"})
    col.set_x_axis({"num_font": {"rotation": -45}})
    col.set_y_axis({"name": "Opportunity Minutes"})
    col.set_y2_axis({"name": "% cumulata", "min": 0, "max": 1, "num_format": "0%"})
    col.set_legend({"position": "bottom"})
    col.set_size({"width": 720, "height": 440})
    ws.insert_chart("E26", col)

    ws.set_column(0, 0, 30)
    ws.set_column(1, 2, 12)


# ── EXPORT (3 blocchi come manualEXPORT, per copia-incolla su PPT) ────────────

def write_export(wb, dd: pd.DataFrame) -> None:
    ws   = wb.add_worksheet("EXPORT")
    fmts = make_formats(wb)

    sect = wb.add_format({"bold": True, "font_color": "white", "bg_color": C_EG_BLUE,
                          "align": "center", "valign": "vcenter", "border": 1})
    hdr  = wb.add_format({"bold": True, "bg_color": C_EG_BLUE_LT, "align": "center",
                          "valign": "vcenter", "border": 1})
    f_txt = wb.add_format({"border": 1})
    f_num = wb.add_format({"border": 1, "num_format": "0.00", "align": "center"})
    f_int = wb.add_format({"border": 1, "num_format": "0", "align": "center"})
    f_pct = wb.add_format({"border": 1, "num_format": "0.00%", "align": "center"})
    f_na  = wb.add_format({"border": 1, "align": "center"})
    cat_lbl = wb.add_format({"bold": True, "italic": True, "bg_color": C_EG_BLUE_LT, "border": 1})

    if dd.empty:
        ws.write(0, 0, "Nessun dato", fmts["data"])
        return

    ph = dd[dd["channel"] == "Phone"].set_index("case_type")
    nl = dd[dd["channel"] == "Non-live"].set_index("case_type")
    case_types = _case_type_order(pd.DataFrame({"Case Type": dd["case_type"]}))

    def write_metrics(r, c0, src, ct):
        if ct in src.index:
            row = src.loc[ct]
            ws.write_number(r, c0 + 0, float(row["aht"]),            f_num)
            ws.write_number(r, c0 + 1, int(row["volume"]),           f_int)
            ws.write_number(r, c0 + 2, float(row["impact_channel"]), f_pct)
            ws.write_number(r, c0 + 3, float(row["out_of_target"]),  f_pct)
            ws.write_number(r, c0 + 4, float(row["std_dev"]),        f_num)
            ws.write_number(r, c0 + 5, float(row["p90p10"]),         f_num)
        else:
            for k in range(6):
                ws.write(r, c0 + k, "N/A", f_na)

    # ── Block1 (Phone) col B-H + Block2 (Non-live) col J-O — header riga 2-3 ──
    ws.merge_range(2, 1, 3, 1, "case_type", hdr)          # B3:B4
    ws.merge_range(2, 2, 2, 7, "[ Phone ]", sect)         # C3:H3
    ws.merge_range(2, 9, 2, 14, "[ Non-live ]", sect)     # J3:O3
    for i, h in enumerate(_METRIC_HEADERS):
        ws.write(3, 2 + i, h, hdr)
        ws.write(3, 9 + i, h, hdr)

    r = 4
    for ct in case_types:
        ws.write(r, 1, ct, f_txt)
        write_metrics(r, 2, ph, ct)
        write_metrics(r, 9, nl, ct)
        r += 1

    # ── Block3 (col Q-W): selezionati per canale raggruppati per Categoria ──
    Q = 16
    ws.merge_range(2, Q, 2, Q + 6, "Selezionati (per Score)", sect)
    ws.write(3, Q, "case_type", hdr)
    for i, h in enumerate(_METRIC_HEADERS):
        ws.write(3, Q + 1 + i, h, hdr)

    r = 4
    for channel, src in [("Phone", ph), ("Non-live", nl)]:
        ws.merge_range(r, Q, r, Q + 6, f"[ {channel} ]", sect)
        r += 1
        chan = dd[(dd["channel"] == channel) &
                  (dd["categoria"].isin(["Highly Actionable", "Actionable"]))]
        chan = chan.sort_values("score", ascending=False)
        if chan.empty:
            ws.write(r, Q, "(nessuno)", f_txt)
            for k in range(6):
                ws.write(r, Q + 1 + k, "", f_na)
            r += 1
            continue
        for cat in ["Highly Actionable", "Actionable"]:
            grp = chan[chan["categoria"] == cat]
            if grp.empty:
                continue
            ws.merge_range(r, Q, r, Q + 6, cat, cat_lbl)
            r += 1
            for _, row in grp.iterrows():
                ct = row["case_type"]
                ws.write(r, Q, ct, f_txt)
                write_metrics(r, Q + 1, src, ct)
                r += 1

    # larghezze (dal template) + separatori stretti I e P
    widths = {1: 32, 2: 10, 3: 7.7, 4: 10.4, 5: 16.7, 6: 8, 7: 13, 8: 1.3,
              9: 10, 10: 7.7, 11: 10.4, 12: 16.7, 13: 8, 14: 13, 15: 3.85,
              16: 27.4, 17: 10, 18: 7.7, 19: 10.4, 20: 16.7, 21: 8, 22: 13}
    for c, w in widths.items():
        ws.set_column(c, c, w)
    ws.set_column(0, 0, 5)
