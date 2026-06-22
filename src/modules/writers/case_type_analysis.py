"""
Foglio Case Type Analysis: tre blocchi tabellari + grafici barre/linea.

Layout:
  - Blocco A (col 0-4):  Case Origin breakdown (tutti i case type, per canale)
  - Blocco B (col 6-10): Case Type summary (solo Actionable/Highly Actionable)
  - Blocco C (col 12-15): Primary Category summary
  - Dati ausiliari grafici: colonne 20-23 (B) e 25-28 (C), fuori dalle tabelle
  - Grafici combinati barre+linea inseriti sotto le tabelle, affiancati
"""

import pandas as pd
from ..classification import build_tbl_class, get_relevant_case_types
from ..data_loader import CHANNEL_CONFIG, weighted_aht
from .styles import (
    make_formats, C_YELLOW_LIGHT, C_BLUE_LIGHT, CLASSIFICATION_COLORS,
)

# Colonne riservate ai dati sorgente dei grafici (lontane dalle tabelle visibili)
_CHART_DATA_COL_B = 20
_CHART_DATA_COL_C = 25


def write_case_type_analysis(wb, df: pd.DataFrame, week: str) -> None:
    ws   = wb.add_worksheet("Case Type Analysis")
    fmts = make_formats(wb)
    tbl  = build_tbl_class(df)

    end_a, chart_b, chart_c = None, None, None

    end_a = _write_block_a(ws, fmts, wb, df)
    end_b, chart_b = _write_block_b(ws, fmts, wb, df, tbl, week, start_col=6)
    end_c, chart_c = _write_block_c(ws, fmts, wb, df, week, start_col=12)

    # I grafici vengono inseriti sotto la tabella più lunga, affiancati
    chart_row = max(end_a, end_b, end_c) + 2
    if chart_b is not None:
        ws.insert_chart(chart_row, 0, chart_b, {"x_offset": 0, "y_offset": 0})
    if chart_c is not None:
        ws.insert_chart(chart_row, 9, chart_c, {"x_offset": 0, "y_offset": 0})

    # ── Larghezze ─────────────────────────────────────────────────────────────
    ws.set_column(0,  0,  28)
    ws.set_column(1,  4,  10)
    ws.set_column(5,  5,  2)   # separatore
    ws.set_column(6,  6,  28)
    ws.set_column(7,  10, 10)
    ws.set_column(11, 11, 2)   # separatore
    ws.set_column(12, 12, 28)
    ws.set_column(13, 15, 10)
    # Nasconde le colonne dei dati sorgente dei grafici
    ws.set_column(_CHART_DATA_COL_B, _CHART_DATA_COL_C + 4, None, None, {"hidden": True})


# ── Blocco A ─────────────────────────────────────────────────────────────────

def _write_block_a(ws, fmts, wb, df: pd.DataFrame) -> int:
    """Case Origin (group) breakdown con subtotali per case type. Ritorna l'ultima riga usata."""
    ws.merge_range(0, 0, 0, 4, "Case Origin Breakdown", fmts["header"])
    headers = ["Case Type", "avg AHT", "volume", "# above target", "% OOT"]
    for i, h in enumerate(headers):
        ws.write(1, i, h, fmts["header"])

    row = 2
    for channel_key, label, bg in [
        ("Phone",    "[ Phone ]",    C_YELLOW_LIGHT),
        ("Non-live", "[ Non-live ]", C_BLUE_LIGHT),
    ]:
        sub = df[df["channel"] == channel_key]
        cfg = CHANNEL_CONFIG[channel_key]
        target = cfg["target_aht"]
        title_fmt = wb.add_format({
            "font_name": "Calibri", "font_size": 10, "bold": True,
            "bg_color": bg, "border": 1, "align": "center",
        })
        ws.merge_range(row, 0, row, 4, label, title_fmt)
        row += 1

        if not sub.empty:
            for ct, grp in sub.groupby("Case Type"):
                vol    = int(grp["cases"].sum())
                aht_v  = weighted_aht(grp)
                above  = int((grp["aht"] > target).sum())
                oot    = above / vol if vol else 0.0
                ws.write(row, 0, ct,    fmts["data_left"])
                ws.write_number(row, 1, aht_v,  fmts["data_num"])
                ws.write_number(row, 2, vol,    fmts["data_int"])
                ws.write_number(row, 3, above,  fmts["data_int"])
                ws.write_number(row, 4, oot,    fmts["data_pct"])
                row += 1

        # Totale canale
        tot_vol   = int(sub["cases"].sum()) if not sub.empty else 0
        tot_aht   = weighted_aht(sub) if not sub.empty else 0.0
        tot_above = int((sub["aht"] > target).sum()) if not sub.empty else 0
        tot_oot   = tot_above / tot_vol if tot_vol else 0.0
        total_fmt = wb.add_format({
            "font_name": "Calibri", "font_size": 10, "bold": True,
            "bg_color": "#DDDDDD", "border": 1, "align": "center",
        })
        ws.write(row, 0, "TOTAL", total_fmt)
        ws.write_number(row, 1, tot_aht,   total_fmt)
        ws.write_number(row, 2, tot_vol,   total_fmt)
        ws.write_number(row, 3, tot_above, total_fmt)
        ws.write_number(row, 4, tot_oot,   total_fmt)
        row += 2

    return row


# ── Blocco B ─────────────────────────────────────────────────────────────────

def _write_block_b(ws, fmts, wb, df: pd.DataFrame, tbl: pd.DataFrame,
                   week: str, start_col: int):
    """
    Case Type summary per canale (solo Actionable/Highly Actionable).
    Ritorna (ultima_riga, chart) — il chart va inserito dal chiamante.
    """
    headers = ["Case Type", "avg AHT", "volume", "target", "class."]
    ws.merge_range(0, start_col, 0, start_col + len(headers) - 1,
                   f"Case Type Summary – W{week}", fmts["header"])
    for i, h in enumerate(headers):
        ws.write(1, start_col + i, h, fmts["header"])

    row = 2
    chart_rows: list[tuple] = []   # (ct, aht, vol, target)

    for channel_key, label, bg in [
        ("Phone",    "[ Phone ]",    C_YELLOW_LIGHT),
        ("Non-live", "[ Non-live ]", C_BLUE_LIGHT),
    ]:
        sub      = df[df["channel"] == channel_key]
        cfg      = CHANNEL_CONFIG[channel_key]
        target   = cfg["target_aht"]
        relevant = get_relevant_case_types(tbl, channel_key)

        title_fmt = wb.add_format({
            "font_name": "Calibri", "font_size": 10, "bold": True,
            "bg_color": bg, "border": 1, "align": "center",
        })
        ws.merge_range(row, start_col, row, start_col + len(headers) - 1, label, title_fmt)
        row += 1

        for ct in relevant:
            grp   = sub[sub["Case Type"] == ct]
            vol   = int(grp["cases"].sum()) if not grp.empty else 0
            aht_v = weighted_aht(grp) if not grp.empty else 0.0
            cls   = _get_class(tbl, channel_key, ct)
            cf    = _cls_fmt(wb, cls)
            cfn   = _cls_fmt(wb, cls, num=True)
            ws.write(row, start_col,     ct,     cf)
            ws.write_number(row, start_col + 1, aht_v,  cfn)
            ws.write_number(row, start_col + 2, vol,    cf)
            ws.write_number(row, start_col + 3, target, cfn)
            ws.write(row, start_col + 4, cls,    cf)
            chart_rows.append((f"{ct} [{channel_key[0]}]", aht_v, vol, target))
            row += 1
        row += 1

    chart = _build_combo_chart(
        ws, fmts, wb, chart_rows, _CHART_DATA_COL_B,
        title=f"Case Type AHT vs Target – W{week}",
        bar_name="Volume", line_name="AHT (min)",
    )
    return row, chart


# ── Blocco C ─────────────────────────────────────────────────────────────────

def _write_block_c(ws, fmts, wb, df: pd.DataFrame, week: str, start_col: int):
    """Primary Category summary. Ritorna (ultima_riga, chart)."""
    headers = ["Primary Category", "avg AHT", "volume", "target"]
    ws.merge_range(0, start_col, 0, start_col + len(headers) - 1,
                   f"Primary Category – W{week}", fmts["header"])
    for i, h in enumerate(headers):
        ws.write(1, start_col + i, h, fmts["header"])

    if "Primary Category" not in df.columns:
        return 2, None

    cfg_ph = CHANNEL_CONFIG["Phone"]
    cfg_nl = CHANNEL_CONFIG["Non-live"]
    target_blended = (cfg_ph["target_aht"] + cfg_nl["target_aht"]) / 2

    cat_agg = (
        df.groupby("Primary Category")
        .apply(lambda g: pd.Series({
            "vol": int(g["cases"].sum()),
            "aht": weighted_aht(g),
        }), include_groups=False)
        .reset_index()
        .sort_values("vol", ascending=False)
    )

    row = 2
    chart_rows: list[tuple] = []
    for _, cat_row in cat_agg.iterrows():
        ws.write(row, start_col, cat_row["Primary Category"], fmts["data_left"])
        ws.write_number(row, start_col + 1, cat_row["aht"], fmts["data_num"])
        ws.write_number(row, start_col + 2, cat_row["vol"], fmts["data_int"])
        ws.write_number(row, start_col + 3, target_blended, fmts["data_num"])
        chart_rows.append((cat_row["Primary Category"], cat_row["aht"],
                           int(cat_row["vol"]), target_blended))
        row += 1

    chart = _build_combo_chart(
        ws, fmts, wb, chart_rows, _CHART_DATA_COL_C,
        title=f"Primary Category AHT – W{week}",
        bar_name="Volume", line_name="AHT (min)",
    )
    return row, chart


# ── Costruzione grafico combinato (barre + 2 linee con UN solo combine) ───────

def _build_combo_chart(ws, fmts, wb, chart_rows: list[tuple], data_col: int,
                       title: str, bar_name: str, line_name: str):
    """
    Scrive i dati sorgente in colonne nascoste (data_col..data_col+3) e
    costruisce un grafico combinato:
      - barre = volume (asse secondario)
      - linea piena = AHT corrente (asse primario)
      - linea tratteggiata = target (asse primario)
    Entrambe le linee stanno in UN solo line chart, così basta un solo combine().
    """
    if len(chart_rows) < 2:
        return None

    # Intestazioni dati sorgente
    ws.write(0, data_col,     "_cat", fmts["data"])
    ws.write(0, data_col + 1, "_aht", fmts["data"])
    ws.write(0, data_col + 2, "_vol", fmts["data"])
    ws.write(0, data_col + 3, "_tgt", fmts["data"])
    for i, (cat, aht_v, vol, tgt) in enumerate(chart_rows, start=1):
        ws.write(i, data_col,     str(cat), fmts["data"])
        ws.write_number(i, data_col + 1, aht_v, fmts["data_num"])
        ws.write_number(i, data_col + 2, vol,   fmts["data_int"])
        ws.write_number(i, data_col + 3, tgt,   fmts["data_num"])

    n     = len(chart_rows)
    sname = ws.get_name()

    bar_chart  = wb.add_chart({"type": "column"})
    line_chart = wb.add_chart({"type": "line"})

    bar_chart.add_series({
        "name":       bar_name,
        "categories": [sname, 1, data_col,     n, data_col],
        "values":     [sname, 1, data_col + 2, n, data_col + 2],
        "fill":       {"color": "#4472C4"},
        "y2_axis":    True,
    })
    line_chart.add_series({
        "name":       line_name,
        "categories": [sname, 1, data_col,     n, data_col],
        "values":     [sname, 1, data_col + 1, n, data_col + 1],
        "line":       {"color": "#C00000", "width": 2.25},
        "marker":     {"type": "circle", "size": 5, "fill": {"color": "#C00000"}},
    })
    line_chart.add_series({
        "name":       "Target AHT",
        "categories": [sname, 1, data_col,     n, data_col],
        "values":     [sname, 1, data_col + 3, n, data_col + 3],
        "line":       {"color": "#70AD47", "width": 1.75, "dash_type": "dash"},
    })

    bar_chart.combine(line_chart)
    bar_chart.set_title({"name": title})
    bar_chart.set_x_axis({"name": "Categoria", "num_font": {"rotation": -45}})
    bar_chart.set_y_axis({"name": "AHT (min)"})
    bar_chart.set_y2_axis({"name": "Volume"})
    bar_chart.set_legend({"position": "bottom"})
    bar_chart.set_size({"width": 560, "height": 340})
    return bar_chart


# ── Helper ────────────────────────────────────────────────────────────────────

def _get_class(tbl: pd.DataFrame, channel: str, ct: str) -> str:
    row = tbl[(tbl["channel"] == channel) & (tbl["case_type"] == ct)]
    return row["classification"].iloc[0] if not row.empty else "Process Driven"


def _cls_fmt(wb, cls: str, num: bool = False):
    bg = CLASSIFICATION_COLORS.get(cls, "#FFFFFF")
    kw = {
        "font_name": "Calibri", "font_size": 10, "border": 1,
        "bg_color": bg, "align": "center", "valign": "vcenter",
    }
    if num:
        kw["num_format"] = "0.00"
    return wb.add_format(kw)
