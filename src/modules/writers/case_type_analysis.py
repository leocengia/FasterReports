"""
Foglio Case Type Analysis: tre blocchi tabellari + grafici barre/linea.
"""

import pandas as pd
from ..classification import build_tbl_class, get_relevant_case_types
from ..data_loader import CHANNEL_CONFIG, weighted_aht
from .styles import (
    make_formats, C_BLUE_HEADER, C_YELLOW_LIGHT, C_BLUE_LIGHT,
    C_RED_LIGHT, C_GREEN_LIGHT, C_LINE_AHT, C_LINE_TARGET,
    CLASSIFICATION_COLORS,
)


def write_case_type_analysis(wb, df: pd.DataFrame, week: str) -> None:
    ws   = wb.add_worksheet("Case Type Analysis")
    fmts = make_formats(wb)
    tbl  = build_tbl_class(df)

    # ── BLOCCO A (col 0-4): Case Origin (group) drill-down ───────────────────
    _write_block_a(ws, fmts, wb, df)

    # ── BLOCCO B (col 6-10): Case Type summary + grafico ─────────────────────
    chart_b_row = _write_block_b(ws, fmts, wb, df, tbl, week, start_col=6)

    # ── BLOCCO C (col 12-16): Primary Category summary + grafico ─────────────
    _write_block_c(ws, fmts, wb, df, tbl, week, start_col=12, chart_anchor_row=chart_b_row)

    # ── Larghezze ─────────────────────────────────────────────────────────────
    ws.set_column(0,  0,  28)
    ws.set_column(1,  4,  10)
    ws.set_column(5,  5,  2)   # separatore
    ws.set_column(6,  6,  28)
    ws.set_column(7,  10, 10)
    ws.set_column(11, 11, 2)   # separatore
    ws.set_column(12, 12, 28)
    ws.set_column(13, 16, 10)


# ── Blocco A ─────────────────────────────────────────────────────────────────

def _write_block_a(ws, fmts, wb, df: pd.DataFrame) -> None:
    """Case Origin (group) breakdown con subtotali per case type."""
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


# ── Blocco B ─────────────────────────────────────────────────────────────────

def _write_block_b(ws, fmts, wb, df: pd.DataFrame, tbl: pd.DataFrame,
                   week: str, start_col: int) -> int:
    """Case Type summary per canale + grafico combo barre/linea."""
    headers = ["Case Type", "avg AHT", "volume", "target", "class."]
    ws.merge_range(0, start_col, 0, start_col + len(headers) - 1,
                   f"Case Type Summary – W{week}", fmts["header"])
    for i, h in enumerate(headers):
        ws.write(1, start_col + i, h, fmts["header"])

    row = 2
    chart_data_rows: list[tuple] = []   # (ct, aht, vol, target, cls) per il grafico

    for channel_key, label, bg in [
        ("Phone",    "[ Phone ]",    C_YELLOW_LIGHT),
        ("Non-live", "[ Non-live ]", C_BLUE_LIGHT),
    ]:
        sub  = df[df["channel"] == channel_key]
        cfg  = CHANNEL_CONFIG[channel_key]
        target = cfg["target_aht"]
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
            ws.write(row, start_col,     ct,      cf)
            ws.write_number(row, start_col + 1, aht_v,  cfn)
            ws.write_number(row, start_col + 2, vol,    cf)
            ws.write_number(row, start_col + 3, target, cfn)
            ws.write(row, start_col + 4, cls,    cf)
            chart_data_rows.append((row, ct, aht_v, vol, target, cls))
            row += 1
        row += 1

    # ── Grafico combinato barre+linea ──────────────────────────────────────────
    chart_start_row = row + 1
    if chart_data_rows:
        # Scrivi dati grafico in una zona nascosta
        g_col = start_col + 7
        ws.write(0, g_col, "_ct",  fmts["data"])
        ws.write(0, g_col + 1, "_aht", fmts["data"])
        ws.write(0, g_col + 2, "_vol", fmts["data"])
        ws.write(0, g_col + 3, "_tgt", fmts["data"])
        for i, (xl_row, ct, aht_v, vol, tgt, _) in enumerate(chart_data_rows, start=1):
            ws.write(i, g_col,     ct,    fmts["data"])
            ws.write_number(i, g_col + 1, aht_v, fmts["data_num"])
            ws.write_number(i, g_col + 2, vol,   fmts["data_int"])
            ws.write_number(i, g_col + 3, tgt,   fmts["data_num"])

        n = len(chart_data_rows)
        sname = ws.get_name()

        bar_chart  = wb.add_chart({"type": "column"})
        line_chart = wb.add_chart({"type": "line"})
        tgt_chart  = wb.add_chart({"type": "line"})

        bar_chart.add_series({
            "name":       "Volume",
            "categories": [sname, 1, g_col,     n, g_col],
            "values":     [sname, 1, g_col + 2, n, g_col + 2],
            "fill":       {"color": "#4472C4"},
            "y2_axis":    True,
        })
        line_chart.add_series({
            "name":       "AHT (min)",
            "categories": [sname, 1, g_col,     n, g_col],
            "values":     [sname, 1, g_col + 1, n, g_col + 1],
            "line":       {"color": "#C00000", "width": 2},
            "marker":     {"type": "circle", "size": 5, "fill": {"color": "#C00000"}},
        })
        tgt_chart.add_series({
            "name":       "Target AHT",
            "categories": [sname, 1, g_col,     n, g_col],
            "values":     [sname, 1, g_col + 3, n, g_col + 3],
            "line":       {"color": "#70AD47", "width": 1.5, "dash_type": "dash"},
        })

        bar_chart.combine(line_chart)
        bar_chart.combine(tgt_chart)
        bar_chart.set_title({"name": f"Case Type AHT vs Target – W{week}"})
        bar_chart.set_x_axis({"name": "Case Type"})
        bar_chart.set_y_axis({"name": "AHT (min)"})
        bar_chart.set_y2_axis({"name": "Volume"})
        bar_chart.set_size({"width": 480, "height": 300})
        ws.insert_chart(chart_start_row, start_col, bar_chart, {"x_offset": 0, "y_offset": 0})

    return chart_start_row


# ── Blocco C ─────────────────────────────────────────────────────────────────

def _write_block_c(ws, fmts, wb, df: pd.DataFrame, tbl: pd.DataFrame,
                   week: str, start_col: int, chart_anchor_row: int) -> None:
    """Primary Category summary + grafico."""
    headers = ["Primary Category", "avg AHT", "volume", "target"]
    ws.merge_range(0, start_col, 0, start_col + len(headers) - 1,
                   f"Primary Category – W{week}", fmts["header"])
    for i, h in enumerate(headers):
        ws.write(1, start_col + i, h, fmts["header"])

    cfg_ph = CHANNEL_CONFIG["Phone"]
    cfg_nl = CHANNEL_CONFIG["Non-live"]
    # Usa target blended (media semplicistica)
    target_blended = (cfg_ph["target_aht"] + cfg_nl["target_aht"]) / 2

    row = 2
    if "Primary Category" not in df.columns:
        return

    cat_agg = (
        df.groupby("Primary Category")
        .apply(lambda g: pd.Series({
            "vol":    int(g["cases"].sum()),
            "aht":    weighted_aht(g),
        }), include_groups=False)
        .reset_index()
        .sort_values("vol", ascending=False)
    )

    g_col = start_col + 6
    ws.write(0, g_col, "_pc", fmts["data"])
    ws.write(0, g_col + 1, "_aht", fmts["data"])
    ws.write(0, g_col + 2, "_vol", fmts["data"])
    ws.write(0, g_col + 3, "_tgt", fmts["data"])

    for i, (_, cat_row) in enumerate(cat_agg.iterrows(), start=1):
        ws.write(row, start_col, cat_row["Primary Category"], fmts["data_left"])
        ws.write_number(row, start_col + 1, cat_row["aht"], fmts["data_num"])
        ws.write_number(row, start_col + 2, cat_row["vol"], fmts["data_int"])
        ws.write_number(row, start_col + 3, target_blended, fmts["data_num"])
        ws.write(i, g_col,     cat_row["Primary Category"], fmts["data"])
        ws.write_number(i, g_col + 1, cat_row["aht"],   fmts["data_num"])
        ws.write_number(i, g_col + 2, cat_row["vol"],   fmts["data_int"])
        ws.write_number(i, g_col + 3, target_blended,   fmts["data_num"])
        row += 1

    n = len(cat_agg)
    if n < 2:
        return
    sname = ws.get_name()

    bar_c  = wb.add_chart({"type": "column"})
    line_c = wb.add_chart({"type": "line"})
    tgt_c  = wb.add_chart({"type": "line"})

    bar_c.add_series({
        "name":       "Volume",
        "categories": [sname, 1, g_col,     n, g_col],
        "values":     [sname, 1, g_col + 2, n, g_col + 2],
        "fill":       {"color": "#4472C4"},
        "y2_axis":    True,
    })
    line_c.add_series({
        "name":       "AHT (min)",
        "categories": [sname, 1, g_col,     n, g_col],
        "values":     [sname, 1, g_col + 1, n, g_col + 1],
        "line":       {"color": "#C00000", "width": 2},
        "marker":     {"type": "circle", "size": 5, "fill": {"color": "#C00000"}},
    })
    tgt_c.add_series({
        "name":       "Target AHT",
        "categories": [sname, 1, g_col,     n, g_col],
        "values":     [sname, 1, g_col + 3, n, g_col + 3],
        "line":       {"color": "#70AD47", "width": 1.5, "dash_type": "dash"},
    })

    bar_c.combine(line_c)
    bar_c.combine(tgt_c)
    bar_c.set_title({"name": f"Primary Category AHT – W{week}"})
    bar_c.set_x_axis({"name": "Category"})
    bar_c.set_y_axis({"name": "AHT (min)"})
    bar_c.set_y2_axis({"name": "Volume"})
    bar_c.set_size({"width": 480, "height": 300})
    ws.insert_chart(chart_anchor_row, start_col, bar_c, {"x_offset": 0, "y_offset": 0})


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
