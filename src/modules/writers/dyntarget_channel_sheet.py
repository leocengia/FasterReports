"""
Fogli DynTarget – Phone e DynTarget – Non-live.
Tabella completa + grafico a barre verticali con linea DynTarget.
"""

import pandas as pd
from ..data_loader import CHANNEL_CONFIG
from .styles import make_formats, C_YELLOW_LIGHT, C_BLUE_LIGHT, C_RED_LIGHT, C_GREEN_LIGHT


def write_dyntarget_channel(wb, df: pd.DataFrame, dyn_targets: dict,
                             channel_summary: dict, channel: str) -> None:
    """
    channel: "Phone" o "Non-live"
    """
    sheet_name = f"DynTarget – {channel}"
    ws   = wb.add_worksheet(sheet_name)
    fmts = make_formats(wb)

    cfg    = CHANNEL_CONFIG[channel]
    target = cfg["target_aht"]
    bg     = C_YELLOW_LIGHT if channel == "Phone" else C_BLUE_LIGHT

    ct_data = dyn_targets.get(channel, {})
    summary  = channel_summary.get(channel, {})

    # ── Titolo ──────────────────────────────────────────────────────────────
    title_fmt = wb.add_format({
        "font_name": "Calibri", "font_size": 13, "bold": True,
        "bg_color": bg, "align": "center", "valign": "vcenter", "border": 0,
    })
    ws.merge_range(0, 0, 0, 6, f"Dynamic Targets — {channel}", title_fmt)
    ws.set_row(0, 22)

    # ── Intestazione tabella ─────────────────────────────────────────────────
    headers = ["Case Type", "Volume", "Avg AHT (min)", "DynTarget (min)",
               "Gap (sec)", "Gap (%)", "AHT Target"]
    for i, h in enumerate(headers):
        ws.write(1, i, h, fmts["header"])

    # ── Dati tabella ─────────────────────────────────────────────────────────
    rows_data: list[dict] = []
    for ct in sorted(ct_data.keys()):
        curr = summary.get(ct, {})
        vol      = curr.get("volume", 0)
        curr_aht = curr.get("avg_aht", 0.0)
        dyn      = ct_data[ct]["dyn_target"]
        gap_sec  = round((curr_aht - dyn) * 60, 2)
        gap_pct  = gap_sec / (dyn * 60) if dyn else 0.0
        rows_data.append({
            "ct":       ct,
            "vol":      vol,
            "curr_aht": curr_aht,
            "dyn":      dyn,
            "gap_sec":  gap_sec,
            "gap_pct":  gap_pct,
            "target":   target,
            "above":    curr_aht > dyn,
        })

    # Ordina per AHT corrente desc
    rows_data.sort(key=lambda x: x["curr_aht"], reverse=True)

    above_fmt = wb.add_format({
        "font_name": "Calibri", "font_size": 10, "border": 1,
        "bg_color": C_RED_LIGHT, "align": "center", "valign": "vcenter",
    })
    below_fmt = wb.add_format({
        "font_name": "Calibri", "font_size": 10, "border": 1,
        "bg_color": C_GREEN_LIGHT, "align": "center", "valign": "vcenter",
    })
    above_num = wb.add_format({
        "font_name": "Calibri", "font_size": 10, "border": 1,
        "bg_color": C_RED_LIGHT, "align": "center", "num_format": "0.00",
    })
    below_num = wb.add_format({
        "font_name": "Calibri", "font_size": 10, "border": 1,
        "bg_color": C_GREEN_LIGHT, "align": "center", "num_format": "0.00",
    })
    above_pct = wb.add_format({
        "font_name": "Calibri", "font_size": 10, "border": 1,
        "bg_color": C_RED_LIGHT, "align": "center", "num_format": "0.0%",
    })
    below_pct = wb.add_format({
        "font_name": "Calibri", "font_size": 10, "border": 1,
        "bg_color": C_GREEN_LIGHT, "align": "center", "num_format": "0.0%",
    })

    for r_i, rd in enumerate(rows_data, start=2):
        rf     = above_fmt if rd["above"] else below_fmt
        rf_num = above_num if rd["above"] else below_num
        rf_pct = above_pct if rd["above"] else below_pct
        ws.write(r_i, 0, rd["ct"],       fmts["data_left"])
        ws.write_number(r_i, 1, rd["vol"],      fmts["data_int"])
        ws.write_number(r_i, 2, rd["curr_aht"], rf_num)
        ws.write_number(r_i, 3, rd["dyn"],      rf_num)
        ws.write_number(r_i, 4, rd["gap_sec"],  rf_num)
        ws.write_number(r_i, 5, rd["gap_pct"],  rf_pct)
        ws.write_number(r_i, 6, rd["target"],   fmts["data_num"])

    n_rows = len(rows_data)
    table_end_row = 2 + n_rows

    # ── Dati ausiliari per grafico ────────────────────────────────────────────
    g_col = 9
    ws.write(1, g_col,     "_ct",    fmts["data"])
    ws.write(1, g_col + 1, "_aht",   fmts["data"])
    ws.write(1, g_col + 2, "_dyn",   fmts["data"])

    for i, rd in enumerate(rows_data, start=2):
        ws.write(i, g_col,     rd["ct"],       fmts["data"])
        ws.write_number(i, g_col + 1, rd["curr_aht"], fmts["data_num"])
        ws.write_number(i, g_col + 2, rd["dyn"],       fmts["data_num"])

    # ── Grafico barre + linea DynTarget ──────────────────────────────────────
    if n_rows >= 2:
        sname = ws.get_name()
        bar   = wb.add_chart({"type": "column"})
        line  = wb.add_chart({"type": "line"})

        bar.add_series({
            "name":       "Avg AHT",
            "categories": [sname, 2, g_col,     2 + n_rows - 1, g_col],
            "values":     [sname, 2, g_col + 1, 2 + n_rows - 1, g_col + 1],
            "fill":       {"color": "#4472C4"},
            "data_labels": {"value": True, "num_format": "0.0"},
        })
        line.add_series({
            "name":       "DynTarget",
            "categories": [sname, 2, g_col,     2 + n_rows - 1, g_col],
            "values":     [sname, 2, g_col + 2, 2 + n_rows - 1, g_col + 2],
            "line":       {"color": "#C00000", "width": 2.5},
            "marker":     {"type": "diamond", "size": 6,
                           "fill": {"color": "#C00000"},
                           "border": {"color": "#C00000"}},
        })

        bar.combine(line)
        bar.set_title({"name": f"Avg AHT vs DynTarget — {channel}"})
        bar.set_x_axis({"name": "Case Type", "text_axis": True})
        bar.set_y_axis({"name": "AHT (min)"})
        bar.set_legend({"position": "bottom"})
        bar.set_size({"width": 680, "height": 400})
        ws.insert_chart(table_end_row + 2, 0, bar)

    # ── Top-15 agenti per ogni case type ─────────────────────────────────────
    detail_row = table_end_row + 22
    ws.merge_range(detail_row, 0, detail_row, 6,
                   "Top-15 Agenti per Case Type (base DynTarget)", fmts["header"])
    detail_row += 1

    for ct, data in dyn_targets.get(channel, {}).items():
        top = data["top_agents"]
        ws.merge_range(detail_row, 0, detail_row, 6, ct, fmts["section_phone"] if channel == "Phone" else fmts["section_nl"])
        detail_row += 1
        ws.write(detail_row, 0, "Agent",      fmts["header"])
        ws.write(detail_row, 1, "Avg AHT",    fmts["header"])
        ws.write(detail_row, 2, "Volume",     fmts["header"])
        ws.write(detail_row, 3, "AHT × Vol",  fmts["header"])
        detail_row += 1
        for _, ag in top.iterrows():
            ws.write(detail_row, 0, ag["agent"],   fmts["data_left"])
            ws.write_number(detail_row, 1, ag["avg_aht"],            fmts["data_num"])
            ws.write_number(detail_row, 2, int(ag["volume"]),        fmts["data_int"])
            ws.write_number(detail_row, 3, ag["avg_aht"] * ag["volume"], fmts["data_num"])
            detail_row += 1
        # Totali + DynTarget
        dyn_v = data["dyn_target"]
        tot_vol = int(top["volume"].sum())
        ws.write(detail_row, 0, "DynTarget (med. pesata)", fmts["bold_left"])
        ws.write_number(detail_row, 1, dyn_v, fmts["bold_num"])
        ws.write_number(detail_row, 2, tot_vol, fmts["bold"])
        detail_row += 2

    # ── Larghezze ─────────────────────────────────────────────────────────────
    ws.set_column(0, 0, 30)
    ws.set_column(1, 6, 14)
    ws.set_column(7, 8, 2)
    ws.set_column(g_col, g_col,     30)
    ws.set_column(g_col + 1, g_col + 2, 12)
