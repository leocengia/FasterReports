"""
Foglio Progress Tracking: storico settimanale per case type tracciati.
"""

import pandas as pd
from ..dynamic_targets import get_progress_data, get_tracked, CHANNELS
from .styles import make_formats, C_YELLOW_LIGHT, C_BLUE_LIGHT, C_GREEN_LIGHT, C_RED_LIGHT


def write_progress(wb, history: dict, week: int) -> None:
    """
    Scrive il foglio Progress Tracking con, per ogni Channel × Case Type
    SELEZIONATO (history[_config][tracked]):
    - Tabella settimanale AHT vs DynTarget
    - Grafico linea AHT vs DynTarget

    Lo storico contiene tutti i case type, ma qui mostriamo solo quelli scelti
    dall'utente con --track, così il foglio resta leggibile.
    """
    ws   = wb.add_worksheet("Progress Tracking")
    fmts = make_formats(wb)

    ws.merge_range(0, 0, 0, 9, "Progress Tracking — Dynamic Targets", fmts["header"])
    ws.set_row(0, 20)

    g_col   = 11   # colonne dati sorgente grafici (fuori dalla tabella visibile)
    cur_row = 2

    tracked = set(get_tracked(history))
    if not tracked:
        ws.merge_range(2, 0, 2, 9,
                       "Nessun case type selezionato. Usa: generate ... "
                       "--track \"Case Type 1\" \"Case Type 2\"",
                       fmts["data_left"])
        ws.set_column(0, 9, 16)
        return

    for channel in CHANNELS:
        ch_history = history.get(channel, {})
        if not ch_history:
            continue

        bg = C_YELLOW_LIGHT if channel == "Phone" else C_BLUE_LIGHT
        ch_title_fmt = wb.add_format({
            "font_name": "Calibri", "font_size": 11, "bold": True,
            "bg_color": bg, "align": "center", "valign": "vcenter", "border": 1,
        })

        for ct in sorted(t for t in tracked if t in ch_history):
            weeks = get_progress_data(history, channel, ct)
            if not weeks:
                continue

            # Titolo sezione
            ws.merge_range(cur_row, 0, cur_row, 9,
                           f"{channel} — {ct}", ch_title_fmt)
            cur_row += 1

            # Header tabella
            cols = ["Week", "Date", "AHT (min)", "AHT (sec)",
                    "DynTarget (min)", "DynTarget (sec)",
                    "Δ vs prev (sec)", "On Track?", "Volume"]
            for i, h in enumerate(cols):
                ws.write(cur_row, i, h, fmts["header"])
            cur_row += 1

            # Dati + grafici ausiliari
            ws.write(cur_row - 1, g_col, "_week",      fmts["data"])
            ws.write(cur_row - 1, g_col + 1, "_aht_min",  fmts["data"])
            ws.write(cur_row - 1, g_col + 2, "_dyn_min",  fmts["data"])

            data_start_row = cur_row
            for wd in weeks:
                delta   = wd.get("delta_sec")
                on_track = (delta is not None and delta >= 0)
                bg_row  = C_GREEN_LIGHT if on_track else C_RED_LIGHT

                def rfmt(num=False, pct=False):
                    kw = {"font_name": "Calibri", "font_size": 10, "border": 1,
                          "bg_color": bg_row, "align": "center", "valign": "vcenter"}
                    if num:
                        kw["num_format"] = "0.00"
                    if pct:
                        kw["num_format"] = "0.0%"
                    return wb.add_format(kw)

                ws.write(cur_row, 0, f"W{wd['week']}",       rfmt())
                ws.write(cur_row, 1, wd.get("date", ""),     rfmt())
                ws.write_number(cur_row, 2, wd["aht_min"],   rfmt(num=True))
                ws.write_number(cur_row, 3, wd["aht_sec"],   rfmt(num=True))
                ws.write_number(cur_row, 4, wd["dyn_target_min"], rfmt(num=True))
                ws.write_number(cur_row, 5, wd["dyn_target_sec"], rfmt(num=True))
                if delta is not None:
                    ws.write_number(cur_row, 6, delta, rfmt(num=True))
                else:
                    ws.write(cur_row, 6, "—", rfmt())
                ws.write(cur_row, 7, "✓" if on_track else "✗", rfmt())
                ws.write_number(cur_row, 8, wd.get("volume", 0), rfmt())

                # Dati grafico
                g_row = cur_row
                ws.write(g_row, g_col, f"W{wd['week']}",             fmts["data"])
                ws.write_number(g_row, g_col + 1, wd["aht_min"],     fmts["data_num"])
                ws.write_number(g_row, g_col + 2, wd["dyn_target_min"], fmts["data_num"])
                cur_row += 1

            data_end_row = cur_row - 1
            n_weeks = data_end_row - data_start_row + 1

            # Grafico linea AHT vs DynTarget
            if n_weeks >= 2:
                sname = ws.get_name()
                chart = wb.add_chart({"type": "line"})
                chart.add_series({
                    "name":       "AHT (min)",
                    "categories": [sname, data_start_row, g_col, data_end_row, g_col],
                    "values":     [sname, data_start_row, g_col + 1, data_end_row, g_col + 1],
                    "line":       {"color": "#C00000", "width": 2},
                    "marker":     {"type": "circle", "size": 5, "fill": {"color": "#C00000"}},
                })
                chart.add_series({
                    "name":       "DynTarget",
                    "categories": [sname, data_start_row, g_col, data_end_row, g_col],
                    "values":     [sname, data_start_row, g_col + 2, data_end_row, g_col + 2],
                    "line":       {"color": "#70AD47", "width": 2, "dash_type": "dash"},
                    "marker":     {"type": "diamond", "size": 5, "fill": {"color": "#70AD47"}},
                })
                chart.set_title({"name": f"{channel} — {ct} — Progress W{week}"})
                chart.set_x_axis({"name": "Week"})
                chart.set_y_axis({"name": "AHT (min)"})
                chart.set_legend({"position": "bottom"})
                chart.set_size({"width": 480, "height": 280})
                ws.insert_chart(cur_row + 1, 0, chart)
                cur_row += 18

            cur_row += 3  # spazio tra sezioni

    ws.set_column(0, 1, 10)
    ws.set_column(2, 8, 14)
    ws.set_column(10, 10, 3)
    ws.set_column(g_col, g_col + 2, 12, None, {"hidden": True})
