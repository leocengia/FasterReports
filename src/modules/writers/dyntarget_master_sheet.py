"""
Foglio DynTarget Master: tabellona con tutti i Channel × Case Type.
"""

import pandas as pd
from ..data_loader import CHANNEL_CONFIG
from .styles import make_formats, C_YELLOW_LIGHT, C_BLUE_LIGHT, C_RED_LIGHT, C_GREEN_LIGHT


def write_dyntarget_master(wb, dyn_targets: dict, channel_summary: dict,
                            week: int, week_date: str) -> None:
    ws   = wb.add_worksheet("DynTarget Master")
    fmts = make_formats(wb)

    headers = ["Channel", "Case Type", "Volume", "Avg AHT (min)",
               "DynTarget (min)", "Gap (sec)", "Gap (%)",
               "AHT Target", "Week", "Date"]
    ws.merge_range(0, 0, 0, len(headers) - 1,
                   "Dynamic Targets Master — tutti i Case Type", fmts["header"])
    for i, h in enumerate(headers):
        ws.write(1, i, h, fmts["header"])

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
    ch_fmt_map = {
        "Phone":    wb.add_format({"font_name": "Calibri", "font_size": 10, "border": 1,
                                   "bg_color": C_YELLOW_LIGHT, "bold": True, "align": "center"}),
        "Non-live": wb.add_format({"font_name": "Calibri", "font_size": 10, "border": 1,
                                   "bg_color": C_BLUE_LIGHT,   "bold": True, "align": "center"}),
    }

    row = 2
    for channel in ("Phone", "Non-live"):
        ct_data = dyn_targets.get(channel, {})
        summary  = channel_summary.get(channel, {})
        cfg      = CHANNEL_CONFIG[channel]
        target   = cfg["target_aht"]
        ch_f     = ch_fmt_map[channel]

        for ct in sorted(ct_data.keys()):
            curr     = summary.get(ct, {})
            vol      = curr.get("volume", 0)
            curr_aht = curr.get("avg_aht", 0.0)
            dyn      = ct_data[ct]["dyn_target"]
            gap_sec  = round((curr_aht - dyn) * 60, 2)
            gap_pct  = gap_sec / (dyn * 60) if dyn else 0.0
            above    = curr_aht > dyn

            rf_num = above_num if above else below_num
            rf_pct = above_pct if above else below_pct

            ws.write(row, 0, channel,   ch_f)
            ws.write(row, 1, ct,        fmts["data_left"])
            ws.write_number(row, 2, vol,      fmts["data_int"])
            ws.write_number(row, 3, curr_aht, rf_num)
            ws.write_number(row, 4, dyn,      rf_num)
            ws.write_number(row, 5, gap_sec,  rf_num)
            ws.write_number(row, 6, gap_pct,  rf_pct)
            ws.write_number(row, 7, target,   fmts["data_num"])
            ws.write(row, 8, str(week),       fmts["data"])
            ws.write(row, 9, week_date,       fmts["data"])
            row += 1
        row += 1  # riga vuota tra canali

    ws.set_column(0, 0, 14)
    ws.set_column(1, 1, 30)
    ws.set_column(2, 9, 14)
    ws.freeze_panes(2, 2)
    ws.autofilter(1, 0, row, len(headers) - 1)
