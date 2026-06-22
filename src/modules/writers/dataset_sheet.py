"""
Scrive il foglio DATASET (copia 1:1 del CSV raw).
"""

import pandas as pd
from .styles import make_formats, C_BLUE_HEADER


def write_dataset(wb, df_raw: pd.DataFrame) -> None:
    """Scrive il foglio DATASET da un DataFrame (colonne originali del CSV)."""
    ws  = wb.add_worksheet("DATASET")
    fmts = make_formats(wb)

    # ── intestazione ──────────────────────────────────────────────────────────
    for col_idx, col_name in enumerate(df_raw.columns):
        ws.write(0, col_idx, col_name, fmts["header"])

    # ── dati ──────────────────────────────────────────────────────────────────
    data_fmt  = wb.add_format({"font_name": "Calibri", "font_size": 9, "border": 0})
    num_fmt   = wb.add_format({"font_name": "Calibri", "font_size": 9, "num_format": "0.##"})

    for row_idx, row in enumerate(df_raw.itertuples(index=False), start=1):
        for col_idx, val in enumerate(row):
            if val is None or (isinstance(val, float) and pd.isna(val)):
                ws.write_blank(row_idx, col_idx, None, data_fmt)
            elif isinstance(val, (int, float)):
                ws.write_number(row_idx, col_idx, float(val), num_fmt)
            else:
                ws.write_string(row_idx, col_idx, str(val), data_fmt)

    # ── larghezze colonne ─────────────────────────────────────────────────────
    ws.set_column(0, len(df_raw.columns) - 1, 18)
    ws.freeze_panes(1, 0)
