"""
Foglio Case Type Analysis.

Layout:
  - Col A-C : tabella pivot interattiva (aggiunta DOPO, via COM — vedi pivot_charts_com)
              origine = foglio DATASET; filtro Case Origin (group); righe Case Type
              → Primary Category; valori media Case AHT (mins) + conteggio case_number.
  - Col D-F : separatore vuoto.
  - Col G-J : due tabelle calcolate (Case Type, Primary Category), top 5 per volume.
  - Col L+  : un grafico per tabella (aggiunto via COM applicando il modello .crtx).

Questa funzione scrive SOLO le tabelle G-J con xlsxwriter e ritorna le coordinate
(range tabelle + ancore grafici) che il post-process COM userà per pivot e grafici.
"""

from xlsxwriter.utility import xl_rowcol_to_cell

from ..data_loader import agg_simple
from .styles import make_formats

# ── Coordinate layout (0-based) ──────────────────────────────────────────────
_TABLE_COL     = 6     # colonna G
_T1_HDR_ROW    = 1     # riga 2: header tabella Case Type
_T2_HDR_ROW    = 9     # riga 10: header tabella Primary Category
_TOP_N         = 5
_CHART1_ANCHOR = "L2"
_CHART2_ANCHOR = "L20"


def write_case_type_analysis(wb, df_raw, week, target_phone, target_nonlive,
                             selected_case_types, dd):
    """
    Scrive le tabelle G-J del foglio Case Type Analysis.
    La tabella Case Type usa i case type selezionati dal deepdive (ordine Score),
    con avg_aht/volume presi dal blocco 'All' di `dd` (coincidono col deepdive).
    Ritorna un dict con i range (formato A1) per pivot/grafici COM.
    """
    ws   = wb.add_worksheet("Case Type Analysis")
    fmts = make_formats(wb)
    target_blended = (float(target_phone) + float(target_nonlive)) / 2.0

    # Tabella 1 — Case Type (selezionati dal deepdive)
    n1 = _write_casetype_table(ws, fmts, dd, selected_case_types, target_blended)
    casetype_range = _range(_T1_HDR_ROW, n1) if n1 else None

    # Tabella 2 — Primary Category (se la colonna esiste)
    primarycat_range = None
    if "Primary Category" in df_raw.columns:
        n2 = _write_table(ws, fmts, df_raw, by_col="Primary Category",
                          hdr_row=_T2_HDR_ROW, label_hdr="primary_category", target=target_blended)
        primarycat_range = _range(_T2_HDR_ROW, n2) if n2 else None
    else:
        print("⚠ Colonna 'Primary Category' assente: tabella/grafico Primary Category saltati.")

    # ── Larghezze colonne ─────────────────────────────────────────────────────
    ws.set_column(0, 2, 22)                              # A-C: spazio per la pivot
    ws.set_column(_TABLE_COL, _TABLE_COL, 28)            # G: label
    ws.set_column(_TABLE_COL + 1, _TABLE_COL + 3, 12)    # H-J: avg/volume/target

    return {
        "casetype_range":   casetype_range,
        "primarycat_range": primarycat_range,
        "chart1_anchor":    _CHART1_ANCHOR,
        "chart2_anchor":    _CHART2_ANCHOR,
    }


# ── Helper ────────────────────────────────────────────────────────────────────

def _write_casetype_table(ws, fmts, dd, selected, target) -> int:
    """
    Tabella Case Type coi case type selezionati dal deepdive (nell'ordine dato),
    con avg_aht/volume dal blocco 'All' di dd. Ritorna il numero di righe scritte.
    """
    headers = ["case_type", "avg_aht", "volume", "target"]
    for i, h in enumerate(headers):
        ws.write(_T1_HDR_ROW, _TABLE_COL + i, h, fmts["header"])

    allr = dd[dd["channel"] == "All"].set_index("case_type") if dd is not None and not dd.empty else None
    r = _T1_HDR_ROW + 1
    n = 0
    for ct in (selected or []):
        if allr is None or ct not in allr.index:
            continue
        ws.write(r, _TABLE_COL,            str(ct),                          fmts["data_left"])
        ws.write_number(r, _TABLE_COL + 1, float(allr.loc[ct, "aht"]),       fmts["data_num"])
        ws.write_number(r, _TABLE_COL + 2, int(allr.loc[ct, "volume"]),      fmts["data_int"])
        ws.write_number(r, _TABLE_COL + 3, float(target),                    fmts["data_num"])
        r += 1
        n += 1
    return n


def _write_table(ws, fmts, df_raw, by_col, hdr_row, label_hdr, target) -> int:
    """
    Scrive una tabella top-N (header + righe) a partire da hdr_row, colonna G.
    Colonne: <label> | avg_aht | volume | target.
    Ritorna il numero di righe dati scritte.
    """
    agg = agg_simple(df_raw, by_col).head(_TOP_N)

    headers = [label_hdr, "avg_aht", "volume", "target"]
    for i, h in enumerate(headers):
        ws.write(hdr_row, _TABLE_COL + i, h, fmts["header"])

    r = hdr_row + 1
    for _, row in agg.iterrows():
        ws.write(r, _TABLE_COL,            str(row[by_col]),        fmts["data_left"])
        ws.write_number(r, _TABLE_COL + 1, float(row["avg_aht"]),   fmts["data_num"])
        ws.write_number(r, _TABLE_COL + 2, int(row["volume"]),      fmts["data_int"])
        ws.write_number(r, _TABLE_COL + 3, float(target),           fmts["data_num"])
        r += 1
    return len(agg)


def _range(hdr_row: int, n_rows: int) -> str:
    """Indirizzo A1 del range (header incluso) per il grafico, su 4 colonne da G."""
    top_left  = xl_rowcol_to_cell(hdr_row, _TABLE_COL)
    bot_right = xl_rowcol_to_cell(hdr_row + n_rows, _TABLE_COL + 3)
    return f"{top_left}:{bot_right}"
