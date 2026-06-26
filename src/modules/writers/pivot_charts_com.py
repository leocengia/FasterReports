"""
Post-elaborazione COM (pywin32) del foglio 'Case Type Analysis'.

Riapre il workbook con Excel e aggiunge:
  - una tabella pivot interattiva (ancorata in A1), origine = foglio DATASET
  - due grafici (a destra delle tabelle G-J) applicando il modello .crtx

Richiede Excel installato e la libreria pywin32. Se non disponibili, salta con
un warning SENZA interrompere la generazione del report (le tabelle G-J restano).

NOTA: questo passo deve essere l'ULTIMO writer sul file. openpyxl (usato dal
drill-down) non preserva pivot/grafici creati da Excel, quindi il drill-down va
eseguito prima di questa funzione.
"""

import os

# ── Costanti Excel (late binding → vanno definite come literal) ──────────────
_xlDatabase   = 1
_xlPageField  = 3
_xlRowField   = 1
_xlAverage    = -4106
_xlCount      = -4112
_xlCompactRow = 0


def add_pivot_and_charts(out_path, crtx_path, layout: dict, theme_path=None) -> None:
    """Aggiunge pivot + grafici a out_path via Excel COM. No-op se pywin32/Excel mancano."""
    try:
        import win32com.client as win32
    except ImportError:
        print("⚠ pywin32 non installato: pivot e grafici del foglio 'Case Type Analysis' "
              "saltati. Esegui setup.bat per installarlo.")
        return

    out_abs  = os.path.abspath(str(out_path))
    crtx_abs = os.path.abspath(str(crtx_path))
    has_crtx = os.path.exists(crtx_abs)
    if not has_crtx:
        print(f"⚠ Modello grafico non trovato: {crtx_abs}\n"
              "  I grafici verranno creati senza lo stile .crtx.")

    excel = None
    wb = None
    try:
        excel = win32.DispatchEx("Excel.Application")   # istanza dedicata
        excel.Visible = False
        excel.DisplayAlerts = False

        wb  = excel.Workbooks.Open(out_abs)

        # Applica il tema colori (EG CC) all'intero workbook, se fornito.
        if theme_path:
            theme_abs = os.path.abspath(str(theme_path))
            if os.path.exists(theme_abs):
                try:
                    wb.ApplyTheme(theme_abs)
                except Exception as e:
                    print(f"⚠ Tema non applicato: {e}")
            else:
                print(f"⚠ Tema non trovato: {theme_abs}")

        cta = wb.Sheets("Case Type Analysis")
        ds  = wb.Sheets("DATASET")

        # I grafici vanno creati PRIMA della pivot: con una PivotTable attiva sul
        # foglio, Excel tenterebbe di creare un PivotChart e SetSourceData su un
        # range non-pivot fallirebbe (DISP_E_EXCEPTION).
        if layout.get("casetype_range"):
            _build_chart(cta, layout["casetype_range"], layout.get("chart1_anchor", "L2"),
                         "Avg. AHT by Case Type", crtx_abs if has_crtx else None)
        if layout.get("primarycat_range"):
            _build_chart(cta, layout["primarycat_range"], layout.get("chart2_anchor", "L20"),
                         "Avg. AHT by Primary Category", crtx_abs if has_crtx else None)

        _build_pivot(wb, cta, ds)

        wb.Save()
        wb.Close(SaveChanges=False)
        wb = None
        print("✓ Pivot e grafici aggiunti al foglio 'Case Type Analysis'.")
    except Exception as e:
        print(f"⚠ Errore COM durante pivot/grafici: {e}\n"
              "  Il report è stato comunque salvato (senza pivot/grafici).\n"
              "  Verifica che Excel sia installato e che il file non sia già aperto.")
    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if excel is not None:
                excel.Quit()
        except Exception:
            pass


def _build_pivot(wb, cta, ds) -> None:
    """Crea la pivot in A1: filtro Case Origin, righe Case Type→Primary Category, valori avg+count."""
    src = f"DATASET!{ds.UsedRange.Address}"
    pc  = wb.PivotCaches().Create(_xlDatabase, src)
    pt  = pc.CreatePivotTable(cta.Range("A1"), "ptCaseType")

    pt.PivotFields("Case Origin (group)").Orientation = _xlPageField

    f_ct = pt.PivotFields("Case Type")
    f_ct.Orientation = _xlRowField
    f_ct.Position = 1

    # Primary Category come secondo livello di riga (se presente nel dataset)
    try:
        f_pc = pt.PivotFields("Primary Category")
        f_pc.Orientation = _xlRowField
        f_pc.Position = 2
    except Exception:
        pass

    pt.AddDataField(pt.PivotFields("Case AHT (mins)"), "Avg AHT", _xlAverage)
    pt.AddDataField(pt.PivotFields("case_number"), "Count", _xlCount)

    try:
        pt.RowAxisLayout(_xlCompactRow)
    except Exception:
        pass

    # Media AHT a 2 decimali. Il formato va messo sul DataField (persiste al
    # refresh della pivot) col separatore decimale del LOCALE di Excel: in
    # italiano "0.00" viene rifiutato perché '.' è separatore delle migliaia,
    # quindi serve "0,00". _set_decimals auto-rileva il caso.
    # NB: va ri-prelevato fresco (DataFields.Item(1)); il riferimento restituito
    # da AddDataField non accetta più il set dopo RowAxisLayout.
    try:
        _set_decimals(pt.DataFields.Item(1), 2)
    except Exception:
        pass


def _set_decimals(data_field, n: int) -> None:
    """
    Imposta n decimali su un campo dati pivot, adattandosi al separatore decimale
    del locale di Excel (en-US usa '.', it-IT usa ','). Prova prima '.': se Excel
    lo rifiuta (locale con ',' decimale), ripiega su ','.
    """
    us = "0." + "0" * n
    try:
        data_field.NumberFormat = us
        if data_field.NumberFormat == us:
            return
    except Exception:
        pass
    try:
        data_field.NumberFormat = "0," + "0" * n
    except Exception:
        pass


def _build_chart(cta, data_range: str, anchor: str, title: str, crtx_abs) -> None:
    """Crea un grafico dal range dati e (se disponibile) applica il modello .crtx."""
    anchor_cell = cta.Range(anchor)
    co = cta.ChartObjects().Add(anchor_cell.Left, anchor_cell.Top, 480, 300)
    ch = co.Chart
    ch.SetSourceData(cta.Range(data_range))
    if crtx_abs:
        ch.ApplyChartTemplate(crtx_abs)
    try:
        ch.HasTitle = True
        ch.ChartTitle.Text = title
    except Exception:
        pass
