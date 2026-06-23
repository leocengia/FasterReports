"""
Genera case_KPI_report_W{N}.xlsx a partire dai dati raw AHT.

Comandi:
    generate  — crea il report settimanale completo
    drill     — aggiunge il foglio drill-down xLori & Costa a un report esistente

Uso:
    python src/case_kpi_report.py generate "raw_data/AHT blended WEEK 24.csv" --week 24
    python src/case_kpi_report.py drill --output output/case_KPI_report_W24.xlsx \\
                                        --type "Rates & Inventory Changes"
"""

import argparse
import sys
from datetime import date, timedelta
from pathlib import Path

import xlsxwriter

from modules.data_loader import load_csv, DataError
from modules.classification import build_tbl_class
from modules.dynamic_targets import (
    compute_dyn_targets,
    compute_channel_summary,
    update_history,
    get_tracked,
)
from modules.writers.dataset_sheet import write_dataset
from modules.writers.wow_sheets import (
    write_wow_pivot,
    write_wow_monthly,
    write_tbl_class,
    write_wow_export,
)
from modules.writers.case_type_analysis import write_case_type_analysis
from modules.writers.dyntarget_channel_sheet import write_dyntarget_channel
from modules.writers.dyntarget_master_sheet import write_dyntarget_master
from modules.writers.progress_sheet import write_progress
from modules.writers.drill_down_sheet import add_drill_down_sheet


DEFAULT_HISTORY = Path("data/dynamic_targets_history.json")


# ── Subcomando: generate ──────────────────────────────────────────────────────

def cmd_generate(args) -> None:
    csv_path = Path(args.input)
    if not csv_path.exists():
        print(f"Errore: file non trovato: {csv_path}", file=sys.stderr)
        sys.exit(1)

    week      = args.week or csv_path.stem.split()[-1]
    week_date = _week_date(args.week_date)
    min_vol   = args.min_vol
    history_p = Path(args.history)

    print(f"Caricamento dati: {csv_path} ...")
    df     = load_csv(csv_path)
    df_raw = _load_raw(csv_path)

    print(f"  {len(df)} righe caricate | canali: {df['channel'].value_counts().to_dict()}")

    print("Calcolo classificazioni ...")
    tbl = build_tbl_class(df)

    print("Calcolo Dynamic Targets ...")
    dyn_targets     = compute_dyn_targets(df)
    channel_summary = compute_channel_summary(df)

    # --track accetta sia più argomenti separati ("A" "B") sia una singola
    # stringa con ';' (comoda da batch): in entrambi i casi normalizziamo.
    tracked = None
    if args.track:
        tracked = [
            t.strip()
            for item in args.track
            for t in item.split(";")
            if t.strip()
        ]

    print(f"Aggiornamento storico: {history_p} ...")
    history = update_history(
        history_p, int(week), week_date,
        channel_summary, dyn_targets, min_vol=min_vol,
        tracked=tracked,
    )

    out_dir  = Path("output")
    out_dir.mkdir(exist_ok=True)
    out_path = out_dir / f"case_KPI_report_W{week}.xlsx"

    print(f"Scrittura workbook: {out_path} ...")
    wb = xlsxwriter.Workbook(str(out_path))

    # Ordine dei fogli
    write_dataset(wb, df_raw)
    write_case_type_analysis(wb, df, week)
    write_dyntarget_channel(wb, df, dyn_targets, channel_summary, "Phone")
    write_dyntarget_channel(wb, df, dyn_targets, channel_summary, "Non-live")
    write_dyntarget_master(wb, dyn_targets, channel_summary, int(week), week_date)

    # Il foglio Progress viene scritto solo se l'utente ha selezionato dei case type
    if get_tracked(history):
        write_progress(wb, history, week)

    write_wow_pivot(wb, df)
    write_wow_monthly(wb, df)
    write_tbl_class(wb, df)
    write_wow_export(wb, df)

    wb.close()

    print(f"\n✓ Report salvato: {out_path}")
    _print_summary(df, dyn_targets)

    if args.drill_type:
        print(f"\nAggiunta drill-down per: {args.drill_type} ...")
        sheet, n = add_drill_down_sheet(out_path, df, args.drill_type)
        if n == 0:
            print(f"⚠ Attenzione: nessun caso trovato per '{args.drill_type}'. "
                  f"Foglio '{sheet}' creato vuoto (controlla il nome del case type).")
        else:
            print(f"✓ Foglio aggiunto: {sheet} ({n} casi)")


# ── Subcomando: drill ─────────────────────────────────────────────────────────

def cmd_drill(args) -> None:
    out_path = Path(args.output)
    if not out_path.exists():
        print(f"Errore: file non trovato: {out_path}", file=sys.stderr)
        sys.exit(1)

    case_type = args.type
    if not case_type:
        print("Errore: specificare --type", file=sys.stderr)
        sys.exit(1)

    # Cerca il CSV raw nello stesso output per ricostruire df
    # Se non disponibile, richiede --input
    if not args.input:
        print("Errore: specificare --input per il CSV raw", file=sys.stderr)
        sys.exit(1)

    csv_path = Path(args.input)
    if not csv_path.exists():
        print(f"Errore: file non trovato: {csv_path}", file=sys.stderr)
        sys.exit(1)

    df = load_csv(csv_path)
    print(f"Aggiunta foglio drill-down '{case_type}' a {out_path} ...")
    sheet, n = add_drill_down_sheet(out_path, df, case_type)
    if n == 0:
        print(f"⚠ Attenzione: nessun caso trovato per '{case_type}'. "
              f"Foglio '{sheet}' creato vuoto (controlla il nome del case type).")
    else:
        print(f"✓ Foglio aggiunto: {sheet} ({n} casi)")


# ── Utility ───────────────────────────────────────────────────────────────────

def _load_raw(csv_path: Path):
    """Carica il CSV senza normalizzazione (per DATASET sheet)."""
    import pandas as pd
    return pd.read_csv(csv_path, sep=";", encoding="utf-8-sig", decimal=",", low_memory=False)


def _week_date(provided: str | None) -> str:
    if provided:
        return provided
    # Default: lunedì scorso
    today = date.today()
    monday = today - timedelta(days=today.weekday())
    return monday.isoformat()


def _print_summary(df, dyn_targets: dict) -> None:
    print("\nRiepilogo:")
    for channel in ("Phone", "Non-live"):
        sub = df[df["channel"] == channel]
        if sub.empty:
            continue
        from modules.data_loader import weighted_aht
        total_vol = int(sub["cases"].sum())
        team_aht  = weighted_aht(sub)
        n_dt      = len(dyn_targets.get(channel, {}))
        print(f"  {channel:12s}  volume={total_vol:5d}  team AHT={team_aht:.2f} min  "
              f"DynTarget calcolati={n_dt}")


# ── Argparse ──────────────────────────────────────────────────────────────────

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Genera case_KPI_report.xlsx",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    sub = parser.add_subparsers(dest="command", required=True)

    # generate
    gen = sub.add_parser("generate", help="Genera il report settimanale")
    gen.add_argument("input",         help="Path al CSV raw (es. 'raw_data/AHT blended WEEK 24.csv')")
    gen.add_argument("--week",        default="", help="Numero settimana (es. 24)")
    gen.add_argument("--week-date",   default=None,
                     help="Data inizio settimana ISO (es. 2026-06-16), default: lunedì scorso")
    gen.add_argument("--min-vol",     type=int, default=3,
                     help="Volume minimo agente per DynTarget (default: 3)")
    gen.add_argument("--history",     default=str(DEFAULT_HISTORY),
                     help=f"Path file history JSON (default: {DEFAULT_HISTORY})")
    gen.add_argument("--track",       nargs="*", default=None, metavar="CASE_TYPE",
                     help="Case type da mostrare nel foglio Progress (si accumulano "
                          "tra esecuzioni). Lo storico raccoglie comunque tutti i case type.")
    gen.add_argument("--drill-type",  default=None,
                     help="Aggiunge anche il foglio drill-down per questo case type")

    # drill
    drv = sub.add_parser("drill", help="Aggiunge foglio drill-down a report esistente")
    drv.add_argument("--output", required=True, help="Path al workbook .xlsx esistente")
    drv.add_argument("--input",  required=True, help="Path al CSV raw originale")
    drv.add_argument("--type",   required=True, help="Case type da analizzare nel drill-down")

    args = parser.parse_args()

    try:
        if args.command == "generate":
            cmd_generate(args)
        elif args.command == "drill":
            cmd_drill(args)
    except DataError as e:
        print(f"\nErrore nei dati: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
