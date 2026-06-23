"""
Lettura CSV e utility condivise per il calcolo AHT.
"""

from pathlib import Path
import pandas as pd

# Mappatura Case Origin (group) → label canale
CHANNEL_MAP = {
    "Phone": "Phone",
    "Other": "Non-live",
}

# Configurazione soglie per canale
CHANNEL_CONFIG = {
    "Phone": {
        "target_aht":   19.98,
        "min_vol":       20,
        "highly":        0.75,
        "actionable":    0.50,
        "impact_high":   0.02,
        "oot_mid":       0.20,
    },
    "Non-live": {
        "target_aht":   18.96,
        "min_vol":       20,
        "highly":        0.75,
        "actionable":    0.50,
        "impact_high":   0.02,
        "oot_mid":       0.20,
    },
}

DYN_TARGET_MIN_VOL = 3   # volume minimo per agente nel calcolo DynTarget
DYN_TARGET_TOP_N   = 15  # top N agenti per AHT più basso

# Colonne indispensabili per generare il report
REQUIRED_COLUMNS = [
    "Case AHT (mins)",
    "Distinct Cases",
    "Case Origin (group)",
    "Case Type",
    "Employee Name",
]


class DataError(Exception):
    """Errore di input dati con messaggio leggibile per l'utente."""


def load_csv(path: str | Path) -> pd.DataFrame:
    """Carica il CSV raw, valida le colonne richieste e normalizza i campi chiave."""
    df = pd.read_csv(
        Path(path),
        sep=";",
        encoding="utf-8-sig",
        decimal=",",
        low_memory=False,
    )

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise DataError(
            "Colonne mancanti nel CSV: "
            + ", ".join(f"'{c}'" for c in missing)
            + ".\n  Verifica che l'export non abbia rinominato le intestazioni "
              "o cambiato separatore (atteso ';')."
        )

    df["aht"]    = pd.to_numeric(df["Case AHT (mins)"], errors="coerce").fillna(0)
    df["cases"]  = pd.to_numeric(df["Distinct Cases"],  errors="coerce").fillna(0)
    df = df[df["cases"] > 0].copy()

    if df.empty:
        raise DataError(
            "Nessuna riga valida nel CSV (dopo il filtro 'Distinct Cases' > 0).\n"
            "  Il file potrebbe essere vuoto o contenere solo l'intestazione."
        )

    # Normalizza canale: "Phone" → "Phone", "Other" → "Non-live"
    df["channel"] = df["Case Origin (group)"].map(CHANNEL_MAP).fillna("Non-live")
    return df.reset_index(drop=True)


def weighted_aht(group: pd.DataFrame,
                 aht_col: str = "aht",
                 vol_col: str = "cases") -> float:
    """Media pesata dell'AHT per un gruppo di righe."""
    total = group[vol_col].sum()
    if total == 0:
        return 0.0
    return float((group[aht_col] * group[vol_col]).sum() / total)


def agg_by_case_type(df: pd.DataFrame, channel: str | None = None) -> pd.DataFrame:
    """
    Aggrega per Case Type dentro un canale (o tutti i canali se None).
    Ritorna: case_type, volume, avg_aht, total_minutes, channel
    """
    if channel is not None:
        sub = df[df["channel"] == channel].copy()
    else:
        sub = df.copy()

    if sub.empty:
        return pd.DataFrame(columns=["case_type", "volume", "avg_aht", "total_minutes", "channel"])

    result = (
        sub.groupby("Case Type")
        .apply(lambda g: pd.Series({
            "volume":        int(g["cases"].sum()),
            "avg_aht":       weighted_aht(g),
            "total_minutes": float((g["aht"] * g["cases"]).sum()),
        }), include_groups=False)
        .reset_index()
        .rename(columns={"Case Type": "case_type"})
    )
    result["channel"] = channel or "All"
    return result.reset_index(drop=True)


def agg_simple(df: pd.DataFrame, by_col: str, count_col: str = "case_number") -> pd.DataFrame:
    """
    Aggrega per `by_col` con la stessa logica di una pivot Excel:
      - avg_aht = media SEMPLICE di "Case AHT (mins)" (non pesata)
      - volume  = CONTEGGIO di `count_col` (case_number) — righe non vuote del gruppo

    Va calcolata sulla stessa popolazione della pivot (tipicamente df_raw), così
    le tabelle G-J del foglio Case Type Analysis coincidono con la pivot.

    Ritorna un DataFrame [by_col, avg_aht, volume] ordinato per volume desc.
    Se `count_col` è assente, volume = numero di righe del gruppo.
    """
    aht = pd.to_numeric(df["Case AHT (mins)"], errors="coerce")
    tmp = pd.DataFrame({by_col: df[by_col], "_aht": aht})

    if count_col in df.columns:
        tmp["_cnt"] = df[count_col].values
        agg_spec = {"avg_aht": ("_aht", "mean"), "volume": ("_cnt", "count")}
    else:
        agg_spec = {"avg_aht": ("_aht", "mean"), "volume": ("_aht", "size")}

    out = (
        tmp.groupby(by_col, dropna=True)
        .agg(**agg_spec)
        .reset_index()
    )
    out["avg_aht"] = out["avg_aht"].fillna(0.0)
    out["volume"]  = out["volume"].astype(int)
    return out.sort_values("volume", ascending=False).reset_index(drop=True)


def agg_by_agent_casetype(df: pd.DataFrame, channel: str | None = None) -> pd.DataFrame:
    """
    Aggrega per (Employee Name, Case Type) per il calcolo dei DynTarget.
    Ritorna: agent, case_type, volume, avg_aht
    """
    if channel is not None:
        sub = df[df["channel"] == channel].copy()
    else:
        sub = df.copy()

    if sub.empty:
        return pd.DataFrame(columns=["agent", "case_type", "volume", "avg_aht"])

    result = (
        sub.groupby(["Employee Name", "Case Type"])
        .apply(lambda g: pd.Series({
            "volume":  int(g["cases"].sum()),
            "avg_aht": weighted_aht(g),
        }), include_groups=False)
        .reset_index()
        .rename(columns={"Employee Name": "agent", "Case Type": "case_type"})
    )
    return result.reset_index(drop=True)
