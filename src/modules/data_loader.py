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


def load_csv(path: str | Path) -> pd.DataFrame:
    """Carica il CSV raw e normalizza le colonne chiave."""
    df = pd.read_csv(
        Path(path),
        sep=";",
        encoding="utf-8-sig",
        decimal=",",
    )
    df["aht"]    = pd.to_numeric(df["Case AHT (mins)"], errors="coerce").fillna(0)
    df["cases"]  = pd.to_numeric(df["Distinct Cases"],  errors="coerce").fillna(0)
    df = df[df["cases"] > 0].copy()

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
