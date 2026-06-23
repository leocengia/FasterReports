"""
Calcolo Dynamic Targets e gestione dello storico settimanale.
"""

import json
from datetime import date
from pathlib import Path

import pandas as pd

from .data_loader import (
    CHANNEL_CONFIG,
    DYN_TARGET_MIN_VOL,
    DYN_TARGET_TOP_N,
    agg_by_agent_casetype,
    weighted_aht,
)


def compute_dyn_targets(df: pd.DataFrame) -> dict:
    """
    Calcola il DynTarget per ogni Channel × Case Type.

    DynTarget = media pesata AHT dei top-N agenti (AHT più basso)
    con volume >= DYN_TARGET_MIN_VOL.

    Ritorna dict: {channel: {case_type: {"dyn_target": float, "top_agents": df}}}
    """
    result = {}
    for channel in CHANNEL_CONFIG:
        agent_ct = agg_by_agent_casetype(df, channel=channel)
        if agent_ct.empty:
            continue
        result[channel] = {}
        for ct, grp in agent_ct.groupby("case_type"):
            eligible = grp[grp["volume"] >= DYN_TARGET_MIN_VOL].copy()
            if eligible.empty:
                continue
            top = eligible.nsmallest(DYN_TARGET_TOP_N, "avg_aht").reset_index(drop=True)
            dyn = weighted_aht(top, aht_col="avg_aht", vol_col="volume")
            result[channel][ct] = {
                "dyn_target": dyn,
                "top_agents": top,
            }
    return result


def compute_channel_summary(df: pd.DataFrame) -> dict:
    """
    Per ogni Channel × Case Type: volume totale e avg_aht corrente.
    Ritorna dict: {channel: {case_type: {"volume": int, "avg_aht": float}}}
    """
    summary = {}
    for channel in CHANNEL_CONFIG:
        sub = df[df["channel"] == channel]
        if sub.empty:
            continue
        summary[channel] = {}
        for ct, grp in sub.groupby("Case Type"):
            vol = int(grp["cases"].sum())
            aht = weighted_aht(grp)
            summary[channel][ct] = {"volume": vol, "avg_aht": aht}
    return summary


# Chiave riservata nel dict history per la configurazione (non è un canale)
CONFIG_KEY = "_config"
CHANNELS   = ("Phone", "Non-live")


def update_history(history_path: Path, week: int, week_date: str,
                   channel_summary: dict, dyn_targets: dict,
                   min_vol: int = DYN_TARGET_MIN_VOL,
                   tracked: list[str] | None = None) -> dict:
    """
    Carica il file history (o crea vuoto), aggiunge la settimana corrente
    per ogni Channel × Case Type con DynTarget disponibile.

    Lo storico accumula TUTTI i case type (così, quando se ne seleziona uno
    per il tracking, la sua storia pregressa è già disponibile). La selezione
    dei case type da mostrare nel foglio Progress è salvata in history[CONFIG_KEY].

    Salva e ritorna il dict aggiornato.
    """
    history_path.parent.mkdir(parents=True, exist_ok=True)
    if history_path.exists():
        with open(history_path, encoding="utf-8") as f:
            history = json.load(f)
    else:
        history = {}

    # Aggiorna l'elenco dei case type tracciati (se fornito)
    if tracked is not None:
        cfg = history.setdefault(CONFIG_KEY, {})
        existing = set(cfg.get("tracked", []))
        cfg["tracked"] = sorted(existing | set(tracked))

    for channel, ct_map in dyn_targets.items():
        if channel not in history:
            history[channel] = {}
        for ct, data in ct_map.items():
            curr_aht = channel_summary.get(channel, {}).get(ct, {}).get("avg_aht", 0.0)
            curr_vol = channel_summary.get(channel, {}).get(ct, {}).get("volume", 0)

            if ct not in history[channel]:
                history[channel][ct] = {
                    "min_vol_threshold": min_vol,
                    "weeks": [],
                }

            existing_weeks = {w["week"] for w in history[channel][ct]["weeks"]}
            if week not in existing_weeks:
                history[channel][ct]["weeks"].append({
                    "week":        week,
                    "date":        week_date,
                    "aht_min":     round(curr_aht, 6),
                    "aht_sec":     round(curr_aht * 60, 4),
                    "volume":      curr_vol,
                    "dyn_target_min": round(data["dyn_target"], 6),
                    "dyn_target_sec": round(data["dyn_target"] * 60, 4),
                })

    with open(history_path, "w", encoding="utf-8") as f:
        json.dump(history, f, indent=2, ensure_ascii=False)

    return history


def get_tracked(history: dict) -> list[str]:
    """Ritorna l'elenco dei case type selezionati per il Progress Tracking."""
    return history.get(CONFIG_KEY, {}).get("tracked", [])


def get_progress_data(history: dict, channel: str, case_type: str) -> list[dict]:
    """
    Ritorna la lista di settimane (ordinate) per un Channel × Case Type,
    con campi aggiuntivi: delta_sec (riduzione vs settimana precedente), on_track.
    """
    weeks = history.get(channel, {}).get(case_type, {}).get("weeks", [])
    if not weeks:
        return []
    weeks = sorted(weeks, key=lambda w: w["week"])
    result = []
    for i, w in enumerate(weeks):
        delta = None
        if i > 0:
            delta = round((weeks[i - 1]["aht_sec"] - w["aht_sec"]), 2)
        result.append({**w, "delta_sec": delta})
    return result
