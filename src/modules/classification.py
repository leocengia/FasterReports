"""
Case Type Deepdive: scoring e classificazione (Actionable / Highly Actionable /
Process Driven) per Channel × Case Type, allineato al template wow_CaseType_Deepdive.

Convenzione di calcolo: media SEMPLICE di "Case AHT (mins)" + conteggio di
"case_number" (come la pivot/tabelle del foglio Case Type Analysis).

I valori score/categoria calcolati qui (con i parametri di default di
CHANNEL_CONFIG) servono per la SELEZIONE dei case type; nel foglio tblClass le
stesse colonne sono riscritte come formule Excel live che leggono tblParam.
"""

import pandas as pd
from .data_loader import CHANNEL_CONFIG


def compute_deepdive(df_raw: pd.DataFrame,
                     target_phone: float,
                     target_nonlive: float) -> pd.DataFrame:
    """
    Calcola la tabella di scoring per gli scope Phone, Non-live e All.
    Ritorna un DataFrame (una riga per scope × case_type) con metriche, rank,
    variability_score, score e categoria.
    """
    blended = (float(target_phone) + float(target_nonlive)) / 2.0
    base    = CHANNEL_CONFIG["Phone"]  # soglie identiche tra i canali

    scopes = [
        ("Phone",    df_raw[df_raw["Case Origin (group)"] == "Phone"],
         float(target_phone),   CHANNEL_CONFIG["Phone"]),
        ("Non-live", df_raw[df_raw["Case Origin (group)"] != "Phone"],
         float(target_nonlive), CHANNEL_CONFIG["Non-live"]),
        ("All",      df_raw,
         blended,               {**base, "target_aht": blended}),
    ]

    frames = []
    for channel, sub, target, params in scopes:
        t = _scope_metrics(sub, target)
        if t.empty:
            continue
        t["channel"]     = channel
        t["target_aht"]  = target
        t["chiave"]      = channel + "|" + t["case_type"]
        _add_ranks(t)
        t["variability_score"] = 0.4 * t["std_dev_rank"] + 0.6 * t["p90p10_rank"]
        t["score"] = (
            0.25 * t["impact_rank"]
            + 0.20 * t["opportunity_rank"]
            + 0.15 * t["volume_rank"]
            + 0.20 * t["oot_rank"]
            + 0.20 * t["variability_score"]
        )
        t["categoria"] = t.apply(lambda r: _classify(r, params), axis=1)
        frames.append(t)

    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


def select_case_types(dd: pd.DataFrame, top_n: int = 5) -> list[str]:
    """
    Seleziona i case type per Case Type Analysis: dal blocco 'All', filtra
    Actionable/Highly Actionable, ordina per Score desc, primi top_n; se sono
    meno di top_n, completa coi migliori per Score successivi.
    """
    if dd.empty:
        return []
    allr = dd[dd["channel"] == "All"].copy()
    act  = (allr[allr["categoria"].isin(["Actionable", "Highly Actionable"])]
            .sort_values("score", ascending=False))
    selected = act["case_type"].tolist()
    if len(selected) < top_n:
        rest = (allr[~allr["case_type"].isin(selected)]
                .sort_values("score", ascending=False))
        selected += rest["case_type"].tolist()
    return selected[:top_n]


# ── Interni ───────────────────────────────────────────────────────────────────

def _scope_metrics(sub: pd.DataFrame, target: float) -> pd.DataFrame:
    """Metriche per case type su un sottoinsieme (convenzione semplice + conteggio)."""
    if sub.empty:
        return pd.DataFrame()

    aht = pd.to_numeric(sub["Case AHT (mins)"], errors="coerce")
    work = pd.DataFrame({"case_type": sub["Case Type"].values, "aht": aht.values})
    if "case_number" in sub.columns:
        work["cnt"] = sub["case_number"].notna().astype(int).values
    else:
        work["cnt"] = 1

    rows = []
    for ct, g in work.groupby("case_type", dropna=True):
        vol = int(g["cnt"].sum())
        if vol == 0:
            continue
        a         = g["aht"].dropna()
        avg       = float(a.mean()) if len(a) else 0.0
        total_min = float(a.sum())
        oot       = float((g["aht"] > target).sum()) / vol
        std       = float(a.std(ddof=1)) if len(a) > 1 else 0.0
        p90p10    = float(a.quantile(0.9) - a.quantile(0.1)) if len(a) else 0.0
        gap       = max(0.0, avg - target)
        rows.append({
            "case_type":      ct,
            "volume":         vol,
            "aht":            avg,
            "total_minutes":  total_min,
            "out_of_target":  oot,
            "std_dev":        std,
            "p90p10":         p90p10,
            "gap_vs_target":  gap,
            "opp_minutes":    vol * gap,
            "std_dev_ratio":  std / avg if avg else 0.0,
            "p90p10_ratio":   p90p10 / avg if avg else 0.0,
        })

    t = pd.DataFrame(rows)
    if t.empty:
        return t
    scope_total = t["total_minutes"].sum()
    t["impact_channel"] = t["total_minutes"] / scope_total if scope_total else 0.0
    return t


def _pct_rank(s: pd.Series) -> pd.Series:
    """Rank percentile entro lo scope (0=min, 1=max)."""
    n = max(1, len(s) - 1)
    return (s.rank(method="min") - 1) / n


def _add_ranks(t: pd.DataFrame) -> None:
    t["std_dev_rank"]     = _pct_rank(t["std_dev_ratio"])
    t["p90p10_rank"]      = _pct_rank(t["p90p10_ratio"])
    t["impact_rank"]      = _pct_rank(t["impact_channel"])
    t["volume_rank"]      = _pct_rank(t["volume"])
    t["opportunity_rank"] = _pct_rank(t["opp_minutes"])
    t["oot_rank"]         = _pct_rank(t["out_of_target"])


def _classify(r: pd.Series, p: dict) -> str:
    if r["volume"] < p["min_vol"]:
        return "Process Driven"
    score, impact, gap = r["score"], r["impact_channel"], r["gap_vs_target"]
    oot, var, vol, opp_rank = (r["out_of_target"], r["variability_score"],
                               r["volume"], r["opportunity_rank"])
    if (score >= p["highly"] and impact >= p["impact_high"] and gap > 0
            and (oot >= p["oot_high"] or var >= p["var_high"])):
        return "Highly Actionable"
    if ((score >= p["actionable"] and (gap > 0 or oot >= p["oot_mid"]))
            or (var >= p["var_high"] and vol >= 2 * p["min_vol"])
            or (opp_rank >= p["opp_rank_min"] and gap > 0)):
        return "Actionable"
    return "Process Driven"
