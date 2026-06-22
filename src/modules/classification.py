"""
Logica di classificazione tblClass (Actionable / Highly Actionable / Process Driven).
"""

import pandas as pd
from .data_loader import CHANNEL_CONFIG


def build_tbl_class(df: pd.DataFrame) -> pd.DataFrame:
    """
    Costruisce la tabella master di scoring (equivalente al foglio tblClass del template).

    Per ogni combinazione Channel × Case Type calcola:
      - Volume, AHT, Total_Minutes, Impact_Channel, Out_of_Target
      - StdDev, P90P10, StdDev_Ratio, P90P10_Ratio
      - Gap_vs_Target, Opportunity_Minutes
      - Ranking percentile: StdDev_Rank, P90P10_Rank, Impact_Rank, Volume_Rank
      - Variability_Score, Classification
    """
    rows = []
    for channel, cfg in CHANNEL_CONFIG.items():
        sub = df[df["channel"] == channel].copy()
        if sub.empty:
            continue

        total_minutes_channel = float((sub["aht"] * sub["cases"]).sum())
        target = cfg["target_aht"]

        for ct, grp in sub.groupby("Case Type"):
            volume        = int(grp["cases"].sum())
            total_min     = float((grp["aht"] * grp["cases"]).sum())
            avg_aht       = total_min / volume if volume else 0.0
            impact        = total_min / total_minutes_channel if total_minutes_channel else 0.0
            oot           = float((grp["aht"] > target).sum()) / volume if volume else 0.0
            std_dev       = float(grp["aht"].std()) if len(grp) > 1 else 0.0
            p90           = float(grp["aht"].quantile(0.9))
            p10           = float(grp["aht"].quantile(0.1))
            p90p10        = p90 - p10
            gap_vs_target = max(0.0, avg_aht - target)
            opp_min       = volume * gap_vs_target
            std_ratio     = std_dev / avg_aht if avg_aht else 0.0
            p90p10_ratio  = p90p10 / avg_aht if avg_aht else 0.0

            rows.append({
                "channel":            channel,
                "case_type":          ct,
                "chiave":             f"{channel}|{ct}",
                "volume":             volume,
                "avg_aht":            avg_aht,
                "total_minutes":      total_min,
                "impact_channel":     impact,
                "out_of_target":      oot,
                "std_dev":            std_dev,
                "p90p10":             p90p10,
                "target_aht":         target,
                "gap_vs_target":      gap_vs_target,
                "opp_minutes":        opp_min,
                "std_dev_ratio":      std_ratio,
                "p90p10_ratio":       p90p10_ratio,
            })

    if not rows:
        return pd.DataFrame()

    tbl = pd.DataFrame(rows)

    # Ranking percentile per canale (0=min, 1=max)
    for channel in tbl["channel"].unique():
        mask = tbl["channel"] == channel
        sub  = tbl[mask].copy()
        n    = max(1, len(sub) - 1)

        for col, rank_col in [
            ("std_dev_ratio",  "std_dev_rank"),
            ("p90p10_ratio",   "p90p10_rank"),
            ("impact_channel", "impact_rank"),
            ("volume",         "volume_rank"),
        ]:
            ranks = sub[col].rank(method="min") - 1
            tbl.loc[mask, rank_col] = (ranks / n).values

    tbl["variability_score"] = (tbl["std_dev_rank"] + tbl["p90p10_rank"]) / 2

    # Classificazione
    tbl["classification"] = tbl.apply(_classify, axis=1)

    return tbl.reset_index(drop=True)


def _classify(row: pd.Series) -> str:
    cfg = CHANNEL_CONFIG.get(row["channel"], {})
    min_vol    = cfg.get("min_vol",    20)
    highly     = cfg.get("highly",     0.75)
    actionable = cfg.get("actionable", 0.50)
    impact_high = cfg.get("impact_high", 0.02)
    oot_mid    = cfg.get("oot_mid",    0.20)

    if row["volume"] < min_vol:
        return "Process Driven"
    oot    = row["out_of_target"]
    impact = row["impact_channel"]
    if oot >= highly:
        return "Highly Actionable"
    if oot >= actionable:
        return "Actionable"
    if oot >= oot_mid and impact >= impact_high:
        return "Actionable"
    return "Process Driven"


def get_relevant_case_types(tbl: pd.DataFrame, channel: str) -> list[str]:
    """
    Ritorna i case type classificati Actionable o Highly Actionable per il canale dato,
    ordinati per impact_channel desc.
    """
    mask = (
        (tbl["channel"] == channel) &
        (tbl["classification"].isin(["Actionable", "Highly Actionable"]))
    )
    return (
        tbl[mask]
        .sort_values("impact_channel", ascending=False)["case_type"]
        .tolist()
    )
