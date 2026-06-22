"""
Palette colori e factory di formati xlsxwriter condivisi.
"""

# ── Palette colori ──────────────────────────────────────────────────────────────
C_BLUE_HEADER  = "#4472C4"   # intestazioni principali
C_BLUE_LIGHT   = "#D9E1F2"   # sfondo sezioni Non-live/blu chiaro
C_YELLOW_LIGHT = "#FFF2CC"   # sfondo sezioni Phone/giallo chiaro
C_GREEN_LIGHT  = "#E2EFDA"   # sfondo verde chiaro (Process Driven / ok)
C_RED_LIGHT    = "#FFE0E0"   # sfondo rosso chiaro (problematico)
C_ORANGE       = "#F4B942"   # arancione (Actionable)
C_RED          = "#C00000"   # rosso scuro (Highly Actionable)
C_GRAY_LIGHT   = "#F2F2F2"   # grigio molto chiaro (zebra)
C_WHITE        = "#FFFFFF"

# Colori barra DynTarget chart
C_BAR_ABOVE    = "#C00000"   # AHT sopra target → rosso
C_BAR_BELOW    = "#70AD47"   # AHT sotto target → verde
C_LINE_TARGET  = "#FF0000"   # linea DynTarget → rosso
C_LINE_AHT     = "#4472C4"   # linea AHT corrente → blu

CLASSIFICATION_COLORS = {
    "Highly Actionable": C_RED_LIGHT,
    "Actionable":        C_YELLOW_LIGHT,
    "Process Driven":    C_GREEN_LIGHT,
}


def make_formats(wb) -> dict:
    """
    Crea e restituisce un dizionario di formati xlsxwriter riutilizzabili.
    wb = xlsxwriter.Workbook instance
    """
    def fmt(**kw):
        defaults = {"font_name": "Calibri", "font_size": 10, "valign": "vcenter"}
        defaults.update(kw)
        return wb.add_format(defaults)

    return {
        # Header
        "header":        fmt(bold=True, font_color="white",  bg_color=C_BLUE_HEADER,
                             align="center", border=1),
        "header_left":   fmt(bold=True, font_color="white",  bg_color=C_BLUE_HEADER,
                             align="left", border=1),
        # Title / sezione
        "title":         fmt(bold=True, font_size=12, align="center"),
        "section_phone": fmt(bold=True, bg_color=C_YELLOW_LIGHT, border=1, align="center"),
        "section_nl":    fmt(bold=True, bg_color=C_BLUE_LIGHT,   border=1, align="center"),
        # Dati base
        "data":          fmt(border=1, align="center"),
        "data_left":     fmt(border=1, align="left"),
        "data_num":      fmt(border=1, align="center", num_format="0.00"),
        "data_pct":      fmt(border=1, align="center", num_format="0.0%"),
        "data_int":      fmt(border=1, align="center", num_format="0"),
        # Colorati
        "ha":  fmt(border=1, bg_color=C_RED_LIGHT,    align="center"),   # Highly Actionable
        "act": fmt(border=1, bg_color=C_YELLOW_LIGHT, align="center"),   # Actionable
        "pd":  fmt(border=1, bg_color=C_GREEN_LIGHT,  align="center"),   # Process Driven
        "ha_num":  fmt(border=1, bg_color=C_RED_LIGHT,    align="center", num_format="0.00"),
        "act_num": fmt(border=1, bg_color=C_YELLOW_LIGHT, align="center", num_format="0.00"),
        "pd_num":  fmt(border=1, bg_color=C_GREEN_LIGHT,  align="center", num_format="0.00"),
        "ha_pct":  fmt(border=1, bg_color=C_RED_LIGHT,    align="center", num_format="0.0%"),
        "act_pct": fmt(border=1, bg_color=C_YELLOW_LIGHT, align="center", num_format="0.0%"),
        "pd_pct":  fmt(border=1, bg_color=C_GREEN_LIGHT,  align="center", num_format="0.0%"),
        # Barre DynTarget
        "bar_above": fmt(border=1, bg_color=C_RED_LIGHT,   align="center", num_format="0.00"),
        "bar_below": fmt(border=1, bg_color=C_GREEN_LIGHT, align="center", num_format="0.00"),
        # Grigio alternato
        "gray": fmt(border=1, bg_color=C_GRAY_LIGHT, align="center"),
        "gray_left": fmt(border=1, bg_color=C_GRAY_LIGHT, align="left"),
        "gray_num":  fmt(border=1, bg_color=C_GRAY_LIGHT, align="center", num_format="0.00"),
        # Wrap text per header multiriga
        "header_wrap": fmt(bold=True, font_color="white", bg_color=C_BLUE_HEADER,
                           align="center", text_wrap=True, border=1),
        # Bold senza sfondo
        "bold":      fmt(bold=True),
        "bold_left": fmt(bold=True, align="left"),
        "bold_num":  fmt(bold=True, num_format="0.00"),
    }


def class_fmt(fmts: dict, cls: str, kind: str = "") -> object:
    """Ritorna il formato corretto per una classificazione (Highly Actionable, Actionable, Process Driven)."""
    prefix = {"Highly Actionable": "ha", "Actionable": "act", "Process Driven": "pd"}.get(cls, "data")
    key = prefix + ("_" + kind if kind else "")
    return fmts.get(key, fmts["data"])
