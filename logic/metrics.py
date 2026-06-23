import pandas as pd
from datetime import timedelta

from config import HOY, HOY_TS, C_CIAN, C_CRITICO, C_ALERTA, C_OK, C_GRIS
from logic.tasks import _activas


def dias_restantes(ts):
    if pd.isna(ts):
        return None
    return (ts.date() - HOY).days


def semaforo(ts):
    d = dias_restantes(ts)
    if d is None:  return C_GRIS,    "Sin fecha"
    if d < 0:      return C_CRITICO, f"Vencida {abs(d)}d"
    if d == 0:     return C_CRITICO, "Vence HOY"
    if d <= 3:     return C_ALERTA,  f"Vence en {d}d"
    if d <= 7:     return C_CIAN,    f"Vence en {d}d"
    return "#475569",                f"Vence en {d}d"


def estado_general(df_raw):
    ac        = _activas(df_raw)
    venc      = ac[ac["FECHA_COMPROMISO"].notna() & (ac["FECHA_COMPROMISO"] < HOY_TS)]
    crit      = ac[ac["PRIORIDAD"] == "Crítica"]
    crit_venc = venc[venc["PRIORIDAD"] == "Crítica"]
    if len(crit_venc) > 0 or len(venc) > 5:
        return "CRÍTICO",      C_CRITICO, "⛔"
    elif len(venc) > 0 or len(crit) >= 2:
        return "BAJO PRESIÓN", C_ALERTA,  "⚠️"
    return "CONTROLADO",       C_OK,      "✅"


def tasa_cumplimiento_semana(df_raw):
    ac       = _activas(df_raw)
    ini      = HOY - timedelta(days=HOY.weekday())
    comp_sem = df_raw[
        (df_raw["ESTADO"] == "Completada") &
        df_raw["FECHA_CIERRE"].notna() &
        (df_raw["FECHA_CIERRE"] >= pd.Timestamp(ini))
    ]
    venc_sem = ac[
        ac["FECHA_COMPROMISO"].notna() &
        (ac["FECHA_COMPROMISO"] >= HOY_TS - pd.Timedelta(days=7)) &
        (ac["FECHA_COMPROMISO"] < HOY_TS)
    ]
    total = len(comp_sem) + len(venc_sem)
    return round(len(comp_sem) / max(total, 1) * 100), len(comp_sem)
