import pandas as pd
import streamlit as st

from config import HOY, HOY_TS


def _activas(df_raw):
    return df_raw[~df_raw["ESTADO"].isin(["Completada", "Cancelada"])]


def _marcar_estado(df_raw, task_id, nuevo_estado, guardar_fn):
    """Cambia el estado de una tarea y guarda en GitHub."""
    df_copy = df_raw.copy()
    mask = df_copy["ID"] == task_id
    df_copy.loc[mask, "ESTADO"] = nuevo_estado
    if nuevo_estado == "Completada":
        df_copy.loc[mask, "FECHA_CIERRE"] = pd.Timestamp(HOY)
    with st.spinner("Guardando..."):
        ok, msg = guardar_fn(df_copy)
    if ok:
        st.toast("Tarea actualizada", icon="✅")
        st.rerun()
    else:
        st.error(msg)


def _merge_edits(df_original, df_edited):
    """Fusiona filas editadas al DataFrame completo usando update() por índice ID."""
    df_result = df_original.copy()
    df_ed     = df_edited.copy()
    for c in ["FECHA_COMPROMISO", "FECHA_CIERRE", "FECHA_CREACION"]:
        if c in df_ed.columns:
            df_ed[c] = pd.to_datetime(df_ed[c], errors="coerce").astype("datetime64[us]")
    for c in ["ID", "IMPACTO", "URGENCIA"]:
        if c in df_ed.columns:
            df_ed[c] = pd.to_numeric(df_ed[c], errors="coerce")
    if "ESFUERZO_HRS" in df_ed.columns:
        df_ed["ESFUERZO_HRS"] = pd.to_numeric(df_ed["ESFUERZO_HRS"], errors="coerce")
    existing = df_ed[df_ed["ID"].notna()].copy()
    new_rows = df_ed[df_ed["ID"].isna()].copy()
    if not existing.empty:
        df_result = df_result.set_index("ID")
        existing  = existing.set_index("ID")
        cols_upd  = [c for c in existing.columns if c in df_result.columns]
        df_result.update(existing[cols_upd])
        df_result = df_result.reset_index()
    if not new_rows.empty:
        max_id = int(df_result["ID"].max() or 0)
        new_rows = new_rows.copy()
        new_rows["ID"]             = list(range(max_id + 1, max_id + 1 + len(new_rows)))
        new_rows["FECHA_CREACION"] = pd.Timestamp(HOY)
        for col, default in [("ESTADO", "Pendiente"), ("PRIORIDAD", "Media"),
                              ("TIPO", "Tarea"), ("AREA", "Trabajo")]:
            if col in new_rows.columns:
                new_rows[col] = new_rows[col].fillna(default)
            else:
                new_rows[col] = default
        for col in df_result.columns:
            if col not in new_rows.columns:
                new_rows[col] = pd.NA
        df_result = pd.concat(
            [df_result, new_rows[df_result.columns]], ignore_index=True)
    return df_result
