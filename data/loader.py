import pandas as pd
import streamlit as st

from config import ARCHIVO, HOY, HOY_TS, PRIO_ORD, _GH_REPO, _GH_FILE

try:
    import requests as _req
except ImportError:
    _req = None


def _token():
    try:
        return st.secrets.get("GITHUB_TOKEN", "")
    except Exception:
        return ""


@st.cache_data(ttl=60)
def _get_file_cache_key():
    """SHA del archivo en GitHub (TTL 60 s). Fallback a mtime local."""
    tok = _token()
    if tok and _req:
        try:
            hdrs = {"Authorization": f"Bearer {tok}",
                    "Accept": "application/vnd.github+json"}
            r = _req.get(
                f"https://api.github.com/repos/{_GH_REPO}/contents/{_GH_FILE}",
                headers=hdrs, timeout=5)
            if r.status_code == 200:
                return r.json()["sha"]
        except Exception:
            pass
    return str(ARCHIVO.stat().st_mtime if ARCHIVO.exists() else 0)


@st.cache_data(ttl=30)
def cargar(mtime):
    data = {}
    try:
        df = pd.read_excel(ARCHIVO, sheet_name="TAREAS")
        for col in ["FECHA_CREACION", "FECHA_COMPROMISO", "FECHA_CIERRE"]:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors="coerce").astype("datetime64[us]")
        defaults = {
            "DESCRIPCION": "", "TIPO": "Tarea", "CATEGORIA": "Personal",
            "PRIORIDAD": "Media", "ESTADO": "Pendiente",
            "IMPACTO": 3, "URGENCIA": 3, "ESFUERZO_HRS": 1.0, "HORAS_REALES": 0.0,
            "TERCERO": "", "PROYECTO": "GENERAL", "AREA": "Trabajo", "NOTAS": "",
        }
        for col, val in defaults.items():
            if col not in df.columns:
                df[col] = val
            if df[col].dtype == object or col in ("DESCRIPCION", "TIPO", "CATEGORIA",
               "PRIORIDAD", "ESTADO", "TERCERO", "PROYECTO", "AREA", "NOTAS"):
                df[col] = df[col].fillna(val)
            else:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(val)
        df = df[df["TAREA"].astype(str).str.strip() != ""].copy()
        if "ORDEN" not in df.columns:
            df["ORDEN"] = list(range(0, len(df) * 10, 10))
        else:
            df["ORDEN"] = pd.to_numeric(df["ORDEN"], errors="coerce")
            max_ord = int(df["ORDEN"].max() or 0)
            df["ORDEN"] = df["ORDEN"].fillna(max_ord + 10).astype(int)
        df["PRIO_ORD"] = df["PRIORIDAD"].map(PRIO_ORD).fillna(99)

        def _urg(row):
            v = row.get("URGENCIA", 0)
            if v and v > 0:
                return v
            fc = row.get("FECHA_COMPROMISO")
            if pd.isna(fc):
                return 1
            d = (fc.date() - HOY).days
            if d < 0:    return 5
            if d <= 3:   return 5
            if d <= 7:   return 4
            if d <= 14:  return 3
            if d <= 30:  return 2
            return 1

        df["URGENCIA"] = df.apply(_urg, axis=1)
        data["tareas"] = df
    except Exception as e:
        data["tareas"] = pd.DataFrame()
        data["_error_tareas"] = str(e)
    try:
        dt = pd.read_excel(ARCHIVO, sheet_name="TERCEROS")
        for col in ["FECHA_INICIO_SEG", "FECHA_ULTIMO_SEG"]:
            if col in dt.columns:
                dt[col] = pd.to_datetime(dt[col], errors="coerce").astype("datetime64[us]")
        if "FECHA_ULTIMO_SEG" in dt.columns:
            dt["DIAS_SIN_RESP"] = dt["FECHA_ULTIMO_SEG"].apply(
                lambda x: (HOY - x.date()).days if pd.notna(x) else 999
            )
        data["terceros"] = dt
    except Exception:
        data["terceros"] = pd.DataFrame()
    return data
