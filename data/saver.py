import base64
import io
import time

import pandas as pd
import requests
import streamlit as st

from config import ARCHIVO, _GH_REPO, _GH_FILE, _GH_BRANCH
from data.loader import cargar, _get_file_cache_key


def _token():
    try:
        return st.secrets.get("GITHUB_TOKEN", "")
    except Exception:
        return ""


def guardar_github(df_tareas_nuevo, df_ter_nuevo=None, df_ter_base=None):
    """Sobreescribe TAREAS (y opcionalmente TERCEROS) en el repo vía GitHub API y limpia caché.

    df_ter_base: DataFrame de terceros actual (fallback si df_ter_nuevo es None).
    """
    token = _token()
    if not token:
        return False, (
            "GITHUB_TOKEN no configurado. "
            "Ve a Streamlit Cloud → Manage app → Secrets y agrega:\n"
            'GITHUB_TOKEN = "ghp_tu_token_aqui"'
        )
    hdrs = {"Authorization": f"Bearer {token}",
            "Accept": "application/vnd.github+json"}
    api  = f"https://api.github.com/repos/{_GH_REPO}/contents/{_GH_FILE}"

    r = requests.get(api, headers=hdrs, timeout=10)
    if r.status_code != 200:
        return False, f"Error leyendo repo ({r.status_code}): {r.json().get('message','')}"
    sha = r.json()["sha"]

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df_save = df_tareas_nuevo.copy()
        for c in ["FECHA_CREACION", "FECHA_COMPROMISO", "FECHA_CIERRE"]:
            if c in df_save.columns:
                df_save[c] = pd.to_datetime(df_save[c], errors="coerce")
        df_save = df_save.drop(columns=["PRIO_ORD"], errors="ignore")
        df_save.to_excel(writer, sheet_name="TAREAS", index=False)
        _ter = df_ter_nuevo if df_ter_nuevo is not None else (
            df_ter_base if df_ter_base is not None else pd.DataFrame())
        ter_save = _ter.drop(columns=["DIAS_SIN_RESP"], errors="ignore")
        if not ter_save.empty:
            ter_save.to_excel(writer, sheet_name="TERCEROS", index=False)
    buf.seek(0)
    raw_bytes = buf.read()

    ts      = pd.Timestamp.now().strftime("%d/%m %H:%M")
    payload = {
        "message": f"update: tareas actualizadas desde app ({ts})",
        "content": base64.b64encode(raw_bytes).decode(),
        "sha":     sha,
        "branch":  _GH_BRANCH,
    }
    _delays   = [0, 2, 5]
    _last_msg = ""
    for _delay in _delays:
        if _delay:
            time.sleep(_delay)
        try:
            r2 = requests.put(api, json=payload, headers=hdrs, timeout=15)
            if r2.status_code in (200, 201):
                try:
                    ARCHIVO.write_bytes(raw_bytes)
                except Exception:
                    pass
                cargar.clear()
                _get_file_cache_key.clear()
                st.session_state.pop("_last_error", None)
                return True, f"Guardado a las {ts}"
            _last_msg = f"Error al guardar ({r2.status_code}): {r2.json().get('message','')}"
        except requests.exceptions.Timeout:
            _last_msg = "Timeout al conectar con GitHub API"
        except Exception as _e:
            _last_msg = f"Error inesperado: {_e}"
    st.session_state["_last_error"] = _last_msg
    return False, _last_msg
