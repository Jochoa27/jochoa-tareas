# ═══════════════════════════════════════════════════════════════════════════════
# CENTRO DE COMANDO PERSONAL — JUAN OCHOA
# Constructora Londres · Planificación y Control · v2.0 · 2026
# ═══════════════════════════════════════════════════════════════════════════════
import streamlit as st
import pandas as pd
import json
import base64
import io
import re
import requests
from pathlib import Path
from datetime import date, timedelta
import streamlit.components.v1 as components

# ─── CONFIG ───────────────────────────────────────────────────────────────────
ARCHIVO = Path(__file__).parent / "TAREAS_JOCHOA.xlsx"
HOY     = date.today()
HOY_TS  = pd.Timestamp(HOY)

# ─── PALETA ───────────────────────────────────────────────────────────────────
C_BG      = "#060B15"
C_BG2     = "#08101F"
C_BG3     = "#0D1830"
C_CIAN    = "#38BDF8"
C_CRITICO = "#FF4757"
C_ALERTA  = "#FFB300"
C_OK      = "#23D160"
C_GRIS    = "#64748B"
C_MORADO  = "#A855F7"
C_VERDE2  = "#10B981"
C_INDIGO  = "#818CF8"

PRIO_ORD = {"Crítica": 0, "Alta": 1, "Media": 2, "Baja": 3}
PRIO_CLR = {"Crítica": C_CRITICO, "Alta": "#FF6B35", "Media": C_CIAN, "Baja": C_GRIS}
PRIO_ICO = {"Crítica": "🔴", "Alta": "🟠", "Media": "🔵", "Baja": "⚪"}

EST_CLR = {
    "Pendiente":          C_GRIS,
    "En Proceso":         C_CIAN,
    "Esperando Terceros": C_ALERTA,
    "Completada":         C_OK,
    "Cancelada":          "#334155",
}
EST_ICO = {
    "Pendiente": "⏳", "En Proceso": "🔄",
    "Esperando Terceros": "⌛", "Completada": "✅", "Cancelada": "❌",
}
TIPO_CLR = {
    "Tarea": C_CIAN, "Seguimiento": C_ALERTA,
    "Compromiso": C_MORADO, "Reunión": C_VERDE2,
}
CAT_CLR = {
    "Planificación": C_INDIGO, "Contratos": "#F59E0B",
    "Compras": C_VERDE2, "Reportes": "#06B6D4",
    "IA y Automatización": C_MORADO, "Gestión Corporativa": "#F97316",
    "Reuniones": "#84CC16", "Salud": "#EC4899", "Personal": C_GRIS,
}

# ─── PAGE CONFIG ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Centro de Comando · Jochoa",
    page_icon="🎯",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── CSS ──────────────────────────────────────────────────────────────────────
st.markdown("""<style>
[data-testid="stAppViewContainer"]{background:#060B15;}
[data-testid="stSidebar"]{background:#05090F;border-right:1px solid rgba(56,189,248,0.08);}
.block-container{padding:1.2rem 1.6rem 3rem;}
[data-testid="stSidebar"] .stButton>button{
    background:rgba(56,189,248,0.04);border:1px solid rgba(56,189,248,0.08);
    color:#64748B;font-weight:700;font-size:0.80rem;letter-spacing:0.04em;
    border-radius:10px;padding:10px 14px;text-align:left;width:100%;
    margin-bottom:3px;transition:all .18s;}
[data-testid="stSidebar"] .stButton>button:hover{
    background:rgba(56,189,248,0.10);border-color:rgba(56,189,248,0.22);
    color:#94A3B8;}
[data-testid="stSidebar"] .stButton>button.active,
[data-testid="stSidebar"] .stButton[data-active="true"]>button{
    background:rgba(56,189,248,0.14);border-color:rgba(56,189,248,0.38);
    color:#38BDF8;box-shadow:0 0 12px rgba(56,189,248,0.12);}
/* Section headers */
.sh{display:flex;align-items:center;gap:12px;margin:1.4rem 0 0.9rem;
    padding-bottom:10px;border-bottom:1px solid rgba(56,189,248,0.10);}
.sh-bar{width:4px;height:24px;border-radius:3px;flex-shrink:0;}
.sh-txt{font-size:0.88rem;font-weight:800;letter-spacing:0.12em;
        text-transform:uppercase;color:#E2E8F0;text-shadow:0 0 18px rgba(56,189,248,0.22);}
/* KPI cards */
.kpi-card{border-radius:14px;padding:14px 16px 12px;min-height:88px;display:flex;
           flex-direction:column;justify-content:center;transition:border-color .2s;}
.kpi-lbl{font-size:0.50rem;font-weight:800;letter-spacing:0.14em;
          text-transform:uppercase;color:#475569;margin-bottom:4px;}
.kpi-val{font-size:1.50rem;font-weight:900;line-height:1;font-variant-numeric:tabular-nums;}
.kpi-sub{font-size:0.60rem;font-weight:600;margin-top:4px;opacity:0.80;}
/* Task cards */
.tc{border-radius:12px;padding:11px 14px;margin-bottom:7px;border:1px solid transparent;}
.tc-t{font-size:0.81rem;font-weight:700;color:#F1F5F9;line-height:1.36;margin:5px 0 6px;}
.tc-tag{font-size:0.50rem;font-weight:700;letter-spacing:0.05em;
         padding:2px 8px;border-radius:8px;text-transform:uppercase;display:inline-block;margin-right:4px;}
.tc-dt{font-size:0.60rem;font-weight:700;margin-top:6px;}
/* Status badge */
.estado-badge{display:inline-flex;align-items:center;gap:8px;padding:6px 16px;
               border-radius:20px;font-size:0.72rem;font-weight:800;letter-spacing:0.12em;}
/* Divider */
hr{border:none;border-top:1px solid rgba(56,189,248,0.08);margin:12px 0;}
/* Scrollable containers */
.scroll-box{overflow-y:auto;max-height:380px;padding-right:4px;}
</style>""", unsafe_allow_html=True)

# ─── TEMAS DE COLOR (3 temas mundo-clase) ─────────────────────────────────────
_THEMES = {
    "void": {
        "name":"Void",  "icon":"⬛", "desc":"Ultra dark · premium",
        "dark":True,
        "bg_css":"linear-gradient(150deg,#050507 0%,#07090F 65%,#040408 100%)",
        "bg_sb":"#030305",
        "bg1":"#050507","bg2":"#07090F","bg3":"#0D0F18",
        "card1":"rgba(8,8,14,0.97)","card2":"rgba(8,8,14,0.82)",
        "border":"rgba(255,255,255,0.055)",
        "nm_clr":"#E2E8F0","meta_clr":"#52525B",
        "txt_dim":"#334155","abr":"56,189,248",
    },
    "zinc": {
        "name":"Zinc",  "icon":"◼", "desc":"Dark gray · profesional",
        "dark":True,
        "bg_css":"linear-gradient(150deg,#18181B 0%,#1C1C1F 65%,#18181B 100%)",
        "bg_sb":"#101012",
        "bg1":"#18181B","bg2":"#1C1C1F","bg3":"#232327",
        "card1":"rgba(28,28,31,0.97)","card2":"rgba(22,22,25,0.82)",
        "border":"rgba(255,255,255,0.065)",
        "nm_clr":"#E4E4E7","meta_clr":"#52525B",
        "txt_dim":"#3F3F46","abr":"56,189,248",
    },
    "pearl": {
        "name":"Pearl", "icon":"☀️", "desc":"Light · sidebar oscuro",
        "dark":False,
        "bg_css":"linear-gradient(150deg,#FAFAFA 0%,#F4F4F5 55%,#ECECEE 100%)",
        "bg_sb":"#18181B",
        "bg1":"#F4F4F5","bg2":"#E4E4E7","bg3":"#D4D4D8",
        "card1":"rgba(255,255,255,0.97)","card2":"rgba(244,244,245,0.95)",
        "border":"rgba(0,0,0,0.075)",
        "nm_clr":"#18181B","meta_clr":"#71717A",
        "txt_dim":"#A1A1AA","abr":"14,165,233",
    },
}

_TK   = st.session_state.get("theme", "void")
_TC   = _THEMES.get(_TK, _THEMES["void"])
_DARK = _TC["dark"]
C_BG  = _TC["bg1"]
C_BG2 = _TC["bg2"]
C_BG3 = _TC["bg3"]
_ABR  = _TC["abr"]

_PEARL_CSS = "" if _DARK else f"""
/* ── PEARL light mode overrides ── */
.kpi-val{{color:#09090B!important;}}
.kpi-lbl{{color:#71717A!important;}}
.kpi-sub{{color:#52525B!important;}}
.sh-txt{{color:#09090B!important;text-shadow:none!important;}}
.sh{{border-bottom-color:rgba({_ABR},0.22)!important;}}
hr{{border-top-color:rgba(0,0,0,0.07)!important;}}
[data-testid="stMetricValue"]{{color:#09090B!important;}}
[data-testid="stMetricLabel"]{{color:#71717A!important;}}
[data-testid="metric-container"]{{
    background:rgba(255,255,255,0.85)!important;
    border-color:rgba(0,0,0,0.07)!important;
    box-shadow:0 1px 6px rgba(0,0,0,0.06)!important;
}}
[data-testid="stDataFrame"]{{border-color:rgba(0,0,0,0.08)!important;}}
[data-testid="stSelectbox"] label,[data-testid="stTextInput"] label,
[data-testid="stTextArea"] label,[data-testid="stDateInput"] label{{
    color:#3F3F46!important;
}}
"""

st.markdown(f"""<style>
[data-testid="stAppViewContainer"]{{
    background:{_TC['bg_css']} !important;
    min-height:100vh;
}}
[data-testid="stSidebar"]{{
    background:{_TC['bg_sb']} !important;
    border-right:2px solid rgba({_ABR},0.20) !important;
    box-shadow:4px 0 28px rgba({_ABR},0.07) !important;
}}
/* ── Responsive móvil / tablet ── */
@media(max-width:900px){{
    .block-container{{padding:0.8rem 0.7rem 2rem!important;}}
    .kpi-val{{font-size:1.15rem!important;}}
    .kpi-card{{min-height:70px!important;padding:10px 12px 8px!important;}}
    .kpi-lbl{{font-size:0.42rem!important;}}
    .sh-txt{{font-size:0.74rem!important;}}
    .sh{{margin:1rem 0 0.7rem!important;}}
}}
@media(max-width:640px){{
    .block-container{{padding:0.5rem 0.4rem 1.5rem!important;}}
    .kpi-val{{font-size:0.95rem!important;}}
    .kpi-lbl{{font-size:0.38rem!important;}}
    .sh-txt{{font-size:0.65rem!important;}}
}}
{_PEARL_CSS}
</style>""", unsafe_allow_html=True)

# ─── DATA LOADING ─────────────────────────────────────────────────────────────
@st.cache_data(ttl=30)
def cargar(mtime):
    data = {}
    try:
        df = pd.read_excel(ARCHIVO, sheet_name="TAREAS")
        for col in ["FECHA_CREACION","FECHA_COMPROMISO","FECHA_CIERRE"]:
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
            if df[col].dtype == object or col in ("DESCRIPCION","TIPO","CATEGORIA",
                "PRIORIDAD","ESTADO","TERCERO","PROYECTO","AREA","NOTAS"):
                df[col] = df[col].fillna(val)
            else:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(val)
        df = df[df["TAREA"].astype(str).str.strip() != ""].copy()
        # ORDEN: columna de ordenamiento persistente dentro de cada grupo
        if "ORDEN" not in df.columns:
            df["ORDEN"] = list(range(0, len(df) * 10, 10))
        else:
            df["ORDEN"] = pd.to_numeric(df["ORDEN"], errors="coerce")
            max_ord = int(df["ORDEN"].max() or 0)
            df["ORDEN"] = df["ORDEN"].fillna(max_ord + 10).astype(int)
        df["PRIO_ORD"] = df["PRIORIDAD"].map(PRIO_ORD).fillna(99)
        # URGENCIA dinámica: si col es 0 o NaN, calcular desde fecha
        def _urg(row):
            v = row.get("URGENCIA", 0)
            if v and v > 0: return v
            fc = row.get("FECHA_COMPROMISO")
            if pd.isna(fc): return 1
            d = (fc.date() - HOY).days
            if d < 0:   return 5
            if d <= 3:  return 5
            if d <= 7:  return 4
            if d <= 14: return 3
            if d <= 30: return 2
            return 1
        df["URGENCIA"] = df.apply(_urg, axis=1)
        data["tareas"] = df
    except Exception as e:
        data["tareas"] = pd.DataFrame()
        data["_error_tareas"] = str(e)
    try:
        dt = pd.read_excel(ARCHIVO, sheet_name="TERCEROS")
        for col in ["FECHA_INICIO_SEG","FECHA_ULTIMO_SEG"]:
            if col in dt.columns:
                dt[col] = pd.to_datetime(dt[col], errors="coerce").astype("datetime64[us]")
        # Días sin respuesta
        if "FECHA_ULTIMO_SEG" in dt.columns:
            dt["DIAS_SIN_RESP"] = dt["FECHA_ULTIMO_SEG"].apply(
                lambda x: (HOY - x.date()).days if pd.notna(x) else 999
            )
        data["terceros"] = dt
    except:
        data["terceros"] = pd.DataFrame()
    return data

# ─── BUSINESS LOGIC ───────────────────────────────────────────────────────────
def _activas():
    return df_raw[~df_raw["ESTADO"].isin(["Completada","Cancelada"])]

def dias_restantes(ts):
    if pd.isna(ts): return None
    return (ts.date() - HOY).days

def semaforo(ts):
    d = dias_restantes(ts)
    if d is None: return C_GRIS, "Sin fecha"
    if d < 0:    return C_CRITICO, f"Vencida {abs(d)}d"
    if d == 0:   return C_CRITICO, "Vence HOY"
    if d <= 3:   return C_ALERTA,  f"Vence en {d}d"
    if d <= 7:   return C_CIAN,    f"Vence en {d}d"
    return "#475569",              f"Vence en {d}d"

def estado_general():
    ac = _activas()
    venc = ac[ac["FECHA_COMPROMISO"].notna() & (ac["FECHA_COMPROMISO"] < HOY_TS)]
    crit = ac[ac["PRIORIDAD"] == "Crítica"]
    crit_venc = venc[venc["PRIORIDAD"] == "Crítica"]
    if len(crit_venc) > 0 or len(venc) > 5:
        return "CRÍTICO", C_CRITICO, "⛔"
    elif len(venc) > 0 or len(crit) >= 2:
        return "BAJO PRESIÓN", C_ALERTA, "⚠️"
    return "CONTROLADO", C_OK, "✅"

def tasa_cumplimiento_semana():
    ac = _activas()
    ini = HOY - timedelta(days=HOY.weekday())
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

# ─── GITHUB SAVE ──────────────────────────────────────────────────────────────
_GH_REPO   = "Jochoa27/jochoa-tareas"
_GH_FILE   = "TAREAS_JOCHOA.xlsx"
_GH_BRANCH = "main"

def _token():
    try:
        return st.secrets.get("GITHUB_TOKEN", "")
    except Exception:
        return ""

@st.cache_data(ttl=60)
def _get_file_cache_key():
    """SHA del archivo en GitHub (TTL 60 s). Fallback a mtime local."""
    tok = _token()
    if tok:
        try:
            hdrs = {"Authorization": f"Bearer {tok}",
                    "Accept": "application/vnd.github+json"}
            r = requests.get(
                f"https://api.github.com/repos/{_GH_REPO}/contents/{_GH_FILE}",
                headers=hdrs, timeout=5)
            if r.status_code == 200:
                return r.json()["sha"]
        except Exception:
            pass
    return str(ARCHIVO.stat().st_mtime if ARCHIVO.exists() else 0)

_cache_key = _get_file_cache_key()
_data  = cargar(_cache_key)
df_raw = _data.get("tareas", pd.DataFrame())
df_ter = _data.get("terceros", pd.DataFrame())

def guardar_github(df_tareas_nuevo, df_ter_nuevo=None):
    """Sobreescribe TAREAS (y opcionalmente TERCEROS) en el repo vía GitHub API y limpia caché."""
    token = _token()
    if not token:
        return False, (
            "GITHUB_TOKEN no configurado. "
            "Ve a Streamlit Cloud → Manage app → Secrets y agrega:\n"
            "GITHUB_TOKEN = \"ghp_tu_token_aqui\""
        )
    hdrs   = {"Authorization": f"Bearer {token}",
              "Accept": "application/vnd.github+json"}
    api    = f"https://api.github.com/repos/{_GH_REPO}/contents/{_GH_FILE}"

    # SHA del archivo actual
    r = requests.get(api, headers=hdrs, timeout=10)
    if r.status_code != 200:
        return False, f"Error leyendo repo ({r.status_code}): {r.json().get('message','')}"
    sha = r.json()["sha"]

    # Construir Excel multi-hoja en memoria
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        # Normalizar tipos antes de escribir
        df_save = df_tareas_nuevo.copy()
        for c in ["FECHA_CREACION","FECHA_COMPROMISO","FECHA_CIERRE"]:
            if c in df_save.columns:
                df_save[c] = pd.to_datetime(df_save[c], errors="coerce")
        # Eliminar columnas internas
        df_save = df_save.drop(columns=["PRIO_ORD"], errors="ignore")
        df_save.to_excel(writer, sheet_name="TAREAS", index=False)
        _ter_base = df_ter_nuevo if df_ter_nuevo is not None else df_ter
        ter_save  = _ter_base.drop(columns=["DIAS_SIN_RESP"], errors="ignore")
        if not ter_save.empty:
            ter_save.to_excel(writer, sheet_name="TERCEROS", index=False)
    buf.seek(0)
    raw_bytes = buf.read()  # leer una sola vez para API y escritura local

    ts      = pd.Timestamp.now().strftime("%d/%m %H:%M")
    payload = {
        "message": f"update: tareas actualizadas desde app ({ts})",
        "content": base64.b64encode(raw_bytes).decode(),
        "sha":     sha,
        "branch":  _GH_BRANCH,
    }
    r2 = requests.put(api, json=payload, headers=hdrs, timeout=15)
    if r2.status_code in (200, 201):
        try:
            ARCHIVO.write_bytes(raw_bytes)  # sincronizar archivo local con GitHub
        except Exception:
            pass
        cargar.clear()
        _get_file_cache_key.clear()
        return True, f"Guardado a las {ts}"
    return False, f"Error al guardar ({r2.status_code}): {r2.json().get('message','')}"

def _marcar_estado(task_id, nuevo_estado):
    """Cambia el estado de una tarea y guarda en GitHub."""
    df_copy = df_raw.copy()
    mask = df_copy["ID"] == task_id
    df_copy.loc[mask, "ESTADO"] = nuevo_estado
    if nuevo_estado == "Completada":
        df_copy.loc[mask, "FECHA_CIERRE"] = pd.Timestamp(HOY)
    with st.spinner("Guardando..."):
        ok, msg = guardar_github(df_copy)
    if ok:
        st.toast("Tarea actualizada", icon="✅")
        st.rerun()
    else:
        st.error(msg)

def _merge_edits(df_original, df_edited):
    """Fusiona filas editadas al DataFrame completo usando update() por índice ID."""
    df_result = df_original.copy()
    df_ed     = df_edited.copy()

    # Normalizar tipos para compatibilidad con df_result
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

    # Actualizar filas existentes vía update() alineado por ID
    if not existing.empty:
        df_result = df_result.set_index("ID")
        existing  = existing.set_index("ID")
        cols_upd  = [c for c in existing.columns if c in df_result.columns]
        df_result.update(existing[cols_upd])
        df_result = df_result.reset_index()

    # Agregar filas nuevas
    if not new_rows.empty:
        max_id = int(df_result["ID"].max() or 0)
        new_rows = new_rows.copy()
        new_rows["ID"]             = list(range(max_id + 1, max_id + 1 + len(new_rows)))
        new_rows["FECHA_CREACION"] = pd.Timestamp(HOY)
        for col, default in [("ESTADO","Pendiente"),("PRIORIDAD","Media"),
                              ("TIPO","Tarea"),("AREA","Trabajo")]:
            if col in new_rows.columns:
                new_rows[col] = new_rows[col].fillna(default)
            else:
                new_rows[col] = default
        # Alinear columnas con df_result
        for col in df_result.columns:
            if col not in new_rows.columns:
                new_rows[col] = pd.NA
        df_result = pd.concat(
            [df_result, new_rows[df_result.columns]], ignore_index=True)

    return df_result

# ─── HELPERS VISUALES ─────────────────────────────────────────────────────────
def seccion(icon, titulo, color=C_CIAN):
    st.markdown(
        f'<div class="sh">'
        f'<div class="sh-bar" style="background:linear-gradient(180deg,{color},{color}44);'
        f'box-shadow:0 0 10px {color}88;"></div>'
        f'<div class="sh-txt">{icon} {titulo}</div></div>',
        unsafe_allow_html=True)

def kpi(label, valor, sub=None, color=C_CIAN):
    sub_h = f'<div class="kpi-sub" style="color:{color};">{sub}</div>' if sub else ""
    return (
        f'<div class="kpi-card" style="background:linear-gradient(135deg,{color}18,{color}04);'
        f'border:1px solid {color}30;">'
        f'<div class="kpi-lbl">{label}</div>'
        f'<div class="kpi-val" style="color:{color};text-shadow:0 0 14px {color}55;">{valor}</div>'
        f'{sub_h}</div>'
    )

def badge_estado(texto, color):
    return (f'<span class="estado-badge" style="background:{color}18;'
            f'border:1px solid {color}44;color:{color};">{texto}</span>')

def chip(texto, color):
    return (f'<span class="tc-tag" style="background:{color}16;'
            f'color:{color};border:1px solid {color}30;">{texto}</span>')

def ec(option, height=260):
    s = json.dumps(option, ensure_ascii=False)
    # Inyectar funciones JS
    s = s.replace('"__SYM_FN__"',   'function(v){return Math.max(v[2]*11,8);}')
    s = s.replace('"__TT_FN__"',    'function(p){var d=p.data;return'
                  ' "<b>"+d[3]+"</b><br/>Urgencia: "+d[0]+"/5<br/>Impacto: "'
                  '+d[1]+"/5<br/>Esfuerzo: "+d[2]+"h";}')
    s = s.replace('"__TT_FN_DIAG__"',
                  'function(p){var d=p.data;var r=\'\';\n'
                  'r+=\'<div style="font-weight:800;font-size:12px;margin-bottom:5px;">\'+d[3]+\'</div>\';\n'
                  'r+=d[4]?\'<div style="color:#94A3B8;font-size:10px;">🏢 \'+d[4]+\'</div>\':\'\';\n'
                  'r+=\'<div style="margin-top:4px;">⚙️ Urgencia \'+d[0]+\'/5 · Impacto \'+d[1]+\'/5 · \'+d[2]+\'h</div>\';\n'
                  'r+=d[5]?\'<div>👤 \'+d[5]+\'</div>\':\'\';\n'
                  'r+=d[6]?\'<div>📅 \'+d[6]+\'</div>\':\'\';\n'
                  'r+=d[7]?\'<div style="color:#FFB300;font-weight:700;margin-top:3px;">\'+d[7]+\'</div>\':\'\';\n'
                  'return r;}')
    html = (f'<!DOCTYPE html><html><head>'
            f'<script src="https://cdn.jsdelivr.net/npm/echarts@5.4.3/dist/echarts.min.js">'
            f'</script></head><body style="margin:0;padding:0;background:transparent;">'
            f'<div id="c" style="width:100%;height:{height}px;"></div>'
            f'<script>var ch=echarts.init(document.getElementById("c"),null,'
            f'{{renderer:"svg",width:"auto",height:{height}}});ch.setOption({s});'
            f'window.addEventListener("resize",function(){{ch.resize();}});</script>'
            f'</body></html>')
    components.html(html, height=height + 10, scrolling=False)

_TT = {
    "trigger": "item",
    "backgroundColor": "rgba(6,11,21,0.97)",
    "borderColor":     "rgba(56,189,248,0.22)",
    "borderRadius":    12,
    "textStyle":       {"color": "#F1F5F9", "fontSize": 11},
}
_TT_AXIS = {**_TT, "trigger": "axis"}

def _axis(color="#38BDF8"):
    return {"axisLine": {"lineStyle": {"color": "#1E293B"}},
            "splitLine": {"lineStyle": {"color": "rgba(255,255,255,0.04)"}},
            "axisLabel": {"color": "#64748B", "fontSize": 10}}

# ─── SIDEBAR ──────────────────────────────────────────────────────────────────
MODULOS = [
    ("🎯", "Centro de Comando"),
    ("📆", "Resumen Semanal"),
    ("⚠️", "Diagnóstico de Riesgo"),
    ("📋", "Bandeja Operacional"),
    ("👥", "Seguimiento de Terceros"),
    ("📈", "Productividad"),
    ("⏱", "Consumo de Tiempo"),
    ("📊", "Gantt"),
]

if "mod" not in st.session_state:
    st.session_state.mod = "Centro de Comando"

with st.sidebar:
    st.markdown(
        '<div style="padding:16px 4px 20px;">'
        '<div style="font-size:0.55rem;font-weight:800;letter-spacing:0.30em;'
        'color:#38BDF8;margin-bottom:2px;text-shadow:0 0 10px rgba(56,189,248,0.40);">'
        'P&C · CONSTRUCTORA LONDRES</div>'
        '<div style="font-size:1.55rem;font-weight:900;color:#F8FAFC;line-height:1.0;'
        'margin-bottom:4px;">🎯 CENTRO DE<br>&nbsp;&nbsp;&nbsp;COMANDO</div>'
        '<div style="font-size:0.62rem;color:#475569;font-weight:600;">Juan Ochoa · 2026</div>'
        '</div>', unsafe_allow_html=True)

    st.markdown('<div style="height:1px;background:rgba(56,189,248,0.08);margin-bottom:12px;"></div>',
                unsafe_allow_html=True)
    st.markdown(
        '<div style="font-size:0.50rem;font-weight:800;letter-spacing:0.16em;'
        'color:#334155;text-transform:uppercase;margin-bottom:8px;">NAVEGACIÓN</div>',
        unsafe_allow_html=True)

    for ico, nombre in MODULOS:
        activo = st.session_state.mod == nombre
        label  = f"{ico}  {nombre}"
        if activo:
            st.markdown(
                f'<div style="background:rgba(56,189,248,0.14);border:1px solid rgba(56,189,248,0.38);'
                f'border-radius:10px;padding:10px 14px;margin-bottom:3px;color:#38BDF8;'
                f'font-weight:800;font-size:0.80rem;letter-spacing:0.04em;'
                f'box-shadow:0 0 12px rgba(56,189,248,0.10);">{label}</div>',
                unsafe_allow_html=True)
        else:
            if st.button(label, key=f"nav_{nombre}", use_container_width=True):
                st.session_state.mod = nombre
                st.rerun()

    st.markdown('<div style="height:1px;background:rgba(56,189,248,0.06);margin:16px 0 12px;"></div>',
                unsafe_allow_html=True)
    # Estado rápido
    eg_txt, eg_clr, eg_ico = estado_general()
    st.markdown(
        f'<div style="text-align:center;padding:8px 6px;">'
        f'<div style="font-size:0.48rem;font-weight:800;letter-spacing:0.16em;'
        f'color:#334155;text-transform:uppercase;margin-bottom:6px;">ESTADO SISTEMA</div>'
        f'<div style="background:{eg_clr}18;border:1px solid {eg_clr}44;'
        f'border-radius:12px;padding:6px 12px;color:{eg_clr};font-weight:800;'
        f'font-size:0.72rem;letter-spacing:0.08em;">{eg_ico} {eg_txt}</div>'
        f'<div style="font-size:0.48rem;color:#334155;margin-top:8px;">'
        f'{HOY.strftime("%A %d de %B")}</div>'
        f'</div>', unsafe_allow_html=True)

    # ── Selector de tema ──────────────────────────────────────────────────────
    st.markdown(
        '<div style="height:1px;background:rgba(255,255,255,0.05);margin:14px 0 10px;"></div>',
        unsafe_allow_html=True)
    # Swatches 1×3
    _sw_html = ('<div style="font-size:0.50rem;font-weight:800;letter-spacing:0.14em;'
                'color:#52525B;text-transform:uppercase;margin-bottom:7px;">🎨 TEMA</div>'
                '<div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:5px;margin-bottom:8px;">')
    for _tkey, _tcc in _THEMES.items():
        _asel = _tkey == _TK
        _rng = f"box-shadow:0 0 0 2px rgba({_tcc['abr']},0.90),0 0 10px rgba({_tcc['abr']},0.25);" if _asel else "border:1px solid rgba(255,255,255,0.08);"
        _sw_html += (
            f'<div style="border-radius:9px;overflow:hidden;{_rng}">'
            f'<div style="background:{_tcc["bg_css"]};height:30px;display:flex;align-items:center;'
            f'justify-content:center;font-size:1.05rem;">{_tcc["icon"]}</div>'
            f'<div style="background:{_tcc["bg_sb"]};padding:3px 0;text-align:center;'
            f'font-size:0.46rem;font-weight:800;letter-spacing:0.06em;'
            f'color:rgba({_tcc["abr"]},1);">{"✓ " if _asel else ""}{_tcc["name"].upper()}</div>'
            f'</div>'
        )
    _sw_html += '</div>'
    st.markdown(_sw_html, unsafe_allow_html=True)
    _tha, _thb, _thc = st.columns(3)
    for _ti, (_tkey, _tcc) in enumerate(_THEMES.items()):
        _col = [_tha, _thb, _thc][_ti]
        with _col:
            if st.button(f"{'✓' if _tkey==_TK else _tcc['icon']}", key=f"th_{_tkey}",
                         use_container_width=True, help=f"{_tcc['name']} — {_tcc['desc']}"):
                st.session_state["theme"] = _tkey
                st.rerun()

    # Aviso de token
    if not _token():
        st.markdown(
            '<div style="background:rgba(255,179,0,0.08);border:1px solid rgba(255,179,0,0.24);'
            'border-radius:10px;padding:10px 12px;margin-top:10px;">'
            '<div style="font-size:0.58rem;font-weight:800;color:#FFB300;margin-bottom:4px;">'
            '⚠️ EDICIÓN DESACTIVADA</div>'
            '<div style="font-size:0.56rem;color:#64748B;line-height:1.5;">'
            'Agrega GITHUB_TOKEN en Streamlit Cloud → Manage app → Secrets para habilitar '
            'guardar desde la plataforma.</div>'
            '</div>', unsafe_allow_html=True)

mod = st.session_state.mod

# ══════════════════════════════════════════════════════════════════════════════
# MÓDULO 1: CENTRO DE COMANDO
# ══════════════════════════════════════════════════════════════════════════════
if mod == "Centro de Comando":
    ac     = _activas()
    comps  = df_raw[df_raw["ESTADO"] == "Completada"]
    venc   = ac[ac["FECHA_COMPROMISO"].notna() & (ac["FECHA_COMPROMISO"] < HOY_TS)]
    hoy_t  = ac[ac["FECHA_COMPROMISO"].notna() & (ac["FECHA_COMPROMISO"] == HOY_TS)]
    próx   = ac[ac["FECHA_COMPROMISO"].notna() &
                (ac["FECHA_COMPROMISO"] >= HOY_TS) &
                (ac["FECHA_COMPROMISO"] <= HOY_TS + pd.Timedelta(days=3))]
    crit   = ac[ac["PRIORIDAD"] == "Crítica"]
    terc_p = ac[ac["ESTADO"] == "Esperando Terceros"]
    ini_sem = HOY - timedelta(days=HOY.weekday())
    comp_s = comps[comps["FECHA_CIERRE"].notna() & (comps["FECHA_CIERRE"] >= pd.Timestamp(ini_sem))]
    comp_m = comps[
        comps["FECHA_CIERRE"].notna() &
        (comps["FECHA_CIERRE"].dt.month == HOY.month) &
        (comps["FECHA_CIERRE"].dt.year  == HOY.year)
    ]
    tasa_s, _ = tasa_cumplimiento_semana()
    proj_d = ac.groupby("PROYECTO").size()
    proj_top = proj_d.idxmax() if not proj_d.empty else "—"
    eg_txt, eg_clr, eg_ico = estado_general()

    # ── Banner ────────────────────────────────────────────────────────────────
    st.markdown(
        f'<div style="background:linear-gradient(100deg,rgba(56,189,248,0.07),transparent);'
        f'border:1px solid rgba(56,189,248,0.14);border-radius:20px;padding:18px 26px;margin-bottom:18px;'
        f'display:flex;align-items:center;justify-content:space-between;">'
        f'<div>'
        f'<div style="font-size:0.58rem;font-weight:800;letter-spacing:0.28em;color:#38BDF8;margin-bottom:3px;">'
        f'PLANIFICACIÓN Y CONTROL · CONSTRUCTORA LONDRES</div>'
        f'<div style="font-size:2.10rem;font-weight:900;color:#F8FAFC;line-height:1.05;">🎯 CENTRO DE COMANDO</div>'
        f'<div style="font-size:0.72rem;color:#64748B;margin-top:4px;font-weight:600;">'
        f'Juan Ochoa · {HOY.strftime("%d de %B de %Y")}</div>'
        f'</div>'
        f'<div style="text-align:right;">'
        f'<div style="background:{eg_clr}18;border:2px solid {eg_clr}44;border-radius:16px;'
        f'padding:10px 22px;color:{eg_clr};font-weight:900;font-size:0.85rem;letter-spacing:0.10em;">'
        f'{eg_ico} {eg_txt}</div>'
        f'<div style="font-size:0.52rem;color:#334155;margin-top:8px;">ESTADO GENERAL DEL SISTEMA</div>'
        f'</div></div>',
        unsafe_allow_html=True)

    # ── KPIs fila 1 ───────────────────────────────────────────────────────────
    c1,c2,c3,c4,c5,c6 = st.columns(6)
    with c1: st.markdown(kpi("TAREAS ACTIVAS",     len(ac),     color=C_CIAN),  unsafe_allow_html=True)
    with c2: st.markdown(kpi("CRÍTICAS",            len(crit),   "requieren atención",
                              color=C_CRITICO if len(crit) else C_GRIS),             unsafe_allow_html=True)
    with c3: st.markdown(kpi("VENCIDAS",            len(venc),   "fuera de plazo",
                              color=C_CRITICO if len(venc) else C_OK),               unsafe_allow_html=True)
    with c4: st.markdown(kpi("VENCEN EN 3 DÍAS",   len(próx),   "alta prioridad",
                              color=C_ALERTA  if len(próx) else C_GRIS),             unsafe_allow_html=True)
    with c5: st.markdown(kpi("ESPERA TERCEROS",     len(terc_p), "bloqueadas",
                              color=C_ALERTA  if len(terc_p) else C_GRIS),          unsafe_allow_html=True)
    with c6: st.markdown(kpi("COMPLETADAS",         len(comps),  color=C_OK),   unsafe_allow_html=True)

    st.markdown('<div style="height:8px;"></div>', unsafe_allow_html=True)

    # ── KPIs fila 2 ───────────────────────────────────────────────────────────
    c7,c8,c9,c10 = st.columns(4)
    tasa_clr = C_OK if tasa_s >= 70 else C_ALERTA if tasa_s >= 40 else C_CRITICO
    _both = comps.dropna(subset=["FECHA_CIERRE","FECHA_CREACION"])
    avg_d = (_both["FECHA_CIERRE"] - _both["FECHA_CREACION"]).dt.days.mean() if not _both.empty else None
    with c7:  st.markdown(kpi("COMP. ESTA SEMANA", len(comp_s), color=C_OK),              unsafe_allow_html=True)
    with c8:  st.markdown(kpi("COMP. ESTE MES",    len(comp_m), color=C_OK),              unsafe_allow_html=True)
    with c9:  st.markdown(kpi("T° PROM. CIERRE",   f"{avg_d:.0f}d" if avg_d else "—",
                               "tiempo promedio",    color=C_CIAN),                        unsafe_allow_html=True)
    with c10: st.markdown(kpi("TASA SEMANAL",      f"{tasa_s}%", "cumplimiento",
                               color=tasa_clr),                                            unsafe_allow_html=True)

    st.markdown('<div style="height:4px;"></div>', unsafe_allow_html=True)

    # ── Inputs compartidos JS→Python (agenda + calendario) ────────────────────
    # Limpiar widget del rerun anterior ANTES de renderizarlo (regla de Streamlit)
    if st.session_state.pop("_clr_action", False):
        st.session_state["ag_task_action"] = ""
    if "ag_task_action" not in st.session_state:
        st.session_state["ag_task_action"] = ""
    if "ag_drag_order" not in st.session_state:
        st.session_state["ag_drag_order"] = ""
    ag_action_raw = st.text_input("TASK_ACTION", key="ag_task_action")
    ag_order_raw  = st.text_input("DRAG_ORDER",  key="ag_drag_order")
    st.markdown(
        '<style>'
        'div[data-testid="stTextInput"]:has(input[aria-label="TASK_ACTION"]),'
        'div[data-testid="stTextInput"]:has(input[aria-label="DRAG_ORDER"])'
        '{position:absolute!important;left:-9999px!important;width:1px!important;'
        'height:1px!important;overflow:hidden!important;}</style>',
        unsafe_allow_html=True)

    # ── Procesar acciones (compartido) ────────────────────────────────────────
    if ag_action_raw.strip():
        try:
            _pts = ag_action_raw.strip().split(":", 2)
            _aty = _pts[0] if len(_pts) >= 2 else "complete"
            # newform no tiene tid numérico — manejar antes del int()
            if _aty == "newform":
                _fn2 = _pts[1] if len(_pts) >= 2 else ""
                _fv2 = _pts[2] if len(_pts) >= 3 else ""
                st.session_state["add_task_defaults"] = {"field": _fn2, "value": _fv2}
                st.session_state["_clr_action"] = True
                st.rerun()
            # reorder: _pts = ["reorder", "<group>", "id1,id2,..."]
            if _aty == "reorder":
                _ids_str = _pts[2] if len(_pts) >= 3 else ""
                st.session_state["_clr_action"] = True
                if _token() and _ids_str:
                    _new_ids = [int(x) for x in _ids_str.split(",") if x.strip()]
                    if _new_ids:
                        _dfc_r = df_raw.copy()
                        if "ORDEN" not in _dfc_r.columns:
                            _dfc_r["ORDEN"] = list(range(0, len(_dfc_r) * 10, 10))
                        _cur_ords = sorted(
                            _dfc_r.loc[_dfc_r["ID"].isin(_new_ids), "ORDEN"]
                            .dropna().astype(int).tolist()
                        )
                        while len(_cur_ords) < len(_new_ids):
                            _cur_ords.append((_cur_ords[-1] if _cur_ords else 0) + 10)
                        for _tid_r, _ord_v in zip(_new_ids, _cur_ords):
                            _dfc_r.loc[_dfc_r["ID"] == _tid_r, "ORDEN"] = _ord_v
                        with st.spinner("Guardando orden..."):
                            _ok_r, _ms_r = guardar_github(_dfc_r)
                        if _ok_r:
                            st.toast("✅ Orden guardado"); st.rerun()
                        else:
                            st.error(_ms_r)
                st.rerun()
            # copy: abrir formulario de nueva tarea pre-cargado con datos de la tarea origen
            if _aty == "copy":
                _src = df_raw[df_raw["ID"] == int(_pts[1])] if len(_pts) >= 2 else pd.DataFrame()
                if not _src.empty:
                    _sr = _src.iloc[0]
                    st.session_state["add_task_defaults"] = {
                        "mode":  "copy",
                        "tarea": str(_sr.get("TAREA", "") or "") + " (copia)",
                        "proj":  str(_sr.get("PROYECTO", "") or ""),
                        "tipo":  str(_sr.get("TIPO", "Tarea") or "Tarea"),
                        "area":  str(_sr.get("AREA", "Trabajo") or "Trabajo"),
                        "prio":  str(_sr.get("PRIORIDAD", "Media") or "Media"),
                        "cat":   str(_sr.get("CATEGORIA", "Planificación") or "Planificación"),
                        "est":   "Pendiente",
                        "desc":  str(_sr.get("NOTAS", "") or ""),
                        "ter":   str(_sr.get("TERCERO", "") or ""),
                        "esf":   float(_sr.get("ESFUERZO_HRS", 0) or 0),
                    }
                st.session_state["_clr_action"] = True
                st.rerun()
            # delete: eliminar tarea permanentemente
            if _aty == "delete":
                _del_id = int(_pts[1]) if len(_pts) >= 2 else None
                st.session_state["_clr_action"] = True
                if _del_id and _token():
                    _dfc_del = df_raw[df_raw["ID"] != _del_id].copy()
                    with st.spinner("Eliminando..."):
                        _ok_del, _ms_del = guardar_github(_dfc_del)
                    if _ok_del:
                        st.toast("✅ Tarea eliminada")
                        st.session_state.pop("detalle_id", None)
                        st.rerun()
                    else:
                        st.error(_ms_del)
                st.rerun()
            _tid = int(_pts[1] if len(_pts) >= 2 else _pts[0])
            _val = _pts[2] if len(_pts) >= 3 else ""
            st.session_state["_clr_action"] = True   # limpiar widget en PRÓXIMO rerun
            if _aty == "open":
                st.session_state["detalle_id"] = _tid
                st.rerun()
            elif _token():
                _dfc = df_raw.copy()
                _mk  = _dfc["ID"] == _tid
                if _aty == "date":
                    _dfc.loc[_mk, "FECHA_COMPROMISO"] = pd.Timestamp(_val) if _val.strip() else pd.NaT
                elif _aty == "field":
                    _fn, _fv = (_val.split(":", 1) + [""])[:2]
                    if _fn and _fn in _dfc.columns:
                        _dfc.loc[_mk, _fn] = _fv
                        if _fn == "ESTADO" and _fv == "Completada":
                            _dfc.loc[_mk, "FECHA_CIERRE"] = pd.Timestamp(HOY)
                else:
                    _dfc.loc[_mk, "ESTADO"] = _val
                    _dfc.loc[_mk, "FECHA_CIERRE"] = (
                        pd.Timestamp(HOY) if _val == "Completada" else pd.NaT)
                with st.spinner("Guardando..."):
                    _ok, _ms = guardar_github(_dfc)
                if _ok:
                    st.toast("✅ Guardado"); st.rerun()
                else:
                    st.error(_ms)
        except Exception:
            st.session_state["_clr_action"] = True

    # ── Toggle de vistas ──────────────────────────────────────────────────────
    _CC_VIEWS = [
        ("calendario", "📅 Kanban Semana"),
        ("estado",     "🔷 Kanban Estado"),
        ("prioridad",  "🎯 Kanban Prioridad"),
        ("area",       "📁 Kanban Área"),
        ("categoria",  "🗂 Kanban Categoría"),
        ("vencidas",   "🚨 Kanban Vencidas"),
    ]
    if st.session_state.get("cc_view") not in {v for v, _ in _CC_VIEWS}:
        st.session_state["cc_view"] = "calendario"
    _tv_cols = st.columns(len(_CC_VIEWS) + 1)
    for _vi, (_vk, _vl) in enumerate(_CC_VIEWS):
        with _tv_cols[_vi]:
            if st.button(_vl, key=f"btn_vw_{_vk}", use_container_width=True,
                         type="primary" if st.session_state["cc_view"]==_vk else "secondary"):
                st.session_state["cc_view"] = _vk; st.rerun()
    with _tv_cols[len(_CC_VIEWS)]:
        if st.button("➕ Nueva tarea", key="btn_global_new", use_container_width=True, type="secondary"):
            st.session_state["add_task_defaults"] = {"mode": "new"}
            st.rerun()
    st.markdown('<div style="height:6px;"></div>', unsafe_allow_html=True)

    # ── Filtros globales del Centro de Comando ────────────────────────────────
    _filt1, _filt2, _filt3 = st.columns([3, 2, 2])
    with _filt1:
        _cc_srch = st.text_input("", value=st.session_state.get("cc_search", ""),
                                  placeholder="🔍 Buscar tarea, proyecto o descripción...",
                                  key="cc_search_inp", label_visibility="collapsed")
        st.session_state["cc_search"] = _cc_srch
    with _filt2:
        _projs_cc = ["Todos los proyectos"] + sorted([p for p in ac["PROYECTO"].dropna().unique() if str(p).strip()])
        _cur_proj = st.session_state.get("cc_proj", "Todos los proyectos")
        if _cur_proj not in _projs_cc:
            _cur_proj = "Todos los proyectos"
        _cc_proj = st.selectbox("", _projs_cc, key="cc_proj_sel",
                                 label_visibility="collapsed",
                                 index=_projs_cc.index(_cur_proj))
        st.session_state["cc_proj"] = _cc_proj
    with _filt3:
        _areas_cc = ["Todas las áreas"] + sorted([a for a in ac["AREA"].dropna().unique() if str(a).strip()]) if "AREA" in ac.columns else ["Todas las áreas"]
        _cur_area = st.session_state.get("cc_area", "Todas las áreas")
        if _cur_area not in _areas_cc:
            _cur_area = "Todas las áreas"
        _cc_area_view = st.session_state.get("cc_view", "calendario")
        _area_disabled = (_cc_area_view == "area")
        _cc_area = st.selectbox(
            "", _areas_cc, key="cc_area_sel",
            label_visibility="collapsed",
            index=_areas_cc.index(_cur_area),
            disabled=_area_disabled,
            help="No aplica en la vista Kanban Área" if _area_disabled else None,
        )
        st.session_state["cc_area"] = _cc_area
    # Aplicar filtros a 'ac' (KPIs ya calculados arriba no se ven afectados)
    if _cc_srch.strip():
        _q = _cc_srch.strip()
        _mask_srch = (
            ac["TAREA"].str.contains(_q, case=False, na=False) |
            ac["PROYECTO"].str.contains(_q, case=False, na=False)
        )
        if "DESCRIPCION" in ac.columns:
            _mask_srch = _mask_srch | ac["DESCRIPCION"].str.contains(_q, case=False, na=False)
        ac = ac[_mask_srch]
    if _cc_proj != "Todos los proyectos":
        ac = ac[ac["PROYECTO"] == _cc_proj]
    if _cc_area != "Todas las áreas" and not _area_disabled and "AREA" in ac.columns:
        ac = ac[ac["AREA"] == _cc_area]

    # ── Formulario nueva tarea (abierto por "+" de kanban / calendario) ────────
    _OPTS_EST_NT  = ["Pendiente","En Proceso","Esperando Terceros","Completada","Cancelada"]
    _OPTS_PRIO_NT = ["Crítica","Alta","Media","Baja"]
    _OPTS_TIPO_NT = ["Tarea","Seguimiento","Compromiso","Reunión"]
    _OPTS_CAT_NT  = ["Planificación","Contratos","Compras","Reportes",
                     "IA y Automatización","Gestión Corporativa","Reuniones","Salud","Personal"]
    _OPTS_AREA_NT = ["Trabajo","Personal"]

    if st.session_state.get("add_task_defaults") is not None:
        _ntd      = st.session_state["add_task_defaults"]
        _is_copy  = _ntd.get("mode") == "copy"
        _nt_fld   = _ntd.get("field", "")
        _nt_val   = _ntd.get("value", "")
        _def_tarea = _ntd.get("tarea", "") if _is_copy else ""
        _def_proj  = _ntd.get("proj",  "") if _is_copy else ""
        _def_tipo  = _ntd.get("tipo", "Tarea") if _is_copy else "Tarea"
        _def_desc  = _ntd.get("desc",  "") if _is_copy else ""
        _def_est   = _ntd.get("est",  "Pendiente") if _is_copy else (_nt_val if _nt_fld == "ESTADO"    else "Pendiente")
        _def_prio  = _ntd.get("prio", "Media")     if _is_copy else (_nt_val if _nt_fld == "PRIORIDAD" else "Media")
        _def_area  = _ntd.get("area", "Trabajo")   if _is_copy else (_nt_val if _nt_fld == "AREA"      else "Trabajo")
        _def_cat   = _ntd.get("cat",  "Planificación") if _is_copy else (_nt_val if _nt_fld == "CATEGORIA" else "Planificación")
        _def_ter   = _ntd.get("ter",  "") if _is_copy else ""
        _def_esf   = float(_ntd.get("esf", 0.0) or 0.0) if _is_copy else 0.0
        try:
            _def_date = pd.Timestamp(_nt_val).date() if _nt_fld == "date" and _nt_val else None
        except Exception:
            _def_date = None
        _form_titulo = "DUPLICAR TAREA" if _is_copy else "NUEVA TAREA"
        seccion("📋" if _is_copy else "➕", _form_titulo, C_CIAN)
        with st.form("frm_add_task", border=False):
            _nt1, _nt2, _nt3 = st.columns([3, 2, 2])
            with _nt1:
                _nt_tarea = st.text_input("Tarea *", value=_def_tarea, placeholder="Nombre de la tarea...")
                _nt_proj  = st.text_input("Proyecto", value=_def_proj, placeholder="Nombre del proyecto")
            with _nt2:
                _nt_tipo  = st.selectbox("Tipo", _OPTS_TIPO_NT,
                    index=_OPTS_TIPO_NT.index(_def_tipo) if _def_tipo in _OPTS_TIPO_NT else 0)
                _nt_area  = st.selectbox("Área", _OPTS_AREA_NT,
                    index=_OPTS_AREA_NT.index(_def_area) if _def_area in _OPTS_AREA_NT else 0)
            with _nt3:
                _nt_prio  = st.selectbox("Prioridad", _OPTS_PRIO_NT,
                    index=_OPTS_PRIO_NT.index(_def_prio) if _def_prio in _OPTS_PRIO_NT else 2)
                _nt_cat   = st.selectbox("Categoría", _OPTS_CAT_NT,
                    index=_OPTS_CAT_NT.index(_def_cat) if _def_cat in _OPTS_CAT_NT else 0)
            _nt4, _nt5, _nt6 = st.columns([2, 2, 1])
            with _nt4:
                _nt_est   = st.selectbox("Estado", _OPTS_EST_NT,
                    index=_OPTS_EST_NT.index(_def_est) if _def_est in _OPTS_EST_NT else 0)
            with _nt5:
                _nt_ter   = st.text_input("Tercero", value=_def_ter, placeholder="Nombre del tercero...")
            with _nt6:
                _nt_fecha = st.date_input("Fecha compromiso", value=_def_date)
            _nt7a, _nt7b = st.columns([3, 1])
            with _nt7b:
                _nt_esf = st.number_input("Esfuerzo (hrs)", min_value=0.0, step=0.5,
                                          format="%.1f", value=_def_esf)
            with _nt7a:
                _nt_desc = st.text_area("Descripción / Notas", value=_def_desc, height=60, placeholder="Detalle opcional...")
            _ntb1, _ntb2 = st.columns(2)
            with _ntb1: _nt_sub = st.form_submit_button("📋 Duplicar" if _is_copy else "➕ Agregar", use_container_width=True, type="primary")
            with _ntb2: _nt_cls = st.form_submit_button("✕ Cancelar", use_container_width=True)
        if _nt_cls:
            st.session_state.pop("add_task_defaults", None); st.rerun()
        if _nt_sub:
            if not _nt_tarea.strip():
                st.warning("⚠️ Ingresa el nombre de la tarea")
            elif not _token():
                st.error("GITHUB_TOKEN no configurado")
            else:
                _nid  = int(df_raw["ID"].max() or 0) + 1
                _nrow = {"ID":_nid,"TAREA":_nt_tarea.strip(),"TIPO":_nt_tipo,
                         "PROYECTO":_nt_proj.strip(),"AREA":_nt_area,"CATEGORIA":_nt_cat,
                         "PRIORIDAD":_nt_prio,"ESTADO":_nt_est,
                         "TERCERO":_nt_ter.strip(),"ESFUERZO_HRS":_nt_esf,
                         "FECHA_COMPROMISO":pd.Timestamp(_nt_fecha) if _nt_fecha else pd.NaT,
                         "FECHA_CREACION":pd.Timestamp(HOY),"FECHA_CIERRE":pd.NaT,
                         "DESCRIPCION":_nt_desc,"NOTAS":"","COMENTARIOS":""}
                _ndf = pd.concat([df_raw, pd.DataFrame([_nrow])], ignore_index=True)
                for _c in ["FECHA_CREACION","FECHA_COMPROMISO","FECHA_CIERRE"]:
                    if _c in _ndf.columns:
                        _ndf[_c] = pd.to_datetime(_ndf[_c], errors="coerce").astype("datetime64[us]")
                with st.spinner("Guardando..."):
                    _ok2, _ms2 = guardar_github(_ndf)
                if _ok2:
                    st.session_state.pop("add_task_defaults", None)
                    st.toast(f"✅ Tarea creada"); st.rerun()
                else:
                    st.error(_ms2)

    # ── Panel de detalle de tarea ──────────────────────────────────────────────
    if st.session_state.get("detalle_id") is not None:
        _dt_id = st.session_state["detalle_id"]
        _dt_df = df_raw[df_raw["ID"] == _dt_id]
        if _dt_df.empty:
            st.session_state["detalle_id"] = None
        else:
            _r = _dt_df.iloc[0]
            _OPTS_EST2  = ["Pendiente","En Proceso","Esperando Terceros","Completada","Cancelada"]
            _OPTS_PRIO2 = ["Crítica","Alta","Media","Baja"]
            _OPTS_TIPO2 = ["Tarea","Seguimiento","Compromiso","Reunión"]
            _OPTS_CAT2  = ["Planificación","Contratos","Compras","Reportes",
                           "IA y Automatización","Gestión Corporativa","Reuniones","Salud","Personal"]
            _OPTS_AREA2 = ["Trabajo","Personal"]
            def _si(lst, v, d=0):
                try: return lst.index(str(v))
                except: return d

            st.markdown(
                f'<div style="background:rgba(13,21,38,0.96);border:1px solid {C_CIAN}35;'
                f'border-radius:16px;padding:14px 20px 4px;margin-bottom:6px;">'
                f'<div style="font-size:0.50rem;font-weight:800;letter-spacing:0.24em;'
                f'color:{C_CIAN};">DETALLE DE TAREA · ID {_dt_id}</div>'
                f'<div style="font-size:1.05rem;font-weight:900;color:#F8FAFC;margin-top:2px;">'
                f'{_r.get("TAREA","")}</div></div>',
                unsafe_allow_html=True)

            with st.form("frm_det", border=False):
                _fa, _fb, _fc = st.columns(3)
                with _fa:
                    _new_tarea = st.text_input("Nombre de la tarea",
                                               value=str(_r.get("TAREA","") or ""))
                    _new_est   = st.selectbox("Estado",    _OPTS_EST2,
                                              index=_si(_OPTS_EST2, _r.get("ESTADO","Pendiente")))
                    _new_prio  = st.selectbox("Prioridad", _OPTS_PRIO2,
                                              index=_si(_OPTS_PRIO2, _r.get("PRIORIDAD","Media")))
                    _new_tipo  = st.selectbox("Tipo",      _OPTS_TIPO2,
                                              index=_si(_OPTS_TIPO2, _r.get("TIPO","Tarea")))
                with _fb:
                    _new_proj = st.text_input("Proyecto",
                                              value=str(_r.get("PROYECTO","") or ""))
                    _new_area = st.selectbox("Área",      _OPTS_AREA2,
                                             index=_si(_OPTS_AREA2, _r.get("AREA","Trabajo")))
                    _new_cat  = st.selectbox("Categoría", _OPTS_CAT2,
                                             index=_si(_OPTS_CAT2,  _r.get("CATEGORIA","Planificación")))
                    _new_ter  = st.text_input("Tercero",
                                              value=str(_r.get("TERCERO","") or ""))
                with _fc:
                    _fc_raw  = _r.get("FECHA_COMPROMISO")
                    _fc_date = pd.Timestamp(_fc_raw).date() if pd.notna(_fc_raw) else None
                    _new_fc  = st.date_input("Fecha de vencimiento",
                                             value=_fc_date, format="DD/MM/YYYY")
                    _new_imp = st.number_input("Impacto (1-5)",  min_value=1, max_value=5, step=1,
                                               value=int(_r.get("IMPACTO",3) or 3))
                    _new_urg = st.number_input("Urgencia (1-5)", min_value=1, max_value=5, step=1,
                                               value=int(_r.get("URGENCIA",3) or 3))
                    _new_esf = st.number_input("Esfuerzo est. (hrs)", min_value=0.0, step=0.5,
                                               format="%.1f",
                                               value=float(_r.get("ESFUERZO_HRS",0) or 0))
                    _cur_est_det = str(_r.get("ESTADO",""))
                    _new_hr = st.number_input(
                        "Horas reales invertidas",
                        min_value=0.0, step=0.5, format="%.1f",
                        value=float(_r.get("HORAS_REALES",0) or 0),
                        help="Registra cuántas horas reales tomó esta tarea",
                        disabled=(_cur_est_det not in ("Completada","Cancelada")),
                    )

                _new_notas = st.text_area("Descripción / Notas",
                                          value=str(_r.get("NOTAS","") or ""), height=88)

                st.markdown(
                    f'<div style="font-size:0.52rem;font-weight:800;letter-spacing:0.18em;'
                    f'color:{C_CIAN};margin:10px 0 4px;">COMENTARIOS</div>',
                    unsafe_allow_html=True)
                _cmt_prev = str(_r.get("COMENTARIOS","") or "") \
                            if "COMENTARIOS" in df_raw.columns else ""
                if _cmt_prev:
                    _lines_html = "".join(
                        f'<div style="font-size:0.70rem;color:#94A3B8;padding:5px 0;'
                        f'border-bottom:1px solid #1E293B;">{ln}</div>'
                        for ln in _cmt_prev.strip().split("\n") if ln.strip()
                    )
                    st.markdown(
                        f'<div style="background:rgba(6,11,21,0.6);border:1px solid #1E293B;'
                        f'border-radius:8px;padding:8px 12px;margin-bottom:8px;max-height:130px;'
                        f'overflow-y:auto;">{_lines_html}</div>',
                        unsafe_allow_html=True)
                _new_cmt = st.text_area("Nuevo comentario",
                                        placeholder="Escribe un comentario y guarda...",
                                        height=68, label_visibility="collapsed",
                                        key="det_cmt_input")

                _sb1, _sb2 = st.columns([3, 2])
                with _sb1:
                    _guardar_d = st.form_submit_button(
                        "💾 Guardar cambios", type="primary",
                        use_container_width=True, disabled=not _token())
                with _sb2:
                    _cerrar_d = st.form_submit_button(
                        "✕ Cerrar", use_container_width=True)

            if _guardar_d and _token():
                _dfc2 = df_raw.copy()
                if "COMENTARIOS" not in _dfc2.columns:
                    _dfc2["COMENTARIOS"] = ""
                _mk2 = _dfc2["ID"] == _dt_id
                _dfc2.loc[_mk2, "TAREA"]            = _new_tarea
                _dfc2.loc[_mk2, "ESTADO"]           = _new_est
                _dfc2.loc[_mk2, "PRIORIDAD"]        = _new_prio
                _dfc2.loc[_mk2, "TIPO"]             = _new_tipo
                _dfc2.loc[_mk2, "PROYECTO"]         = _new_proj
                _dfc2.loc[_mk2, "AREA"]             = _new_area
                _dfc2.loc[_mk2, "CATEGORIA"]        = _new_cat
                _dfc2.loc[_mk2, "TERCERO"]          = _new_ter
                _dfc2.loc[_mk2, "FECHA_COMPROMISO"] = (
                    pd.Timestamp(_new_fc) if _new_fc else pd.NaT)
                _dfc2.loc[_mk2, "IMPACTO"]          = _new_imp
                _dfc2.loc[_mk2, "URGENCIA"]         = _new_urg
                _dfc2.loc[_mk2, "ESFUERZO_HRS"]     = _new_esf
                _dfc2.loc[_mk2, "HORAS_REALES"]     = _new_hr
                _dfc2.loc[_mk2, "NOTAS"]            = _new_notas
                if _new_est == "Completada":
                    _cierre_prev = _dfc2.loc[_mk2, "FECHA_CIERRE"].values[0]
                    if pd.isna(_cierre_prev):
                        _dfc2.loc[_mk2, "FECHA_CIERRE"] = pd.Timestamp(HOY)
                # Historial automático: detectar cambios de campos clave
                _auto_lines = []
                _ts_auto = pd.Timestamp.now().strftime("%d/%m/%Y %H:%M")
                _track = [
                    ("ESTADO",    str(_r.get("ESTADO","") or ""),    _new_est),
                    ("PRIORIDAD", str(_r.get("PRIORIDAD","") or ""), _new_prio),
                    ("PROYECTO",  str(_r.get("PROYECTO","") or ""),  _new_proj),
                    ("ESFUERZO_HRS", str(float(_r.get("ESFUERZO_HRS",0) or 0)),
                                     str(_new_esf)),
                ]
                _fc_old_s = pd.Timestamp(_fc_raw).strftime("%d/%m/%Y") if pd.notna(_fc_raw) else "—"
                _fc_new_s = pd.Timestamp(_new_fc).strftime("%d/%m/%Y") if _new_fc else "—"
                if _fc_old_s != _fc_new_s:
                    _auto_lines.append(f"FECHA_COMPROMISO: '{_fc_old_s}' → '{_fc_new_s}'")
                for _fn, _fv_old, _fv_new in _track:
                    if str(_fv_old).strip() != str(_fv_new).strip():
                        _auto_lines.append(f"{_fn}: '{_fv_old}' → '{_fv_new}'")
                # Construir bloque de prepend (auto + manual)
                _prepend = []
                if _auto_lines:
                    _prepend.append(f"[{_ts_auto} 🔄] " + " | ".join(_auto_lines))
                if _new_cmt.strip():
                    _prepend.append(f"[{_ts_auto}] {_new_cmt.strip()}")
                if _prepend:
                    _p_cmt = str(_dfc2.loc[_mk2, "COMENTARIOS"].values[0] or "")
                    _blq   = "\n".join(_prepend)
                    _dfc2.loc[_mk2, "COMENTARIOS"] = (
                        _blq + ("\n" + _p_cmt if _p_cmt else "")).strip()
                with st.spinner("Guardando..."):
                    _ok_d, _ms_d = guardar_github(_dfc2)
                if _ok_d:
                    st.session_state["detalle_id"] = None
                    st.toast("✅ Guardado"); st.rerun()
                else:
                    st.error(_ms_d)

            if _cerrar_d:
                st.session_state["detalle_id"] = None
                st.rerun()

            # ── Eliminar tarea ─────────────────────────────────────────────
            _del_sp, _del_btn_col = st.columns([4, 2])
            with _del_btn_col:
                if st.button("🗑 Eliminar tarea", key="btn_del_task", use_container_width=True):
                    st.session_state["confirm_delete"] = _dt_id
                    st.rerun()
            if st.session_state.get("confirm_delete") == _dt_id:
                st.warning("⚠️ ¿Eliminar esta tarea permanentemente? Esta acción no se puede deshacer.")
                _dc1, _dc2 = st.columns(2)
                with _dc1:
                    if st.button("⛔ Sí, eliminar", type="primary", key="btn_del_ok", use_container_width=True):
                        if _token():
                            _dfc_del = df_raw[df_raw["ID"] != _dt_id].copy()
                            with st.spinner("Eliminando..."):
                                _ok_del, _ms_del = guardar_github(_dfc_del)
                            if _ok_del:
                                st.session_state["detalle_id"] = None
                                st.session_state.pop("confirm_delete", None)
                                st.toast("✅ Tarea eliminada"); st.rerun()
                            else:
                                st.error(_ms_del)
                with _dc2:
                    if st.button("✕ Cancelar eliminación", key="btn_del_cancel", use_container_width=True):
                        st.session_state.pop("confirm_delete", None); st.rerun()

        st.markdown('<div style="height:4px;"></div>', unsafe_allow_html=True)

    # ══════════════════════════════════════════════════════════════════════════
    # VISTA CALENDARIO — mismo formato que kanban, drag cross-day
    # ══════════════════════════════════════════════════════════════════════════
    if st.session_state["cc_view"] == "calendario":
        seccion("📅",
                f"CALENDARIO SEMANAL · {HOY.strftime('%d %b')} – "
                f"{(HOY+timedelta(days=6)).strftime('%d %b %Y')}",
                C_CIAN)
        _cal_ac = ac[
            ac["FECHA_COMPROMISO"].notna() &
            (ac["FECHA_COMPROMISO"] >= HOY_TS) &
            (ac["FECHA_COMPROMISO"] <= HOY_TS + pd.Timedelta(days=6))
        ]
        _sf_ac = ac[ac["FECHA_COMPROMISO"].isna()].sort_values("ORDEN")
        # ── Construir columna "Sin fecha" ──────────────────────────────────
        _SF_ST_ABB = {"Pendiente":"⏸","En Proceso":"▶","Esperando Terceros":"🕐","Completada":"✓","Cancelada":"✗"}
        _sf_th = ""
        for _, _t in _sf_ac.iterrows():
            _ti2 = int(_t["ID"])
            _est2 = str(_t.get("ESTADO","Pendiente"))
            _dn2 = _est2 == "Completada"
            _p2  = str(_t.get("PRIORIDAD","Media"))
            _pc2 = PRIO_CLR.get(_p2, C_GRIS)
            _n2  = (str(_t.get("TAREA",""))
                    .replace("&","&amp;").replace("<","&lt;").replace(">","&gt;"))
            _pj2 = str(_t.get("PROYECTO","") or "")[:18]
            _ch2 = "checked" if _dn2 else ""
            _s2  = "text-decoration:line-through;opacity:0.38;" if _dn2 else ""
            _esf2 = float(_t.get("ESFUERZO_HRS", 0) or 0)
            _esf2_s = f"{int(_esf2)}h" if _esf2 == int(_esf2) else f"{_esf2:.1f}h"
            _meta2 = (f'<div class="kc-meta">'
                      + (f'<span class="kc-proj">{_pj2}</span>' if _pj2 else '')
                      + (f'<span class="kc-hrs">{_esf2_s}</span>' if _esf2 > 0 else '')
                      + '</div>') if (_pj2 or _esf2 > 0) else ''
            _st_abb2 = _SF_ST_ABB.get(_est2, "⏸")
            _sf_th += (
                f'<div class="kc" data-id="{_ti2}">'
                f'<div class="kc-top">'
                f'<span class="kc-dot" style="background:{_pc2};"></span>'
                f'<span class="kc-ico" style="color:{_pc2};">{PRIO_ICO.get(_p2,"")}</span>'
                f'<button class="kc-st" data-id="{_ti2}" data-state="{_est2}" title="Cambiar estado">{_st_abb2}</button>'
                f'<button class="kc-cp" data-id="{_ti2}" title="Duplicar tarea">⊕</button>'
                f'<button class="kc-del" data-id="{_ti2}" title="Eliminar tarea">🗑</button>'
                f'<input class="kc-chk" type="checkbox" data-id="{_ti2}" {_ch2}>'
                f'</div>'
                f'<div class="kc-nm" style="{_s2}">{_n2}</div>'
                + _meta2 + f'</div>'
            )
        _sf_hc  = _TC["meta_clr"]
        _sf_hrs = float(_sf_ac["ESFUERZO_HRS"].fillna(0).sum())
        _sf_hrs_s = (f"{int(_sf_hrs)}h" if _sf_hrs == int(_sf_hrs) else f"{_sf_hrs:.1f}h") if _sf_hrs > 0 else ""
        _sf_bs = (f'color:{_sf_hc};font-size:0.65rem;font-weight:800;background:{_sf_hc}18;'
                  f'border:1px solid {_sf_hc}44;border-radius:5px;padding:2px 8px;')
        _sf_cnt_b = f'<span style="{_sf_bs}">{len(_sf_ac)} tareas</span>'
        _sf_hrs_b = f'<span style="{_sf_bs}">⏳ {_sf_hrs_s}</span>' if _sf_hrs_s else ""
        _sf_col_h = (
            f'<div class="kk" style="border-top:3px solid {_sf_hc}55;opacity:0.78;">'
            f'<div class="kk-top"><div>'
            f'<div class="kk-hdr" style="color:{_sf_hc};">📥 Sin fecha</div>'
            f'</div></div>'
            f'<div class="kk-cnt" style="display:flex;gap:5px;flex-wrap:wrap;">'
            f'{_sf_cnt_b}{_sf_hrs_b}</div>'
            f'<div class="dz" data-date="" data-field="date" data-group="" id="dz-sf">{_sf_th}</div></div>'
        )
        _DIAS_ES = ["Lun","Mar","Mié","Jue","Vie","Sáb","Dom"]
        _dcols_h = _sf_col_h
        _mx_t    = max(len(_sf_ac), 0)
        for _i in range(7):
            _d     = HOY + timedelta(days=_i)
            _d_str = _d.strftime("%Y-%m-%d")
            _td    = _cal_ac[_cal_ac["FECHA_COMPROMISO"].dt.date == _d].sort_values("ORDEN")
            if len(_td) > _mx_t: _mx_t = len(_td)
            _es_h  = (_i == 0)
            _hc    = f"rgb({_ABR})" if _es_h else _TC["meta_clr"]
            _th    = ""
            for _, _t in _td.iterrows():
                _ti2 = int(_t["ID"])
                _est2 = str(_t.get("ESTADO","Pendiente"))
                _dn2 = _est2 == "Completada"
                _p2  = str(_t.get("PRIORIDAD","Media"))
                _pc2 = PRIO_CLR.get(_p2, C_GRIS)
                _n2  = (str(_t.get("TAREA",""))
                        .replace("&","&amp;").replace("<","&lt;").replace(">","&gt;"))
                _pj2 = str(_t.get("PROYECTO","") or "")[:18]
                _ch2 = "checked" if _dn2 else ""
                _s2  = "text-decoration:line-through;opacity:0.38;" if _dn2 else ""
                _esf2 = float(_t.get("ESFUERZO_HRS", 0) or 0)
                _esf2_s = f"{int(_esf2)}h" if _esf2 == int(_esf2) else f"{_esf2:.1f}h"
                _meta2 = (f'<div class="kc-meta">'
                          + (f'<span class="kc-proj">{_pj2}</span>' if _pj2 else '')
                          + (f'<span class="kc-hrs">{_esf2_s}</span>' if _esf2 > 0 else '')
                          + '</div>') if (_pj2 or _esf2 > 0) else ''
                _st_abb2 = _SF_ST_ABB.get(_est2, "⏸")
                _th += (
                    f'<div class="kc" data-id="{_ti2}">'
                    f'<div class="kc-top">'
                    f'<span class="kc-dot" style="background:{_pc2};"></span>'
                    f'<span class="kc-ico" style="color:{_pc2};">{PRIO_ICO.get(_p2,"")}</span>'
                    f'<button class="kc-st" data-id="{_ti2}" data-state="{_est2}" title="Cambiar estado">{_st_abb2}</button>'
                    f'<button class="kc-cp" data-id="{_ti2}" title="Duplicar tarea">⊕</button>'
                    f'<button class="kc-del" data-id="{_ti2}" title="Eliminar tarea">🗑</button>'
                    f'<input class="kc-chk" type="checkbox" data-id="{_ti2}" {_ch2}>'
                    f'</div>'
                    f'<div class="kc-nm" style="{_s2}">{_n2}</div>'
                    + _meta2
                    + f'</div>'
                )
            _add_btn = f'<button class="kk-add" data-field="date" data-group="{_d_str}" title="Agregar tarea">+</button>' if _token() else ''
            _dhrs   = float(_td["ESFUERZO_HRS"].fillna(0).sum())
            _dhrs_s = (f"{int(_dhrs)}h" if _dhrs == int(_dhrs) else f"{_dhrs:.1f}h") if _dhrs > 0 else ""
            _dbs    = (f'color:{_hc};font-size:0.65rem;font-weight:800;background:{_hc}18;'
                       f'border:1px solid {_hc}44;border-radius:5px;padding:2px 8px;')
            _dcnt_b = f'<span style="{_dbs}">{len(_td)} tareas</span>'
            _dhrs_b = f'<span style="{_dbs}">⏳ {_dhrs_s}</span>' if _dhrs_s else ""
            _dcols_h += (
                f'<div class="kk{"  kk-today" if _es_h else ""}" '
                f'style="border-top:3px solid {_hc}{"80" if not _es_h else ""};">'
                f'<div class="kk-top">'
                f'<div>'
                f'<div class="kk-hdr" style="color:{_hc};">{_DIAS_ES[_d.weekday()]}</div>'
                f'<div class="kk-num">{_d.day}</div>'
                f'</div>'
                f'{_add_btn}</div>'
                f'<div class="kk-cnt" style="display:flex;gap:5px;flex-wrap:wrap;">'
                f'{_dcnt_b}{_dhrs_b}</div>'
                f'<div class="dz" data-date="{_d_str}" data-field="date" '
                f'data-group="{_d_str}" id="dz{_d_str}">{_th}</div></div>'
            )
        _cal_h   = max(400, 240 + _mx_t * 90)
        _tok_ok2 = "true" if _token() else "false"
        components.html(f"""<!DOCTYPE html><html><head><meta charset="utf-8">
<script src="https://cdn.jsdelivr.net/npm/sortablejs@1.15.2/Sortable.min.js"></script>
<style>
*{{box-sizing:border-box;margin:0;padding:0;}}
html,body{{height:100%;background:transparent;
           font-family:-apple-system,'Segoe UI',sans-serif;}}
body{{overflow-x:auto;overflow-y:auto;}}
.cal{{display:flex;flex-direction:row;flex-wrap:nowrap;gap:8px;padding:2px 2px 10px;
     align-items:flex-start;}}
::-webkit-scrollbar{{height:5px;width:5px;}}
::-webkit-scrollbar-track{{background:transparent;}}
::-webkit-scrollbar-thumb{{background:rgba({_ABR},0.25);border-radius:3px;}}
::-webkit-scrollbar-thumb:hover{{background:rgba({_ABR},0.45);}}
.kk{{flex:0 0 220px;width:220px;background:{_TC['card1']};
     border:1px solid {_TC['border']};border-radius:14px;padding:11px 9px;}}
.kk-top{{display:flex;align-items:flex-start;justify-content:space-between;margin-bottom:4px;}}
.kk-hdr{{font-size:0.54rem;font-weight:800;letter-spacing:0.10em;}}
.kk-num{{font-size:1.30rem;font-weight:900;line-height:1;color:{_TC['nm_clr']};}}
.kk-cnt{{font-size:0.48rem;color:{_TC['txt_dim']};font-weight:700;margin-bottom:8px;}}
.kk-add{{background:rgba({_ABR},0.06);border:1px solid rgba({_ABR},0.16);
         border-radius:6px;color:{_TC['meta_clr']};cursor:pointer;font-size:1rem;
         font-weight:700;line-height:1;padding:3px 7px;transition:all .15s;}}
.kk-add:hover{{background:rgba({_ABR},0.16);border-color:rgba({_ABR},0.40);
              color:rgb({_ABR});}}
.dz{{min-height:48px;border-radius:9px;border:2px dashed transparent;
    padding:3px;transition:all .18s;}}
.dz.ov{{border-color:rgba({_ABR},.38)!important;background:rgba({_ABR},.04);}}
.kc{{background:{_TC['card2']};border:1px solid {_TC['border']};border-radius:9px;
     padding:7px;margin-bottom:5px;cursor:grab;}}
.kc:active{{cursor:grabbing;}}
.kc-top{{display:flex;align-items:center;gap:5px;margin-bottom:4px;}}
.kc-dot{{width:7px;height:7px;border-radius:50%;flex-shrink:0;}}
.kc-ico{{font-size:0.58rem;flex:1;}}
.kc-chk{{width:13px;height:13px;accent-color:#23D160;cursor:pointer;flex-shrink:0;}}
.kc-nm{{font-size:0.68rem;font-weight:600;color:{_TC['nm_clr']};
        line-height:1.28;word-break:break-word;}}
.kc-meta{{display:flex;gap:4px;margin-top:4px;}}
.kc-proj{{font-size:0.52rem;color:{_TC['meta_clr']};background:rgba({_ABR},0.08);
          border-radius:4px;padding:1px 4px;}}
.kc-hrs{{font-size:0.52rem;color:rgb({_ABR});font-weight:700;
         background:rgba({_ABR},0.13);border-radius:4px;padding:1px 5px;}}
.kc-cp{{background:transparent;border:none;color:#334155;cursor:pointer;
        font-size:0.70rem;line-height:1;padding:1px 3px;flex-shrink:0;transition:color .15s;}}
.kc-cp:hover{{color:rgb({_ABR});}}
.kc-st{{background:rgba(100,116,139,0.12);border:1px solid rgba(100,116,139,0.22);
        border-radius:4px;color:#64748B;cursor:pointer;font-size:0.50rem;font-weight:800;
        line-height:1;padding:2px 4px;flex-shrink:0;transition:all .15s;}}
.kc-st:hover{{background:rgba({_ABR},0.14);border-color:rgba({_ABR},0.30);color:rgb({_ABR});}}
.kc-del{{background:transparent;border:none;color:#334155;cursor:pointer;
         font-size:0.62rem;line-height:1;padding:1px 3px;flex-shrink:0;transition:color .15s;}}
.kc-del:hover{{color:#FF4757;}}
.kc-del.conf{{color:#FF4757!important;font-weight:700;background:rgba(255,71,87,0.10);
              border-radius:3px;}}
.sortable-ghost{{opacity:.20;transform:scale(.96);}}
.sortable-chosen{{box-shadow:0 4px 18px rgba({_ABR},.24);}}
</style></head><body>
<div class="cal">{_dcols_h}</div>
<script>
var HT={_tok_ok2};
function nfy(v){{
  try{{
    var el=window.parent.document.querySelector('input[aria-label="TASK_ACTION"]');
    if(!el)return;
    var s=Object.getOwnPropertyDescriptor(window.parent.HTMLInputElement.prototype,'value').set;
    s.call(el,v);
    el.dispatchEvent(new Event('input',{{bubbles:true}}));
    el.dispatchEvent(new KeyboardEvent('keydown',{{bubbles:true,cancelable:true,key:'Enter',keyCode:13}}));
    el.dispatchEvent(new KeyboardEvent('keypress',{{bubbles:true,cancelable:true,key:'Enter',keyCode:13}}));
    el.dispatchEvent(new KeyboardEvent('keyup',{{bubbles:true,cancelable:true,key:'Enter',keyCode:13}}));
  }}catch(e){{console.error(e);}}
}}
document.querySelectorAll('.dz').forEach(function(dz){{
  Sortable.create(dz,{{
    group:'cal',animation:150,ghostClass:'sortable-ghost',chosenClass:'sortable-chosen',
    onAdd:function(evt){{if(HT)nfy('date:'+evt.item.dataset.id+':'+evt.to.dataset.date);}},
    onEnd:function(evt){{
      if(HT&&evt.from===evt.to){{
        var ids=Array.from(evt.from.children).map(function(c){{return c.dataset.id;}}).join(',');
        nfy('reorder:'+evt.from.dataset.date+':'+ids);
      }}
    }},
    onOver:function(evt){{evt.to.classList.add('ov');}},
    onLeave:function(evt){{evt.from.classList.remove('ov');}}
  }});
}});
document.querySelectorAll('.kk-add').forEach(function(btn){{
  btn.addEventListener('click',function(e){{
    e.stopPropagation();
    if(HT)nfy('newform:date:'+this.dataset.group);
  }});
}});
document.querySelectorAll('.kc-cp').forEach(function(btn){{
  btn.addEventListener('click',function(e){{
    e.stopPropagation();
    nfy('copy:'+this.dataset.id+':');
  }});
}});
var _ST_CYCLE=['Pendiente','En Proceso','Esperando Terceros','Completada'];
var _ST_ABB={{'Pendiente':'⏸','En Proceso':'▶','Esperando Terceros':'🕐','Completada':'✓','Cancelada':'✗'}};
document.querySelectorAll('.kc-st').forEach(function(btn){{
  btn.addEventListener('click',function(e){{
    e.stopPropagation();
    var cur=this.dataset.state||'Pendiente';
    var idx=_ST_CYCLE.indexOf(cur);
    var nxt=_ST_CYCLE[(idx+1)%_ST_CYCLE.length];
    this.dataset.state=nxt;
    this.textContent=_ST_ABB[nxt]||nxt;
    nfy('field:'+this.dataset.id+':ESTADO:'+nxt);
  }});
}});
var _del_timer={{}};
document.querySelectorAll('.kc-del').forEach(function(btn){{
  btn.addEventListener('click',function(e){{
    e.stopPropagation();
    var id=this.dataset.id;
    if(this.classList.contains('conf')){{
      clearTimeout(_del_timer[id]);
      delete _del_timer[id];
      this.classList.remove('conf');
      this.textContent='🗑';
      if(HT)nfy('delete:'+id+':');
    }}else{{
      var self=this;
      this.classList.add('conf');
      this.textContent='¿OK?';
      _del_timer[id]=setTimeout(function(){{
        self.classList.remove('conf');
        self.textContent='🗑';
      }},2200);
    }}
  }});
}});
if(HT){{
  document.querySelectorAll('.kc-chk').forEach(function(chk){{
    chk.addEventListener('change',function(){{
      var ns=this.checked?'Completada':'Pendiente';
      var nm=this.closest('.kc').querySelector('.kc-nm');
      if(this.checked){{nm.style.textDecoration='line-through';nm.style.opacity='.38';}}
      else{{nm.style.textDecoration='';nm.style.opacity='';}}
      nfy('complete:'+this.dataset.id+':'+ns);
    }});
  }});
}}
var _dbc={{}};
document.querySelectorAll('.kc').forEach(function(c){{
  c.addEventListener('click',function(e){{
    if(e.target.closest('.kc-chk')||e.target.closest('.kc-cp')||
       e.target.closest('.kc-st')||e.target.closest('.kc-del'))return;
    var tid=this.dataset.id,now=Date.now();
    if(_dbc[tid]&&(now-_dbc[tid])<380){{_dbc[tid]=0;nfy('open:'+tid+':');}}
    else{{_dbc[tid]=now;}}
  }});
}});
</script></body></html>""", height=_cal_h, scrolling=True)
        st.stop()

    # ══════════════════════════════════════════════════════════════════════════
    # VISTAS KANBAN (Estado / Prioridad / Área / Categoría / Vencidas)
    # ══════════════════════════════════════════════════════════════════════════
    _cv = st.session_state["cc_view"]
    if _cv in ("estado","prioridad","area","categoria","vencidas"):

        _KB_CFG = {
            "estado": {
                "field":  "ESTADO",
                "groups": ["Pendiente","En Proceso","Esperando Terceros","Completada"],
                "colors": {"Pendiente":C_GRIS,"En Proceso":C_CIAN,
                           "Esperando Terceros":C_ALERTA,"Completada":C_OK},
                "icons":  {"Pendiente":"⏳","En Proceso":"🔄",
                           "Esperando Terceros":"🕐","Completada":"✅"},
            },
            "prioridad": {
                "field":  "PRIORIDAD",
                "groups": ["Crítica","Alta","Media","Baja"],
                "colors": {"Crítica":C_CRITICO,"Alta":C_ALERTA,"Media":C_CIAN,"Baja":C_GRIS},
                "icons":  {"Crítica":"🔴","Alta":"🟡","Media":"🔵","Baja":"⚫"},
            },
            "area": {
                "field":  "AREA",
                "groups": ["Trabajo","Personal"],
                "colors": {"Trabajo":C_CIAN,"Personal":C_MORADO},
                "icons":  {"Trabajo":"💼","Personal":"🏠"},
            },
            "categoria": {
                "field":  "CATEGORIA",
                "groups": ["Planificación","Contratos","Compras","Reportes",
                           "IA y Automatización","Gestión Corporativa",
                           "Reuniones","Salud","Personal"],
                "colors": {g:CAT_CLR.get(g,C_CIAN) for g in
                           ["Planificación","Contratos","Compras","Reportes",
                            "IA y Automatización","Gestión Corporativa",
                            "Reuniones","Salud","Personal"]},
                "icons":  {"Planificación":"📐","Contratos":"📝","Compras":"🛒",
                           "Reportes":"📊","IA y Automatización":"🤖",
                           "Gestión Corporativa":"🏢","Reuniones":"📅",
                           "Salud":"❤️","Personal":"🏠"},
            },
            "vencidas": {
                "field":  "PRIORIDAD",
                "groups": ["Crítica","Alta","Media","Baja"],
                "colors": {"Crítica":C_CRITICO,"Alta":C_ALERTA,"Media":C_CIAN,"Baja":C_GRIS},
                "icons":  {"Crítica":"🔴","Alta":"🟡","Media":"🔵","Baja":"⚫"},
            },
        }

        _cfg     = _KB_CFG[_cv]
        _kfield  = _cfg["field"]
        _kgroups = _cfg["groups"]
        _kcolors = _cfg["colors"]
        _kicons  = _cfg["icons"]

        # Dataset: vencidas = solo atrasadas activas; estado = sin Cancelada; resto = activas
        if _cv == "vencidas":
            _kac = ac[ac["FECHA_COMPROMISO"].notna() &
                      (ac["FECHA_COMPROMISO"] < HOY_TS)]
        elif _cv == "estado":
            _kac = df_raw[~df_raw["ESTADO"].isin(["Cancelada"])]
        else:
            _kac = ac

        _sect_ico = {"estado":"🔷","prioridad":"🎯","area":"📁",
                     "categoria":"🗂","vencidas":"🚨"}
        seccion(_sect_ico.get(_cv,"📋"), f"KANBAN · {_cv.upper()}", C_CIAN)

        # ── Construir columnas (ancho fijo uniforme, scroll horizontal) ──────────
        _kn   = len(_kgroups)
        _mx_t = 0
        _dcols_h = ""
        _tok_add = _token()
        for _gval in _kgroups:
            _gdf = _kac[_kac[_kfield].fillna("") == _gval].sort_values("ORDEN")
            if len(_gdf) > _mx_t: _mx_t = len(_gdf)
            _gc2  = _kcolors.get(_gval, C_GRIS)
            _gi   = _kicons.get(_gval, "")
            _sg   = _gval.replace(" ","-").replace("/","-")
            _th2  = ""
            _KB_ST_ABB = {"Pendiente":"⏸","En Proceso":"▶","Esperando Terceros":"🕐","Completada":"✓","Cancelada":"✗"}
            for _, _t in _gdf.iterrows():
                _ti2 = int(_t["ID"])
                _est3 = str(_t.get("ESTADO","Pendiente"))
                _dn2 = _est3 == "Completada"
                _nm2 = (str(_t.get("TAREA",""))
                        .replace("&","&amp;").replace("<","&lt;").replace(">","&gt;"))
                _pj2 = str(_t.get("PROYECTO","") or "")[:20]
                _fc2 = _t.get("FECHA_COMPROMISO")
                _fs2 = pd.Timestamp(_fc2).strftime("%d/%m") if pd.notna(_fc2) else ""
                _pr2 = str(_t.get("PRIORIDAD","Media"))
                _pc3 = PRIO_CLR.get(_pr2, C_GRIS)
                _ch3 = "checked" if _dn2 else ""
                _sy3 = "text-decoration:line-through;opacity:0.38;" if _dn2 else ""
                _esf3 = float(_t.get("ESFUERZO_HRS", 0) or 0)
                _esf3_s = f"{int(_esf3)}h" if _esf3 == int(_esf3) else f"{_esf3:.1f}h"
                _delay = ""
                if _cv == "vencidas" and pd.notna(_fc2):
                    _days = (HOY_TS - pd.Timestamp(_fc2)).days
                    _delay = f'<span class="kc-delay">−{_days}d</span>'
                _st_abb3 = _KB_ST_ABB.get(_est3, "⏸")
                _th2 += (
                    f'<div class="kc" data-id="{_ti2}">'
                    f'<div class="kc-top">'
                    f'<span class="kc-dot" style="background:{_pc3};"></span>'
                    f'<span class="kc-ico" style="color:{_pc3};">{PRIO_ICO.get(_pr2,"")}</span>'
                    f'<button class="kc-st" data-id="{_ti2}" data-state="{_est3}" title="Cambiar estado">{_st_abb3}</button>'
                    f'<button class="kc-cp" data-id="{_ti2}" title="Duplicar tarea">⊕</button>'
                    f'<button class="kc-del" data-id="{_ti2}" title="Eliminar tarea">🗑</button>'
                    f'<input class="kc-chk" type="checkbox" data-id="{_ti2}" {_ch3}>'
                    f'</div>'
                    f'<div class="kc-nm" style="{_sy3}">{_nm2}</div>'
                    f'<div class="kc-meta">'
                    + (f'<span class="kc-proj">{_pj2}</span>' if _pj2 else '')
                    + (f'<span class="kc-fc">{_fs2}</span>' if _fs2 else '')
                    + (f'<span class="kc-hrs">{_esf3_s}</span>' if _esf3 > 0 else '')
                    + _delay
                    + f'</div></div>'
                )
            _add_btn = (f'<button class="kk-add" data-field="{_kfield}" '
                        f'data-group="{_gval}" title="Nueva tarea">+</button>'
                        if _tok_add and _cv != "vencidas" else "")
            _ghrs   = float(_gdf["ESFUERZO_HRS"].fillna(0).sum())
            _ghrs_s = (f"{int(_ghrs)}h" if _ghrs == int(_ghrs) else f"{_ghrs:.1f}h") if _ghrs > 0 else ""
            _bs = (f'color:{_gc2};font-size:0.65rem;font-weight:800;background:{_gc2}18;'
                   f'border:1px solid {_gc2}44;border-radius:5px;padding:2px 8px;')
            _cnt_badge  = f'<span style="{_bs}">{len(_gdf)} tareas</span>'
            _ghrs_badge = f'<span style="{_bs}">⏳ {_ghrs_s}</span>' if _ghrs_s else ""
            _dcols_h += (
                f'<div class="kk" style="border-top:3px solid {_gc2}55;">'
                f'<div class="kk-top">'
                f'<div class="kk-hdr" style="color:{_gc2};">{_gi} {_gval}</div>'
                f'{_add_btn}</div>'
                f'<div class="kk-cnt" style="display:flex;gap:5px;flex-wrap:wrap;">'
                f'{_cnt_badge}{_ghrs_badge}</div>'
                f'<div class="dz" data-field="{_kfield}" data-group="{_gval}" '
                f'id="dz-{_sg}">{_th2}</div></div>'
            )

        _COL_W = 220          # ancho fijo de cada columna (px) — mismo en todas las vistas
        _kh    = max(440, 260 + _mx_t * 92)
        _tokk  = "true" if _token() else "false"
        components.html(f"""<!DOCTYPE html><html><head><meta charset="utf-8">
<script src="https://cdn.jsdelivr.net/npm/sortablejs@1.15.2/Sortable.min.js"></script>
<style>
*{{box-sizing:border-box;margin:0;padding:0;}}
html,body{{height:100%;background:transparent;
           font-family:-apple-system,'Segoe UI',sans-serif;}}
body{{overflow-x:auto;overflow-y:auto;}}
/* ── fila única de columnas con scroll horizontal ── */
.kb{{display:flex;flex-direction:row;flex-wrap:nowrap;
    gap:10px;padding:2px 2px 10px;align-items:flex-start;}}
.kk{{flex:0 0 {_COL_W}px;width:{_COL_W}px;
     background:{_TC['card1']};border:1px solid {_TC['border']};
     border-radius:14px;padding:11px 9px;}}
.kk-top{{display:flex;align-items:center;justify-content:space-between;margin-bottom:2px;}}
.kk-hdr{{font-size:0.68rem;font-weight:900;letter-spacing:0.05em;}}
.kk-add{{background:rgba({_ABR},0.06);border:1px solid rgba({_ABR},0.16);
         border-radius:6px;color:{_TC['meta_clr']};cursor:pointer;
         font-size:1rem;font-weight:700;line-height:1;padding:2px 7px;
         transition:all .15s;flex-shrink:0;}}
.kk-add:hover{{background:rgba({_ABR},0.16);border-color:rgba({_ABR},0.40);
              color:rgb({_ABR});}}
.kk-cnt{{font-size:0.48rem;color:{_TC['txt_dim']};font-weight:700;margin-bottom:9px;}}
.dz{{min-height:54px;border-radius:9px;border:2px dashed transparent;
    padding:3px;transition:all .18s;}}
.dz.ov{{border-color:rgba({_ABR},.38)!important;background:rgba({_ABR},.04);}}
.kc{{background:{_TC['card2']};border:1px solid {_TC['border']};
     border-radius:9px;padding:8px;margin-bottom:5px;cursor:grab;}}
.kc:active{{cursor:grabbing;}}
.kc-top{{display:flex;align-items:center;gap:5px;margin-bottom:5px;}}
.kc-dot{{width:7px;height:7px;border-radius:50%;flex-shrink:0;}}
.kc-ico{{font-size:0.58rem;flex:1;}}
.kc-chk{{width:13px;height:13px;accent-color:#23D160;cursor:pointer;flex-shrink:0;}}
.kc-nm{{font-size:0.70rem;font-weight:600;color:{_TC['nm_clr']};
        line-height:1.3;word-break:break-word;}}
.kc-meta{{display:flex;gap:4px;margin-top:4px;flex-wrap:wrap;}}
.kc-proj{{font-size:0.52rem;color:{_TC['meta_clr']};background:rgba({_ABR},0.08);
          border-radius:4px;padding:1px 4px;}}
.kc-fc{{font-size:0.52rem;color:{_TC['meta_clr']};}}
.kc-hrs{{font-size:0.52rem;color:rgb({_ABR});font-weight:700;
         background:rgba({_ABR},0.13);border-radius:4px;padding:1px 5px;}}
.kc-delay{{font-size:0.52rem;color:#FF4757;font-weight:700;
           background:rgba(255,71,87,0.10);border-radius:4px;padding:1px 4px;}}
.kc-cp{{background:transparent;border:none;color:#334155;cursor:pointer;
        font-size:0.70rem;line-height:1;padding:1px 3px;flex-shrink:0;transition:color .15s;}}
.kc-cp:hover{{color:rgb({_ABR});}}
.kc-st{{background:rgba(100,116,139,0.12);border:1px solid rgba(100,116,139,0.22);
        border-radius:4px;color:#64748B;cursor:pointer;font-size:0.50rem;font-weight:800;
        line-height:1;padding:2px 4px;flex-shrink:0;transition:all .15s;}}
.kc-st:hover{{background:rgba({_ABR},0.14);border-color:rgba({_ABR},0.30);color:rgb({_ABR});}}
.kc-del{{background:transparent;border:none;color:#334155;cursor:pointer;
         font-size:0.62rem;line-height:1;padding:1px 3px;flex-shrink:0;transition:color .15s;}}
.kc-del:hover{{color:#FF4757;}}
.kc-del.conf{{color:#FF4757!important;font-weight:700;background:rgba(255,71,87,0.10);
              border-radius:3px;}}
.sortable-ghost{{opacity:.18;transform:scale(.94);}}
.sortable-chosen{{box-shadow:0 4px 18px rgba({_ABR},.26);}}
/* scrollbar delgado */
::-webkit-scrollbar{{height:5px;width:5px;}}
::-webkit-scrollbar-track{{background:transparent;}}
::-webkit-scrollbar-thumb{{background:rgba({_ABR},0.25);border-radius:3px;}}
::-webkit-scrollbar-thumb:hover{{background:rgba({_ABR},0.45);}}
</style></head><body>
<div class="kb">{_dcols_h}</div>
<script>
var HT={_tokk};
function nfy(v){{
  try{{
    var el=window.parent.document.querySelector('input[aria-label="TASK_ACTION"]');
    if(!el)return;
    var s=Object.getOwnPropertyDescriptor(window.parent.HTMLInputElement.prototype,'value').set;
    s.call(el,v);
    el.dispatchEvent(new Event('input',{{bubbles:true}}));
    el.dispatchEvent(new KeyboardEvent('keydown',{{bubbles:true,cancelable:true,key:'Enter',keyCode:13}}));
    el.dispatchEvent(new KeyboardEvent('keypress',{{bubbles:true,cancelable:true,key:'Enter',keyCode:13}}));
    el.dispatchEvent(new KeyboardEvent('keyup',{{bubbles:true,cancelable:true,key:'Enter',keyCode:13}}));
  }}catch(e){{console.error(e);}}
}}
document.querySelectorAll('.dz').forEach(function(dz){{
  Sortable.create(dz,{{
    group:'kb',animation:150,ghostClass:'sortable-ghost',chosenClass:'sortable-chosen',
    onAdd:function(evt){{
      if(HT)nfy('field:'+evt.item.dataset.id+':'+evt.to.dataset.field+':'+evt.to.dataset.group);
    }},
    onEnd:function(evt){{
      if(HT&&evt.from===evt.to){{
        var ids=Array.from(evt.from.children).map(function(c){{return c.dataset.id;}}).join(',');
        nfy('reorder:'+evt.from.dataset.group+':'+ids);
      }}
    }},
    onOver:function(evt){{evt.to.classList.add('ov');}},
    onLeave:function(evt){{evt.from.classList.remove('ov');}}
  }});
}});
document.querySelectorAll('.kk-add').forEach(function(btn){{
  btn.addEventListener('click',function(e){{
    e.stopPropagation();
    if(HT)nfy('newform:'+this.dataset.field+':'+this.dataset.group);
  }});
}});
document.querySelectorAll('.kc-cp').forEach(function(btn){{
  btn.addEventListener('click',function(e){{
    e.stopPropagation();
    nfy('copy:'+this.dataset.id+':');
  }});
}});
var _ST_CYCLE=['Pendiente','En Proceso','Esperando Terceros','Completada'];
var _ST_ABB={{'Pendiente':'⏸','En Proceso':'▶','Esperando Terceros':'🕐','Completada':'✓','Cancelada':'✗'}};
document.querySelectorAll('.kc-st').forEach(function(btn){{
  btn.addEventListener('click',function(e){{
    e.stopPropagation();
    var cur=this.dataset.state||'Pendiente';
    var idx=_ST_CYCLE.indexOf(cur);
    var nxt=_ST_CYCLE[(idx+1)%_ST_CYCLE.length];
    this.dataset.state=nxt;
    this.textContent=_ST_ABB[nxt]||nxt;
    nfy('field:'+this.dataset.id+':ESTADO:'+nxt);
  }});
}});
var _del_timer={{}};
document.querySelectorAll('.kc-del').forEach(function(btn){{
  btn.addEventListener('click',function(e){{
    e.stopPropagation();
    var id=this.dataset.id;
    if(this.classList.contains('conf')){{
      clearTimeout(_del_timer[id]);
      delete _del_timer[id];
      this.classList.remove('conf');
      this.textContent='🗑';
      if(HT)nfy('delete:'+id+':');
    }}else{{
      var self=this;
      this.classList.add('conf');
      this.textContent='¿OK?';
      _del_timer[id]=setTimeout(function(){{
        self.classList.remove('conf');
        self.textContent='🗑';
      }},2200);
    }}
  }});
}});
if(HT){{
  document.querySelectorAll('.kc-chk').forEach(function(chk){{
    chk.addEventListener('change',function(){{
      var ns=this.checked?'Completada':'Pendiente';
      var nm=this.closest('.kc').querySelector('.kc-nm');
      if(this.checked){{nm.style.textDecoration='line-through';nm.style.opacity='.38';}}
      else{{nm.style.textDecoration='';nm.style.opacity='';}}
      nfy('complete:'+this.dataset.id+':'+ns);
    }});
  }});
}}
var _dbt={{}};
document.querySelectorAll('.kc').forEach(function(c){{
  c.addEventListener('click',function(e){{
    if(e.target.closest('.kc-chk')||e.target.closest('.kc-cp')||
       e.target.closest('.kc-st')||e.target.closest('.kc-del'))return;
    var tid=this.dataset.id,now=Date.now();
    if(_dbt[tid]&&(now-_dbt[tid])<380){{_dbt[tid]=0;nfy('open:'+tid+':');}}
    else{{_dbt[tid]=now;}}
  }});
}});
</script></body></html>""", height=_kh, scrolling=True)
        st.stop()

    # fallback (vista no reconocida)
    col_ag, col_prox = st.columns([3, 2])

    # ── Agenda del Día ────────────────────────────────────────────────────────
    with col_ag:
        seccion("📌", "AGENDA DEL DÍA", C_CRITICO)
        agenda_base = ac[
            (ac["PRIORIDAD"].isin(["Crítica","Alta"])) |
            (ac["FECHA_COMPROMISO"].notna() &
             (ac["FECHA_COMPROMISO"] <= HOY_TS + pd.Timedelta(days=1)))
        ].head(14)
        agenda_ids_all = set(agenda_base["ID"].astype(int))
        id_to_row      = {int(r["ID"]): r for _, r in agenda_base.iterrows()}

        if not id_to_row:
            st.markdown('<div style="color:#334155;font-size:0.80rem;padding:20px;text-align:center;">'
                        '✅ No hay tareas urgentes pendientes</div>', unsafe_allow_html=True)
        else:
            # ── Reconciliar orden ────────────────────────────────────────────
            prev = st.session_state.get("agenda_order", [])
            ordered_ids = [i for i in prev if i in agenda_ids_all]
            ordered_ids += [i for i in id_to_row if i not in ordered_ids]
            st.session_state["agenda_order"] = ordered_ids

            # ── Procesar cambio de orden ──────────────────────────────────────
            if ag_order_raw.strip():
                try:
                    new_ids = [int(x) for x in ag_order_raw.split(",") if x.strip()]
                    if set(new_ids) == agenda_ids_all:
                        ordered_ids = new_ids
                        st.session_state["agenda_order"] = new_ids
                except Exception:
                    pass

            # ── Construir HTML de tarjetas ────────────────────────────────────
            _PICO = {"Crítica":"🔴","Alta":"🟡","Media":"🔵","Baja":"⚫"}
            cards_html = ""
            for task_id in ordered_ids:
                r = id_to_row.get(task_id)
                if r is None:
                    continue
                is_done = str(r.get("ESTADO","")) == "Completada"
                p   = str(r.get("PRIORIDAD","Media"))
                pc  = PRIO_CLR.get(p, C_GRIS)
                fc  = r.get("FECHA_COMPROMISO")
                sc_c, sd_t = semaforo(fc)
                venc_r = pd.notna(fc) and fc < HOY_TS and not is_done
                tipo  = str(r.get("TIPO","Tarea"))
                tc    = TIPO_CLR.get(tipo, C_CIAN)
                proj  = str(r.get("PROYECTO",""))
                if venc_r:
                    bg_c, brd_c = "rgba(255,71,87,0.05)", "rgba(255,71,87,0.22)"
                elif is_done:
                    bg_c, brd_c = "rgba(35,209,96,0.03)", "rgba(35,209,96,0.16)"
                else:
                    bg_c, brd_c = "rgba(13,21,38,0.85)", f"{pc}22"
                chk_attr = "checked" if is_done else ""
                nm_sty   = "text-decoration:line-through;opacity:0.42;color:#64748B;" if is_done else ""
                nm_text  = (str(r.get("TAREA",""))
                            .replace("&","&amp;").replace("<","&lt;").replace(">","&gt;"))
                pchip = (f'<span style="background:{pc}18;color:{pc};border:1px solid {pc}30;'
                         f'font-size:0.55rem;font-weight:700;padding:2px 7px;border-radius:20px;">'
                         f'{_PICO.get(p,"")} {p}</span>')
                tchip = (f'<span style="background:{tc}18;color:{tc};border:1px solid {tc}30;'
                         f'font-size:0.55rem;font-weight:700;padding:2px 7px;border-radius:20px;">'
                         f'{tipo}</span>')
                prchip = (f'<span style="background:#6366F118;color:#6366F1;border:1px solid #6366F130;'
                          f'font-size:0.55rem;font-weight:700;padding:2px 7px;border-radius:20px;">'
                          f'{proj}</span>') if proj else ""
                esf_ag = float(r.get("ESFUERZO_HRS", 0) or 0)
                esf_ag_s = f"{int(esf_ag)}h" if esf_ag == int(esf_ag) else f"{esf_ag:.1f}h"
                esfchip = (f'<span style="background:{C_CIAN}15;color:{C_CIAN};border:1px solid {C_CIAN}30;'
                           f'font-size:0.55rem;font-weight:700;padding:2px 7px;border-radius:20px;">'
                           f'⏳ {esf_ag_s}</span>') if esf_ag > 0 else ""
                cards_html += (
                    f'<div class="sc" data-id="{task_id}" '
                    f'style="background:{bg_c};border:1px solid {brd_c};">'
                    f'<label class="chkw"><input type="checkbox" data-id="{task_id}" {chk_attr}></label>'
                    f'<div class="body">'
                    f'<div class="chips">{pchip} {tchip} {prchip} {esfchip}</div>'
                    f'<div class="nm" style="{nm_sty}">{nm_text}</div>'
                    f'<div class="meta" style="color:{sc_c};">⏱ {sd_t}</div>'
                    f'</div>'
                    f'<div class="hdl" title="Arrastrar para reordenar">⠿</div>'
                    f'</div>'
                )

            # ── Componente con SortableJS ─────────────────────────────────────
            token_ok = "true" if _token() else "false"
            comp_h   = len(ordered_ids) * 92 + 20
            components.html(f"""<!DOCTYPE html><html><head><meta charset="utf-8">
<script src="https://cdn.jsdelivr.net/npm/sortablejs@1.15.2/Sortable.min.js"></script>
<style>
*{{box-sizing:border-box;margin:0;padding:0;}}
body{{background:transparent;font-family:-apple-system,'Segoe UI',sans-serif;overflow:hidden;}}
.list{{display:flex;flex-direction:column;gap:5px;padding:2px;}}
.sc{{border-radius:12px;padding:9px 11px;display:flex;align-items:flex-start;
     gap:9px;user-select:none;transition:box-shadow .15s;}}
.sortable-ghost{{opacity:.28;transform:scale(.98);}}
.sortable-chosen{{box-shadow:0 6px 28px rgba(56,189,248,.22);}}
.hdl{{color:#2A3A52;font-size:1.1rem;cursor:grab;padding-top:1px;flex-shrink:0;}}
.hdl:active{{cursor:grabbing;}}
.chkw{{padding-top:2px;flex-shrink:0;}}
input[type=checkbox]{{width:16px;height:16px;accent-color:#23D160;cursor:pointer;}}
.body{{flex:1;min-width:0;}}
.chips{{display:flex;flex-wrap:wrap;gap:3px;margin-bottom:5px;}}
.nm{{font-size:.79rem;font-weight:700;color:#E2E8F0;line-height:1.32;margin-bottom:3px;word-break:break-word;}}
.meta{{font-size:.60rem;}}
</style></head><body>
<div class="list" id="sl">{cards_html}</div>
<script>
var hasToken={token_ok};
function notify(type,val){{
  try{{
    var sel=type==='order'?'input[aria-label="DRAG_ORDER"]':'input[aria-label="TASK_ACTION"]';
    var el=window.parent.document.querySelector(sel);
    if(!el)return;
    var s=Object.getOwnPropertyDescriptor(window.parent.HTMLInputElement.prototype,'value').set;
    s.call(el,val);
    el.dispatchEvent(new Event('input',{{bubbles:true}}));
    el.dispatchEvent(new KeyboardEvent('keydown',{{bubbles:true,cancelable:true,key:'Enter',keyCode:13}}));
    el.dispatchEvent(new KeyboardEvent('keypress',{{bubbles:true,cancelable:true,key:'Enter',keyCode:13}}));
    el.dispatchEvent(new KeyboardEvent('keyup',{{bubbles:true,cancelable:true,key:'Enter',keyCode:13}}));
  }}catch(e){{console.error('notify:',e);}}
}}
Sortable.create(document.getElementById('sl'),{{
  animation:160,handle:'.hdl',ghostClass:'sortable-ghost',chosenClass:'sortable-chosen',
  onEnd:function(){{
    if(hasToken){{
      var ids=Array.from(document.getElementById('sl').children).map(function(c){{return c.dataset.id;}}).join(',');
      notify('action','reorder::'+ids);
    }}
  }}
}});
if(hasToken){{
  document.querySelectorAll('input[type=checkbox]').forEach(function(chk){{
    chk.addEventListener('change',function(){{
      var tid=this.dataset.id;
      var ns=this.checked?'Completada':'Pendiente';
      var nm=this.closest('.sc').querySelector('.nm');
      if(this.checked){{nm.style.textDecoration='line-through';nm.style.opacity='.42';nm.style.color='#64748B';}}
      else{{nm.style.textDecoration='';nm.style.opacity='';nm.style.color='';}}
      notify('action','complete:'+tid+':'+ns);
    }});
  }});
}}
var _dbt={{}};
document.querySelectorAll('.sc').forEach(function(c){{
  c.addEventListener('click',function(e){{
    if(e.target.closest('.hdl')||e.target.closest('.chkw'))return;
    var tid=this.dataset.id,now=Date.now();
    if(_dbt[tid]&&(now-_dbt[tid])<380){{
      _dbt[tid]=0;
      notify('action','open:'+tid+':');
    }}else{{_dbt[tid]=now;}}
  }});
}});
</script></body></html>""", height=comp_h, scrolling=False)

    # ── Próximos 7 días (lista compacta) ─────────────────────────────────────
    with col_prox:
        seccion("📅", "PRÓXIMOS 7 DÍAS", C_CIAN)
        prox7 = ac[
            ac["FECHA_COMPROMISO"].notna() &
            (ac["FECHA_COMPROMISO"] >= HOY_TS) &
            (ac["FECHA_COMPROMISO"] <= HOY_TS + pd.Timedelta(days=7))
        ].sort_values("FECHA_COMPROMISO")

        if prox7.empty:
            st.markdown('<div style="color:#334155;font-size:0.80rem;padding:20px;text-align:center;">'
                        '📭 Sin vencimientos los próximos 7 días</div>', unsafe_allow_html=True)
        else:
            items_h = ""
            last_day = None
            for _, r in prox7.iterrows():
                fc   = r["FECHA_COMPROMISO"]
                day  = fc.date()
                sc, sd = semaforo(fc)
                p    = str(r.get("PRIORIDAD","Media"))
                pc   = PRIO_CLR.get(p, C_GRIS)
                if day != last_day:
                    label = "HOY" if day == HOY else day.strftime("%a %d/%m").upper()
                    items_h += (f'<div style="font-size:0.52rem;font-weight:800;letter-spacing:0.12em;'
                                f'color:#334155;margin:10px 0 5px;">{label}</div>')
                    last_day = day
                items_h += (
                    f'<div style="background:rgba(56,189,248,0.03);border:1px solid rgba(56,189,248,0.08);'
                    f'border-radius:10px;padding:8px 12px;margin-bottom:5px;">'
                    f'<div style="font-size:0.75rem;font-weight:700;color:#E2E8F0;line-height:1.3;">'
                    f'{PRIO_ICO.get(p,"")} {r["TAREA"]}</div>'
                    f'<div style="font-size:0.58rem;color:{sc};margin-top:3px;">{sd} · '
                    f'{chip(str(r.get("PROYECTO","")), C_INDIGO)}</div>'
                    f'</div>'
                )
            st.markdown(f'<div class="scroll-box">{items_h}</div>', unsafe_allow_html=True)
        st.markdown('<div style="height:4px;"></div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# MÓDULO 2: DIAGNÓSTICO DE RIESGO
# ══════════════════════════════════════════════════════════════════════════════
# ══════════════════════════════════════════════════════════════════════════════
# MÓDULO 2: RESUMEN SEMANAL
# ══════════════════════════════════════════════════════════════════════════════
elif mod == "Resumen Semanal":
    ac    = _activas()
    comps = df_raw[df_raw["ESTADO"] == "Completada"]
    ini_sem = HOY - timedelta(days=HOY.weekday())
    fin_sem = ini_sem + timedelta(days=6)

    st.markdown(
        f'<div style="background:linear-gradient(100deg,rgba(56,189,248,0.07),transparent);'
        f'border:1px solid rgba(56,189,248,0.16);border-radius:20px;padding:16px 24px;margin-bottom:18px;">'
        f'<div style="font-size:0.58rem;font-weight:800;letter-spacing:0.24em;color:{C_CIAN};">PLANIFICACIÓN Y CONTROL · SEMANA {ini_sem.strftime("%V / %Y")}</div>'
        f'<div style="font-size:1.65rem;font-weight:900;color:#F8FAFC;">📆 RESUMEN SEMANAL</div>'
        f'<div style="font-size:0.70rem;color:{C_GRIS};margin-top:3px;">'
        f'{ini_sem.strftime("%d %b")} — {fin_sem.strftime("%d %b %Y")}</div>'
        f'</div>', unsafe_allow_html=True)

    # ── KPIs de la semana ────────────────────────────────────────────────────
    comp_hoy = comps[comps["FECHA_CIERRE"].notna() & (comps["FECHA_CIERRE"].dt.date == HOY)]
    en_proc  = ac[ac["ESTADO"] == "En Proceso"]
    venc_act = ac[ac["FECHA_COMPROMISO"].notna() & (ac["FECHA_COMPROMISO"] < HOY_TS)]
    crit_act = ac[ac["PRIORIDAD"] == "Crítica"]
    prox7    = ac[ac["FECHA_COMPROMISO"].notna() &
                  (ac["FECHA_COMPROMISO"] >= HOY_TS) &
                  (ac["FECHA_COMPROMISO"] <= HOY_TS + pd.Timedelta(days=7))]
    hh_prox7 = float(prox7["ESFUERZO_HRS"].fillna(0).sum())
    comp_sem_r = comps[comps["FECHA_CIERRE"].notna() & (comps["FECHA_CIERRE"] >= pd.Timestamp(ini_sem))]
    venc_sem_r = ac[ac["FECHA_COMPROMISO"].notna() &
                    (ac["FECHA_COMPROMISO"] >= HOY_TS - pd.Timedelta(days=7)) &
                    (ac["FECHA_COMPROMISO"] < HOY_TS)]
    _tot_s = len(comp_sem_r) + len(venc_sem_r)
    tasa_rs = round(len(comp_sem_r) / max(_tot_s, 1) * 100)
    tasa_clr_rs = C_OK if tasa_rs >= 70 else C_ALERTA if tasa_rs >= 40 else C_CRITICO

    _k1, _k2, _k3, _k4, _k5, _k6 = st.columns(6)
    with _k1: st.markdown(kpi("COMPLETADAS HOY",   len(comp_hoy),
                               "tareas cerradas", color=C_OK),          unsafe_allow_html=True)
    with _k2: st.markdown(kpi("EN PROCESO",         len(en_proc),
                               "en ejecución", color=C_CIAN),           unsafe_allow_html=True)
    with _k3: st.markdown(kpi("VENCIDAS ACTIVAS",   len(venc_act),
                               "fuera de plazo",
                               color=C_CRITICO if len(venc_act) else C_OK), unsafe_allow_html=True)
    with _k4: st.markdown(kpi("CRÍTICAS ACTIVAS",   len(crit_act),
                               "requieren atención",
                               color=C_CRITICO if len(crit_act) else C_GRIS), unsafe_allow_html=True)
    with _k5:
        _hh_s = f"{hh_prox7:.0f}h" if hh_prox7 > 0 else "0h"
        st.markdown(kpi("HH PRÓX. 7 DÍAS",   _hh_s,
                         "esfuerzo comprometido", color=C_ALERTA),      unsafe_allow_html=True)
    with _k6: st.markdown(kpi("TASA SEMANAL",       f"{tasa_rs}%",
                               "cumplimiento", color=tasa_clr_rs),      unsafe_allow_html=True)

    st.markdown('<div style="height:8px;"></div>', unsafe_allow_html=True)

    # ── Agenda de la semana ──────────────────────────────────────────────────
    seccion("📅", "AGENDA DE LA SEMANA", C_CIAN)
    _DIAS_RS = ["Lun","Mar","Mié","Jue","Vie","Sáb","Dom"]
    _agenda_h = f'<div style="display:grid;grid-template-columns:repeat(7,1fr);gap:6px;margin-bottom:4px;">'
    for _ii in range(7):
        _dd = ini_sem + timedelta(days=_ii)
        _es_hoy_rs = (_dd == HOY)
        _td_rs = ac[ac["FECHA_COMPROMISO"].notna() & (ac["FECHA_COMPROMISO"].dt.date == _dd)]
        _td_rs = _td_rs.sort_values("PRIO_ORD") if not _td_rs.empty else _td_rs
        _hc_rs = f"rgb({_ABR})" if _es_hoy_rs else "#475569"
        _bg_rs = f"rgba({_ABR},0.06)" if _es_hoy_rs else "rgba(13,21,38,0.45)"
        _brd_rs = f"rgba({_ABR},0.30)" if _es_hoy_rs else "rgba(255,255,255,0.05)"
        _hh_d  = float(_td_rs["ESFUERZO_HRS"].fillna(0).sum())
        _hh_d_s = f" · {_hh_d:.0f}h" if _hh_d > 0 else ""
        _cards_rs = ""
        for _, _t in _td_rs.iterrows():
            _p_rs = str(_t.get("PRIORIDAD","Media"))
            _pc_rs = PRIO_CLR.get(_p_rs, C_GRIS)
            _nm_rs = str(_t["TAREA"])[:26]
            _est_rs = str(_t.get("ESTADO",""))
            _done_rs = _est_rs == "Completada"
            _opacity_rs = "opacity:0.40;" if _done_rs else ""
            _strike_rs  = "text-decoration:line-through;" if _done_rs else ""
            _esf_rs = float(_t.get("ESFUERZO_HRS",0) or 0)
            _esf_rs_s = f"{_esf_rs:.0f}h" if _esf_rs > 0 else ""
            _cards_rs += (
                f'<div style="background:rgba(13,21,38,0.75);border:1px solid {_pc_rs}22;'
                f'border-left:3px solid {_pc_rs};border-radius:6px;padding:5px 7px;margin-bottom:3px;{_opacity_rs}">'
                f'<div style="font-size:0.62rem;color:#E2E8F0;line-height:1.28;{_strike_rs}">{_nm_rs}</div>'
                + (f'<span style="font-size:0.50rem;color:rgb({_ABR});font-weight:700;">{_esf_rs_s}</span>' if _esf_rs_s else "")
                + f'</div>'
            )
        _agenda_h += (
            f'<div style="background:{_bg_rs};border:1px solid {_brd_rs};'
            f'border-top:3px solid {_hc_rs}{"80" if not _es_hoy_rs else ""};'
            f'border-radius:10px;padding:10px 8px;min-height:90px;">'
            f'<div style="font-size:0.52rem;font-weight:800;color:{_hc_rs};letter-spacing:0.08em;">'
            f'{_DIAS_RS[_ii].upper()}</div>'
            f'<div style="font-size:1.10rem;font-weight:900;color:{_hc_rs if _es_hoy_rs else "#94A3B8"};'
            f'margin-bottom:5px;">{_dd.day}<span style="font-size:0.55rem;color:{C_CIAN};">{_hh_d_s}</span></div>'
            + _cards_rs
            + f'</div>'
        )
    _agenda_h += '</div>'
    st.markdown(_agenda_h, unsafe_allow_html=True)

    st.markdown('<div style="height:10px;"></div>', unsafe_allow_html=True)

    _col_risg, _col_time = st.columns([1, 1])

    # ── Top 5 riesgos ────────────────────────────────────────────────────────
    with _col_risg:
        seccion("⚠️", "TOP RIESGOS ACTIVOS", C_CRITICO)
        _ac_r2 = ac.copy()
        _ac_r2["RIESGO"] = _ac_r2["URGENCIA"].astype(float) * _ac_r2["IMPACTO"].astype(float)
        _top5 = _ac_r2.nlargest(5, "RIESGO")
        for _, _rr in _top5.iterrows():
            _pr = str(_rr.get("PRIORIDAD","Media"))
            _pc_r = PRIO_CLR.get(_pr, C_GRIS)
            _rsk = float(_rr["RIESGO"])
            _clr_r = C_CRITICO if _rsk >= 20 else C_ALERTA if _rsk >= 12 else C_CIAN
            _urg_r = float(_rr.get("URGENCIA",3))
            _imp_r = float(_rr.get("IMPACTO",3))
            _action_r = ("🔴 ACTUAR AHORA" if _urg_r >= 3 and _imp_r >= 3 else
                         "🟡 PLANIFICAR"   if _imp_r >= 3 else
                         "🔵 DELEGAR"      if _urg_r >= 3 else
                         "⚪ MONITOREAR")
            _fc_r = _rr.get("FECHA_COMPROMISO")
            _dias_rr = (pd.Timestamp(_fc_r).date() - HOY).days if pd.notna(_fc_r) else None
            _dias_rs_s = ("Vence hoy" if _dias_rr == 0 else
                          f"Vence en {_dias_rr}d" if _dias_rr and _dias_rr > 0 else
                          f"Vencida {abs(_dias_rr)}d" if _dias_rr is not None else "Sin fecha")
            _ter_r = str(_rr.get("TERCERO","") or "")
            st.markdown(
                f'<div style="background:rgba(13,21,38,0.85);border:1px solid {_clr_r}22;'
                f'border-left:3px solid {_clr_r};border-radius:10px;padding:10px 13px;margin-bottom:6px;">'
                f'<div style="display:flex;justify-content:space-between;align-items:flex-start;">'
                f'<div style="font-size:0.72rem;font-weight:700;color:#E2E8F0;flex:1;line-height:1.28;">'
                f'{PRIO_ICO.get(_pr,"")} {str(_rr["TAREA"])[:36]}</div>'
                f'<span style="font-size:0.50rem;font-weight:800;background:{_clr_r}18;'
                f'color:{_clr_r};border:1px solid {_clr_r}44;border-radius:5px;padding:2px 7px;'
                f'white-space:nowrap;margin-left:6px;">{_action_r}</span></div>'
                f'<div style="display:flex;gap:10px;margin-top:5px;">'
                f'<span style="font-size:0.58rem;color:{_clr_r};font-weight:700;">Score {_rsk:.0f}/25</span>'
                f'<span style="font-size:0.58rem;color:{C_GRIS};">📅 {_dias_rs_s}</span>'
                + (f'<span style="font-size:0.58rem;color:{C_ALERTA};">👤 {_ter_r}</span>' if _ter_r else "")
                + f'</div></div>',
                unsafe_allow_html=True)

    # ── Compromisos próximos 14 días ─────────────────────────────────────────
    with _col_time:
        seccion("📆", "COMPROMISOS PRÓXIMOS 14 DÍAS", C_ALERTA)
        _dias14_lbl, _cnt14, _hh14 = [], [], []
        _alertas14 = []
        for _jj in range(14):
            _dj = HOY + timedelta(days=_jj)
            _tdj = ac[ac["FECHA_COMPROMISO"].notna() & (ac["FECHA_COMPROMISO"].dt.date == _dj)]
            _dias14_lbl.append(_dj.strftime("%d/%m"))
            _cnt14.append(len(_tdj))
            _hh14.append(round(float(_tdj["ESFUERZO_HRS"].fillna(0).sum()), 1))
            if len(_tdj) >= 3:
                _alertas14.append(_dj.strftime("%d/%m"))
        ec({
            "backgroundColor": "transparent",
            "tooltip": {**_TT_AXIS,
                        "formatter": "[function(p){return p[0].axisValue+'<br/>'+p.map(function(s){return s.seriesName+': '+s.value;}).join('<br/>')}]"},
            "legend": {"bottom":"0%","textStyle":{"color":C_GRIS,"fontSize":9}},
            "grid": {"left":"40px","right":"50px","top":"14px","bottom":"44px"},
            "xAxis": {**_axis(),"type":"category","data":_dias14_lbl,
                      "axisLabel":{"color":"#64748B","fontSize":9,"rotate":35}},
            "yAxis": [
                {**_axis(),"type":"value","name":"Tareas","nameTextStyle":{"color":"#64748B","fontSize":8}},
                {**_axis(),"type":"value","name":"HH","nameTextStyle":{"color":"#64748B","fontSize":8},
                 "splitLine":{"show":False}},
            ],
            "series": [
                {"name":"Tareas","type":"bar","barMaxWidth":"28px","data":_cnt14,
                 "itemStyle":{"borderRadius":[4,4,0,0],
                              "color":{"type":"linear","x":0,"y":0,"x2":0,"y2":1,
                                       "colorStops":[{"offset":0,"color":"rgba(255,179,0,0.90)"},
                                                     {"offset":1,"color":"rgba(255,179,0,0.25)"}]}},
                 "label":{"show":True,"position":"top","color":"#94A3B8","fontSize":9}},
                {"name":"HH estimadas","type":"line","yAxisIndex":1,"data":_hh14,
                 "smooth":True,"symbol":"circle","symbolSize":4,
                 "lineStyle":{"color":f"rgb({_ABR})","width":2},
                 "itemStyle":{"color":f"rgb({_ABR})"}},
            ],
        }, height=250)
        if _alertas14:
            st.markdown(
                f'<div style="background:rgba(255,71,87,0.06);border:1px solid rgba(255,71,87,0.20);'
                f'border-radius:8px;padding:8px 12px;margin-top:6px;">'
                f'<span style="font-size:0.60rem;font-weight:800;color:{C_CRITICO};">⚠️ DÍAS CON ALTA CARGA (+3 tareas): </span>'
                f'<span style="font-size:0.60rem;color:{C_ALERTA};">{", ".join(_alertas14)}</span>'
                f'</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# MÓDULO 3: DIAGNÓSTICO DE RIESGO (mejorado)
# ══════════════════════════════════════════════════════════════════════════════
elif mod == "Diagnóstico de Riesgo":
    ac = _activas()

    st.markdown(
        f'<div style="background:linear-gradient(100deg,rgba(255,71,87,0.07),transparent);'
        f'border:1px solid rgba(255,71,87,0.16);border-radius:20px;padding:16px 24px;margin-bottom:18px;">'
        f'<div style="font-size:0.58rem;font-weight:800;letter-spacing:0.24em;color:{C_CRITICO};">ANÁLISIS DE RIESGO OPERACIONAL</div>'
        f'<div style="font-size:1.65rem;font-weight:900;color:#F8FAFC;">⚠️ DIAGNÓSTICO DE RIESGO</div>'
        f'<div style="font-size:0.70rem;color:{C_GRIS};margin-top:3px;">Matriz Urgencia × Impacto · Tamaño = Esfuerzo estimado</div>'
        f'</div>', unsafe_allow_html=True)

    # ── Filtros del diagnóstico ───────────────────────────────────────────────
    _df1, _df2, _df3 = st.columns([3, 2, 2])
    with _df1:
        _projs_diag = sorted([p for p in ac["PROYECTO"].dropna().unique() if str(p).strip()])
        _f_proj_d = st.multiselect("Filtrar por proyecto", _projs_diag, key="diag_proj",
                                    placeholder="Todos los proyectos")
    with _df2:
        _areas_diag = ["Todas las áreas"] + sorted([a for a in ac["AREA"].dropna().unique() if str(a).strip()])
        _f_area_d = st.selectbox("Área", _areas_diag, key="diag_area", label_visibility="visible")
    with _df3:
        _f_est_d = st.selectbox("Estado", ["Todos"]+["Pendiente","En Proceso","Esperando Terceros"],
                                  key="diag_est", label_visibility="visible")
    ac_d = ac.copy()
    if _f_proj_d:
        ac_d = ac_d[ac_d["PROYECTO"].isin(_f_proj_d)]
    if _f_area_d != "Todas las áreas":
        ac_d = ac_d[ac_d["AREA"] == _f_area_d]
    if _f_est_d != "Todos":
        ac_d = ac_d[ac_d["ESTADO"] == _f_est_d]
    ac_d["RIESGO_"] = ac_d["URGENCIA"].astype(float) * ac_d["IMPACTO"].astype(float)
    _score_total = int(ac_d["RIESGO_"].sum())
    _score_max   = 25 * max(len(ac_d), 1)

    # KPI de riesgo total
    _sk1, _sk2, _sk3, _sk4 = st.columns(4)
    with _sk1: st.markdown(kpi("TAREAS EN ANÁLISIS", len(ac_d), color=C_CRITICO), unsafe_allow_html=True)
    with _sk2: st.markdown(kpi("SCORE RIESGO TOTAL", _score_total,
                                f"de {_score_max} máx.", color=C_CRITICO), unsafe_allow_html=True)
    with _sk3:
        _n_critico = int((ac_d["RIESGO_"] >= 20).sum())
        st.markdown(kpi("ZONA CRÍTICA (≥20)", _n_critico, "urgente+alto impacto",
                        color=C_CRITICO if _n_critico else C_GRIS), unsafe_allow_html=True)
    with _sk4:
        _n_alerta = int(((ac_d["RIESGO_"] >= 12) & (ac_d["RIESGO_"] < 20)).sum())
        st.markdown(kpi("ZONA ALERTA (12-19)", _n_alerta, "requiere seguimiento",
                        color=C_ALERTA if _n_alerta else C_GRIS), unsafe_allow_html=True)
    st.markdown('<div style="height:6px;"></div>', unsafe_allow_html=True)

    col_mat, col_info = st.columns([3, 1])

    with col_mat:
        series_data = {}
        for _, r in ac_d.iterrows():
            p   = str(r.get("PRIORIDAD","Media"))
            urg = float(r.get("URGENCIA", 3) or 3)
            imp = float(r.get("IMPACTO",  3) or 3)
            esf = float(r.get("ESFUERZO_HRS", 1) or 1)
            nm  = str(r["TAREA"])[:44]
            proj_d = str(r.get("PROYECTO","") or "")
            ter_d  = str(r.get("TERCERO","") or "")
            fc_d   = r.get("FECHA_COMPROMISO")
            fc_d_s = pd.Timestamp(fc_d).strftime("%d/%m/%Y") if pd.notna(fc_d) else ""
            dias_d = (pd.Timestamp(fc_d).date() - HOY).days if pd.notna(fc_d) else None
            dias_d_s = (f"Vence en {dias_d}d" if dias_d and dias_d > 0 else
                        "Vence hoy" if dias_d == 0 else
                        f"Vencida {abs(dias_d)}d" if dias_d is not None else "Sin fecha")
            if p not in series_data:
                series_data[p] = []
            series_data[p].append([urg, imp, esf, nm, proj_d, ter_d, fc_d_s, dias_d_s])

        series_list = []
        for p, pts in series_data.items():
            clr = PRIO_CLR.get(p, C_GRIS)
            series_list.append({
                "name": p, "type": "scatter",
                "data": pts,
                "symbolSize": "__SYM_FN__",
                "itemStyle": {
                    "color": clr,
                    "borderColor": f"{clr}",
                    "borderWidth": 1.5,
                    "opacity": 0.88,
                },
                "emphasis": {"scale": 1.15},
            })

        # Cuadrantes via markArea en una serie fantasma
        series_list.append({
            "name": "_bg", "type": "scatter", "data": [],
            "markArea": {
                "silent": True,
                "data": [
                    [{"coord":[0.5,3],"itemStyle":{"color":"rgba(255,179,0,0.04)"}},{"coord":[3,5.5]}],
                    [{"coord":[3,3],"itemStyle":{"color":"rgba(255,71,87,0.06)"}},  {"coord":[5.5,5.5]}],
                    [{"coord":[0.5,0.5],"itemStyle":{"color":"rgba(100,116,139,0.03)"}},{"coord":[3,3]}],
                    [{"coord":[3,0.5],"itemStyle":{"color":"rgba(56,189,248,0.03)"}},{"coord":[5.5,3]}],
                ],
            },
            "markLine": {
                "silent": True, "symbol": "none",
                "lineStyle": {"color":"rgba(255,255,255,0.08)","type":"dashed","width":1},
                "data": [{"xAxis":3},{"yAxis":3}],
            },
        })

        opt_mat = {
            "backgroundColor": "transparent",
            "tooltip": {"trigger":"item","formatter":"__TT_FN_DIAG__",
                        "backgroundColor":"rgba(6,11,21,0.97)",
                        "borderColor":"rgba(255,71,87,0.22)","borderRadius":12,
                        "textStyle":{"color":"#F1F5F9","fontSize":11}},
            "legend": {"bottom":"2%","textStyle":{"color":C_GRIS,"fontSize":10}},
            "grid": {"left":"60px","right":"60px","top":"50px","bottom":"60px"},
            "xAxis": {**_axis(),"type":"value","min":0.5,"max":5.5,
                      "name":"URGENCIA →","nameLocation":"end",
                      "nameTextStyle":{"color":"#475569","fontSize":9},
                      "splitNumber":4,
                      "axisLabel":{**_axis()["axisLabel"],
                                   "formatter":"{value}"}},
            "yAxis": {**_axis(),"type":"value","min":0.5,"max":5.5,
                      "name":"↑ IMPACTO","nameLocation":"end",
                      "nameTextStyle":{"color":"#475569","fontSize":9},
                      "splitNumber":4},
            "series": series_list,
            "graphic": [
                {"type":"text","left":"8%","top":"12%",
                 "style":{"text":"PLANIFICAR","fill":"rgba(255,179,0,0.22)",
                          "fontSize":9,"fontWeight":"bold","letterSpacing":3}},
                {"type":"text","right":"4%","top":"12%",
                 "style":{"text":"ACTUAR AHORA","fill":"rgba(255,71,87,0.30)",
                          "fontSize":9,"fontWeight":"bold","letterSpacing":2}},
                {"type":"text","left":"8%","bottom":"14%",
                 "style":{"text":"MONITOREAR","fill":"rgba(100,116,139,0.22)",
                          "fontSize":9,"fontWeight":"bold","letterSpacing":3}},
                {"type":"text","right":"4%","bottom":"14%",
                 "style":{"text":"DELEGAR","fill":"rgba(56,189,248,0.22)",
                          "fontSize":9,"fontWeight":"bold","letterSpacing":3}},
            ],
        }
        ec(opt_mat, height=440)

    with col_info:
        st.markdown('<div style="height:8px;"></div>', unsafe_allow_html=True)
        # Top riesgo (usa ac_d filtrado)
        ac2 = ac_d.copy()
        ac2["RIESGO"] = ac2["URGENCIA"].astype(float) * ac2["IMPACTO"].astype(float)
        top_r = ac2.nlargest(8, "RIESGO")
        cards = ""
        for _, r in top_r.iterrows():
            p  = str(r.get("PRIORIDAD","Media"))
            pc = PRIO_CLR.get(p, C_GRIS)
            rsk = float(r["RIESGO"])
            clr = C_CRITICO if rsk >= 20 else C_ALERTA if rsk >= 12 else C_CIAN
            cards += (
                f'<div style="background:rgba(13,21,38,0.80);border:1px solid {clr}22;'
                f'border-radius:10px;padding:9px 11px;margin-bottom:6px;">'
                f'<div style="font-size:0.68rem;font-weight:700;color:#E2E8F0;line-height:1.3;">'
                f'{PRIO_ICO.get(p,"")} {str(r["TAREA"])[:32]}</div>'
                f'<div style="font-size:0.56rem;color:{clr};margin-top:4px;font-weight:700;">'
                f'Riesgo {rsk:.0f}/25 · U:{r["URGENCIA"]:.0f} I:{r["IMPACTO"]:.0f}</div>'
                f'</div>'
            )
        st.markdown(
            f'<div style="font-size:0.50rem;font-weight:800;letter-spacing:0.14em;'
            f'color:#334155;text-transform:uppercase;margin-bottom:8px;">MAYOR RIESGO</div>'
            f'<div class="scroll-box" style="max-height:420px;">{cards}</div>',
            unsafe_allow_html=True)

    # Tabla completa de riesgo (sincronizada con filtros)
    seccion("📊", "TABLA DE RIESGO", C_CRITICO)
    ac_t = ac_d.copy()
    ac_t["RIESGO"] = ac_t["URGENCIA"].astype(float) * ac_t["IMPACTO"].astype(float)
    ac_t["ZONA"] = ac_t["RIESGO"].apply(
        lambda x: "🔴 CRÍTICO" if x >= 20 else "🟡 ALERTA" if x >= 12 else "🔵 NORMAL"
    )
    cols_show = [c for c in ["TAREA","ZONA","PRIORIDAD","ESTADO","URGENCIA","IMPACTO","ESFUERZO_HRS",
                              "TERCERO","FECHA_COMPROMISO"] if c in ac_t.columns]
    ac_tshow = ac_t.sort_values("RIESGO", ascending=False)[cols_show].copy()
    if "FECHA_COMPROMISO" in ac_tshow.columns:
        ac_tshow["FECHA_COMPROMISO"] = ac_tshow["FECHA_COMPROMISO"].dt.strftime("%d/%m/%Y").fillna("—")
    st.dataframe(ac_tshow.fillna("—"), use_container_width=True, hide_index=True,
                 height=min(500, 55 + len(ac_tshow)*36))

    # ── B7: Compromisos 30-60-90 días ────────────────────────────────────────
    seccion("📅", "COMPROMISOS 30 · 60 · 90 DÍAS", C_CRITICO)
    _ac_fc = _activas().copy()
    _ac_fc = _ac_fc[_ac_fc["FECHA_COMPROMISO"].notna()].copy()
    _ac_fc["_dias_hasta"] = (_ac_fc["FECHA_COMPROMISO"].dt.date.apply(
        lambda d: (d - HOY).days))
    _b7_30 = _ac_fc[(_ac_fc["_dias_hasta"] >= 0) & (_ac_fc["_dias_hasta"] <= 30)]
    _b7_60 = _ac_fc[(_ac_fc["_dias_hasta"] > 30) & (_ac_fc["_dias_hasta"] <= 60)]
    _b7_90 = _ac_fc[(_ac_fc["_dias_hasta"] > 60) & (_ac_fc["_dias_hasta"] <= 90)]

    _b7c1, _b7c2, _b7c3 = st.columns(3)
    _b7_bands = [
        (_b7c1, _b7_30, "0–30 días", C_CRITICO),
        (_b7c2, _b7_60, "31–60 días", C_ALERTA),
        (_b7c3, _b7_90, "61–90 días", C_CIAN),
    ]
    for _col_b7, _df_b7, _lbl_b7, _clr_b7 in _b7_bands:
        with _col_b7:
            st.markdown(kpi(f"COMPROMISOS {_lbl_b7}", len(_df_b7),
                            f"{_df_b7['ESFUERZO_HRS'].sum():.0f}h estimadas" if 'ESFUERZO_HRS' in _df_b7.columns else "",
                            color=_clr_b7), unsafe_allow_html=True)

    # Gráfico barras: cantidad de compromisos por semana (próximas 13 sem.)
    _b7_sem_lbl, _b7_sem_cnt, _b7_sem_hrs = [], [], []
    _b7_cur = HOY - timedelta(days=HOY.weekday())
    for _wi in range(13):
        _b7_ini = _b7_cur + timedelta(weeks=_wi)
        _b7_fin = _b7_ini + timedelta(days=6)
        _b7_mask = (
            (_ac_fc["FECHA_COMPROMISO"].dt.date >= _b7_ini) &
            (_ac_fc["FECHA_COMPROMISO"].dt.date <= _b7_fin)
        )
        _b7_sub = _ac_fc[_b7_mask]
        _b7_sem_lbl.append(f"S{_wi+1}\n{_b7_ini.strftime('%d/%m')}")
        _b7_sem_cnt.append(int(len(_b7_sub)))
        _b7_sem_hrs.append(round(float(_b7_sub["ESFUERZO_HRS"].sum()) if "ESFUERZO_HRS" in _b7_sub.columns else 0, 1))

    _b7_col_a, _b7_col_b = st.columns([3, 2])
    with _b7_col_a:
        seccion("📊", "COMPROMISOS SEMANALES — PRÓXIMAS 13 SEMANAS", C_CRITICO)
        _b7_clrs_cnt = [
            C_CRITICO if c >= 5 else C_ALERTA if c >= 3 else C_CIAN
            for c in _b7_sem_cnt
        ]
        ec({
            "backgroundColor": "transparent",
            "tooltip": {**_TT_AXIS},
            "legend": {"data": ["Tareas", "Horas est."], "bottom": "0%",
                       "textStyle": {"color": C_GRIS, "fontSize": 9}},
            "grid": {"left": "20px", "right": "20px", "top": "10px", "bottom": "34px", "containLabel": True},
            "xAxis": {**_axis(), "type": "category", "data": _b7_sem_lbl,
                      "axisLabel": {"color": "#94A3B8", "fontSize": 8}},
            "yAxis": [
                {**_axis(), "type": "value", "name": "Tareas",
                 "nameTextStyle": {"color": C_GRIS, "fontSize": 8}},
                {**_axis(), "type": "value", "name": "Horas", "position": "right",
                 "nameTextStyle": {"color": C_MORADO, "fontSize": 8}},
            ],
            "series": [
                {"name": "Tareas", "type": "bar", "barMaxWidth": "24px", "yAxisIndex": 0,
                 "data": [{"value": v, "itemStyle": {"color": _b7_clrs_cnt[i], "borderRadius": [4, 4, 0, 0]}}
                          for i, v in enumerate(_b7_sem_cnt)],
                 "label": {"show": True, "position": "top", "color": "#94A3B8", "fontSize": 8}},
                {"name": "Horas est.", "type": "line", "yAxisIndex": 1,
                 "data": _b7_sem_hrs,
                 "lineStyle": {"color": C_MORADO, "width": 2},
                 "itemStyle": {"color": C_MORADO},
                 "symbol": "circle", "symbolSize": 5, "smooth": True},
            ],
        }, height=260)

    with _b7_col_b:
        seccion("🗓", "DÍAS CRÍTICOS (3+ compromisos)", C_CRITICO)
        _b7_dias = {}
        for _, _rr in _ac_fc.iterrows():
            _fd = _rr["FECHA_COMPROMISO"].date()
            if 0 <= (_fd - HOY).days <= 90:
                _b7_dias.setdefault(_fd, []).append(str(_rr.get("TAREA", ""))[:24])
        _dias_crit = {d: t for d, t in sorted(_b7_dias.items()) if len(t) >= 3}
        if _dias_crit:
            for _dc, _tc in list(_dias_crit.items())[:8]:
                _dias_n = (_dc - HOY).days
                _dc_clr = C_CRITICO if _dias_n <= 7 else C_ALERTA if _dias_n <= 30 else C_CIAN
                st.markdown(
                    f'<div style="background:rgba(13,21,38,0.80);border:1px solid {_dc_clr}33;'
                    f'border-radius:10px;padding:9px 12px;margin-bottom:6px;">'
                    f'<div style="font-size:0.62rem;font-weight:800;color:{_dc_clr};">'
                    f'📅 {_dc.strftime("%d/%m/%Y")} · {_dias_n}d · {len(_tc)} tareas</div>'
                    f'<div style="font-size:0.56rem;color:#94A3B8;margin-top:3px;">'
                    + " · ".join(_tc[:4]) + ("…" if len(_tc) > 4 else "") +
                    f'</div></div>',
                    unsafe_allow_html=True)
        else:
            st.info("No hay días con 3+ compromisos en los próximos 90 días.", icon="✅")

# ══════════════════════════════════════════════════════════════════════════════
# MÓDULO 3: BANDEJA OPERACIONAL
# ══════════════════════════════════════════════════════════════════════════════
elif mod == "Bandeja Operacional":
    st.markdown(
        f'<div style="background:linear-gradient(100deg,rgba(56,189,248,0.06),transparent);'
        f'border:1px solid rgba(56,189,248,0.12);border-radius:20px;padding:16px 24px;margin-bottom:18px;">'
        f'<div style="font-size:0.58rem;font-weight:800;letter-spacing:0.24em;color:{C_CIAN};">VISTA DE TRABAJO</div>'
        f'<div style="font-size:1.65rem;font-weight:900;color:#F8FAFC;">📋 BANDEJA OPERACIONAL</div>'
        f'</div>', unsafe_allow_html=True)

    ac = _activas()
    # ── Fila 1: filtros categoriales ──────────────────────────────────────────
    f1,f2,f3,f4,f5 = st.columns(5)
    with f1:
        projs = ["Todos"] + sorted(df_raw["PROYECTO"].dropna().unique().tolist())
        fp = st.selectbox("Proyecto", projs, key="bo_proj")
    with f2:
        tipos = ["Todos"] + sorted(ac["TIPO"].dropna().unique().tolist()) if "TIPO" in ac.columns else ["Todos"]
        ft = st.selectbox("Tipo", tipos, key="bo_tipo")
    with f3:
        prios = ["Todos","Crítica","Alta","Media","Baja"]
        fpr = st.selectbox("Prioridad", prios, key="bo_prio")
    with f4:
        ests = ["Activas","Pendiente","En Proceso","Esperando Terceros","Completada","Todas"]
        fe = st.selectbox("Estado", ests, key="bo_est")
    with f5:
        buscar = st.text_input("🔍 Buscar", placeholder="Nombre de tarea...", key="bo_q",
                               label_visibility="visible")

    # ── Fila 2: filtro de fecha ────────────────────────────────────────────────
    fd1, fd2, fd3, _ = st.columns([2, 2, 2, 6])
    with fd1:
        usar_fecha = st.checkbox("Filtrar por fecha de vencimiento", key="bo_use_date")
    if usar_fecha:
        with fd2:
            date_desde = st.date_input("Desde", value=HOY, key="bo_desde")
        with fd3:
            date_hasta = st.date_input("Hasta", value=HOY + timedelta(days=30), key="bo_hasta")
    else:
        date_desde = date_hasta = None

    df_f = df_raw.copy()
    if fe == "Activas":
        df_f = df_f[~df_f["ESTADO"].isin(["Completada","Cancelada"])]
    elif fe != "Todas":
        df_f = df_f[df_f["ESTADO"] == fe]
    if fp != "Todos":      df_f = df_f[df_f["PROYECTO"] == fp]
    if ft != "Todos" and "TIPO" in df_f.columns: df_f = df_f[df_f["TIPO"] == ft]
    if fpr != "Todos":     df_f = df_f[df_f["PRIORIDAD"] == fpr]
    if buscar.strip():
        df_f = df_f[df_f["TAREA"].str.contains(buscar.strip(), case=False, na=False)]
    if usar_fecha and date_desde and date_hasta:
        df_f = df_f[
            df_f["FECHA_COMPROMISO"].notna() &
            (df_f["FECHA_COMPROMISO"] >= pd.Timestamp(date_desde)) &
            (df_f["FECHA_COMPROMISO"] <= pd.Timestamp(date_hasta))
        ]
    df_f = df_f.sort_values(["PRIO_ORD","FECHA_COMPROMISO"])

    # Stats rápidas
    n_total  = len(df_f)
    n_pend   = int((df_f["ESTADO"]=="Pendiente").sum())
    n_proc   = int((df_f["ESTADO"]=="En Proceso").sum())
    _mask_venc = (
        df_f["FECHA_COMPROMISO"].notna() &
        (df_f["FECHA_COMPROMISO"] < HOY_TS) &
        ~df_f["ESTADO"].isin(["Completada","Cancelada"])
    )
    n_venc = int(_mask_venc.sum())
    st.markdown(
        f'<div style="display:flex;gap:10px;margin:8px 0 10px;flex-wrap:wrap;">'
        f'<span style="background:rgba(56,189,248,0.08);border:1px solid rgba(56,189,248,0.18);'
        f'border-radius:10px;padding:5px 14px;font-size:0.70rem;font-weight:800;color:{C_CIAN};">'
        f'{n_total} TAREAS</span>'
        f'<span style="background:rgba(100,116,139,0.08);border:1px solid rgba(100,116,139,0.18);'
        f'border-radius:10px;padding:5px 14px;font-size:0.70rem;font-weight:800;color:{C_GRIS};">'
        f'⏳ {n_pend} pendientes</span>'
        f'<span style="background:rgba(56,189,248,0.08);border:1px solid rgba(56,189,248,0.16);'
        f'border-radius:10px;padding:5px 14px;font-size:0.70rem;font-weight:800;color:{C_CIAN};">'
        f'🔄 {n_proc} en proceso</span>'
        f'</div>',
        unsafe_allow_html=True)

    # ── Vencidas expandibles ───────────────────────────────────────────────────
    if n_venc > 0:
        with st.expander(f"🔴 {n_venc} TAREAS VENCIDAS — click para ver", expanded=False):
            df_venc = df_f[_mask_venc].sort_values("FECHA_COMPROMISO")
            venc_cols = [c for c in ["TAREA","PROYECTO","PRIORIDAD","ESTADO",
                                     "FECHA_COMPROMISO","TERCERO","NOTAS"]
                         if c in df_venc.columns]
            df_venc_show = df_venc[venc_cols].copy()
            if "FECHA_COMPROMISO" in df_venc_show.columns:
                df_venc_show["FECHA_COMPROMISO"] = (
                    df_venc_show["FECHA_COMPROMISO"].dt.strftime("%d/%m/%Y")
                )
            df_venc_show = df_venc_show.rename(columns={"FECHA_COMPROMISO": "VENCIÓ"})
            st.dataframe(
                df_venc_show.fillna("—"),
                use_container_width=True,
                hide_index=True,
                column_config={
                    "TAREA":    st.column_config.TextColumn("Tarea", width="large"),
                    "VENCIÓ":   st.column_config.TextColumn("Venció", width="small"),
                    "PRIORIDAD":st.column_config.TextColumn("Prioridad", width="small"),
                },
            )

    # ── Columnas y config del editor ──────────────────────────────────────────
    COLS_EDIT = [c for c in ["ID","TAREA","TIPO","PROYECTO","AREA","CATEGORIA",
                              "PRIORIDAD","ESTADO","IMPACTO","URGENCIA","ESFUERZO_HRS",
                              "TERCERO","FECHA_COMPROMISO","FECHA_CIERRE","NOTAS"]
                 if c in df_f.columns]
    df_edit = df_f[COLS_EDIT].copy()
    for dc in ["FECHA_COMPROMISO","FECHA_CIERRE"]:
        if dc in df_edit.columns:
            df_edit[dc] = pd.to_datetime(df_edit[dc], errors="coerce").dt.date

    # ── Controles de ordenamiento ──────────────────────────────────────────────
    _SORT_LBL = {
        "ID":"ID","TAREA":"Tarea","TIPO":"Tipo","PROYECTO":"Proyecto","AREA":"Área",
        "CATEGORIA":"Categoría","PRIORIDAD":"Prioridad","ESTADO":"Estado",
        "IMPACTO":"Impacto","URGENCIA":"Urgencia","ESFUERZO_HRS":"Horas",
        "TERCERO":"Tercero","FECHA_COMPROMISO":"Fecha de vencimiento",
        "FECHA_CIERRE":"Fecha de cierre",
    }
    _SORT_COLS = ["—"] + [c for c in COLS_EDIT if c != "NOTAS"]
    _srt1, _srt2 = st.columns([4, 3])
    with _srt1:
        _sort_by = st.selectbox(
            "Ordenar por",
            options=_SORT_COLS,
            format_func=lambda x: "— Sin ordenar —" if x == "—" else _SORT_LBL.get(x, x),
            key="bop_sort_col",
        )
    with _srt2:
        _sort_asc = st.radio(
            "Dirección",
            options=["↑ Ascendente", "↓ Descendente"],
            horizontal=True,
            key="bop_sort_dir",
        ) == "↑ Ascendente"
    if _sort_by != "—" and _sort_by in df_edit.columns:
        if _sort_by == "PRIORIDAD":
            _pord = {"Crítica":0,"Alta":1,"Media":2,"Baja":3}
            df_edit = df_edit.sort_values(
                "PRIORIDAD", ascending=_sort_asc,
                key=lambda s: s.map(_pord).fillna(9), na_position="last")
        elif _sort_by == "ESTADO":
            _eord = {"Pendiente":0,"En Proceso":1,"Esperando Terceros":2,"Completada":3,"Cancelada":4}
            df_edit = df_edit.sort_values(
                "ESTADO", ascending=_sort_asc,
                key=lambda s: s.map(_eord).fillna(9), na_position="last")
        else:
            df_edit = df_edit.sort_values(_sort_by, ascending=_sort_asc, na_position="last")

    _OPTS_EST  = ["Pendiente","En Proceso","Esperando Terceros","Completada","Cancelada"]
    _OPTS_PRIO = ["Crítica","Alta","Media","Baja"]
    _OPTS_TIPO = ["Tarea","Seguimiento","Compromiso","Reunión"]
    _OPTS_CAT  = ["Planificación","Contratos","Compras","Reportes","IA y Automatización",
                  "Gestión Corporativa","Reuniones","Salud","Personal"]
    _OPTS_AREA = ["Trabajo","Personal"]

    col_cfg = {
        "ID":              st.column_config.NumberColumn("ID", disabled=True, width="small"),
        "TAREA":           st.column_config.TextColumn("Tarea", width="large"),
        "TIPO":            st.column_config.SelectboxColumn("Tipo",      options=_OPTS_TIPO),
        "PROYECTO":        st.column_config.TextColumn("Proyecto"),
        "AREA":            st.column_config.SelectboxColumn("Área",      options=_OPTS_AREA),
        "CATEGORIA":       st.column_config.SelectboxColumn("Categoría", options=_OPTS_CAT),
        "PRIORIDAD":       st.column_config.SelectboxColumn("Prioridad", options=_OPTS_PRIO),
        "ESTADO":          st.column_config.SelectboxColumn("Estado",    options=_OPTS_EST),
        "IMPACTO":         st.column_config.NumberColumn("Impacto",  min_value=1, max_value=5, step=1),
        "URGENCIA":        st.column_config.NumberColumn("Urgencia",  min_value=1, max_value=5, step=1),
        "ESFUERZO_HRS":    st.column_config.NumberColumn("Horas",    min_value=0.0, step=0.5, format="%.1f"),
        "TERCERO":         st.column_config.TextColumn("Tercero"),
        "FECHA_COMPROMISO":st.column_config.DateColumn("Vence",  format="DD/MM/YYYY"),
        "FECHA_CIERRE":    st.column_config.DateColumn("Cierre", format="DD/MM/YYYY"),
        "NOTAS":           st.column_config.TextColumn("Notas", width="medium"),
    }

    edited = st.data_editor(
        df_edit,
        column_config=col_cfg,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        key="editor_bandeja",
        height=min(620, 55 + len(df_edit)*36 + 60),
    )

    # ── Barra de acción ────────────────────────────────────────────────────────
    sa1, sa2, sa3, sa4 = st.columns([2, 2, 2, 6])
    with sa1:
        guardar_btn = st.button("💾 Guardar en GitHub", type="primary",
                                use_container_width=True, disabled=not _token())
    with sa2:
        if st.button("↩ Descartar cambios", use_container_width=True):
            st.rerun()
    with sa3:
        _dl_buf = io.BytesIO()
        df_f.to_excel(_dl_buf, index=False, engine="openpyxl")
        _dl_buf.seek(0)
        st.download_button(
            "📥 Descargar Excel",
            data=_dl_buf.getvalue(),
            file_name=f"tareas_{HOY.strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    if guardar_btn:
        with st.spinner("Guardando en GitHub..."):
            df_merged = _merge_edits(df_raw, edited)
            ok, msg   = guardar_github(df_merged)
        if ok:
            st.success(f"✅ {msg}")
            st.rerun()
        else:
            st.error(msg)

    if not _token():
        st.info(
            "**Edición local desactivada.** Para guardar cambios desde la plataforma, "
            "configura `GITHUB_TOKEN` en Streamlit Cloud → Manage app → Secrets. "
            "O bien edita el Excel localmente y haz `git push`.",
            icon="ℹ️"
        )

    # ── Nueva tarea rápida ─────────────────────────────────────────────────────
    with st.expander("➕ Agregar tarea rápida"):
        with st.form("nueva_tarea_form", clear_on_submit=True):
            nt1, nt2, nt3 = st.columns(3)
            with nt1: nt_nombre = st.text_input("Nombre de la tarea *")
            with nt2: nt_proj   = st.text_input("Proyecto", value="GENERAL")
            with nt3: nt_prio   = st.selectbox("Prioridad", _OPTS_PRIO, index=2)
            nt4, nt5, nt6 = st.columns(3)
            with nt4: nt_tipo   = st.selectbox("Tipo", _OPTS_TIPO)
            with nt5: nt_cat    = st.selectbox("Categoría", _OPTS_CAT)
            with nt6: nt_fecha  = st.date_input("Fecha compromiso", value=None)
            nt_notas = st.text_area("Notas", height=60)
            submitted = st.form_submit_button("Agregar tarea", type="primary",
                                              use_container_width=True, disabled=not _token())
            if submitted:
                if not nt_nombre.strip():
                    st.error("El nombre es obligatorio.")
                else:
                    nueva = {
                        "ID":               int(df_raw["ID"].max() or 0) + 1,
                        "TAREA":            nt_nombre.strip(),
                        "DESCRIPCION":      "",
                        "TIPO":             nt_tipo,
                        "PROYECTO":         nt_proj.strip() or "GENERAL",
                        "AREA":             "Trabajo",
                        "CATEGORIA":        nt_cat,
                        "PRIORIDAD":        nt_prio,
                        "ESTADO":           "Pendiente",
                        "IMPACTO":          3,
                        "URGENCIA":         3,
                        "ESFUERZO_HRS":     1.0,
                        "TERCERO":          "",
                        "FECHA_CREACION":   pd.Timestamp(HOY),
                        "FECHA_COMPROMISO": pd.Timestamp(nt_fecha) if nt_fecha else pd.NaT,
                        "FECHA_CIERRE":     pd.NaT,
                        "NOTAS":            nt_notas,
                    }
                    df_nuevo = pd.concat(
                        [df_raw, pd.DataFrame([nueva])], ignore_index=True)
                    with st.spinner("Guardando..."):
                        ok, msg = guardar_github(df_nuevo)
                    if ok:
                        st.success(f"Tarea agregada · {msg}")
                        st.rerun()
                    else:
                        st.error(msg)

# ══════════════════════════════════════════════════════════════════════════════
# MÓDULO 4: SEGUIMIENTO DE TERCEROS
# ══════════════════════════════════════════════════════════════════════════════
elif mod == "Seguimiento de Terceros":
    st.markdown(
        f'<div style="background:linear-gradient(100deg,rgba(255,179,0,0.06),transparent);'
        f'border:1px solid rgba(255,179,0,0.14);border-radius:20px;padding:16px 24px;margin-bottom:18px;">'
        f'<div style="font-size:0.58rem;font-weight:800;letter-spacing:0.24em;color:{C_ALERTA};">CONTROL DE EXTERNOS</div>'
        f'<div style="font-size:1.65rem;font-weight:900;color:#F8FAFC;">👥 SEGUIMIENTO DE TERCEROS</div>'
        f'</div>', unsafe_allow_html=True)

    if df_ter.empty:
        st.warning("No se encontró la hoja TERCEROS en el Excel. Regenera el archivo con crear_excel_v2.py")
    else:
        pend_t = df_ter[df_ter["ESTADO"] != "Resuelto"] if "ESTADO" in df_ter.columns else df_ter
        dias_col = "DIAS_SIN_RESP"
        prom_dias = int(pend_t[dias_col].mean()) if dias_col in pend_t.columns and not pend_t.empty else 0
        crit_t = pend_t[pend_t.get("PRIORIDAD","") == "Alta"] if "PRIORIDAD" in pend_t.columns else pd.DataFrame()

        sin_contacto_7 = int((pend_t[dias_col] >= 7).sum()) if dias_col in pend_t.columns else 0
        sc7_clr = C_CRITICO if sin_contacto_7 > 0 else C_OK

        c1,c2,c3,c4,c5 = st.columns(5)
        with c1: st.markdown(kpi("TERCEROS ACTIVOS",  len(pend_t),
                                  color=C_ALERTA if len(pend_t) else C_GRIS),    unsafe_allow_html=True)
        with c2: st.markdown(kpi("DÍAS PROM. ESPERA", f"{prom_dias}d",
                                  "sin respuesta", color=C_ALERTA),              unsafe_allow_html=True)
        with c3: st.markdown(kpi("PRIORIDAD ALTA",    len(crit_t),
                                  color=C_CRITICO if len(crit_t) else C_GRIS),  unsafe_allow_html=True)
        with c4: st.markdown(kpi("RESUELTOS",
                                  len(df_ter)-len(pend_t), color=C_OK),         unsafe_allow_html=True)
        with c5: st.markdown(kpi("SIN CONTACTO +7d",  sin_contacto_7,
                                  "requieren seguimiento", color=sc7_clr),      unsafe_allow_html=True)

        if sin_contacto_7 > 0:
            _nom_sc7 = pend_t[pend_t[dias_col] >= 7]["NOMBRE"].tolist() if "NOMBRE" in pend_t.columns else []
            st.warning(
                f"⚠️ **{sin_contacto_7} tercero{'s' if sin_contacto_7>1 else ''} sin contacto hace 7+ días:** "
                + ", ".join(_nom_sc7[:5]) + ("…" if len(_nom_sc7) > 5 else ""),
                icon="⚠️",
            )

        st.markdown('<div style="height:6px;"></div>', unsafe_allow_html=True)
        seccion("📊", "ESTADO DE TERCEROS", C_ALERTA)

        # ── Tabla editable de terceros ────────────────────────────────────────
        _ter_edit_cols = [c for c in ["NOMBRE","ORGANIZACION","ROL","TEMA_PENDIENTE",
                                       "FECHA_ULTIMO_SEG","ESTADO","PRIORIDAD","NOTAS"]
                          if c in df_ter.columns]
        _ter_ed_df = df_ter[_ter_edit_cols].copy()
        for _fc in ["FECHA_ULTIMO_SEG"]:
            if _fc in _ter_ed_df.columns:
                _ter_ed_df[_fc] = pd.to_datetime(_ter_ed_df[_fc], errors="coerce")

        _col_cfg_ter = {}
        if "ESTADO" in _ter_ed_df.columns:
            _col_cfg_ter["ESTADO"] = st.column_config.SelectboxColumn(
                "Estado", options=["Pendiente", "En Seguimiento", "Resuelto", "Bloqueado"], required=True)
        if "PRIORIDAD" in _ter_ed_df.columns:
            _col_cfg_ter["PRIORIDAD"] = st.column_config.SelectboxColumn(
                "Prioridad", options=["Alta", "Media", "Baja"], required=True)
        if "FECHA_ULTIMO_SEG" in _ter_ed_df.columns:
            _col_cfg_ter["FECHA_ULTIMO_SEG"] = st.column_config.DateColumn("Último seguimiento", format="DD/MM/YYYY")

        _ter_edited = st.data_editor(
            _ter_ed_df, use_container_width=True, hide_index=True,
            column_config=_col_cfg_ter, num_rows="fixed",
            key="ter_data_editor",
        )

        # Botones de acción
        _tb1, _tb2, _tb3 = st.columns([2, 2, 6])
        with _tb1:
            if st.button("💾 Guardar cambios", key="btn_ter_save", use_container_width=True,
                         type="primary"):
                if _token():
                    _ter_upd = df_ter.copy()
                    for _ec in _ter_edit_cols:
                        if _ec in _ter_edited.columns:
                            _ter_upd[_ec] = _ter_edited[_ec].values
                    with st.spinner("Guardando terceros..."):
                        _ok_ter, _ms_ter = guardar_github(df_raw, df_ter_nuevo=_ter_upd)
                    if _ok_ter:
                        st.toast("✅ Terceros guardados")
                        st.rerun()
                    else:
                        st.error(_ms_ter)
                else:
                    st.error("Token no configurado")
        with _tb2:
            if st.button("📅 Registrar seg. hoy", key="btn_ter_hoy", use_container_width=True):
                _ter_sel = st.session_state.get("ter_data_editor", {}).get("selected_rows", [])
                if not _ter_sel:
                    st.info("Selecciona filas en la tabla para registrar seguimiento de hoy.", icon="ℹ️")
                else:
                    if _token():
                        _ter_upd2 = df_ter.copy()
                        for _ri in _ter_sel:
                            _ter_upd2.loc[_ri, "FECHA_ULTIMO_SEG"] = pd.Timestamp(HOY)
                        with st.spinner("Actualizando..."):
                            _ok2, _ms2 = guardar_github(df_raw, df_ter_nuevo=_ter_upd2)
                        if _ok2:
                            st.toast("✅ Seguimiento registrado")
                            st.rerun()
                        else:
                            st.error(_ms2)
        # DIAS_SIN_RESP se muestra aparte (calculado, no editable)
        if dias_col in df_ter.columns:
            _dias_show = df_ter[["NOMBRE", dias_col]].copy() if "NOMBRE" in df_ter.columns else df_ter[[dias_col]].copy()
            _dias_show = _dias_show.rename(columns={dias_col: "Días sin resp."})
            st.caption("Días sin respuesta (calculado automáticamente):")
            st.dataframe(_dias_show, use_container_width=True, hide_index=True, height=120)

        # Gráfico: días por tercero
        if dias_col in pend_t.columns and not pend_t.empty and "NOMBRE" in pend_t.columns:
            seccion("📈", "DÍAS SIN RESPUESTA POR TERCERO", C_ALERTA)
            names = pend_t["NOMBRE"].tolist()
            dias  = pend_t[dias_col].tolist()
            clrs  = [C_CRITICO if d >= 5 else C_ALERTA if d >= 2 else C_OK for d in dias]
            ec({
                "backgroundColor": "transparent",
                "tooltip": _TT_AXIS,
                "grid": {"left":"120px","right":"40px","top":"20px","bottom":"20px"},
                "xAxis": {**_axis(),"type":"value","name":"Días"},
                "yAxis": {**_axis(),"type":"category","data":names,
                          "axisLabel":{"color":"#94A3B8","fontSize":10}},
                "series": [{"type":"bar","barMaxWidth":"28px",
                             "data":[{"value":d,"itemStyle":{"color":c,"borderRadius":[0,6,6,0]}}
                                     for d, c in zip(dias, clrs)],
                             "label":{"show":True,"position":"right","color":"#94A3B8","fontSize":10}}],
            }, height=200)

        # Tareas de TAREAS relacionadas a terceros
        if "TERCERO" in df_raw.columns:
            tareas_terc = _activas()[df_raw["TERCERO"].astype(str).str.strip() != ""]
            if not tareas_terc.empty:
                seccion("📋", "TAREAS CON TERCEROS PENDIENTES", C_ALERTA)
                cols_t = [c for c in ["TAREA","TERCERO","PROYECTO","ESTADO",
                                       "FECHA_COMPROMISO"] if c in tareas_terc.columns]
                df_tt = tareas_terc[cols_t].copy()
                if "FECHA_COMPROMISO" in df_tt.columns:
                    df_tt["FECHA_COMPROMISO"] = df_tt["FECHA_COMPROMISO"].dt.strftime("%d/%m/%Y").fillna("—")
                st.dataframe(df_tt.fillna("—"), use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════════════════════
# MÓDULO 5: PRODUCTIVIDAD
# ══════════════════════════════════════════════════════════════════════════════
elif mod == "Productividad":
    st.markdown(
        f'<div style="background:linear-gradient(100deg,rgba(35,209,96,0.06),transparent);'
        f'border:1px solid rgba(35,209,96,0.14);border-radius:20px;padding:16px 24px;margin-bottom:18px;">'
        f'<div style="font-size:0.58rem;font-weight:800;letter-spacing:0.24em;color:{C_OK};">MÉTRICAS DE DESEMPEÑO</div>'
        f'<div style="font-size:1.65rem;font-weight:900;color:#F8FAFC;">📈 PRODUCTIVIDAD</div>'
        f'</div>', unsafe_allow_html=True)

    comps  = df_raw[df_raw["ESTADO"] == "Completada"]
    ac     = _activas()
    ini_s  = HOY - timedelta(days=HOY.weekday())
    comp_s = comps[comps["FECHA_CIERRE"].notna() & (comps["FECHA_CIERRE"] >= pd.Timestamp(ini_s))]
    comp_m = comps[
        comps["FECHA_CIERRE"].notna() &
        (comps["FECHA_CIERRE"].dt.month == HOY.month) &
        (comps["FECHA_CIERRE"].dt.year  == HOY.year)
    ]
    _both = comps.dropna(subset=["FECHA_CIERRE","FECHA_CREACION"])
    avg_d = (_both["FECHA_CIERRE"] - _both["FECHA_CREACION"]).dt.days.mean() if not _both.empty else None
    tasa_s, _ = tasa_cumplimiento_semana()
    tasa_clr  = C_OK if tasa_s >= 70 else C_ALERTA if tasa_s >= 40 else C_CRITICO

    c1,c2,c3,c4 = st.columns(4)
    with c1: st.markdown(kpi("COMP. ESTA SEMANA", len(comp_s), color=C_OK),                     unsafe_allow_html=True)
    with c2: st.markdown(kpi("COMP. ESTE MES",    len(comp_m), color=C_OK),                     unsafe_allow_html=True)
    with c3: st.markdown(kpi("T° PROM. CIERRE",   f"{avg_d:.0f}d" if avg_d else "—",
                               "desde creación", color=C_CIAN),                                  unsafe_allow_html=True)
    with c4: st.markdown(kpi("TASA SEMANAL",       f"{tasa_s}%", "cumplimiento",
                               color=tasa_clr),                                                   unsafe_allow_html=True)

    st.markdown('<div style="height:6px;"></div>', unsafe_allow_html=True)

    # ── Selector de rango de análisis ─────────────────────────────────────────
    _pr1, _pr2, _pr3 = st.columns([2, 2, 4])
    with _pr1:
        _prod_desde = st.date_input("Desde", value=HOY - timedelta(days=56), key="prod_desde")
    with _pr2:
        _prod_hasta = st.date_input("Hasta", value=HOY, key="prod_hasta")
    with _pr3:
        _prod_n_sem = max(1, ((_prod_hasta - _prod_desde).days // 7) + 1)
        _comps_rng = comps[
            comps["FECHA_CIERRE"].notna() &
            (comps["FECHA_CIERRE"].dt.date >= _prod_desde) &
            (comps["FECHA_CIERRE"].dt.date <= _prod_hasta)
        ]
        _tasa_prom_rng = round(len(_comps_rng) / max(_prod_n_sem, 1), 1)
        st.markdown(
            f'<div style="margin-top:24px;">'
            + kpi("PROM. COMPLETADAS/SEMANA", f"{_tasa_prom_rng:.1f}", f"en {_prod_n_sem} semanas analizadas", color=C_OK)
            + '</div>', unsafe_allow_html=True)

    # Gráfico 1: completadas + creadas por semana (rango seleccionado)
    seccion("📊", "FLUJO SEMANAL — COMPLETADAS VS CREADAS", C_OK)
    _sem_lbl, _comp_sem_data, _crea_sem_data = [], [], []
    _cur = _prod_desde - timedelta(days=_prod_desde.weekday())
    while _cur <= _prod_hasta:
        _fin_w = _cur + timedelta(days=6)
        _n_comp = len(comps[
            comps["FECHA_CIERRE"].notna() &
            (comps["FECHA_CIERRE"].dt.date >= _cur) &
            (comps["FECHA_CIERRE"].dt.date <= _fin_w)
        ])
        _n_crea = len(df_raw[
            df_raw["FECHA_CREACION"].notna() &
            (df_raw["FECHA_CREACION"].dt.date >= _cur) &
            (df_raw["FECHA_CREACION"].dt.date <= _fin_w)
        ])
        _sem_lbl.append(_cur.strftime("W%V\n%d/%m"))
        _comp_sem_data.append(_n_comp)
        _crea_sem_data.append(_n_crea)
        _cur += timedelta(days=7)

    # Media móvil 3 semanas para tendencia
    _comp_mv = []
    for _si in range(len(_comp_sem_data)):
        _win = _comp_sem_data[max(0,_si-2):_si+1]
        _comp_mv.append(round(sum(_win)/len(_win), 1))

    ec({
        "backgroundColor": "transparent",
        "tooltip": _TT_AXIS,
        "legend": {"bottom":"2%","textStyle":{"color":C_GRIS,"fontSize":10}},
        "grid": {"left":"40px","right":"20px","top":"20px","bottom":"50px"},
        "xAxis": {**_axis(),"type":"category","data":_sem_lbl,
                  "axisLabel":{"color":"#64748B","fontSize":9}},
        "yAxis": {**_axis(),"type":"value"},
        "series": [
            {"name":"Completadas","type":"bar","barMaxWidth":"30px","data":_comp_sem_data,
             "itemStyle":{"borderRadius":[5,5,0,0],
                          "color":{"type":"linear","x":0,"y":0,"x2":0,"y2":1,
                                   "colorStops":[{"offset":0,"color":"rgba(35,209,96,0.95)"},
                                                 {"offset":1,"color":"rgba(35,209,96,0.25)"}]}}},
            {"name":"Creadas","type":"bar","barMaxWidth":"30px","data":_crea_sem_data,
             "itemStyle":{"borderRadius":[5,5,0,0],
                          "color":{"type":"linear","x":0,"y":0,"x2":0,"y2":1,
                                   "colorStops":[{"offset":0,"color":"rgba(56,189,248,0.55)"},
                                                 {"offset":1,"color":"rgba(56,189,248,0.12)"}]}}},
            {"name":"Tendencia completadas","type":"line","data":_comp_mv,
             "smooth":True,"symbol":"none",
             "lineStyle":{"color":"rgba(35,209,96,0.70)","width":2,"type":"dashed"},
             "itemStyle":{"color":"rgba(35,209,96,0.70)"}},
        ],
    }, height=230)

    col_p1, col_p2 = st.columns(2)
    with col_p1:
        seccion("📊", "POR CATEGORÍA", C_INDIGO)
        cat_c = df_raw.groupby("CATEGORIA")["ESTADO"].apply(
            lambda x: (x == "Completada").sum()
        ).sort_values()
        ec({
            "backgroundColor": "transparent",
            "tooltip": _TT_AXIS,
            "grid": {"left":"140px","right":"40px","top":"10px","bottom":"10px"},
            "xAxis": {**_axis(),"type":"value"},
            "yAxis": {**_axis(),"type":"category","data":cat_c.index.tolist(),
                      "axisLabel":{"color":"#94A3B8","fontSize":9}},
            "series": [{"type":"bar","barMaxWidth":"22px",
                        "data":[int(v) for v in cat_c.values],
                        "label":{"show":True,"position":"right","color":"#94A3B8","fontSize":9},
                        "itemStyle":{"borderRadius":[0,6,6,0],
                                     "color":"rgba(168,85,247,0.80)"}}],
        }, height=230)

    with col_p2:
        seccion("📊", "PENDIENTES POR PRIORIDAD", C_ALERTA)
        prio_p_all = ac["PRIORIDAD"].value_counts()
        ec({
            "backgroundColor": "transparent",
            "tooltip": _TT,
            "legend": {"bottom":"2%","textStyle":{"color":C_GRIS,"fontSize":10}},
            "series": [{"type":"pie","radius":["40%","68%"],"center":["50%","44%"],
                        "avoidLabelOverlap":True,
                        "label":{"show":False},
                        "emphasis":{"label":{"show":True,"fontSize":13,"fontWeight":"bold","color":"#F1F5F9"}},
                        "data":[{"name":k,"value":int(v),
                                  "itemStyle":{"color":PRIO_CLR.get(k,C_GRIS),"borderWidth":2,"borderColor":"#060B15"}}
                                 for k,v in prio_p_all.items()]}],
        }, height=230)

    # ── Tabla de avance por proyecto ──────────────────────────────────────────
    seccion("📋", "AVANCE POR PROYECTO", C_CIAN)
    _proj_list = [p for p in df_raw["PROYECTO"].dropna().unique() if str(p).strip()]
    _proj_tbl  = []
    for _pj in sorted(_proj_list):
        _dfpj = df_raw[df_raw["PROYECTO"] == _pj]
        _n_comp_pj = int((_dfpj["ESTADO"] == "Completada").sum())
        _n_proc_pj = int((_dfpj["ESTADO"] == "En Proceso").sum())
        _n_pend_pj = int((_dfpj["ESTADO"] == "Pendiente").sum())
        _n_canc_pj = int((_dfpj["ESTADO"].isin(["Cancelada"])).sum())
        _n_tot_pj  = len(_dfpj) - _n_canc_pj
        _avance_pj = round(_n_comp_pj / max(_n_tot_pj, 1) * 100)
        _hh_est_pj = round(float(_dfpj[~_dfpj["ESTADO"].isin(["Completada","Cancelada"])]["ESFUERZO_HRS"].fillna(0).sum()), 1)
        _proj_tbl.append({
            "Proyecto": _pj,
            "Completadas": _n_comp_pj,
            "En Proceso": _n_proc_pj,
            "Pendientes": _n_pend_pj,
            "% Avance": f"{_avance_pj}%",
            "HH pendientes": f"{_hh_est_pj}h",
        })
    _df_proj_tbl = pd.DataFrame(_proj_tbl)
    if not _df_proj_tbl.empty:
        st.dataframe(_df_proj_tbl, use_container_width=True, hide_index=True,
                     height=min(400, 55 + len(_df_proj_tbl)*36))

# ══════════════════════════════════════════════════════════════════════════════
# MÓDULO 7: CONSUMO DE TIEMPO
# ══════════════════════════════════════════════════════════════════════════════
elif mod == "Consumo de Tiempo":
    st.markdown(
        f'<div style="background:linear-gradient(100deg,rgba(168,85,247,0.06),transparent);'
        f'border:1px solid rgba(168,85,247,0.14);border-radius:20px;padding:16px 24px;margin-bottom:18px;">'
        f'<div style="font-size:0.58rem;font-weight:800;letter-spacing:0.24em;color:{C_MORADO};">DISTRIBUCIÓN DE ESFUERZO</div>'
        f'<div style="font-size:1.65rem;font-weight:900;color:#F8FAFC;">⏱ CONSUMO DE TIEMPO</div>'
        f'<div style="font-size:0.70rem;color:{C_GRIS};margin-top:3px;">Basado en ESFUERZO_HRS de tareas activas</div>'
        f'</div>', unsafe_allow_html=True)

    ac = _activas()
    if "ESFUERZO_HRS" not in ac.columns:
        st.warning("Columna ESFUERZO_HRS no encontrada. Regenera el Excel con crear_excel_v2.py")
    else:
        tot_hrs   = ac["ESFUERZO_HRS"].sum()
        cat_hrs   = ac.groupby("CATEGORIA")["ESFUERZO_HRS"].sum().sort_values(ascending=False)
        proj_hrs  = ac.groupby("PROYECTO")["ESFUERZO_HRS"].sum().sort_values(ascending=False)
        area_hrs  = ac.groupby("AREA")["ESFUERZO_HRS"].sum() if "AREA" in ac.columns else pd.Series()
        cat_top   = cat_hrs.idxmax() if not cat_hrs.empty else "—"
        proj_top  = proj_hrs.idxmax() if not proj_hrs.empty else "—"

        c1,c2,c3 = st.columns(3)
        with c1: st.markdown(kpi("HORAS ESTIMADAS TOTALES", f"{tot_hrs:.0f}h", "tareas activas",
                                  color=C_MORADO),                              unsafe_allow_html=True)
        with c2: st.markdown(kpi("CATEGORÍA MÁS DEMANDANTE", cat_top,
                                  f"{cat_hrs.get(cat_top, 0):.0f}h estimadas",
                                  color=CAT_CLR.get(cat_top, C_MORADO)),        unsafe_allow_html=True)
        with c3: st.markdown(kpi("PROYECTO MÁS DEMANDANTE", proj_top,
                                  f"{proj_hrs.get(proj_top, 0):.0f}h estimadas",
                                  color=C_CIAN),                                unsafe_allow_html=True)

        st.markdown('<div style="height:6px;"></div>', unsafe_allow_html=True)
        col_pie, col_bar = st.columns([1, 1])

        with col_pie:
            seccion("🕐", "DISTRIBUCIÓN POR CATEGORÍA", C_MORADO)
            pie_data = [{"name": k, "value": round(float(v), 1),
                         "itemStyle": {"color": CAT_CLR.get(k, C_GRIS),
                                       "borderWidth": 2, "borderColor": "#060B15"}}
                        for k, v in cat_hrs.items() if v > 0]
            ec({
                "backgroundColor": "transparent",
                "tooltip": {**_TT, "formatter": "{b}: {c}h ({d}%)"},
                "legend": {"bottom":"2%","left":"center",
                           "textStyle":{"color":C_GRIS,"fontSize":9},"orient":"horizontal"},
                "series": [{"type":"pie","radius":["38%","65%"],"center":["50%","46%"],
                             "avoidLabelOverlap":True,
                             "label":{"show":False},
                             "emphasis":{"label":{"show":True,"fontSize":12,"fontWeight":"bold",
                                                  "color":"#F1F5F9"}},
                             "data": pie_data}],
            }, height=290)

        with col_bar:
            seccion("📊", "HORAS POR PROYECTO", C_CIAN)
            ph = proj_hrs.sort_values()
            clrs_p = [PRIO_CLR.get("Alta","#FF6B35")] * len(ph)
            ec({
                "backgroundColor": "transparent",
                "tooltip": {**_TT_AXIS, "valueFormatter": "(v)=>{return v.toFixed(1)+'h'}"},
                "grid": {"left":"90px","right":"50px","top":"10px","bottom":"10px"},
                "xAxis": {**_axis(),"type":"value","name":"Horas"},
                "yAxis": {**_axis(),"type":"category","data":ph.index.tolist(),
                          "axisLabel":{"color":"#94A3B8","fontSize":9}},
                "series": [{"type":"bar","barMaxWidth":"24px",
                             "data":[{"value":round(float(v),1),
                                      "itemStyle":{"color":C_CIAN,"borderRadius":[0,6,6,0]}}
                                     for v in ph.values],
                             "label":{"show":True,"position":"right","color":"#94A3B8","fontSize":9,
                                      "formatter":"{c}h"}}],
            }, height=290)

        # ── B5: Estimado vs Real ─────────────────────────────────────────────
        df_all = cargar()
        comp_b5 = df_all[
            (df_all["ESTADO"].isin(["Completada", "Cancelada"])) &
            (df_all["HORAS_REALES"].notna()) &
            (df_all["HORAS_REALES"] > 0) &
            (df_all["ESFUERZO_HRS"].notna()) &
            (df_all["ESFUERZO_HRS"] > 0)
        ].copy()

        if not comp_b5.empty:
            seccion("⚖️", "ESTIMADO vs REAL (TAREAS COMPLETADAS)", C_MORADO)
            tot_est_b5  = float(comp_b5["ESFUERZO_HRS"].sum())
            tot_real_b5 = float(comp_b5["HORAS_REALES"].sum())
            ratio_b5    = tot_real_b5 / tot_est_b5 if tot_est_b5 > 0 else 1.0
            ratio_lbl   = f"{ratio_b5:.2f}×"
            ratio_clr   = C_OK if ratio_b5 <= 1.05 else (C_ALERTA if ratio_b5 <= 1.30 else C_RIESGO)
            ratio_txt   = "Bajo presupuesto" if ratio_b5 <= 1.0 else ("Leve sobreestimación" if ratio_b5 <= 1.30 else "Alta desviación")

            kb1, kb2, kb3 = st.columns(3)
            with kb1:
                st.markdown(kpi("TOTAL ESTIMADO", f"{tot_est_b5:.1f}h",
                                f"{len(comp_b5)} tareas completadas", color=C_MORADO), unsafe_allow_html=True)
            with kb2:
                st.markdown(kpi("TOTAL REAL INVERTIDO", f"{tot_real_b5:.1f}h",
                                f"Δ {tot_real_b5-tot_est_b5:+.1f}h vs estimado", color=C_CIAN), unsafe_allow_html=True)
            with kb3:
                st.markdown(kpi("RATIO DE EFICIENCIA", ratio_lbl,
                                ratio_txt, color=ratio_clr), unsafe_allow_html=True)

            st.markdown('<div style="height:6px;"></div>', unsafe_allow_html=True)

            # Barras agrupadas por categoría
            cat_b5 = comp_b5.groupby("CATEGORIA").agg(
                est=("ESFUERZO_HRS", "sum"), real=("HORAS_REALES", "sum")
            ).sort_values("est", ascending=False)

            if not cat_b5.empty:
                col_b5a, col_b5b = st.columns([3, 2])
                with col_b5a:
                    seccion("📊", "ESTIMADO VS REAL POR CATEGORÍA", C_MORADO)
                    _b5_cats = cat_b5.index.tolist()
                    _b5_est  = [round(float(v), 1) for v in cat_b5["est"]]
                    _b5_real = [round(float(v), 1) for v in cat_b5["real"]]
                    ec({
                        "backgroundColor": "transparent",
                        "tooltip": {**_TT_AXIS},
                        "legend": {"data": ["Estimado", "Real"],
                                   "textStyle": {"color": C_GRIS, "fontSize": 9},
                                   "bottom": "0%"},
                        "grid": {"left": "20px", "right": "20px", "top": "10px", "bottom": "30px",
                                 "containLabel": True},
                        "xAxis": {**_axis(), "type": "category", "data": _b5_cats,
                                  "axisLabel": {"color": "#94A3B8", "fontSize": 9, "rotate": 20}},
                        "yAxis": {**_axis(), "type": "value", "name": "Horas",
                                  "nameTextStyle": {"color": C_GRIS, "fontSize": 9}},
                        "series": [
                            {"name": "Estimado", "type": "bar", "barMaxWidth": "28px",
                             "data": [{"value": v, "itemStyle": {"color": C_MORADO, "borderRadius": [4, 4, 0, 0]}}
                                      for v in _b5_est],
                             "label": {"show": True, "position": "top", "color": "#94A3B8",
                                       "fontSize": 8, "formatter": "{c}h"}},
                            {"name": "Real", "type": "bar", "barMaxWidth": "28px",
                             "data": [{"value": v, "itemStyle": {
                                           "color": C_OK if _b5_real[i] <= _b5_est[i] * 1.05
                                           else (C_ALERTA if _b5_real[i] <= _b5_est[i] * 1.30 else C_RIESGO),
                                           "borderRadius": [4, 4, 0, 0]}}
                                      for i, v in enumerate(_b5_real)],
                             "label": {"show": True, "position": "top", "color": "#94A3B8",
                                       "fontSize": 8, "formatter": "{c}h"}},
                        ],
                    }, height=280)

                with col_b5b:
                    seccion("📋", "TABLA DE EFICIENCIA", C_MORADO)
                    _tbl_rows = []
                    for cat_nm, row in cat_b5.iterrows():
                        _r_b5 = row["real"] / row["est"] if row["est"] > 0 else 1.0
                        _ic   = "🟢" if _r_b5 <= 1.05 else ("🟡" if _r_b5 <= 1.30 else "🔴")
                        _tbl_rows.append({
                            "Categoría": cat_nm,
                            "Est (h)": f'{row["est"]:.1f}',
                            "Real (h)": f'{row["real"]:.1f}',
                            "Ratio": f'{_r_b5:.2f}×',
                            "": _ic,
                        })
                    if _tbl_rows:
                        import pandas as _pd_b5
                        st.dataframe(
                            _pd_b5.DataFrame(_tbl_rows),
                            use_container_width=True, hide_index=True,
                            height=min(38 * len(_tbl_rows) + 40, 280),
                        )
        else:
            st.info("Aún no hay tareas completadas con horas reales registradas. "
                    "Completa una tarea e ingresa las horas reales en el panel de detalle.", icon="ℹ️")

        # Trabajo vs Personal
        if not area_hrs.empty:
            seccion("⚖️", "TRABAJO VS PERSONAL", C_ALERTA)
            trab = float(area_hrs.get("Trabajo", 0))
            pers = float(area_hrs.get("Personal", 0))
            total_a = trab + pers or 1
            st.markdown(
                f'<div style="display:flex;gap:16px;align-items:center;">'
                f'<div style="flex:1;background:rgba(56,189,248,0.06);border:1px solid rgba(56,189,248,0.16);'
                f'border-radius:14px;padding:18px;text-align:center;">'
                f'<div style="font-size:0.52rem;font-weight:800;letter-spacing:0.14em;color:#475569;margin-bottom:6px;">TRABAJO</div>'
                f'<div style="font-size:2.20rem;font-weight:900;color:{C_CIAN};">{trab:.0f}h</div>'
                f'<div style="font-size:0.70rem;color:{C_CIAN};font-weight:700;">{trab/total_a*100:.0f}%</div>'
                f'</div>'
                f'<div style="flex:1;background:rgba(100,116,139,0.06);border:1px solid rgba(100,116,139,0.14);'
                f'border-radius:14px;padding:18px;text-align:center;">'
                f'<div style="font-size:0.52rem;font-weight:800;letter-spacing:0.14em;color:#475569;margin-bottom:6px;">PERSONAL</div>'
                f'<div style="font-size:2.20rem;font-weight:900;color:{C_GRIS};">{pers:.0f}h</div>'
                f'<div style="font-size:0.70rem;color:{C_GRIS};font-weight:700;">{pers/total_a*100:.0f}%</div>'
                f'</div>'
                f'<div style="flex:2;background:rgba(13,21,38,0.60);border:1px solid rgba(56,189,248,0.08);'
                f'border-radius:14px;padding:18px;">'
                f'<div style="font-size:0.52rem;font-weight:800;letter-spacing:0.14em;color:#475569;margin-bottom:10px;">BALANCE DE CARGA</div>'
                f'<div style="height:14px;border-radius:8px;background:rgba(255,255,255,0.06);overflow:hidden;">'
                f'<div style="height:100%;width:{trab/total_a*100:.0f}%;background:{C_CIAN};border-radius:8px;"></div>'
                f'</div>'
                f'<div style="display:flex;justify-content:space-between;margin-top:6px;">'
                f'<span style="font-size:0.58rem;color:{C_CIAN};">{trab:.0f}h Trabajo</span>'
                f'<span style="font-size:0.58rem;color:{C_GRIS};">{pers:.0f}h Personal</span>'
                f'</div></div>'
                f'</div>',
                unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# MÓDULO 7: GANTT
# ══════════════════════════════════════════════════════════════════════════════
elif mod == "Gantt":
    st.markdown(
        f'<div style="background:linear-gradient(100deg,rgba(56,189,248,0.06),transparent);'
        f'border:1px solid rgba(56,189,248,0.14);border-radius:20px;padding:16px 24px;margin-bottom:18px;">'
        f'<div style="font-size:0.58rem;font-weight:800;letter-spacing:0.24em;color:{C_CIAN};">LÍNEA DE TIEMPO · PLANIFICACIÓN Y CONTROL</div>'
        f'<div style="font-size:1.65rem;font-weight:900;color:#F8FAFC;">📊 GANTT</div>'
        f'<div style="font-size:0.70rem;color:{C_GRIS};margin-top:3px;">'
        f'Tareas agrupadas por proyecto · Arrastre el slider inferior para navegar en el tiempo</div>'
        f'</div>', unsafe_allow_html=True)

    # ── Filtros ───────────────────────────────────────────────────────────────
    _gf1, _gf2, _gf3, _gf4 = st.columns(4)
    with _gf1:
        _g_projs = ["Todos"] + sorted(df_raw["PROYECTO"].dropna().unique().tolist())
        _g_proj = st.selectbox("Proyecto", _g_projs, key="gantt_proj")
    with _gf2:
        _g_prio = st.selectbox("Prioridad", ["Todos","Crítica","Alta","Media","Baja"], key="gantt_prio")
    with _gf3:
        _g_est = st.selectbox("Estado", ["Activas","Todas","Completadas"], key="gantt_est")
    with _gf4:
        _g_color = st.selectbox("Color por", ["Prioridad","Estado","Proyecto"], key="gantt_color")

    # ── Filtrar datos ─────────────────────────────────────────────────────────
    _dg = df_raw[~df_raw["ESTADO"].isin(["Cancelada"])].copy()
    if _g_est == "Activas":
        _dg = _dg[~_dg["ESTADO"].isin(["Completada"])]
    elif _g_est == "Completadas":
        _dg = _dg[_dg["ESTADO"] == "Completada"]
    if _g_proj != "Todos":
        _dg = _dg[_dg["PROYECTO"] == _g_proj]
    if _g_prio != "Todos":
        _dg = _dg[_dg["PRIORIDAD"] == _g_prio]

    if _dg.empty:
        st.info("Sin tareas para los filtros seleccionados.")
    else:
        # ── Paleta por modo de color ──────────────────────────────────────────
        _PROJ_COLORS = [
            "#38BDF8","#A855F7","#10B981","#F59E0B","#F97316",
            "#EC4899","#06B6D4","#84CC16","#818CF8","#FF4757",
        ]
        _proj_list = sorted(_dg["PROYECTO"].dropna().unique().tolist())
        _proj_clr  = {p: _PROJ_COLORS[i % len(_PROJ_COLORS)] for i, p in enumerate(_proj_list)}

        def _gantt_color(row):
            if _g_color == "Estado":
                return EST_CLR.get(str(row.get("ESTADO","Pendiente")), C_GRIS)
            elif _g_color == "Proyecto":
                return _proj_clr.get(str(row.get("PROYECTO","")), C_GRIS)
            else:
                return PRIO_CLR.get(str(row.get("PRIORIDAD","Media")), C_GRIS)

        # ── Construir y_labels y series_data ─────────────────────────────────
        _dg = _dg.sort_values(["PROYECTO","PRIO_ORD","FECHA_COMPROMISO"])
        _y_labels   = []   # list de dicts para yAxis.data
        _series_data = []  # list de dicts para series[0].data

        for _gproj in _proj_list:
            _grp = _dg[_dg["PROYECTO"] == _gproj]
            if _grp.empty:
                continue
            # Cabecera de proyecto
            _y_labels.append({
                "value": f"◆  {_gproj}",
                "textStyle": {
                    "color": _proj_clr.get(_gproj, C_CIAN),
                    "fontWeight": "bold",
                    "fontSize": 11,
                    "padding": [3, 0, 3, 0],
                },
            })

            for _, _gr in _grp.iterrows():
                _y_idx = len(_y_labels)

                _t_start = _gr.get("FECHA_CREACION")
                _t_end   = _gr.get("FECHA_COMPROMISO")
                if pd.isna(_t_start):
                    _t_start = pd.Timestamp(HOY)
                if str(_gr.get("ESTADO","")) == "Completada" and pd.notna(_gr.get("FECHA_CIERRE")):
                    _t_end = _gr["FECHA_CIERRE"]
                if pd.isna(_t_end):
                    _esf_raw = float(_gr.get("ESFUERZO_HRS", 0) or 0)
                    _t_end = _t_start + pd.Timedelta(days=max(3, int(_esf_raw)))

                _clr   = _gantt_color(_gr)
                _estado = str(_gr.get("ESTADO","Pendiente"))
                _prio   = str(_gr.get("PRIORIDAD","Media"))
                _esf_v  = float(_gr.get("ESFUERZO_HRS", 0) or 0)
                _terc_v = str(_gr.get("TERCERO","") or "")
                _nm_raw = str(_gr["TAREA"])

                _start_ms = int(_t_start.timestamp() * 1000)
                _end_ms   = int(_t_end.timestamp() * 1000)
                if _end_ms <= _start_ms:
                    _end_ms = _start_ms + 2 * 86400 * 1000

                _series_data.append({
                    "value":     [_y_idx, _start_ms, _end_ms],
                    "name":      _nm_raw,
                    "itemStyle": {"color": _clr, "opacity": 0.45 if _estado == "Completada" else 0.90},
                    "_prio":   _prio,
                    "_estado": _estado,
                    "_proj":   _gproj,
                    "_esf":    _esf_v,
                    "_terc":   _terc_v,
                    "_id":     int(_gr.get("ID", 0)),
                })
                _y_labels.append({
                    "value": f"   {_nm_raw[:42]}",
                    "textStyle": {"color": "#94A3B8", "fontSize": 10},
                })

        # ── Rango de fechas ───────────────────────────────────────────────────
        _today_ms = int(pd.Timestamp(HOY).timestamp() * 1000)
        if _series_data:
            _all_s = [d["value"][1] for d in _series_data]
            _all_e = [d["value"][2] for d in _series_data]
            _xmin  = min(_all_s) - 14 * 86400 * 1000
            _xmax  = max(_all_e) + 21 * 86400 * 1000
        else:
            _xmin = _today_ms - 30 * 86400 * 1000
            _xmax = _today_ms + 90 * 86400 * 1000

        _zoom_start = max(_xmin, _today_ms - 20 * 86400 * 1000)
        _zoom_end   = min(_xmax, _today_ms + 70 * 86400 * 1000)

        _n_rows      = len(_y_labels)
        _row_px      = 30
        _gantt_h     = max(520, _n_rows * _row_px + 130)
        _today_label = HOY.strftime("%d/%m/%Y")

        # ── Construir opción ECharts ──────────────────────────────────────────
        _opt_g = {
            "backgroundColor": "transparent",
            "tooltip": {
                "trigger": "item",
                "backgroundColor": "rgba(6,11,21,0.97)",
                "borderColor": f"rgba({_ABR},0.22)",
                "borderRadius": 12,
                "padding": [10, 14],
                "textStyle": {"color": "#F1F5F9", "fontSize": 11},
                "formatter": "__GT_TT__",
            },
            "legend": {
                "show": _g_color != "Proyecto",
                "data": (
                    [{"name":"Crítica","itemStyle":{"color":C_CRITICO}},
                     {"name":"Alta","itemStyle":{"color":"#FF6B35"}},
                     {"name":"Media","itemStyle":{"color":C_CIAN}},
                     {"name":"Baja","itemStyle":{"color":C_GRIS}}]
                    if _g_color == "Prioridad" else
                    [{"name":k,"itemStyle":{"color":v}} for k, v in EST_CLR.items()]
                ),
                "top": 6, "right": 20,
                "textStyle": {"color": "#64748B", "fontSize": 10},
            },
            "grid": {"top": 50, "bottom": 70, "left": 290, "right": 30},
            "xAxis": {
                "type": "time",
                "min": _xmin,
                "max": _xmax,
                "splitLine": {"lineStyle": {"color": "rgba(255,255,255,0.04)", "type": "dashed"}},
                "axisLabel": {"color": "#64748B", "fontSize": 10, "formatter": "__GT_XFMT__"},
                "axisLine": {"lineStyle": {"color": "#1E293B"}},
                "axisTick": {"lineStyle": {"color": "#1E293B"}},
            },
            "yAxis": {
                "type": "category",
                "data": _y_labels,
                "inverse": True,
                "axisLine": {"show": False},
                "axisTick": {"show": False},
                "splitLine": {"show": True, "lineStyle": {"color": "rgba(255,255,255,0.025)"}},
                "axisLabel": {
                    "color": "#94A3B8",
                    "fontSize": 10,
                    "width": 270,
                    "overflow": "truncate",
                    "align": "right",
                    "rich": {},
                },
            },
            "dataZoom": [
                {
                    "type": "slider",
                    "xAxisIndex": 0,
                    "startValue": _zoom_start,
                    "endValue": _zoom_end,
                    "height": 20,
                    "bottom": 12,
                    "borderColor": f"rgba({_ABR},0.18)",
                    "fillerColor": f"rgba({_ABR},0.07)",
                    "handleStyle": {"color": f"rgb({_ABR})"},
                    "moveHandleStyle": {"color": f"rgba({_ABR},0.50)"},
                    "textStyle": {"color": "#475569", "fontSize": 10},
                    "labelFormatter": "__GT_DZ__",
                },
                {"type": "inside", "xAxisIndex": 0, "startValue": _zoom_start, "endValue": _zoom_end},
            ],
            "series": [{
                "type": "custom",
                "renderItem": "__GT_RENDER__",
                "encode": {"x": [1, 2], "y": 0},
                "data": _series_data,
                "clip": True,
                "markLine": {
                    "silent": True,
                    "symbol": ["none", "none"],
                    "lineStyle": {"color": "rgba(255,71,87,0.70)", "type": "dashed", "width": 1.5},
                    "label": {
                        "show": True,
                        "position": "insideEndBottom",
                        "formatter": f"HOY {_today_label}",
                        "color": "#FF4757",
                        "fontSize": 9,
                        "fontWeight": "bold",
                        "backgroundColor": "rgba(255,71,87,0.12)",
                        "padding": [2, 5],
                        "borderRadius": 4,
                    },
                    "data": [{"xAxis": _today_ms}],
                },
            }],
        }

        _sg = json.dumps(_opt_g, ensure_ascii=False)

        # Inyectar funciones JS
        _sg = _sg.replace('"__GT_RENDER__"', r"""function(params, api) {
    var yIdx = api.value(0);
    var start = api.coord([api.value(1), yIdx]);
    var end   = api.coord([api.value(2), yIdx]);
    var h = Math.max(api.size([0,1])[1] * 0.52, 16);
    var w = Math.max(end[0] - start[0], 3);
    var clr = api.visual('color');
    var children = [{
        type: 'rect',
        transition: ['shape'],
        shape: {x: start[0], y: start[1]-h/2, width: w, height: h, r: 4},
        style: {fill: clr, shadowBlur: 5, shadowColor: 'rgba(0,0,0,0.30)', shadowOffsetY: 2},
        emphasis: {style: {shadowBlur: 14, shadowColor: 'rgba(0,0,0,0.50)'}},
    }];
    if (w > 38) {
        children.push({
            type: 'text',
            style: {
                text: params.data.name.length > 22 ? params.data.name.substring(0,20)+'…' : params.data.name,
                x: start[0]+6, y: start[1],
                textVerticalAlign: 'middle',
                fill: 'rgba(255,255,255,0.92)',
                fontSize: 9, fontWeight: '600',
                width: w-12, overflow: 'truncate',
            },
            silent: true,
        });
    }
    return {type:'group', children: children};
}""")

        _sg = _sg.replace('"__GT_TT__"', r"""function(p) {
    if (!p.data || !Array.isArray(p.data.value)) return '';
    var d = p.data;
    var fmt = function(ms) {
        var dt = new Date(ms);
        return ('0'+dt.getDate()).slice(-2)+'/'+('0'+(dt.getMonth()+1)).slice(-2)+'/'+dt.getFullYear();
    };
    var dur = Math.round((d.value[2]-d.value[1])/(86400*1000));
    return '<div style="max-width:260px;line-height:1.7;">'
        +'<b style="font-size:12px;color:#F8FAFC;">'+d.name+'</b><br>'
        +'<span style="color:#64748B;font-size:10px;">📁 '+d._proj+'</span><br>'
        +'<span style="color:#94A3B8;font-size:10px;">'
        +d._prio+' &nbsp;·&nbsp; '+d._estado+'</span><br>'
        +(d._esf>0?'<span style="color:#38BDF8;font-size:10px;">⏳ '+d._esf+'h estimadas</span><br>':'')
        +(d._terc?'<span style="color:#FFB300;font-size:10px;">👤 '+d._terc+'</span><br>':'')
        +'<span style="color:#475569;font-size:10px;">📅 '+fmt(d.value[1])+' → '+fmt(d.value[2])
        +' ('+dur+' días)</span>'
        +'</div>';
}""")

        _sg = _sg.replace('"__GT_XFMT__"', r"""function(val) {
    var d = new Date(val);
    var m = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
    return d.getDate()+'\n'+m[d.getMonth()];
}""")

        _sg = _sg.replace('"__GT_DZ__"', r"""function(val) {
    var d = new Date(val);
    var m = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
    return d.getDate()+' '+m[d.getMonth()];
}""")

        _html_gantt = (
            f'<!DOCTYPE html><html><head><meta charset="utf-8">'
            f'<script src="https://cdn.jsdelivr.net/npm/echarts@5.4.3/dist/echarts.min.js"></script>'
            f'</head>'
            f'<body style="margin:0;padding:0;background:transparent;">'
            f'<div id="g" style="width:100%;height:{_gantt_h}px;"></div>'
            f'<script>'
            f'var ch=echarts.init(document.getElementById("g"),null,'
            f'{{renderer:"canvas",width:"auto",height:{_gantt_h}}});'
            f'ch.setOption({_sg});'
            f'window.addEventListener("resize",function(){{ch.resize();}});'
            f'</script>'
            f'</body></html>'
        )
        components.html(_html_gantt, height=_gantt_h + 12, scrolling=True)

        # ── Leyenda de proyectos (si color = Proyecto) ────────────────────────
        if _g_color == "Proyecto":
            _leg_h = ""
            for _pp, _pc in _proj_clr.items():
                if _pp in _proj_list:
                    _leg_h += (f'<span style="display:inline-flex;align-items:center;gap:5px;'
                               f'margin:3px 6px;font-size:0.62rem;font-weight:700;color:{_pc};">'
                               f'<span style="width:14px;height:8px;border-radius:3px;'
                               f'background:{_pc};display:inline-block;"></span>{_pp}</span>')
            st.markdown(
                f'<div style="display:flex;flex-wrap:wrap;gap:2px;padding:6px 0;">{_leg_h}</div>',
                unsafe_allow_html=True)

        # ── KPIs de contexto ──────────────────────────────────────────────────
        st.markdown('<div style="height:8px;"></div>', unsafe_allow_html=True)
        seccion("📋", "RESUMEN DE TAREAS EN VISTA", C_CIAN)
        _gk1, _gk2, _gk3, _gk4, _gk5 = st.columns(5)
        _g_venc = _dg[_dg["FECHA_COMPROMISO"].notna() & (_dg["FECHA_COMPROMISO"] < HOY_TS) &
                      ~_dg["ESTADO"].isin(["Completada"])]
        _g_comp = _dg[_dg["ESTADO"] == "Completada"]
        _g_esf_tot = _dg["ESFUERZO_HRS"].fillna(0).sum()
        with _gk1: st.markdown(kpi("EN VISTA", len(_dg), color=C_CIAN), unsafe_allow_html=True)
        with _gk2: st.markdown(kpi("COMPLETADAS", len(_g_comp), color=C_OK), unsafe_allow_html=True)
        with _gk3: st.markdown(kpi("VENCIDAS", len(_g_venc),
                                    color=C_CRITICO if len(_g_venc) else C_GRIS), unsafe_allow_html=True)
        with _gk4: st.markdown(kpi("HH ESTIMADAS", f"{_g_esf_tot:.0f}h",
                                    color=C_MORADO), unsafe_allow_html=True)
        with _gk5:
            _g_tasa = round(len(_g_comp) / max(len(_dg), 1) * 100)
            _g_tclr = C_OK if _g_tasa >= 60 else C_ALERTA if _g_tasa >= 30 else C_CRITICO
            st.markdown(kpi("% COMPLETADO", f"{_g_tasa}%", color=_g_tclr), unsafe_allow_html=True)
