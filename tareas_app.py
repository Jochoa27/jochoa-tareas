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
            "IMPACTO": 3, "URGENCIA": 3, "ESFUERZO_HRS": 1.0,
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
    try:
        dr = pd.read_excel(ARCHIVO, sheet_name="REUNIONES")
        for col in ["FECHA","FECHA_COMP"]:
            if col in dr.columns:
                dr[col] = pd.to_datetime(dr[col], errors="coerce").astype("datetime64[us]")
        data["reuniones"] = dr
    except:
        data["reuniones"] = pd.DataFrame()
    return data

mtime  = ARCHIVO.stat().st_mtime if ARCHIVO.exists() else 0
_data  = cargar(mtime)
df_raw = _data.get("tareas", pd.DataFrame())
df_ter = _data.get("terceros", pd.DataFrame())
df_reu = _data.get("reuniones", pd.DataFrame())

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

def guardar_github(df_tareas_nuevo):
    """Sobreescribe TAREAS sheet en el repo vía GitHub API y limpia caché."""
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
        # Preservar otras hojas
        ter_save = df_ter.drop(columns=["DIAS_SIN_RESP"], errors="ignore")
        if not ter_save.empty:
            ter_save.to_excel(writer, sheet_name="TERCEROS", index=False)
        if not df_reu.empty:
            df_reu.to_excel(writer, sheet_name="REUNIONES", index=False)
    buf.seek(0)

    ts      = pd.Timestamp.now().strftime("%d/%m %H:%M")
    payload = {
        "message": f"update: tareas actualizadas desde app ({ts})",
        "content": base64.b64encode(buf.read()).decode(),
        "sha":     sha,
        "branch":  _GH_BRANCH,
    }
    r2 = requests.put(api, json=payload, headers=hdrs, timeout=15)
    if r2.status_code in (200, 201):
        cargar.clear()
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
    ("⚠️", "Diagnóstico de Riesgo"),
    ("📋", "Bandeja Operacional"),
    ("👥", "Seguimiento de Terceros"),
    ("📈", "Productividad"),
    ("⏱", "Consumo de Tiempo"),
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
                    _dfc.loc[_mk, "FECHA_COMPROMISO"] = pd.Timestamp(_val)
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
        ("calendario", "📅 Cal."),
        ("estado",     "🔷 Estado"),
        ("prioridad",  "🎯 Prio."),
        ("area",       "📁 Área"),
        ("categoria",  "🗂 Cat."),
        ("vencidas",   "🚨 Vencidas"),
    ]
    if st.session_state.get("cc_view") not in {v for v, _ in _CC_VIEWS}:
        st.session_state["cc_view"] = "calendario"
    _tv_cols = st.columns(len(_CC_VIEWS))
    for _vi, (_vk, _vl) in enumerate(_CC_VIEWS):
        with _tv_cols[_vi]:
            if st.button(_vl, key=f"btn_vw_{_vk}", use_container_width=True,
                         type="primary" if st.session_state["cc_view"]==_vk else "secondary"):
                st.session_state["cc_view"] = _vk; st.rerun()
    st.markdown('<div style="height:6px;"></div>', unsafe_allow_html=True)

    # ── Formulario nueva tarea (abierto por "+" de kanban / calendario) ────────
    _OPTS_EST_NT  = ["Pendiente","En Proceso","Esperando Terceros","Completada","Cancelada"]
    _OPTS_PRIO_NT = ["Crítica","Alta","Media","Baja"]
    _OPTS_TIPO_NT = ["Tarea","Seguimiento","Compromiso","Reunión"]
    _OPTS_CAT_NT  = ["Planificación","Contratos","Compras","Reportes",
                     "IA y Automatización","Gestión Corporativa","Reuniones","Salud","Personal"]
    _OPTS_AREA_NT = ["Trabajo","Personal"]

    if st.session_state.get("add_task_defaults") is not None:
        _ntd    = st.session_state["add_task_defaults"]
        _nt_fld = _ntd.get("field", "")
        _nt_val = _ntd.get("value", "")
        _def_est  = _nt_val if _nt_fld == "ESTADO"    else "Pendiente"
        _def_prio = _nt_val if _nt_fld == "PRIORIDAD" else "Media"
        _def_area = _nt_val if _nt_fld == "AREA"      else "Trabajo"
        _def_cat  = _nt_val if _nt_fld == "CATEGORIA" else "Planificación"
        try:
            _def_date = pd.Timestamp(_nt_val).date() if _nt_fld == "date" and _nt_val else None
        except Exception:
            _def_date = None
        seccion("➕", "NUEVA TAREA", C_CIAN)
        with st.form("frm_add_task", border=False):
            _nt1, _nt2, _nt3 = st.columns([3, 2, 2])
            with _nt1:
                _nt_tarea = st.text_input("Tarea *", placeholder="Nombre de la tarea...")
                _nt_proj  = st.text_input("Proyecto", placeholder="Nombre del proyecto")
            with _nt2:
                _nt_tipo  = st.selectbox("Tipo", _OPTS_TIPO_NT)
                _nt_area  = st.selectbox("Área", _OPTS_AREA_NT,
                    index=_OPTS_AREA_NT.index(_def_area) if _def_area in _OPTS_AREA_NT else 0)
            with _nt3:
                _nt_prio  = st.selectbox("Prioridad", _OPTS_PRIO_NT,
                    index=_OPTS_PRIO_NT.index(_def_prio) if _def_prio in _OPTS_PRIO_NT else 2)
                _nt_cat   = st.selectbox("Categoría", _OPTS_CAT_NT,
                    index=_OPTS_CAT_NT.index(_def_cat) if _def_cat in _OPTS_CAT_NT else 0)
            _nt4, _nt5 = st.columns([2, 1])
            with _nt4:
                _nt_est   = st.selectbox("Estado", _OPTS_EST_NT,
                    index=_OPTS_EST_NT.index(_def_est) if _def_est in _OPTS_EST_NT else 0)
            with _nt5:
                _nt_fecha = st.date_input("Fecha compromiso", value=_def_date)
            _nt_desc = st.text_area("Descripción / Notas", height=60, placeholder="Detalle opcional...")
            _ntb1, _ntb2 = st.columns(2)
            with _ntb1: _nt_sub = st.form_submit_button("➕ Agregar",  use_container_width=True, type="primary")
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
                    _new_esf = st.number_input("Esfuerzo (hrs)", min_value=0.0, step=0.5,
                                               format="%.1f",
                                               value=float(_r.get("ESFUERZO_HRS",0) or 0))

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
                _dfc2.loc[_mk2, "NOTAS"]            = _new_notas
                if _new_est == "Completada":
                    _cierre_prev = _dfc2.loc[_mk2, "FECHA_CIERRE"].values[0]
                    if pd.isna(_cierre_prev):
                        _dfc2.loc[_mk2, "FECHA_CIERRE"] = pd.Timestamp(HOY)
                if _new_cmt.strip():
                    _ts_c  = pd.Timestamp.now().strftime("%d/%m/%Y %H:%M")
                    _p_cmt = str(_dfc2.loc[_mk2, "COMENTARIOS"].values[0] or "")
                    _blq   = f"[{_ts_c}] {_new_cmt.strip()}"
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
        _DIAS_ES = ["Lun","Mar","Mié","Jue","Vie","Sáb","Dom"]
        _dcols_h = ""
        _mx_t    = 0
        for _i in range(7):
            _d     = HOY + timedelta(days=_i)
            _d_str = _d.strftime("%Y-%m-%d")
            _td    = _cal_ac[_cal_ac["FECHA_COMPROMISO"].dt.date == _d]
            if len(_td) > _mx_t: _mx_t = len(_td)
            _es_h  = (_i == 0)
            _hc    = f"rgb({_ABR})" if _es_h else _TC["meta_clr"]
            _th    = ""
            for _, _t in _td.iterrows():
                _ti2 = int(_t["ID"])
                _dn2 = str(_t.get("ESTADO","")) == "Completada"
                _p2  = str(_t.get("PRIORIDAD","Media"))
                _pc2 = PRIO_CLR.get(_p2, C_GRIS)
                _n2  = (str(_t.get("TAREA",""))
                        .replace("&","&amp;").replace("<","&lt;").replace(">","&gt;"))
                _pj2 = str(_t.get("PROYECTO","") or "")[:18]
                _ch2 = "checked" if _dn2 else ""
                _s2  = "text-decoration:line-through;opacity:0.38;" if _dn2 else ""
                _th += (
                    f'<div class="kc" data-id="{_ti2}">'
                    f'<div class="kc-top">'
                    f'<span class="kc-dot" style="background:{_pc2};"></span>'
                    f'<span class="kc-ico" style="color:{_pc2};">{PRIO_ICO.get(_p2,"")}</span>'
                    f'<input class="kc-chk" type="checkbox" data-id="{_ti2}" {_ch2}>'
                    f'</div>'
                    f'<div class="kc-nm" style="{_s2}">{_n2}</div>'
                    + (f'<div class="kc-meta"><span class="kc-proj">{_pj2}</span></div>' if _pj2 else '')
                    + f'</div>'
                )
            _add_btn = f'<button class="kk-add" data-field="date" data-group="{_d_str}" title="Agregar tarea">+</button>' if _token() else ''
            _dcols_h += (
                f'<div class="kk{"  kk-today" if _es_h else ""}" '
                f'style="border-top:3px solid {_hc}{"80" if not _es_h else ""};">'
                f'<div class="kk-top">'
                f'<div>'
                f'<div class="kk-hdr" style="color:{_hc};">{_DIAS_ES[_d.weekday()]}</div>'
                f'<div class="kk-num">{_d.day}</div>'
                f'</div>'
                f'{_add_btn}</div>'
                f'<div class="kk-cnt">{len(_td)}</div>'
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
    if(e.target.closest('.kc-chk'))return;
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
            _gdf = _kac[_kac[_kfield].fillna("") == _gval]
            if len(_gdf) > _mx_t: _mx_t = len(_gdf)
            _gc2  = _kcolors.get(_gval, C_GRIS)
            _gi   = _kicons.get(_gval, "")
            _sg   = _gval.replace(" ","-").replace("/","-")
            _th2  = ""
            for _, _t in _gdf.iterrows():
                _ti2 = int(_t["ID"])
                _dn2 = str(_t.get("ESTADO","")) == "Completada"
                _nm2 = (str(_t.get("TAREA",""))
                        .replace("&","&amp;").replace("<","&lt;").replace(">","&gt;"))
                _pj2 = str(_t.get("PROYECTO","") or "")[:20]
                _fc2 = _t.get("FECHA_COMPROMISO")
                _fs2 = pd.Timestamp(_fc2).strftime("%d/%m") if pd.notna(_fc2) else ""
                _pr2 = str(_t.get("PRIORIDAD","Media"))
                _pc3 = PRIO_CLR.get(_pr2, C_GRIS)
                _ch3 = "checked" if _dn2 else ""
                _sy3 = "text-decoration:line-through;opacity:0.38;" if _dn2 else ""
                # Para vista vencidas: mostrar días de retraso
                _delay = ""
                if _cv == "vencidas" and pd.notna(_fc2):
                    _days = (HOY_TS - pd.Timestamp(_fc2)).days
                    _delay = f'<span class="kc-delay">−{_days}d</span>'
                _th2 += (
                    f'<div class="kc" data-id="{_ti2}">'
                    f'<div class="kc-top">'
                    f'<span class="kc-dot" style="background:{_pc3};"></span>'
                    f'<span class="kc-ico" style="color:{_pc3};">{PRIO_ICO.get(_pr2,"")}</span>'
                    f'<input class="kc-chk" type="checkbox" data-id="{_ti2}" {_ch3}>'
                    f'</div>'
                    f'<div class="kc-nm" style="{_sy3}">{_nm2}</div>'
                    f'<div class="kc-meta">'
                    + (f'<span class="kc-proj">{_pj2}</span>' if _pj2 else '')
                    + (f'<span class="kc-fc">{_fs2}</span>' if _fs2 else '')
                    + _delay
                    + f'</div></div>'
                )
            _add_btn = (f'<button class="kk-add" data-field="{_kfield}" '
                        f'data-group="{_gval}" title="Nueva tarea">+</button>'
                        if _tok_add and _cv != "vencidas" else "")
            _dcols_h += (
                f'<div class="kk" style="border-top:3px solid {_gc2}55;">'
                f'<div class="kk-top">'
                f'<div class="kk-hdr" style="color:{_gc2};">{_gi} {_gval}</div>'
                f'{_add_btn}</div>'
                f'<div class="kk-cnt">{len(_gdf)} tareas</div>'
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
.kc-delay{{font-size:0.52rem;color:#FF4757;font-weight:700;
           background:rgba(255,71,87,0.10);border-radius:4px;padding:1px 4px;}}
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
    if(e.target.closest('.kc-chk'))return;
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
                cards_html += (
                    f'<div class="sc" data-id="{task_id}" '
                    f'style="background:{bg_c};border:1px solid {brd_c};">'
                    f'<label class="chkw"><input type="checkbox" data-id="{task_id}" {chk_attr}></label>'
                    f'<div class="body">'
                    f'<div class="chips">{pchip} {tchip} {prchip}</div>'
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
    var ids=Array.from(document.getElementById('sl').children).map(function(c){{return c.dataset.id;}}).join(',');
    notify('order',ids);
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
elif mod == "Diagnóstico de Riesgo":
    ac = _activas()

    st.markdown(
        f'<div style="background:linear-gradient(100deg,rgba(255,71,87,0.07),transparent);'
        f'border:1px solid rgba(255,71,87,0.16);border-radius:20px;padding:16px 24px;margin-bottom:18px;">'
        f'<div style="font-size:0.58rem;font-weight:800;letter-spacing:0.24em;color:{C_CRITICO};">ANÁLISIS DE RIESGO OPERACIONAL</div>'
        f'<div style="font-size:1.65rem;font-weight:900;color:#F8FAFC;">⚠️ DIAGNÓSTICO DE RIESGO</div>'
        f'<div style="font-size:0.70rem;color:{C_GRIS};margin-top:3px;">Matriz Urgencia × Impacto · Tamaño = Esfuerzo estimado</div>'
        f'</div>', unsafe_allow_html=True)

    col_mat, col_info = st.columns([3, 1])

    with col_mat:
        series_data = {}
        for _, r in ac.iterrows():
            p   = str(r.get("PRIORIDAD","Media"))
            urg = float(r.get("URGENCIA", 3) or 3)
            imp = float(r.get("IMPACTO",  3) or 3)
            esf = float(r.get("ESFUERZO_HRS", 1) or 1)
            nm  = str(r["TAREA"])[:40]
            if p not in series_data:
                series_data[p] = []
            series_data[p].append([urg, imp, esf, nm])

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
            "tooltip": {"trigger":"item","formatter":"__TT_FN__",
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
        # Top riesgo
        ac2 = ac.copy()
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

    # Tabla completa de riesgo
    seccion("📊", "TABLA DE RIESGO", C_CRITICO)
    ac_t = ac.copy()
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
    sa1, sa2, sa3 = st.columns([2, 2, 8])
    with sa1:
        guardar_btn = st.button("💾 Guardar en GitHub", type="primary",
                                use_container_width=True, disabled=not _token())
    with sa2:
        if st.button("↩ Descartar cambios", use_container_width=True):
            st.rerun()

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

        c1,c2,c3,c4 = st.columns(4)
        with c1: st.markdown(kpi("TERCEROS ACTIVOS",  len(pend_t),
                                  color=C_ALERTA if len(pend_t) else C_GRIS),    unsafe_allow_html=True)
        with c2: st.markdown(kpi("DÍAS PROM. ESPERA", f"{prom_dias}d",
                                  "sin respuesta", color=C_ALERTA),              unsafe_allow_html=True)
        with c3: st.markdown(kpi("PRIORIDAD ALTA",    len(crit_t),
                                  color=C_CRITICO if len(crit_t) else C_GRIS),  unsafe_allow_html=True)
        with c4: st.markdown(kpi("RESUELTOS",
                                  len(df_ter)-len(pend_t), color=C_OK),         unsafe_allow_html=True)

        st.markdown('<div style="height:6px;"></div>', unsafe_allow_html=True)
        seccion("📊", "ESTADO DE TERCEROS", C_ALERTA)

        # Tabla de terceros con color
        df_ter_show = df_ter.copy()
        for col in ["FECHA_INICIO_SEG","FECHA_ULTIMO_SEG"]:
            if col in df_ter_show.columns:
                df_ter_show[col] = df_ter_show[col].dt.strftime("%d/%m/%Y").fillna("—")
        cols_show = [c for c in ["NOMBRE","ORGANIZACION","ROL","TEMA_PENDIENTE",
                                  "FECHA_ULTIMO_SEG","DIAS_SIN_RESP","ESTADO","PRIORIDAD","NOTAS"]
                     if c in df_ter_show.columns]
        st.dataframe(df_ter_show[cols_show].fillna("—"), use_container_width=True, hide_index=True)

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

    # Gráfico 1: completadas por semana (últimas 8 semanas)
    seccion("📊", "COMPLETADAS POR SEMANA", C_OK)
    semanas, cantidades = [], []
    for i in range(7, -1, -1):
        ini_w = HOY - timedelta(days=HOY.weekday() + i*7)
        fin_w = ini_w + timedelta(days=6)
        n = len(comps[
            comps["FECHA_CIERRE"].notna() &
            (comps["FECHA_CIERRE"] >= pd.Timestamp(ini_w)) &
            (comps["FECHA_CIERRE"] <= pd.Timestamp(fin_w))
        ])
        semanas.append(ini_w.strftime("W%V\n%d/%m"))
        cantidades.append(n)

    ec({
        "backgroundColor": "transparent",
        "tooltip": _TT_AXIS,
        "grid": {"left":"40px","right":"20px","top":"20px","bottom":"40px"},
        "xAxis": {**_axis(),"type":"category","data":semanas},
        "yAxis": {**_axis(),"type":"value"},
        "series": [{"type":"bar","barMaxWidth":"36px","data":cantidades,
                    "label":{"show":True,"position":"top","color":"#94A3B8","fontSize":10},
                    "itemStyle":{
                        "borderRadius":[6,6,0,0],
                        "color":{"type":"linear","x":0,"y":0,"x2":0,"y2":1,
                                 "colorStops":[{"offset":0,"color":"rgba(35,209,96,0.95)"},
                                               {"offset":1,"color":"rgba(35,209,96,0.25)"}]}
                    }}],
    }, height=220)

    col_p1, col_p2 = st.columns(2)
    with col_p1:
        seccion("📊", "POR PROYECTO", C_CIAN)
        proj_c = df_raw.groupby("PROYECTO")["ESTADO"].apply(
            lambda x: (x == "Completada").sum()
        ).sort_values()
        ec({
            "backgroundColor": "transparent",
            "tooltip": _TT_AXIS,
            "grid": {"left":"110px","right":"40px","top":"10px","bottom":"10px"},
            "xAxis": {**_axis(),"type":"value"},
            "yAxis": {**_axis(),"type":"category","data":proj_c.index.tolist(),
                      "axisLabel":{"color":"#94A3B8","fontSize":9}},
            "series": [{"type":"bar","barMaxWidth":"22px",
                        "data":[int(v) for v in proj_c.values],
                        "label":{"show":True,"position":"right","color":"#94A3B8","fontSize":9},
                        "itemStyle":{"borderRadius":[0,6,6,0],
                                     "color":"rgba(56,189,248,0.80)"}}],
        }, height=230)

    with col_p2:
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

    # Pendientes por prioridad
    seccion("📊", "PENDIENTES POR PRIORIDAD", C_ALERTA)
    prio_p = ac.groupby("PRIORIDAD").size().reindex(["Crítica","Alta","Media","Baja"]).fillna(0)
    prio_p_all = ac["PRIORIDAD"].value_counts()
    ec({
        "backgroundColor": "transparent",
        "tooltip": _TT,
        "legend": {"bottom":"2%","textStyle":{"color":C_GRIS,"fontSize":10}},
        "series": [{"type":"pie","radius":["40%","68%"],"center":["50%","48%"],
                    "avoidLabelOverlap":True,
                    "label":{"show":False},
                    "emphasis":{"label":{"show":True,"fontSize":13,"fontWeight":"bold","color":"#F1F5F9"}},
                    "data":[{"name":k,"value":int(v),
                              "itemStyle":{"color":PRIO_CLR.get(k,C_GRIS),"borderWidth":2,"borderColor":"#060B15"}}
                             for k,v in prio_p_all.items()]}],
    }, height=220)

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
