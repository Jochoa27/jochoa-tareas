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
from streamlit_sortables import sort_items

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
    ("📅", "Reuniones"),
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
    col_ag, col_prox = st.columns([3, 2])

    # ── Agenda del Día ────────────────────────────────────────────────────────
    with col_ag:
        seccion("📌", "AGENDA DEL DÍA · arrastra para reordenar", C_CRITICO)
        agenda_base = ac[
            (ac["PRIORIDAD"].isin(["Crítica","Alta"])) |
            (ac["FECHA_COMPROMISO"].notna() &
             (ac["FECHA_COMPROMISO"] <= HOY_TS + pd.Timedelta(days=1)))
        ].head(14)

        # Incluir también completadas visibles (para mostrar tachadas)
        agenda_ids_all = set(agenda_base["ID"].astype(int))
        id_to_row = {int(r["ID"]): r for _, r in agenda_base.iterrows()}

        if not id_to_row:
            st.markdown('<div style="color:#334155;font-size:0.80rem;padding:20px;text-align:center;">'
                        '✅ No hay tareas urgentes pendientes</div>', unsafe_allow_html=True)
        else:
            # Mantener/actualizar orden en session_state
            prev = st.session_state.get("agenda_order", [])
            cur_ids = list(id_to_row.keys())
            ordered = [i for i in prev if i in agenda_ids_all]
            ordered += [i for i in cur_ids if i not in ordered]
            st.session_state["agenda_order"] = ordered

            # Labels codificadas con ID para el drag
            labels = [f"[{i}] {str(id_to_row[i]['TAREA'])[:42]}" for i in ordered]
            sorted_labels = sort_items(labels, key="agenda_drag")
            # Actualizar orden según arrastre
            new_order = []
            for lbl in sorted_labels:
                m = re.match(r'\[(\d+)\]', lbl)
                if m:
                    new_order.append(int(m.group(1)))
            st.session_state["agenda_order"] = new_order

            st.markdown('<div style="height:6px;"></div>', unsafe_allow_html=True)

            # Renderizar tarjetas en el nuevo orden
            for task_id in new_order:
                r = id_to_row.get(task_id)
                if r is None:
                    continue
                is_done = str(r.get("ESTADO","")) == "Completada"
                p   = str(r.get("PRIORIDAD","Media"))
                pc  = PRIO_CLR.get(p, C_GRIS)
                fc  = r.get("FECHA_COMPROMISO")
                sc, sd = semaforo(fc)
                venc_r  = pd.notna(fc) and fc < HOY_TS and not is_done
                bg  = "rgba(255,71,87,0.04)" if venc_r else ("rgba(35,209,96,0.03)" if is_done else "rgba(13,21,38,0.70)")
                brd = "rgba(255,71,87,0.18)" if venc_r else (f"{C_OK}20" if is_done else f"{pc}20")
                tipo = str(r.get("TIPO","Tarea"))
                tc   = TIPO_CLR.get(tipo, C_CIAN)
                proj = str(r.get("PROYECTO",""))
                strike = "text-decoration:line-through;opacity:0.45;" if is_done else ""
                nombre_display = f'<span style="{strike}">{r["TAREA"]}</span>'

                c_chk, c_card = st.columns([0.6, 11])
                with c_chk:
                    checked = st.checkbox(
                        "", value=is_done, key=f"chk_{task_id}",
                        label_visibility="collapsed",
                        help="Marcar como completada / desmarcar para reactivar"
                    )
                    if checked != is_done and _token():
                        _marcar_estado(task_id, "Completada" if checked else "Pendiente")
                with c_card:
                    st.markdown(
                        f'<div class="tc" style="background:{bg};border-color:{brd};margin:0 0 4px;">'
                        f'<div>{chip(f"{PRIO_ICO.get(p,"")} {p}", pc)}'
                        f'{chip(tipo, tc)}{chip(proj, C_INDIGO)}</div>'
                        f'<div class="tc-t">{nombre_display}</div>'
                        f'<div class="tc-dt" style="color:{sc};">⏱ {sd}</div>'
                        f'</div>',
                        unsafe_allow_html=True)

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
        st.markdown('<div style="height:6px;"></div>', unsafe_allow_html=True)
        if st.button("📅 Ver en calendario →", use_container_width=True,
                     key="btn_cal_open"):
            st.session_state["show_cal"] = not st.session_state.get("show_cal", False)

    # ── Calendario 7 días (expandible a pantalla completa) ────────────────────
    if st.session_state.get("show_cal", False):
        prox7_cal = ac[
            ac["FECHA_COMPROMISO"].notna() &
            (ac["FECHA_COMPROMISO"] >= HOY_TS) &
            (ac["FECHA_COMPROMISO"] <= HOY_TS + pd.Timedelta(days=7))
        ]
        DIAS_ES = ["Lun","Mar","Mié","Jue","Vie","Sáb","Dom"]
        cal_html = (
            '<div style="display:grid;grid-template-columns:repeat(7,1fr);gap:8px;'
            'margin-top:16px;">'
        )
        for i in range(7):
            d = HOY + timedelta(days=i)
            d_ts = pd.Timestamp(d)
            tareas_dia = prox7_cal[prox7_cal["FECHA_COMPROMISO"].dt.date == d]
            es_hoy = (i == 0)
            hdr_clr = C_CIAN if es_hoy else "#475569"
            num_clr = "#F8FAFC" if es_hoy else "#94A3B8"
            brd_clr = f"{C_CIAN}55" if es_hoy else "#1E293B"
            dia_nombre = DIAS_ES[d.weekday()]
            cal_html += (
                f'<div style="background:rgba(13,21,38,0.85);border:1px solid {brd_clr};'
                f'border-radius:12px;padding:10px 8px;min-height:110px;">'
                f'<div style="font-size:0.54rem;font-weight:800;color:{hdr_clr};'
                f'letter-spacing:0.10em;">{dia_nombre}</div>'
                f'<div style="font-size:1.35rem;font-weight:900;color:{num_clr};'
                f'line-height:1.1;margin-bottom:8px;">{d.day}</div>'
            )
            if tareas_dia.empty:
                cal_html += '<div style="color:#1E293B;font-size:0.60rem;text-align:center;">—</div>'
            else:
                for _, t in tareas_dia.iterrows():
                    p2  = str(t.get("PRIORIDAD","Media"))
                    pc2 = PRIO_CLR.get(p2, C_GRIS)
                    nombre_c = str(t["TAREA"])[:24]
                    cal_html += (
                        f'<div title="{t["TAREA"]}" style="background:{pc2}16;border:1px solid {pc2}30;'
                        f'border-radius:6px;padding:3px 6px;margin-bottom:4px;'
                        f'font-size:0.58rem;color:{pc2};white-space:nowrap;'
                        f'overflow:hidden;text-overflow:ellipsis;">'
                        f'{PRIO_ICO.get(p2,"")} {nombre_c}</div>'
                    )
            cal_html += '</div>'
        cal_html += '</div>'

        btn_cerrar, _ = st.columns([2, 10])
        with btn_cerrar:
            if st.button("✕ Cerrar calendario", key="btn_cal_close"):
                st.session_state["show_cal"] = False
                st.rerun()
        st.markdown(
            f'<div style="background:rgba(6,11,21,0.95);border:1px solid rgba(56,189,248,0.14);'
            f'border-radius:16px;padding:20px 24px;margin-top:4px;">'
            f'<div style="font-size:0.54rem;font-weight:800;letter-spacing:0.22em;'
            f'color:{C_CIAN};margin-bottom:2px;">VISTA CALENDARIO</div>'
            f'<div style="font-size:1.1rem;font-weight:900;color:#F8FAFC;margin-bottom:12px;">'
            f'Próximos 7 días · {HOY.strftime("%d %b")} – {(HOY+timedelta(days=6)).strftime("%d %b %Y")}'
            f'</div>'
            f'{cal_html}</div>',
            unsafe_allow_html=True
        )

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
# MÓDULO 5: REUNIONES Y COMPROMISOS
# ══════════════════════════════════════════════════════════════════════════════
elif mod == "Reuniones":
    st.markdown(
        f'<div style="background:linear-gradient(100deg,rgba(16,185,129,0.06),transparent);'
        f'border:1px solid rgba(16,185,129,0.14);border-radius:20px;padding:16px 24px;margin-bottom:18px;">'
        f'<div style="font-size:0.58rem;font-weight:800;letter-spacing:0.24em;color:{C_VERDE2};">COMPROMISOS Y ACTAS</div>'
        f'<div style="font-size:1.65rem;font-weight:900;color:#F8FAFC;">📅 REUNIONES Y COMPROMISOS</div>'
        f'</div>', unsafe_allow_html=True)

    if df_reu.empty:
        st.warning("No se encontró la hoja REUNIONES. Regenera con crear_excel_v2.py")
    else:
        prox_r = df_reu[
            (df_reu["ESTADO"] != "Realizada") &
            df_reu["FECHA"].notna() &
            (df_reu["FECHA"] >= HOY_TS)
        ] if "FECHA" in df_reu.columns else pd.DataFrame()
        comp_pend = df_reu[df_reu.get("ESTADO_COMP","") == "Pendiente"] \
            if "ESTADO_COMP" in df_reu.columns else pd.DataFrame()
        comp_proc = df_reu[df_reu.get("ESTADO_COMP","") == "En Proceso"] \
            if "ESTADO_COMP" in df_reu.columns else pd.DataFrame()

        c1,c2,c3 = st.columns(3)
        with c1: st.markdown(kpi("REUNIONES PRÓXIMAS", len(prox_r), color=C_VERDE2),  unsafe_allow_html=True)
        with c2: st.markdown(kpi("COMPROMISOS PENDIENTES", len(comp_pend),
                                  color=C_ALERTA if len(comp_pend) else C_GRIS),       unsafe_allow_html=True)
        with c3: st.markdown(kpi("EN PROCESO", len(comp_proc), color=C_CIAN),          unsafe_allow_html=True)

        st.markdown('<div style="height:6px;"></div>', unsafe_allow_html=True)
        seccion("📅", "AGENDA DE REUNIONES", C_VERDE2)

        cards_r = ""
        df_reu_s = df_reu.sort_values("FECHA") if "FECHA" in df_reu.columns else df_reu
        for _, r in df_reu_s.iterrows():
            ft   = r.get("FECHA")
            flab = ft.strftime("%d/%m/%Y") if pd.notna(ft) else "—"
            est  = str(r.get("ESTADO","Pendiente"))
            ec_  = C_OK if est == "Realizada" else C_ALERTA
            tipo = str(r.get("TIPO",""))
            fc   = r.get("FECHA_COMP")
            ecmp = str(r.get("ESTADO_COMP",""))
            ecmp_clr = C_OK if ecmp == "Completada" else C_ALERTA if ecmp == "En Proceso" else C_GRIS
            comprm = str(r.get("COMPROMISO",""))
            resp   = str(r.get("RESPONSABLE_COMP",""))
            cards_r += (
                f'<div style="background:rgba(16,185,129,0.04);border:1px solid rgba(16,185,129,0.12);'
                f'border-radius:14px;padding:16px;margin-bottom:10px;">'
                f'<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;">'
                f'<div style="font-size:0.90rem;font-weight:800;color:#F1F5F9;">'
                f'📅 {str(r.get("TITULO",""))}</div>'
                f'<div>{chip(flab, C_VERDE2)}{chip(tipo, C_INDIGO)}'
                f'<span style="font-size:0.55rem;font-weight:700;color:{ec_};margin-left:6px;">{est}</span></div>'
                f'</div>'
                f'<div style="font-size:0.72rem;color:#94A3B8;margin-bottom:6px;">'
                f'<b style="color:#CBD5E1;">Participantes:</b> {str(r.get("PARTICIPANTES",""))}</div>'
                f'<div style="background:rgba(0,0,0,0.20);border-radius:8px;padding:10px;margin-top:8px;">'
                f'<div style="font-size:0.55rem;font-weight:800;letter-spacing:0.10em;'
                f'color:#334155;text-transform:uppercase;margin-bottom:4px;">COMPROMISOS</div>'
                f'<div style="font-size:0.72rem;color:#CBD5E1;">{comprm}</div>'
                f'<div style="font-size:0.62rem;margin-top:6px;">'
                f'{chip(f"Resp: {resp}", C_CIAN)}'
                f'<span style="font-size:0.58rem;color:{ecmp_clr};margin-left:8px;font-weight:700;">'
                f'{ecmp}</span>'
                f'{" · " + fc.strftime("%d/%m") if pd.notna(fc) else ""}'
                f'</div></div></div>'
            )
        st.markdown(cards_r, unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# MÓDULO 6: PRODUCTIVIDAD
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
