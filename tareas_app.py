# ═══════════════════════════════════════════════════════════════════════════════
# CENTRO DE COMANDO PERSONAL — JUAN OCHOA
# Constructora Londres · Planificación y Control · v2.0 · 2026
# ═══════════════════════════════════════════════════════════════════════════════
import streamlit as st
import pandas as pd
import json
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
        seccion("📌", "AGENDA DEL DÍA", C_CRITICO)
        agenda = ac[
            (ac["PRIORIDAD"].isin(["Crítica","Alta"])) |
            (ac["FECHA_COMPROMISO"].notna() &
             (ac["FECHA_COMPROMISO"] <= HOY_TS + pd.Timedelta(days=1)))
        ].sort_values(["PRIO_ORD","FECHA_COMPROMISO"]).head(12)

        if agenda.empty:
            st.markdown('<div style="color:#334155;font-size:0.80rem;padding:20px;text-align:center;">'
                        '✅ No hay tareas urgentes pendientes</div>', unsafe_allow_html=True)
        else:
            cards_html = ""
            for _, r in agenda.iterrows():
                p   = str(r.get("PRIORIDAD","Media"))
                pc  = PRIO_CLR.get(p, C_GRIS)
                est = str(r.get("ESTADO","Pendiente"))
                ec2 = EST_CLR.get(est, C_GRIS)
                fc  = r.get("FECHA_COMPROMISO")
                sc, sd = semaforo(fc)
                venc_r = pd.notna(fc) and fc < HOY_TS
                bg  = f"rgba(255,71,87,0.06)" if venc_r else "rgba(13,21,38,0.70)"
                brd = f"rgba(255,71,87,0.28)" if venc_r else f"{pc}24"
                tipo = str(r.get("TIPO","Tarea"))
                tc   = TIPO_CLR.get(tipo, C_CIAN)
                proj = str(r.get("PROYECTO",""))
                cards_html += (
                    f'<div class="tc" style="background:{bg};border-color:{brd};">'
                    f'<div>{chip(f"{PRIO_ICO.get(p,"")} {p}", pc)}'
                    f'{chip(tipo, tc)}{chip(proj, C_INDIGO)}</div>'
                    f'<div class="tc-t">{r["TAREA"]}</div>'
                    f'<div class="tc-dt" style="color:{sc};">⏱ {sd}'
                    f' &nbsp;·&nbsp; {chip(est, ec2)}</div>'
                    f'</div>'
                )
            st.markdown(f'<div class="scroll-box">{cards_html}</div>', unsafe_allow_html=True)

    # ── Próximos 7 días ───────────────────────────────────────────────────────
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
            items = ""
            last_day = None
            for _, r in prox7.iterrows():
                fc   = r["FECHA_COMPROMISO"]
                day  = fc.date()
                sc, sd = semaforo(fc)
                p    = str(r.get("PRIORIDAD","Media"))
                pc   = PRIO_CLR.get(p, C_GRIS)
                if day != last_day:
                    label = "HOY" if day == HOY else day.strftime("%a %d/%m").upper()
                    items += (f'<div style="font-size:0.52rem;font-weight:800;letter-spacing:0.12em;'
                              f'color:#334155;margin:10px 0 5px;">{label}</div>')
                    last_day = day
                items += (
                    f'<div style="background:rgba(56,189,248,0.03);border:1px solid rgba(56,189,248,0.08);'
                    f'border-radius:10px;padding:8px 12px;margin-bottom:5px;">'
                    f'<div style="font-size:0.75rem;font-weight:700;color:#E2E8F0;line-height:1.3;">'
                    f'{PRIO_ICO.get(p,"")} {r["TAREA"]}</div>'
                    f'<div style="font-size:0.58rem;color:{sc};margin-top:3px;">{sd} · '
                    f'{chip(str(r.get("PROYECTO","")), C_INDIGO)}</div>'
                    f'</div>'
                )
            st.markdown(f'<div class="scroll-box">{items}</div>', unsafe_allow_html=True)

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
    ac_tshow = ac_t[cols_show].sort_values("RIESGO", ascending=False).copy()
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
    # Filtros
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
    df_f = df_f.sort_values(["PRIO_ORD","FECHA_COMPROMISO"])

    # Stats rápidas
    n_total  = len(df_f)
    n_pend   = int((df_f["ESTADO"]=="Pendiente").sum())
    n_proc   = int((df_f["ESTADO"]=="En Proceso").sum())
    n_venc   = int((df_f["FECHA_COMPROMISO"].notna() &
                    (df_f["FECHA_COMPROMISO"]<HOY_TS) &
                    ~df_f["ESTADO"].isin(["Completada","Cancelada"])).sum())
    st.markdown(
        f'<div style="display:flex;gap:10px;margin:8px 0 14px;flex-wrap:wrap;">'
        f'<span style="background:rgba(56,189,248,0.08);border:1px solid rgba(56,189,248,0.18);'
        f'border-radius:10px;padding:5px 14px;font-size:0.70rem;font-weight:800;color:{C_CIAN};">'
        f'{n_total} TAREAS</span>'
        f'<span style="background:rgba(100,116,139,0.08);border:1px solid rgba(100,116,139,0.18);'
        f'border-radius:10px;padding:5px 14px;font-size:0.70rem;font-weight:800;color:{C_GRIS};">'
        f'⏳ {n_pend} pendientes</span>'
        f'<span style="background:rgba(56,189,248,0.08);border:1px solid rgba(56,189,248,0.16);'
        f'border-radius:10px;padding:5px 14px;font-size:0.70rem;font-weight:800;color:{C_CIAN};">'
        f'🔄 {n_proc} en proceso</span>'
        f'<span style="background:rgba(255,71,87,0.08);border:1px solid rgba(255,71,87,0.18);'
        f'border-radius:10px;padding:5px 14px;font-size:0.70rem;font-weight:800;color:{C_CRITICO};">'
        f'🔴 {n_venc} vencidas</span>'
        f'</div>',
        unsafe_allow_html=True)

    # Tabla
    cols_t = [c for c in ["TAREA","TIPO","PROYECTO","AREA","CATEGORIA","PRIORIDAD",
                            "ESTADO","URGENCIA","IMPACTO","ESFUERZO_HRS","TERCERO",
                            "FECHA_COMPROMISO","NOTAS"] if c in df_f.columns]
    df_show = df_f[cols_t].copy()
    if "FECHA_COMPROMISO" in df_show.columns:
        df_show["FECHA_COMPROMISO"] = df_show["FECHA_COMPROMISO"].dt.strftime("%d/%m/%Y").fillna("—")
    st.dataframe(df_show.fillna("—"), use_container_width=True, hide_index=True,
                 height=min(650, 55 + len(df_show)*36))
    st.markdown(
        f'<div style="font-size:0.58rem;color:#334155;margin-top:8px;text-align:right;">'
        f'📝 Editar: abre TAREAS_JOCHOA.xlsx localmente → git push para actualizar</div>',
        unsafe_allow_html=True)

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
