import streamlit as st
import pandas as pd
import json
from pathlib import Path
from datetime import date, timedelta
import streamlit.components.v1 as components

# ─── CONFIG ───────────────────────────────────────────────────────────────────
ARCHIVO = Path(__file__).parent / "TAREAS_JOCHOA.xlsx"
HOY = date.today()

PRIO_ORD = {"Crítica": 0, "Alta": 1, "Media": 2, "Baja": 3}
PRIO_CLR = {"Crítica": "#FF4757", "Alta": "#FF6B35", "Media": "#38BDF8", "Baja": "#64748B"}
PRIO_ICO = {"Crítica": "🔴", "Alta": "🟠", "Media": "🔵", "Baja": "⚪"}
EST_CLR  = {
    "Pendiente":          "#64748B",
    "En Proceso":         "#38BDF8",
    "Esperando Terceros": "#FFB300",
    "Completada":         "#23D160",
}
EST_ICO  = {"Pendiente": "⏳", "En Proceso": "🔄", "Esperando Terceros": "⌛", "Completada": "✅"}

st.set_page_config(
    page_title="Tareas — Jochoa",
    page_icon="✅",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── CSS ──────────────────────────────────────────────────────────────────────
st.markdown("""<style>
[data-testid="stAppViewContainer"]{background:#060B15;}
[data-testid="stSidebar"]{background:#08101F;border-right:1px solid rgba(56,189,248,0.10);}
.block-container{padding:1.2rem 1.5rem 2rem;}
[data-testid="stSidebar"] label{color:#94A3B8 !important;}
div[data-baseweb="tab-list"]{background:rgba(56,189,248,0.05);border-radius:12px;padding:4px;gap:4px;}
button[data-baseweb="tab"]{color:#64748B;font-weight:700;font-size:0.78rem;letter-spacing:0.05em;border-radius:8px;}
button[data-baseweb="tab"][aria-selected="true"]{color:#38BDF8 !important;background:rgba(56,189,248,0.12) !important;}
[data-testid="stTextInput"] input{background:rgba(56,189,248,0.05);border:1px solid rgba(56,189,248,0.20) !important;color:#F1F5F9;border-radius:10px;}
.stMultiSelect [data-baseweb="select"]{background:rgba(56,189,248,0.04) !important;border:1px solid rgba(56,189,248,0.18) !important;border-radius:10px;}
/* KPI cards */
.tk-kpi{border-radius:14px;padding:14px 16px 12px;min-height:90px;display:flex;flex-direction:column;justify-content:center;transition:border-color .2s,box-shadow .2s;}
.tk-lbl{font-size:0.52rem;font-weight:800;letter-spacing:0.14em;text-transform:uppercase;color:#475569;margin-bottom:4px;}
.tk-val{font-size:1.55rem;font-weight:900;line-height:1;font-variant-numeric:tabular-nums;}
.tk-sub{font-size:0.62rem;font-weight:600;margin-top:4px;opacity:0.78;}
/* Kanban */
.kb-wrap{background:rgba(255,255,255,0.018);border:1px solid rgba(255,255,255,0.06);border-radius:16px;padding:14px 12px;min-height:300px;}
.kb-hdr{font-size:0.65rem;font-weight:800;letter-spacing:0.16em;text-transform:uppercase;margin-bottom:12px;padding-bottom:9px;border-bottom:1px solid rgba(255,255,255,0.07);display:flex;align-items:center;justify-content:space-between;}
.kb-n{font-size:0.60rem;font-weight:700;padding:2px 10px;border-radius:12px;color:#060B15;}
/* Task card */
.tc{border-radius:12px;padding:12px 14px;margin-bottom:8px;border:1px solid transparent;}
.tc-title{font-size:0.80rem;font-weight:700;color:#F1F5F9;line-height:1.38;margin:5px 0 7px;}
.tc-tags{display:flex;gap:5px;flex-wrap:wrap;align-items:center;margin-top:3px;}
.tc-tag{font-size:0.52rem;font-weight:700;letter-spacing:0.05em;padding:2px 8px;border-radius:8px;text-transform:uppercase;}
.tc-date{font-size:0.60rem;font-weight:700;margin-top:7px;}
/* Section headers */
.sh{display:flex;align-items:center;gap:12px;margin:1.6rem 0 0.8rem;padding-bottom:10px;border-bottom:1px solid rgba(56,189,248,0.12);}
.sh-bar{width:4px;height:26px;border-radius:3px;flex-shrink:0;}
.sh-txt{font-size:0.92rem;font-weight:800;letter-spacing:0.10em;text-transform:uppercase;color:#E2E8F0;text-shadow:0 0 20px rgba(56,189,248,0.26);}
</style>""", unsafe_allow_html=True)

# ─── DATOS ────────────────────────────────────────────────────────────────────
@st.cache_data(ttl=30)
def cargar(mtime):
    df = pd.read_excel(ARCHIVO, sheet_name="TAREAS")
    for col in ["FECHA_COMPROMISO", "FECHA_INICIO", "FECHA_CIERRE"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce").astype("datetime64[us]")
    defaults = {"PRIORIDAD": "Media", "ESTADO": "Pendiente",
                "PROYECTO": "GENERAL", "AREA": "Trabajo", "NOTAS": ""}
    for col, val in defaults.items():
        if col not in df.columns:
            df[col] = val
        df[col] = df[col].fillna(val)
    df = df[df["TAREA"].astype(str).str.strip() != ""].copy()
    df["PRIO_ORD"] = df["PRIORIDAD"].map(PRIO_ORD).fillna(99)
    return df

mtime  = ARCHIVO.stat().st_mtime if ARCHIVO.exists() else 0
df_raw = cargar(mtime)

# ─── SIDEBAR ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown(
        '<div style="font-size:0.60rem;font-weight:800;letter-spacing:0.26em;text-transform:uppercase;'
        'color:#38BDF8;margin-bottom:2px;text-shadow:0 0 10px rgba(56,189,248,0.40);">'
        'P&C · CONSTRUCTORA LONDRES</div>'
        '<div style="font-size:1.70rem;font-weight:900;color:#F8FAFC;line-height:1.0;margin-bottom:18px;">'
        '✅ TAREAS</div>', unsafe_allow_html=True)

    def msel(label, opciones, key):
        st.markdown(
            f'<div style="font-size:0.58rem;font-weight:800;letter-spacing:0.12em;'
            f'text-transform:uppercase;color:#475569;margin:10px 0 5px;">{label}</div>',
            unsafe_allow_html=True)
        return st.multiselect("", opciones, default=opciones, key=key, label_visibility="collapsed")

    projs   = sorted(df_raw["PROYECTO"].dropna().unique())
    areas   = sorted(df_raw["AREA"].dropna().unique())
    prios   = [p for p in ["Crítica","Alta","Media","Baja"] if p in df_raw["PRIORIDAD"].values]
    estados = [e for e in ["Pendiente","En Proceso","Esperando Terceros"]
               if e in df_raw["ESTADO"].values]

    sel_proj = msel("Proyecto", projs,   "f_proj")
    sel_area = msel("Área",     areas,   "f_area")
    sel_prio = msel("Prioridad",prios,   "f_prio")
    sel_est  = msel("Estado",   estados, "f_est")

    st.divider()
    st.markdown(
        f'<div style="font-size:0.58rem;color:#334155;text-align:center;">'
        f'📅 {HOY.strftime("%d de %B de %Y")}</div>'
        f'<div style="font-size:0.52rem;color:#1E293B;text-align:center;margin-top:4px;">'
        f'Edita en Excel → git push para actualizar</div>',
        unsafe_allow_html=True)

# ─── FILTRO ACTIVO ────────────────────────────────────────────────────────────
df = df_raw[
    df_raw["PROYECTO"].isin(sel_proj) &
    df_raw["AREA"].isin(sel_area) &
    df_raw["PRIORIDAD"].isin(sel_prio) &
    df_raw["ESTADO"].isin(sel_est)
].copy()

# ─── KPIS GLOBALES ────────────────────────────────────────────────────────────
activas  = df_raw[df_raw["ESTADO"] != "Completada"]
comps    = df_raw[df_raw["ESTADO"] == "Completada"]
vencidas = activas[activas["FECHA_COMPROMISO"].notna() &
                    (activas["FECHA_COMPROMISO"] < pd.Timestamp(HOY))]
ini_sem  = HOY - timedelta(days=HOY.weekday())
c_sem    = comps[comps["FECHA_CIERRE"].notna() &
                  (comps["FECHA_CIERRE"] >= pd.Timestamp(ini_sem))]
c_mes    = comps[comps["FECHA_CIERRE"].notna() &
                  (comps["FECHA_CIERRE"].dt.month == HOY.month) &
                  (comps["FECHA_CIERRE"].dt.year  == HOY.year)]

n_act  = len(activas)
n_pend = int((df_raw["ESTADO"] == "Pendiente").sum())
n_proc = int((df_raw["ESTADO"] == "En Proceso").sum())
n_esp  = int((df_raw["ESTADO"] == "Esperando Terceros").sum())
n_comp = len(comps)
n_venc = len(vencidas)
n_csem = len(c_sem)
n_cmes = len(c_mes)
_both  = comps.dropna(subset=["FECHA_CIERRE","FECHA_INICIO"])
avg_dias = (_both["FECHA_CIERRE"] - _both["FECHA_INICIO"]).dt.days.mean() if not _both.empty else None
tasa_sem = round(n_csem / max(n_csem + n_venc, 1) * 100) if (n_csem + n_venc) > 0 else 0

# ─── HELPERS ──────────────────────────────────────────────────────────────────
def kpi_card(label, valor, sub=None, color="#38BDF8"):
    sub_h = f'<div class="tk-sub" style="color:{color};">{sub}</div>' if sub else ""
    return (
        f'<div class="tk-kpi" style="background:linear-gradient(135deg,{color}18,{color}04);'
        f'border:1px solid {color}30;">'
        f'<div class="tk-lbl">{label}</div>'
        f'<div class="tk-val" style="color:{color};text-shadow:0 0 14px {color}55;">{valor}</div>'
        f'{sub_h}</div>'
    )

def seccion(icon, titulo, color="#38BDF8"):
    st.markdown(
        f'<div class="sh">'
        f'<div class="sh-bar" style="background:linear-gradient(180deg,{color},{color}44);'
        f'box-shadow:0 0 10px {color}88;"></div>'
        f'<div class="sh-txt">{icon} {titulo}</div></div>',
        unsafe_allow_html=True)

def fecha_tag(dt):
    if pd.isna(dt): return ""
    d = dt.date()
    if d < HOY:    c, i = "#FF4757", "🔴"
    elif d == HOY: c, i = "#FFB300", "🟡"
    else:          c, i = "#64748B", "📅"
    return f'<div class="tc-date" style="color:{c};">{i} Vence {d.strftime("%d/%m/%Y")}</div>'

def card_html(row):
    p  = str(row.get("PRIORIDAD", "Media"))
    pc = PRIO_CLR.get(p, "#64748B")
    pr = str(row.get("PROYECTO", ""))
    ar = str(row.get("AREA", ""))
    nm = str(row["TAREA"])
    dt = row.get("FECHA_COMPROMISO", pd.NaT)
    venc = pd.notna(dt) and dt.date() < HOY
    bg   = "rgba(255,71,87,0.07)" if venc else "rgba(13,21,38,0.80)"
    brd  = "rgba(255,71,87,0.32)" if venc else f"{pc}28"
    return (
        f'<div class="tc" style="background:{bg};border-color:{brd};">'
        f'<div class="tc-tags">'
        f'<span class="tc-tag" style="background:{pc}22;color:{pc};border:1px solid {pc}40;">'
        f'{PRIO_ICO.get(p,"")} {p}</span>'
        f'<span class="tc-tag" style="background:rgba(56,189,248,0.08);color:#38BDF8;'
        f'border:1px solid rgba(56,189,248,0.18);">{pr}</span>'
        f'</div>'
        f'<div class="tc-title">{nm}</div>'
        f'<div class="tc-tags">'
        f'<span class="tc-tag" style="background:rgba(100,116,139,0.12);color:#475569;">{ar}</span>'
        f'</div>'
        f'{fecha_tag(dt)}</div>'
    )

_TT = {
    "trigger": "item",
    "backgroundColor": "rgba(6,11,21,0.97)",
    "borderColor": "rgba(56,189,248,0.22)",
    "borderRadius": 12,
    "textStyle": {"color": "#F1F5F9", "fontSize": 11},
}

def ec(option, height=260):
    opt_str = json.dumps(option, ensure_ascii=False)
    html = f"""<!DOCTYPE html><html><head>
<script src="https://cdn.jsdelivr.net/npm/echarts@5.4.3/dist/echarts.min.js"></script>
</head><body style="margin:0;padding:0;background:transparent;">
<div id="c" style="width:100%;height:{height}px;"></div>
<script>
var ch = echarts.init(document.getElementById('c'),null,{{renderer:'svg',width:'auto',height:{height}}});
ch.setOption({opt_str});
window.addEventListener('resize',function(){{ch.resize();}});
</script></body></html>"""
    components.html(html, height=height + 8, scrolling=False)

# ══════════════════════════════════════════════════════════════════════════════
# TABS
# ══════════════════════════════════════════════════════════════════════════════
tab1, tab2, tab3, tab4 = st.tabs(["📊 Dashboard", "📋 Kanban", "📄 Lista", "✅ Historial"])

# ─────────────────────── TAB 1: DASHBOARD ────────────────────────────────────
with tab1:
    # Banner
    st.markdown(
        '<div style="background:linear-gradient(90deg,rgba(56,189,248,0.08),transparent);'
        'border:1px solid rgba(56,189,248,0.15);border-radius:18px;padding:16px 22px;margin-bottom:16px;'
        'display:flex;align-items:center;gap:18px;">'
        '<div>'
        '<div style="font-size:0.60rem;font-weight:800;letter-spacing:0.26em;color:#38BDF8;margin-bottom:3px;">'
        'PLANIFICACIÓN Y CONTROL · CONSTRUCTORA LONDRES</div>'
        '<div style="font-size:1.85rem;font-weight:900;color:#F8FAFC;">✅ GESTIÓN DE TAREAS</div>'
        '</div>'
        f'<div style="margin-left:auto;text-align:right;">'
        f'<div style="font-size:0.55rem;color:#475569;letter-spacing:0.10em;">HOY</div>'
        f'<div style="font-size:0.95rem;font-weight:800;color:#94A3B8;">{HOY.strftime("%d/%m/%Y")}</div>'
        f'</div></div>', unsafe_allow_html=True)

    # KPI fila 1 — actividad
    c1,c2,c3,c4,c5,c6 = st.columns(6)
    with c1: st.markdown(kpi_card("TOTAL ACTIVAS",  n_act,  color="#38BDF8"), unsafe_allow_html=True)
    with c2: st.markdown(kpi_card("PENDIENTES",     n_pend, color="#64748B"), unsafe_allow_html=True)
    with c3: st.markdown(kpi_card("EN PROCESO",     n_proc, color="#38BDF8"), unsafe_allow_html=True)
    with c4: st.markdown(kpi_card("ESP. TERCEROS",  n_esp,  color="#FFB300"), unsafe_allow_html=True)
    with c5: st.markdown(kpi_card("VENCIDAS", n_venc,
                                   "requieren atención" if n_venc else "al día",
                                   color="#FF4757" if n_venc else "#23D160"), unsafe_allow_html=True)
    with c6: st.markdown(kpi_card("COMPLETADAS", n_comp, color="#23D160"), unsafe_allow_html=True)

    st.markdown('<div style="height:8px;"></div>', unsafe_allow_html=True)

    # KPI fila 2 — productividad
    c7,c8,c9,c10 = st.columns(4)
    with c7:  st.markdown(kpi_card("COMP. ESTA SEMANA", n_csem, color="#23D160"), unsafe_allow_html=True)
    with c8:  st.markdown(kpi_card("COMP. ESTE MES",    n_cmes, color="#23D160"), unsafe_allow_html=True)
    with c9:
        avg_str = f"{avg_dias:.0f} días" if avg_dias is not None else "—"
        st.markdown(kpi_card("T° PROM. CIERRE", avg_str, color="#38BDF8"), unsafe_allow_html=True)
    with c10:
        tc = "#23D160" if tasa_sem >= 70 else "#FFB300" if tasa_sem >= 40 else "#FF4757"
        st.markdown(kpi_card("TASA SEMANAL", f"{tasa_sem}%",
                              "completadas vs vencidas", color=tc), unsafe_allow_html=True)

    st.markdown('<div style="height:6px;"></div>', unsafe_allow_html=True)
    seccion("📈", "ANÁLISIS DE CARGA", color="#38BDF8")

    ch1, ch2, ch3 = st.columns(3)

    with ch1:
        prio_d = df_raw[df_raw["ESTADO"] != "Completada"]["PRIORIDAD"].value_counts()
        ec({
            "backgroundColor": "transparent",
            "tooltip": _TT,
            "title": {"text": "Por Prioridad", "left": "center", "top": "2%",
                      "textStyle": {"color": "#CBD5E1", "fontSize": 12, "fontWeight": "bold"}},
            "legend": {"bottom": "2%", "left": "center",
                       "textStyle": {"color": "#64748B", "fontSize": 10}},
            "series": [{"type": "pie", "radius": ["42%", "68%"], "center": ["50%", "52%"],
                        "avoidLabelOverlap": True,
                        "label": {"show": False},
                        "emphasis": {"label": {"show": True, "fontSize": 13,
                                               "fontWeight": "bold", "color": "#F1F5F9"}},
                        "data": [{"name": k, "value": int(v),
                                  "itemStyle": {"color": PRIO_CLR.get(k, "#64748B"),
                                                "borderRadius": 4, "borderWidth": 2,
                                                "borderColor": "#060B15"}}
                                 for k, v in prio_d.items()]}]
        })

    with ch2:
        est_d = df_raw[df_raw["ESTADO"] != "Completada"]["ESTADO"].value_counts()
        ec({
            "backgroundColor": "transparent",
            "tooltip": _TT,
            "title": {"text": "Por Estado", "left": "center", "top": "2%",
                      "textStyle": {"color": "#CBD5E1", "fontSize": 12, "fontWeight": "bold"}},
            "legend": {"bottom": "2%", "left": "center",
                       "textStyle": {"color": "#64748B", "fontSize": 10}},
            "series": [{"type": "pie", "radius": ["42%", "68%"], "center": ["50%", "52%"],
                        "avoidLabelOverlap": True,
                        "label": {"show": False},
                        "emphasis": {"label": {"show": True, "fontSize": 13,
                                               "fontWeight": "bold", "color": "#F1F5F9"}},
                        "data": [{"name": k, "value": int(v),
                                  "itemStyle": {"color": EST_CLR.get(k, "#64748B"),
                                                "borderRadius": 4, "borderWidth": 2,
                                                "borderColor": "#060B15"}}
                                 for k, v in est_d.items()]}]
        })

    with ch3:
        proj_d = (df_raw[df_raw["ESTADO"] != "Completada"]
                  .groupby("PROYECTO").size().sort_values(ascending=True))
        ec({
            "backgroundColor": "transparent",
            "tooltip": {"trigger": "axis", "backgroundColor": "rgba(6,11,21,0.97)",
                        "borderColor": "rgba(56,189,248,0.22)", "borderRadius": 12,
                        "textStyle": {"color": "#F1F5F9", "fontSize": 11}},
            "title": {"text": "Por Proyecto", "left": "center", "top": "2%",
                      "textStyle": {"color": "#CBD5E1", "fontSize": 12, "fontWeight": "bold"}},
            "grid": {"left": "110px", "right": "40px", "top": "38px", "bottom": "18px"},
            "xAxis": {"type": "value",
                      "axisLabel": {"color": "#64748B", "fontSize": 9},
                      "splitLine": {"lineStyle": {"color": "rgba(255,255,255,0.045)"}},
                      "axisLine": {"lineStyle": {"color": "#1E293B"}}},
            "yAxis": {"type": "category", "data": proj_d.index.tolist(),
                      "axisLabel": {"color": "#94A3B8", "fontSize": 10},
                      "axisLine": {"lineStyle": {"color": "#1E293B"}},
                      "splitLine": {"show": False}},
            "series": [{"type": "bar", "barMaxWidth": "26px",
                        "data": [int(v) for v in proj_d.values],
                        "label": {"show": True, "position": "right",
                                  "color": "#94A3B8", "fontSize": 10},
                        "itemStyle": {
                            "borderRadius": [0, 6, 6, 0],
                            "color": {"type": "linear", "x": 0, "y": 0, "x2": 1, "y2": 0,
                                      "colorStops": [{"offset": 0, "color": "rgba(56,189,248,0.90)"},
                                                     {"offset": 1, "color": "rgba(56,189,248,0.25)"}]}
                        }}]
        })

# ─────────────────────── TAB 2: KANBAN ───────────────────────────────────────
with tab2:
    buscar_k = st.text_input("🔍", placeholder="Filtrar tareas...", key="kb_s",
                              label_visibility="collapsed")
    df_kb = df_raw[df_raw["ESTADO"].isin(["Pendiente","En Proceso","Esperando Terceros"])].copy()
    if buscar_k.strip():
        df_kb = df_kb[df_kb["TAREA"].str.contains(buscar_k.strip(), case=False, na=False)]
    df_kb = df_kb.sort_values("PRIO_ORD")

    COLS_KB = [("Pendiente", "#64748B"), ("En Proceso", "#38BDF8"), ("Esperando Terceros", "#FFB300")]
    c_kb1, c_kb2, c_kb3 = st.columns(3)
    for col_st, (est, clr) in zip([c_kb1, c_kb2, c_kb3], COLS_KB):
        grp = df_kb[df_kb["ESTADO"] == est]
        ico = EST_ICO[est]
        cards = "".join(card_html(r) for _, r in grp.iterrows())
        with col_st:
            st.markdown(
                f'<div class="kb-wrap" style="border-color:{clr}22;">'
                f'<div class="kb-hdr" style="color:{clr};">'
                f'{ico} {est}'
                f'<span class="kb-n" style="background:{clr};">&nbsp;{len(grp)}&nbsp;</span>'
                f'</div>{cards}</div>',
                unsafe_allow_html=True)

# ─────────────────────── TAB 3: LISTA ────────────────────────────────────────
with tab3:
    buscar_l = st.text_input("🔍", placeholder="Buscar tarea...", key="ls_s",
                              label_visibility="collapsed")
    df_l = df.copy()
    if buscar_l.strip():
        df_l = df_l[df_l["TAREA"].str.contains(buscar_l.strip(), case=False, na=False)]
    df_l = df_l.sort_values(["PRIO_ORD", "FECHA_COMPROMISO"])

    st.markdown(
        f'<div style="font-size:0.62rem;color:#475569;margin-bottom:6px;">'
        f'{len(df_l)} tarea(s) · filtros aplicados del sidebar</div>',
        unsafe_allow_html=True)

    cols_show = [c for c in ["TAREA","PROYECTO","AREA","PRIORIDAD","ESTADO",
                              "FECHA_COMPROMISO","FECHA_CIERRE","NOTAS"] if c in df_l.columns]
    df_ld = df_l[cols_show].copy()
    for c in ["FECHA_COMPROMISO", "FECHA_CIERRE"]:
        if c in df_ld.columns:
            df_ld[c] = df_ld[c].dt.strftime("%d/%m/%Y").fillna("—")
    st.dataframe(df_ld.fillna("—"), use_container_width=True, hide_index=True,
                 height=min(650, 50 + len(df_l) * 36))

# ─────────────────────── TAB 4: HISTORIAL ────────────────────────────────────
with tab4:
    df_h = df_raw[df_raw["ESTADO"] == "Completada"].sort_values("FECHA_CIERRE", ascending=False)
    h1, h2, h3, h4 = st.columns(4)
    with h1: st.markdown(kpi_card("TOTAL COMPLETADAS", len(df_h), color="#23D160"), unsafe_allow_html=True)
    with h2: st.markdown(kpi_card("ESTA SEMANA", n_csem, color="#23D160"), unsafe_allow_html=True)
    with h3: st.markdown(kpi_card("ESTE MES",    n_cmes, color="#23D160"), unsafe_allow_html=True)
    with h4:
        avg_s = f"{avg_dias:.0f} días" if avg_dias is not None else "—"
        st.markdown(kpi_card("T° PROM. CIERRE", avg_s, color="#38BDF8"), unsafe_allow_html=True)

    st.markdown('<div style="height:8px;"></div>', unsafe_allow_html=True)
    cols_h = [c for c in ["TAREA","PROYECTO","AREA","PRIORIDAD","FECHA_COMPROMISO","FECHA_CIERRE"]
              if c in df_h.columns]
    df_hd = df_h[cols_h].copy()
    for c in ["FECHA_COMPROMISO", "FECHA_CIERRE"]:
        if c in df_hd.columns:
            df_hd[c] = df_hd[c].dt.strftime("%d/%m/%Y").fillna("—")
    st.dataframe(df_hd.fillna("—"), use_container_width=True, hide_index=True)

# ─── AUTO-REFRESH ─────────────────────────────────────────────────────────────
from streamlit_autorefresh import st_autorefresh
st_autorefresh(interval=60_000, key="tareas_auto")
