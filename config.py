# ─── CONFIG GLOBAL ────────────────────────────────────────────────────────────
from pathlib import Path
from datetime import date
import pandas as pd

ARCHIVO = Path(__file__).parent / "TAREAS_JOCHOA.xlsx"
HOY     = date.today()
HOY_TS  = pd.Timestamp(HOY)

# ─── PALETA SEMÁNTICA (fija — no cambia con tema) ─────────────────────────────
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

# ─── TEMAS ────────────────────────────────────────────────────────────────────
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

# ─── GITHUB ───────────────────────────────────────────────────────────────────
_GH_REPO   = "Jochoa27/jochoa-tareas"
_GH_FILE   = "TAREAS_JOCHOA.xlsx"
_GH_BRANCH = "main"

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

# ─── ECHARTS TOOLTIP BASE ─────────────────────────────────────────────────────
_TT = {
    "trigger": "item",
    "backgroundColor": "rgba(6,11,21,0.97)",
    "borderColor":     "rgba(56,189,248,0.22)",
    "borderRadius":    12,
    "textStyle":       {"color": "#F1F5F9", "fontSize": 11},
}
_TT_AXIS = {**_TT, "trigger": "axis"}
