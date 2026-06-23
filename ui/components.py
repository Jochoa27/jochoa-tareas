import json
import streamlit as st
import streamlit.components.v1 as components

from config import C_CIAN

# ─── SECTION HEADER ───────────────────────────────────────────────────────────
def seccion(icon, titulo, color=C_CIAN):
    st.markdown(
        f'<div class="sh">'
        f'<div class="sh-bar" style="background:linear-gradient(180deg,{color},{color}44);'
        f'box-shadow:0 0 10px {color}88;"></div>'
        f'<div class="sh-txt">{icon} {titulo}</div></div>',
        unsafe_allow_html=True)

# ─── KPI CARD ─────────────────────────────────────────────────────────────────
def kpi(label, valor, sub=None, color=C_CIAN):
    sub_h = f'<div class="kpi-sub" style="color:{color};">{sub}</div>' if sub else ""
    return (
        f'<div class="kpi-card" style="background:linear-gradient(135deg,{color}18,{color}04);'
        f'border:1px solid {color}30;">'
        f'<div class="kpi-lbl">{label}</div>'
        f'<div class="kpi-val" style="color:{color};text-shadow:0 0 14px {color}55;">{valor}</div>'
        f'{sub_h}</div>'
    )

# ─── BADGES Y CHIPS ───────────────────────────────────────────────────────────
def badge_estado(texto, color):
    return (f'<span class="estado-badge" style="background:{color}18;'
            f'border:1px solid {color}44;color:{color};">{texto}</span>')

def chip(texto, color):
    return (f'<span class="tc-tag" style="background:{color}16;'
            f'color:{color};border:1px solid {color}30;">{texto}</span>')

# ─── ECHARTS RENDERER ─────────────────────────────────────────────────────────
def ec(option, height=260):
    s = json.dumps(option, ensure_ascii=False)
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

# ─── ECHARTS AXIS STYLE ───────────────────────────────────────────────────────
def _axis(color="#38BDF8"):
    return {"axisLine": {"lineStyle": {"color": "#1E293B"}},
            "splitLine": {"lineStyle": {"color": "rgba(255,255,255,0.04)"}},
            "axisLabel": {"color": "#64748B", "fontSize": 10}}
