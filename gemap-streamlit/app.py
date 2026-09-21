"""
GEMAP — Panel de seguimiento de empresas de la gestoría.
Streamlit + SQLite. Sin dependencias externas más allá de streamlit.
"""
import calendar as cal
from datetime import date, datetime, timedelta

import streamlit as st

import db

# ---------------------------------------------------------------- Setup
st.set_page_config(
    page_title="GEMAP — Panel de empresas",
    page_icon="assets/logo.png",
    layout="wide",
    initial_sidebar_state="expanded",
)
db.init_db()

BRAND_1 = "#7a1224"
BRAND_2 = "#b8213a"
BRAND_SOFT = "#f7e3e6"
WARN = "#a1483a"
WARN_SOFT = "#f6e2de"
INFO = "#4a6b8a"
INFO_SOFT = "#e4ecf3"
DONE = "#4c7a63"

TIPO_LABEL = {"contable": "Contable", "informacion": "Información", "reunion": "Reunión"}

st.markdown(f"""
<style>
  html, body, [class*="css"] {{ font-family: 'Source Sans Pro', sans-serif; }}
  h1, h2, h3 {{ font-family: 'Georgia', serif; }}
  .gemap-alert {{ padding: 10px 14px; border-radius: 6px; margin-bottom: 8px; font-size: 0.92rem; }}
  .gemap-alert.warn {{ background: {WARN_SOFT}; color: {WARN}; }}
  .gemap-alert.info {{ background: {INFO_SOFT}; color: {INFO}; }}
  .gemap-tag {{ display:inline-block; font-size:0.72rem; padding:1px 8px; border-radius:20px; border:1px solid #ccc; margin-right:6px;}}
  .gemap-tag.contable {{ color: {BRAND_1}; border-color: {BRAND_1}; }}
  .gemap-tag.informacion {{ color: {WARN}; border-color: {WARN}; }}
  .gemap-tag.reunion {{ color: {INFO}; border-color: {INFO}; }}
  .gemap-profile {{ background:#fff; border:1px solid #eee; border-radius:8px; padding:16px 20px; margin-bottom:20px; }}
  section[data-testid="stSidebar"] {{ background: #ffffff; }}
  .gemap-cal-day {{ border:1px solid #eee; border-radius:6px; padding:6px; min-height:64px; font-size:0.8rem; }}
  .gemap-cal-day.today {{ background: {BRAND_SOFT}; }}
  .gemap-cal-day.other {{ color:#bbb; }}
</style>
""", unsafe_allow_html=True)


def days_until(fecha_str):
    if not fecha_str:
        return None
    try:
        d = datetime.strptime(fecha_str[:10], "%Y-%m-%d").date()
    except ValueError:
        return None
    return (d - date.today()).days


def fmt_date(fecha_str):
    """Formato español dd/mm/aaaa, sin depender del locale del servidor."""
    if not fecha_str:
        return ""
    try:
        d = datetime.strptime(fecha_str[:10], "%Y-%m-%d").date()
    except ValueError:
        return ""
    return d.strftime("%d/%m/%Y")


MESES_ES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
            "agosto", "septiembre", "octubre", "noviembre", "diciembre"]


def outlook_link(descripcion, empresa_nombre, fecha_str, tipo):
    """Enlace que abre Outlook Web con el evento precargado (sin necesitar conectar cuenta)."""
    if not fecha_str:
        return None
    from urllib.parse import quote
    start = f"{fecha_str[:10]}T09:00:00"
    end = f"{fecha_str[:10]}T09:30:00"
    subject = f"{empresa_nombre}: {descripcion}"
    body = f"Recordatorio GEMAP ({TIPO_LABEL.get(tipo, tipo)}) — {empresa_nombre}"
    return (
        "https://outlook.office.com/calendar/0/deeplink/compose"
        f"?path=/calendar/action/compose&rru=addevent"
        f"&subject={quote(subject)}&startdt={quote(start)}&enddt={quote(end)}"
        f"&body={quote(body)}&location={quote(empresa_nombre)}"
    )


# ---------------------------------------------------------------- Sidebar
with st.sidebar:
    col_logo, col_title = st.columns([1, 3])
    with col_logo:
        st.image("assets/logo.png", width=44)
    with col_title:
        st.markdown("### GEMAP")
        st.caption("Gestión de empresas · Panel")

    empresas = db.list_empresas()
    tareas_all = db.list_tareas()

    view = st.radio("Vista", ["Resumen", "Calendario"], label_visibility="collapsed")

    st.markdown("**Empresas**")
    counts = {}
    for t in tareas_all:
        if t["estado"] != "hecho":
            counts[t["empresa_id"]] = counts.get(t["empresa_id"], 0) + 1

    name_to_id = {e["nombre"]: e["id"] for e in empresas}
    buscada = st.selectbox(
        "Buscar empresa",
        options=sorted(name_to_id.keys(), key=lambda n: n.lower()),
        index=None,
        placeholder="Escribe para buscar...",
        label_visibility="collapsed",
    )
    if buscada:
        st.session_state["selected_empresa"] = name_to_id[buscada]
        view = "Empresa"

    selected_empresa = st.session_state.get("selected_empresa")
    for e in empresas:
        label = f'{e["nombre"]}  ·  {counts.get(e["id"], 0)}'
        if st.button(label, key=f'company_{e["id"]}', use_container_width=True):
            st.session_state["selected_empresa"] = e["id"]
            view = "Empresa"
            st.rerun()

    with st.expander("+ Añadir empresa"):
        with st.form("new_company_form", clear_on_submit=True):
            new_name = st.text_input("Nombre de la empresa")
            submitted = st.form_submit_button("Crear")
            if submitted and new_name.strip():
                eid = db.ensure_empresa(new_name.strip())
                st.session_state["selected_empresa"] = eid
                st.rerun()

if st.session_state.get("selected_empresa") and view not in ("Resumen", "Calendario"):
    view = "Empresa"


# ---------------------------------------------------------------- Alerts
def compute_alerts(tareas, empresas_by_id):
    alerts = []
    for t in tareas:
        if t["estado"] == "hecho":
            continue
        du = days_until(t["fecha"])
        nombre = empresas_by_id.get(t["empresa_id"], {}).get("nombre", t["empresa_id"])
        if t["tipo"] == "reunion":
            if du is not None and 0 <= du <= 7:
                tag = "Hoy" if du == 0 else ("Mañana" if du == 1 else f"En {du} días")
                alerts.append(("info", tag, f'Reunión con {nombre}: {t["descripcion"]}', du))
            elif du is not None and du < 0:
                alerts.append(("warn", "Pasada", f'Reunión con {nombre} ({fmt_date(t["fecha"])}): {t["descripcion"]}', du))
        else:
            if du is not None and du < 0:
                alerts.append(("warn", "Vencido", f'{nombre} — {t["descripcion"]} ({fmt_date(t["fecha"])})', du))
            elif du is not None and du <= 3:
                tag = "Hoy" if du == 0 else f"En {du} días"
                alerts.append(("info", tag, f'{nombre} — {t["descripcion"]}', du))
    alerts.sort(key=lambda a: a[3])
    return alerts


def render_alerts(alerts):
    for kind, tag, text, _ in alerts:
        st.markdown(f'<div class="gemap-alert {kind}"><b>{tag}:</b> {text}</div>', unsafe_allow_html=True)


def render_task_row(t, empresas_by_id, show_company=False):
    nombre = empresas_by_id.get(t["empresa_id"], {}).get("nombre", t["empresa_id"])
    is_done = t["estado"] == "hecho"
    c1, c2, c3 = st.columns([0.06, 0.82, 0.12])
    with c1:
        checked = st.checkbox("", value=is_done, key=f'chk_{t["id"]}', label_visibility="collapsed")
        if checked != is_done:
            db.toggle_tarea(t["id"], checked)
            st.rerun()
    with c2:
        desc = t["descripcion"]
        if is_done:
            st.markdown(f"~~{desc}~~")
        else:
            st.markdown(desc)
        meta = f'<span class="gemap-tag {t["tipo"]}">{TIPO_LABEL.get(t["tipo"], t["tipo"])}</span>'
        if t["fecha"]:
            meta += f' <span style="font-size:0.75rem;color:#888;">{fmt_date(t["fecha"])}</span>'
        if show_company:
            meta += f' <span style="font-size:0.75rem;color:#888;">· {nombre}</span>'
        st.markdown(meta, unsafe_allow_html=True)
    with c3:
        if t["fecha"] and not is_done:
            link = outlook_link(t["descripcion"], nombre, t["fecha"], t["tipo"])
            st.link_button("Outlook", link, use_container_width=True)


def render_quick_add(fixed_empresa_id=None, empresa_names=None):
    st.markdown("---")
    st.markdown("##### Añadir algo rápido")
    with st.form(f"quickadd_{fixed_empresa_id or 'global'}", clear_on_submit=True):
        cols = st.columns([2, 3, 2, 2, 1])
        empresa_nombre = fixed_empresa_id
        if not fixed_empresa_id:
            empresa_nombre = cols[0].text_input("Empresa", label_visibility="collapsed", placeholder="Empresa")
        else:
            cols[0].markdown(f"**{empresa_names.get(fixed_empresa_id)}**")
        descripcion = cols[1].text_input("Descripción", label_visibility="collapsed", placeholder="¿Qué falta?")
        tipo = cols[2].selectbox("Tipo", ["contable", "informacion", "reunion"],
                                  format_func=lambda x: TIPO_LABEL[x], label_visibility="collapsed")
        fecha = cols[3].date_input("Fecha", value=None, format="DD/MM/YYYY", label_visibility="collapsed")
        submitted = cols[4].form_submit_button("Añadir")
        if submitted and descripcion.strip():
            eid = fixed_empresa_id if fixed_empresa_id else db.ensure_empresa(empresa_nombre.strip())
            if eid:
                fecha_str = fecha.isoformat() if fecha else None
                db.add_tarea(eid, descripcion.strip(), tipo, fecha_str)
                st.rerun()


empresas_by_id = {e["id"]: e for e in empresas}

# ---------------------------------------------------------------- Views
if view == "Resumen":
    st.title("Resumen")
    alerts = compute_alerts(tareas_all, empresas_by_id)
    if alerts:
        render_alerts(alerts)

    pending = [t for t in tareas_all if t["estado"] != "hecho"]
    if not pending and not alerts:
        st.info("Nada por aquí todavía. Dime qué falta por hacer, pedir o qué reunión tienes en cada empresa.")
    else:
        by_company = {}
        for t in sorted(pending, key=lambda t: t["fecha"] or "9999"):
            by_company.setdefault(t["empresa_id"], []).append(t)
        for eid in sorted(by_company, key=lambda x: empresas_by_id.get(x, {}).get("nombre", x)):
            st.markdown(f'#### {empresas_by_id.get(eid, {}).get("nombre", eid)}')
            for t in by_company[eid]:
                render_task_row(t, empresas_by_id)

    render_quick_add()

elif view == "Calendario":
    st.title("Calendario")
    if "cal_month" not in st.session_state:
        st.session_state["cal_month"] = date.today().replace(day=1)
    cur = st.session_state["cal_month"]

    c1, c2, c3 = st.columns([1, 4, 1])
    if c1.button("◀"):
        st.session_state["cal_month"] = (cur.replace(day=1) - timedelta(days=1)).replace(day=1)
        st.rerun()
    c2.markdown(f"### {MESES_ES[cur.month - 1]} {cur.year}")
    if c3.button("▶"):
        nxt_month = cur.month % 12 + 1
        nxt_year = cur.year + (1 if cur.month == 12 else 0)
        st.session_state["cal_month"] = date(nxt_year, nxt_month, 1)
        st.rerun()

    reuniones = [t for t in tareas_all if t["tipo"] == "reunion" and t["fecha"]]
    by_date = {}
    for t in reuniones:
        by_date.setdefault(t["fecha"][:10], []).append(t)

    weeks = cal.Calendar(firstweekday=0).monthdatescalendar(cur.year, cur.month)
    dow_labels = ["Lun", "Mar", "Mié", "Jue", "Vie", "Sáb", "Dom"]
    header_cols = st.columns(7)
    for i, lbl in enumerate(dow_labels):
        header_cols[i].markdown(f"<div style='text-align:center;font-size:0.75rem;color:#999;'>{lbl}</div>",
                                 unsafe_allow_html=True)

    if "cal_selected" not in st.session_state:
        st.session_state["cal_selected"] = date.today().isoformat()

    for week in weeks:
        cols = st.columns(7)
        for i, d in enumerate(week):
            iso = d.isoformat()
            is_other = d.month != cur.month
            is_today = iso == date.today().isoformat()
            n_items = len(by_date.get(iso, []))
            css_class = "gemap-cal-day" + (" other" if is_other else "") + (" today" if is_today else "")
            dots = "●" * min(n_items, 3)
            with cols[i]:
                if st.button(f"{d.day}\n{dots}", key=f"cal_{iso}", use_container_width=True):
                    st.session_state["cal_selected"] = iso
                    st.rerun()

    sel = st.session_state["cal_selected"]
    st.markdown(f"#### {fmt_date(sel)}")
    day_items = by_date.get(sel, [])
    if not day_items:
        st.caption("Sin reuniones este día.")
    else:
        for t in day_items:
            render_task_row(t, empresas_by_id, show_company=True)

elif view == "Empresa" and st.session_state.get("selected_empresa"):
    eid = st.session_state["selected_empresa"]
    e = db.get_empresa(eid)
    if not e:
        st.warning("Esta empresa ya no existe.")
    else:
        st.title(e["nombre"])
        with st.container():
            st.markdown('<div class="gemap-profile">', unsafe_allow_html=True)
            edit_mode = st.session_state.get(f"edit_{eid}", False)
            if edit_mode:
                with st.form(f"edit_form_{eid}"):
                    actividad = st.text_input("A qué se dedican", value=e.get("actividad", ""))
                    responsable = st.text_input("Responsable", value=e["responsable"])
                    c1, c2 = st.columns(2)
                    analitica = c1.checkbox("Lleva analítica", value=bool(e.get("analitica", 0)))
                    sii = c2.checkbox("Está en el SII", value=bool(e.get("sii", 0)))
                    particularidades = st.text_area("Apreciaciones / notas", value=e["particularidades"])
                    c3, c4 = st.columns(2)
                    if c3.form_submit_button("Guardar"):
                        db.update_empresa(eid, responsable, particularidades, actividad, analitica, sii)
                        st.session_state[f"edit_{eid}"] = False
                        st.rerun()
                    if c4.form_submit_button("Cancelar"):
                        st.session_state[f"edit_{eid}"] = False
                        st.rerun()
            else:
                c1, c2 = st.columns(2)
                c1.markdown(f"**A qué se dedican**  \n{e.get('actividad') or '_Sin especificar_'}")
                c2.markdown(f"**Responsable**  \n{e['responsable'] or '_Sin asignar_'}")
                c3, c4 = st.columns(2)
                c3.markdown(f"**Analítica**  \n{'Sí' if e.get('analitica') else 'No'}")
                c4.markdown(f"**SII**  \n{'Sí' if e.get('sii') else 'No'}")
                st.markdown(f"**Apreciaciones**  \n{e['particularidades'] or '_Sin notas_'}")
                if st.button("Editar", key=f"editbtn_{eid}"):
                    st.session_state[f"edit_{eid}"] = True
                    st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)

        tareas_e = [t for t in tareas_all if t["empresa_id"] == eid]
        pending = [t for t in tareas_e if t["estado"] != "hecho"]
        if not pending:
            st.info(f"Todo al día para {e['nombre']}.")
        else:
            for t in sorted(pending, key=lambda t: t["fecha"] or "9999"):
                render_task_row(t, empresas_by_id)

        render_quick_add(fixed_empresa_id=eid, empresa_names={eid: e["nombre"]})
else:
    st.title("Resumen")
    st.info("Selecciona una vista en el menú lateral.")
