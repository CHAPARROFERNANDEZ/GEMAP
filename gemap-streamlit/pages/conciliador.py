# -*- coding: utf-8 -*-
"""
Página "Conciliador de mayores" para la app Streamlit de GEMAP.

Colócala en la carpeta  pages/  de la app y deja  conciliar.py  en la raíz del proyecto
(junto al archivo principal de la app). Necesita pandas y openpyxl en requirements.txt.
"""
import io
import os
import sys
import tempfile
import zipfile

import hmac

import streamlit as st

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if RAIZ not in sys.path:
    sys.path.insert(0, RAIZ)

import conciliar  # noqa: E402  (programa de conciliación)


def procesar_archivos(archivos, prov, cli, partidas, progreso=None):
    """archivos: lista de (nombre, bytes). Devuelve (resultados, resumen_bytes, errores).
    resultados: lista de (nombre_salida, bytes). Todo se hace en una carpeta temporal que se borra."""
    resultados, errores = [], []
    with tempfile.TemporaryDirectory() as tmp:
        entrada = os.path.join(tmp, "entrada")
        salida = os.path.join(tmp, "resultados")
        os.makedirs(entrada)
        os.makedirs(salida)
        for k, (nombre, datos) in enumerate(archivos):
            ruta = os.path.join(entrada, os.path.basename(nombre))
            with open(ruta, "wb") as f:
                f.write(datos)
            try:
                conciliar.procesar(ruta, prov, cli, partidas, salida)
            except Exception as ex:  # una empresa con error no para las demás
                m = str(ex)
                if "zip" in m.lower():
                    m = "No parece un Excel (.xlsx) válido. Expórtalo de nuevo desde la contabilidad."
                errores.append((nombre, m))
            if progreso:
                progreso((k + 1) / len(archivos), nombre)
        resumen = None
        for f in sorted(os.listdir(salida)):
            with open(os.path.join(salida, f), "rb") as fh:
                datos = fh.read()
            if f == "RESUMEN_EMPRESAS.xlsx":
                resumen = datos
            else:
                resultados.append((f, datos))
    return resultados, resumen, errores


# ---------------------------------------------------------------------------
# Interfaz
# ---------------------------------------------------------------------------
if __name__ == "__main__" or True:
    st.set_page_config(page_title="Conciliador de mayores", page_icon="🧾", layout="wide")
    st.title("🧾 Conciliador de mayores")

    # --- Acceso protegido: la contraseña se define en Railway (Variables) como CONCILIADOR_PASSWORD.
    #     Si la variable no existe, la página no se activa (por seguridad).
    CLAVE = os.environ.get("CONCILIADOR_PASSWORD", "")
    if not CLAVE:
        st.warning("Esta página está desactivada: falta definir la variable CONCILIADOR_PASSWORD en Railway "
                   "(pestaña Variables). Contiene datos de clientes y no debe quedar abierta.")
        st.stop()
    if not st.session_state.get("conc_ok"):
        pwd = st.text_input("Contraseña", type="password")
        if pwd:
            if hmac.compare_digest(pwd.encode(), CLAVE.encode()):
                st.session_state["conc_ok"] = True
                st.rerun()
            else:
                st.error("Contraseña incorrecta.")
        st.stop()

    st.caption("Sube el mayor («Cuentas corrientes», .xlsx) de una o varias empresas. Casa facturas con pagos, "
               "revisa la 555 y te devuelve un Excel por empresa con colores, resumen e incidencias.")

    with st.expander("Opciones avanzadas (solo si la empresa usa otras cuentas)"):
        c1, c2, c3 = st.columns(3)
        prov = [x.strip() for x in c1.text_input("Cuentas de proveedores (prefijos)", "40,41").split(",") if x.strip()]
        cli = [x.strip() for x in c2.text_input("Cuentas de clientes (prefijos)", "43").split(",") if x.strip()]
        partidas = int(c3.text_input("Cuenta de partidas pendientes de aplicación", "55500000"))

    subidos = st.file_uploader("Mayores en Excel", type=["xlsx"], accept_multiple_files=True)

    if subidos and st.button("Conciliar", type="primary"):
        barra = st.progress(0.0, text="Empezando…")
        archivos = [(f.name, f.getvalue()) for f in subidos]

        def avance(p, nombre):
            barra.progress(p, text=f"Procesado: {nombre}")

        with st.spinner("Analizando… los mayores grandes pueden tardar un minuto."):
            res, resumen, errores = procesar_archivos(archivos, prov, cli, partidas, avance)
        barra.empty()
        st.session_state["res"] = res
        st.session_state["resumen"] = resumen
        st.session_state["errores"] = errores

    if "res" in st.session_state:
        for nombre, msg in st.session_state["errores"]:
            st.error(f"{nombre}: {msg}")
        res = st.session_state["res"]
        if res:
            st.success(f"{len(res)} empresa(s) procesada(s). Descarga los resultados:")
            if st.session_state["resumen"]:
                st.download_button("📊 RESUMEN_EMPRESAS.xlsx (una línea por empresa)", st.session_state["resumen"],
                                   file_name="RESUMEN_EMPRESAS.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            for nombre, datos in res:
                st.download_button(f"⬇️ {nombre}", datos, file_name=nombre, key="dl_" + nombre,
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            if len(res) > 1:
                buf = io.BytesIO()
                with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
                    for nombre, datos in res:
                        z.writestr(nombre, datos)
                    if st.session_state["resumen"]:
                        z.writestr("RESUMEN_EMPRESAS.xlsx", st.session_state["resumen"])
                st.download_button("📦 Descargar todo en un ZIP", buf.getvalue(), file_name="conciliaciones.zip",
                                   mime="application/zip")
        st.info("Las propuestas son sugerencias basadas en importes, fechas y referencias. Valídalas antes de "
                "contabilizar. Los archivos no se guardan en el servidor: se procesan y se borran.")
