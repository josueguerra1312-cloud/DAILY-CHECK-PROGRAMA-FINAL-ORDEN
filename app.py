from datetime import date
import streamlit as st
from generar_daily_check import generar_daily_check

st.set_page_config(page_title="Generador Daily Check", page_icon="📋", layout="centered")
st.title("Generador Daily Check GDL")
st.caption("Carga la plantilla y el programa procesado. La aplicacion devuelve un Excel nuevo sin modificar los archivos originales.")

plantilla = st.file_uploader("1. DAILY CHECK PLANTILLA.xlsx", type=["xlsx"], key="plantilla")
programa = st.file_uploader("2. PROGRAMA PROCESADO.xlsx", type=["xlsx"], key="programa")
fecha = st.date_input("Fecha del archivo de salida", value=date.today())
nombre = f"Daily check {fecha.strftime('%b %d').lower()}.xlsx"

if st.button("Generar Daily Check", type="primary", use_container_width=True):
    if plantilla is None or programa is None:
        st.error("Carga los dos archivos Excel antes de continuar.")
    else:
        try:
            with st.spinner("Procesando archivos..."):
                contenido, resumen = generar_daily_check(plantilla, programa)
            st.session_state["resultado"] = contenido
            st.session_state["resumen"] = resumen
            st.session_state["nombre"] = nombre
            st.success(f"Archivo generado: {resumen['matriculas']} matriculas y {resumen['tareas']} tareas.")
        except Exception as error:
            st.session_state.pop("resultado", None)
            st.error(f"No fue posible generar el archivo: {error}")

if "resultado" in st.session_state:
    st.download_button(
        "Descargar Excel final",
        data=st.session_state["resultado"],
        file_name=st.session_state["nombre"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )
    sin_hm = st.session_state["resumen"].get("sin_hm", [])
    if sin_hm:
        with st.expander(f"Tareas sin coincidencia de M/H ({len(sin_hm)})"):
            st.write("\n".join(f"• {x}" for x in sin_hm))
