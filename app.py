from datetime import date
import streamlit as st

st.set_page_config(page_title="Daily Check GDL", page_icon="DC", layout="centered")
st.title("Generador Daily Check GDL")
st.write("Carga DAILY CHECK PLANTILLA y PROGRAMA PROCESADO para generar el archivo final.")

plantilla = st.file_uploader("1. DAILY CHECK PLANTILLA.xlsx", type=["xlsx"], key="plantilla")
programa = st.file_uploader("2. PROGRAMA PROCESADO.xlsx", type=["xlsx"], key="programa")
fecha = st.date_input("Fecha del Daily Check", value=date.today())

if "resultado" not in st.session_state:
    st.session_state.resultado = None
    st.session_state.nombre = None
    st.session_state.resumen = None

if st.button("Generar archivo final", type="primary", use_container_width=True):
    if plantilla is None or programa is None:
        st.warning("Carga los dos archivos de entrada.")
    else:
        try:
            from daily_check import generar
            with st.spinner("Incorporando tareas y ordenando el Daily Check..."):
                contenido, resumen = generar(plantilla.getvalue(), programa.getvalue())
            meses = ["ENE", "FEB", "MAR", "ABR", "MAY", "JUN", "JUL", "AGO", "SEP", "OCT", "NOV", "DIC"]
            st.session_state.nombre = f"Daily check GDL {meses[fecha.month-1]} {fecha.day:02d}.xlsx"
            st.session_state.resultado = contenido
            st.session_state.resumen = resumen
            st.success(f"Archivo listo: {resumen['matriculas']} matriculas y {resumen['tareas']} tareas.")
        except Exception as exc:
            st.session_state.resultado = None
            st.error("No fue posible generar el archivo.")
            st.exception(exc)

if st.session_state.resultado:
    st.download_button(
        "Descargar Daily Check final",
        data=st.session_state.resultado,
        file_name=st.session_state.nombre,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
