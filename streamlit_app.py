import io
import streamlit as st
from processor import generate_combined

st.set_page_config(page_title="Daily Check Combinado", page_icon="✈️", layout="centered")
st.title("Daily Check Combinado")
st.write("Carga los dos archivos diarios. La aplicación conserva el libro plantilla y completa la hoja GDL.")

programa = st.file_uploader("1. PROGRAMA PROCESADO", type=["xlsx"], key="programa")
plantilla = st.file_uploader("2. DAILY CHECK PLANTILLA", type=["xlsx"], key="plantilla")

if programa and plantilla:
    try:
        resultado = generate_combined(io.BytesIO(programa.getvalue()), io.BytesIO(plantilla.getvalue()))
        st.success("Archivo generado correctamente.")
        st.download_button(
            "Descargar DAILY CHECK COMBINADO.xlsx",
            data=resultado,
            file_name="DAILY CHECK COMBINADO.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    except Exception as exc:
        st.error(f"No fue posible generar el archivo: {exc}")
else:
    st.info("Selecciona ambos archivos para iniciar el proceso.")
