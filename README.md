# Generador Daily Check GDL

Aplicacion Streamlit que incorpora `PROGRAMA PROCESADO.xlsx` en `DAILY CHECK PLANTILLA.xlsx` y permite descargar el resultado.

## Archivos en la raiz

```text
app.py
generar_daily_check.py
requirements.txt
README.md
.gitignore
```

Los Excel no se guardan en GitHub. Se cargan desde la pantalla de la aplicacion cada vez que se genera un Daily Check.

## Ejecucion local

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
py -m pip install -r requirements.txt
py -m streamlit run app.py
```

## Streamlit Community Cloud

1. Sube los cinco archivos a la raiz del repositorio.
2. En Streamlit Cloud selecciona el repositorio.
3. En `Main file path` escribe `app.py`.
4. Despliega la aplicacion.
5. Carga la plantilla y el programa desde la interfaz.

## Correccion del FileNotFoundError

La aplicacion ya no busca archivos Excel por nombre dentro del servidor. Ambos libros se reciben mediante `st.file_uploader` y el resultado se genera en memoria. Por ello no es necesario subir archivos operativos al repositorio.

## Proteccion del formato

- Se abre la plantilla existente y no se crea un libro nuevo.
- No se insertan ni eliminan filas.
- No se mueve el bloque de resumen.
- Se conservan hojas, estilos, formulas y validaciones existentes.
- Se sustituyen solo las celdas operativas A:P entre el encabezado y `RON AC`.
- Los campos STATUS, DONE, PENDING, responsables y comentarios quedan vacios.
