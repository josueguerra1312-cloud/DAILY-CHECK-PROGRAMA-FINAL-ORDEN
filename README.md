# Generador Daily Check GDL

## Archivos en la raiz

- `app.py`
- `daily_check.py`
- `requirements.txt`
- `.python-version`
- `.gitignore`

## Despliegue

1. Borra del repositorio los Python anteriores, especialmente `streamlit_app.py`, `generar_daily_check.py` y `procesar_daily_check.py`.
2. Copia estos cinco archivos en la raiz.
3. En Streamlit Community Cloud elimina la aplicacion anterior y crea una nueva.
4. Usa `app.py` como Main file path.
5. Selecciona Python 3.12 en Advanced settings.
6. Despliega.

La interfaz solicita `DAILY CHECK PLANTILLA.xlsx` y `PROGRAMA PROCESADO.xlsx`. El resultado sigue la logica del archivo `DAILY CHECK GDL SEP 01.xlsx`.
