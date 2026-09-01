# Daily Check GDL

Aplicacion Streamlit con dos archivos de entrada:

1. `Daily Check Plantilla.xlsx`
2. `Programa Procesado.xlsx`

La descarga conserva la estructura de la plantilla y ordena vuelos, transitos, matriculas, WO y tareas siguiendo la logica del archivo `Daily Check GDL SEP 01`.

## Archivos del repositorio

Coloca estos archivos directamente en la raiz:

```text
app.py
requirements.txt
README.md
.gitignore
```

## Ejecutar localmente

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
py -m pip install -r requirements.txt
py -m streamlit run app.py
```

## Streamlit Community Cloud

- Main file path: `app.py`
- No subas los Excel de trabajo al repositorio.
- Carga ambos Excel desde la pantalla de la aplicacion.
- Si reemplazas archivos en GitHub, reinicia la aplicacion desde Streamlit Cloud.

## Funcionamiento

- Lee vuelos desde `Sheet1` de la plantilla.
- Lee matriculas, WO, tareas y descripciones desde Programa Procesado.
- Ordena vuelos en jornada operacional desde las 18:00.
- Para matriculas repetidas, asocia el programa al ultimo vuelo del dia.
- Los vuelos sin programa se marcan como TRANSIT CHECK o TRANSIT CHECK / RON.
- Las matriculas programadas sin vuelo se colocan como STORAGE.
- Agrupa las tareas combinando las columnas A:E de cada matricula.
- Mantiene STATUS, DONE, PENDING, responsables y comentarios vacios.
- Busca M/H en la hoja HM.
- Desplaza el resumen solo si se requieren filas adicionales y actualiza sus formulas.
