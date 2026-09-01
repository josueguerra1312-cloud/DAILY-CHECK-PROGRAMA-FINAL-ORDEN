# Daily Check GDL

## Archivos del repositorio

Copia directamente en la raiz de GitHub:

```text
streamlit_app.py
app.py
requirements.txt
README.md
.gitignore
```

`streamlit_app.py` y `app.py` contienen la misma aplicacion. Para un despliegue nuevo se recomienda seleccionar `streamlit_app.py` como archivo principal.

## Entradas

La pantalla solicita:

1. `DAILY CHECK PLANTILLA.xlsx`
2. `PROGRAMA PROCESADO.xlsx`

El Excel de referencia `DAILY CHECK GDL SEP 01.xlsx` no se carga en la aplicacion. Su logica ya esta incorporada en el codigo.

## Salida

La aplicacion genera y permite descargar:

```text
Daily check GDL SEP 01.xlsx
```

El nombre cambia de acuerdo con la fecha seleccionada.

## Ejecucion local

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
py -m pip install -r requirements.txt
py -m streamlit run streamlit_app.py
```

## Despliegue limpio en Streamlit Community Cloud

1. Sustituye los archivos anteriores por esta version.
2. Haz commit y push a GitHub.
3. Elimina la aplicacion anterior de Streamlit Community Cloud.
4. Crea un despliegue nuevo.
5. Selecciona Python 3.12 en Advanced settings.
6. Usa `streamlit_app.py` como Main file path.
7. Pulsa Deploy.

No uses `runtime.txt`: la version de Python se selecciona desde Advanced settings durante el despliegue.

## Logica incluida

- Lee los vuelos desde `Sheet1` de la plantilla.
- Ordena la jornada desde las 18:00 y continua despues de medianoche.
- Inserta matrículas, vuelos, horarios, WO, tareas y descripciones.
- Coloca el programa en la ultima aparicion de una matricula repetida.
- Clasifica los vuelos sin programa como `TRANSIT CHECK` o `TRANSIT CHECK / RON`.
- Coloca como `STORAGE` las matriculas programadas sin vuelo.
- Combina A:E dentro de cada grupo de tareas.
- Conserva hojas, formato general, catalogos y bloque de resumen.
- Deja vacios STATUS, DONE, PENDING, responsables y comentarios.
- Busca M/H en la hoja `HM`.
