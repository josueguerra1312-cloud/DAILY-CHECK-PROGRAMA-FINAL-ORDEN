# Daily Check Combinado

Aplicación Streamlit que combina `PROGRAMA PROCESADO.xlsx` con `DAILY CHECK PLANTILLA.xlsx` y genera `DAILY CHECK COMBINADO.xlsx`.

## Archivos obligatorios en la raíz del repositorio

- `streamlit_app.py`
- `processor.py`
- `requirements.txt`
- `runtime.txt`

No coloques estos archivos dentro de otra carpeta.

## Configuración en Streamlit Community Cloud

Al crear o editar la aplicación, usa estos valores:

- Repository: el repositorio correspondiente
- Branch: `main`
- Main file path: `streamlit_app.py`

El nombre distingue mayúsculas y minúsculas.

## Lógica

- Conserva las hojas y estructura de la plantilla.
- Solo agrega tareas cuando la matrícula coincide.
- Para aviones `STORAGE`, conserva los trabajos existentes y agrega las tareas coincidentes del programa.
- Si un avión `STORAGE` no coincide, no agrega trabajos nuevos.
- Para matrículas repetidas en vuelos, asigna el trabajo al vuelo con mayor tiempo disponible.
- Los vuelos restantes se clasifican como `TRANSIT CHECK` o `TRANSIT CHECK / RON`.

## Ejecución local

```bash
python -m venv .venv
```

Windows:

```bash
.venv\Scripts\activate
```

Instalación y ejecución:

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```
