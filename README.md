# Daily Check Combinado

Aplicación en Python y Streamlit que combina **PROGRAMA PROCESADO.xlsx** con **DAILY CHECK PLANTILLA.xlsx** y genera **DAILY CHECK COMBINADO.xlsx**.

## Lógica implementada

- Conserva todas las hojas del archivo plantilla.
- Trabaja sobre la hoja `GDL` sin reconstruir el libro desde cero.
- Solo incorpora grupos cuya matrícula exista en el programa y en la plantilla.
- Si una matrícula está marcada como `STORAGE`, conserva sus trabajos existentes y agrega debajo las tareas coincidentes de `PROGRAMA PROCESADO`.
- Si un avión `STORAGE` no aparece en el programa, permanece sin cambios y no recibe tareas nuevas.
- Si una matrícula aparece en varios vuelos, asigna las tareas al vuelo con mayor tiempo entre llegada y salida.
- Los vuelos sin tareas coincidentes reciben:
  - `TRANSIT CHECK` si salida es posterior a llegada.
  - `TRANSIT CHECK / RON` si la salida ocurre al día siguiente.
- Conserva celdas combinadas, estilos, fórmulas, validaciones y hojas auxiliares de la plantilla.
- Las tareas y descripciones se toman directamente de `PROGRAMA PROCESADO`.

## Archivos del repositorio

- `app.py`: interfaz Streamlit.
- `processor.py`: lectura, asignación y generación del Excel.
- `requirements.txt`: dependencias.
- `test_processor.py`: prueba rápida con los tres archivos de ejemplo.

## Ejecución local

1. Instala Python 3.11 o superior.
2. Coloca estos archivos en la raíz del repositorio.
3. Abre una terminal en esa carpeta.
4. Ejecuta:

```bash
python -m venv .venv
```

En Windows:

```bash
.venv\Scripts\activate
```

En macOS o Linux:

```bash
source .venv/bin/activate
```

Instala dependencias:

```bash
pip install -r requirements.txt
```

Inicia la aplicación:

```bash
streamlit run app.py
```

## Prueba con los archivos de ejemplo

Coloca en la misma carpeta:

- `PROGRAMA_PROCESADO.xlsx`
- `DAILY CHECK PLANTILLA.xlsx`
- `DAILY CHECK COMBINADO.xlsx`

Después ejecuta:

```bash
python test_processor.py
```

La prueba genera `RESULTADO_PRUEBA.xlsx` y verifica estructura, hojas y matrículas principales.
