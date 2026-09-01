# Generador Daily Check

El programa incorpora la información de `PROGRAMA PROCESADO.xlsx` dentro de `DAILY CHECK PLANTILLA.xlsx` sin crear un diseño nuevo.

## Archivos

Copia estos archivos en la raiz del repositorio:

- `generar_daily_check.py`
- `requirements.txt`
- `README.md`
- `.gitignore`

Los Excel de trabajo tambien deben estar en la raiz:

- `DAILY CHECK PLANTILLA.xlsx`
- `PROGRAMA PROCESADO.xlsx`

## Instalacion en Windows

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
py -m pip install -r requirements.txt
```

## Ejecucion rapida

```powershell
py generar_daily_check.py
```

El resultado se guarda como:

```text
Daily check sep 01.xlsx
```

## Nombres personalizados

```powershell
py generar_daily_check.py --plantilla "DAILY CHECK PLANTILLA.xlsx" --programa "PROGRAMA PROCESADO.xlsx" --salida "Daily check sep 01.xlsx"
```

## Logica

1. Abre la plantilla existente, no crea un libro desde cero.
2. Conserva hojas, encabezados, anchos, colores, formulas, listas y bloque de resumen.
3. Lee matrícula, WO, tarea y descripción desde el programa procesado.
4. Inserta cada matrícula y todas sus tareas consecutivamente.
5. Conserva la descripción completa, incluyendo P/N y S/N.
6. Busca horas hombre en la hoja `HM`.
7. Deja en blanco STATUS, DONE, PENDING, causas y comentarios.
8. Guarda el resultado en un archivo nuevo.

## Importante

La plantilla debe tener suficiente espacio entre el encabezado `A/C` y la fila `RON AC`. Si faltan filas, el programa se detiene antes de alterar el bloque de resumen.
