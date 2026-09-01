#!/usr/bin/env python3
"""Incorpora PROGRAMA PROCESADO en un Daily Check sin reconstruir el libro."""
from __future__ import annotations
import argparse
import copy
import re
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell

AC_RE = re.compile(r"^(?:XA-[A-Z0-9]{3}|N[0-9]{3}[A-Z]{2})$", re.I)


def txt(value):
    return "" if value is None else str(value).strip()


def is_ac(value):
    return bool(AC_RE.fullmatch(txt(value).upper()))


def normalize(value):
    return re.sub(r"\s+", " ", txt(value).replace("_x000D_", " ")).upper().strip()


def tarea_base(value):
    value = normalize(value)
    return re.sub(r"\s+(?:ENG\s*[12]|LH|RH)$", "", value).strip()


def leer_programa(path):
    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]
    programa = []
    actual = None
    for fila in ws.iter_rows(values_only=True):
        a, b, c, d = (list(fila) + [None] * 4)[:4]
        if is_ac(a):
            actual = {"ac": txt(a).upper(), "wo": txt(b), "tareas": []}
            programa.append(actual)
            if txt(c):
                actual["tareas"].append((txt(c), txt(d) or txt(c)))
        elif actual and txt(c):
            actual["tareas"].append((txt(c), txt(d) or txt(c)))
    return programa


def buscar_fila(ws, texto, columna=1):
    objetivo = normalize(texto)
    for r in range(1, ws.max_row + 1):
        if normalize(ws.cell(r, columna).value) == objetivo:
            return r
    raise ValueError(f"No se encontro {texto!r} en la hoja {ws.title}")


def copiar_estilo_fila(ws, origen, destino):
    ws.row_dimensions[destino].height = ws.row_dimensions[origen].height
    ws.row_dimensions[destino].hidden = ws.row_dimensions[origen].hidden
    for col in range(1, ws.max_column + 1):
        src, dst = ws.cell(origen, col), ws.cell(destino, col)
        if not isinstance(src, MergedCell):
            dst._style = copy.copy(src._style)
            dst.number_format = src.number_format
            dst.alignment = copy.copy(src.alignment)
            dst.protection = copy.copy(src.protection)


def hm_catalogo(wb):
    exacto, base = {}, {}
    if "HM" not in wb.sheetnames:
        return exacto, base
    ws = wb["HM"]
    for card, _descripcion, mh in ws.iter_rows(min_row=2, max_col=3, values_only=True):
        if txt(card):
            exacto.setdefault(normalize(card), mh)
            base.setdefault(tarea_base(card), mh)
    return exacto, base


def filas_vuelos(wb):
    if "Sheet1" not in wb.sheetnames:
        return []
    ws = wb["Sheet1"]
    encabezado = buscar_fila(ws, "A/C")
    resultado = []
    for r in range(encabezado + 1, ws.max_row + 1):
        ac = txt(ws.cell(r, 1).value).upper()
        if is_ac(ac):
            resultado.append((ac, ws.cell(r, 2).value, ws.cell(r, 3).value, ws.cell(r, 4).value))
    return resultado


def escoger_vuelo(vuelos, ac):
    candidatos = [v for v in vuelos if v[0] == ac]
    if not candidatos:
        return None
    # Para matriculas repetidas se usa la ultima aparicion, igual que en el ejemplo final.
    return candidatos[-1]


def generar(plantilla, programa_path, salida):
    # No se crea un libro nuevo. Se abre y edita la plantilla original para conservar su esencia.
    wb = load_workbook(plantilla, keep_links=True)
    ws = wb["GDL"]
    programa = leer_programa(programa_path)
    vuelos = filas_vuelos(wb)
    hm_exacto, hm_base = hm_catalogo(wb)

    inicio = buscar_fila(ws, "A/C") + 1
    resumen = buscar_fila(ws, "RON AC")

    # Se conserva sin cambios todo lo que esta debajo del resumen.
    # Dentro del cuerpo solo se reemplaza contenido; no se modifican anchos, colores ni encabezados.
    capacidad = resumen - inicio
    filas_necesarias = sum(1 + len(p["tareas"]) for p in programa)
    if filas_necesarias > capacidad:
        raise ValueError(
            f"La plantilla tiene {capacidad} filas disponibles y se requieren {filas_necesarias}. "
            "Amplia el area operativa en Excel sin modificar el bloque de resumen."
        )

    # Limpia exclusivamente datos operativos del cuerpo. Mantiene estilos y validaciones.
    for r in range(inicio, resumen):
        for c in range(1, 17):
            if not isinstance(ws.cell(r, c), MergedCell):
                ws.cell(r, c).value = None

    fila = inicio
    for bloque in programa:
        copiar_estilo_fila(ws, inicio, fila)
        vuelo = escoger_vuelo(vuelos, bloque["ac"])
        ws.cell(fila, 1).value = bloque["ac"]
        if vuelo:
            ws.cell(fila, 2).value = vuelo[1]
            ws.cell(fila, 3).value = vuelo[2]
            ws.cell(fila, 4).value = vuelo[3]
        else:
            ws.cell(fila, 2).value = "STORAGE"
        ws.cell(fila, 5).value = bloque["wo"]
        fila += 1

        for card, descripcion in bloque["tareas"]:
            copiar_estilo_fila(ws, inicio, fila)
            ws.cell(fila, 6).value = card
            ws.cell(fila, 7).value = descripcion
            mh = hm_exacto.get(normalize(card))
            if mh is None:
                mh = hm_base.get(tarea_base(card))
            ws.cell(fila, 10).value = mh
            # STATUS, DONE, PENDING y comentarios quedan vacios.
            for c in range(8, 10):
                ws.cell(fila, c).value = None
            for c in range(11, 17):
                ws.cell(fila, c).value = None
            fila += 1

    # Fuerza recalculo al abrir en Excel sin sustituir formulas existentes.
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.save(salida)
    return len(programa), filas_necesarias


def main():
    parser = argparse.ArgumentParser(description="Genera el Daily Check final conservando la plantilla")
    parser.add_argument("--plantilla", default="DAILY CHECK PLANTILLA.xlsx")
    parser.add_argument("--programa", default="PROGRAMA PROCESADO.xlsx")
    parser.add_argument("--salida", default="Daily check sep 01.xlsx")
    args = parser.parse_args()
    n, filas = generar(Path(args.plantilla), Path(args.programa), Path(args.salida))
    print(f"Listo: {args.salida} | matriculas: {n} | filas incorporadas: {filas}")


if __name__ == "__main__":
    main()
