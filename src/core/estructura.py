import argparse
import os
import re
import shutil
import subprocess
import sys
import tempfile
import unicodedata
from collections import Counter, defaultdict
from datetime import datetime
import yaml
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from rapidfuzz import fuzz
from utilidades import localizador
from utilidades import reglas_revisor
from reportes.reporte_resumen import escribir_resumen, redactar
from reportes.reporte_word import escribir_word
from .utils import *
from .config import DATOS, RECURSOS

def es_indice(p):
    """lineas de tabla de contenido / indice de tablas (repiten los titulos con numero de pagina)"""
    estilo = (p.style.name or "").lower() if p.style is not None else ""
    if re.match(r"^(toc|tdc|table of figures|tabla de ilustraciones)", estilo):
        return True
    return re.search(r"(\t|\.{4,}|…{2,})\s*\d+\s*$", p.text) is not None


def detectar_secciones(bloques, reglas):
    """devuelve lista de (indice_bloque, nombre_canonico) de los titulos encontrados"""
    candidatos = []
    for sec in reglas["secciones"]:
        for forma in [sec["nombre"]] + sec.get("alias", []):
            candidatos.append((norm(forma), sec["nombre"]))
    encontrados = []
    for i, b in enumerate(bloques):
        if not isinstance(b, Paragraph):
            continue
        txt = b.text.strip()
        if not txt or len(txt) > 400 or es_indice(b):
            continue
        limpio = limpiar_titulo(txt)
        if not limpio or len(limpio) > 70:
            continue
        mejor, puntaje = None, 0
        for forma, nombre in candidatos:
            p = fuzz.ratio(limpio, forma)
            if p > puntaje:
                mejor, puntaje = nombre, p
        if puntaje < 82:
            # el tesista alargó el título ("Referencias bibliográficas", "Metodología de la investigación"):
            # vale si el nombre esperado aparece completo dentro del título
            for forma, nombre in candidatos:
                if len(forma) >= 8 and re.search(rf"\b{re.escape(forma)}\b", limpio):
                    mejor, puntaje = nombre, 82
                    break
        if puntaje >= 82:
            encontrados.append((i, mejor))
    # si un titulo aparece repetido (ej. "Referencias" citado en el texto), nos quedamos con el primero
    vistos, unicos = set(), []
    for i, n in encontrados:
        if n not in vistos:
            unicos.append((i, n))
            vistos.add(n)
    return unicos, encontrados


def revisar_estructura(bloques, reglas, obs):
    secciones, todos = detectar_secciones(bloques, reglas)
    esperado = [s["nombre"] for s in reglas["secciones"]]
    presentes = [n for _, n in secciones]

    for n in esperado:
        if n not in presentes:
            obs.add("Estructura", "Error", "Documento", f"Falta la sección '{n}'",
                    "No se encontró un título que corresponda. Si existe con otro nombre, agregar el alias en reglas.yaml.")

    cnt = Counter(n for _, n in todos)
    for n, c in cnt.items():
        if c > 1:
            obs.add("Estructura", "Revisar", n, f"El título '{n}' aparece {c} veces", "Puede ser un duplicado o una mención en el texto.")

    orden = [esperado.index(n) for n in presentes]
    for k in range(1, len(orden)):
        if orden[k] < orden[k - 1]:
            obs.add("Estructura", "Error", presentes[k],
                    f"Sección fuera de orden: '{presentes[k]}' aparece después de '{presentes[k-1]}'",
                    "Respetar el orden del esquema oficial.")

    # rangos de cada seccion
    rangos = {}
    for k, (i, n) in enumerate(secciones):
        fin = secciones[k + 1][0] if k + 1 < len(secciones) else len(bloques)
        rangos[n] = (i, fin)

    for n, (ini, fin) in rangos.items():
        contenido = bloques[ini + 1:fin]
        texto_titulo = bloques[ini].text
        resto_titulo = re.split(r":", re.sub(r"\([^()]*\)?", "", texto_titulo), maxsplit=1)
        en_linea = len(resto_titulo) > 1 and len(resto_titulo[1].strip()) > 3
        tiene = en_linea or any(
            (isinstance(b, Paragraph) and b.text.strip()) or isinstance(b, Table) for b in contenido)
        if not tiene:
            obs.add("Estructura", "Error", n, f"La sección '{n}' está vacía")

    for n in reglas.get("secciones_con_tabla", []):
        if n in rangos:
            ini, fin = rangos[n]
            tablas = [b for b in bloques[ini + 1:fin] if isinstance(b, Table)]
            if not tablas:
                obs.add("Estructura", "Error", n, f"La sección '{n}' no tiene tabla")
            else:
                t = tablas[0]
                celdas = {c.text.strip() for r in t.rows for c in r.cells if c.text.strip()}
                propias = {x for x in celdas if norm(x) not in TEXTOS_PLANTILLA_TABLAS}
                llenas = propias
                if not propias:
                    obs.add("Estructura", "Error", n, f"La tabla de '{n}' está vacía o casi vacía",
                            "Solo tiene el encabezado de la plantilla.")

    for padre, hijos in reglas.get("subsecciones", {}).items():
        if padre not in rangos:
            continue
        ini, fin = rangos[padre]
        textos = [limpiar_titulo(b.text) for b in bloques[ini + 1:fin] if isinstance(b, Paragraph) and b.text.strip()]
        for h in hijos:
            if not any(fuzz.ratio(norm(h), t) >= 75 or fuzz.partial_ratio(norm(h), t) >= 90 for t in textos):
                obs.add("Estructura", "Error", padre, f"Falta la subsección '{h}'")

    return rangos



TEXTOS_PLANTILLA_TABLAS = set()


def cargar_textos_tablas(ruta_plantilla):
    global TEXTOS_PLANTILLA_TABLAS
    if ruta_plantilla and os.path.exists(ruta_plantilla):
        d = Document(ruta_plantilla)
        TEXTOS_PLANTILLA_TABLAS = {norm(c.text) for t in d.tables for r in t.rows for c in r.cells if c.text.strip()}


def marcas_guia(ruta_plantilla):
    """extrae el texto guia entre parentesis de la plantilla (lo que el tesista debe borrar)"""
    if not ruta_plantilla or not os.path.exists(ruta_plantilla):
        return []
    doc = Document(ruta_plantilla)
    marcas = set()
    for p in doc.paragraphs:
        for m in re.findall(r"\(([^()]{20,})\)?", p.text):
            m = norm(m)
            if len(m) >= 20:
                marcas.add(m[:45])
    return sorted(marcas)


def revisar_marcas(bloques, rangos, marcas, obs):
    for i, b in enumerate(bloques):
        if not isinstance(b, Paragraph) or not b.text.strip():
            continue
        t = norm(b.text)
        for m in marcas:
            if m in t or (len(t) > 30 and fuzz.partial_ratio(m, t) >= 93):
                obs.add("Formato", "Error", ubic(i, rangos, bloques, b.text, 50),
                        "Quedó texto guía de la plantilla sin borrar", f"Coincide con: '{m}...'")
                break


