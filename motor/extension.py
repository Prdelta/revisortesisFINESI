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
import localizador
import reglas_revisor
from reporte_resumen import escribir_resumen, redactar
from reporte_word import escribir_word
from .utils import *

def buscar_soffice():
    import shutil
    for nombre in ("soffice", "libreoffice"):
        if shutil.which(nombre):
            return shutil.which(nombre)
    for ruta in (r"C:\Program Files\LibreOffice\program\soffice.exe",
                 r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
                 "/Applications/LibreOffice.app/Contents/MacOS/soffice",
                 "/usr/bin/soffice",
                 "/usr/bin/libreoffice"):
        if os.path.exists(ruta):
            return ruta
    return None


def paginas_pdf(pdf):
    try:
        from pypdf import PdfReader
        return len(PdfReader(pdf).pages)
    except Exception:
        with open(pdf, "rb") as fh:
            return len(re.findall(rb"/Type\s*/Page[^s]", fh.read())) or None


def contar_paginas(ruta, mapa=None):
    """usa el conteo del PDF ya generado; si no hay, el dato que guarda Word en el archivo"""
    if mapa is not None and mapa.paginas:
        return mapa.paginas, "LibreOffice"
    try:
        import zipfile
        with zipfile.ZipFile(ruta) as z:
            xml = z.read("docProps/app.xml").decode("utf8", "ignore")
        m = re.search(r"<Pages>(\d+)</Pages>", xml)
        if m and int(m.group(1)) > 0:
            return int(m.group(1)), "dato guardado por Word al último guardado"
    except Exception as ex:
        import traceback
        print(f"Error al contar_paginas zipfile: {ex}\n{traceback.format_exc()}")
    return None, None


def texto_de_seccion(bloques, rangos, nombre):
    if nombre not in rangos:
        return ""
    ini, fin = rangos[nombre]
    partes = []
    tit = bloques[ini].text.split(":", 1)
    if len(tit) > 1:
        partes.append(tit[1])
    partes += [b.text for b in bloques[ini + 1:fin] if isinstance(b, Paragraph)]
    return "\n".join(p for p in partes if p.strip())


def revisar_extension(ruta, bloques, rangos, reglas, obs, mapa=None):
    e = reglas.get("extension") or {}
    if e.get("max_paginas"):
        pags, metodo = contar_paginas(ruta, mapa)
        if pags is None:
            obs.add("Extensión", "Revisar", "Documento", "No se pudo contar las páginas", "Verificar manualmente en Word.")
        elif pags > e["max_paginas"]:
            obs.add("Extensión", "Error", "Documento", f"El documento tiene {pags} páginas, máximo {e['max_paginas']}",
                    f"Conteo: {metodo}. Puede variar en una página respecto a lo que muestra Word.")

    tit = texto_de_seccion(bloques, rangos, e.get("seccion_titulo", "Título"))
    if tit and e.get("titulo_max_palabras"):
        n = len(tit.split())
        if n > e["titulo_max_palabras"]:
            obs.add("Extensión", "Error", e.get("seccion_titulo", "Título"), f"El título tiene {n} palabras, máximo {e['titulo_max_palabras']}", corto(tit, 120))

    kw = texto_de_seccion(bloques, rangos, e.get("seccion_palabras_clave", "Palabras claves"))
    if kw and e.get("palabras_clave_max"):
        # si el tesista puso "Keywords:" en ingles, contamos solo la primera linea (en espanol)
        linea = [l for l in kw.split("\n") if l.strip()][0]
        linea = re.sub(r"^\s*(palabras? claves?|keywords)\s*:?", "", linea, flags=re.I)
        items = [k for k in re.split(r"[;,]", linea) if k.strip()]
        if len(items) > e["palabras_clave_max"]:
            obs.add("Extensión", "Error", e.get("seccion_palabras_clave", "Palabras claves"), f"Tiene {len(items)} palabras clave, máximo {e['palabras_clave_max']}", corto(linea, 120))


