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

APELLIDO = r"[A-ZÁÉÍÓÚÑÜ][A-Za-zÁÉÍÓÚÑÜáéíóúñü'\-]+"
PARTICULA = r"(?:(?:de|del|la|las|los|van|von|da|di|le|mc)\s+)*"
AUTOR = rf"{PARTICULA}{APELLIDO}(?:[\s\-]{APELLIDO})?"
ANIO = r"(?:\d{4}[a-z]?|s\.\s?f\.)"

RE_PAREN = re.compile(r"\(([^()]*?(?:\d{4}[a-z]?|s\.\s?f\.)[^()]*)\)")
RE_SIGLA = re.compile(r"\s*\[[^\]]+\]")   # (Ministerio de Salud [MINSA], 2020)
RE_NARR = re.compile(rf"({AUTOR}(?:\s+et\s+al\.?|\s+(?:y|&)\s+{AUTOR})?)\s+\(({ANIO})(?:[,;][^)]*)?\)")
RE_PARTE = re.compile(rf"^(?:(?:ver|véase|cf\.|p\.\s?ej\.)\s+)?(.+?),?\s+({ANIO})(?:[,:].*)?$", re.I)


def primer_autor(autores):
    a = re.split(r"\s+et\s+al|\s+y\s+|\s*&\s*|,", autores)[0]
    return norm(a)


def extraer_citas(texto):
    citas = []
    for m in RE_PAREN.finditer(texto):
        for parte in m.group(1).split(";"):
            parte = RE_SIGLA.sub("", parte.strip()).strip()
            pm = RE_PARTE.match(parte)
            if pm and re.search(r"[A-ZÁÉÍÓÚÑ]", pm.group(1)) and not re.fullmatch(r"[\d\s,.\-]+", pm.group(1)):
                citas.append((primer_autor(pm.group(1)), norm(pm.group(2)), parte))
    for m in RE_NARR.finditer(texto):
        citas.append((primer_autor(m.group(1)), norm(m.group(2)), m.group(0)))
    return citas


def revisar_citas(bloques, rangos, reglas, obs):
    c = reglas.get("citas") or {}
    if not c:
        return
    sec_ref = c.get("seccion_referencias", "Referencias")
    if str(c.get("estilo", "APA7")).upper().replace(" ", "") != "APA7":
        obs.add("Citas", "Revisar", "Documento", "Solo está implementada la verificación APA 7")
        return
    ref_ini, ref_fin = rangos.get(sec_ref, (None, None))

    # citas en el texto (todo lo que no es la seccion de referencias)
    citas = []
    numericas = 0
    for i, b in enumerate(bloques):
        if not isinstance(b, Paragraph) or (ref_ini is not None and ref_ini <= i < ref_fin):
            continue
        t = b.text
        citas += [(a, y, txt, i) for a, y, txt in extraer_citas(t)]
        numericas += len(re.findall(r"\[\d+(?:[\-,–]\s*\d+)*\]", t))
        for m in re.finditer(r"\bet al(?!\.)\b", t):
            obs.add("Citas", "Error", ubic(i, rangos, bloques, t[max(0, m.start()-30):m.end()+10]), "'et al' sin punto", "Debe escribirse 'et al.'")
        for m in re.finditer(rf"\(({APELLIDO})\s+(\d{{4}})\)", t):
            obs.add("Citas", "Error", ubic(i, rangos, bloques, rangos, bloques), f"Cita sin coma entre autor y año: '{m.group(0)}'",
                    f"En APA 7: ({m.group(1)}, {m.group(2)})")

    if numericas:
        obs.add("Citas", "Error", "Documento", f"Se encontraron {numericas} citas numéricas tipo [1]",
                "Mezcla de estilos: el documento debería usar APA 7 (autor, año).")

    if ref_ini is None:
        obs.add("Citas", "Error", "Documento",
                f"No se pudo verificar las citas contra la lista de referencias",
                f"No se identificó la sección '{sec_ref}'. Si el tesista la tituló de otra forma, "
                f"agrega ese nombre en 'alias' dentro de reglas/{reglas.get('tipo', 'proyecto')}.yaml.")
        return

    # lista de referencias
    refs = []
    for i in range(ref_ini + 1, ref_fin):
        b = bloques[i]
        if not isinstance(b, Paragraph) or len(b.text.strip()) < 15:
            continue
        t = b.text.strip()
        autor = primer_autor(t.split(",")[0]) if "," in t else norm(t.split(".")[0])
        m = re.search(r"\((\d{4}[a-z]?|s\.\s?f\.)(?:,[^)]*)?\)", t)
        anio = norm(m.group(1)) if m else None
        refs.append(dict(i=i, texto=t, autor=autor, anio=anio, par=b))
        if not m:
            obs.add("Citas", "Error", ubic(i, rangos, bloques, t, 50),
                    "Referencia sin año entre paréntesis", "Formato APA 7: Apellido, A. A. (año). Título...")
        fli = b.paragraph_format.first_line_indent
        if fli is None or fli >= 0:
            pass  # se evalua en bloque abajo
        if re.search(r"doi\.org|doi:", t, re.I) and not re.search(r"https://doi\.org/", t):
            obs.add("Citas", "Advertencia", ubic(i, rangos, bloques, t), "DOI con formato antiguo",
                    "En APA 7 el DOI va como https://doi.org/xxxxx")

    if not refs:
        obs.add("Citas", "Error", sec_ref, "La sección de referencias no tiene entradas reconocibles")
        return

    sin_francesa = [r for r in refs if not (r["par"].paragraph_format.first_line_indent is not None
                                            and r["par"].paragraph_format.first_line_indent < 0)]
    if len(sin_francesa) > len(refs) / 2:
        obs.add("Citas", "Error", sec_ref, f"{len(sin_francesa)} de {len(refs)} referencias sin sangría francesa",
                "APA 7 usa sangría francesa de 1.27 cm")

    autores_ref = [r["autor"] for r in refs]
    if autores_ref != sorted(autores_ref):
        for a, b in zip(refs, refs[1:]):
            if a["autor"] > b["autor"]:
                obs.add("Citas", "Error", sec_ref, "Referencias no están en orden alfabético",
                        f"'{corto(b['texto'], 40)}' debería ir antes de '{corto(a['texto'], 40)}'")
                break

    def coincide(c_aut, c_anio, r):
        return (c_aut == r["autor"] or c_aut.split()[-1:] == r["autor"].split()[-1:]) and (r["anio"] is None or c_anio == r["anio"])

    ya = set()
    for a, y, txt, i in citas:
        if (a, y) in ya:
            continue
        ya.add((a, y))
        if not any(coincide(a, y, r) for r in refs):
            mismo_autor = [r for r in refs if a == r["autor"]]
            det = f"Hay referencia de ese autor con año {mismo_autor[0]['anio']}" if mismo_autor else "No figura en Referencias"
            obs.add("Citas", "Error", ubic(i, rangos, bloques, txt), f"Cita sin referencia: '{txt}'", det)

    for r in refs:
        if r["anio"] and not any(coincide(a, y, r) for a, y, _, _ in citas):
            obs.add("Citas", "Advertencia", ubic(r['i'], rangos, bloques, r['texto']), f"Referencia no citada en el texto: '{corto(r['texto'], 60)}'",
                    "En APA 7 toda referencia debe estar citada en el texto (revisar a mano, la detección no es perfecta).")


