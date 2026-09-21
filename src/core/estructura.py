"""
Estructura del documento: detecta los títulos del esquema oficial, verifica que
estén todos, en orden y con contenido, y marca el texto guía de la plantilla que
el tesista no borró.
"""
import os
import re
from collections import Counter

from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
from rapidfuzz import fuzz

from .utils import corto, limpiar_titulo, norm, ubic


def es_indice(p):
    """lineas de tabla de contenido / indice de tablas (repiten los titulos con numero de pagina)"""
    estilo = (p.style.name or "").lower() if p.style is not None else ""
    if re.match(r"^(toc|tdc|table of figures|tabla de ilustraciones)", estilo):
        return True
    return re.search(r"(\t|\.{4,}|…{2,})\s*\d+\s*$", p.text) is not None


# Un párrafo es un título por varias razones independientes del texto que dice.
# Apoyarse solo en el parecido textual dejaba fuera títulos legítimos ("Fuentes
# bibliográficas" puntúa 79 contra 'bibliografia'), y eso no producía ruido sino
# silencio: sin la sección, los chequeos que dependen de ella no reportan nada.
RE_CAPITULO = re.compile(r"^\s*CAP[IÍ]TULO\s+(?:[IVXLC]+|\d+)\b", re.I)
RE_NUMERADO = re.compile(r"^\s*\d+(?:\.\d+)*\.?\s+\S")

UMBRAL = 82              # parecido exigido a un párrafo que no da ninguna señal de formato
UMBRAL_CON_SENAL = 72    # cuando el formato confirma que es un título, se puede relajar
UMBRAL_SUGERENCIA = 55   # por debajo de esto no se sugiere nada
MARGEN_MINIMO = 6        # el mejor candidato debe despegarse del segundo para aceptarlo


def senales_de_titulo(p, texto):
    """Indicios de formato de que un párrafo es un título, independientes de su texto."""
    senales = []
    estilo = (p.style.name or "").lower() if p.style is not None else ""
    if re.search(r"(heading|t[ií]tulo)\s*\d", estilo):
        senales.append("estilo de título de Word")
    if RE_CAPITULO.match(texto) or RE_NUMERADO.match(texto):
        senales.append("numeración")
    runs = [r for r in p.runs if r.text.strip()]
    if runs and all(r.bold for r in runs):
        senales.append("negrita")
    letras = [c for c in texto if c.isalpha()]
    if len(letras) > 3 and all(c.isupper() for c in letras):
        senales.append("mayúsculas")
    return senales


def _dos_mejores(limpio, candidatos):
    """(mejor_nombre, mejor_puntaje, puntaje_del_segundo_nombre_distinto)"""
    mejor, punt, segundo = None, 0, 0
    for forma, nombre in candidatos:
        p = fuzz.ratio(limpio, forma)
        if p > punt:
            if nombre != mejor:
                segundo = punt
            mejor, punt = nombre, p
        elif nombre != mejor and p > segundo:
            segundo = p
    return mejor, punt, segundo


def detectar_secciones(bloques, reglas):
    """
    Devuelve (unicos, todos, sueltos):
      unicos   [(i, nombre)]  primera aparición de cada sección del esquema
      todos    [(i, nombre)]  todas las coincidencias, para detectar duplicados
      sueltos  [(i, texto, texto_normalizado)]  párrafos que parecen título por su
               formato pero no alcanzaron el umbral. Permiten sugerir un alias en vez
               de afirmar que falta la sección.
    """
    candidatos = []
    for sec in reglas["secciones"]:
        for forma in [sec["nombre"]] + sec.get("alias", []):
            candidatos.append((norm(forma), sec["nombre"]))

    encontrados, sueltos = [], []
    for i, b in enumerate(bloques):
        if not isinstance(b, Paragraph):
            continue
        txt = b.text.strip()
        if not txt or len(txt) > 400 or es_indice(b):
            continue
        limpio = limpiar_titulo(txt)
        if not limpio or len(limpio) > 70:
            continue

        mejor, puntaje, segundo = _dos_mejores(limpio, candidatos)
        senales = senales_de_titulo(b, txt)
        umbral = UMBRAL_CON_SENAL if senales else UMBRAL

        if puntaje < umbral:
            # el tesista alargó el título ("Referencias bibliográficas"): vale si el
            # nombre esperado aparece completo dentro del título
            for forma, nombre in candidatos:
                if len(forma) >= 8 and re.search(rf"\b{re.escape(forma)}\b", limpio):
                    mejor, puntaje, segundo = nombre, max(puntaje, umbral), 0
                    break

        # con el umbral relajado exigimos además que el candidato no sea ambiguo
        ambiguo = umbral == UMBRAL_CON_SENAL and puntaje < UMBRAL and (puntaje - segundo) < MARGEN_MINIMO
        if puntaje >= umbral and not ambiguo:
            encontrados.append((i, mejor))
        elif senales:
            sueltos.append((i, txt, limpio))

    # si un titulo aparece repetido (ej. "Referencias" citado en el texto), nos quedamos con el primero
    vistos, unicos = set(), []
    for i, n in encontrados:
        if n not in vistos:
            unicos.append((i, n))
            vistos.add(n)
    return unicos, encontrados, sueltos


def revisar_estructura(bloques, reglas, obs, deteccion=None):
    secciones, todos, sueltos = deteccion or detectar_secciones(bloques, reglas)
    esperado = [s["nombre"] for s in reglas["secciones"]]
    presentes = [n for _, n in secciones]
    tipo = reglas.get("tipo", "proyecto")

    formas_de = {s["nombre"]: [norm(f) for f in [s["nombre"]] + s.get("alias", [])]
                 for s in reglas["secciones"]}

    for n in esperado:
        if n in presentes:
            continue
        # ¿hay algún párrafo con formato de título que se parezca a ESTA sección?
        # Se puntúa contra las formas de 'n', no contra su mejor coincidencia global:
        # un título puede parecerse más a otra sección ya encontrada y aun así ser este.
        # Solo se sugiere si además 'n' es la sección a la que ESE título más se parece:
        # si el candidato se explica mejor por otra sección, sugerirlo es mala pista.
        parecidos = []
        for i, txt, limpio in sueltos:
            propio = max(fuzz.ratio(limpio, f) for f in formas_de[n])
            ajeno = max((max(fuzz.ratio(limpio, f) for f in formas)
                         for otra, formas in formas_de.items() if otra != n), default=0)
            if propio >= UMBRAL_SUGERENCIA and propio >= ajeno - MARGEN_MINIMO:
                parecidos.append((propio, txt, i))
        parecidos.sort()
        if parecidos:
            punt, txt, i = parecidos[-1]
            punt = int(punt)
            obs.add("Estructura", "Error", ubic(i, {}, bloques, txt, 50),
                    f"No se reconoció la sección '{n}'",
                    f"El título más parecido del documento es '{corto(txt, 60)}' "
                    f"({punt}% de parecido). Si es esa sección, agregar "
                    f"'{corto(txt, 45)}' a los 'alias' de '{n}' en reglas/{tipo}.yaml. "
                    f"Si no lo es, entonces la sección falta de verdad.")
        else:
            obs.add("Estructura", "Error", "Documento", f"Falta la sección '{n}'",
                    f"No se encontró un título que corresponda. Si existe con otro nombre, "
                    f"agregar el alias en reglas/{tipo}.yaml.")

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


