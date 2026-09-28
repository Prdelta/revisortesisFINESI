"""
Hoja de revisión resumida, en el formato que usa la Dirección de Investigación:
título, ficha del proyecto, asesor y una lista corta de observaciones, cada una con
las líneas del documento donde aparece el problema ("Líneas 11, 126 y 633: …").

El revisor puede cambiar el formato sin tocar el programa:
  - reporte/hoja_revision.docx   plantilla de Word con marcadores {{TITULO}},
                                 {{ASESOR}}, {{OBSERVACIONES}}… (diseño, letra, logo)
  - reporte/frases_reporte.yaml  el texto de cada observación y cómo se escriben
                                 los números de línea
Si faltan, se usan los que trae el programa (ver crear_plantilla_hoja y FRASES).
"""
import copy
import os
import re

import yaml
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

TINTA = "1A1A1A"
ACENTO = "1F3864"
SUAVE = "6E6E6E"
LINEA = "BFBFBF"
FUENTE = "Arial"


# --------------------------------------------------------------------------- frases editables

FRASES = {
    "lineas": {
        "una": "Línea {n}",
        "varias": "Líneas {lista}",
        "separador": ", ",
        "ultimo": " y ",
        "maximo": 10,              # más líneas que esto: se muestran las primeras y "y otros"
        "y_otros": " y otros",
    },
    "observaciones": {
        "esquema": "Revisar y cumplir el esquema de {tipo} de tesis PGI",
        "formato": "corregir el formato del documento: {detalle}",
        "titulo_incompleto": "completar el título de la sección: debe decir «{seccion}»",
        "titulo_distinto": "titular la sección como indica el esquema: «{seccion}»",
        "texto_guia": "eliminar el texto guía de la plantilla que quedó sin borrar",
        "revision_ortografica": "Revisión gramatical y ortográfica de todo el proyecto",
        "citas": "corregir la forma de citar, estilo APA 7ª ed",
        "ortografia": "corregir ortografía y gramática",
        "referencias": "referenciar según norma APA 7ª ed",
        "sin_referencias": "Titular la sección de referencias como indica el esquema; no se pudo "
                           "verificar la correspondencia entre citas y referencias",
        "extension_paginas": "Reducir la extensión a {maximo} páginas como máximo (tiene {valor})",
        "extension_titulo": "Reducir el título a {maximo} palabras como máximo (tiene {valor})",
        "extension_palabras_clave": "Dejar {maximo} palabras clave como máximo (tiene {valor})",
        "sin_observaciones": "Sin observaciones de formato. Pasa a la siguiente etapa.",
    },
}


def cargar_frases(ruta):
    """
    Las frases del revisor encima de las del programa: lo que no escribió, queda como
    viene. Un archivo con error de formato se informa con ValueError.
    """
    frases = copy.deepcopy(FRASES)
    if not ruta or not os.path.exists(ruta):
        return frases
    try:
        with open(ruta, encoding="utf8") as fh:
            propias = yaml.safe_load(fh) or {}
    except yaml.YAMLError as ex:
        raise ValueError(f"{os.path.basename(ruta)} tiene un error de formato: {ex}") from ex
    for grupo, valores in propias.items():
        if isinstance(valores, dict) and grupo in frases:
            frases[grupo].update({k: v for k, v in valores.items() if v is not None})
    return frases


class Observacion(str):
    """una línea de la hoja; 'prefijo' es la parte que va en negrita ("Líneas 11 y 126")"""

    def __new__(cls, texto, prefijo=""):
        obj = super().__new__(cls, texto)
        obj.prefijo = prefijo
        return obj


def lista_de_lineas(numeros, frases=None):
    """[11, 126, 633] -> 'Líneas 11, 126 y 633'; [] -> ''"""
    f = (frases or FRASES)["lineas"]
    nums = sorted({int(n) for n in numeros if n})
    if not nums:
        return ""
    if len(nums) == 1:
        return f["una"].format(n=nums[0])
    maximo = int(f.get("maximo") or 0) or len(nums)
    if len(nums) > maximo:
        lista = f["separador"].join(str(n) for n in nums[:maximo]) + f["y_otros"]
    else:
        lista = f["separador"].join(str(n) for n in nums[:-1]) + f["ultimo"] + str(nums[-1])
    return f["varias"].format(lista=lista)


def _frase(numeros, texto, frases=None):
    prefijo = lista_de_lineas(numeros, frases)
    if prefijo:
        return Observacion(f"{prefijo}: {texto}", prefijo)
    return Observacion(texto[0].upper() + texto[1:])


def _lineas_de(items, filtro=lambda x: True):
    """todas las líneas de las observaciones que cumplen el filtro"""
    nums = []
    for it in items:
        if filtro(it):
            nums += it.get("nlineas") or ([it["nlinea"]] if it.get("nlinea") else [])
    return nums


def _cuantas(items, filtro):
    return sum(1 for it in items if filtro(it))


CITA_TEXTO = ("Cita sin referencia", "'et al", "Cita sin coma", "citas numéricas")
REFERENCIA = ("Referencia sin año", "sangría francesa", "orden alfabético", "DOI", "Referencia no citada",
              "no tiene entradas", "Referencia con datos", "Referencia cortada", "Título en mayúsculas",
              "Páginas mal escritas", "Artículo sin páginas", "Número de artículo")


def redactar(items, filas_orto, extras=(), tipo="proyecto", frases=None):
    """convierte las observaciones detalladas en la lista corta de la hoja de revisión"""
    fr = frases or FRASES
    o = fr["observaciones"]
    obs = []
    esquema = _cuantas(items, lambda x: x["categoria"] == "Estructura" and x["severidad"] == "Error")
    formato = [x for x in items if x["categoria"] == "Formato" and x["severidad"] == "Error"]
    guia = [x for x in items if "texto guía" in x["observacion"]]
    extension = [x for x in items if x["categoria"] == "Extensión" and x["severidad"] == "Error"]

    if esquema:
        obs.append(Observacion(o["esquema"].format(tipo=tipo)))
    de_formato = [x for x in formato if x not in guia]
    if de_formato:
        detalle = ", ".join(sorted({_que_formato(x["observacion"]) for x in de_formato}))
        obs.append(_frase(_lineas_de(de_formato), o["formato"].format(detalle=detalle), fr))

    # títulos de sección cortados o distintos a la plantilla: uno por uno, con su línea
    for x in items:
        if x["categoria"] == "Estructura" and "Título de sección incompleto" in x["observacion"]:
            obs.append(_frase(_lineas_de([x]), o["titulo_incompleto"].format(seccion=x["ubicacion"]), fr))
        elif x["categoria"] == "Estructura" and "Título distinto" in x["observacion"]:
            obs.append(_frase(_lineas_de([x]), o["titulo_distinto"].format(seccion=x["ubicacion"]), fr))
    if guia:
        obs.append(_frase(_lineas_de(guia), o["texto_guia"], fr))
    if filas_orto:
        obs.append(Observacion(o["revision_ortografica"]))

    es_cita = lambda x: x["categoria"] == "Citas" and any(c in x["observacion"] for c in CITA_TEXTO)  # noqa: E731
    if _cuantas(items, es_cita):
        obs.append(_frase(_lineas_de(items, es_cita), o["citas"], fr))

    # la tipografía (doble espacio, espacio antes de signo) entra en la misma línea
    tipografia = [x for x in items if x["categoria"] == "Ortografía" and x["severidad"] == "Advertencia"]
    if filas_orto or tipografia:
        nums = _lineas_de(filas_orto) + _lineas_de(tipografia)
        obs.append(_frase(nums, o["ortografia"], fr))

    es_ref = lambda x: x["categoria"] == "Citas" and any(c in x["observacion"] for c in REFERENCIA)  # noqa: E731
    if _cuantas(items, es_ref):
        obs.append(_frase(_lineas_de(items, es_ref), o["referencias"], fr))

    for x in extension:
        obs.append(Observacion(_redactar_extension(x["observacion"], o)))

    if [x for x in items if "No se pudo verificar las citas" in x["observacion"]]:
        obs.append(Observacion(o["sin_referencias"]))

    # Las de similitud ya vienen redactadas con el porcentaje y el máximo: se copian
    # tal cual, que es lo que el tesista necesita leer.
    for x in [i for i in items if i["categoria"] == "Similitud"]:
        obs.append(Observacion(x["observacion"]))

    for x in [i for i in items if i["categoria"] == "Revisor"]:
        texto = x["observacion"]
        nums = _lineas_de([x])
        obs.append(_frase(nums, texto[0].lower() + texto[1:], fr) if nums else Observacion(texto))

    obs += [Observacion(e.strip()) for e in extras if e.strip()]
    return obs


def _redactar_extension(texto, o=None):
    o = o or FRASES["observaciones"]
    for patron, clave in ((r"El documento tiene (\d+) páginas, máximo (\d+)", "extension_paginas"),
                          (r"El título tiene (\d+) palabras, máximo (\d+)", "extension_titulo"),
                          (r"Tiene (\d+) palabras clave, máximo (\d+)", "extension_palabras_clave")):
        m = re.search(patron, texto)
        if m:
            return o[clave].format(valor=m.group(1), maximo=m.group(2))
    return texto


def _que_formato(observacion):
    o = observacion.lower()
    if "margen" in o:
        return "márgenes"
    if "fuente" in o:
        return "tipo de letra"
    if "tamaño" in o:
        return "tamaño de letra"
    if "interlineado" in o:
        return "interlineado"
    if "justificar" in o:
        return "texto justificado"
    if "encabezado" in o:
        return "encabezado institucional"
    if "papel" in o:
        return "tamaño de papel"
    return "formato general"


# --------------------------------------------------------------------------- plantilla de Word

def _fuente(run):
    rpr = run._r.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rpr.append(rf)
    for a in ("w:ascii", "w:hAnsi", "w:cs"):
        rf.set(qn(a), FUENTE)


def txt(par, texto, negrita=False, tam=11, color=TINTA, mayus=False):
    r = par.add_run(texto)
    r.bold = negrita
    r.font.size = Pt(tam)
    r.font.name = FUENTE
    r.font.color.rgb = RGBColor.from_string(color)
    if mayus:
        r.font.all_caps = True
    _fuente(r)
    return r


def sombrear(par, color):
    ppr = par._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:fill"), color)
    ppr.append(shd)


def borde_inferior(par, color=LINEA, grosor=6):
    ppr = par._p.get_or_add_pPr()
    bordes = OxmlElement("w:pBdr")
    b = OxmlElement("w:bottom")
    b.set(qn("w:val"), "single")
    b.set(qn("w:sz"), str(grosor))
    b.set(qn("w:space"), "4")
    b.set(qn("w:color"), color)
    bordes.append(b)
    ppr.append(bordes)


def crear_plantilla_hoja(ruta=None):
    """
    La plantilla que trae el programa, con el diseño de la Dirección de Investigación.
    Si se da 'ruta' se guarda ahí; siempre devuelve el Document.
    """
    doc = Document()
    s = doc.sections[0]
    s.page_width, s.page_height = Cm(21), Cm(29.7)
    s.left_margin = s.right_margin = Cm(2.5)
    s.top_margin = Cm(2.2)
    s.bottom_margin = Cm(2.2)
    normal = doc.styles["Normal"]
    normal.font.name = FUENTE
    normal.font.size = Pt(11)
    normal.paragraph_format.space_after = Pt(8)

    p = doc.add_paragraph()
    txt(p, "REVISIÓN PGI {{FECHA_REVISION}} – {{N_REVISION}}", negrita=True, tam=12)

    # ficha oscura del proyecto
    for texto, estilo in (("{{CODIGO?}}", dict(negrita=True, tam=11, color="FFFFFF")),
                          ("{{TITULO}}", dict(tam=8, color="D9D9D9", mayus=True)),
                          ("{{ETAPA}}   ·   1 tesista   ·   {{ARCHIVO}}", dict(tam=8, color="BFBFBF"))):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.left_indent = Cm(0.2)
        sombrear(p, "1A1A1A")
        txt(p, texto, **estilo)
    p.paragraph_format.space_after = Pt(14)

    for etiqueta, marcador in (("Tesista: ", "{{TESISTA?}}"), ("Asesor: ", "{{ASESOR}}"),
                               ("Turnitin: ", "{{SIMILITUD?}}")):
        p = doc.add_paragraph()
        txt(p, etiqueta, negrita=True)
        txt(p, marcador, tam=10 if "SIMILITUD" in marcador else 11)

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    txt(p, "OBSERVACIONES", negrita=True)
    borde_inferior(p)

    q = doc.add_paragraph()
    q.paragraph_format.space_after = Pt(6)
    q.alignment = WD_ALIGN_PARAGRAPH.LEFT
    txt(q, "{{OBSERVACIONES}}")

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(36)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    txt(p, "_" * 42, tam=10, color=LINEA)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(0)
    txt(p, "{{REVISOR}}", negrita=True, tam=10)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    txt(p, "Dirección de Investigación – FINESI", tam=9, color=SUAVE)
    if ruta:
        os.makedirs(os.path.dirname(os.path.abspath(ruta)), exist_ok=True)
        doc.save(ruta)
    return doc


MARCADOR = re.compile(r"\{\{\s*([A-Za-z_]+)\s*(\?)?\s*\}\}")
# lo que se muestra cuando un dato obligatorio viene vacío
SI_FALTA = {"ASESOR": "(completar)", "TESISTA": "(completar)", "TITULO": "(título del proyecto)",
            "REVISOR": "Revisor"}


def _valores(datos):
    return {
        "FECHA_REVISION": datos.get("fecha_revision", ""),
        "N_REVISION": datos.get("n_revision", ""),
        "CODIGO": datos.get("codigo", ""),
        "TITULO": datos.get("proyecto", ""),
        "TIPO": datos.get("tipo", ""),
        "ETAPA": datos.get("etapa") or "Etapa 2: Revisión de formato de proyecto",
        "ARCHIVO": datos.get("archivo", ""),
        "TESISTA": datos.get("tesista", ""),
        "ASESOR": datos.get("asesor", ""),
        "FECHA_PRESENTACION": datos.get("fecha_subida", ""),
        "SIMILITUD": datos.get("similitud", ""),
        "REVISOR": datos.get("revisor", ""),
    }


def _reemplazar_en_parrafo(par, valores):
    """
    Reemplaza los marcadores aunque Word los haya partido en varios fragmentos (runs)
    al editar la plantilla: el texto nuevo queda con el formato del fragmento donde
    empieza el marcador. Devuelve False si el párrafo debe borrarse ({{DATO?}} vacío).
    """
    runs = par.runs
    completo = "".join(r.text for r in runs)
    hallados = list(MARCADOR.finditer(completo))
    if not hallados:
        return True
    for m in hallados:
        if m.group(2) and not valores.get(m.group(1).upper(), ""):
            return False
    inicios, total = [], 0
    for r in runs:
        inicios.append(total)
        total += len(r.text)

    def run_de(pos):
        k = 0
        for i, ini in enumerate(inicios):
            if ini <= pos:
                k = i
        return k

    for m in reversed(hallados):
        clave = m.group(1).upper()
        valor = valores.get(clave, "") or ("" if m.group(2) else SI_FALTA.get(clave, ""))
        a, b = run_de(m.start()), run_de(m.end() - 1)
        ta, tb = runs[a].text, runs[b].text
        ini_a, ini_b = m.start() - inicios[a], m.end() - inicios[b]
        if a == b:
            runs[a].text = ta[:ini_a] + valor + ta[ini_b:]
        else:
            runs[a].text = ta[:ini_a] + valor
            for k in range(a + 1, b):
                runs[k].text = ""
            runs[b].text = tb[ini_b:]
    # un separador que quedó colgando porque el dato estaba vacío: "… – " o "· "
    ultimo = next((r for r in reversed(runs) if r.text), None)
    if ultimo is not None:
        ultimo.text = re.sub(r"\s+[–\-·]\s*$", "", ultimo.text)
    return True


def _parrafos(doc):
    """todos los párrafos: cuerpo, tablas, encabezados y pies"""
    contenedores = [doc]
    for s in doc.sections:
        contenedores += [s.header, s.footer]
    vistos = set()
    for c in contenedores:
        for p in c.paragraphs:
            yield p
        for t in c.tables:
            for fila in t.rows:
                for celda in fila.cells:
                    if id(celda._tc) in vistos:
                        continue
                    vistos.add(id(celda._tc))
                    yield from celda.paragraphs


def _escribir_observacion(molde, observacion, antes_de):
    """copia el párrafo molde (el de {{OBSERVACIONES}}) con el texto de una observación"""
    nuevo = copy.deepcopy(molde._p)
    antes_de.addprevious(nuevo)
    runs = nuevo.findall(qn("w:r"))
    base = runs[0] if runs else None
    for r in runs:
        nuevo.remove(r)

    def agregar(texto, negrita=False):
        r = copy.deepcopy(base) if base is not None else OxmlElement("w:r")
        for t in r.findall(qn("w:t")):
            r.remove(t)
        rpr = r.find(qn("w:rPr"))
        if negrita:
            if rpr is None:
                rpr = OxmlElement("w:rPr")
                r.insert(0, rpr)
            # la plantilla puede traer la negrita desactivada (w:b w:val="0"): se reemplaza
            for b in rpr.findall(qn("w:b")):
                rpr.remove(b)
            rpr.append(OxmlElement("w:b"))
        t = OxmlElement("w:t")
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        t.text = texto
        r.append(t)
        nuevo.append(r)

    prefijo = getattr(observacion, "prefijo", "")
    if prefijo and observacion.startswith(prefijo):
        agregar(prefijo + ": ", negrita=True)
        agregar(observacion[len(prefijo):].lstrip(": ").lstrip())
    else:
        agregar(str(observacion))


def escribir_resumen(ruta_salida, datos, observaciones, plantilla=None, frases=None):
    """
    Llena la plantilla de la hoja de revisión. 'plantilla' es la ruta del .docx con
    marcadores; si no existe se usa la del programa.
    """
    doc = Document(plantilla) if plantilla and os.path.exists(plantilla) else crear_plantilla_hoja()
    valores = _valores(datos)
    molde = None
    for p in list(_parrafos(doc)):
        if "{{OBSERVACIONES}}" in p.text.replace(" ", ""):
            molde = molde or p
            continue
        if not _reemplazar_en_parrafo(p, valores):
            p._p.getparent().remove(p._p)

    lista = list(observaciones) or [Observacion((frases or FRASES)["observaciones"]["sin_observaciones"])]
    if molde is None:
        # la plantilla editada perdió el marcador: las observaciones van al final
        molde = doc.add_paragraph()
        for o in lista:
            _escribir_observacion(molde, o, molde._p)
    else:
        for o in lista:
            _escribir_observacion(molde, o, molde._p)
    molde._p.getparent().remove(molde._p)
    doc.save(ruta_salida)
