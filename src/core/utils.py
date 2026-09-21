"""
Piezas compartidas por los módulos de revisión: normalización de texto, ubicación
legible de una observación, recorrido del documento en orden real y resolución de
los estilos efectivos de Word (run > estilo > Normal > docDefaults).
"""
import re
import unicodedata

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

__all__ = ["norm", "corto", "limpiar_titulo", "juntar", "ubic", "Obs", "SEVERIDADES",
           "bloques_en_orden", "Estilos"]


def norm(txt):
    """minusculas, sin tildes, sin signos, espacios simples"""
    txt = unicodedata.normalize("NFD", txt.lower())
    txt = "".join(c for c in txt if unicodedata.category(c) != "Mn")
    txt = re.sub(r"[^a-z0-9 ]+", " ", txt)   # la ñ ya se volvió n al quitar los diacríticos
    return re.sub(r"\s+", " ", txt).strip()


def corto(txt, n=70):
    txt = re.sub(r"\s+", " ", txt).strip()
    return txt if len(txt) <= n else txt[:n] + "..."


ROMANO = r"(?:[IVXLC]+|[ivx]+)"
PREFIJO_NUM = re.compile(rf"^\s*(?:{ROMANO}|\d+(?:\.\d+)*|[a-z])[\.\)\-:]?\s+")


def limpiar_titulo(txt):
    """quita numeracion inicial y lo que viene desde el primer parentesis o dos puntos"""
    txt = PREFIJO_NUM.sub("", txt.strip())
    txt = re.split(r"[\(:]", txt, maxsplit=1)[0]
    return norm(txt)




def juntar(lugares):
    """une ubicaciones sin repetir el nombre de la seccion"""
    salida, previa = [], None
    for l in lugares:
        sec, _, resto = l.partition(": ")
        salida.append(resto if sec == previa and resto else l)
        previa = sec
    return "; ".join(salida)


def ubic(i, rangos, bloques, txt=None, n=40):
    """ubicacion legible: seccion del esquema + inicio del texto (el numero de parrafo no se ve en Word)"""
    sec = "Inicio del documento"
    for nombre, (ini, fin) in rangos.items():
        if ini <= i < fin:
            sec = nombre
            break
    if txt is None and i < len(bloques) and isinstance(bloques[i], Paragraph):
        txt = bloques[i].text
    if i < len(bloques) and isinstance(bloques[i], Table):
        return f"{sec} (tabla): '{corto(txt or '', n)}'"
    return f"{sec}: '{corto(txt or '', n)}'" if txt else sec


SEVERIDADES = ("Error", "Advertencia", "Revisar")


class Obs:
    """
    Observaciones de una revisión. 'ajustes' permite bajarle (o subirle) la severidad
    a una categoría entera desde el archivo de reglas, sin tocar código:

        severidades:
          Citas: Advertencia
    """

    def __init__(self, ajustes=None):
        self.items = []
        self.ajustes = {str(k).strip().lower(): v for k, v in (ajustes or {}).items()
                        if v in SEVERIDADES}

    def add(self, categoria, severidad, ubicacion, observacion, detalle=""):
        severidad = self.ajustes.get(str(categoria).strip().lower(), severidad)
        self.items.append(dict(categoria=categoria, severidad=severidad, ubicacion=ubicacion,
                               observacion=observacion, detalle=detalle))



def bloques_en_orden(doc):
    """parrafos y tablas del cuerpo en el orden real del documento"""
    for hijo in doc.element.body.iterchildren():
        if hijo.tag == qn("w:p"):
            yield Paragraph(hijo, doc)
        elif hijo.tag == qn("w:tbl"):
            yield Table(hijo, doc)


class Estilos:
    """resuelve fuente, tamano, interlineado y alineacion efectivos (run > estilo > Normal > docDefaults)"""

    def __init__(self, doc):
        self.doc = doc
        self.tema = {}
        try:
            for rel in doc.part.package.parts:
                if rel.partname.endswith("theme1.xml"):
                    xml = rel.blob.decode("utf8", "ignore")
                    may = re.search(r'<a:majorFont><a:latin typeface="([^"]*)"', xml)
                    men = re.search(r'<a:minorFont><a:latin typeface="([^"]*)"', xml)
                    if may:
                        self.tema["major"] = may.group(1)
                    if men:
                        self.tema["minor"] = men.group(1)
        except Exception:
            pass
        st = doc.styles.element
        self.def_fuente = self._fuente_de(st.find(qn("w:docDefaults") + "/" + qn("w:rPrDefault") + "/" + qn("w:rPr")))
        sz = st.find(qn("w:docDefaults") + "/" + qn("w:rPrDefault") + "/" + qn("w:rPr") + "/" + qn("w:sz"))
        self.def_tam = int(sz.get(qn("w:val"))) / 2 if sz is not None else 10.0

    def _fuente_de(self, rpr):
        if rpr is None:
            return None
        rf = rpr.find(qn("w:rFonts"))
        if rf is None:
            return None
        if rf.get(qn("w:ascii")):
            return rf.get(qn("w:ascii"))
        tema = rf.get(qn("w:asciiTheme")) or rf.get(qn("w:hAnsiTheme"))
        if tema:
            return self.tema.get("major" if "major" in tema.lower() else "minor")
        return None

    def _cadena(self, estilo):
        while estilo is not None:
            yield estilo
            estilo = estilo.base_style

    def fuente(self, run, par):
        f = self._fuente_de(run._r.rPr)
        if f:
            return f
        if run.style is not None:
            for s in self._cadena(run.style):
                f = self._fuente_de(s.element.rPr)
                if f:
                    return f
        for s in self._cadena(par.style):
            f = self._fuente_de(s.element.rPr)
            if f:
                return f
        return self.def_fuente

    def tamano(self, run, par):
        if run.font.size:
            return run.font.size.pt
        if run.style is not None:
            for s in self._cadena(run.style):
                if s.font.size:
                    return s.font.size.pt
        for s in self._cadena(par.style):
            if s.font.size:
                return s.font.size.pt
        return self.def_tam

    def interlineado(self, par):
        v = par.paragraph_format.line_spacing
        if v is None:
            for s in self._cadena(par.style):
                if s.paragraph_format.line_spacing is not None:
                    v = s.paragraph_format.line_spacing
                    break
        if v is None:
            return 1.0
        if hasattr(v, "pt"):  # interlineado exacto en puntos
            return round(v.pt / 12.0, 2)
        return round(float(v), 2)

    def alineacion(self, par):
        if par.alignment is not None:
            return par.alignment
        for s in self._cadena(par.style):
            if s.paragraph_format.alignment is not None:
                return s.paragraph_format.alignment
        return WD_ALIGN_PARAGRAPH.LEFT


