"""
Extensión del documento: número de páginas, palabras del título y cantidad de
palabras clave.
"""
import re

from docx.text.paragraph import Paragraph

from .utils import corto


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
    except Exception:
        pass   # sin el dato de Word se informa más abajo que no se pudo contar
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


