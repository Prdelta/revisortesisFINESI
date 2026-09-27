"""
Extensión del documento: número de páginas, palabras del título y cantidad de
palabras clave.
"""
import re

from docx.text.paragraph import Paragraph

from .utils import celdas_unicas, corto


# Palabras por página de un documento a 10 pt, interlineado sencillo y con algunas
# tablas: un proyecto real de 5818 palabras ocupó 10 páginas en Word (582 por página).
# Se usa un valor algo menor para que la estimación, si se equivoca, lo haga hacia arriba.
PALABRAS_POR_PAGINA = 550
# El dato de Word se acepta si no es menor que esta fracción de la estimación
PLAUSIBLE = 0.6


def contar_paginas(ruta, mapa=None, palabras=None, por_pagina=PALABRAS_POR_PAGINA):
    """
    Devuelve (páginas, método, exacto).

    Con Word o LibreOffice el conteo es el del PDF generado. Sin ellos queda el dato <Pages> que
    Word guarda en el archivo, que no siempre se actualiza: un proyecto de 10 páginas
    lo traía en 1, y así una tesis de 30 páginas pasaba sin aviso. Por eso se lo
    contrasta con una estimación por número de palabras y, si no es creíble, manda la
    estimación.
    """
    if mapa is not None and mapa.paginas:
        return mapa.paginas, getattr(mapa, "motor", "") or "LibreOffice", True
    estimado = max(1, round(palabras / por_pagina)) if palabras else None
    guardado = None
    try:
        import zipfile
        with zipfile.ZipFile(ruta) as z:
            xml = z.read("docProps/app.xml").decode("utf8", "ignore")
        m = re.search(r"<Pages>(\d+)</Pages>", xml)
        if m and int(m.group(1)) > 0:
            guardado = int(m.group(1))
    except Exception:
        pass
    if guardado and (estimado is None or guardado >= estimado * PLAUSIBLE):
        return guardado, "dato guardado por Word al último guardado", False
    if estimado:
        detalle = f"estimado por {palabras} palabras (~{por_pagina} por página)"
        if guardado:
            detalle += f"; el archivo dice {guardado}, dato desactualizado"
        return estimado, detalle, False
    return None, None, False


def contar_palabras(bloques):
    """
    Palabras del cuerpo, contadas como las cuenta Turnitin: párrafos y tablas.

    Las celdas combinadas se cuentan una sola vez. Contándolas una vez por columna,
    el total se inflaba (un encabezado combinado a tres columnas valía el triple) y
    la comparación con las palabras que declara el informe acusaba en falso al tesista
    de haber pasado por Turnitin otro documento.
    """
    total = 0
    for b in bloques:
        if isinstance(b, Paragraph):
            total += len(b.text.split())
        else:                       # tabla
            for celda in celdas_unicas(b):
                total += len(celda.text.split())
    return total


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
        pags, metodo, exacto = contar_paginas(ruta, mapa, contar_palabras(bloques),
                                              e.get("palabras_por_pagina", PALABRAS_POR_PAGINA))
        if pags is None:
            obs.add("Extensión", "Revisar", "Documento", "No se pudo contar las páginas", "Verificar manualmente en Word.")
        elif pags > e["max_paginas"]:
            if exacto or metodo.startswith("dato guardado"):
                obs.add("Extensión", "Error", "Documento", f"El documento tiene {pags} páginas, máximo {e['max_paginas']}",
                        f"Conteo: {metodo}. Puede variar en una página respecto a lo que muestra Word.")
            else:
                obs.add("Extensión", "Revisar", "Documento",
                        f"El documento tendría unas {pags} páginas, máximo {e['max_paginas']}",
                        f"Conteo aproximado: {metodo}. Verificar en Word.")

    tit = texto_de_seccion(bloques, rangos, e.get("seccion_titulo", "Título"))
    if tit and e.get("titulo_max_palabras"):
        n = len(tit.split())
        if n > e["titulo_max_palabras"]:
            obs.add("Extensión", "Error", e.get("seccion_titulo", "Título"), f"El título tiene {n} palabras, máximo {e['titulo_max_palabras']}", corto(tit, 120))

    kw = texto_de_seccion(bloques, rangos, e.get("seccion_palabras_clave", "Palabras claves"))
    if kw and e.get("palabras_clave_max"):
        # si el tesista puso "Keywords:" en ingles, contamos solo la primera linea (en espanol)
        linea = [x for x in kw.split("\n") if x.strip()][0]
        linea = re.sub(r"^\s*(palabras? claves?|keywords)\s*:?", "", linea, flags=re.I)
        items = [k for k in re.split(r"[;,]", linea) if k.strip()]
        if len(items) > e["palabras_clave_max"]:
            obs.add("Extensión", "Error", e.get("seccion_palabras_clave", "Palabras claves"), f"Tiene {len(items)} palabras clave, máximo {e['palabras_clave_max']}", corto(linea, 120))


