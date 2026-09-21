"""
Informe de similitud de Turnitin (PDF).

Además del .docx, lo único que el tesista entrega es el PDF que devuelve Turnitin
después de pasar el documento. De ahí salen el índice de similitud, el porcentaje de
texto detectado como generado por IA, las fuentes que más coinciden y los filtros que
se aplicaron al medir. Este módulo lee ese PDF y traduce esos datos a observaciones.

Se leen los formatos que Turnitin ha usado, en español y en inglés:
  - "Informe de similitud" con "Portada de integridad" (Turnitin actual).
  - "Informe de originalidad" de Feedback Studio (el clásico, con ÍNDICE DE SIMILITUD).
  - "Recibo digital", que suele venir pegado al informe y trae los datos de la entrega.

El PDF se lee por posición, no por orden de lectura: Turnitin pone los rótulos en una
fila y los números en la de abajo, en columnas. Por eso cada porcentaje se asigna al
rótulo que tiene más cerca en la página y no al que le sigue en el texto; leerlo de
corrido intercambia el porcentaje de IA con el de similitud.

Lo que no se puede leer se dice. Si no aparece el índice de similitud, se genera una
observación pidiendo verificarlo a mano: el modo de fallo peligroso no es el ruido,
es el silencio.
"""
import os
import re
import unicodedata
from dataclasses import dataclass, field
from datetime import datetime

from .utils import corto

__all__ = ["Informe", "analizar", "revisar_similitud", "resumen", "opciones_similitud",
           "nombre_comparable", "PREDETERMINADO"]


# --------------------------------------------------------------------------- texto del PDF

def _plano(texto):
    """minúsculas y sin tildes, conservando la longitud (las columnas importan)"""
    salida = []
    for c in texto:
        d = unicodedata.normalize("NFD", c.lower())
        d = "".join(x for x in d if unicodedata.category(x) != "Mn")
        salida.append(d[:1] or " ")
    return "".join(salida)


def _paginas_de(ruta):
    """texto de cada página, respetando la disposición en columnas; (páginas, error)"""
    try:
        from pypdf import PdfReader
    except ImportError:
        return [], "falta la librería pypdf (reinstalar con: pip install -r requirements.txt)"
    try:
        lector = PdfReader(ruta)
        if lector.is_encrypted:
            try:
                lector.decrypt("")
            except Exception:
                return [], "el PDF está protegido con contraseña"
        paginas = []
        for pagina in lector.pages:
            try:
                texto = pagina.extract_text(extraction_mode="layout") or ""
            except Exception:
                try:
                    texto = pagina.extract_text() or ""
                except Exception:
                    texto = ""
            paginas.append(texto)
    except Exception as ex:
        return [], f"no se pudo abrir el PDF ({type(ex).__name__}: {ex})"
    if not any(p.strip() for p in paginas):
        return paginas, ("el PDF no tiene texto que se pueda leer; puede ser un escaneo o una "
                         "captura de pantalla del informe, no el PDF que descarga Turnitin")
    return paginas, ""


RE_SOLO_PCT = re.compile(r"^[\s%]*%[\s%]*$")
RE_NUMERO = re.compile(r"\d+(?:[.,]\d)?")


def _fusionar_porcentajes(lineas):
    """
    En la portada, Turnitin pone el número en una fila y el signo % en la de abajo,
    cada uno bajo su columna. Se junta el % con el número que tiene encima para que
    el porcentaje se pueda leer como uno solo.
    """
    salida = []
    for linea in lineas:
        if salida and "%" in linea and RE_SOLO_PCT.match(linea) and RE_NUMERO.search(salida[-1]):
            previa, usados, posiciones = salida[-1], set(), []
            for marca in re.finditer("%", linea):
                col, mejor = marca.start(), None
                for n in RE_NUMERO.finditer(previa):
                    if n.end() in usados:
                        continue
                    d = min(abs(n.start() - col), abs(n.end() - col))
                    if mejor is None or d < mejor[0]:
                        mejor = (d, n.end())
                if mejor and mejor[0] <= 12:
                    usados.add(mejor[1])
                    posiciones.append(mejor[1])
            for pos in sorted(posiciones, reverse=True):
                previa = previa[:pos] + "%" + previa[pos:]
            salida[-1] = previa
            continue
        salida.append(linea)
    return salida


# --------------------------------------------------------------------------- porcentajes

# Los rótulos van en plural a propósito: en el informe clásico, "FUENTES DE INTERNET"
# es el total y "Fuente de Internet" es el tipo de cada fuente de la lista de abajo.
ETIQUETAS = {
    "similitud": ("similitud general", "indice de similitud", "similitud total",
                  "porcentaje de similitud", "overall similarity", "similarity index"),
    "ia": ("deteccion de ia", "ia detectada", "texto generado por ia", "generado por ia",
           "escritura de ia", "escritura con ia", "ai detection", "ai writing"),
    "internet": ("fuentes de internet", "internet sources"),
    "publicaciones": ("publicaciones", "publications"),
    "estudiantes": ("trabajos del estudiante", "trabajos entregados", "trabajos de estudiantes",
                    "student papers", "submitted works"),
}

# En el informe clásico los rótulos vienen partidos en dos líneas ("ÍNDICE DE" /
# "SIMILITUD"). Estos trozos solo se buscan para los porcentajes que no se hayan
# encontrado con el rótulo completo: sueltos son demasiado ambiguos.
ETIQUETAS_PARTIDAS = {
    "similitud": ("indice de", "similitud", "similarity"),
    "ia": ("de ia", "ia", "ai"),
    "internet": ("internet",),
    "publicaciones": ("publicacion", "publication"),
    "estudiantes": ("trabajos del", "trabajos", "estudiante", "student", "submitted"),
}

RE_PORCENTAJE = re.compile(r"(?<![\d.,])(\d{1,3}(?:[.,]\d)?)\s*%")


def _ubicar_etiquetas(lineas, diccionario):
    """[(campo, fila, col_inicio, col_fin)] de cada rótulo reconocido"""
    hallados = []
    for fila, linea in enumerate(lineas):
        plano = _plano(linea)
        for campo, etiquetas in diccionario.items():
            for et in etiquetas:
                desde = 0
                while True:
                    k = plano.find(et, desde)
                    if k < 0:
                        break
                    hallados.append((campo, fila, k, k + len(et)))
                    desde = k + 1
    return hallados


def _ubicar_porcentajes(lineas):
    """[(fila, columna, valor)] de cada porcentaje de la página"""
    return [(fila, m.start(), float(m.group(1).replace(",", ".")))
            for fila, linea in enumerate(lineas)
            for m in RE_PORCENTAJE.finditer(linea)]


def _costo(etiqueta, porcentaje):
    """qué tan lejos está el porcentaje del rótulo; None si no pueden ir juntos"""
    _campo, fila_e, ini, fin = etiqueta
    fila_p, col_p, _valor = porcentaje
    salto = fila_p - fila_e
    if salto == 0:       # en la misma fila: primero el que va a la derecha del rótulo
        if col_p >= fin:
            return (col_p - fin) * 0.02
        return 1.0 + (ini - col_p) * 0.02
    if 0 < salto <= 2:   # el número va debajo del rótulo (portada de integridad)
        return 2.0 + (salto - 1) * 2 + abs(col_p - ini) * 0.05
    if -3 <= salto < 0:  # el número va encima del rótulo (informe clásico)
        return 2.5 + (-salto - 1) * 2 + abs(col_p - ini) * 0.05
    return None


def _emparejar(lineas, diccionario, porcentajes, valores, usados_pct):
    """asigna a cada campo que falte el porcentaje más cercano que quede libre"""
    pares = []
    for e in _ubicar_etiquetas(lineas, diccionario):
        if e[0] in valores:
            continue
        for i, p in enumerate(porcentajes):
            c = _costo(e, p)
            if c is not None:
                pares.append((c, e[1], i, e, p[2]))
    pares.sort(key=lambda x: (x[0], x[1], x[2]))
    for _c, _fila, i, e, valor in pares:
        if e[0] in valores or i in usados_pct:
            continue
        valores[e[0]] = valor
        usados_pct.add(i)


# Donde empieza la lista de fuentes una por una. De ahí para abajo los porcentajes
# son de cada fuente, no los totales, y no deben emparejarse con los rótulos.
INICIO_LISTA = ("fuentes primarias", "primary sources", "coincidencias principales",
                "top matches", "principales fuentes coincidentes")


def _recortar_resumen(lineas):
    for k, linea in enumerate(lineas):
        plano = _plano(linea).strip()
        if any(plano.startswith(m) for m in INICIO_LISTA):
            return lineas[:k]
    return lineas


def _asignar(lineas):
    """
    Empareja cada rótulo con el porcentaje que tiene más cerca. Se resuelve de la
    pareja más cercana a la más lejana y ningún porcentaje se usa dos veces: así
    'Detección de IA   Similitud general' / '8%   23%' queda cada uno en su sitio.
    Primero se usan los rótulos completos; los trozos sueltos solo cubren lo que
    quedó sin encontrar, porque el informe clásico parte los rótulos en dos líneas.
    """
    lineas = _recortar_resumen(lineas)
    porcentajes = _ubicar_porcentajes(lineas)
    valores, usados_pct = {}, set()
    _emparejar(lineas, ETIQUETAS, porcentajes, valores, usados_pct)
    _emparejar(lineas, ETIQUETAS_PARTIDAS, porcentajes, valores, usados_pct)
    return valores


MARCAS_RESUMEN = ETIQUETAS["similitud"] + (
    "informe de originalidad", "originality report", "portada de integridad",
    "integrity overview", "descripcion general de integridad")


def _region_resumen(paginas):
    """las páginas donde Turnitin pone el cuadro de porcentajes"""
    for k, texto in enumerate(paginas):
        plano = _plano(texto)
        if any(et in plano for et in MARCAS_RESUMEN):
            return paginas[k:k + 2]
    return paginas[:2]


# --------------------------------------------------------------------------- datos de la entrega

MARCAS = ("turnitin", "informe de originalidad", "originality report", "indice de similitud",
          "similarity index", "similitud general", "overall similarity", "trn:oid",
          "identificador de la entrega", "submission id", "recibo digital", "digital receipt",
          "portada de integridad", "integrity overview", "fuentes primarias", "primary sources")


def _buscar(texto, *patrones):
    for p in patrones:
        m = re.search(p, texto, re.I)
        if m and (m.group(1) or "").strip():
            return re.sub(r"\s+", " ", m.group(1)).strip(" .:;, ")
    return ""


def _entero(texto):
    digitos = re.sub(r"[^\d]", "", texto or "")
    return int(digitos) if digitos else None


def _valor_tras(lineas, *etiquetas):
    """valor de un rótulo, esté después de los dos puntos o en la línea siguiente"""
    for k, linea in enumerate(lineas):
        plano = _plano(linea).strip()
        for et in etiquetas:
            if not plano.startswith(et):
                continue
            resto = linea.strip()[len(et):].strip(" : ")
            if resto:
                return re.sub(r"\s+", " ", resto)
            for siguiente in lineas[k + 1:k + 3]:
                if siguiente.strip():
                    return re.sub(r"\s+", " ", siguiente.strip())
    return ""


ENCENDIDO = ("activo", "activado", "si", "on", "yes", "excluido", "excluida")
APAGADO = ("apagado", "desactivado", "no", "off", "incluido", "incluida")


def _estado(plano, etiqueta):
    """lee un 'Excluir bibliografía   Activo' del informe clásico"""
    k = plano.find(etiqueta)
    if k < 0:
        return None
    cola = plano[k + len(etiqueta):k + len(etiqueta) + 60]
    for palabra in re.sub(r"[^a-z ]+", " ", cola).split():
        if palabra in ENCENDIDO:
            return True
        if palabra in APAGADO:
            return False
    return None


def _filtros(plano):
    """qué dejó fuera Turnitin al medir: bibliografía, texto citado, coincidencias cortas"""
    filtros = {"bibliografia": None, "citas": None, "coincidencias_menores": None}
    for clave, etiquetas in (("bibliografia", ("excluir bibliografia", "exclude bibliography")),
                             ("citas", ("excluir citas", "exclude quotes"))):
        for et in etiquetas:
            estado = _estado(plano, et)
            if estado is not None:
                filtros[clave] = estado
                break

    # Turnitin actual: una lista bajo "Filtrado desde el informe"
    k = plano.find("filtrado desde el informe")
    if k < 0:
        k = plano.find("filtered from the report")
    if k >= 0:
        bloque = plano[k:k + 400]
        corte = re.search(r"(fuentes principales|top sources|coincidencias principales|"
                          r"descripcion general|marcas de integridad|integrity flags)", bloque)
        if corte:
            bloque = bloque[:corte.start()]
        if filtros["bibliografia"] is None:
            filtros["bibliografia"] = "bibliografia" in bloque or "bibliography" in bloque
        if filtros["citas"] is None:
            filtros["citas"] = ("texto citado" in bloque or "quoted text" in bloque
                                or "citas" in bloque or "quotes" in bloque)

    m = (re.search(r"coincidencias menores \(menos de (\d+) palabras?\)", plano)
         or re.search(r"small matches \(less than (\d+) words?\)", plano)
         or re.search(r"excluir coincidencias[^\d]{0,40}(\d+)", plano)
         or re.search(r"exclude matches[^\d]{0,40}(\d+)", plano))
    if m:
        filtros["coincidencias_menores"] = int(m.group(1))
    return filtros


AGREGADOS = tuple(et for grupo in ETIQUETAS.values() for et in grupo) + (
    "fuentes principales", "top sources", "coincidencias principales", "top matches",
    "fuentes primarias", "primary sources", "filtrado desde el informe", "filtered from the report",
    "coincidencias menores", "small matches", "excluir", "exclude", "identificador de la entrega",
    "submission id", "recuento de", "word count", "character count", "total de", "total paginas",
    "pagina ", "page ", "informe de", "report", "turnitin")


# En el informe clásico, la fuente ocupa dos líneas: el nombre arriba y debajo el
# tipo con el porcentaje. Si en la línea del porcentaje solo está el tipo, el nombre
# hay que ir a buscarlo a la de arriba.
TIPOS_FUENTE = ("fuente de internet", "internet source", "publicacion", "publication",
                "trabajo del estudiante", "student paper", "submitted works", "internet",
                "trabajo entregado", "trabajos entregados", "publicaciones")


def _limpiar_fuente(texto):
    texto = re.sub(r"^\s*\d{1,2}[\.\)]?\s+", "", texto)      # el número de orden
    return re.sub(r"\s{2,}", " ", texto).strip(" .- ")


def _fuentes(paginas, tope):
    """[(descripción, porcentaje)] de las fuentes individuales que lista el informe"""
    hallados = {}
    for texto in paginas:
        anterior = ""
        for linea in texto.split("\n"):
            m = RE_PORCENTAJE.search(linea)
            if not m:
                if linea.strip():
                    anterior = linea
                continue
            previa, anterior = anterior, linea
            valor = float(m.group(1).replace(",", "."))
            if tope is not None and valor > tope + 0.001:
                continue        # una sola fuente no puede pasar el total
            nombre = _limpiar_fuente(RE_PORCENTAJE.sub(" ", linea))
            if _plano(nombre).strip(" .:,") in TIPOS_FUENTE and previa.strip():
                tipo, nombre = nombre, _limpiar_fuente(previa)
                if len(nombre) > 3:
                    nombre = f"{nombre} ({tipo.lower()})"
            plano = _plano(nombre)
            if len(re.sub(r"[^a-z0-9]", "", plano)) < 5:
                continue
            if any(plano.startswith(a.strip()) for a in AGREGADOS):
                continue
            if any(a in plano for a in ("similitud", "similarity", "deteccion de ia", "ai detection")):
                continue
            clave = plano[:60]
            if clave not in hallados or hallados[clave][1] < valor:
                hallados[clave] = (corto(nombre, 110), valor)
    return sorted(hallados.values(), key=lambda x: -x[1])[:15]


MESES = {"ene": 1, "feb": 2, "mar": 3, "abr": 4, "may": 5, "jun": 6, "jul": 7, "ago": 8,
         "set": 9, "sep": 9, "oct": 10, "nov": 11, "dic": 12,
         "jan": 1, "apr": 4, "aug": 8, "dec": 12}


def _fecha_segura(anio, mes, dia):
    try:
        return datetime(anio, mes, dia)
    except ValueError:
        return None


def fecha_informe(texto):
    """datetime de la fecha de entrega que declara el informe, o None"""
    if not texto:
        return None
    t = _plano(texto)
    m = re.search(r"(\d{4})-(\d{2})-(\d{2})", t)
    if m:
        return _fecha_segura(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    m = re.search(r"(\d{1,2})[-/](\d{1,2})[-/](\d{4})", t)
    if m:
        return _fecha_segura(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    m = re.search(r"(\d{1,2})\s*(?:de\s+)?[-/ ]?\s*([a-z]{3,10})\.?\s*(?:de\s+)?[-/ ]?\s*(\d{4})", t)
    if m:
        mes = MESES.get(m.group(2)[:3])
        if mes:
            return _fecha_segura(int(m.group(3)), mes, int(m.group(1)))
    return None


# --------------------------------------------------------------------------- informe

@dataclass
class Informe:
    """lo que declara el PDF de Turnitin; lo que no se pudo leer queda en None"""
    ruta: str = ""
    archivo: str = ""
    es_turnitin: bool = False
    formato: str = ""               # similitud | originalidad | recibo
    similitud: float = None
    ia: float = None
    internet: float = None
    publicaciones: float = None
    estudiantes: float = None
    fuentes: list = field(default_factory=list)
    filtros: dict = field(default_factory=dict)
    entrega: str = ""
    titulo: str = ""
    autor: str = ""
    archivo_entregado: str = ""
    palabras: int = None
    caracteres: int = None
    paginas: int = None
    paginas_pdf: int = 0
    fecha: str = ""
    error: str = ""
    texto: str = ""


def analizar(ruta):
    """lee el PDF y devuelve lo que declara el informe de Turnitin"""
    paginas, error = _paginas_de(ruta)
    informe = analizar_paginas(paginas, ruta=ruta, archivo=os.path.basename(ruta))
    if error:
        informe.error = error
    return informe


def analizar_paginas(paginas, ruta="", archivo=""):
    """interpreta el texto ya extraído de cada página; separado para poder probarlo"""
    informe = Informe(ruta=ruta, archivo=archivo or os.path.basename(ruta))
    informe.paginas_pdf = len(paginas)
    informe.texto = "\n".join(paginas)

    plano = _plano(informe.texto)
    informe.es_turnitin = any(m in plano for m in MARCAS)
    if "portada de integridad" in plano or "integrity overview" in plano or "similitud general" in plano:
        informe.formato = "similitud"
    elif "informe de originalidad" in plano or "originality report" in plano or "indice de similitud" in plano:
        informe.formato = "originalidad"
    elif "recibo digital" in plano or "digital receipt" in plano:
        informe.formato = "recibo"
    if not informe.es_turnitin:
        return informe

    valores = _asignar(_fusionar_porcentajes("\n".join(_region_resumen(paginas)).split("\n")))
    informe.similitud = valores.get("similitud")
    informe.ia = valores.get("ia")
    informe.internet = valores.get("internet")
    informe.publicaciones = valores.get("publicaciones")
    informe.estudiantes = valores.get("estudiantes")
    informe.fuentes = _fuentes(paginas, informe.similitud)
    informe.filtros = _filtros(plano)

    texto = informe.texto
    lineas = texto.split("\n")
    informe.entrega = _buscar(texto,
                              r"identificador de la entrega[:\s]*([^\s\n]+)",
                              r"submission id[:\s]*([^\s\n]+)",
                              r"(trn:oid:::[\d:]+)")
    informe.palabras = _entero(_buscar(texto,
                                       r"recuento de palabras[:\s]*([\d.,  ]{1,15})",
                                       r"total de palabras[:\s]*([\d.,  ]{1,15})",
                                       r"word count[:\s]*([\d.,  ]{1,15})"))
    informe.caracteres = _entero(_buscar(texto,
                                         r"recuento de caracteres[:\s]*([\d.,  ]{1,15})",
                                         r"total de caracteres[:\s]*([\d.,  ]{1,15})",
                                         r"character count[:\s]*([\d.,  ]{1,15})"))
    informe.paginas = _entero(_buscar(texto,
                                      r"total p[áa]ginas[:\s]*(\d{1,4})",
                                      r"p[áa]gina\s+\d+\s+de\s+(\d{1,4})",
                                      r"page\s+\d+\s+of\s+(\d{1,4})"))
    informe.fecha = _buscar(texto,
                            r"fecha de entrega[:\s]*([^\n]{6,60})",
                            r"fecha de env[íi]o[:\s]*([^\n]{6,60})",
                            r"submission date[:\s]*([^\n]{6,60})")
    informe.titulo = (_buscar(texto,
                              r"t[íi]tulo de la entrega[:\s]*([^\n]{3,200})",
                              r"submission title[:\s]*([^\n]{3,200})")
                      or _valor_tras(lineas, "nombre del documento", "titulo del documento",
                                     "document title"))
    informe.autor = (_buscar(texto,
                             r"\bautor[:\s]*([^\n]{3,120})",
                             r"\bauthor[:\s]*([^\n]{3,120})")
                     or _valor_tras(lineas, "nombre del estudiante", "student name"))
    informe.archivo_entregado = (_buscar(texto,
                                         r"nombre del archivo[:\s]*([^\n]{3,200})",
                                         r"file name[:\s]*([^\n]{3,200})")
                                 or _valor_tras(lineas, "nombre del documento", "document title"))
    return informe


# --------------------------------------------------------------------------- revisión

# Valores de partida. CONFIRMAR contra el reglamento de la FINESI: el tope de similitud
# y el de IA son decisiones de la Dirección de Investigación, no del programa.
PREDETERMINADO = {
    "max_pct": 25,                  # índice de similitud máximo admitido (null = no se revisa)
    "margen_alerta_pct": 5,         # avisa cuando queda a menos de esto del tope
    "max_fuente_pct": 10,           # lo máximo que puede aportar una sola fuente
    "max_ia_pct": None,             # texto detectado como generado por IA (null = no se revisa)
    "filtros_exigidos": [],         # "bibliografia", "citas"
    "max_antiguedad_dias": None,    # antigüedad máxima del informe (null = no se revisa)
    "tolerancia_palabras_pct": 15,  # diferencia admitida entre el .docx y lo que midió Turnitin
    "exigir_informe": False,        # exigir el PDF de Turnitin junto a cada .docx
}


def opciones_similitud(reglas):
    """los valores de 'similitud:' del archivo de reglas, completados con los de partida"""
    cfg = dict(PREDETERMINADO)
    propias = (reglas or {}).get("similitud")
    if isinstance(propias, dict):
        cfg.update({k: v for k, v in propias.items() if k in PREDETERMINADO})
    filtros = cfg.get("filtros_exigidos")
    cfg["filtros_exigidos"] = ([_plano(str(f)).strip() for f in filtros]
                               if isinstance(filtros, list) else [])
    return cfg


def pct(valor):
    if valor is None:
        return "—"
    return f"{valor:.0f}%" if float(valor).is_integer() else f"{valor:.1f}%"


def resumen(informe):
    """una línea con lo que declara el informe, para el registro y la hoja de revisión"""
    if informe is None:
        return ""
    if informe.error or not informe.es_turnitin:
        return f"{informe.archivo}: no se pudo leer el informe de similitud"
    partes = [f"similitud {pct(informe.similitud)}"]
    desglose = [f"{nombre} {pct(v)}" for nombre, v in
                (("internet", informe.internet), ("publicaciones", informe.publicaciones),
                 ("trabajos de estudiantes", informe.estudiantes)) if v is not None]
    if desglose:
        partes.append("(" + ", ".join(desglose) + ")")
    if informe.ia is not None:
        partes.append(f"· IA {pct(informe.ia)}")
    if informe.palabras:
        partes.append(f"· {informe.palabras} palabras")
    return " ".join(partes)


NOMBRE_FILTRO = {"bibliografia": "la bibliografía", "citas": "el texto citado"}


def _detalle_fuentes(informe):
    partes = []
    desglose = [f"{nombre} {pct(v)}" for nombre, v in
                (("internet", informe.internet), ("publicaciones", informe.publicaciones),
                 ("trabajos de estudiantes", informe.estudiantes)) if v is not None]
    if desglose:
        partes.append("Desglose: " + ", ".join(desglose) + ".")
    if informe.fuentes:
        partes.append("Fuentes que más coinciden: "
                      + "; ".join(f"{n} ({pct(v)})" for n, v in informe.fuentes[:3]) + ".")
    fuera = [NOMBRE_FILTRO[k] for k in ("bibliografia", "citas") if informe.filtros.get(k)]
    if fuera:
        partes.append("El informe midió excluyendo " + " y ".join(fuera) + ".")
    return " ".join(partes)


def revisar_similitud(informe, reglas, obs, palabras_docx=None, nombre_docx=""):
    """agrega a 'obs' lo que haya que observar del informe de Turnitin"""
    cfg = opciones_similitud(reglas)
    if informe is None:
        if cfg["exigir_informe"]:
            obs.add("Similitud", "Error", "Documento",
                    "Falta el informe de similitud de Turnitin",
                    "Adjuntar el PDF que devuelve Turnitin, con el mismo nombre del documento y "
                    "en la misma carpeta que el .docx.")
        return
    ubicacion = f"Informe de Turnitin: '{corto(informe.archivo, 50)}'"

    if informe.error:
        obs.add("Similitud", "Revisar", ubicacion, "No se pudo leer el informe de similitud",
                f"{informe.error}. Verificar el índice de similitud a mano.")
        return
    if not informe.es_turnitin:
        obs.add("Similitud", "Revisar", ubicacion, "El PDF no parece un informe de Turnitin",
                "Solo se lee el PDF que devuelve Turnitin (informe de similitud o de originalidad). "
                "Si es el documento de tesis exportado a PDF, pedir el .docx: el formato solo se "
                "puede revisar sobre el original de Word.")
        return

    tope = cfg["max_pct"]
    detalle = _detalle_fuentes(informe)
    if informe.similitud is None:
        obs.add("Similitud", "Revisar", ubicacion,
                "No se pudo leer el índice de similitud del informe",
                "Verificarlo a mano en el PDF. Puede ser una versión del informe que el programa "
                'todavía no reconoce; para ver qué leyó, ejecutar: python -m core.turnitin "archivo.pdf"')
    elif tope is not None and informe.similitud > tope:
        obs.add("Similitud", "Error", ubicacion,
                f"El índice de similitud es {pct(informe.similitud)}, el máximo permitido es {pct(tope)}",
                detalle)
    elif tope is not None and cfg["margen_alerta_pct"] and informe.similitud >= tope - cfg["margen_alerta_pct"]:
        obs.add("Similitud", "Advertencia", ubicacion,
                f"El índice de similitud es {pct(informe.similitud)}, al límite del máximo de {pct(tope)}",
                detalle)

    if cfg["max_fuente_pct"] is not None:
        excedidas = [(n, v) for n, v in informe.fuentes if v > cfg["max_fuente_pct"]]
        for nombre, valor in excedidas[:3]:
            obs.add("Similitud", "Advertencia", ubicacion,
                    f"Una sola fuente aporta {pct(valor)} de coincidencia, máximo "
                    f"{pct(cfg['max_fuente_pct'])}",
                    f"Fuente: {nombre}. Suele ser una cita larga sin comillas o un párrafo copiado; "
                    "verificar en el informe si está citada y parafraseada.")

    if cfg["max_ia_pct"] is not None:
        if informe.ia is None:
            obs.add("Similitud", "Revisar", ubicacion,
                    "El informe no incluye el porcentaje de texto generado por IA",
                    "Pedir el informe con la detección de IA activada, o verificarlo en Turnitin.")
        elif informe.ia > cfg["max_ia_pct"]:
            obs.add("Similitud", "Advertencia", ubicacion,
                    f"Turnitin marca {pct(informe.ia)} del texto como generado por IA, máximo "
                    f"{pct(cfg['max_ia_pct'])}",
                    "La detección de IA es referencial y se equivoca; contrastarla con el tesista "
                    "antes de observarla formalmente.")

    for filtro in cfg["filtros_exigidos"]:
        if filtro not in NOMBRE_FILTRO:
            continue
        estado = informe.filtros.get(filtro)
        if estado is False:
            obs.add("Similitud", "Advertencia", ubicacion,
                    f"El informe no excluyó {NOMBRE_FILTRO[filtro]} al medir la similitud",
                    "Volver a generar el informe en Turnitin con ese filtro activado: el porcentaje "
                    "que se compara con el reglamento se mide con la exclusión puesta.")
        elif estado is None:
            obs.add("Similitud", "Revisar", ubicacion,
                    f"No se pudo determinar si el informe excluyó {NOMBRE_FILTRO[filtro]}",
                    "Verificar en el PDF con qué filtros se generó el informe.")

    if palabras_docx and informe.palabras and cfg["tolerancia_palabras_pct"] is not None:
        diferencia = abs(informe.palabras - palabras_docx) / max(palabras_docx, 1) * 100
        if diferencia > cfg["tolerancia_palabras_pct"]:
            obs.add("Similitud", "Advertencia", ubicacion,
                    "El documento que se pasó por Turnitin no coincide con el que se está revisando",
                    f"El informe declara {informe.palabras} palabras y {nombre_docx or 'el .docx'} "
                    f"tiene {palabras_docx} ({diferencia:.0f}% de diferencia). Puede ser otra versión "
                    "del documento, o el informe de otro trabajo; confirmar antes de dar por válida "
                    "la similitud.")

    if cfg["max_antiguedad_dias"] and informe.fecha:
        fecha = fecha_informe(informe.fecha)
        if fecha:
            dias = (datetime.now() - fecha).days
            if dias > cfg["max_antiguedad_dias"]:
                obs.add("Similitud", "Advertencia", ubicacion,
                        f"El informe de Turnitin tiene {dias} días, máximo "
                        f"{cfg['max_antiguedad_dias']}",
                        f"Fecha de entrega declarada: {informe.fecha}. Pedir un informe del "
                        "documento en su versión actual.")


# --------------------------------------------------------------------------- emparejar con el .docx

RUIDO = ("turnitin", "informe", "reporte", "similitud", "originalidad", "integridad", "recibo",
         "digital", "report", "similarity", "originality", "receipt", "final", "copia")


def nombre_comparable(ruta):
    """nombre del archivo sin extensión ni las palabras con que se rotula el informe"""
    base = _plano(os.path.splitext(os.path.basename(ruta))[0])
    base = re.sub(r"[^a-z0-9]+", " ", base)
    return " ".join(p for p in base.split() if p not in RUIDO and len(p) > 1).strip()


# --------------------------------------------------------------------------- diagnóstico

def main():
    """python -m core.turnitin informe.pdf  ->  muestra lo que el programa leyó del PDF"""
    import argparse
    ap = argparse.ArgumentParser(
        description="Muestra lo que el revisor lee de un informe de Turnitin en PDF")
    ap.add_argument("pdf", nargs="+", help="uno o varios informes de Turnitin en PDF")
    ap.add_argument("--texto", action="store_true", help="volcar además el texto extraído del PDF")
    args = ap.parse_args()
    for ruta in args.pdf:
        informe = analizar(ruta)
        print(f"\n=== {informe.archivo}")
        if informe.error:
            print(f"  error: {informe.error}")
        print(f"  ¿es de Turnitin?: {'sí' if informe.es_turnitin else 'NO'}   "
              f"formato: {informe.formato or '—'}   páginas del PDF: {informe.paginas_pdf}")
        print(f"  similitud: {pct(informe.similitud)}   IA: {pct(informe.ia)}   "
              f"internet: {pct(informe.internet)}   publicaciones: {pct(informe.publicaciones)}   "
              f"estudiantes: {pct(informe.estudiantes)}")
        print(f"  entrega: {informe.entrega or '—'}   fecha: {informe.fecha or '—'}")
        print(f"  autor: {informe.autor or '—'}")
        print(f"  documento: {informe.archivo_entregado or informe.titulo or '—'}")
        print(f"  palabras: {informe.palabras or '—'}   caracteres: {informe.caracteres or '—'}   "
              f"páginas del documento: {informe.paginas or '—'}")
        print(f"  filtros: {informe.filtros}")
        for nombre, valor in informe.fuentes[:10]:
            print(f"    {pct(valor):>6}  {nombre}")
        if args.texto:
            print("-" * 70)
            print(informe.texto)


if __name__ == "__main__":
    main()
