"""
Orquestador de la revisión: carga las reglas, recorre el documento una sola vez y
va pasando los bloques a cada módulo de revisión, ubica cada observación en su
página y línea, y delega la escritura de los reportes a la capa de reportes.
"""
import argparse
import os
import re
import sys
from dataclasses import dataclass, field, replace
from datetime import datetime

from typing import Any, Callable, Dict, List, Optional, Tuple

import yaml
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
from rapidfuzz import fuzz

from reportes.reporte_excel import escribir_reporte
from reportes.reporte_resumen import escribir_resumen, redactar
from reportes.reporte_word import escribir_word
from utilidades import localizador, reglas_revisor

from .citas import extraer_citas, revisar_citas
from .config import (DATOS, carpeta_reglas, preparar_datos, ruta_recurso,
                     tipos_disponibles)
from .estructura import (cargar_textos_tablas, detectar_secciones, marcas_guia,
                         revisar_estructura, revisar_marcas)
from .extension import contar_palabras, revisar_extension, texto_de_seccion
from .formato import revisar_formato
from .ortografia import revisar_ortografia
from .turnitin import PREDETERMINADO as PREDETERMINADO_SIMILITUD
from .turnitin import (analizar as analizar_informe, nombre_comparable,
                       resumen as resumen_informe, revisar_similitud)
from .utils import SEVERIDADES, Obs, bloques_en_orden, desenvolver_controles, corto, limpiar_titulo, norm


def ubicar_lineas(obs, filas_orto, mapa, bloques, rangos):
    """agrega a cada observacion la pagina y linea donde esta el texto"""
    for it in obs.items:
        if not it.get("propia"):        # las reglas del revisor ya traen su línea
            it["linea"], it["nlinea"] = "", None
    for f in filas_orto:
        f["linea"], f["nlinea"] = "", None
    if not mapa:
        return
    def titulo_de(nombre):
        if nombre not in rangos:
            return None
        ini, fin = rangos[nombre]
        texto = bloques[ini].text
        if len(re.sub(r"\W", "", texto)) < 8:   # titulos muy cortos: se agrega el inicio del contenido
            sigue = next((b.text for b in bloques[ini + 1:fin] if isinstance(b, Paragraph) and b.text.strip()), "")
            texto = (texto + " " + sigue).strip()
        return texto

    for it in obs.items:
        if it.get("propia"):
            continue
        seccion = (it["ubicacion"] or "").split(":")[0].replace(" (tabla)", "").strip()
        ancla = titulo_de(seccion)
        frag = localizador.fragmento(it["ubicacion"])
        if not frag and seccion in rangos:
            frag, ancla = titulo_de(seccion), None
        if frag:
            it["linea"] = mapa.buscar(frag, ancla)
            it["nlinea"] = mapa.numero(frag, ancla)
    for f in filas_orto:
        ancla = titulo_de(f.get("seccion", ""))
        f["linea"] = mapa.buscar(f.get("frag", ""), ancla) or mapa.buscar(f.get("palabra", ""), ancla)
        f["nlinea"] = mapa.numero(f.get("frag", ""), ancla) or mapa.numero(f.get("palabra", ""), ancla)



def leer_registro(ruta_registro):
    """lee registro.csv: archivo;tesista;proyecto;asesor;fecha_subida"""
    if not ruta_registro or not os.path.exists(ruta_registro):
        return {}
    import csv
    reg = {}
    contenido = None
    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin1"):
        try:
            with open(ruta_registro, encoding=enc, newline="") as fh:
                contenido = fh.read()
            break
        except UnicodeDecodeError:
            continue
            
    if contenido is None:
        return {}
        
    import io
    fh = io.StringIO(contenido)
    muestra = fh.read(2048)
    fh.seek(0)
    try:
        dial = csv.Sniffer().sniff(muestra, delimiters=";,\t")
    except Exception:
        dial = csv.excel
        dial.delimiter = ";"
    for fila in csv.DictReader(fh, dialect=dial):
        clave = (fila.get("archivo") or "").strip().lower()
        if clave:
            reg[clave] = {k: (v or "").strip() for k, v in fila.items() if k}
            reg[os.path.splitext(clave)[0]] = reg[clave]
    return reg


def fecha_local(marca):
    """'2026-09-27T01:00:00Z' (UTC, como la guarda Word) -> '26/09/2026' en hora local"""
    from datetime import timezone
    marca = marca.strip()
    if len(marca) == 10:                       # solo fecha, sin hora: no hay nada que convertir
        return datetime.strptime(marca, "%Y-%m-%d").strftime("%d/%m/%Y")
    momento = datetime.fromisoformat(marca.replace("Z", "+00:00"))
    if momento.tzinfo is None:
        momento = momento.replace(tzinfo=timezone.utc)
    return momento.astimezone().strftime("%d/%m/%Y")


def fecha_archivo(ruta):
    """
    Fecha en que el tesista entregó: la del último guardado que registra el docx, si
    no la del archivo. Word la guarda en UTC ('2026-09-27T01:00:00Z'): tomar solo la
    fecha adelantaba un día todo lo guardado en Perú después de las 19:00, así que se
    pasa a la hora local del equipo.
    """
    try:
        import zipfile
        with zipfile.ZipFile(ruta) as z:
            xml = z.read("docProps/core.xml").decode("utf8", "ignore")
        m = re.search(r"<dcterms:modified[^>]*>([^<]+)<", xml)
        if m:
            return fecha_local(m.group(1))
    except Exception:
        pass
    try:
        return datetime.fromtimestamp(os.path.getmtime(ruta)).strftime("%d/%m/%Y")
    except Exception:
        return ""


def datos_expediente(ruta: str, bloques: List[Any], rangos: Dict[str, Tuple[int, int]], reglas: Dict[str, Any], registro: Dict[str, Any], args: Any) -> Dict[str, str]:
    base = os.path.basename(ruta)
    fila = registro.get(base.lower()) or registro.get(os.path.splitext(base)[0].lower()) or {}
    titulo = texto_de_seccion(bloques, rangos, (reglas.get("extension") or {}).get("seccion_titulo") or "Título")
    titulo = re.sub(r"\s+", " ", titulo).strip()
    return dict(
        tipo=reglas.get("tipo", ""),
        codigo=getattr(args, "codigo", "") or fila.get("codigo", ""),
        n_revision=getattr(args, "n_revision", "") or fila.get("n_revision", ""),
        etapa=fila.get("etapa", ""),
        archivo=base,
        tesista=getattr(args, "tesista", "") or fila.get("tesista", ""),
        proyecto=getattr(args, "proyecto", "") or fila.get("proyecto", "") or (corto(titulo, 300) if titulo else ""),
        asesor=getattr(args, "asesor", "") or fila.get("asesor", ""),
        fecha_subida=getattr(args, "fecha", "") or fila.get("fecha_subida", "") or fecha_archivo(ruta),
        fecha_revision=datetime.now().strftime("%d/%m/%Y"),
        revisor=getattr(args, "revisor", "") or "",
    )


def parece_otro_tipo(bloques: List[Any], reglas: Dict[str, Any]) -> List[str]:
    """devuelve los titulos encontrados que pertenecen al otro tipo de documento"""
    ajenos = [norm(x) for x in reglas.get("secciones_de_otro_tipo", [])]
    hallados = set()
    for b in bloques:
        if isinstance(b, Paragraph) and 0 < len(b.text.strip()) < 80:
            t = limpiar_titulo(b.text)
            for a in ajenos:
                if t and fuzz.ratio(t, a) >= 90:
                    hallados.add(b.text.strip())
    return sorted(hallados)


def _protegido(obs: Obs, categoria: str, descripcion: str, fn: Callable, por_defecto: Any = None) -> Any:
    """
    Ejecuta un chequeo sin que su fallo se lleve el documento entero. Antes, una
    excepción en cualquier módulo descartaba todo lo ya revisado y el revisor solo
    veía 'no se pudo revisar'; ahora recibe su hoja con lo que sí se pudo verificar
    y una línea explícita sobre lo que quedó sin revisar.
    """
    try:
        return fn()
    except Exception as ex:
        obs.add(categoria, "Revisar", "Documento",
                f"No se pudo completar la revisión de {descripcion}",
                f"Error interno del programa ({type(ex).__name__}: {ex}). "
                f"El resto del documento sí se revisó. Revisar {descripcion} a mano "
                f"y avisar al responsable del sistema.")
        return por_defecto


def _base_reporte(prefijo: str, base: str, usados: Optional[set] = None) -> str:
    """
    Nombre del reporte, sin extensión. Reservarlo en 'usados' evita que dos documentos
    que se llaman igual, venidos de carpetas distintas, escriban los dos sobre la misma
    hoja: el segundo sale como '... (2)'. Volver a revisar el mismo documento sí
    sobrescribe su hoja anterior, que es lo que el revisor espera.
    """
    nombre = f"{prefijo}{base}"
    if usados is None:
        return nombre
    candidato, n = nombre, 1
    while candidato.lower() in usados:
        n += 1
        candidato = f"{nombre} ({n})"
    usados.add(candidato.lower())
    return candidato


def revisar(ruta: str, reglas: Dict[str, Any], carpeta_salida: str, registro: Dict[str, Any], args: Any, forzar: bool = False, aviso: Callable = print, informe: Any = None, usados: Optional[set] = None) -> Optional[str]:
    """
    Revisa un .docx. 'informe' es el PDF que devolvió Turnitin para ese mismo
    documento, ya leído: sus observaciones van en la misma hoja, porque el revisor
    entrega una sola por expediente.
    """
    doc = Document(ruta)
    desenvolver_controles(doc)   # citas y bibliografía de Mendeley/Zotero
    bloques = list(bloques_en_orden(doc))
    ajenos = parece_otro_tipo(bloques, reglas)
    if len(ajenos) >= 2 and not forzar:
        aviso(f"{os.path.basename(ruta)}: NO REVISADO. Parece que no es un {reglas['tipo']}: tiene los títulos "
              f"{', '.join(repr(x) for x in ajenos[:4])}. Revisa si elegiste bien --tipo (o usa --forzar).")
        return None
    obs = Obs(reglas.get("severidades"))
    filas_orto = []
    if getattr(args, "motor", ""):
        reglas = dict(reglas)
        reglas["ortografia"] = dict(reglas.get("ortografia") or {}, motor=args.motor)
    mapa = localizador.construir(ruta)

    plantilla = ruta_recurso(reglas.get("plantilla", ""))
    _protegido(obs, "Formato", "la plantilla oficial", lambda: cargar_textos_tablas(plantilla))

    # La detección de títulos se calcula una sola vez: la usan estructura y formato.
    deteccion = _protegido(obs, "Estructura", "los títulos del esquema",
                           lambda: detectar_secciones(bloques, reglas), ([], [], []))
    rangos = _protegido(obs, "Estructura", "la estructura",
                        lambda: revisar_estructura(bloques, reglas, obs, deteccion), {}) or {}
    _protegido(obs, "Formato", "el texto guía de la plantilla",
               lambda: revisar_marcas(bloques, rangos, marcas_guia(plantilla), obs))
    _protegido(obs, "Formato", "el formato",
               lambda: revisar_formato(doc, bloques, rangos, reglas, obs, deteccion))
    _protegido(obs, "Extensión", "la extensión",
               lambda: revisar_extension(ruta, bloques, rangos, reglas, obs, mapa))
    _protegido(obs, "Citas", "las citas y referencias",
               lambda: revisar_citas(bloques, rangos, reglas, obs))
    _protegido(obs, "Similitud", "el informe de similitud de Turnitin",
               lambda: revisar_similitud(informe, reglas, obs,
                                         palabras_docx=contar_palabras(bloques),
                                         nombre_docx=os.path.basename(ruta)))
    motor = _protegido(obs, "Ortografía", "la ortografía",
                       lambda: revisar_ortografia(bloques, rangos, reglas, obs, filas_orto),
                       "no disponible")
    _protegido(obs, "Revisor", "tus reglas propias",
               lambda: reglas_revisor.aplicar(
                   reglas_revisor.crear_si_falta(os.path.join(DATOS, "mis_reglas.yaml")),
                   bloques, rangos, mapa, obs, lambda t: len(extraer_citas(t)), Paragraph, Table))

    ubicar_lineas(obs, filas_orto, mapa, bloques, rangos)
    datos = datos_expediente(ruta, bloques, rangos, reglas, registro, args)
    datos["similitud"] = resumen_informe(informe)
    base = os.path.splitext(os.path.basename(ruta))[0]
    orden_sev = {"Error": 0, "Advertencia": 1, "Revisar": 2}
    orden_cat = {"Estructura": 0, "Formato": 1, "Extensión": 2, "Citas": 3, "Similitud": 4,
                 "Ortografía": 5, "Gramática/puntuación": 6, "Revisor": 7}
    items = sorted(obs.items, key=lambda x: (orden_cat.get(x["categoria"], 9), orden_sev.get(x["severidad"], 9)))
    resumen = {c: sum(1 for x in obs.items if x["categoria"] == c) for c in orden_cat}

    formato = getattr(args, "formato", "resumen")
    stem = _base_reporte("Observaciones - ", base, usados)
    salida = os.path.join(carpeta_salida, f"{stem}.docx")
    if formato == "detallado":
        escribir_word(salida, datos, items, filas_orto, motor, resumen, mapa)
    else:
        lista = redactar(items, filas_orto, getattr(args, "extras", ()) or (), reglas.get("tipo", "proyecto"))
        escribir_resumen(salida, datos, lista)
        if formato == "ambos":
            detalle = os.path.join(carpeta_salida, f"{stem} (detalle).docx")
            escribir_word(detalle, datos, items, filas_orto, motor, resumen, mapa)
    if args.excel:
        xls = os.path.join(carpeta_salida, f"{stem}.xlsx")
        escribir_reporte(xls, f"{os.path.basename(ruta)} ({reglas['tipo']})", obs, filas_orto, motor)

    n_err = sum(1 for x in obs.items if x["severidad"] == "Error")
    faltan = [k for k in ("tesista", "asesor") if not datos.get(k)]
    nota = f"  [completar en el Word: {', '.join(faltan)}]" if faltan else ""
    turnitin = f", {datos['similitud']}" if datos.get("similitud") else ""
    aviso(f"{os.path.basename(ruta)}: {n_err} errores, {len(obs.items) - n_err} "
          f"advertencias/revisar{turnitin} -> {salida}{nota}")
    return salida


@dataclass
class Opciones:
    """Opciones de una corrida; las usan tanto la línea de comandos como la interfaz."""

    tipo: str = "proyecto"
    salida: str = "reportes"
    excel: bool = False
    forzar: bool = False
    registro: str = None        # None = data/registro.csv
    revisor: str = ""
    tesista: str = ""
    proyecto: str = ""
    asesor: str = ""
    fecha: str = ""
    motor: str = ""             # "" = lo que diga el archivo de reglas
    formato: str = "resumen"    # resumen | detallado | ambos
    codigo: str = ""
    n_revision: str = ""
    turnitin: str = ""          # informe de Turnitin en PDF para el documento que se revisa
    extras: List[str] = field(default_factory=list)

    def __post_init__(self):
        if not self.registro:
            self.registro = os.path.join(DATOS, "registro.csv")
        self.extras = list(self.extras)


# Claves que los módulos de revisión leen de forma directa: si falta una, el chequeo
# revienta con un KeyError que no le dice nada al revisor. Se validan al cargar.
FORMATO_OBLIGATORIO = ["fuente", "tamano_pt", "interlineado", "alineacion_cuerpo",
                       "tolerancia_margen_cm", "tolerancia_fuera_de_norma"]
MARGENES_OBLIGATORIOS = ["izquierdo", "derecho", "superior", "inferior"]


def validar_reglas(reglas):
    """Devuelve la lista de problemas del archivo de reglas, en lenguaje del revisor."""
    fallos = []
    secciones = reglas.get("secciones")
    if not secciones:
        fallos.append("falta la lista 'secciones' (el esquema del documento)")
    elif not isinstance(secciones, list):
        fallos.append("'secciones' debe ser una lista")
    else:
        for k, sec in enumerate(secciones, 1):
            if not isinstance(sec, dict) or not sec.get("nombre"):
                fallos.append(f"la sección número {k} no tiene 'nombre'")

    f = reglas.get("formato")
    if f is not None:
        if not isinstance(f, dict):
            fallos.append("'formato' debe ser un bloque con claves, no un valor suelto")
        else:
            for clave in FORMATO_OBLIGATORIO:
                if clave not in f:
                    fallos.append(f"en 'formato' falta '{clave}'")
            papel = f.get("papel")
            if not isinstance(papel, dict) or "ancho_cm" not in papel or "alto_cm" not in papel:
                fallos.append("en 'formato' falta 'papel' con 'ancho_cm' y 'alto_cm'")
            margenes = f.get("margenes_cm")
            if not isinstance(margenes, dict):
                fallos.append("en 'formato' falta 'margenes_cm'")
            else:
                for lado in MARGENES_OBLIGATORIOS:
                    if lado not in margenes:
                        fallos.append(f"en 'formato.margenes_cm' falta '{lado}'")

    similitud = reglas.get("similitud")
    if similitud is not None:
        if not isinstance(similitud, dict):
            fallos.append("'similitud' debe ser un bloque con claves, no un valor suelto")
        else:
            for clave, valor in similitud.items():
                if clave not in PREDETERMINADO_SIMILITUD:
                    fallos.append(f"en 'similitud' sobra '{clave}' (se usan: "
                                  f"{', '.join(PREDETERMINADO_SIMILITUD)})")
                elif clave == "filtros_exigidos":
                    if not isinstance(valor, list) or any(v not in ("bibliografia", "citas")
                                                          for v in valor):
                        fallos.append("'similitud.filtros_exigidos' debe ser una lista con "
                                      "'bibliografia' y/o 'citas'")
                elif clave == "exigir_informe":
                    if not isinstance(valor, bool):
                        fallos.append("'similitud.exigir_informe' debe ser true o false")
                elif valor is not None and not isinstance(valor, (int, float)):
                    fallos.append(f"'similitud.{clave}' debe ser un número o null")

    ajustes = reglas.get("severidades") or {}
    if not isinstance(ajustes, dict):
        fallos.append("'severidades' debe ser un bloque 'Categoría: Severidad'")
    else:
        for cat, sev in ajustes.items():
            if sev not in SEVERIDADES:
                fallos.append(f"la severidad '{sev}' de '{cat}' no es válida "
                              f"(usar: {', '.join(SEVERIDADES)})")
    return fallos


def cargar_reglas(tipo):
    ruta = os.path.join(carpeta_reglas(), tipo + ".yaml")
    try:
        with open(ruta, encoding="utf8") as fh:
            reglas = yaml.safe_load(fh) or {}
    except FileNotFoundError:
        raise ValueError(f"No existe el archivo de reglas '{ruta}'.")
    except yaml.YAMLError as ex:
        raise ValueError(f"El archivo de reglas '{tipo}.yaml' tiene un error de formato YAML:\n{ex}")
    if not isinstance(reglas, dict):
        raise ValueError(f"El archivo de reglas '{tipo}.yaml' debería empezar con claves como 'secciones:'.")
    reglas.setdefault("tipo", tipo)

    fallos = validar_reglas(reglas)
    if fallos:
        detalle = "\n".join(f"  - {x}" for x in fallos)
        raise ValueError(
            f"El archivo de reglas '{tipo}.yaml' está incompleto:\n{detalle}\n\n"
            f"Si todavía no lo generaste, copia la plantilla oficial en "
            f"data/plantillas/{tipo}.docx y ejecuta:\n"
            f"    python scripts/generar_reglas.py data/plantillas/{tipo}.docx --tipo {tipo}")
    return reglas


def expandir(entradas):
    """
    Convierte archivos y carpetas en lo que se puede revisar: los .docx y los PDF
    que devuelve Turnitin. Devuelve (documentos, pdfs, avisos).
    """
    documentos, pdfs, avisos = [], [], []
    for entrada in entradas:
        if not os.path.exists(entrada):
            avisos.append(f"No existe: {entrada}")
        elif os.path.isdir(entrada):
            todos = [f for f in sorted(os.listdir(entrada)) if not f.startswith("~$")]
            documentos += [os.path.join(entrada, f) for f in todos if f.lower().endswith(".docx")]
            pdfs += [os.path.join(entrada, f) for f in todos if f.lower().endswith(".pdf")]
            if any(f.lower().endswith(".doc") for f in todos):
                avisos.append(f"En {os.path.basename(entrada)} hay archivos .doc que no se revisan. "
                              "Pedir el .docx (un .doc se abre en Word y se guarda como .docx).")
        elif entrada.lower().endswith(".pdf"):
            pdfs.append(entrada)
        elif not entrada.lower().endswith(".docx"):
            avisos.append(f"Se omite {os.path.basename(entrada)}: solo se revisan archivos .docx "
                          "y los informes de Turnitin en .pdf.")
        else:
            documentos.append(entrada)
    return documentos, pdfs, avisos


def leer_informes(pdfs, aviso=print):
    """lee los PDF y se queda con los que son informes de Turnitin"""
    informes = []
    for ruta in pdfs:
        informe = analizar_informe(ruta)
        if informe.error:
            aviso(f"{os.path.basename(ruta)}: no se pudo leer el PDF ({informe.error}).")
        elif not informe.es_turnitin:
            aviso(f"{os.path.basename(ruta)}: no parece un informe de Turnitin, no se revisa. "
                  "Del PDF solo se revisa el informe de similitud; el formato necesita el .docx.")
        else:
            informes.append(informe)
    return informes


def _parecido(a, b):
    """
    Cuánto se parecen dos nombres de archivo ya normalizados.

    Con fuzz.ratio solo, 'proyecto cruz apaza' y 'cruz apaza' dan 69 y no se emparejaban:
    la palabra de más castiga por diferencia de largo, y el PDF de Turnitin casi nunca
    repite el rótulo del .docx. token_set_ratio compara por palabras, que es lo que
    importa acá: sube esos casos a 90 o más y deja los de tesistas distintos en 63 o menos.

    Pero token_set_ratio con una sola palabra en común es peligroso: 'proyecto quispe
    mamani' contra 'quispe' da 100 aunque sean dos tesistas, porque un nombre contiene al
    otro. Con menos de dos palabras compartidas se vuelve a comparar letra por letra, que
    para un apellido solo es lo correcto.
    """
    if not a or not b:
        return 0
    if len(set(a.split()) & set(b.split())) >= 2:
        return max(fuzz.ratio(a, b), fuzz.token_set_ratio(a, b))
    return fuzz.ratio(a, b)


def emparejar(documentos, informes):
    """
    Decide qué informe de Turnitin corresponde a cada .docx. Se empareja por el
    nombre del archivo y por el nombre del documento que declara el propio informe;
    si el PDF se llama solo 'Informe Turnitin.pdf' y hay un único documento, va con
    ese. Devuelve ({docx: informe}, informes que quedaron sueltos).
    """
    candidatos = []
    for doc in documentos:
        clave = nombre_comparable(doc)
        for informe in informes:
            puntajes = [_parecido(clave, nombre_comparable(informe.ruta))]
            if informe.archivo_entregado:
                puntajes.append(_parecido(clave, nombre_comparable(informe.archivo_entregado)))
            candidatos.append((max(puntajes), doc, informe))
    candidatos.sort(key=lambda x: -x[0])

    parejas, usados = {}, set()
    for puntaje, doc, informe in candidatos:
        if puntaje < 75 or doc in parejas or id(informe) in usados:
            continue
        parejas[doc] = informe
        usados.add(id(informe))

    sueltos = [i for i in informes if id(i) not in usados]
    # Un informe con nombre genérico ('Informe Turnitin.pdf') y un solo documento:
    # no hay con qué confundirlo. Si no corresponde, el conteo de palabras lo delata.
    if len(documentos) == 1 and documentos[0] not in parejas and len(sueltos) == 1:
        if not nombre_comparable(sueltos[0].ruta):
            parejas[documentos[0]] = sueltos.pop()
    return parejas, sueltos


def buscar_informe_junto(ruta_docx, aviso=print):
    """busca en la carpeta del documento un informe de Turnitin que se llame igual"""
    carpeta = os.path.dirname(os.path.abspath(ruta_docx))
    clave = nombre_comparable(ruta_docx)
    if not clave or not os.path.isdir(carpeta):
        return None
    for nombre in sorted(os.listdir(carpeta)):
        if not nombre.lower().endswith(".pdf") or nombre.startswith("~$"):
            continue
        if _parecido(clave, nombre_comparable(nombre)) < 80:
            continue
        informe = analizar_informe(os.path.join(carpeta, nombre))
        if informe.es_turnitin and not informe.error:
            return informe
    return None


def revisar_informe(informe, reglas, carpeta_salida, registro, args, aviso=print, usados=None):
    """
    Genera la hoja de un informe de Turnitin que vino sin su .docx. Del PDF solo se
    revisa la similitud: el formato, la estructura y las citas se revisan sobre el
    documento de Word, porque el PDF de Turnitin no conserva ni los márgenes ni los
    estilos del original.
    """
    obs = Obs(reglas.get("severidades"))
    revisar_similitud(informe, reglas, obs)
    base = os.path.splitext(os.path.basename(informe.ruta))[0]
    fila = registro.get(os.path.basename(informe.ruta).lower()) or registro.get(base.lower()) or {}
    datos = dict(
        tipo=reglas.get("tipo", ""),
        codigo=getattr(args, "codigo", "") or fila.get("codigo", ""),
        n_revision=getattr(args, "n_revision", "") or fila.get("n_revision", ""),
        etapa=fila.get("etapa", "") or "Revisión del informe de similitud (Turnitin)",
        archivo=os.path.basename(informe.ruta),
        tesista=args.tesista or fila.get("tesista", "") or informe.autor,
        proyecto=args.proyecto or fila.get("proyecto", "") or corto(informe.titulo, 300),
        asesor=args.asesor or fila.get("asesor", ""),
        fecha_subida=args.fecha or fila.get("fecha_subida", "") or informe.fecha,
        fecha_revision=datetime.now().strftime("%d/%m/%Y"),
        revisor=args.revisor or "",
        similitud=resumen_informe(informe),
    )
    items = sorted(obs.items, key=lambda x: {"Error": 0, "Advertencia": 1}.get(x["severidad"], 2))
    stem = _base_reporte("Similitud - ", base, usados)
    salida = os.path.join(carpeta_salida, f"{stem}.docx")
    formato = getattr(args, "formato", "resumen")
    if formato == "detallado":
        escribir_word(salida, datos, items, [], "no revisada", {"Similitud": len(items)})
    else:
        lista = redactar(items, [], getattr(args, "extras", ()) or (), reglas.get("tipo", "proyecto"))
        escribir_resumen(salida, datos, lista)
        if formato == "ambos":
            escribir_word(os.path.join(carpeta_salida, f"{stem} (detalle).docx"),
                          datos, items, [], "no revisada", {"Similitud": len(items)})
    n_err = sum(1 for x in items if x["severidad"] == "Error")
    aviso(f"{os.path.basename(informe.ruta)}: informe de Turnitin, {datos['similitud']}, "
          f"{n_err} errores -> {salida}")
    return salida


def ejecutar(entradas, opciones, aviso=print, progreso=None):
    """
    Revisa todo lo que haya en 'entradas' y devuelve la lista de reportes generados.

    'aviso' recibe cada línea del registro de la revisión. 'progreso' se llama antes
    de empezar cada documento, como progreso(hechas, total, nombre): hasta que no se
    emparejan los PDF con los .docx no se sabe cuántas hojas van a salir, así que el
    total lo informa esta función y no quien la llama.
    """
    preparar_datos()
    reglas = cargar_reglas(opciones.tipo)
    documentos, pdfs, avisos = expandir(entradas)
    for a in avisos:
        aviso(a)
    elegido = getattr(opciones, "turnitin", "")
    if elegido:
        if not os.path.exists(elegido):
            raise ValueError(f"No existe el informe de Turnitin indicado: {elegido}")
        if len(documentos) != 1:
            raise ValueError("El informe indicado con --turnitin solo vale al revisar un documento.")
        pdfs = [p for p in pdfs if os.path.abspath(p) != os.path.abspath(elegido)]
    informes = leer_informes(pdfs, aviso)
    parejas, sueltos = emparejar(documentos, informes)
    if elegido:
        # El revisor nombró este PDF a mano: si no se puede leer, se le dice y no se
        # sale a buscar otro a la carpeta, que sería cambiarle el archivo en silencio.
        propios = leer_informes([elegido], aviso)
        if not propios:
            raise ValueError(f"No se pudo usar el informe de Turnitin indicado: "
                             f"{os.path.basename(elegido)}. Revisar el mensaje anterior.")
        parejas[documentos[0]] = propios[0]
    if not documentos and not sueltos:
        raise ValueError("No hay archivos .docx ni informes de Turnitin en .pdf para revisar.")
    opciones_sueltos = opciones
    if len(documentos) + len(sueltos) > 1 and any([opciones.tesista, opciones.asesor, opciones.fecha]):
        if len(documentos) != 1:
            raise ValueError("Los datos de tesista, asesor y fecha solo valen para un archivo. "
                             "Para varios, usa registro.csv.")
        # Un solo .docx y algún PDF que no se pudo emparejar: los datos escritos son
        # los de ese documento. Abortar dejaba al revisor sin ninguna hoja; en cambio
        # los informes sueltos toman sus datos de registro.csv o del propio PDF.
        opciones_sueltos = replace(opciones, tesista="", asesor="", fecha="", codigo="", proyecto="")
        for informe in sueltos:
            aviso(f"{os.path.basename(informe.ruta)}: no corresponde a "
                  f"{os.path.basename(documentos[0])}; los datos del expediente no se le aplican.")
    registro = leer_registro(opciones.registro)
    os.makedirs(opciones.salida, exist_ok=True)

    total = len(documentos) + len(sueltos)
    hechos, avance, usados = [], 0, set()
    for f in documentos:
        if progreso:
            progreso(avance, total, os.path.basename(f))
        informe = parejas.get(f) if elegido else (parejas.get(f) or buscar_informe_junto(f, aviso))
        if informe is not None:
            aviso(f"{os.path.basename(f)}: se usa el informe de Turnitin "
                  f"'{os.path.basename(informe.ruta)}'.")
        try:
            r = revisar(f, reglas, opciones.salida, registro, opciones, opciones.forzar, aviso,
                        informe, usados)
            if r:
                hechos.append(r)
        except Exception as ex:
            aviso(f"{os.path.basename(f)}: no se pudo revisar ({type(ex).__name__}: {ex})")
        avance += 1
    for informe in sueltos:
        if progreso:
            progreso(avance, total, os.path.basename(informe.ruta))
        try:
            hechos.append(revisar_informe(informe, reglas, opciones.salida, registro, opciones_sueltos,
                                          aviso, usados))
        except Exception as ex:
            aviso(f"{os.path.basename(informe.ruta)}: no se pudo revisar "
                  f"({type(ex).__name__}: {ex})")
        avance += 1
    if progreso:
        progreso(avance, total, "")
    return hechos


def main():
    preparar_datos()
    tipos = [t for t, _ in tipos_disponibles()]
    ap = argparse.ArgumentParser(description="Revisor de formato de tesis FINESI")
    ap.add_argument("entrada", nargs="+",
                    help="uno o varios .docx, los informes de Turnitin en .pdf, o una carpeta")
    ap.add_argument("--tipo", required=True, choices=tipos, help="tipo de documento: " + ", ".join(tipos))
    ap.add_argument("--salida", default=os.path.join(DATOS, "reportes"), help="carpeta donde se guardan los reportes")
    ap.add_argument("--forzar", action="store_true", help="revisar aunque parezca de otro tipo")
    ap.add_argument("--formato", default="resumen", choices=["resumen", "detallado", "ambos"],
                    help="resumen: hoja corta estilo PGI (por defecto); detallado: tabla completa")
    ap.add_argument("--codigo", default="", help="código del expediente, ej. P26-183522A")
    ap.add_argument("--n-revision", dest="n_revision", default="", help="número de revisión, ej. 1")
    ap.add_argument("--excel", action="store_true", help="generar también el Excel (solo con --formato detallado)")
    ap.add_argument("--registro", default=os.path.join(DATOS, "registro.csv"),
                    help="CSV con archivo;tesista;proyecto;asesor;fecha_subida")
    ap.add_argument("--revisor", default="", help="nombre de quien revisa (va en la firma)")
    ap.add_argument("--turnitin", default="",
                    help="informe de Turnitin en PDF del documento que se revisa; si el PDF está "
                         "en la misma carpeta y se llama igual que el .docx, se toma solo")
    g = ap.add_argument_group("datos del expediente (solo si revisas UN archivo; si no, usa registro.csv)")
    g.add_argument("--tesista", default="")
    g.add_argument("--proyecto", default="", help="título del proyecto; por defecto se toma del documento")
    g.add_argument("--asesor", default="")
    g.add_argument("--fecha", default="", help="fecha de presentación, ej. 05/09/2026")
    g.add_argument("--extra", dest="extras", action="append", default=[],
                   help="observación adicional; se puede repetir")
    a = ap.parse_args()

    try:
        ejecutar(a.entrada, a, print)   # valida las reglas y revisa
    except ValueError as ex:
        sys.exit(str(ex))


if __name__ == "__main__":
    main()
