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


def fecha_archivo(ruta):
    """fecha en que el tesista entregó: la de creación guardada en el docx, si no la del archivo"""
    try:
        import zipfile
        with zipfile.ZipFile(ruta) as z:
            xml = z.read("docProps/core.xml").decode("utf8", "ignore")
        m = re.search(r"<dcterms:modified[^>]*>(\d{4}-\d{2}-\d{2})", xml)
        if m:
            return datetime.strptime(m.group(1), "%Y-%m-%d").strftime("%d/%m/%Y")
    except Exception:
        pass
    try:
        return datetime.fromtimestamp(os.path.getmtime(ruta)).strftime("%d/%m/%Y")
    except Exception:
        return ""


def datos_expediente(ruta, bloques, rangos, reglas, registro, args):
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
        tesista=args.tesista or fila.get("tesista", ""),
        proyecto=args.proyecto or fila.get("proyecto", "") or (corto(titulo, 300) if titulo else ""),
        asesor=args.asesor or fila.get("asesor", ""),
        fecha_subida=args.fecha or fila.get("fecha_subida", "") or fecha_archivo(ruta),
        fecha_revision=datetime.now().strftime("%d/%m/%Y"),
        revisor=args.revisor or "",
    )


COLORES = {"Error": "F8CBAD", "Advertencia": "FFE699", "Revisar": "BDD7EE"}


def escribir_reporte(ruta_salida, nombre_doc, obs, filas_orto, motor, paginas_info):
    wb = Workbook()
    ws = wb.active
    ws.title = "Observaciones"
    ws.append([f"Revisión automática: {nombre_doc}"])
    ws["A1"].font = Font(bold=True, size=12)
    ws.append(["Borrador generado por el revisor. Validar cada observación antes de enviarla al tesista."])
    ws.append([])
    enc = ["N°", "Categoría", "Severidad", "Página y línea", "Ubicación", "Observación", "Detalle / corrección"]
    ws.append(enc)
    for c in ws[4]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E78")
    orden_sev = {"Error": 0, "Advertencia": 1, "Revisar": 2}
    orden_cat = {"Estructura": 0, "Formato": 1, "Extensión": 2, "Citas": 3, "Ortografía": 4}
    items = sorted(obs.items, key=lambda x: (orden_cat.get(x["categoria"], 9), orden_sev.get(x["severidad"], 9)))
    for k, it in enumerate(items, 1):
        ws.append([k, it["categoria"], it["severidad"], it.get("linea", ""), it["ubicacion"], it["observacion"], it["detalle"]])
        ws.cell(row=ws.max_row, column=3).fill = PatternFill("solid", fgColor=COLORES.get(it["severidad"], "FFFFFF"))
    for col, ancho in zip("ABCDEFG", [5, 13, 12, 18, 40, 55, 60]):
        ws.column_dimensions[col].width = ancho
    for fila in ws.iter_rows(min_row=5):
        for c in fila:
            c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "A5"

    r = wb.create_sheet("Resumen")
    r.append(["Categoría", "Errores", "Advertencias", "Revisar"])
    for c in r[1]:
        c.font = Font(bold=True)
    for cat in orden_cat:
        sub = [x for x in obs.items if x["categoria"] == cat]
        r.append([cat] + [sum(1 for x in sub if x["severidad"] == s) for s in ("Error", "Advertencia", "Revisar")])
    r.append([])
    r.append(["Motor de ortografía", motor])
    r.column_dimensions["A"].width = 22
    r.column_dimensions["B"].width = 14

    o = wb.create_sheet("Ortografía")
    o.append(["Página y línea", "Ubicación", "Palabra", "Problema", "Sugerencias", "Contexto"])
    for c in o[1]:
        c.font = Font(bold=True)
    for f in filas_orto:
        o.append([f.get("linea", ""), f["parrafo"], f["palabra"], f["problema"], f["sugerencias"], f["contexto"]])
    for col, ancho in zip("ABCDEF", [18, 40, 20, 45, 30, 70]):
        o.column_dimensions[col].width = ancho
    wb.save(ruta_salida)



def parece_otro_tipo(bloques, reglas):
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


def revisar(ruta, reglas, carpeta_salida, registro, args, forzar=False, aviso=print):
    doc = Document(ruta)
    bloques = list(bloques_en_orden(doc))
    ajenos = parece_otro_tipo(bloques, reglas)
    if len(ajenos) >= 2 and not forzar:
        aviso(f"{os.path.basename(ruta)}: NO REVISADO. Parece que no es un {reglas['tipo']}: tiene los títulos "
              f"{', '.join(repr(x) for x in ajenos[:4])}. Revisa si elegiste bien --tipo (o usa --forzar).")
        return None
    obs = Obs()
    filas_orto = []
    if getattr(args, "motor", ""):
        reglas = dict(reglas)
        reglas["ortografia"] = dict(reglas.get("ortografia") or {}, motor=args.motor)
    mapa = localizador.construir(ruta)

    bloques = list(bloques_en_orden(doc))
    rangos = {}
    cargar_textos_tablas(ruta_recurso(reglas.get("plantilla", "")))
    rangos = revisar_estructura(bloques, reglas, obs)
    revisar_marcas(bloques, rangos, marcas_guia(ruta_recurso(reglas.get("plantilla", ""))), obs)
    revisar_formato(doc, bloques, rangos, reglas, obs)
    revisar_extension(ruta, bloques, rangos, reglas, obs, mapa)
    revisar_citas(bloques, rangos, reglas, obs)
    motor = revisar_ortografia(bloques, rangos, reglas, obs, filas_orto)
    reglas_revisor.aplicar(reglas_revisor.crear_si_falta(os.path.join(DATOS, "mis_reglas.yaml")),
                           bloques, rangos, mapa, obs, lambda t: len(extraer_citas(t)), Paragraph, Table)

    ubicar_lineas(obs, filas_orto, mapa, bloques, rangos)
    datos = datos_expediente(ruta, bloques, rangos, reglas, registro, args)
    base = os.path.splitext(os.path.basename(ruta))[0]
    orden_sev = {"Error": 0, "Advertencia": 1, "Revisar": 2}
    orden_cat = {"Estructura": 0, "Formato": 1, "Extensión": 2, "Citas": 3, "Ortografía": 4,
                 "Gramática/puntuación": 5, "Revisor": 6}
    items = sorted(obs.items, key=lambda x: (orden_cat.get(x["categoria"], 9), orden_sev.get(x["severidad"], 9)))
    resumen = {c: sum(1 for x in obs.items if x["categoria"] == c) for c in orden_cat}

    formato = getattr(args, "formato", "resumen")
    salida = os.path.join(carpeta_salida, f"Observaciones - {base}.docx")
    if formato == "detallado":
        escribir_word(salida, datos, items, filas_orto, motor, resumen, mapa)
    else:
        lista = redactar(items, filas_orto, getattr(args, "extras", ()) or (), reglas.get("tipo", "proyecto"))
        escribir_resumen(salida, datos, lista)
        if formato == "ambos":
            detalle = os.path.join(carpeta_salida, f"Observaciones - {base} (detalle).docx")
            escribir_word(detalle, datos, items, filas_orto, motor, resumen, mapa)
    if args.excel:
        xls = os.path.join(carpeta_salida, f"Observaciones - {base}.xlsx")
        escribir_reporte(xls, f"{os.path.basename(ruta)} ({reglas['tipo']})", obs, filas_orto, motor, None)

    n_err = sum(1 for x in obs.items if x["severidad"] == "Error")
    faltan = [k for k in ("tesista", "asesor") if not datos.get(k)]
    nota = f"  [completar en el Word: {', '.join(faltan)}]" if faltan else ""
    aviso(f"{os.path.basename(ruta)}: {n_err} errores, {len(obs.items) - n_err} advertencias/revisar -> {salida}{nota}")
    return salida


class Opciones:
    """opciones de una corrida; la usan tanto la línea de comandos como la interfaz"""

    def __init__(self, tipo="proyecto", salida="reportes", excel=False, forzar=False, registro=None,
                 revisor="", tesista="", proyecto="", asesor="", fecha="", motor="",
                 formato="resumen", codigo="", n_revision="", extras=()):
        self.tipo = tipo
        self.salida = salida
        self.excel = excel
        self.forzar = forzar
        self.registro = registro or os.path.join(DATOS, "registro.csv")
        self.revisor = revisor
        self.tesista = tesista
        self.proyecto = proyecto
        self.asesor = asesor
        self.fecha = fecha
        self.motor = motor          # "" = lo que diga el archivo de reglas
        self.formato = formato      # resumen | detallado | ambos
        self.codigo = codigo
        self.n_revision = n_revision
        self.extras = list(extras)


def carpeta_reglas():
    preparar_datos()
    propia = os.path.join(DATOS, "reglas")
    return propia if os.path.isdir(propia) else os.path.join(RECURSOS, "reglas")


def tipos_disponibles():
    """[(tipo, configurado)]"""
    salida = []
    for f in sorted(os.listdir(carpeta_reglas())):
        if not f.endswith(".yaml"):
            continue
        tipo = os.path.splitext(f)[0]
        try:
            with open(os.path.join(carpeta_reglas(), f), encoding="utf8") as fh:
                reglas = yaml.safe_load(fh) or {}
            salida.append((tipo, bool(reglas.get("secciones"))))
        except Exception:
            salida.append((tipo, False))
    return salida


def cargar_reglas(tipo):
    with open(os.path.join(carpeta_reglas(), tipo + ".yaml"), encoding="utf8") as fh:
        reglas = yaml.safe_load(fh) or {}
    reglas.setdefault("tipo", tipo)
    if not reglas.get("secciones"):
        raise ValueError(
            f"Las reglas de '{tipo}' no están configuradas todavía.\n"
            f"Copia la plantilla oficial en plantillas/{tipo}.docx y ejecuta:\n"
            f"    python generar_reglas.py plantillas/{tipo}.docx --tipo {tipo}")
    return reglas


def expandir(entradas):
    """convierte archivos y carpetas en una lista de .docx; devuelve (archivos, avisos)"""
    archivos, avisos = [], []
    for entrada in entradas:
        if not os.path.exists(entrada):
            avisos.append(f"No existe: {entrada}")
        elif os.path.isdir(entrada):
            todos = sorted(os.listdir(entrada))
            archivos += [os.path.join(entrada, f) for f in todos
                         if f.lower().endswith(".docx") and not f.startswith("~$")]
            if any(f.lower().endswith((".pdf", ".doc")) for f in todos):
                avisos.append(f"En {os.path.basename(entrada)} hay archivos .pdf o .doc que no se revisan. "
                              "Pedir el .docx (un .doc se abre en Word y se guarda como .docx).")
        elif not entrada.lower().endswith(".docx"):
            avisos.append(f"Se omite {os.path.basename(entrada)}: solo se revisan archivos .docx.")
        else:
            archivos.append(entrada)
    return archivos, avisos


def ejecutar(entradas, opciones, aviso=print):
    """revisa todo y devuelve la lista de reportes generados"""
    preparar_datos()
    reglas = cargar_reglas(opciones.tipo)
    archivos, avisos = expandir(entradas)
    for a in avisos:
        aviso(a)
    if not archivos:
        raise ValueError("No hay archivos .docx para revisar.")
    if len(archivos) > 1 and any([opciones.tesista, opciones.asesor, opciones.fecha]):
        raise ValueError("Los datos de tesista, asesor y fecha solo valen para un archivo. "
                         "Para varios, usa registro.csv.")
    registro = leer_registro(opciones.registro)
    os.makedirs(opciones.salida, exist_ok=True)
    hechos = []
    for f in archivos:
        try:
            r = revisar(f, reglas, opciones.salida, registro, opciones, opciones.forzar, aviso)
            if r:
                hechos.append(r)
        except Exception as ex:
            aviso(f"{os.path.basename(f)}: no se pudo revisar ({type(ex).__name__}: {ex})")
    return hechos


def main():
    preparar_datos()
    tipos = [t for t, _ in tipos_disponibles()]
    ap = argparse.ArgumentParser(description="Revisor de formato de tesis FINESI")
    ap.add_argument("entrada", nargs="+", help="uno o varios archivos .docx, o una carpeta")
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
    g = ap.add_argument_group("datos del expediente (solo si revisas UN archivo; si no, usa registro.csv)")
    g.add_argument("--tesista", default="")
    g.add_argument("--proyecto", default="", help="título del proyecto; por defecto se toma del documento")
    g.add_argument("--asesor", default="")
    g.add_argument("--fecha", default="", help="fecha de presentación, ej. 05/09/2026")
    g.add_argument("--extra", dest="extras", action="append", default=[],
                   help="observación adicional; se puede repetir")
    a = ap.parse_args()

    try:
        reglas = cargar_reglas(a.tipo)
    except ValueError as ex:
        sys.exit(str(ex))

    try:
        ejecutar(a.entrada, a, print)
    except ValueError as ex:
        sys.exit(str(ex))


if __name__ == "__main__":
    main()
