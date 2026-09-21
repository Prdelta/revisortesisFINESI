"""
Ortografía, gramática y tipografía. Usa LanguageTool si hay Java disponible;
si no, cae al diccionario hunspell es_PE (solo ortografía).
"""
import os
import re
import shutil
from collections import defaultdict

from docx.text.paragraph import Paragraph

from .config import ruta_recurso
from .citas import RE_NARR, RE_PAREN
from .utils import corto, juntar, ubic


def cargar_permitidas(ruta):
    if not ruta or not os.path.exists(ruta):
        return set()
    with open(ruta, encoding="utf8") as fh:
        return {l.strip().lower() for l in fh if l.strip() and not l.startswith("#")}


def revisar_tipografia(bloques, rangos, obs):
    reglas_tipo = [
        (r"\S  +\S", "Doble espacio entre palabras"),
        (r"\s+[,.;:](?!\d)", "Espacio antes de signo de puntuación"),
        (r"[a-záéíóúñ][,;](?=[A-Za-zÁÉÍÓÚÑáéíóúñ])", "Falta espacio después de coma o punto y coma"),
        (r"[a-záéíóúñ]\.(?=[A-ZÁÉÍÓÚÑ][a-záéíóúñ])", "Falta espacio después de punto"),
        (r"\(\s+|\s+\)", "Espacio dentro de paréntesis"),
    ]
    conteo = defaultdict(list)
    for i, b in enumerate(bloques):
        if not isinstance(b, Paragraph):
            continue
        t = b.text
        for patron, nombre in reglas_tipo:
            for m in re.finditer(patron, t):
                if "doi" in t[max(0, m.start()-30):m.end()].lower() or "http" in t[max(0, m.start()-40):m.end()].lower():
                    continue
                conteo[nombre].append(ubic(i, rangos, bloques, "..." + t[max(0, m.start()-15):m.end()+15]))
    for nombre, lugares in conteo.items():
        obs.add("Ortografía", "Advertencia", juntar(lugares[:4]), f"{nombre} ({len(lugares)} caso(s))")


def spans_citas(t):
    return [(m.start(), m.end()) for r in (RE_PAREN, RE_NARR) for m in r.finditer(t)] + \
           [(m.start(), m.end()) for m in re.finditer(r"https?://\S+", t)]


def ortografia_languagetool(textos, permitidas):
    import language_tool_python
    tool = language_tool_python.LanguageTool("es")
    try:
        return _languagetool(tool, textos, permitidas)
    finally:
        tool.close()   # sin esto, un fallo a mitad del bucle dejaba el proceso Java vivo


def _languagetool(tool, textos, permitidas):
    salida = []
    for i, t in textos:
        excluir = spans_citas(t)
        for m in tool.check(t):
            largo = getattr(m, "error_length", None) or getattr(m, "errorLength", 0)
            regla = getattr(m, "rule_id", None) or getattr(m, "ruleId", "")
            tipo = getattr(m, "rule_issue_type", None) or getattr(m, "ruleIssueType", "")
            frag = t[m.offset:m.offset + largo]
            if any(a <= m.offset < b for a, b in excluir) or frag.lower() in permitidas:
                continue
            if tipo == "misspelling" and frag[:1].isupper() and m.offset > 0 and not re.search(r"[.!?:]\s*$", t[:m.offset]):
                continue  # probable nombre propio
            if "WHITESPACE" in regla:
                continue  # los espacios los revisa revisar_tipografia
            if regla == "UPPERCASE_SENTENCE_START" and "." not in t.strip()[:-1]:
                continue  # listas o palabras clave, no son oraciones
            clase = "Ortografía" if tipo == "misspelling" else "Gramática/puntuación"
            salida.append((i, frag, f"[{clase}] {m.message}", ", ".join(m.replacements[:3]), m.context))
    return salida


def ortografia_hunspell(textos, permitidas, ruta_dic):
    from spylls.hunspell import Dictionary
    base = ruta_dic if os.path.isabs(ruta_dic) else ruta_recurso(ruta_dic)
    if not os.path.exists(base + ".dic"):
        raise FileNotFoundError(
            f"No se encontró el diccionario '{os.path.basename(base)}' en "
            f"{os.path.dirname(base)}. Sin él no se puede revisar la ortografía "
            f"cuando no hay Java para LanguageTool.")
    d = Dictionary.from_files(base)
    desconocidas = defaultdict(list)
    for i, t in textos:
        # quitamos citas y URLs para no marcar apellidos ni direcciones
        limpio = RE_NARR.sub(" ", RE_PAREN.sub(" ", t))
        limpio = re.sub(r"https?://\S+|\S+@\S+", " ", limpio)
        palabras = re.finditer(r"[A-Za-zÁÉÍÓÚÑÜáéíóúñü]+", limpio)
        for m in palabras:
            w = m.group(0)
            if len(w) < 3 or w.isupper() or w.lower() in permitidas:
                continue
            ini = m.start()
            inicio_oracion = ini == 0 or re.search(r"[.!?:]\s*$", limpio[:ini]) is not None
            if w[0].isupper() and not inicio_oracion:
                continue  # probable nombre propio
            if d.lookup(w) or d.lookup(w.lower()):
                continue
            desconocidas[w].append((i, limpio[max(0, ini-25):ini+len(w)+25]))
    salida = []
    for w, lugares in desconocidas.items():
        sug = []
        try:
            for s in d.suggest(w):
                sug.append(s)
                if len(sug) == 3:
                    break
        except Exception:
            pass   # sin sugerencias: la palabra se reporta igual, que es lo que importa
        i, ctx = lugares[0]
        salida.append((i, w, f"Palabra no reconocida ({len(lugares)} vez/veces)", ", ".join(sug), ctx))
    return salida


def revisar_ortografia(bloques, rangos, reglas, obs, filas_orto):
    o = reglas.get("ortografia") or {}
    omitir = set()
    for n in o.get("omitir_secciones", []):
        if n in rangos:
            omitir.update(range(*rangos[n]))
    textos = [(i, b.text) for i, b in enumerate(bloques)
              if isinstance(b, Paragraph) and b.text.strip() and i not in omitir]
    permitidas = cargar_permitidas(ruta_recurso(o.get("palabras_permitidas", "permitidas.txt")))

    resultados, motor = None, None
    modo = o.get("motor", "auto")
    if modo == "auto" and not shutil.which("java"):
        modo = "diccionario"   # sin Java, LanguageTool no puede funcionar
    if modo in ("auto", "languagetool"):
        try:
            resultados = ortografia_languagetool(textos, permitidas)
            motor = "LanguageTool (ortografía, gramática y puntuación)"
        except Exception as ex:
            if modo == "languagetool":
                obs.add("Ortografía", "Revisar", "Sistema", "LanguageTool no disponible", str(ex)[:150])
    if resultados is None:
        resultados = ortografia_hunspell(textos, permitidas, o.get("diccionario", "dic/es_PE"))
        motor = "Diccionario hunspell es_PE (solo ortografía, no gramática)"

    resultados = [r for r in resultados if r[1].lower() not in permitidas]
    for i, palabra, msg, sug, ctx in resultados:
        seccion = ubic(i, rangos, bloques, None, 25).split(":")[0].strip()
        # para ubicar la linea se busca desde la palabra con error hacia adelante
        pos = ctx.find(palabra)
        desde_palabra = ctx[pos:pos + 55] if pos >= 0 else palabra
        filas_orto.append(dict(parrafo=seccion, seccion=seccion, palabra=palabra, problema=msg,
                               sugerencias=sug, contexto=corto(ctx, 90), frag=desde_palabra))
    if resultados:
        obs.add("Ortografía", "Revisar", "Ver el detalle al final del documento",
                f"{len(resultados)} posible(s) error(es) de ortografía o gramática",
                f"Motor: {motor}. Revisar la lista: puede haber términos técnicos correctos (agregarlos a permitidas.txt).")
    revisar_tipografia(bloques, rangos, obs)
    return motor


