"""
Reglas propias del revisor.

Se escriben en mis_reglas.yaml y se aplican además de las reglas del programa.
No hay código que tocar: el revisor agrega una regla nueva y se aplica en la siguiente revisión.

Tipos de condición:
    si_aparece      lista de textos o expresiones; si alguno aparece, se observa
    si_no_aparece   lista de textos; si ninguno aparece, se observa
    minimo_citas    número mínimo de citas (autor, año) en esa sección
    minimo_palabras número mínimo de palabras en esa sección
"""
import os
import re
import unicodedata

import yaml


def _norm(t):
    t = unicodedata.normalize("NFD", (t or "").lower())
    t = "".join(c for c in t if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", t)


def cargar(ruta):
    if not ruta or not os.path.exists(ruta):
        return []
    try:
        with open(ruta, encoding="utf8") as fh:
            datos = yaml.safe_load(fh) or {}
        reglas = datos.get("reglas") or []
        return [r for r in reglas if isinstance(r, dict) and r.get("observacion") and r.get("activa", True)]
    except Exception:
        return []


def _texto_de(bloques, rangos, donde, Paragraph, Table):
    """(texto, indice del primer bloque) del ámbito pedido"""
    if not donde or _norm(donde) in ("documento", "todo", "todo el documento"):
        ini, fin = 0, len(bloques)
    elif donde in rangos:
        ini, fin = rangos[donde]
    else:
        return None, None
    partes = []
    for b in bloques[ini:fin]:
        if isinstance(b, Paragraph):
            partes.append(b.text)
        elif isinstance(b, Table):
            partes += [c.text for f in b.rows for c in f.cells]
    return "\n".join(partes), ini


def aplicar(ruta_reglas, bloques, rangos, mapa, obs, contar_citas, Paragraph, Table):
    """agrega al informe las observaciones de las reglas propias del revisor"""
    for regla in cargar(ruta_reglas):
        donde = regla.get("donde") or "documento"
        texto, ini = _texto_de(bloques, rangos, donde, Paragraph, Table)
        if texto is None:
            continue   # la sección no existe: ya lo reporta la revisión de estructura
        severidad = regla.get("severidad", "Error")
        if severidad not in ("Error", "Advertencia", "Revisar"):
            severidad = "Error"
        ubicacion = donde if donde in rangos else "Documento"
        ancla = bloques[rangos[donde][0]].text if donde in rangos else None
        plano = _norm(texto)
        disparo, fragmento = False, ""

        patrones = regla.get("si_aparece") or []
        if patrones:
            for p in patrones:
                try:
                    m = re.search(_norm(p), plano)
                except re.error:
                    m = re.search(re.escape(_norm(p)), plano)
                if m:
                    disparo = True
                    fragmento = texto[max(0, m.start() - 10):m.start() + 50]
                    break

        faltantes = regla.get("si_no_aparece") or []
        if faltantes and not disparo:
            disparo = not any(_norm(p) in plano for p in faltantes)

        minimo = regla.get("minimo_citas")
        if minimo and not disparo:
            n = contar_citas(texto)
            if n < int(minimo):
                disparo = True
                regla = dict(regla, detalle=f"Tiene {n} cita(s); se esperan al menos {minimo}.")

        minimo_p = regla.get("minimo_palabras")
        if minimo_p and not disparo:
            n = len(texto.split())
            if n < int(minimo_p):
                disparo = True
                regla = dict(regla, detalle=f"Tiene {n} palabra(s); se esperan al menos {minimo_p}.")

        if not disparo:
            continue

        linea, nlinea = "", None
        if mapa:
            base = fragmento or ancla or ""
            if base:
                linea = mapa.buscar(base, ancla if fragmento else None)
                nlinea = mapa.numero(base, ancla if fragmento else None)
        obs.add("Revisor", severidad, ubicacion, regla["observacion"], regla.get("detalle", ""))
        obs.items[-1]["linea"] = linea
        obs.items[-1]["nlinea"] = nlinea
        obs.items[-1]["propia"] = True


PLANTILLA = """# Reglas propias del revisor
#
# Se aplican además de las que ya trae el programa. Para agregar una, copia un bloque y edítalo.
# No hace falta reconstruir el ejecutable: los cambios valen en la siguiente revisión.
#
# Campos:
#   observacion      texto que sale en la hoja de revisión (obligatorio)
#   donde            "documento" o el nombre exacto de una sección del esquema
#   si_aparece       lista de textos; si alguno aparece, se observa
#   si_no_aparece    lista de textos; si NINGUNO aparece, se observa
#   minimo_citas     número mínimo de citas (autor, año) en esa sección
#   minimo_palabras  número mínimo de palabras en esa sección
#   severidad        Error, Advertencia o Revisar
#   activa           false para desactivarla sin borrarla

reglas:
  - observacion: "En anexos, incluir matriz de consistencia"
    donde: documento
    si_no_aparece: ["matriz de consistencia"]
    severidad: Error

  - observacion: "Incrementar número de citas (antecedentes)"
    donde: "Antecedentes del proyecto"
    minimo_citas: 5
    severidad: Error

  - observacion: "Redactar en tercera persona o impersonal"
    donde: documento
    si_aparece: ["\\\\brealizare\\\\b", "\\\\bhice\\\\b", "\\\\bmi tesis\\\\b", "\\\\byo \\\\b", "\\\\bnuestro trabajo\\\\b"]
    severidad: Advertencia
    activa: false

  - observacion: "Los objetivos deben empezar con verbo en infinitivo"
    donde: "Objetivo general"
    si_no_aparece: ["ar ", "er ", "ir "]
    severidad: Revisar
    activa: false
"""


def crear_si_falta(ruta):
    if ruta and not os.path.exists(ruta):
        try:
            with open(ruta, "w", encoding="utf8") as fh:
                fh.write(PLANTILLA)
        except Exception:
            pass
    return ruta
