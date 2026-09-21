"""
Mide la calidad del motor contra un corpus de tesis ya revisadas a mano.

Los tests de tests/ congelan el comportamiento: avisan si algo cambia. Esto es otra
cosa: dice si el motor acierta. Sin esto, los umbrales del motor (el 82 del parecido
de títulos, el 88 del emparejado de autores, el 90 de 'parece otro tipo') son
corazonadas que nadie puede evaluar.

Cómo armar el corpus
--------------------
1. Junta en una carpeta los .docx de tesis que ya revisaste a mano.
2. Genera las fichas de partida:

       python scripts/medir_calidad.py corpus/ --tipo proyecto --crear-fichas

   Por cada documento se crea <nombre>.esperado.yaml con lo que el motor encontró.
3. Edita cada ficha: borra las observaciones que son falsos positivos y agrega a mano
   las que el motor no vio. Eso convierte la ficha en la verdad de referencia.
4. A partir de ahí, mide cuando quieras:

       python scripts/medir_calidad.py corpus/ --tipo proyecto

Qué reporta
-----------
Por categoría y en total:
  aciertos          observaciones reales que el motor encontró
  falsos positivos  observaciones que el motor inventó (le cuestan credibilidad)
  no detectadas     observaciones reales que el motor no vio (el fallo más grave:
                    el revisor cree que el documento está limpio)
  precisión         aciertos / (aciertos + falsos positivos)
  exhaustividad     aciertos / (aciertos + no detectadas)
"""
import argparse
import os
import sys

import yaml
from rapidfuzz import fuzz

sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "src"))
from core import orquestador  # noqa: E402
from core.orquestador import Opciones, cargar_reglas  # noqa: E402

PARECIDO_MINIMO = 85   # dos textos de observación que se parecen tanto son el mismo hallazgo


def observaciones_del_motor(ruta_docx, tipo, motor):
    """Corre la revisión real y devuelve [(categoria, observacion)]."""
    capturado = {}

    def espia(ruta, datos, obs_items, filas_orto, motor_orto, resumen, mapa=None):
        capturado["items"] = obs_items

    original = orquestador.escribir_word
    orquestador.escribir_word = espia
    try:
        opciones = Opciones(tipo=tipo, salida=os.path.dirname(ruta_docx) or ".",
                            motor=motor, formato="detallado")
        orquestador.revisar(ruta_docx, cargar_reglas(tipo), opciones.salida, {}, opciones,
                            forzar=True, aviso=lambda _m: None)
    finally:
        orquestador.escribir_word = original
    return [(it["categoria"], it["observacion"]) for it in capturado.get("items", [])]


def ruta_ficha(ruta_docx):
    return os.path.splitext(ruta_docx)[0] + ".esperado.yaml"


def leer_ficha(ruta_docx):
    ruta = ruta_ficha(ruta_docx)
    if not os.path.exists(ruta):
        return None
    with open(ruta, encoding="utf8") as fh:
        datos = yaml.safe_load(fh) or {}
    return [(x.get("categoria", ""), x.get("observacion", ""))
            for x in (datos.get("esperadas") or []) if isinstance(x, dict)]


CABECERA_FICHA = """# Verdad de referencia para este documento.
#
# Esto es lo que el motor encontró. Revisalo a mano:
#   - borra las observaciones que sean FALSOS POSITIVOS
#   - agrega las que el motor NO vio y deberían estar
#
# Luego: python scripts/medir_calidad.py <carpeta> --tipo <tipo>

"""


def escribir_ficha(ruta_docx, halladas):
    # Se serializa la estructura completa con yaml en vez de armar el texto a mano:
    # las observaciones traen comillas, dos puntos y saltos que hay que escapar bien.
    cuerpo = yaml.safe_dump(
        {"esperadas": [{"categoria": cat, "observacion": obs} for cat, obs in halladas]},
        allow_unicode=True, sort_keys=False, default_flow_style=False, width=100)
    with open(ruta_ficha(ruta_docx), "w", encoding="utf8") as fh:
        fh.write(CABECERA_FICHA + cuerpo)


def emparejar(esperadas, halladas):
    """(aciertos, falsos_positivos, no_detectadas) por comparación difusa del texto."""
    libres = list(halladas)
    aciertos, no_detectadas = [], []
    for cat, obs in esperadas:
        mejor, puntaje = None, 0
        for k, (c2, o2) in enumerate(libres):
            if c2 != cat:
                continue
            p = fuzz.ratio(obs.lower(), o2.lower())
            if p > puntaje:
                mejor, puntaje = k, p
        if mejor is not None and puntaje >= PARECIDO_MINIMO:
            aciertos.append((cat, obs))
            libres.pop(mejor)
        else:
            no_detectadas.append((cat, obs))
    return aciertos, libres, no_detectadas


def tabla(resultados):
    categorias = sorted({c for grupo in resultados.values() for c, _ in grupo})
    print(f"\n{'Categoría':<22}{'aciertos':>9}{'falsos+':>9}{'no det.':>9}"
          f"{'precisión':>11}{'exhaust.':>10}")
    print("-" * 70)
    for cat in categorias + ["TOTAL"]:
        if cat == "TOTAL":
            ac = len(resultados["aciertos"])
            fp = len(resultados["falsos"])
            nd = len(resultados["no_detectadas"])
            print("-" * 70)
        else:
            ac = sum(1 for c, _ in resultados["aciertos"] if c == cat)
            fp = sum(1 for c, _ in resultados["falsos"] if c == cat)
            nd = sum(1 for c, _ in resultados["no_detectadas"] if c == cat)
        prec = ac / (ac + fp) if (ac + fp) else float("nan")
        exh = ac / (ac + nd) if (ac + nd) else float("nan")
        print(f"{cat:<22}{ac:>9}{fp:>9}{nd:>9}{prec:>10.0%}{exh:>10.0%}")


def main():
    ap = argparse.ArgumentParser(description="Mide precisión y exhaustividad del motor")
    ap.add_argument("carpeta", help="carpeta con los .docx del corpus")
    ap.add_argument("--tipo", required=True, help="tipo de documento (proyecto, borrador)")
    ap.add_argument("--motor", default="diccionario",
                    help="motor de ortografía; 'diccionario' evita depender de Java")
    ap.add_argument("--crear-fichas", action="store_true",
                    help="genera las fichas .esperado.yaml de partida y no mide nada")
    ap.add_argument("--detalle", action="store_true",
                    help="lista cada falso positivo y cada observación no detectada")
    a = ap.parse_args()

    if not os.path.isdir(a.carpeta):
        sys.exit(f"No existe la carpeta '{a.carpeta}'.")
    docs = sorted(os.path.join(a.carpeta, f) for f in os.listdir(a.carpeta)
                  if f.lower().endswith(".docx") and not f.startswith("~$"))
    if not docs:
        sys.exit(f"No hay .docx en '{a.carpeta}'.")

    acumulado = {"aciertos": [], "falsos": [], "no_detectadas": []}
    sin_ficha = []
    for ruta in docs:
        nombre = os.path.basename(ruta)
        try:
            halladas = observaciones_del_motor(ruta, a.tipo, a.motor)
        except Exception as ex:
            print(f"  {nombre}: no se pudo revisar ({type(ex).__name__}: {ex})")
            continue

        if a.crear_fichas:
            escribir_ficha(ruta, halladas)
            print(f"  ficha creada: {os.path.basename(ruta_ficha(ruta))} "
                  f"({len(halladas)} observaciones para revisar)")
            continue

        esperadas = leer_ficha(ruta)
        if esperadas is None:
            sin_ficha.append(nombre)
            continue
        ac, fp, nd = emparejar(esperadas, halladas)
        acumulado["aciertos"] += ac
        acumulado["falsos"] += fp
        acumulado["no_detectadas"] += nd
        print(f"  {nombre}: {len(ac)} aciertos, {len(fp)} falsos+, {len(nd)} no detectadas")
        if a.detalle:
            for c, o in fp:
                print(f"      falso positivo   [{c}] {o[:90]}")
            for c, o in nd:
                print(f"      no detectada     [{c}] {o[:90]}")

    if a.crear_fichas:
        print("\nAhora edita las fichas: borra los falsos positivos y agrega lo que falte.")
        return

    if sin_ficha:
        print(f"\nSin ficha de referencia ({len(sin_ficha)}): {', '.join(sin_ficha)}")
        print("Generalas con --crear-fichas.")
    if any(acumulado.values()):
        tabla(acumulado)
        print("\nLas 'no detectadas' son el número que más importa: son los errores que el\n"
              "revisor no va a ver y va a dar por buenos.")
    else:
        print("\nNo hay nada medido todavía.")


if __name__ == "__main__":
    main()
