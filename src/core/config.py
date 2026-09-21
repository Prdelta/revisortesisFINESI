import os
import sys
import glob
import shutil
import yaml

def _rutas():
    """
    Determina las rutas base para recursos (plantillas, reglas por defecto)
    y datos (reglas personalizadas, registro.csv).
    En el .exe, los recursos van dentro (MEIPASS) y los datos van junto al .exe.
    En desarrollo, ambos apuntan a la carpeta 'data/'.
    """
    if getattr(sys, "frozen", False):
        return sys._MEIPASS, os.path.dirname(sys.executable)
    
    # En desarrollo: src/core/config.py -> subir 3 niveles y entrar a 'data'
    base_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    data_dir = os.path.join(base_dir, "data")
    return data_dir, data_dir

RECURSOS, DATOS = _rutas()
BASE = DATOS

def carpeta_reglas():
    """
    Carpeta de reglas en uso: la editable del usuario si tiene algún .yaml,
    si no la que viene empaquetada. Es la única fuente de verdad: la usan
    tanto tipos_disponibles() como cargar_reglas().
    """
    propia = os.path.join(DATOS, "reglas")
    if glob.glob(os.path.join(propia, "*.yaml")):
        return propia
    return os.path.join(RECURSOS, "reglas")


def tipos_disponibles():
    """[(tipo, configurado)] de los tipos de documento con reglas disponibles"""
    tipos = []
    for ruta in sorted(glob.glob(os.path.join(carpeta_reglas(), "*.yaml"))):
        nombre = os.path.splitext(os.path.basename(ruta))[0]
        try:
            with open(ruta, "r", encoding="utf8") as f:
                reglas = yaml.safe_load(f) or {}
            tipos.append((nombre, bool(reglas.get("secciones"))))
        except Exception:
            tipos.append((nombre, False))
    return tipos

def ruta_recurso(nombre):
    """Busca un recurso (como el ícono) en la carpeta de recursos"""
    return os.path.join(RECURSOS, nombre)

def preparar_datos():
    """
    Asegura que los datos editables existan. Desde el .exe, DATOS (junto al ejecutable)
    empieza vacío: se copian ahí las reglas y plantillas empaquetadas para que el programa
    funcione en una PC nueva y el revisor pueda editarlas. Nunca pisa un archivo existente.
    """
    for sub in ("reglas", "plantillas"):
        destino = os.path.join(DATOS, sub)
        os.makedirs(destino, exist_ok=True)
        origen = os.path.join(RECURSOS, sub)
        if os.path.abspath(origen) == os.path.abspath(destino) or not os.path.isdir(origen):
            continue
        for nombre in os.listdir(origen):
            if os.path.exists(os.path.join(destino, nombre)):
                continue
            try:
                shutil.copy2(os.path.join(origen, nombre), os.path.join(destino, nombre))
            except OSError:
                pass

