import os
import sys
import glob
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

def tipos_disponibles(filtro=None):
    """Carga la lista de tipos de documentos (reglas YAML) disponibles en la carpeta data/reglas"""
    rutas = glob.glob(os.path.join(DATOS, "reglas", "*.yaml"))
    tipos = []
    for r in rutas:
        nombre = os.path.basename(r).replace(".yaml", "")
        if filtro and nombre != filtro:
            continue
        try:
            with open(r, "r", encoding="utf8") as f:
                tipos.append((nombre, yaml.safe_load(f)))
        except Exception:
            pass
    return tipos

def ruta_recurso(nombre):
    """Busca un recurso (como el ícono) en la carpeta de recursos"""
    return os.path.join(RECURSOS, nombre)

def preparar_datos():
    """Asegura que los datos existan (útil al iniciar por primera vez o desde el ejecutable)"""
    os.makedirs(os.path.join(DATOS, "reglas"), exist_ok=True)
    os.makedirs(os.path.join(DATOS, "plantillas"), exist_ok=True)

