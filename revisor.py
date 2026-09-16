"""
Revisor automatico de Proyectos de Tesis FINESI (formato, estructura, citas APA y ortografia).
Fachada (Facade) que redirige al motor modularizado para mantener compatibilidad hacia atrás.
"""

from motor.orquestador import *
from motor.utils import *
from motor.estructura import *
from motor.formato import *
from motor.citas import *
from motor.ortografia import *
from motor.extension import *

# Si este script se ejecuta directamente (CLI)
if __name__ == '__main__':
    import argparse
    import json
    parser = argparse.ArgumentParser()
    parser.add_argument("archivo", help="Ruta al documento .docx o carpeta")
    parser.add_argument("--tipo", required=True, help="Tipo de documento (proyecto, borrador, etc.)")
    parser.add_argument("--motor", help="URL del servidor LanguageTool (opcional)")
    parser.add_argument("--forzar", action="store_true", help="Revisar aunque no parezca del tipo indicado")
    parser.add_argument("--salida", help="Carpeta de salida para reportes (defecto: dist/reportes)")
    parser.add_argument("--json", action="store_true", help="Salida en JSON por stdout")
    args = parser.parse_args()

    reglas = tipos_disponibles(args.tipo)
    if not reglas:
        print(f"Error: tipo '{args.tipo}' no reconocido. Tipos disponibles: {', '.join(t[0] for t in tipos_disponibles())}")
        sys.exit(1)

    salida = args.salida or os.path.join(DATOS, "dist", "reportes")
    
    rutas = [args.archivo]
    if os.path.isdir(args.archivo):
        rutas = [os.path.join(args.archivo, a) for a in os.listdir(args.archivo) if a.endswith(".docx") and not a.startswith("~")]
        if not rutas:
            print(f"No se encontraron archivos .docx en {args.archivo}")
            sys.exit(1)

    for ruta in rutas:
        res = revisar(ruta, reglas, salida, args=args, forzar=args.forzar)
        if args.json:
            print(json.dumps(res, ensure_ascii=False))
