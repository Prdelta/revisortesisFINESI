import os
import glob

replacements = {
    "import localizador": "from utilidades import localizador",
    "import reglas_revisor": "from utilidades import reglas_revisor",
    "from reporte_resumen import": "from reportes.reporte_resumen import",
    "from reporte_word import": "from reportes.reporte_word import",
    "import revisor": "from core.orquestador import revisar, preparar_datos, Opciones\nfrom core.config import DATOS, RECURSOS, tipos_disponibles, ruta_recurso\nfrom utilidades import reglas_revisor\n",
    "revisor.DATOS": "DATOS",
    "revisor.RECURSOS": "RECURSOS",
    "revisor.tipos_disponibles": "tipos_disponibles",
    "revisor.ruta_recurso": "ruta_recurso",
    "revisor.preparar_datos": "preparar_datos",
    "revisor.Opciones": "Opciones",
    "revisor.revisar(": "revisar(",
    "revisor.cargar_reglas": "reglas_revisor.cargar_reglas",
    "revisor.leer_registro": "orquestador.leer_registro" # wait, leer_registro is in orquestador?
}

# Actually leer_registro is in orquestador. I'll just let app.py import orquestador
replacements["import revisor"] = "import core.orquestador as orquestador\nfrom core.config import DATOS, RECURSOS, tipos_disponibles, ruta_recurso\nfrom utilidades import reglas_revisor\nfrom core.orquestador import revisar, preparar_datos, Opciones"
replacements["revisor.leer_registro"] = "orquestador.leer_registro"

# For motor/utils.py -> core/utils.py, it imports DATOS from config now
replacements["from motor.utils import *"] = "from core.utils import *\nfrom core.config import DATOS, RECURSOS"
replacements["from .utils import *"] = "from .utils import *\nfrom .config import DATOS, RECURSOS"

# App.py imports
replacements["sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))"] = "sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))\nimport core\nimport utilidades\nimport reportes"


for root, dirs, files in os.walk("src"):
    for file in files:
        if file.endswith(".py"):
            path = os.path.join(root, file)
            with open(path, "r", encoding="utf8") as f:
                content = f.read()
            
            for old, new in replacements.items():
                content = content.replace(old, new)
                
            with open(path, "w", encoding="utf8") as f:
                f.write(content)

print("Imports actualizados")
