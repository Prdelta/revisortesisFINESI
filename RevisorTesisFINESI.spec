# Receta de PyInstaller para generar el ejecutable de Windows.
# Se usa así (desde esta carpeta, con el entorno ya instalado):
#     pyinstaller --noconfirm --clean RevisorTesisFINESI.spec
# El resultado queda en dist/RevisorTesisFINESI.exe

import os

from PyInstaller.utils.hooks import collect_all, collect_data_files

datos = [
    ("data/reglas", "reglas"),
    ("data/plantillas", "plantillas"),
    ("data/dic", "dic"),
    ("data/permitidas.txt", "."),
    ("data/mis_reglas.yaml", "."),
    ("data/registro.csv", "."),
    ("icono.ico", "."),
    ("icono.png", "."),
]
# Antes se filtraba en silencio lo que no existiera: así el diccionario quedó fuera
# del .exe durante varias compilaciones sin que nadie se enterara. Ahora falla fuerte.
faltan = [o for o, _ in datos if not os.path.exists(o)]
if faltan:
    raise SystemExit("Faltan recursos que deben ir dentro del .exe: " + ", ".join(faltan))

# python-docx necesita su plantilla interna; language_tool_python sus datos
datos += collect_data_files("docx")
datos += collect_data_files("spylls")
lt_datos, lt_bin, lt_ocultos = collect_all("language_tool_python")
datos += lt_datos

ocultos = lt_ocultos + [
    "docx", "openpyxl", "yaml", "rapidfuzz", "spylls", "spylls.hunspell", "pypdf",
    "ttkbootstrap"
]

a = Analysis(
    ["src/app.py"],
    pathex=["src"],
    binaries=lt_bin,
    datas=datos,
    hiddenimports=ocultos,
    hookspath=[],
    runtime_hooks=[],
    excludes=["matplotlib", "numpy", "pandas", "scipy", "PyQt5", "PySide2", "notebook", "IPython"],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="RevisorTesisFINESI",
    console=False,          # aplicación de ventana, sin consola negra
    icon="icono.ico" if os.path.exists("icono.ico") else None,
    upx=False,
    strip=False,
    disable_windowed_traceback=False,
)
