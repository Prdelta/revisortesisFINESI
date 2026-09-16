# Receta de PyInstaller para generar el ejecutable de Windows.
# Se usa así (desde esta carpeta, con el entorno ya instalado):
#     pyinstaller --noconfirm --clean RevisorTesisFINESI.spec
# El resultado queda en dist/RevisorTesisFINESI.exe

import os

from PyInstaller.utils.hooks import collect_all, collect_data_files

datos = [
    ("data/reglas", "reglas"),
    ("data/plantillas", "plantillas"),
    ("dic", "dic"),
    ("data/permitidas.txt", "."),
    ("data/mis_reglas.yaml", "."),
    ("data/registro.csv", "."),
    ("LEEME.md", "."),
    ("icono.ico", "."),
]
datos += [("icono.png", ".")]
datos = [(o, d) for o, d in datos if os.path.exists(o)]

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
