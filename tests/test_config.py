"""
Dónde guarda sus datos el .exe. Instalado en Archivos de programa no puede escribir
junto a sí mismo: ahí los datos van a Documentos. Suelto (portable), junto al .exe.
"""
import os

from core.config import MARCADOR_INSTALADO, NOMBRE_CARPETA_DATOS, carpeta_datos_exe


def test_el_exe_suelto_guarda_junto_a_si_mismo(tmp_path):
    assert carpeta_datos_exe(str(tmp_path), documentos="D:/Docs") == str(tmp_path)


def test_el_exe_instalado_guarda_en_documentos(tmp_path):
    (tmp_path / MARCADOR_INSTALADO).write_text("instalado", encoding="utf-8")
    assert carpeta_datos_exe(str(tmp_path), documentos="D:/Docs") == \
        os.path.join("D:/Docs", NOMBRE_CARPETA_DATOS)
