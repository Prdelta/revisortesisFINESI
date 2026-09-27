"""
Conteo de páginas sin LibreOffice. El dato <Pages> que Word guarda en el archivo no
siempre está al día: un proyecto de 10 páginas lo traía en 1, y así una tesis larga
pasaba el control de extensión sin ningún aviso.
"""
import zipfile

import pytest

from core.extension import contar_paginas


def _docx_con_paginas(tmp_path, paginas):
    ruta = tmp_path / "doc.docx"
    with zipfile.ZipFile(ruta, "w") as z:
        z.writestr("docProps/app.xml", f"<Properties><Pages>{paginas}</Pages></Properties>")
    return str(ruta)


def test_el_dato_desactualizado_de_word_no_manda(tmp_path):
    pags, metodo, exacto = contar_paginas(_docx_con_paginas(tmp_path, 1), palabras=16500)
    assert pags == 30 and not exacto
    assert "desactualizado" in metodo


def test_el_dato_de_word_creible_se_usa(tmp_path):
    pags, metodo, _ = contar_paginas(_docx_con_paginas(tmp_path, 10), palabras=5818)
    assert pags == 10 and metodo.startswith("dato guardado")


def test_sin_dato_de_word_se_estima(tmp_path):
    ruta = tmp_path / "vacio.docx"
    with zipfile.ZipFile(ruta, "w") as z:
        z.writestr("word/document.xml", "<x/>")
    assert contar_paginas(str(ruta), palabras=1100)[0] == 2


def test_libreoffice_manda_sobre_todo(tmp_path):
    class Mapa:
        paginas = 12
    assert contar_paginas(_docx_con_paginas(tmp_path, 1), Mapa(), palabras=16500) == (12, "LibreOffice", True)
