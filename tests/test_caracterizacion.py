"""
Tests de caracterización: no dicen lo que el programa *debería* observar, sino que
congelan lo que observa hoy sobre los documentos de 'ejemplos/'. Si un cambio altera
el resultado, salta acá y hay que decidir si es una mejora (se regenera la instantánea)
o una regresión.
"""
import os

import pytest

from conftest import comparar, observaciones_del_excel

from core.config import tipos_disponibles
from core.orquestador import cargar_reglas, ejecutar, validar_reglas

TESIS = ["proyecto_con_errores.docx", "proyecto_correcto.docx"]

HAY_BORRADOR = "borrador" in [t for t, _ in tipos_disponibles()]
sin_borrador = pytest.mark.skipif(not HAY_BORRADOR,
                                  reason="no hay reglas de borrador en esta instalación")


def _revisar(ruta, opciones):
    """revisa un documento en silencio y devuelve las observaciones del Excel"""
    hechos = ejecutar([ruta], opciones, aviso=lambda m: None)
    base = os.path.splitext(os.path.basename(ruta))[0]
    xlsx = os.path.join(opciones.salida, f"Observaciones - {base}.xlsx")
    return hechos, xlsx


@pytest.mark.parametrize("documento", TESIS)
def test_observaciones_congeladas(documento, ejemplos, opciones):
    hechos, xlsx = _revisar(os.path.join(ejemplos, documento), opciones)
    assert len(hechos) == 1, "debería salir una hoja de revisión por documento"
    assert os.path.exists(xlsx), "con excel=True tiene que salir también el .xlsx"
    filas = observaciones_del_excel(xlsx)
    comparar(os.path.splitext(documento)[0], "\n".join(filas) + "\n")


def test_el_documento_con_errores_da_mas_observaciones(ejemplos, opciones):
    """
    La única afirmación de fondo que vale fijar: el documento preparado con errores
    tiene que dar más observaciones que el preparado correcto. Si esto se invierte,
    el motor está roto de una forma que una instantánea no delata.
    """
    conteos = {}
    for documento in TESIS:
        _, xlsx = _revisar(os.path.join(ejemplos, documento), opciones)
        conteos[documento] = len(observaciones_del_excel(xlsx))
    assert conteos["proyecto_con_errores.docx"] > conteos["proyecto_correcto.docx"], conteos


def test_la_hoja_de_revision_se_puede_abrir(ejemplos, opciones):
    """el entregable formal es el Word: tiene que ser un .docx válido y traer la firma"""
    from docx import Document
    ejecutar([os.path.join(ejemplos, "proyecto_con_errores.docx")], opciones,
             aviso=lambda m: None)
    hoja = os.path.join(opciones.salida, "Observaciones - proyecto_con_errores.docx")
    doc = Document(hoja)
    texto = "\n".join(p.text for p in doc.paragraphs)
    assert "Revisor de prueba" in texto, "la firma del revisor tiene que aparecer en la hoja"
    assert len(texto.strip()) > 200


def test_formato_ambos_genera_los_dos_archivos(ejemplos, opciones):
    opciones.formato = "ambos"
    ejecutar([os.path.join(ejemplos, "proyecto_con_errores.docx")], opciones,
             aviso=lambda m: None)
    for nombre in ("Observaciones - proyecto_con_errores.docx",
                   "Observaciones - proyecto_con_errores (detalle).docx"):
        assert os.path.exists(os.path.join(opciones.salida, nombre)), nombre


@sin_borrador
def test_tipo_equivocado_no_genera_reporte(ejemplos, opciones):
    """
    Si se elige 'borrador' para un proyecto, el programa avisa y no genera la hoja:
    una hoja con el esquema equivocado es peor que ninguna.
    """
    opciones.tipo = "borrador"
    avisos = []
    hechos = ejecutar([os.path.join(ejemplos, "proyecto_con_errores.docx")], opciones, avisos.append)
    assert hechos == []
    assert any("NO REVISADO" in a for a in avisos), avisos


@sin_borrador
def test_forzar_revisa_aunque_parezca_de_otro_tipo(ejemplos, opciones):
    opciones.tipo = "borrador"
    opciones.forzar = True
    hechos = ejecutar([os.path.join(ejemplos, "proyecto_con_errores.docx")], opciones,
                      aviso=lambda m: None)
    assert len(hechos) == 1


def test_las_reglas_de_cada_tipo_son_validas():
    """las reglas que se distribuyen tienen que pasar su propia validación"""
    for tipo, _ in tipos_disponibles():
        assert validar_reglas(cargar_reglas(tipo)) == [], f"las reglas de '{tipo}' tienen fallos"
