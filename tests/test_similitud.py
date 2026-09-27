"""
Tests de la revisión del informe de Turnitin: qué se observa según lo que declara el
PDF y los topes del archivo de reglas.

No hacen falta PDF de verdad: se arma el Informe ya leído, que es justo la frontera
entre 'leer el PDF' y 'decidir qué observar'. La lectura del PDF se prueba a mano con
    python -m core.turnitin "informe.pdf"
porque los informes reales son de estudiantes y no se versionan.
"""
import pytest

from core.turnitin import (PREDETERMINADO, Informe, opciones_similitud, pct, resumen,
                           revisar_similitud)
from core.utils import Obs


def informe(**campos):
    base = dict(ruta="C:/x/Turnitin - Quispe.pdf", archivo="Turnitin - Quispe.pdf",
                es_turnitin=True, similitud=10.0)
    base.update(campos)
    return Informe(**base)


def observar(inf, **reglas_similitud):
    obs = Obs()
    revisar_similitud(inf, {"similitud": reglas_similitud} if reglas_similitud else {}, obs)
    return obs.items


def textos(items):
    return " || ".join(x["observacion"] for x in items)


# --------------------------------------------------------------------------- opciones

def test_las_opciones_se_completan_con_los_valores_de_partida():
    cfg = opciones_similitud({"similitud": {"max_pct": 30}})
    assert cfg["max_pct"] == 30
    assert cfg["max_fuente_pct"] == PREDETERMINADO["max_fuente_pct"]


def test_se_ignora_una_clave_que_no_existe():
    cfg = opciones_similitud({"similitud": {"inventada": 1}})
    assert "inventada" not in cfg


def test_sin_bloque_de_similitud_se_usan_los_de_partida():
    assert opciones_similitud({}) == dict(PREDETERMINADO, filtros_exigidos=[])


# --------------------------------------------------------------------------- el índice

def test_pasarse_del_tope_es_error():
    items = observar(informe(similitud=31.0), max_pct=25)
    assert [x["severidad"] for x in items] == ["Error"]
    assert "31%" in textos(items) and "25%" in textos(items)


def test_quedar_al_limite_es_advertencia():
    items = observar(informe(similitud=22.0), max_pct=25, margen_alerta_pct=5)
    assert [x["severidad"] for x in items] == ["Advertencia"]
    assert "al límite" in textos(items)


def test_muy_por_debajo_del_tope_no_se_observa():
    assert observar(informe(similitud=8.0), max_pct=25, margen_alerta_pct=5) == []


def test_el_tope_en_null_no_revisa_el_indice():
    assert observar(informe(similitud=90.0), max_pct=None) == []


def test_un_indice_que_no_se_pudo_leer_se_manda_a_revisar_a_mano():
    items = observar(informe(similitud=None), max_pct=25)
    assert [x["severidad"] for x in items] == ["Revisar"]
    assert "core.turnitin" in items[0]["detalle"], "hay que decir cómo ver lo que se leyó"


# --------------------------------------------------------------------------- fuentes

def test_una_sola_fuente_que_se_pasa_es_advertencia():
    items = observar(informe(fuentes=[("scielo.org", 14.0)]), max_pct=25, max_fuente_pct=10)
    assert [x["severidad"] for x in items] == ["Advertencia"]
    assert "scielo.org" in items[0]["detalle"]


def test_no_se_listan_mas_de_tres_fuentes_pasadas():
    fuentes = [(f"sitio{i}.org", 20.0) for i in range(6)]
    items = observar(informe(fuentes=fuentes), max_pct=None, max_fuente_pct=10)
    assert len(items) == 3, "la hoja no se llena con veinte fuentes"


# --------------------------------------------------------------------------- IA

def test_la_ia_solo_se_revisa_si_hay_tope():
    assert observar(informe(ia=40.0), max_ia_pct=None) == []


def test_pasarse_del_tope_de_ia_es_advertencia_y_se_aclara_que_es_referencial():
    items = observar(informe(ia=40.0), max_pct=None, max_ia_pct=20)
    assert [x["severidad"] for x in items] == ["Advertencia"]
    assert "se equivoca" in items[0]["detalle"], "no se observa un porcentaje de IA sin matizarlo"


def test_si_se_exige_ia_y_el_informe_no_la_trae_se_pide_de_nuevo():
    items = observar(informe(ia=None), max_pct=None, max_ia_pct=20)
    assert [x["severidad"] for x in items] == ["Revisar"]


# --------------------------------------------------------------------------- filtros

def test_si_se_exige_excluir_bibliografia_y_no_se_excluyo_se_observa():
    items = observar(informe(filtros={"bibliografia": False}), max_pct=None,
                     filtros_exigidos=["bibliografia"])
    assert [x["severidad"] for x in items] == ["Advertencia"]
    assert "bibliografía" in textos(items)


def test_si_no_se_sabe_con_que_filtros_se_genero_se_manda_a_verificar():
    items = observar(informe(filtros={}), max_pct=None, filtros_exigidos=["citas"])
    assert [x["severidad"] for x in items] == ["Revisar"]


def test_el_filtro_puesto_no_se_observa():
    assert observar(informe(filtros={"bibliografia": True}), max_pct=None,
                    filtros_exigidos=["bibliografia"]) == []


# --------------------------------------------------------------------------- el informe es de otro documento

def test_un_conteo_de_palabras_muy_distinto_avisa_que_el_informe_es_de_otra_version():
    obs = Obs()
    revisar_similitud(informe(palabras=4000), {"similitud": {"max_pct": None}}, obs,
                      palabras_docx=9000, nombre_docx="tesis.docx")
    assert len(obs.items) == 1
    assert "no coincide" in obs.items[0]["observacion"]
    assert "tesis.docx" in obs.items[0]["detalle"]


def test_una_diferencia_chica_de_palabras_no_se_observa():
    obs = Obs()
    revisar_similitud(informe(palabras=9200), {"similitud": {"max_pct": None}}, obs,
                      palabras_docx=9000)
    assert obs.items == []


# --------------------------------------------------------------------------- falta el informe

def test_sin_informe_no_se_observa_nada_si_no_se_exige():
    assert observar(None, exigir_informe=False) == []


def test_sin_informe_se_observa_si_se_exige():
    items = observar(None, exigir_informe=True)
    assert [x["severidad"] for x in items] == ["Error"]
    assert "Falta el informe" in items[0]["observacion"]


def test_un_pdf_que_no_es_de_turnitin_se_manda_a_revisar_y_se_pide_el_docx():
    items = observar(informe(es_turnitin=False, similitud=None))
    assert [x["severidad"] for x in items] == ["Revisar"]
    assert ".docx" in items[0]["detalle"], "si mandaron la tesis en PDF hay que pedir el Word"


def test_un_informe_que_no_se_pudo_leer_no_se_da_por_bueno():
    items = observar(informe(error="PDF cifrado"))
    assert [x["severidad"] for x in items] == ["Revisar"]
    assert "PDF cifrado" in items[0]["detalle"]


# --------------------------------------------------------------------------- severidad ajustable

def test_se_puede_bajar_la_severidad_de_similitud_desde_las_reglas():
    """
    La Dirección puede querer que la similitud avise pero no bloquee, sin tocar código.
    """
    obs = Obs({"Similitud": "Advertencia"})
    revisar_similitud(informe(similitud=31.0), {"similitud": {"max_pct": 25}}, obs)
    assert [x["severidad"] for x in obs.items] == ["Advertencia"]


# --------------------------------------------------------------------------- resumen de una línea

def test_el_resumen_junta_indice_desglose_e_ia():
    linea = resumen(informe(similitud=18.0, internet=12.0, publicaciones=4.0, estudiantes=2.0,
                            ia=5.0, palabras=9000))
    assert "similitud 18%" in linea
    assert "internet 12%" in linea
    assert "IA 5%" in linea
    assert "9000 palabras" in linea


def test_sin_informe_el_resumen_esta_vacio():
    assert resumen(None) == ""


def test_el_resumen_dice_cuando_no_se_pudo_leer():
    assert "no se pudo leer" in resumen(informe(error="roto"))


@pytest.mark.parametrize("valor, texto", [(25, "25%"), (25.0, "25%"), (24.5, "24.5%"), (None, "—")])
def test_los_porcentajes_se_escriben_sin_decimales_de_mas(valor, texto):
    assert pct(valor) == texto
