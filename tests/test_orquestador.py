"""
Tests del armado de la corrida: qué archivos entran, cómo se emparejan con su informe
de Turnitin, cómo se nombran las hojas y qué se rechaza antes de empezar.

Son los que faltaban cuando el emparejado con Turnitin quedó escrito en el orquestador
pero la interfaz seguía llamando a revisar() por su cuenta y nunca lo usaba.
"""
import os
import shutil

import pytest

from core.orquestador import (_base_reporte, ejecutar, emparejar, expandir,
                              leer_informes, validar_reglas)
from core.turnitin import Informe, nombre_comparable


# --------------------------------------------------------------------------- expandir

def test_expandir_separa_documentos_de_informes(tmp_path):
    (tmp_path / "tesis.docx").write_bytes(b"x")
    (tmp_path / "turnitin.pdf").write_bytes(b"x")
    (tmp_path / "viejo.doc").write_bytes(b"x")
    (tmp_path / "~$tesis.docx").write_bytes(b"x")
    documentos, pdfs, avisos = expandir([str(tmp_path)])
    assert [os.path.basename(d) for d in documentos] == ["tesis.docx"]
    assert [os.path.basename(p) for p in pdfs] == ["turnitin.pdf"]
    assert any(".doc" in a for a in avisos), "hay que pedir el .docx cuando llega un .doc"


def test_expandir_avisa_de_lo_que_no_revisa(tmp_path):
    hoja = tmp_path / "notas.txt"
    hoja.write_text("x", encoding="utf-8")
    documentos, pdfs, avisos = expandir([str(hoja)])
    assert documentos == [] and pdfs == []
    assert any("notas.txt" in a for a in avisos)


def test_expandir_avisa_si_no_existe():
    _, _, avisos = expandir(["no_existe_por_ningun_lado.docx"])
    assert any("No existe" in a for a in avisos)


# --------------------------------------------------------------------------- emparejar

def _informe(ruta, **extra):
    return Informe(ruta=ruta, archivo=os.path.basename(ruta), es_turnitin=True, **extra)


def test_empareja_por_nombre_parecido():
    docs = ["C:/x/PROYECTO Quispe Mamani.docx", "C:/x/PROYECTO Cruz Apaza.docx"]
    informes = [_informe("C:/x/Turnitin - Cruz Apaza.pdf")]
    parejas, sueltos = emparejar(docs, informes)
    assert parejas == {"C:/x/PROYECTO Cruz Apaza.docx": informes[0]}
    assert sueltos == []


def test_un_informe_con_nombre_generico_va_con_el_unico_documento():
    docs = ["C:/x/PROYECTO Quispe.docx"]
    informes = [_informe("C:/x/Informe Turnitin.pdf")]
    parejas, sueltos = emparejar(docs, informes)
    assert parejas[docs[0]] is informes[0]
    assert sueltos == []


def test_con_varios_documentos_un_nombre_generico_no_se_adivina():
    """
    Adivinar acá sería peor que no emparejar: pondría el índice de similitud de un
    tesista en la hoja de otro.
    """
    docs = ["C:/x/A Quispe.docx", "C:/x/B Mamani.docx"]
    informes = [_informe("C:/x/Informe Turnitin.pdf"), _informe("C:/x/Informe Turnitin (1).pdf")]
    parejas, sueltos = emparejar(docs, informes)
    assert parejas == {}
    assert len(sueltos) == 2, "los dos quedan sueltos y llevan su propia hoja"


def test_un_informe_no_se_reparte_entre_dos_documentos():
    docs = ["C:/x/Quispe Mamani.docx", "C:/x/Quispe Mamani (v2).docx"]
    informes = [_informe("C:/x/Turnitin Quispe Mamani.pdf")]
    parejas, _ = emparejar(docs, informes)
    assert len(parejas) == 1, "el mismo informe no puede ir en dos hojas"


def test_empareja_por_el_documento_que_declara_el_informe():
    """el PDF puede llamarse cualquier cosa y declarar adentro el archivo que se entregó"""
    docs = ["C:/x/tesis final Chambi.docx"]
    informes = [_informe("C:/x/descarga_98213.pdf", archivo_entregado="tesis final Chambi.docx")]
    parejas, sueltos = emparejar(docs, informes)
    assert parejas[docs[0]] is informes[0]


@pytest.mark.parametrize("docx, pdf", [
    ("PROYECTO Cruz Apaza.docx", "Turnitin - Cruz Apaza.pdf"),
    ("tesis Quispe Mamani.docx", "Quispe Mamani - informe de similitud.pdf"),
    ("PROYECTO DE TESIS - Chambi Luque.docx", "Chambi Luque.pdf"),
    ("Mamani Coila Jose.docx", "Recibo digital Mamani Coila Jose.pdf"),
])
def test_empareja_aunque_el_pdf_no_repita_el_rotulo_del_docx(docx, pdf):
    """
    El caso normal: el .docx se llama 'PROYECTO Apellido' y el PDF solo 'Apellido'.
    Comparando letra por letra esto daba 69 y no emparejaba; hay que comparar por
    palabras (ver _parecido).
    """
    informes = [_informe("C:/x/" + pdf)]
    parejas, sueltos = emparejar(["C:/x/" + docx], informes)
    assert parejas == {"C:/x/" + docx: informes[0]}, "tendría que haber emparejado"
    assert sueltos == []


@pytest.mark.parametrize("docx, pdf", [
    ("PROYECTO Cruz Apaza.docx", "Turnitin - Quispe Mamani.pdf"),
    ("PROYECTO Cruz Apaza.docx", "PROYECTO Quispe Mamani.pdf"),
    ("Cruz Apaza.docx", "Turnitin Cruz Condori.pdf"),
    # un apellido solo no alcanza, aunque un nombre contenga al otro: comparar por
    # palabras daba 100 acá y son dos tesistas distintos
    ("PROYECTO Quispe Mamani.docx", "Turnitin - Quispe.pdf"),
    ("Quispe.docx", "Turnitin Quispe Mamani Jose Luis.pdf"),
])
def test_no_empareja_documentos_de_tesistas_distintos(docx, pdf):
    """el falso positivo acá es grave: el porcentaje de uno en la hoja de otro"""
    informes = [_informe("C:/x/" + pdf)]
    parejas, sueltos = emparejar(["C:/x/" + docx], informes)
    assert parejas == {}, "no debería emparejar apellidos distintos"
    assert len(sueltos) == 1


def test_nombre_comparable_quita_las_palabras_de_rotulo():
    assert nombre_comparable("Informe Turnitin.pdf") == ""
    assert nombre_comparable("PROYECTO_Quispe_Mamani.docx") == "proyecto quispe mamani"
    # las preposiciones sueltas quedan: no estorban, porque el parecido se mide por
    # palabras y el apellido sigue pesando
    assert nombre_comparable("Reporte de similitud - Quispe.pdf") == "de quispe"


# --------------------------------------------------------------------------- nombres de salida

def test_dos_documentos_con_el_mismo_nombre_no_se_sobrescriben():
    usados = set()
    assert _base_reporte("Observaciones - ", "tesis", usados) == "Observaciones - tesis"
    assert _base_reporte("Observaciones - ", "tesis", usados) == "Observaciones - tesis (2)"
    assert _base_reporte("Observaciones - ", "tesis", usados) == "Observaciones - tesis (3)"


def test_sin_registro_de_usados_el_nombre_es_el_de_siempre():
    """revisar el mismo documento otra vez sí pisa su hoja anterior: eso es lo que se espera"""
    assert _base_reporte("Observaciones - ", "tesis") == "Observaciones - tesis"
    assert _base_reporte("Observaciones - ", "tesis") == "Observaciones - tesis"


def test_las_hojas_de_similitud_y_de_observaciones_no_chocan():
    usados = set()
    assert _base_reporte("Observaciones - ", "tesis", usados) == "Observaciones - tesis"
    assert _base_reporte("Similitud - ", "tesis", usados) == "Similitud - tesis"


# --------------------------------------------------------------------------- leer_informes

def test_un_pdf_roto_se_avisa_y_no_rompe_la_corrida(tmp_path):
    roto = tmp_path / "anexo.pdf"
    roto.write_bytes(b"%PDF-1.4\nesto no es un PDF\n")
    avisos = []
    assert leer_informes([str(roto)], avisos.append) == []
    assert any("anexo.pdf" in a for a in avisos)


# --------------------------------------------------------------------------- ejecutar

def test_ejecutar_rechaza_una_entrada_sin_nada_revisable(tmp_path, opciones):
    with pytest.raises(ValueError, match="No hay archivos"):
        ejecutar([str(tmp_path)], opciones, aviso=lambda m: None)


def test_ejecutar_rechaza_datos_de_expediente_con_varios_documentos(ejemplos, opciones):
    opciones.tesista = "Quispe Mamani"
    with pytest.raises(ValueError, match="registro.csv"):
        ejecutar([os.path.join(ejemplos, "proyecto_con_errores.docx"),
                  os.path.join(ejemplos, "proyecto_correcto.docx")], opciones, aviso=lambda m: None)


def test_ejecutar_rechaza_un_informe_indicado_que_no_existe(ejemplos, opciones):
    opciones.turnitin = "no_existe.pdf"
    with pytest.raises(ValueError, match="No existe el informe"):
        ejecutar([os.path.join(ejemplos, "proyecto_con_errores.docx")], opciones,
                 aviso=lambda m: None)


def test_un_informe_indicado_que_no_se_puede_leer_es_error_no_silencio(ejemplos, opciones, tmp_path):
    """
    El revisor nombró ese PDF a mano. Si no se puede leer hay que decirlo, no seguir
    como si no lo hubiera pedido ni salir a buscar otro a la carpeta.
    """
    roto = tmp_path / "Turnitin.pdf"
    roto.write_bytes(b"%PDF-1.4\nroto\n")
    opciones.turnitin = str(roto)
    with pytest.raises(ValueError, match="No se pudo usar el informe"):
        ejecutar([os.path.join(ejemplos, "proyecto_con_errores.docx")], opciones,
                 aviso=lambda m: None)


def test_dos_documentos_homonimos_generan_dos_hojas(ejemplos, opciones, tmp_path):
    a, b = tmp_path / "A", tmp_path / "B"
    a.mkdir(), b.mkdir()
    shutil.copy(os.path.join(ejemplos, "proyecto_con_errores.docx"), a / "tesis.docx")
    shutil.copy(os.path.join(ejemplos, "proyecto_correcto.docx"), b / "tesis.docx")
    hechos = ejecutar([str(a / "tesis.docx"), str(b / "tesis.docx")], opciones,
                      aviso=lambda m: None)
    assert len(hechos) == 2
    assert len(set(hechos)) == 2, "las dos hojas no pueden ser el mismo archivo"


def test_el_progreso_informa_el_total_y_avanza(ejemplos, opciones):
    """
    La interfaz dibuja su barra con esto. El total lo decide ejecutar() porque hasta que
    no se emparejan los PDF con los .docx no se sabe cuántas hojas van a salir.
    """
    eventos = []
    ejecutar([os.path.join(ejemplos, "proyecto_con_errores.docx"),
              os.path.join(ejemplos, "proyecto_correcto.docx")], opciones,
             aviso=lambda m: None, progreso=lambda h, t, n: eventos.append((h, t, n)))
    assert [e[1] for e in eventos] == [2, 2, 2], "el total no debería cambiar a mitad de camino"
    assert [e[0] for e in eventos] == [0, 1, 2], "el avance tiene que ir subiendo"
    assert eventos[-1][2] == "", "la última llamada cierra la barra, sin nombre"


def test_el_progreso_avanza_aunque_un_documento_falle(tmp_path, opciones):
    """si la barra solo contara los reportes que salieron, se quedaría clavada"""
    malo = tmp_path / "roto.docx"
    malo.write_bytes(b"no soy un docx")
    eventos = []
    hechos = ejecutar([str(malo)], opciones, aviso=lambda m: None,
                      progreso=lambda h, t, n: eventos.append((h, t, n)))
    assert hechos == []
    assert eventos[-1][0] == 1, "el documento fallado igual cuenta como avance"


# --------------------------------------------------------------------------- validar_reglas

def _fallos_de_similitud(similitud):
    """
    Los fallos que validar_reglas achaca al bloque 'similitud'. Se filtra porque estos
    tests le pasan un bloque suelto, no unas reglas completas: los avisos por las
    secciones que faltan son de otra cosa.
    """
    fallos = validar_reglas({"similitud": similitud})
    return [f for f in fallos if "similitud" in f.lower()]


def test_validar_reglas_acepta_el_bloque_de_similitud():
    assert _fallos_de_similitud({"max_pct": 25, "max_ia_pct": None, "exigir_informe": True,
                                 "filtros_exigidos": ["bibliografia", "citas"]}) == []


@pytest.mark.parametrize("similitud, esperado", [
    ({"max_pct": "veinte"}, "número"),
    ({"exigir_informe": "si"}, "true o false"),
    ({"filtros_exigidos": ["resumen"]}, "filtros_exigidos"),
    ({"filtros_exigidos": "bibliografia"}, "filtros_exigidos"),
    ({"tope_maximo": 25}, "sobra"),
])
def test_validar_reglas_rechaza_similitud_mal_escrita(similitud, esperado):
    fallos = _fallos_de_similitud(similitud)
    assert any(esperado in f for f in fallos), fallos


def test_validar_reglas_rechaza_similitud_que_no_es_un_bloque():
    fallos = _fallos_de_similitud(25)
    assert any("bloque" in f for f in fallos), fallos


def test_un_docx_con_un_pdf_sin_pareja_no_aborta_la_corrida(ejemplos, opciones, tmp_path, monkeypatch):
    """
    La app deja escribir tesista y asesor cuando hay un solo .docx, pero no puede saber
    de antemano si el PDF que lo acompaña se va a emparejar. Antes, si no se emparejaba,
    ejecutar() rechazaba toda la corrida y el revisor se quedaba sin ninguna hoja.
    """
    from core import orquestador
    ajeno = tmp_path / "Informe Quispe Mamani.pdf"
    ajeno.write_bytes(b"x")
    monkeypatch.setattr(orquestador, "leer_informes",
                        lambda pdfs, aviso=print: [Informe(ruta=p, archivo=os.path.basename(p),
                                                           es_turnitin=True) for p in pdfs])
    recibidas = []
    monkeypatch.setattr(orquestador, "revisar_informe",
                        lambda informe, reglas, salida, registro, args, aviso, usados:
                        recibidas.append(args) or "Similitud.docx")
    opciones.tesista = "Cruz Apaza"
    avisos = []
    hechos = ejecutar([os.path.join(ejemplos, "proyecto_con_errores.docx"), str(ajeno)],
                      opciones, avisos.append)

    assert len(hechos) == 2, "tienen que salir la hoja del .docx y la del informe"
    assert recibidas[0].tesista == "", "los datos del .docx no se le aplican al informe ajeno"
    assert opciones.tesista == "Cruz Apaza", "las opciones del revisor no se tocan"
    assert any("no corresponde" in a for a in avisos)


# --------------------------------------------------------------------------- fecha de presentación

def test_la_fecha_utc_de_word_se_pasa_a_hora_local():
    """
    Word guarda la fecha en UTC: lo guardado el 26/09 a las 20:00 en Perú queda como
    27/09 01:00. Tomar solo la fecha adelantaba un día; se convierte a la zona del equipo.
    """
    import datetime as dt
    from core.orquestador import fecha_local
    esperado = dt.datetime(2026, 9, 27, 1, tzinfo=dt.timezone.utc).astimezone().strftime("%d/%m/%Y")
    assert fecha_local("2026-09-27T01:00:00Z") == esperado
    lima = dt.timezone(dt.timedelta(hours=-5))
    assert dt.datetime(2026, 9, 27, 1, tzinfo=dt.timezone.utc).astimezone(lima).day == 26


def test_una_fecha_sin_hora_queda_igual():
    from core.orquestador import fecha_local
    assert fecha_local("2026-09-27") == "27/09/2026"
