"""
Número de línea en la hoja de revisión: "Línea 33 y otros: corregir la forma de
citar". El revisor observa así, con los números que el tesista ve en el margen.
"""
from reportes.reporte_resumen import redactar
from utilidades.localizador import MapaLineas, _apretar


def _mapa(textos):
    lineas = [(1, n, _apretar(t), t) for n, t in enumerate(textos, 1)]
    return MapaLineas(lineas, 1, True, "numeración de líneas del documento", "Word")


def test_un_fragmento_que_cruza_el_salto_de_linea_se_ubica():
    """el comienzo del fragmento queda al final de una línea y el resto en la siguiente"""
    mapa = _mapa(["La referencia hidrológica externa procede de un marco calibrado (Llau-",
                  "ca et al., 2023) y se evalúa sobre 2022-2025, fuera de su ventana"])
    assert mapa.numero("...ca et al., 2023)  y se evalúa sobr") == 2


def _item(categoria, severidad, observacion, nlineas, ubicacion="Documento"):
    nlineas = [nlineas] if isinstance(nlineas, int) else (nlineas or [])
    return dict(categoria=categoria, severidad=severidad, observacion=observacion, ubicacion=ubicacion,
                nlinea=nlineas[0] if nlineas else None, nlineas=nlineas, detalle="")


def test_la_hoja_lleva_todas_las_lineas_de_cada_tipo_de_observacion():
    items = [
        _item("Formato", "Error", "Texto en fuente 'Courier New' (4.2% del documento)", [220, 407]),
        _item("Estructura", "Advertencia", "Título de sección incompleto: 'Uso de los resultados'", 445,
              "Uso de los resultados y contribuciones del proyecto"),
        _item("Citas", "Advertencia", "Referencia cortada ('…')", 363, "Referencias: 'Béjar'"),
        _item("Citas", "Advertencia", "Título en mayúsculas", 387, "Referencias: 'Llanque'"),
        _item("Ortografía", "Advertencia", "Doble espacio entre palabras (2 caso(s))", [295, 633]),
    ]
    filas = [dict(nlinea=11, nlineas=[11], palabra="hidro"), dict(nlinea=126, nlineas=[126], palabra=".Lees")]
    hoja = redactar(items, filas)
    assert "Líneas 220 y 407: corregir el formato del documento: tipo de letra" in hoja
    assert ("Línea 445: completar el título de la sección: debe decir "
            "«Uso de los resultados y contribuciones del proyecto»") in hoja
    assert "Líneas 363 y 387: referenciar según norma APA 7ª ed" in hoja
    assert "Líneas 11, 126, 295 y 633: corregir ortografía y gramática" in hoja


def test_sin_numero_de_linea_la_observacion_sale_igual():
    hoja = redactar([_item("Formato", "Error", "Margen izquierdo incorrecto", None)], [])
    assert "Corregir el formato del documento: márgenes" in hoja


def test_un_final_repetido_no_manda_a_otra_parte_del_documento():
    """
    'mlforecast' salía en la línea 43: su fragmento terminaba igual que una frase de la
    Justificación y se buscaba por el final. Buscando sobre el texto unido, no pasa.
    """
    mapa = _mapa([
        "Los modelos estadísticos y de aprendizaje automático ofrecen otra vía.",
        "XIV. Recursos",
        "INFRAESTRUCTURA Y EQUIPOS",
        "Python con las bibliotecas statsforecast y mlforecast para los",
        "modelos estadísticos y de aprendizaje automático, neuralforecast y",
    ])
    ancla = "Recursos INFRAESTRUCTURA Y EQUIPOS"          # título corto + primer contenido
    assert mapa.numero("mlforecast para los modelos estadísticos y de aprendizaje automático", ancla) == 4


def test_el_ancla_de_varias_lineas_se_encuentra():
    mapa = _mapa(["recursos hídricos del país", "XIV. Recursos", "INFRAESTRUCTURA Y EQUIPOS", "Computadora"])
    assert mapa._indice_ancla("Recursos INFRAESTRUCTURA Y EQUIPOS") == 1
