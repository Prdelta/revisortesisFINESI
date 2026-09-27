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


def _item(categoria, severidad, observacion, nlinea, ubicacion="Documento"):
    return dict(categoria=categoria, severidad=severidad, observacion=observacion,
                ubicacion=ubicacion, nlinea=nlinea, detalle="")


def test_la_hoja_lleva_la_linea_de_cada_tipo_de_observacion():
    items = [
        _item("Formato", "Error", "Texto en fuente 'Courier New' (4.2% del documento)", 220),
        _item("Estructura", "Advertencia", "Título de sección incompleto: 'Uso de los resultados'", 445,
              "Uso de los resultados y contribuciones del proyecto"),
        _item("Citas", "Advertencia", "Referencia cortada ('…')", 363, "Referencias: 'Béjar'"),
        _item("Citas", "Advertencia", "Título en mayúsculas", 387, "Referencias: 'Llanque'"),
        _item("Ortografía", "Advertencia", "Doble espacio entre palabras (1 caso(s))", 295),
    ]
    filas = [dict(nlinea=11, palabra="hidro")]
    hoja = redactar(items, filas)
    assert "Línea 220 y otros: corregir el formato del documento: tipo de letra" in hoja
    assert ("Línea 445: completar el título de la sección: debe decir "
            "«Uso de los resultados y contribuciones del proyecto»") in hoja
    assert "Línea 363 y otros: referenciar según norma APA 7ª ed" in hoja
    assert "Línea 11 y otros: corregir ortografía y gramática" in hoja


def test_sin_numero_de_linea_la_observacion_sale_igual():
    hoja = redactar([_item("Formato", "Error", "Margen izquierdo incorrecto", None)], [])
    assert "Corregir el formato del documento: márgenes" in hoja
