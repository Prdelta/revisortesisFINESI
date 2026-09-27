"""
Tests de la interfaz: que el botón «Revisar» pase por el mismo camino que la línea de
comandos, y que el PDF de Turnitin se pueda cargar desde la app.

Existen por una regresión concreta: el emparejado de cada .docx con su informe de
Turnitin se escribió en orquestador.ejecutar(), pero _trabajar() seguía llamando a
revisar() documento por documento y nunca le pasaba el informe. La función estaba
escrita, probada a mano por CLI, y en el .exe no hacía nada.

No se abre ninguna ventana: se prueban los métodos con un objeto mínimo que solo tiene
la cola de mensajes, que es todo lo que estos métodos usan de la ventana.
"""
import os
import queue
import types

import pytest

from core import orquestador

app = pytest.importorskip("app", reason="hace falta tkinter/ttkbootstrap")
App = app.App


@pytest.fixture
def ventana():
    """lo mínimo de la ventana que _trabajar necesita: la cola y el coloreado"""
    return types.SimpleNamespace(cola=queue.Queue(), _color=App._color)


def vaciar(cola):
    eventos = []
    while not cola.empty():
        eventos.append(cola.get())
    return eventos


# --------------------------------------------------------------------------- el camino compartido

def test_la_app_revisa_por_ejecutar_y_no_por_su_cuenta(ventana, opciones, monkeypatch):
    """
    Lo que esta regresión costó: si _trabajar vuelve a recorrer los documentos por su
    cuenta, se saltea el emparejado con Turnitin y la función no llega al .exe.
    """
    llamadas = []

    def espia(entradas, op, aviso=print, progreso=None):
        llamadas.append((list(entradas), op))
        return ["hoja.docx"]

    monkeypatch.setattr(orquestador, "ejecutar", espia)
    App._trabajar(ventana, ["a.docx", "Turnitin - a.pdf"], opciones)

    assert len(llamadas) == 1, "la app tiene que delegar en ejecutar() una sola vez"
    entradas, op = llamadas[0]
    assert entradas == ["a.docx", "Turnitin - a.pdf"], "el PDF tiene que llegar a ejecutar()"
    assert op is opciones


def test_el_pdf_llega_hasta_la_revision_de_similitud(ventana, opciones, ejemplos, monkeypatch):
    """
    El otro extremo de lo mismo: que el informe emparejado llegue como argumento a
    revisar(), que es donde se revisa la similitud.
    """
    recibidos = []
    original = orquestador.revisar

    def espia(*a, **kw):
        recibidos.append(kw.get("informe", a[7] if len(a) > 7 else None))
        return original(*a, **kw)

    monkeypatch.setattr(orquestador, "revisar", espia)
    App._trabajar(ventana, [os.path.join(ejemplos, "proyecto_con_errores.docx")], opciones)
    assert len(recibidos) == 1, "se tiene que revisar el documento"
    # sin PDF junto al .docx el informe es None, pero el argumento tiene que existir:
    # antes ni se pasaba y revisar() no podía observar la similitud nunca
    assert recibidos == [None]


# --------------------------------------------------------------------------- corrida real

def test_una_corrida_completa_deja_la_cola_en_orden(ventana, opciones, ejemplos):
    App._trabajar(ventana, [os.path.join(ejemplos, "proyecto_con_errores.docx")], opciones)
    eventos = vaciar(ventana.cola)
    clases = [c for c, _ in eventos]

    assert clases[-1] == "fin"
    hechos, salida = eventos[-1][1]
    assert hechos == 1
    assert salida == opciones.salida
    assert "total" in clases, "la barra necesita saber el total"
    assert "estado" in clases, "hay que decir qué documento se está revisando"
    avances = [v for c, v in eventos if c == "avance"]
    assert avances == sorted(avances), "el avance no puede ir para atrás"
    assert avances[-1] == 1


def test_un_error_esperado_se_muestra_sin_volcar_el_traceback(ventana, opciones, tmp_path):
    """
    Lo que ejecutar() rechaza a propósito es un mensaje para el revisor, no un fallo del
    programa: no tiene por qué ver un traceback.
    """
    App._trabajar(ventana, [str(tmp_path)], opciones)
    eventos = vaciar(ventana.cola)
    logs = [v[0] for c, v in eventos if c == "log"]
    assert any("No hay archivos" in m for m in logs), logs
    assert not any("Traceback" in m for m in logs), "un ValueError no lleva traceback"
    assert eventos[-1] == ("fin", (0, opciones.salida))


# --------------------------------------------------------------------------- selección de archivos

@pytest.mark.parametrize("archivos, esperado", [
    ([], 0),
    (["a.docx"], 1),
    (["a.docx", "Turnitin - a.pdf"], 1),                  # un expediente, no dos
    (["a.docx", "b.docx", "t1.pdf", "t2.pdf"], 2),
    (["solo1.pdf", "solo2.pdf"], 2),                      # informes sueltos: una hoja cada uno
])
def test_cuantas_hojas_se_esperan(archivos, esperado):
    """
    De esto depende que los campos de tesista y asesor queden habilitados: un .docx con
    su PDF es un solo expediente y el revisor tiene que poder escribir sus datos.
    """
    assert App._unidades(types.SimpleNamespace(archivos=archivos)) == esperado


@pytest.mark.parametrize("mensaje, etiqueta", [
    ("tesis.docx: 0 errores, 2 advertencias/revisar", "ok"),
    ("tesis.docx: 17 errores, 6 advertencias/revisar", None),
    ("tesis.docx: NO REVISADO. Parece que no es un borrador", "err"),
    ("anexo.pdf: no se pudo leer el PDF (roto).", "err"),
    ("anexo.pdf: no parece un informe de Turnitin, no se revisa.", "suave"),
    ("Se omite notas.txt: solo se revisan archivos .docx", "suave"),
])
def test_el_color_de_cada_linea_del_registro(mensaje, etiqueta):
    assert App._color(mensaje) == etiqueta


@pytest.mark.parametrize("texto, valida", [
    ("15/03/2026", True),
    ("1/3/2026", True),
    ("2026-03-15", False),
    ("15 de marzo", False),
    ("31/02/2026", False),
])
def test_la_fecha_de_presentacion_se_valida(texto, valida):
    assert App._fecha_valida(texto) is valida


# --------------------------------------------------------------------------- estado en la tabla

@pytest.mark.parametrize("mensaje, archivo, estado, etiqueta", [
    ("tesis.docx: 0 errores, 2 advertencias/revisar", "tesis.docx", "✓  0 errores · 2 advertencias", "ok"),
    ("tesis.docx: 17 errores, 6 advertencias/revisar", "tesis.docx", "✗  17 errores · 6 advertencias", "error"),
    ("tesis.docx: NO REVISADO. Parece que no es un borrador", "tesis.docx",
     "No revisado: no parece de este tipo", "error"),
    ("tesis.docx: no se pudo revisar (KeyError: 'x')", "tesis.docx", "No se pudo revisar", "error"),
    ("anexo.pdf: no se pudo leer el PDF (roto).", "anexo.pdf", "PDF ilegible", "error"),
    ("anexo.pdf: no parece un informe de Turnitin, no se revisa.", "anexo.pdf",
     "No es un informe de Turnitin", "pendiente"),
    ("tesis.docx: se usa el informe de Turnitin 'Turnitin - tesis.pdf'.", "Turnitin - tesis.pdf",
     "Emparejado con tesis.docx", "ok"),
    ("solo.pdf: informe de Turnitin, similitud 12% (internet 5%), 1 errores", "solo.pdf",
     "Hoja propia · similitud 12 %", "error"),
])
def test_cada_linea_del_registro_se_vuelve_un_estado(mensaje, archivo, estado, etiqueta):
    info = App.interpretar(mensaje)
    assert (info["archivo"], info["estado"], info["etiqueta"]) == (archivo, estado, etiqueta)


def test_los_totales_salen_de_la_linea_de_resultado():
    info = App.interpretar("tesis.docx: 3 errores, 4 advertencias/revisar, similitud 18.5% (internet 9%)")
    assert (info["errores"], info["advertencias"], info["similitud"], info["hoja"]) == (3, 4, 18.5, True)


@pytest.mark.parametrize("mensaje", [
    "Tipo: proyecto   ·   2 archivo(s)",
    "No hay archivos .docx ni informes de Turnitin en .pdf para revisar.",
    "Se omite notas.txt: solo se revisan archivos .docx",
])
def test_las_lineas_que_no_son_de_un_archivo_no_tocan_la_tabla(mensaje):
    assert App.interpretar(mensaje) is None


def test_el_orquestador_sigue_escribiendo_lo_que_la_tabla_entiende(opciones, ejemplos):
    """si cambia el texto del resultado en el orquestador, la tabla dejaría de actualizarse"""
    avisos = []
    orquestador.ejecutar([os.path.join(ejemplos, "proyecto_con_errores.docx")], opciones, avisos.append)
    resultados = [App.interpretar(a.split(" -> ")[0]) for a in avisos]
    assert any(r and r["hoja"] and r["errores"] > 0 for r in resultados), avisos


@pytest.mark.parametrize("falla", [RuntimeError("se rompió"), SystemExit(1)])
def test_la_revision_siempre_avisa_que_termino(ventana, opciones, monkeypatch, falla):
    """
    Si el hilo muere sin mandar «fin», la ventana se queda en «Revisando…» para siempre
    y en el .exe, sin consola, no hay forma de saber por qué.
    """
    def revienta(*a, **kw):
        raise falla

    monkeypatch.setattr(orquestador, "ejecutar", revienta)
    App._trabajar(ventana, ["a.docx"], opciones)
    eventos = vaciar(ventana.cola)
    assert eventos[-1] == ("fin", (0, opciones.salida))
