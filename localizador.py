"""
Localiza en qué página y línea del documento está cada observación.

Convierte el .docx a PDF con LibreOffice y lee el texto con posiciones. Si la plantilla
tiene activada la numeración de líneas de Word (la del proyecto la tiene), se usan esos
mismos números, que son los que el tesista ve en el margen de su documento.
Si no la tiene, se cuentan las líneas desde el inicio de cada página.
"""
import os
import re
import subprocess
import tempfile
import unicodedata


def _norm(t):
    t = unicodedata.normalize("NFD", t.lower())
    t = "".join(c for c in t if unicodedata.category(c) != "Mn")
    return re.sub(r"[^a-z0-9ñ ]+", " ", t).strip()


def _apretar(t):
    return re.sub(r"\s+", "", _norm(t))


class MapaLineas:
    """lineas = [(pagina, numero_de_linea, texto_apretado, texto)]"""

    def __init__(self, lineas, paginas, numeradas, metodo):
        self.lineas = lineas
        self.paginas = paginas
        self.numeradas = numeradas
        self.metodo = metodo
        # numeracion continua = los numeros siguen creciendo de una pagina a otra
        nums = [n for _, n, _, _ in lineas if n is not None]
        crecientes = sum(1 for a, b in zip(nums, nums[1:]) if b >= a)
        self.continua = bool(numeradas) and len(nums) > 3 and crecientes >= len(nums) - 2

    def __bool__(self):
        return bool(self.lineas)

    def etiqueta(self, pagina, numero):
        if self.continua:
            return f"línea {numero}"
        if self.numeradas:
            return f"línea {numero}" if self.paginas == 1 else f"pág. {pagina}, línea {numero}"
        return f"pág. {pagina}, línea {numero} (contada desde el inicio de la página)"

    def _indice_ancla(self, ancla):
        """posicion en la lista de lineas donde empieza la seccion, para no confundir textos repetidos"""
        if not ancla:
            return 0
        clave = _apretar(ancla)[:40]
        if len(clave) < 6:
            return 0
        for k, (_, numero, apretado, _t) in enumerate(self.lineas):
            if numero is not None and clave[:30] in apretado:
                return k
        return 0

    def _posiciones(self, texto, desde=0):
        clave = _apretar(texto)
        if len(clave) < 8:
            return []
        for trozo in (clave[:40], clave[:22], clave[:12]):
            if len(trozo) < 8:
                break
            hallados = [k for k in range(desde, len(self.lineas))
                        if self.lineas[k][1] is not None and trozo in self.lineas[k][2]]
            if hallados:
                return hallados
        # el texto puede abarcar varias lineas: se busca una linea contenida en el texto
        return [k for k in range(desde, len(self.lineas))
                if self.lineas[k][1] is not None and len(self.lineas[k][2]) > 20 and self.lineas[k][2] in clave]

    def numero(self, texto, ancla=None):
        """solo el numero de linea (int) o None"""
        desde = self._indice_ancla(ancla)
        hallados = self._posiciones(texto, desde) or (self._posiciones(texto, 0) if desde else [])
        return self.lineas[hallados[0]][1] if hallados else None

    def buscar(self, texto, ancla=None, maximo=1):
        """etiqueta de la linea donde aparece el texto. 'ancla' es un texto anterior
        (normalmente el titulo de la seccion) desde el cual empezar a buscar."""
        desde = self._indice_ancla(ancla)
        hallados = self._posiciones(texto, desde)
        if not hallados and desde:
            hallados = self._posiciones(texto, 0)
        if not hallados:
            return ""
        return "; ".join(self.etiqueta(self.lineas[k][0], self.lineas[k][1]) for k in hallados[:maximo])


def _buscar_soffice():
    import shutil
    for nombre in ("soffice", "libreoffice"):
        if shutil.which(nombre):
            return shutil.which(nombre)
    for ruta in (r"C:\Program Files\LibreOffice\program\soffice.exe",
                 r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
                 "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        if os.path.exists(ruta):
            return ruta
    return None


RE_NUMERO = re.compile(r"^\s*(\d{1,4})\s{2,}(\S.*)$")


def construir(ruta_docx):
    """devuelve MapaLineas; vacio si no hay LibreOffice o falla la conversion"""
    soffice = _buscar_soffice()
    if not soffice:
        return MapaLineas([], None, False, "sin LibreOffice")
    try:
        from pypdf import PdfReader
    except ImportError:
        return MapaLineas([], None, False, "falta pypdf")
    try:
        with tempfile.TemporaryDirectory() as tmp:
            subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", tmp, ruta_docx],
                           capture_output=True, timeout=240)
            pdf = os.path.join(tmp, os.path.splitext(os.path.basename(ruta_docx))[0] + ".pdf")
            if not os.path.exists(pdf):
                return MapaLineas([], None, False, "LibreOffice no generó el PDF")
            lector = PdfReader(pdf)
            paginas = len(lector.pages)
            crudo = []
            for i, pagina in enumerate(lector.pages, 1):
                try:
                    texto = pagina.extract_text(extraction_mode="layout")
                except Exception:
                    texto = pagina.extract_text() or ""
                crudo.append((i, texto.split("\n")))
    except Exception:
        return MapaLineas([], None, False, "error al convertir")

    # ¿el documento trae numeración de líneas de Word en el margen?
    con_numero = sum(1 for _, ls in crudo for l in ls if RE_NUMERO.match(l))
    con_texto = sum(1 for _, ls in crudo for l in ls if l.strip())
    numeradas = con_texto and con_numero / con_texto > 0.5

    # encabezado y pie se repiten en cada pagina y LibreOffice tambien los numera: se descartan
    repetidos = set()
    if len(crudo) > 1:
        from collections import Counter
        bordes = Counter()
        for _, ls in crudo:
            utiles = [x.strip() for x in ls if x.strip()]
            for x in utiles[:3] + utiles[-3:]:
                bordes[_apretar(RE_NUMERO.sub(r"\2", x))] += 1
        repetidos = {k for k, n in bordes.items() if n >= len(crudo) and len(k) > 4}

    lineas = []
    for pagina, ls in crudo:
        contador = 0
        for l in ls:
            if numeradas:
                m = RE_NUMERO.match(l)
                if not m:
                    continue
                numero, texto = int(m.group(1)), m.group(2)
                if _apretar(texto) in repetidos:
                    continue
            else:
                if not l.strip():
                    continue
                if _apretar(l) in repetidos:
                    continue
                contador += 1
                numero, texto = contador, l.strip()
            lineas.append((pagina, numero, _apretar(texto), texto.strip()))
    lineas = _depurar(lineas)
    return MapaLineas(lineas, paginas, bool(numeradas),
                      "numeración de líneas del documento" if numeradas else "conteo de líneas por página")


def _depurar(lineas):
    """quita lineas sueltas cuya numeracion se sale de la secuencia (tablas, pies de pagina);
    si son demasiadas es que la numeracion reinicia por pagina y se deja todo como está"""
    limpias, fuera, tope = [], [], 0
    for fila in lineas:
        numero = fila[1]
        if numero is None or numero >= tope - 3:
            tope = max(tope, numero or 0)
            limpias.append(fila)
        else:
            fuera.append(fila)
    return limpias if len(fuera) <= len(lineas) * 0.2 else lineas


RE_ENTRECOMILLADO = re.compile(r"'([^']{6,})'")


def fragmento(ubicacion):
    """saca el texto citado dentro de una ubicacion tipo  Seccion: 'texto...'  """
    m = RE_ENTRECOMILLADO.search(ubicacion or "")
    if not m:
        return ""
    return m.group(1).replace("...", " ").replace("…", " ").strip()
