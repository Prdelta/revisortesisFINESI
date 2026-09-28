"""
Localiza en qué página y línea del documento está cada observación.

Convierte el .docx a PDF y lee el texto con posiciones. En Windows se usa Word, que
compagina igual que lo que ve el tesista; si no hay Word, LibreOffice. Si la plantilla
tiene activada la numeración de líneas de Word (la del proyecto la tiene), se usan esos
mismos números, que son los que el tesista ve en el margen de su documento.
Si no la tiene, se cuentan las líneas desde el inicio de cada página.
"""
import os
import re
import subprocess
import sys
import tempfile
import unicodedata
from bisect import bisect_right


def _norm(t):
    t = unicodedata.normalize("NFD", t.lower())
    t = "".join(c for c in t if unicodedata.category(c) != "Mn")
    return re.sub(r"[^a-z0-9ñ ]+", " ", t).strip()


def _apretar(t):
    return re.sub(r"\s+", "", _norm(t))


class MapaLineas:
    """lineas = [(pagina, numero_de_linea, texto_apretado, texto)]"""

    def __init__(self, lineas, paginas, numeradas, metodo, motor=""):
        self.motor = motor          # "Word" o "LibreOffice": con qué se generó el PDF
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

    def _texto_unido(self):
        """
        Todas las líneas pegadas en un solo texto, con el comienzo de cada una. Buscar
        ahí hace que un salto de línea del PDF no importe: antes, un fragmento partido
        entre dos líneas no se encontraba, y buscar trozos sueltos acertaba en otra frase
        igual de otra parte del documento.
        """
        if not hasattr(self, "_unido"):
            inicios, total = [], 0
            for _, _, apretado, _ in self.lineas:
                inicios.append(total)
                total += len(apretado)
            self._unido = "".join(apretado for _, _, apretado, _ in self.lineas)
            self._inicios = inicios
        return self._unido, self._inicios

    def _linea_en(self, posicion):
        _, inicios = self._texto_unido()
        return max(0, bisect_right(inicios, posicion) - 1)

    def _indice_ancla(self, ancla):
        """posicion en la lista de lineas donde empieza la seccion, para no confundir textos repetidos"""
        if not ancla:
            return 0
        clave = _apretar(ancla)[:30]
        if len(clave) < 6:
            return 0
        unido, _ = self._texto_unido()
        pos = unido.find(clave)
        return self._linea_en(pos) if pos >= 0 else 0

    def _posiciones(self, texto, desde=0):
        clave = _apretar(texto)
        if len(clave) < 8 or not self.lineas:
            return []
        unido, inicios = self._texto_unido()
        base = inicios[desde] if desde < len(inicios) else len(unido)
        for trozo in (clave[:40], clave[:22], clave[:12]):
            if len(trozo) < 8:
                break
            hallados, pos = [], unido.find(trozo, base)
            while pos >= 0:
                k = self._linea_en(pos)
                if not hallados or hallados[-1] != k:
                    hallados.append(k)
                pos = unido.find(trozo, pos + 1)
            if hallados:
                return hallados
        return []

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


def hay_word():
    """¿está instalado Microsoft Word? (su servidor COM está registrado)"""
    if not sys.platform.startswith("win"):
        return False
    try:
        import winreg
        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Word.Application\CLSID"))
        return True
    except OSError:
        return False


# Exporta a PDF con una instancia de Word propia, invisible, que se cierra al terminar:
# el Word que el revisor tenga abierto no se toca. El documento se abre en solo
# lectura y con las macros desactivadas. Las rutas llegan por variables de entorno
# para no tener que escapar comillas ni espacios.
SCRIPT_WORD = r"""
$ErrorActionPreference = 'Stop'
$w = New-Object -ComObject Word.Application
try {
    $w.Visible = $false
    $w.DisplayAlerts = 0
    $w.AutomationSecurity = 3
    $d = $w.Documents.Open($env:RT_DOCX, $false, $true, $false)
    try { $d.ExportAsFixedFormat($env:RT_PDF, 17) } finally { $d.Close($false) }
} finally {
    $w.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($w)
}
"""


def _pdf_con_word(ruta_docx, pdf):
    entorno = dict(os.environ, RT_DOCX=os.path.abspath(ruta_docx), RT_PDF=os.path.abspath(pdf))
    subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
                    "-Command", SCRIPT_WORD], env=entorno, capture_output=True, timeout=240,
                   creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    return os.path.exists(pdf)


def _pdf_con_libreoffice(soffice, ruta_docx, tmp):
    subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", tmp, ruta_docx],
                   capture_output=True, timeout=240,
                   creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    pdf = os.path.join(tmp, os.path.splitext(os.path.basename(ruta_docx))[0] + ".pdf")
    return pdf if os.path.exists(pdf) else None


RE_NUMERO = re.compile(r"^\s*(\d{1,4})\s{2,}(\S.*)$")


def construir(ruta_docx):
    """devuelve MapaLineas; vacio si no hay Word ni LibreOffice o falla la conversion"""
    soffice = _buscar_soffice()
    word = hay_word()
    if not word and not soffice:
        return MapaLineas([], None, False, "sin Word ni LibreOffice")
    try:
        from pypdf import PdfReader
    except ImportError:
        return MapaLineas([], None, False, "falta pypdf")
    try:
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as tmp:
            pdf, motor = None, ""
            if word:
                destino = os.path.join(tmp, "documento.pdf")
                try:
                    if _pdf_con_word(ruta_docx, destino):
                        pdf, motor = destino, "Word"
                except Exception:
                    pass   # Word falló (licencia, documento protegido): se intenta LibreOffice
            if pdf is None and soffice:
                pdf, motor = _pdf_con_libreoffice(soffice, ruta_docx, tmp), "LibreOffice"
            if pdf is None:
                return MapaLineas([], None, False, "no se pudo generar el PDF")
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
    con_numero = sum(1 for _, ls in crudo for x in ls if RE_NUMERO.match(x))
    con_texto = sum(1 for _, ls in crudo for x in ls if x.strip())
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
        for linea in ls:
            if numeradas:
                m = RE_NUMERO.match(linea)
                if not m:
                    continue
                numero, texto = int(m.group(1)), m.group(2)
                if _apretar(texto) in repetidos:
                    continue
            else:
                if not linea.strip():
                    continue
                if _apretar(linea) in repetidos:
                    continue
                contador += 1
                numero, texto = contador, linea.strip()
            lineas.append((pagina, numero, _apretar(texto), texto.strip()))
    lineas = _depurar(lineas)
    return MapaLineas(lineas, paginas, bool(numeradas),
                      "numeración de líneas del documento" if numeradas else "conteo de líneas por página",
                      motor)


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
    todos = fragmentos(ubicacion)
    return todos[0] if todos else ""


def fragmentos(ubicacion):
    """
    Todos los textos citados de una ubicación. Una observación que junta varios casos
    ("Doble espacio (3 casos)") trae un ejemplo por caso, y cada uno tiene su línea.
    """
    return [m.group(1).replace("...", " ").replace("…", " ").strip()
            for m in RE_ENTRECOMILLADO.finditer(ubicacion or "")]
