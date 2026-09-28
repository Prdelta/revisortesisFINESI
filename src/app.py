"""
Revisor de Tesis FINESI. Aplicación de escritorio.

Ejecutar:  python app.py
"""
import os
import queue
import re
import shutil
import subprocess
import sys
import threading
import time
from datetime import datetime
import tkinter as tk
import ttkbootstrap as tb
from tkinter import filedialog, messagebox

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import core          # noqa: F401  (import explicito para PyInstaller)
import utilidades    # noqa: F401
import reportes      # noqa: F401

import core.orquestador as orquestador
from core.config import DATOS, tipos_disponibles, ruta_recurso
from core.app_config import AppConfig
from utilidades import reglas_revisor
from core.orquestador import preparar_datos, Opciones  # noqa: E402
from reportes.reporte_resumen import crear_plantilla_hoja  # noqa: E402

try:
    # arrastrar y soltar archivos; sin esta librería la app funciona igual, con los botones
    from tkinterdnd2 import DND_FILES, TkinterDnD
except ImportError:
    DND_FILES = TkinterDnD = None

EN_WINDOWS = sys.platform.startswith("win")
FUENTE = "Segoe UI" if EN_WINDOWS else "DejaVu Sans"
MONO = "Consolas" if EN_WINDOWS else "monospace"


def _fuerte(tamano):
    return ("Segoe UI Semibold", tamano) if EN_WINDOWS else (FUENTE, tamano, "bold")


def _tema(estilo, oscuro):
    """los temas de ttkbootstrap 2.x; con una versión anterior, sus equivalentes"""
    disponibles = estilo.theme_names()
    nombre = "bootstrap-dark" if oscuro else "bootstrap-light"
    return nombre if nombre in disponibles else ("darkly" if oscuro else "flatly")


# Cómo leer las líneas que emite el orquestador para mostrar el estado de cada archivo.
_ARCHIVO = re.compile(r"^(.+?\.(?:docx|pdf)): (.*)$", re.IGNORECASE | re.DOTALL)
_RESULTADO = re.compile(r"^(\d+) errores, (\d+) advertencias/revisar")
_SIMILITUD = re.compile(r"similitud (\d+(?:\.\d+)?)%")
_ERRORES_FINAL = re.compile(r", (\d+) errores\s*$")
_PAREJA = re.compile(r"^se usa el informe de Turnitin '(.+)'")


def _pct(valor):
    return f"{valor:g} %"


class App(tb.Window):
    def __init__(self):
        self.cfg = AppConfig(os.path.join(DATOS, "config.json")).leer_config()
        self.oscuro = self.cfg.get("tema", "oscuro") != "claro"
        super().__init__(themename="darkly")
        self.style.theme_use(_tema(self.style, self.oscuro))
        preparar_datos()
        self.config_manager = AppConfig(os.path.join(DATOS, "config.json"))
        self.title("Revisor de Tesis  ·  FINESI")
        self.geometry("1180x780")
        self.minsize(1040, 700)
        self.arrastre = False
        if TkinterDnD is not None:
            try:
                TkinterDnD._require(self)
                self.arrastre = True
            except Exception:
                pass
        self._icono()
        self.archivos = []
        self.filas = {}          # nombre del archivo -> filas de la tabla
        self.parejas = {}        # .docx -> nombre de su informe de Turnitin
        self.cola = queue.Queue()
        self.trabajando = False
        self._hubo_error = False
        self._totales = {}
        self._estado_base = ""
        self._inicio = 0.0

        self._construir()
        self.protocol("WM_DELETE_WINDOW", self._cerrar)
        self.after(120, self._vaciar_cola)

    # ---------------------------------------------------------------- preferencias
    def _guardar_config(self):
        datos = dict(revisor=self.campos["revisor"][0].get().strip(),
                     n_revision=self.campos["n_revision"][0].get().strip(),
                     extras=self.extras.get("1.0", "end").strip(),
                     formato=self.formato.get(),
                     tipo=self.tipo.get(), salida=self.salida.get().strip(),
                     excel=bool(self.excel.get()), gramatica=bool(self.gramatica.get()),
                     abrir=bool(self.abrir.get()), tema="oscuro" if self.oscuro else "claro")
        self.config_manager.guardar_config(datos)

    def _cerrar(self):
        # el hilo de la revisión es daemon: cerrar lo corta y puede dejar una hoja a medias
        if self.trabajando and not messagebox.askyesno(
                "Revisor de Tesis",
                "Hay una revisión en curso. Si cierras ahora se interrumpe y el reporte "
                "que se está escribiendo puede quedar incompleto.\n\n¿Cerrar de todos modos?"):
            return
        self._guardar_config()
        self.destroy()

    @staticmethod
    def _recurso(nombre):
        """en el .exe el ícono va entre los recursos; al correr con python, en la raíz del repo"""
        for ruta in (ruta_recurso(nombre),
                     os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), nombre)):
            if os.path.exists(ruta):
                return ruta
        return ""

    def _icono(self):
        self.logo = None
        png = self._recurso("icono.png")
        if png:
            try:
                from PIL import Image, ImageTk
                self.logo = ImageTk.PhotoImage(Image.open(png).resize((44, 44), Image.LANCZOS))
            except Exception:
                self.logo = None
        ico = self._recurso("icono.ico")
        try:
            if EN_WINDOWS and ico:
                self.iconbitmap(ico)
            elif png:
                self.iconphoto(True, tk.PhotoImage(file=png))
        except Exception:
            pass

    # ---------------------------------------------------------------- interfaz
    def _construir(self):
        self._cabecera()
        self._barra_inferior()
        cuerpo = tb.Frame(self, padding=(24, 18, 24, 8))
        cuerpo.pack(fill="both", expand=True)
        cuerpo.columnconfigure(0, weight=3, uniform="c")
        cuerpo.columnconfigure(1, weight=2, uniform="c")
        cuerpo.rowconfigure(0, weight=1)
        izq = tb.Frame(cuerpo)
        izq.grid(row=0, column=0, sticky="nsew", padx=(0, 18))
        der = tb.Frame(cuerpo)
        der.grid(row=0, column=1, sticky="nsew")
        self._panel_documentos(izq)
        self._panel_ajustes(der)
        self._refrescar()
        self._recolorear()

    def _cabecera(self):
        tb.Frame(self, bootstyle="primary", height=4).pack(fill="x")
        cab = tb.Frame(self, padding=(24, 14, 24, 14))
        cab.pack(fill="x")
        if self.logo:
            tb.Label(cab, image=self.logo).pack(side="left", padx=(0, 14))
        textos = tb.Frame(cab)
        textos.pack(side="left")
        tb.Label(textos, text="Revisor de Tesis", font=_fuerte(18)).pack(anchor="w")
        tb.Label(textos, text="Universidad Nacional del Altiplano · Facultad de Ingeniería "
                              "Estadística e Informática", bootstyle="secondary",
                 font=(FUENTE, 9)).pack(anchor="w")
        self.boton_tema = tb.Button(cab, bootstyle="secondary-outline", command=self._cambiar_tema)
        self.boton_tema.pack(side="right")
        tb.Separator(self).pack(fill="x")

    # ---- columna izquierda: documentos y resultado
    def _panel_documentos(self, padre):
        cab = tb.Frame(padre)
        cab.pack(fill="x")
        tb.Label(cab, text="Documentos", font=_fuerte(12)).pack(side="left")
        self.conteo = tb.Label(cab, text="", bootstyle="secondary", font=(FUENTE, 9))
        self.conteo.pack(side="left", padx=10, pady=(3, 0))
        tb.Button(cab, text="Limpiar", bootstyle="link", command=self.limpiar).pack(side="right")
        tb.Button(cab, text="Quitar", bootstyle="link", command=self.quitar).pack(side="right")

        self._zona_arrastre(padre)

        marco = tb.Frame(padre)
        marco.pack(fill="both", expand=True, pady=(10, 0))
        self.tabla = tb.Treeview(marco, columns=("tipo", "estado"), show="tree headings",
                                 bootstyle="primary", selectmode="extended", height=4)
        self.tabla.heading("#0", text="Archivo", anchor="w")
        self.tabla.heading("tipo", text="Tipo", anchor="w")
        self.tabla.heading("estado", text="Estado", anchor="w")
        self.tabla.column("#0", width=250, stretch=True)
        self.tabla.column("tipo", width=80, stretch=False)
        self.tabla.column("estado", width=230, stretch=False)
        barra = tb.Scrollbar(marco, orient="vertical", command=self.tabla.yview)
        self.tabla.configure(yscrollcommand=barra.set)
        self.tabla.pack(side="left", fill="both", expand=True)
        barra.pack(side="right", fill="y")
        self.tabla.bind("<Delete>", lambda _: self.quitar())
        if self.arrastre:
            self.tabla.drop_target_register(DND_FILES)
            self.tabla.dnd_bind("<<Drop>>", self._soltar)

        self._panel_resultado(padre)

    def _zona_arrastre(self, padre):
        """recuadro punteado donde se sueltan los .docx y los PDF de Turnitin"""
        self.zona = tk.Canvas(padre, height=96, highlightthickness=0, bd=0)
        self.zona.pack(fill="x", pady=(10, 0))
        self.zona_botones = tb.Frame(self.zona)
        tb.Button(self.zona_botones, text="+  Agregar archivos", bootstyle="primary",
                  command=self.agregar_archivos).pack(side="left")
        tb.Button(self.zona_botones, text="Agregar carpeta", bootstyle="primary-outline",
                  command=self.agregar_carpeta).pack(side="left", padx=8)
        self.zona.bind("<Configure>", lambda _: self._dibujar_zona())
        if self.arrastre:
            self.zona.drop_target_register(DND_FILES)
            self.zona.dnd_bind("<<Drop>>", self._soltar)

    def _dibujar_zona(self):
        z = self.zona
        z.delete("all")
        ancho, alto = z.winfo_width(), z.winfo_height()
        gris = self.style.colors.secondary
        z.create_rectangle(2, 2, ancho - 2, alto - 2, outline=gris, dash=(6, 4))
        texto = ("Arrastra aquí los .docx y los informes de Turnitin (.pdf)" if self.arrastre
                 else "Agrega los .docx y, si los tienes, los informes de Turnitin (.pdf)")
        z.create_text(ancho / 2, 26, text=texto, fill=gris, font=(FUENTE, 10))
        z.create_window(ancho / 2, 64, window=self.zona_botones)

    def _panel_resultado(self, padre):
        tb.Label(padre, text="Resultado de la última revisión", font=_fuerte(12)).pack(
            anchor="w", pady=(18, 8))
        fila = tb.Frame(padre)
        fila.pack(fill="x")
        self.tarjetas = {}
        for i, (clave, texto, estilo) in enumerate((("hojas", "hojas generadas", "primary"),
                                                    ("errores", "errores", "danger"),
                                                    ("advertencias", "advertencias", "warning"),
                                                    ("similitud", "similitud máxima", "info"))):
            fila.columnconfigure(i, weight=1, uniform="t")
            borde = tb.Frame(fila, bootstyle=estilo, padding=1)
            borde.grid(row=0, column=i, sticky="nsew", padx=(0 if i == 0 else 8, 0))
            interior = tb.Frame(borde, padding=(14, 10))
            interior.pack(fill="both", expand=True)
            cifra = tb.Label(interior, text="—", bootstyle=estilo, font=_fuerte(20))
            cifra.pack(anchor="w")
            tb.Label(interior, text=texto, bootstyle="secondary", font=(FUENTE, 9)).pack(anchor="w")
            self.tarjetas[clave] = (borde, cifra)
        self.tarjetas["similitud"][0].grid_remove()

        acciones = tb.Frame(padre)
        acciones.pack(fill="x", pady=(10, 0))
        tb.Button(acciones, text="Abrir carpeta de reportes", bootstyle="success",
                  command=lambda: self.abrir_carpeta(self._carpeta_salida())).pack(side="left")
        self.ver_registro = tk.BooleanVar(value=False)
        tb.Checkbutton(acciones, text="Ver registro técnico", variable=self.ver_registro,
                       bootstyle="secondary-round-toggle", command=self._mostrar_registro).pack(side="right")

        self.log = tk.Text(padre, height=7, font=(MONO, 9), bd=0, highlightthickness=1,
                           wrap="word", padx=10, pady=8, state="disabled")

    def _mostrar_registro(self):
        if self.ver_registro.get():
            self.log.pack(fill="both", expand=False, pady=(10, 0))
        else:
            self.log.pack_forget()

    # ---- columna derecha: pestañas
    def _panel_ajustes(self, padre):
        pestanas = tb.Notebook(padre)
        pestanas.pack(fill="both", expand=True)
        pestanas.add(self._pestana_revision(pestanas), text="  Revisión  ")
        pestanas.add(self._pestana_expediente(pestanas), text="  Expediente  ")
        pestanas.add(self._pestana_ajustes(pestanas), text="  Ajustes  ")

    @staticmethod
    def _titulo(padre, texto, arriba=16):
        tb.Label(padre, text=texto, font=_fuerte(10)).pack(anchor="w", pady=(arriba, 6))

    @staticmethod
    def _ajustar_al_ancho(padre, etiquetas):
        """el texto se corta al ancho de la pestaña, que cambia con la ventana"""
        def ajustar(evento):
            for etiqueta in etiquetas:
                etiqueta.configure(wraplength=max(evento.width - 40, 150))
        padre.bind("<Configure>", ajustar, add="+")

    def _pestana_revision(self, nb):
        p = tb.Frame(nb, padding=18)
        self._titulo(p, "Tipo de documento", arriba=0)
        self.tipos = tipos_disponibles() or [("proyecto", False)]
        nombres = [t for t, _ in self.tipos]
        guardado = self.cfg.get("tipo")
        self.tipo = tk.StringVar(value=guardado if guardado in nombres else
                                 ("proyecto" if "proyecto" in nombres else nombres[0]))
        selector = tb.Frame(p)
        selector.pack(fill="x")
        for t in nombres:
            tb.Radiobutton(selector, text=t.capitalize(), value=t, variable=self.tipo,
                           bootstyle="primary-outline-toolbutton", padding=(0, 7),
                           command=self.cambio_tipo).pack(side="left", fill="x", expand=True)
        self.aviso_tipo = tb.Label(p, text="", bootstyle="secondary", font=(FUENTE, 9), justify="left")
        self.aviso_tipo.pack(anchor="w", pady=(6, 0))

        self._titulo(p, "Formato del reporte")
        self.formato = tk.StringVar(value=self.cfg.get("formato", "resumen"))
        for valor, texto, ayuda in (("resumen", "Hoja de revisión", "Una página con las observaciones"),
                                    ("detallado", "Detallado", "Tabla completa, observación por observación"),
                                    ("ambos", "Los dos", "")):
            fila = tb.Frame(p)
            fila.pack(fill="x", pady=2)
            tb.Radiobutton(fila, text=texto, value=valor, variable=self.formato).pack(side="left")
            if ayuda:
                tb.Label(fila, text=ayuda, bootstyle="secondary", font=(FUENTE, 8)).pack(side="left", padx=8)

        self._titulo(p, "Opciones")
        self.excel = tk.BooleanVar(value=bool(self.cfg.get("excel", False)))
        tb.Checkbutton(p, text="Generar también el Excel", variable=self.excel,
                       bootstyle="success-round-toggle").pack(anchor="w", pady=4)
        hay_java = bool(shutil.which("java"))
        self.gramatica = tk.BooleanVar(value=bool(self.cfg.get("gramatica", hay_java)) and hay_java)
        tb.Checkbutton(p, text="Revisar gramática (requiere Java)", variable=self.gramatica,
                       bootstyle="success-round-toggle", command=self.aviso_gramatica).pack(anchor="w", pady=4)
        self.abrir = tk.BooleanVar(value=bool(self.cfg.get("abrir", True)))
        tb.Checkbutton(p, text="Abrir la carpeta al terminar", variable=self.abrir,
                       bootstyle="success-round-toggle").pack(anchor="w", pady=4)
        self._ajustar_al_ancho(p, [self.aviso_tipo])
        self.cambio_tipo()
        return p

    def _pestana_expediente(self, nb):
        p = tb.Frame(nb, padding=18)
        self.aviso_varios = tb.Frame(p, bootstyle="info", padding=(4, 0, 0, 0))
        nota = tb.Label(self.aviso_varios, text="ⓘ  Hay varios documentos: código, tesista, asesor y "
                                                "fecha se toman de registro.csv.", padding=(10, 8))
        nota.pack(fill="x")
        self.rejilla = tb.Frame(p)
        self.rejilla.pack(fill="x", pady=(14, 0))
        self.rejilla.columnconfigure(0, weight=1, uniform="e")
        self.rejilla.columnconfigure(1, weight=1, uniform="e")
        self.campos = {}
        for i, (clave, etiqueta) in enumerate((("codigo", "Código del expediente"),
                                               ("n_revision", "N° de revisión"),
                                               ("tesista", "Tesista"), ("asesor", "Asesor"),
                                               ("fecha", "Fecha de presentación (dd/mm/aaaa)"),
                                               ("revisor", "Revisor (firma)"))):
            celda = tb.Frame(self.rejilla)
            celda.grid(row=i // 2, column=i % 2, sticky="ew", padx=(0 if i % 2 == 0 else 10, 0), pady=5)
            tb.Label(celda, text=etiqueta, bootstyle="secondary", font=(FUENTE, 9)).pack(anchor="w")
            v = tk.StringVar(value=self.cfg.get(clave, "") if clave in ("revisor", "n_revision") else "")
            campo = tb.Entry(celda, textvariable=v)
            campo.pack(fill="x", pady=(3, 0))
            self.campos[clave] = (v, campo)

        self._titulo(p, "Observaciones adicionales")
        tb.Label(p, text="Una por línea. Se agregan al final de la hoja.", bootstyle="secondary",
                 font=(FUENTE, 8)).pack(anchor="w", pady=(0, 4))
        self.extras = tk.Text(p, height=4, wrap="word", font=(FUENTE, 9), bd=0,
                              highlightthickness=1, padx=8, pady=6)
        self.extras.pack(fill="x")
        self.extras.insert("1.0", self.cfg.get("extras", ""))
        self._ajustar_al_ancho(p, [nota])
        return p

    def _pestana_ajustes(self, nb):
        p = tb.Frame(nb, padding=18)
        self._titulo(p, "Reglas del revisor", arriba=0)
        explicacion = tb.Label(p, text="Criterios propios que se aplican en cada revisión, además de "
                                       "los del programa. Se escriben una vez y quedan.",
                               bootstyle="secondary", font=(FUENTE, 9), justify="left")
        explicacion.pack(anchor="w")
        fila = tb.Frame(p)
        fila.pack(fill="x", pady=(8, 0))
        tb.Button(fila, text="Editar reglas", bootstyle="secondary-outline",
                  command=self.editar_mis_reglas).pack(side="left")
        tb.Button(fila, text="Palabras permitidas", bootstyle="secondary-outline",
                  command=self.editar_permitidas).pack(side="left", padx=8)

        self._titulo(p, "Diseño de la hoja de revisión")
        formato = tb.Label(p, text="El diseño de la hoja de revisión (letra, logo, encabezado) se edita en "
                                   "Word; el texto de cada observación, en el archivo de frases.",
                           bootstyle="secondary", font=(FUENTE, 9), justify="left")
        formato.pack(anchor="w")
        fila = tb.Frame(p)
        fila.pack(fill="x", pady=(8, 0))
        tb.Button(fila, text="Diseño (Word)", bootstyle="secondary-outline",
                  command=self.editar_hoja).pack(side="left")
        tb.Button(fila, text="Frases", bootstyle="secondary-outline",
                  command=self.editar_frases).pack(side="left", padx=8)
        tb.Button(fila, text="Restablecer", bootstyle="link",
                  command=self.restablecer_formato).pack(side="left")

        self._titulo(p, "Carpeta de reportes")
        fila = tb.Frame(p)
        fila.pack(fill="x")
        self.salida = tk.StringVar(value=self.cfg.get("salida") or os.path.join(DATOS, "reportes"))
        tb.Entry(fila, textvariable=self.salida).pack(side="left", fill="x", expand=True)
        tb.Button(fila, text="Cambiar", bootstyle="secondary-outline",
                  command=self.elegir_salida).pack(side="left", padx=(8, 0))

        self._titulo(p, "Datos del programa")
        tb.Label(p, text="Registro de expedientes (registro.csv), reglas y configuración.",
                 bootstyle="secondary", font=(FUENTE, 9)).pack(anchor="w")
        tb.Button(p, text="Abrir carpeta de datos", bootstyle="link",
                  command=lambda: self.abrir_carpeta(DATOS)).pack(anchor="w", pady=(2, 0))
        self._ajustar_al_ancho(p, [explicacion, formato])
        return p

    def _barra_inferior(self):
        pie = tb.Frame(self)
        pie.pack(fill="x", side="bottom")
        self.progreso = tb.Progressbar(pie, mode="determinate", bootstyle="success-striped")
        self._separador_inferior = tb.Separator(pie)
        self._separador_inferior.pack(fill="x")
        barra = tb.Frame(pie, padding=(24, 12, 24, 14))
        barra.pack(fill="x")
        self.estado = tb.Label(barra, text="Listo.", bootstyle="secondary")
        self.estado.pack(side="left")
        self.boton = tb.Button(barra, text="▶  Revisar", bootstyle="primary", padding=(28, 10),
                               command=self.revisar)
        self.boton.pack(side="right")

    # ---------------------------------------------------------------- tema
    def _cambiar_tema(self):
        self.oscuro = not self.oscuro
        self.style.theme_use(_tema(self.style, self.oscuro))
        self._recolorear()

    def _recolorear(self):
        """lo que no es ttk no sigue al tema solo: el lienzo, los cuadros de texto y la tabla"""
        c = self.style.colors
        fondo = self.style.lookup("TFrame", "background")
        self.boton_tema.configure(text="☀  Tema claro" if self.oscuro else "☾  Tema oscuro")
        self.zona.configure(bg=fondo)
        self._dibujar_zona()
        campo = c.inputbg if hasattr(c, "inputbg") else fondo
        for texto in (self.log, self.extras):
            texto.configure(bg=campo, fg=c.inputfg, insertbackground=c.inputfg,
                            highlightbackground=c.border, highlightcolor=c.primary)
        self.log.tag_configure("err", foreground=c.danger)
        self.log.tag_configure("ok", foreground=c.success)
        self.log.tag_configure("suave", foreground=c.secondary)
        self.tabla.tag_configure("error", foreground=c.danger)
        self.tabla.tag_configure("activo", foreground=c.primary)
        self.tabla.tag_configure("pendiente", foreground=c.secondary)
        self.tabla.tag_configure("ok", foreground=c.fg)

    # ---------------------------------------------------------------- acciones
    def cambio_tipo(self):
        tipo = self.tipo.get()
        try:
            reglas = orquestador.cargar_reglas(tipo)
        except Exception as ex:
            self.aviso_tipo.configure(text=str(ex).split("\n")[0] +
                                      "  Genera las reglas desde la plantilla oficial (ver README).",
                                      bootstyle="danger")
            return
        self.aviso_tipo.configure(text=f"Esquema con {len(reglas['secciones'])} secciones.",
                                  bootstyle="secondary")

    def agregar_archivos(self):
        rutas = filedialog.askopenfilenames(
            title="Elegir documentos e informes de Turnitin",
            filetypes=[("Documentos e informes", "*.docx *.pdf"),
                       ("Documentos de Word", "*.docx"),
                       ("Informes de Turnitin", "*.pdf"),
                       ("Todos", "*.*")])
        self._sumar(rutas)

    def agregar_carpeta(self):
        carpeta = filedialog.askdirectory(title="Elegir carpeta con documentos")
        if carpeta:
            self._sumar_carpeta(carpeta)

    def _sumar_carpeta(self, carpeta):
        # el mismo filtro que usa la línea de comandos: descarta los ~$ y avisa de los .doc
        try:
            documentos, pdfs, avisos = orquestador.expandir([carpeta])
        except OSError as ex:
            messagebox.showerror("Revisor de Tesis", f"No se pudo leer la carpeta:\n{carpeta}\n\n{ex}")
            return
        self._sumar(documentos + pdfs)
        for a in avisos:
            self.avisar(a)
        if not documentos and not pdfs:
            self.avisar(f"{os.path.basename(carpeta)}: no hay archivos .docx ni informes .pdf.")

    def _soltar(self, evento):
        """archivos o carpetas arrastrados desde el explorador"""
        if self.trabajando:
            return evento.action
        archivos = []
        for ruta in self.tk.splitlist(evento.data):
            if os.path.isdir(ruta):
                self._sumar_carpeta(ruta)
            else:
                archivos.append(ruta)
        if archivos:
            self._sumar(archivos)
        return evento.action

    def _sumar(self, rutas):
        omitidos = 0
        for r in rutas:
            if os.path.basename(r).startswith("~$"):
                continue
            # el .pdf entra porque es el informe de similitud de Turnitin: se empareja
            # con su .docx por el nombre y sus observaciones van en la misma hoja
            if not r.lower().endswith((".docx", ".pdf")):
                omitidos += 1
                continue
            if r not in self.archivos:
                self.archivos.append(r)
        self._refrescar()
        if omitidos:
            self.avisar(f"Se omitieron {omitidos} archivo(s): solo se revisan .docx y los "
                        "informes de Turnitin en .pdf.")

    def quitar(self):
        if self.trabajando:
            return
        for iid in sorted((int(i) for i in self.tabla.selection()), reverse=True):
            del self.archivos[iid]
        self._refrescar()

    def limpiar(self):
        if self.trabajando:
            return
        self.archivos = []
        self._refrescar()

    def _unidades(self):
        """
        Cuántas hojas de revisión van a salir, aproximadamente: una por .docx, porque
        el PDF de Turnitin se suma a la hoja de su documento en vez de llevar una aparte.
        Solo cuando llegan PDF sin ningún .docx cada uno lleva su propia hoja.
        """
        docx = [r for r in self.archivos if r.lower().endswith(".docx")]
        return len(docx) if docx else len(self.archivos)

    def _refrescar(self):
        self.tabla.delete(*self.tabla.get_children())
        self.filas = {}
        for i, r in enumerate(self.archivos):
            nombre = os.path.basename(r)
            word = r.lower().endswith(".docx")
            self.tabla.insert("", "end", iid=str(i), text=f" {'📄' if word else '📑'}  {nombre}",
                              values=("Word" if word else "Turnitin", "Pendiente"), tags=("pendiente",))
            self.filas.setdefault(nombre.lower(), []).append(str(i))
        n = len(self.archivos)
        pdfs = sum(1 for r in self.archivos if r.lower().endswith(".pdf"))
        detalle = f" · {pdfs} informe(s) de Turnitin" if pdfs else ""
        self.conteo.configure(text="Sin documentos" if not n else f"{n} archivo(s){detalle}")
        varios = self._unidades() > 1
        for clave in ("codigo", "tesista", "asesor", "fecha"):
            self.campos[clave][1].configure(state="disabled" if varios else "normal")
        if varios:
            self.aviso_varios.pack(fill="x", before=self.rejilla)
        else:
            self.aviso_varios.pack_forget()

    def aviso_gramatica(self):
        if self.gramatica.get() and not shutil.which("java"):
            messagebox.showwarning("Revisor de Tesis",
                                   "No se encontró Java en este equipo.\n\nSin Java solo se revisa ortografía con "
                                   "el diccionario incluido. Para revisar gramática instala Java 17 o superior "
                                   "desde adoptium.net.")
            self.gramatica.set(False)
        elif self.gramatica.get():
            self.avisar("La primera revisión con gramática descarga LanguageTool (~260 MB). Solo ocurre una vez.")

    def _abrir_archivo(self, ruta):
        try:
            if EN_WINDOWS:
                os.startfile(ruta)  # noqa
            elif sys.platform == "darwin":
                subprocess.Popen(["open", ruta])
            else:
                subprocess.Popen(["xdg-open", ruta])
        except Exception as ex:
            messagebox.showerror("Revisor de Tesis", f"No se pudo abrir el archivo:\n{ruta}\n\n{ex}")

    def editar_mis_reglas(self):
        ruta = reglas_revisor.crear_si_falta(os.path.join(DATOS, "mis_reglas.yaml"))
        self._abrir_archivo(ruta)
        self.avisar("Se abrió mis_reglas.yaml. Guarda el archivo y vuelve a revisar; "
                    "los cambios se aplican de inmediato.")

    def editar_permitidas(self):
        ruta = ruta_recurso("permitidas.txt")
        destino = os.path.join(DATOS, "permitidas.txt")
        if ruta != destino and not os.path.exists(destino):
            try:
                shutil.copy2(ruta, destino)
            except Exception:
                destino = ruta
        self._abrir_archivo(destino)
        self.avisar("Agrega una palabra por línea: términos técnicos, apellidos y siglas que el "
                    "corrector marca por error.")

    # ---- formato editable de la hoja de revisión (carpeta de datos / reporte)
    FORMATO_REPORTE = ("hoja_revision.docx", "frases_reporte.yaml")

    def _archivo_de_reporte(self, nombre, reponer=False):
        """la copia editable del revisor; se crea desde la del programa si falta"""
        destino = os.path.join(DATOS, "reporte", nombre)
        original = ruta_recurso(os.path.join("reporte", nombre))
        if os.path.abspath(destino) == os.path.abspath(original):
            if reponer and nombre.endswith(".docx"):
                crear_plantilla_hoja(destino)     # al correr con python, la original es esta
            return destino
        if reponer or not os.path.exists(destino):
            os.makedirs(os.path.dirname(destino), exist_ok=True)
            if os.path.exists(original):
                shutil.copy2(original, destino)
            elif nombre.endswith(".docx"):
                crear_plantilla_hoja(destino)
        return destino

    def editar_hoja(self):
        self._abrir_archivo(self._archivo_de_reporte("hoja_revision.docx"))
        self.avisar("Se abrió la plantilla de la hoja. Cambia letra, logo o textos fijos y guarda; "
                    "no borres los marcadores {{…}} que necesites (ej. {{OBSERVACIONES}}).")

    def editar_frases(self):
        self._abrir_archivo(self._archivo_de_reporte("frases_reporte.yaml"))
        self.avisar("Se abrió frases_reporte.yaml. Cambia el texto entre comillas y guarda; "
                    "vale desde la siguiente revisión.")

    def restablecer_formato(self):
        if not messagebox.askyesno("Revisor de Tesis",
                                   "¿Volver al formato original de la hoja de revisión?\n\n"
                                   "Se pierden los cambios hechos en el diseño y en las frases."):
            return
        try:
            for nombre in self.FORMATO_REPORTE:
                self._archivo_de_reporte(nombre, reponer=True)
        except OSError as ex:
            messagebox.showerror("Revisor de Tesis", f"No se pudo restablecer el formato "
                                                     f"(¿está abierto en Word?):\n\n{ex}")
            return
        self.avisar("Se restableció el formato original de la hoja de revisión.")

    def elegir_salida(self):
        carpeta = filedialog.askdirectory(title="Carpeta donde guardar los reportes")
        if carpeta:
            self.salida.set(carpeta)

    def _carpeta_salida(self):
        return self.salida.get().strip() or os.path.join(DATOS, "reportes")

    def abrir_carpeta(self, ruta):
        try:
            os.makedirs(ruta, exist_ok=True)
            if EN_WINDOWS:
                os.startfile(ruta)  # noqa
            elif sys.platform == "darwin":
                subprocess.Popen(["open", ruta])
            else:
                subprocess.Popen(["xdg-open", ruta])
        except Exception:
            pass

    def escribir(self, texto, etiqueta=None):
        self.log.configure(state="normal")
        self.log.insert("end", texto + "\n", etiqueta or ())
        self.log.see("end")
        self.log.configure(state="disabled")

    def avisar(self, texto):
        """un aviso para el revisor: va al registro y a la barra de estado, que siempre se ve"""
        self.escribir(texto, "suave")
        self.estado.configure(text=texto if len(texto) <= 110 else texto[:107] + "…")

    # ---------------------------------------------------------------- proceso
    def revisar(self):
        if self.trabajando:
            return
        if not self.archivos:
            messagebox.showinfo("Revisor de Tesis",
                                "Agrega al menos un documento .docx (y, si lo tienes, "
                                "el informe de Turnitin en .pdf).")
            return
        try:
            orquestador.cargar_reglas(self.tipo.get())   # avisar del error antes de empezar
        except Exception as ex:
            messagebox.showerror("Revisor de Tesis", str(ex))
            return
        salida = self._carpeta_salida()
        try:
            os.makedirs(salida, exist_ok=True)
        except Exception as ex:
            messagebox.showerror("Revisor de Tesis", f"No se pudo crear la carpeta de reportes:\n{ex}")
            return

        uno = self._unidades() == 1
        fecha = self.campos["fecha"][0].get().strip()
        if uno and fecha and not self._fecha_valida(fecha) and not messagebox.askyesno(
                "Revisor de Tesis",
                f"La fecha de presentación «{fecha}» no tiene la forma dd/mm/aaaa "
                "(por ejemplo 15/03/2026).\n\n¿Usarla así en la hoja de revisión?"):
            return
        opciones = Opciones(
            tipo=self.tipo.get(), salida=salida, excel=self.excel.get(),
            revisor=self.campos["revisor"][0].get().strip(),
            tesista=self.campos["tesista"][0].get().strip() if uno else "",
            asesor=self.campos["asesor"][0].get().strip() if uno else "",
            fecha=fecha if uno else "",
            codigo=self.campos["codigo"][0].get().strip() if uno else "",
            n_revision=self.campos["n_revision"][0].get().strip(),
            formato=self.formato.get(),
            extras=[x for x in self.extras.get("1.0", "end").splitlines() if x.strip()],
            motor="languagetool" if self.gramatica.get() else "diccionario")

        self.trabajando = True
        self._hubo_error = False
        self._inicio = time.monotonic()
        self._estado_base = "Preparando la revisión…"
        self._totales = dict(errores=0, advertencias=0, similitud=None)
        self.parejas = {}
        self._refrescar()
        self.boton.configure(state="disabled", text="Revisando…")
        self.progreso.pack(fill="x", side="top", before=self._separador_inferior)
        self.progreso.configure(maximum=max(self._unidades(), 1), value=0)
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")
        self.escribir(f"Tipo: {opciones.tipo}   ·   {len(self.archivos)} archivo(s)", "suave")

        self._guardar_config()
        threading.Thread(target=self._trabajar, args=(list(self.archivos), opciones),
                         daemon=True).start()

    @staticmethod
    def _fecha_valida(texto):
        try:
            datetime.strptime(texto, "%d/%m/%Y")
            return True
        except ValueError:
            return False

    @staticmethod
    def _color(mensaje):
        """con qué color va cada línea del registro"""
        bajo = mensaje.lower()
        if "no revisado" in bajo or "no se pudo" in bajo or "no existe" in bajo:
            return "err"
        if " 0 errores" in mensaje:
            return "ok"
        if "no se revisan" in bajo or "no parece un informe" in bajo or bajo.startswith("se omite"):
            return "suave"
        return None

    @staticmethod
    def interpretar(mensaje):
        """
        Traduce una línea del registro al estado de un archivo en la tabla. Devuelve
        None si la línea no habla de un archivo. Las líneas son las que emite el
        orquestador; si cambian allá, los tests de la app lo delatan.
        """
        m = _ARCHIVO.match(mensaje)
        if not m:
            return None
        archivo, resto = m.groups()
        similitud = _SIMILITUD.search(resto)
        similitud = float(similitud.group(1)) if similitud else None
        estado = dict(archivo=archivo, similitud=similitud, errores=0, advertencias=0, hoja=False)

        r = _RESULTADO.match(resto)
        if r:
            err, adv = int(r.group(1)), int(r.group(2))
            texto = f"{'✗' if err else '✓'}  {err} errores · {adv} advertencias"
            return dict(estado, estado=texto, etiqueta="error" if err else "ok",
                        errores=err, advertencias=adv, hoja=True)
        if resto.startswith("informe de Turnitin,"):
            e = _ERRORES_FINAL.search(resto)
            err = int(e.group(1)) if e else 0
            texto = "Hoja propia" + (f" · similitud {_pct(similitud)}" if similitud is not None else "")
            return dict(estado, estado=texto, etiqueta="error" if err else "ok", errores=err, hoja=True)
        pareja = _PAREJA.match(resto)
        if pareja:
            return dict(estado, archivo=pareja.group(1), estado=f"Emparejado con {archivo}",
                        etiqueta="ok", pareja=archivo)
        if "NO REVISADO" in resto:
            return dict(estado, estado="No revisado: no parece de este tipo", etiqueta="error")
        if resto.startswith("no se pudo revisar"):
            return dict(estado, estado="No se pudo revisar", etiqueta="error")
        if resto.startswith("no se pudo leer el PDF"):
            return dict(estado, estado="PDF ilegible", etiqueta="error")
        if resto.startswith("no parece un informe de Turnitin"):
            return dict(estado, estado="No es un informe de Turnitin", etiqueta="pendiente")
        return None

    def _trabajar(self, archivos, opciones):
        """
        Corre la revisión en su hilo. Delega todo en orquestador.ejecutar para que la app
        y la línea de comandos hagan exactamente lo mismo: el emparejado de cada .docx con
        su informe de Turnitin vive ahí, y antes esta función se lo saltaba.
        """
        def aviso(mensaje):
            cuerpo = mensaje.split(" -> ")[0]
            self.cola.put(("log", (cuerpo, self._color(cuerpo))))

        def progreso(hechas, total, nombre):
            self.cola.put(("total", total))
            self.cola.put(("avance", hechas))
            if nombre:
                self.cola.put(("estado", f"Revisando {nombre}…"))

        hechos = []
        try:
            if opciones.motor == "languagetool":
                self.cola.put(("log", ("Iniciando LanguageTool para la gramática: el primer documento "
                                       "tarda más.", "suave")))
            hechos = orquestador.ejecutar(archivos, opciones, aviso, progreso)
        except ValueError as ex:
            # lo que ejecutar() rechaza a propósito: no hay nada que revisar, datos de
            # expediente con varios archivos, el informe indicado no se pudo leer
            self.cola.put(("log", (str(ex), "err")))
        except BaseException as ex:
            import traceback
            tb = traceback.format_exc()
            print(tb)  # para la consola cuando se ejecuta con python
            self.cola.put(("log", (f"No se pudo completar la revisión "
                                   f"({type(ex).__name__}: {ex})", "err")))
            self.cola.put(("log", (tb, "err")))
        finally:
            # pase lo que pase, la interfaz tiene que salir de «Revisando…»
            self.cola.put(("fin", (len(hechos), opciones.salida)))

    # ---------------------------------------------------------------- tabla y tarjetas
    def _marcar(self, nombre, texto, etiqueta):
        for iid in self.filas.get(nombre.lower(), ()):
            if self.tabla.exists(iid):
                self.tabla.set(iid, "estado", texto)
                self.tabla.item(iid, tags=(etiqueta,) if etiqueta else ())

    def _registrar(self, texto):
        """lleva una línea del registro a la tabla y a los totales"""
        info = self.interpretar(texto)
        if not info:
            return
        self._marcar(info["archivo"], info["estado"], info["etiqueta"])
        if "pareja" in info:
            self.parejas[info["pareja"].lower()] = info["archivo"]
        if info["hoja"]:
            self._totales["errores"] += info["errores"]
            self._totales["advertencias"] += info["advertencias"]
        if info["similitud"] is not None:
            actual = self._totales["similitud"]
            self._totales["similitud"] = max(actual or 0, info["similitud"])
            pdf = self.parejas.get(info["archivo"].lower())
            if pdf:
                self._marcar(pdf, f"Emparejado · similitud {_pct(info['similitud'])}", "ok")

    def _mostrar_totales(self, hojas):
        self.tarjetas["hojas"][1].configure(text=str(hojas))
        self.tarjetas["errores"][1].configure(text=str(self._totales["errores"]))
        self.tarjetas["advertencias"][1].configure(text=str(self._totales["advertencias"]))
        similitud = self._totales["similitud"]
        borde, cifra = self.tarjetas["similitud"]
        if similitud is None:
            borde.grid_remove()
        else:
            cifra.configure(text=_pct(similitud))
            borde.grid()

    def _vaciar_cola(self):
        """
        Pasa a la pantalla lo que manda el hilo de la revisión. Un error al mostrar un
        mensaje no puede cortar este ciclo: si se corta, la ventana queda en «Revisando…»
        para siempre, y en el .exe sin consola nadie ve por qué.
        """
        try:
            while True:
                try:
                    clase, valor = self.cola.get_nowait()
                except queue.Empty:
                    break
                try:
                    self._procesar(clase, valor)
                except Exception:
                    self._anotar_error("mostrando un mensaje de la revisión")
                    if clase == "fin":
                        self._terminar_interfaz()
            if self.trabajando and self._estado_base:
                segundos = int(time.monotonic() - self._inicio)
                self.estado.configure(text=f"{self._estado_base}   {segundos // 60}:{segundos % 60:02d}")
        finally:
            self.after(120, self._vaciar_cola)

    def _procesar(self, clase, valor):
        if clase == "log":
            texto, etiqueta = valor
            self.escribir(texto, etiqueta)
            self._registrar(texto)
            if etiqueta == "err":
                self._hubo_error = True
        elif clase == "estado":
            self._estado_base = valor
            nombre = valor[len("Revisando "):].rstrip("…")
            self._marcar(nombre, "Revisando…", "activo")
        elif clase == "total":
            self.progreso.configure(maximum=max(valor, 1))
        elif clase == "avance":
            # valor absoluto, no un paso: ejecutar() informa cuántas van hechas
            self.progreso.configure(value=valor)
        elif clase == "fin":
            hechos, salida = valor
            self._terminar_interfaz()
            self._mostrar_totales(hechos)
            if not hechos:
                self.estado.configure(text="La revisión no generó reportes. Ver el registro técnico.")
            elif self._hubo_error:
                self.estado.configure(text=f"Terminado con errores: {hechos} reporte(s) "
                                           "generado(s). Ver el registro técnico.")
            else:
                self.estado.configure(text=f"Terminado. {hechos} reporte(s) generado(s).")
            if self._hubo_error and not self.ver_registro.get():
                self.ver_registro.set(True)
                self._mostrar_registro()
            if hechos:
                self.escribir(f"Reportes en: {salida}", "suave")
            self._guardar_registro()
            if hechos and self.abrir.get():
                self.abrir_carpeta(salida)

    def _terminar_interfaz(self):
        self.trabajando = False
        self._estado_base = ""
        self.boton.configure(state="normal", text="▶  Revisar")
        self.progreso.pack_forget()

    # ---------------------------------------------------------------- errores y registro
    def _guardar_registro(self):
        """deja el registro de la última revisión en un archivo, para poder diagnosticar"""
        try:
            with open(os.path.join(DATOS, "ultima_revision.log"), "w", encoding="utf-8") as fh:
                fh.write(datetime.now().strftime("%d/%m/%Y %H:%M:%S") + "\n")
                fh.write(self.log.get("1.0", "end"))
        except Exception:
            pass

    def _anotar_error(self, donde):
        """un error de la interfaz: al registro, a errores.log y a la barra de estado"""
        import traceback
        tb = traceback.format_exc()
        print(tb)
        try:
            with open(os.path.join(DATOS, "errores.log"), "a", encoding="utf-8") as fh:
                fh.write(f"--- {datetime.now():%d/%m/%Y %H:%M:%S}  {donde}\n{tb}\n")
        except Exception:
            pass
        try:
            self.escribir(f"Error de la interfaz {donde}. Detalle en errores.log (carpeta de datos).", "err")
            self.escribir(tb, "err")
            self._hubo_error = True
        except Exception:
            pass

    def report_callback_exception(self, tipo, valor, rastro):
        """
        Tk manda aquí los errores de botones y temporizadores. Por defecto van a la
        consola, que en el .exe no existe: se perdían sin dejar rastro.
        """
        try:
            raise valor.with_traceback(rastro)
        except BaseException:
            self._anotar_error("al atender la ventana")


def main():
    App().mainloop()


if __name__ == "__main__":
    main()
