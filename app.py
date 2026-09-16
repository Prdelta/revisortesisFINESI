"""
Revisor de Tesis FINESI. Aplicación de escritorio.

Ejecutar:  python app.py
"""
import json
import os
import queue
import shutil
import subprocess
import sys
import threading
import tkinter as tk
import ttkbootstrap as tb
from tkinter import filedialog, messagebox, ttk

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import revisor  # noqa: E402

FONDO = "#222222"
PANEL = "#303030"
BORDE = "#444444"
TINTA = "#EEEEEE"
SUAVE = "#AAAAAA"
ACENTO = "#375A7F"
ACENTO_CLARO = "#2B4764"
ROJO = "#E74C3C"
VERDE = "#00BC8C"
FUENTE = "Segoe UI" if sys.platform.startswith("win") else "DejaVu Sans"
MONO = "Consolas" if sys.platform.startswith("win") else "monospace"


class App(tb.Window):
    def __init__(self):
        super().__init__(themename="darkly")
        revisor.preparar_datos()
        self.title("Revisor de Tesis  ·  FINESI")
        self.configure(bg=FONDO)
        self.geometry("1020x790")
        self.minsize(920, 700)
        self._icono()
        self.archivos = []
        self.cola = queue.Queue()
        self.trabajando = False
        self.cfg = self._leer_config()
        self._estilos()
        self._construir()
        self.protocol("WM_DELETE_WINDOW", self._cerrar)
        self.after(120, self._vaciar_cola)

    # ---------------------------------------------------------------- preferencias
    def _ruta_config(self):
        return os.path.join(revisor.DATOS, "config.json")

    def _leer_config(self):
        try:
            with open(self._ruta_config(), encoding="utf8") as fh:
                return json.load(fh)
        except Exception:
            return {}

    def _guardar_config(self):
        try:
            datos = dict(revisor=self.campos["revisor"][0].get().strip(),
                         n_revision=self.campos["n_revision"][0].get().strip(),
                         extras=self.extras.get("1.0", "end").strip(),
                         formato=self.formato.get(),
                         tipo=self.tipo.get(), salida=self.salida.get().strip(),
                         excel=bool(self.excel.get()), gramatica=bool(self.gramatica.get()),
                         abrir=bool(self.abrir.get()))
            with open(self._ruta_config(), "w", encoding="utf8") as fh:
                json.dump(datos, fh, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def _cerrar(self):
        self._guardar_config()
        self.destroy()

    def _icono(self):
        for nombre in ("icono.ico", "icono.png"):
            ruta = revisor.ruta_recurso(nombre)
            if not os.path.exists(ruta):
                continue
            try:
                if nombre.endswith(".ico") and sys.platform.startswith("win"):
                    self.iconbitmap(ruta)
                else:
                    self.iconphoto(True, tk.PhotoImage(file=ruta))
                return
            except Exception:
                continue

    # ---------------------------------------------------------------- apariencia
    def _estilos(self):
        e = ttk.Style(self)
        
        e.configure(".", background=FONDO, foreground=TINTA, font=(FUENTE, 10))
        e.configure("TFrame", background=FONDO)
        e.configure("TLabel", background=FONDO, foreground=TINTA)
        e.configure("Institucion.TLabel", font=(FUENTE, 8, "bold"), foreground=SUAVE)
        e.configure("Titulo.TLabel", font=(FUENTE, 17, "bold"), foreground=ACENTO)
        e.configure("Sub.TLabel", font=(FUENTE, 9), foreground=SUAVE)
        e.configure("Seccion.TLabel", font=(FUENTE, 9, "bold"), foreground=ACENTO)
        e.configure("Campo.TLabel", font=(FUENTE, 9), foreground=SUAVE)
        e.configure("TEntry", fieldbackground="#303030", bordercolor=BORDE,
                    lightcolor=BORDE, darkcolor=BORDE, insertcolor=TINTA)
        e.configure("TRadiobutton", background=FONDO, foreground=TINTA, font=(FUENTE, 10))
        e.map("TRadiobutton", background=[("active", "#303030")])
        e.configure("TCheckbutton", background=FONDO, foreground=TINTA, font=(FUENTE, 9))
        e.map("TCheckbutton", background=[("active", "#303030")])
        e.configure("TButton", font=(FUENTE, 9), padding=(12, 6), background=PANEL,
                    foreground=TINTA, borderwidth=1, relief="flat")
        e.map("TButton", background=[("active", "#444444")])
        e.configure("Primario.TButton", font=(FUENTE, 10, "bold"), padding=(24, 9),
                    background=ACENTO, foreground="#FFFFFF", borderwidth=0)
        e.map("Primario.TButton", background=[("active", ACENTO_CLARO), ("disabled", "#A7AEBA")])
        e.configure("TProgressbar", background=ACENTO, troughcolor=PANEL, borderwidth=0, thickness=3)

    def _linea(self, padre):
        tk.Frame(padre, bg=BORDE, height=1).pack(fill="x")

    # ---------------------------------------------------------------- interfaz
    def _construir(self):
        cab = ttk.Frame(self, padding=(26, 14, 26, 12))
        cab.pack(fill="x")
        ttk.Label(cab, text="UNIVERSIDAD NACIONAL DEL ALTIPLANO – PUNO",
                  style="Institucion.TLabel").pack(anchor="w")
        ttk.Label(cab, text="Revisor de Tesis", style="Titulo.TLabel").pack(anchor="w", pady=(2, 0))
        ttk.Label(cab, text="Facultad de Ingeniería Estadística e Informática  ·  Dirección de Investigación",
                  style="Sub.TLabel").pack(anchor="w", pady=(3, 0))
        self._linea(self)

        cuerpo = ttk.Frame(self, padding=(26, 16, 26, 0))
        cuerpo.pack(fill="both", expand=True)
        cuerpo.columnconfigure(0, weight=3, uniform="c")
        cuerpo.columnconfigure(1, weight=2, uniform="c")
        cuerpo.rowconfigure(0, weight=1)
        izq = ttk.Frame(cuerpo)
        izq.grid(row=0, column=0, sticky="nsew", padx=(0, 20))
        der = self._columna_con_scroll(cuerpo)
        self._panel_izquierdo(izq)
        self._panel_derecho(der)
        self._barra_inferior()

    def _columna_con_scroll(self, cuerpo):
        """la columna derecha es larga: se desplaza con la rueda si la ventana es pequeña"""
        caja = ttk.Frame(cuerpo)
        caja.grid(row=0, column=1, sticky="nsew")
        lienzo = tk.Canvas(caja, bg=FONDO, highlightthickness=0, bd=0)
        barra = ttk.Scrollbar(caja, orient="vertical", command=lienzo.yview)
        interior = ttk.Frame(lienzo)
        ventana = lienzo.create_window((0, 0), window=interior, anchor="nw")
        lienzo.configure(yscrollcommand=barra.set)
        caja.columnconfigure(0, weight=1)
        caja.rowconfigure(0, weight=1)
        lienzo.grid(row=0, column=0, sticky="nsew")

        def ajustar(_=None):
            lienzo.configure(scrollregion=lienzo.bbox("all"))
            lienzo.itemconfigure(ventana, width=lienzo.winfo_width())
            hace_falta = interior.winfo_reqheight() > lienzo.winfo_height()
            if hace_falta and not barra.winfo_ismapped():
                barra.grid(row=0, column=1, sticky="ns")
            elif not hace_falta and barra.winfo_ismapped():
                barra.grid_remove()

        interior.bind("<Configure>", ajustar)
        lienzo.bind("<Configure>", ajustar)

        def rueda(evento):
            if interior.winfo_reqheight() <= lienzo.winfo_height():
                return
            paso = -1 if getattr(evento, "delta", 0) > 0 or evento.num == 4 else 1
            lienzo.yview_scroll(paso, "units")

        for widget, secuencia in ((self, "<MouseWheel>"), (self, "<Button-4>"), (self, "<Button-5>")):
            widget.bind_all(secuencia, rueda, add="+")
        return interior

    def _panel_izquierdo(self, padre):
        ttk.Label(padre, text="DOCUMENTOS A REVISAR", style="Seccion.TLabel").pack(anchor="w")
        barra = ttk.Frame(padre)
        barra.pack(fill="x", pady=(8, 8))
        ttk.Button(barra, text="Agregar archivos", command=self.agregar_archivos).pack(side="left")
        ttk.Button(barra, text="Agregar carpeta", command=self.agregar_carpeta).pack(side="left", padx=6)
        ttk.Button(barra, text="Quitar", command=self.quitar).pack(side="left")
        ttk.Button(barra, text="Limpiar", command=self.limpiar).pack(side="left", padx=6)

        marco = tk.Frame(padre, bg=FONDO, highlightthickness=1, highlightbackground=BORDE)
        marco.pack(fill="both", expand=True)
        self.lista = tk.Listbox(marco, bg="#303030", fg=TINTA, font=(FUENTE, 9), bd=0,
                                highlightthickness=0, selectmode="extended", activestyle="none",
                                selectbackground="#375A7F", selectforeground="#FFFFFF")
        scroll = ttk.Scrollbar(marco, orient="vertical", command=self.lista.yview)
        self.lista.configure(yscrollcommand=scroll.set)
        self.lista.pack(side="left", fill="both", expand=True, padx=1, pady=1)
        scroll.pack(side="right", fill="y")
        self.conteo = ttk.Label(padre, text="Sin documentos.", style="Sub.TLabel")
        self.conteo.pack(anchor="w", pady=(6, 0))

        ttk.Label(padre, text="REGISTRO DE LA REVISIÓN", style="Seccion.TLabel").pack(anchor="w", pady=(14, 6))
        self.log = tk.Text(padre, height=9, bg=PANEL, fg=TINTA, font=(MONO, 9), bd=0,
                           highlightthickness=1, highlightbackground=BORDE, wrap="word", padx=10, pady=8)
        self.log.pack(fill="both", expand=True)
        self.log.tag_configure("err", foreground=ROJO)
        self.log.tag_configure("ok", foreground=VERDE)
        self.log.tag_configure("suave", foreground=SUAVE)
        self.log.configure(state="disabled")

    def _panel_derecho(self, padre):
        ttk.Label(padre, text="TIPO DE DOCUMENTO", style="Seccion.TLabel").pack(anchor="w")
        self.tipos = revisor.tipos_disponibles() or [("proyecto", False)]
        nombres = [t for t, _ in self.tipos]
        guardado = self.cfg.get("tipo")
        self.tipo = tk.StringVar(value=guardado if guardado in nombres else
                                 ("proyecto" if "proyecto" in nombres else nombres[0]))
        fila = ttk.Frame(padre)
        fila.pack(fill="x", pady=(8, 2))
        for t, _ in self.tipos:
            ttk.Radiobutton(fila, text=t.capitalize(), value=t, variable=self.tipo,
                            command=self.cambio_tipo).pack(side="left", padx=(0, 16))
        self.aviso_tipo = ttk.Label(padre, text="", style="Sub.TLabel", wraplength=310, justify="left")
        self.aviso_tipo.pack(anchor="w", pady=(2, 0))

        ttk.Label(padre, text="DATOS DEL EXPEDIENTE", style="Seccion.TLabel").pack(anchor="w", pady=(16, 2))
        self.nota_datos = ttk.Label(padre, text="Se usan al revisar un solo documento.",
                                    style="Sub.TLabel", wraplength=310, justify="left")
        self.nota_datos.pack(anchor="w", pady=(0, 8))
        self.campos = {}
        for clave, etiqueta in (("codigo", "Código del expediente"), ("tesista", "Tesista"),
                                ("asesor", "Asesor"), ("fecha", "Fecha de presentación"),
                                ("n_revision", "N° de revisión"), ("revisor", "Revisor (firma)")):
            ttk.Label(padre, text=etiqueta, style="Campo.TLabel").pack(anchor="w")
            v = tk.StringVar(value=self.cfg.get(clave, "") if clave in ("revisor", "n_revision") else "")
            campo = ttk.Entry(padre, textvariable=v, font=(FUENTE, 10))
            campo.pack(fill="x", pady=(2, 8), ipady=3)
            self.campos[clave] = (v, campo)

        ttk.Label(padre, text="OBSERVACIONES ADICIONALES", style="Seccion.TLabel").pack(anchor="w", pady=(6, 2))
        ttk.Label(padre, text="Una por línea. Se agregan al final de la hoja.",
                  style="Sub.TLabel").pack(anchor="w", pady=(0, 6))
        self.extras = tk.Text(padre, height=3, bg="#303030", fg=TINTA, font=(FUENTE, 9), bd=0,
                              highlightthickness=1, highlightbackground=BORDE, wrap="word", padx=8, pady=6)
        self.extras.pack(fill="x")
        self.extras.insert("1.0", self.cfg.get("extras", ""))

        ttk.Label(padre, text="REGLAS DEL REVISOR", style="Seccion.TLabel").pack(anchor="w", pady=(16, 2))
        ttk.Label(padre, text="Criterios propios que se aplican en cada revisión, además de los del "
                              "programa. Se escriben una vez y quedan.",
                  style="Sub.TLabel", wraplength=310, justify="left").pack(anchor="w", pady=(0, 6))
        f = ttk.Frame(padre)
        f.pack(fill="x")
        ttk.Button(f, text="Editar reglas", command=self.editar_mis_reglas).pack(side="left")
        ttk.Button(f, text="Palabras permitidas", command=self.editar_permitidas).pack(side="left", padx=6)

        ttk.Label(padre, text="FORMATO DEL REPORTE", style="Seccion.TLabel").pack(anchor="w", pady=(16, 6))
        self.formato = tk.StringVar(value=self.cfg.get("formato", "resumen"))
        for valor, etiqueta in (("resumen", "Hoja de revisión (corta)"),
                                ("detallado", "Detallado (tabla completa)"),
                                ("ambos", "Los dos")):
            ttk.Radiobutton(padre, text=etiqueta, value=valor, variable=self.formato).pack(anchor="w")

        ttk.Label(padre, text="OPCIONES", style="Seccion.TLabel").pack(anchor="w", pady=(16, 6))
        self.excel = tk.BooleanVar(value=bool(self.cfg.get("excel", False)))
        ttk.Checkbutton(padre, text="Generar también el Excel", variable=self.excel).pack(anchor="w")
        self.gramatica = tk.BooleanVar(value=bool(self.cfg.get("gramatica", bool(shutil.which("java"))))
                                       and bool(shutil.which("java")))
        ttk.Checkbutton(padre, text="Revisar gramática (requiere Java)", variable=self.gramatica,
                        command=self.aviso_gramatica).pack(anchor="w")
        self.abrir = tk.BooleanVar(value=bool(self.cfg.get("abrir", True)))
        ttk.Checkbutton(padre, text="Abrir la carpeta al terminar", variable=self.abrir).pack(anchor="w")

        ttk.Label(padre, text="CARPETA DE REPORTES", style="Seccion.TLabel").pack(anchor="w", pady=(16, 6))
        f = ttk.Frame(padre)
        f.pack(fill="x")
        self.salida = tk.StringVar(value=self.cfg.get("salida") or os.path.join(revisor.DATOS, "reportes"))
        ttk.Entry(f, textvariable=self.salida, font=(FUENTE, 9)).pack(side="left", fill="x", expand=True, ipady=3)
        ttk.Button(f, text="Cambiar", command=self.elegir_salida).pack(side="left", padx=(6, 0))
        self.cambio_tipo()

    def _barra_inferior(self):
        contenedor = ttk.Frame(self)
        contenedor.pack(fill="x", side="bottom")
        self.progreso = ttk.Progressbar(contenedor, mode="determinate")
        self._linea(contenedor)
        barra = ttk.Frame(contenedor, padding=(26, 12, 26, 16))
        barra.pack(fill="x")
        self.estado = ttk.Label(barra, text="Listo.", style="Sub.TLabel")
        self.estado.pack(side="left")
        self.boton = ttk.Button(barra, text="Revisar", style="Primario.TButton", command=self.revisar)
        self.boton.pack(side="right")
        ttk.Button(barra, text="Carpeta de datos", command=lambda: self.abrir_carpeta(revisor.DATOS)).pack(
            side="right", padx=10)

    # ---------------------------------------------------------------- acciones
    def cambio_tipo(self):
        tipo = self.tipo.get()
        try:
            reglas = revisor.cargar_reglas(tipo)
        except Exception as ex:
            self.aviso_tipo.configure(text=str(ex).split("\n")[0] +
                                      f"  Genera las reglas desde la plantilla oficial (ver LEEME).",
                                      foreground=ROJO)
            return
        self.aviso_tipo.configure(text=f"Esquema con {len(reglas['secciones'])} secciones.", foreground=SUAVE)

    def agregar_archivos(self):
        rutas = filedialog.askopenfilenames(title="Elegir documentos",
                                            filetypes=[("Documentos de Word", "*.docx"), ("Todos", "*.*")])
        self._sumar(rutas)

    def agregar_carpeta(self):
        carpeta = filedialog.askdirectory(title="Elegir carpeta con documentos")
        if carpeta:
            self._sumar([os.path.join(carpeta, f) for f in sorted(os.listdir(carpeta))])

    def _sumar(self, rutas):
        omitidos = 0
        for r in rutas:
            if os.path.basename(r).startswith("~$"):
                continue
            if not r.lower().endswith(".docx"):
                omitidos += 1
                continue
            if r not in self.archivos:
                self.archivos.append(r)
        self._refrescar()
        if omitidos:
            self.escribir(f"Se omitieron {omitidos} archivo(s) que no son .docx.", "suave")

    def quitar(self):
        for i in sorted(self.lista.curselection(), reverse=True):
            del self.archivos[i]
        self._refrescar()

    def limpiar(self):
        self.archivos = []
        self._refrescar()

    def _refrescar(self):
        self.lista.delete(0, "end")
        for r in self.archivos:
            self.lista.insert("end", "  " + os.path.basename(r))
        n = len(self.archivos)
        self.conteo.configure(text="Sin documentos." if not n else f"{n} documento(s) en la lista.")
        varios = n > 1
        for clave in ("codigo", "tesista", "asesor", "fecha"):
            self.campos[clave][1].configure(state="disabled" if varios else "normal")
        self.nota_datos.configure(
            text="Varios documentos: código, tesista, asesor y fecha se toman de registro.csv." if varios
            else "Se usan al revisar un solo documento.")

    def aviso_gramatica(self):
        if self.gramatica.get() and not shutil.which("java"):
            messagebox.showwarning("Revisor de Tesis",
                                   "No se encontró Java en este equipo.\n\nSin Java solo se revisa ortografía con "
                                   "el diccionario incluido. Para revisar gramática instala Java 17 o superior "
                                   "desde adoptium.net.")
            self.gramatica.set(False)
        elif self.gramatica.get():
            self.escribir("La primera revisión con gramática descarga LanguageTool (~260 MB). Solo ocurre una vez.",
                          "suave")

    def _abrir_archivo(self, ruta):
        try:
            if sys.platform.startswith("win"):
                os.startfile(ruta)  # noqa
            elif sys.platform == "darwin":
                subprocess.Popen(["open", ruta])
            else:
                subprocess.Popen(["xdg-open", ruta])
        except Exception as ex:
            messagebox.showerror("Revisor de Tesis", f"No se pudo abrir el archivo:\n{ruta}\n\n{ex}")

    def editar_mis_reglas(self):
        import reglas_revisor
        ruta = reglas_revisor.crear_si_falta(os.path.join(revisor.DATOS, "mis_reglas.yaml"))
        self._abrir_archivo(ruta)
        self.escribir("Se abrió mis_reglas.yaml. Guarda el archivo y vuelve a revisar; "
                      "los cambios se aplican de inmediato.", "suave")

    def editar_permitidas(self):
        ruta = revisor.ruta_recurso("permitidas.txt")
        destino = os.path.join(revisor.DATOS, "permitidas.txt")
        if ruta != destino and not os.path.exists(destino):
            try:
                shutil.copy2(ruta, destino)
            except Exception:
                destino = ruta
        self._abrir_archivo(destino)
        self.escribir("Agrega una palabra por línea: términos técnicos, apellidos y siglas que el "
                      "corrector marca por error.", "suave")

    def elegir_salida(self):
        carpeta = filedialog.askdirectory(title="Carpeta donde guardar los reportes")
        if carpeta:
            self.salida.set(carpeta)

    def abrir_carpeta(self, ruta):
        try:
            os.makedirs(ruta, exist_ok=True)
            if sys.platform.startswith("win"):
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

    # ---------------------------------------------------------------- proceso
    def revisar(self):
        if self.trabajando:
            return
        if not self.archivos:
            messagebox.showinfo("Revisor de Tesis", "Agrega al menos un documento .docx.")
            return
        try:
            reglas = revisor.cargar_reglas(self.tipo.get())
        except Exception as ex:
            messagebox.showerror("Revisor de Tesis", str(ex))
            return
        salida = self.salida.get().strip() or os.path.join(revisor.DATOS, "reportes")
        try:
            os.makedirs(salida, exist_ok=True)
        except Exception as ex:
            messagebox.showerror("Revisor de Tesis", f"No se pudo crear la carpeta de reportes:\n{ex}")
            return

        uno = len(self.archivos) == 1
        opciones = revisor.Opciones(
            tipo=self.tipo.get(), salida=salida, excel=self.excel.get(),
            revisor=self.campos["revisor"][0].get().strip(),
            tesista=self.campos["tesista"][0].get().strip() if uno else "",
            asesor=self.campos["asesor"][0].get().strip() if uno else "",
            fecha=self.campos["fecha"][0].get().strip() if uno else "",
            codigo=self.campos["codigo"][0].get().strip() if uno else "",
            n_revision=self.campos["n_revision"][0].get().strip(),
            formato=self.formato.get(),
            extras=[x for x in self.extras.get("1.0", "end").splitlines() if x.strip()],
            motor="languagetool" if self.gramatica.get() else "diccionario")

        self.trabajando = True
        self.boton.configure(state="disabled", text="Revisando…")
        self.progreso.pack(fill="x", side="top")
        self.progreso.configure(maximum=len(self.archivos), value=0)
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")
        self.escribir(f"Tipo: {opciones.tipo}   ·   {len(self.archivos)} documento(s)", "suave")

        self._guardar_config()
        threading.Thread(target=self._trabajar, args=(list(self.archivos), reglas, opciones, salida),
                         daemon=True).start()

    def _trabajar(self, archivos, reglas, opciones, salida):
        registro = revisor.leer_registro(opciones.registro)
        hechos = 0
        for ruta in archivos:
            nombre = os.path.basename(ruta)
            self.cola.put(("estado", f"Revisando {nombre}…"))
            try:
                mensajes = []
                res = revisor.revisar(ruta, reglas, salida, registro, opciones,
                                      opciones.forzar, mensajes.append)
                if res is None:
                    self.cola.put(("log", (f"{nombre}: no parece un {reglas['tipo']}. No se generó reporte.", "err")))
                else:
                    hechos += 1
                    for m in mensajes:
                        cuerpo = m.split(" -> ")[0]
                        self.cola.put(("log", (cuerpo, "ok" if " 0 errores" in cuerpo else None)))
            except Exception as ex:
                import traceback
                tb = traceback.format_exc()
                print(tb)  # Para la consola
                self.cola.put(("log", (f"{nombre}: no se pudo revisar ({type(ex).__name__}: {ex})", "err")))
                self.cola.put(("log", (tb, "err")))
            self.cola.put(("avance", 1))
        self.cola.put(("fin", (hechos, salida)))

    def _vaciar_cola(self):
        try:
            while True:
                clase, valor = self.cola.get_nowait()
                if clase == "log":
                    texto, etiqueta = valor
                    self.escribir(texto, etiqueta)
                elif clase == "estado":
                    self.estado.configure(text=valor)
                elif clase == "avance":
                    self.progreso.step(valor)
                elif clase == "fin":
                    hechos, salida = valor
                    self.trabajando = False
                    self.boton.configure(state="normal", text="Revisar")
                    self.progreso.pack_forget()
                    self.estado.configure(text=f"Terminado. {hechos} reporte(s) generado(s).")
                    self.escribir(f"Reportes en: {salida}", "suave")
                    if hechos and self.abrir.get():
                        self.abrir_carpeta(salida)
        except queue.Empty:
            pass
        self.after(120, self._vaciar_cola)


def main():
    App().mainloop()


if __name__ == "__main__":
    main()
