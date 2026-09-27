# Revisor de Tesis · FINESI

![Python](https://img.shields.io/badge/Python-3.10+-blue.svg)
![Arquitectura](https://img.shields.io/badge/Arquitectura-Clean_Architecture-success.svg)
![Plataforma](https://img.shields.io/badge/Plataforma-Windows-lightgrey.svg)

Aplicación de escritorio desarrollada para la Dirección de Investigación de FINESI. Revisa documentos de Word (`.docx`) y genera una hoja de observaciones formal (también en Word), detallando la página y línea de cada error, con espacio para la firma del revisor.

> **Nota:** El sistema revisa formato, estructura, citas, ortografía y el informe de similitud de Turnitin. No evalúa el contenido ni la coherencia metodológica.

---

## Arquitectura del Sistema

El proyecto ha sido rediseñado utilizando **Clean Architecture** y el patrón **src/ layout** para garantizar escalabilidad y fácil mantenimiento:

* **`src/app.py`** (Presentación): Interfaz gráfica en Tkinter / ttkbootstrap.
* **`src/core/`** (Dominio): El motor de validación. `orquestador.py` coordina, `utils.py` guarda las piezas compartidas (normalización, ubicación, estilos de Word) y cada submódulo revisa un aspecto: `estructura.py`, `formato.py`, `citas.py`, `ortografia.py`, `extension.py`, `turnitin.py`.
* **`src/reportes/`** (Infraestructura): Exportadores a Word (`reporte_resumen.py`, `reporte_word.py`) y a Excel (`reporte_excel.py`).
* **`src/utilidades/`**: Mapeo de líneas y lectura de reglas.
* **`data/`**: Reglas YAML, plantillas oficiales, diccionarios, y tu archivo `mis_reglas.yaml` y `registro.csv`.
* **`scripts/`**: Herramientas CLI: el generador automático de reglas (`generar_reglas.py`) y la medición de calidad contra un corpus (`medir_calidad.py`).
* **`tests/`**: Tests de caracterización (instantáneas sobre `ejemplos/`), del armado de la corrida, de la revisión de similitud y de la interfaz (ver sección 7).

Las dependencias van en una sola dirección: `app` → `core` → `reportes`/`utilidades`. Ningún módulo de `core/` importa a otro con `import *`, para que siempre se pueda ver de dónde sale cada nombre.

---

## 1. Crear el Ejecutable (Windows)

1. Instala **Python 3.10 o superior** desde python.org. Es **crucial** marcar la casilla **Add Python to PATH**.
2. Haz doble clic en **`construir_exe.bat`**. Este script limpiará el entorno, instalará dependencias y compilará la aplicación.
3. El resultado quedará en **`dist\RevisorTesisFINESI.exe`**.

Este `.exe` es completamente autónomo. Puedes copiarlo a cualquier PC con Windows (ej. `C:\RevisorTesis`) y funcionará sin tener Python instalado. Así, suelto, es **portable**: guarda sus datos (reglas, `registro.csv`, reportes, configuración) en la misma carpeta que el `.exe`.

### Instalador (recomendado para entregar)

1. Instala **Inno Setup 6** desde jrsoftware.org/isdl.php.
2. Haz doble clic en **`construir_instalador.bat`**. Construye el `.exe` y lo empaqueta.
3. El resultado queda en **`instalador\salida\RevisorTesisFINESI-Setup-<versión>.exe`**.

El instalador crea el acceso directo en el menú Inicio (y, si se elige, en el escritorio), aparece en *Aplicaciones instaladas* para desinstalarlo, y no pide permisos de administrador: se instala para el usuario actual. Instalado, el programa guarda sus datos en **`Documentos\Revisor de Tesis FINESI`**, porque en la carpeta de instalación Windows no deja escribir. Desinstalar no borra esa carpeta.

Para una versión nueva, cambia `Version` en `instalador\RevisorTesisFINESI.iss`; instalarla encima reemplaza la anterior y conserva los datos.

> **Dependencias Externas Recomendadas:**
> * **Java 17+** (adoptium.net): Necesario para revisar *gramática* con LanguageTool. Si falta, solo se revisa ortografía.
> * **LibreOffice**: Necesario para calcular con precisión la página y línea de cada error.
>
> Para leer el informe de similitud de Turnitin no hace falta nada extra: el PDF se lee con `pypdf`, que ya viene incluido.

---

## 2. Uso del Programa

1. Abre el `.exe`.
2. Haz clic en **Agregar archivos** o **Agregar carpeta** para cargar los `.docx`. Si el tesista adjuntó el informe de similitud de Turnitin en `.pdf`, agrégalo también: se empareja solo con su documento y sus observaciones salen en la misma hoja (ver sección 5).
3. Selecciona el tipo de documento: **Proyecto** o **Borrador**. (Si te equivocas, el sistema lo detectará automáticamente por los títulos y te avisará).
4. Completa la ficha: tesista, asesor, fecha. Si revisas por lotes, esto se carga automáticamente desde `data/registro.csv`.
5. Coloca tu nombre en **Revisor (firma)**.
6. Clic en **Revisar**. Los reportes terminados se guardarán en `data/reportes/`.

---

## 3. Formato de los Reportes

Elige el formato de salida según tu necesidad:
- **Hoja de revisión (corta):** Reproduce el formato oficial PGI. Da un resumen tipo "Línea 33 y otros: corregir formato APA".
- **Detallado:** Tabla exhaustiva con cada uno de los hallazgos aislados, su ubicación exacta y la corrección sugerida.

**Observaciones Adicionales:** Puedes escribir a mano en la app cosas que el software no detecta ("Falta matriz de consistencia"). Quedan guardadas para tu próxima sesión.

---

## 4. Reglas Propias (mis_reglas.yaml)

No necesitas programar para enseñarle cosas nuevas al sistema. Haz clic en **Editar reglas** en la app para abrir `data/mis_reglas.yaml` y agrega las tuyas:

```yaml
reglas:
  - observacion: "En anexos, incluir matriz de consistencia"
    donde: documento
    si_no_aparece: ["matriz de consistencia"]
    severidad: Error

  - observacion: "Incrementar número de citas (antecedentes)"
    donde: "Antecedentes del proyecto"
    minimo_citas: 5
```

### Ajustar la severidad de una categoría

Si una categoría te está dando falsos positivos y no querés que salga como error, bajale la severidad en `data/reglas/<tipo>.yaml` sin tocar código:

```yaml
severidades:
  Citas: Advertencia
  Ortografía: Revisar
```

Valores válidos: `Error`, `Advertencia`, `Revisar`. Aplica a cualquier categoría: `Estructura`, `Formato`, `Extensión`, `Citas`, `Similitud`, `Ortografía`.

### Topes del informe de Turnitin

Los umbrales de similitud son una decisión de la Dirección de Investigación, no del programa. Están en `data/reglas/<tipo>.yaml`, bajo `similitud:`, y **hay que confirmarlos contra el reglamento vigente de la FINESI** antes de usarlos en serio:

```yaml
similitud:
  max_pct: 25                  # índice máximo admitido (null = no se revisa)
  margen_alerta_pct: 5         # avisa cuando queda a menos de esto del tope
  max_fuente_pct: 10           # lo máximo que puede aportar una sola fuente
  max_ia_pct: null             # texto detectado como generado por IA (null = no se revisa)
  filtros_exigidos: []         # ["bibliografia", "citas"] si el informe debe excluirlos al medir
  max_antiguedad_dias: null    # antigüedad máxima del informe (null = no se revisa)
  tolerancia_palabras_pct: 15  # diferencia admitida entre el .docx y lo que midió Turnitin
  exigir_informe: false        # true = observar los documentos que llegan sin informe
```

Poner `null` en un tope desactiva esa comprobación. `exigir_informe: true` hace que cada `.docx` que llegue sin su PDF salga observado; déjalo en `false` mientras el informe siga siendo opcional en el trámite.

> La detección de IA de Turnitin es referencial y se equivoca. El programa la reporta como **advertencia** y con esa aclaración escrita en el detalle, nunca como error.

---

## 5. ¿Qué revisa exactamente el Core?

- **Estructura**: Secciones obligatorias, orden correcto, secciones vacías.
- **Plantilla**: Texto guía entre paréntesis `(...)` que el tesista olvidó borrar.
- **Formato**: Tamaño de papel, márgenes institucionales, fuente, tamaño, interlineado, justificado.
- **Extensión**: Límite de páginas, conteo de palabras del título y palabras clave.
- **Citas APA 7**: Citas huérfanas (sin referencia al final), discrepancia de años, referencias no citadas, orden alfabético, sangría francesa.
- **Similitud (Turnitin)**: Lee el PDF del informe y compara el índice contra el tope, revisa cuánto aporta cada fuente por separado, el porcentaje de IA, con qué filtros se generó y su antigüedad.
- **Ortografía, Gramática y Tipografía**: Revisa errores ortográficos, dobles espacios, comas mal espaciadas.

### Cómo reconoce las secciones

No se fía solo del parecido del texto. Un párrafo se toma por título cuando el formato lo confirma —estilo de título de Word, numeración, negrita o mayúsculas— y en ese caso se acepta un parecido menor. Así entran títulos legítimos como «FUENTES BIBLIOGRÁFICAS», que por texto puro quedaban fuera.

Cuando una sección obligatoria no se reconoce, en vez de afirmar que falta, el reporte nombra el título más parecido del documento y te dice qué agregar como `alias`. Esto importa porque el modo de fallo peligroso no es el ruido: es el silencio. Si no se identifica «Referencias», la verificación de citas no reporta nada y el documento parece limpio.

### Cómo se empareja el informe de Turnitin con su documento

El PDF se empareja con su `.docx` por el nombre del archivo y por el nombre del documento que el propio informe declara adentro. Si el PDF se llama solo `Informe Turnitin.pdf` y hay un único documento en la lista, va con ese; si hay varios, **no se adivina**: poner el índice de similitud de un tesista en la hoja de otro es peor que no ponerlo, así que cada informe suelto sale en su propia hoja `Similitud - <nombre>.docx`.

El parecido se mide por palabras, no letra por letra: `PROYECTO Cruz Apaza.docx` empareja con `Turnitin - Cruz Apaza.pdf` aunque el rótulo no coincida. Pero **un solo apellido compartido no alcanza**: `Quispe Mamani` y `Quispe` no se emparejan, porque pueden ser dos tesistas distintos.

Como último control, el programa compara las palabras que declara el informe con las del `.docx`. Si difieren más de `tolerancia_palabras_pct`, avisa que el informe probablemente sea de otra versión del documento o de otro trabajo.

Del PDF **solo** se revisa la similitud. El formato, los márgenes y los estilos se revisan sobre el original de Word: un PDF no los conserva. Si llega la tesis exportada a PDF en lugar del informe, el programa lo detecta y pide el `.docx`.

Para ver exactamente qué leyó de un informe (útil cuando Turnitin cambia el diseño del PDF):

```powershell
$env:PYTHONPATH="src"
python -m core.turnitin "ruta\del\informe.pdf"
```

### Cuando un chequeo falla

Cada chequeo corre aislado. Si uno revienta, el resto del documento se revisa igual y en la hoja aparece una línea diciendo qué quedó sin verificar. Antes, un error en cualquier módulo descartaba la revisión completa.

---

## 6. Uso por Línea de Comandos (CLI)

Para usuarios avanzados, el orquestador funciona por terminal:

```powershell
# Definir el PYTHONPATH es crucial con la nueva arquitectura
$env:PYTHONPATH="src"

# Revisar un archivo
python -m core.orquestador --tipo proyecto "ejemplos/tesis.docx"

# Revisar una carpeta completa a tu nombre
python -m core.orquestador --tipo proyecto "ejemplos/" --revisor "Tu Nombre"

# Con el informe de Turnitin. Si el PDF está en la misma carpeta y se llama
# parecido al .docx, se toma solo y no hace falta nombrarlo:
python -m core.orquestador --tipo proyecto "tesis.docx" --turnitin "Turnitin - tesis.pdf"

# Al revisar una carpeta, cada .docx se empareja con su PDF automáticamente
python -m core.orquestador --tipo proyecto "expedientes/" --revisor "Tu Nombre"
```

`--turnitin` solo vale al revisar un documento. Si ese PDF no se puede leer, la corrida falla con un mensaje: se pidió ese archivo a propósito y cambiarlo en silencio por otro sería peor.

## 7. Tests

```powershell
python -m pip install pytest
python -m pytest tests/
```

(No hace falta definir `PYTHONPATH`: `tests/conftest.py` se encarga.)

No hacen falta Java ni LibreOffice: la ortografía se fuerza al diccionario hunspell para que el resultado sea igual en cualquier máquina, y las instantáneas no incluyen la columna de página y línea, que sí depende de LibreOffice.

Son cinco archivos, con propósitos distintos:

| Archivo | Qué cubre |
|---|---|
| `test_caracterizacion.py` | **Instantáneas** sobre los documentos de `ejemplos/`: congelan lo que el motor observa hoy, no lo que *debería* observar. Cualquier cambio que altere el resultado salta de inmediato. |
| `test_orquestador.py` | El armado de la corrida: qué archivos entran, cómo se empareja cada `.docx` con su informe de Turnitin, cómo se nombran las hojas, qué se rechaza antes de empezar. |
| `test_similitud.py` | Qué se observa del informe según los topes de las reglas. Arma el `Informe` ya leído, que es la frontera entre *leer el PDF* y *decidir qué observar*. |
| `test_app.py` | Que el botón **Revisar** pase por el mismo camino que la línea de comandos. Existe por una regresión concreta: el emparejado con Turnitin estaba escrito y probado por CLI, pero la interfaz lo salteaba y en el `.exe` no hacía nada. |
| `test_lectura_word.py` | Las dos trampas de python-docx que producían observaciones falsas sobre documentos correctos: la sangría francesa definida en el estilo (no en el párrafo) y la celda combinada que `row.cells` devuelve repetida. |

Si un cambio modifica el comportamiento a propósito, se regeneran las instantáneas:

```powershell
$env:REGENERAR_INSTANTANEAS="1"; python -m pytest tests/
```

Revisa el diff de `tests/instantaneas/` antes de confirmarlo: ahí se ve exactamente qué observaciones se ganaron o se perdieron.

> La lectura de los PDF reales de Turnitin no se testea automáticamente: los informes son documentos de estudiantes y no se versionan. Se comprueba a mano con `python -m core.turnitin "informe.pdf"`.

---

## 8. Medir si el motor acierta

Los tests congelan el comportamiento: avisan si algo cambia. Eso es distinto de saber si el motor **acierta**. Los umbrales internos (el parecido de títulos, el emparejado de autores) solo se pueden ajustar con evidencia, y para eso hace falta un corpus de tesis ya revisadas a mano.

```powershell
# 1. Junta en una carpeta los .docx que ya revisaste
# 2. Genera las fichas de partida con lo que el motor encontró
python scripts/medir_calidad.py corpus/ --tipo proyecto --crear-fichas

# 3. Edita cada corpus/<nombre>.esperado.yaml: borra los falsos positivos
#    y agrega lo que el motor no vio. Eso es la verdad de referencia.

# 4. Mide
python scripts/medir_calidad.py corpus/ --tipo proyecto --detalle
```

Reporta por categoría los aciertos, los falsos positivos y las **no detectadas**, más precisión y exhaustividad. Las no detectadas son el número que más importa: son los errores que el revisor no va a ver y va a dar por buenos.

La carpeta `corpus/` está en `.gitignore`: son documentos reales de estudiantes y no deben versionarse.

---

##  9. Límites Conocidos
- Solo revisa `.docx`. De un `.pdf` únicamente lee el informe de similitud de Turnitin; el formato necesita el original de Word.
- La lectura del informe depende del diseño del PDF que entrega Turnitin. Si cambia el diseño, el programa avisa que no pudo leer el índice y manda a verificarlo a mano en vez de darlo por bueno. Para ver qué leyó: `python -m core.turnitin "informe.pdf"`.
- Los topes de similitud que vienen configurados son un punto de partida, **no** el reglamento de la FINESI. Confirmarlos antes de usarlos.
- La línea exacta del error depende del mapeo interno con LibreOffice, puede tener un margen de error mínimo si el tesista desconfiguró los márgenes.
- Detección de citas usa Expresiones Regulares (RegEx); apellidos muy inusuales compuestos podrían generar falsos positivos. Por ello, "referencia no citada" es una advertencia, no un error fatal.
