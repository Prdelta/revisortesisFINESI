# Revisor de Tesis · FINESI

![Python](https://img.shields.io/badge/Python-3.10+-blue.svg)
![Arquitectura](https://img.shields.io/badge/Arquitectura-Clean_Architecture-success.svg)
![Plataforma](https://img.shields.io/badge/Plataforma-Windows-lightgrey.svg)

Aplicación de escritorio desarrollada para la Dirección de Investigación de FINESI. Revisa documentos de Word (`.docx`) y genera una hoja de observaciones formal (también en Word), detallando la página y línea de cada error, con espacio para la firma del revisor.

> **Nota:** El sistema revisa formato, estructura, citas y ortografía. No evalúa el contenido ni la coherencia metodológica.

---

## Arquitectura del Sistema

El proyecto ha sido rediseñado utilizando **Clean Architecture** y el patrón **src/ layout** para garantizar escalabilidad y fácil mantenimiento:

* **`src/app.py`** (Presentación): Interfaz gráfica en Tkinter / ttkbootstrap.
* **`src/core/`** (Dominio): El motor de validación. Contiene el orquestador y los submódulos especializados (`citas.py`, `estructura.py`, `formato.py`, `ortografia.py`, `extension.py`).
* **`src/reportes/`** (Infraestructura): Exportadores a Word.
* **`src/utilidades/`**: Mapeo de líneas y lectura de reglas.
* **`data/`**: Reglas YAML, plantillas oficiales, diccionarios, y tu archivo `mis_reglas.yaml` y `registro.csv`.
* **`scripts/`**: Herramientas CLI como el generador automático de reglas (`generar_reglas.py`).

---

## 1. Crear el Ejecutable (Windows)

1. Instala **Python 3.10 o superior** desde python.org. Es **crucial** marcar la casilla **Add Python to PATH**.
2. Haz doble clic en **`construir_exe.bat`**. Este script limpiará el entorno, instalará dependencias y compilará la aplicación.
3. El resultado quedará en **`dist\RevisorTesisFINESI.exe`**.

Este `.exe` es completamente autónomo. Puedes copiarlo a cualquier PC con Windows (ej. `C:\RevisorTesis`) y funcionará sin tener Python instalado. 

> **Dependencias Externas Recomendadas:**
> * **Java 17+** (adoptium.net): Necesario para revisar *gramática* con LanguageTool. Si falta, solo se revisa ortografía.
> * **LibreOffice**: Necesario para calcular con precisión la página y línea de cada error.

---

## 2. Uso del Programa

1. Abre el `.exe`.
2. Haz clic en **Agregar archivos** o **Agregar carpeta** para cargar los `.docx`.
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

---

## 5. ¿Qué revisa exactamente el Core?

- **Estructura**: Secciones obligatorias, orden correcto, secciones vacías.
- **Plantilla**: Texto guía entre paréntesis `(...)` que el tesista olvidó borrar.
- **Formato**: Tamaño de papel, márgenes institucionales, fuente, tamaño, interlineado, justificado.
- **Extensión**: Límite de páginas, conteo de palabras del título y palabras clave.
- **Citas APA 7**: Citas huérfanas (sin referencia al final), discrepancia de años, referencias no citadas, orden alfabético, sangría francesa.
- **Ortografía, Gramática y Tipografía**: Revisa errores ortográficos, dobles espacios, comas mal espaciadas.

---

## 6. Uso por Línea de Comandos (CLI)

Para usuarios avanzados, el orquestador funciona por terminal:

```bash
# Definir el PYTHONPATH es crucial con la nueva arquitectura
$env:PYTHONPATH="src"

# Revisar un archivo
python -m core.orquestador --tipo proyecto "ejemplos/tesis.docx"

# Revisar una carpeta completa a tu nombre
python -m core.orquestador --tipo proyecto "ejemplos/" --revisor "Tu Nombre"
```

##  7. Límites Conocidos
- Solo lee `.docx`.
- La línea exacta del error depende del mapeo interno con LibreOffice, puede tener un margen de error mínimo si el tesista desconfiguró los márgenes.
- Detección de citas usa Expresiones Regulares (RegEx); apellidos muy inusuales compuestos podrían generar falsos positivos. Por ello, "referencia no citada" es una advertencia, no un error fatal.
