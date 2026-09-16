# Revisor de Tesis · FINESI

Aplicación de escritorio para la Dirección de Investigación. Revisa documentos de Word (.docx) y
genera una hoja de observaciones, también en Word, con los datos del expediente, la tabla de
observaciones con página y línea de cada error, y espacio para la firma del revisor.

Revisa formato, estructura, citas y ortografía. No evalúa el contenido ni la metodología.
El reporte es un borrador: cada observación se valida antes de comunicarla al tesista.

Proyecto de tesis y borrador de tesis usan reglas separadas y nunca se mezclan.

---

## 1. Crear el ejecutable (una sola vez, en una PC con Windows)

1. Instala Python 3.10 o superior desde python.org. Marca la casilla **Add Python to PATH**.
2. Doble clic en **construir_exe.bat**. Instala lo necesario y arma el programa.
3. Al terminar queda **dist\RevisorTesisFINESI.exe**.

Ese .exe es el programa. Se copia a cualquier PC con Windows y funciona sin Python instalado.
Cópialo a una carpeta propia, por ejemplo `C:\RevisorTesis`, y crea un acceso directo en el
escritorio. La primera vez que se abre crea junto a él las carpetas `reglas`, `plantillas`,
`reportes` y los archivos `permitidas.txt` y `registro.csv`, que son los que se pueden editar.

### Programas que conviene tener en la PC donde se usa

| Programa | Para qué | Si falta |
|---|---|---|
| Java 17 o superior (adoptium.net) | revisar gramática con LanguageTool | solo revisa ortografía con el diccionario incluido |
| LibreOffice (libreoffice.org) | calcular página y línea de cada error, contar páginas | las observaciones salen sin número de línea |

La primera revisión con gramática descarga LanguageTool, unos 260 MB. Ocurre una sola vez.

## 2. Usar el programa

1. Abre el .exe.
2. **Agregar archivos** o **Agregar carpeta** para elegir los .docx.
3. Elige el tipo: Proyecto o Borrador.
4. Si revisas un solo documento, escribe tesista, asesor y fecha. Si son varios, esos datos salen
   de `registro.csv`.
5. Escribe tu nombre en **Revisor (firma)**. Queda guardado para la próxima vez.
6. **Revisar**. Los reportes van a la carpeta `reportes`, uno por documento.

Si eliges el tipo equivocado, el programa lo detecta por los títulos, no genera el reporte y avisa.

## 3. Formato del reporte

Hay dos formatos y se eligen en la ventana:

- **Hoja de revisión (corta)**, el que se usa por defecto. Reproduce el formato de la Dirección:
  título REVISIÓN PGI con fecha y número, ficha del proyecto con su código, asesor y una lista
  corta de observaciones del tipo "Línea 33 y otros: corregir la forma de citar, estilo APA 7ª ed".
  Cada observación apunta a la primera línea donde aparece el problema.
- **Detallado**, una tabla con cada hallazgo por separado, su ubicación exacta y la corrección.
  Sirve cuando el tesista pide el detalle o cuando se quiere revisar qué encontró el programa.

En **Observaciones adicionales** se escriben las que el programa no puede detectar, una por línea,
por ejemplo "En anexos, incluir matriz de consistencia". Quedan guardadas para la próxima vez.

## 4. Datos del expediente

El reporte lleva código del expediente, tesista, título, asesor, fecha y número de revisión. De
esos datos, el título sale del propio documento y la fecha de presentación de `registro.csv` o, si
no está, de la última fecha de guardado del archivo Word. El código, el tesista y el asesor solo
pueden venir de ti.

Para revisar varios documentos a la vez, llena `registro.csv` (se abre con Excel, separador punto
y coma):

    archivo;codigo;tesista;proyecto;asesor;fecha_subida;n_revision;etapa
    tesis_perez.docx;P26-183522A;Juan Pérez Mamani;;ELQUI YEYE PARI C.;05/09/2026;1;Etapa 2: Revisión de formato

Lo que falte aparece como "(completar)" en gris en el Word.

## 5. Reglas propias del revisor

Además de lo que trae el programa, cada revisor puede escribir sus propios criterios en
`mis_reglas.yaml` (botón **Editar reglas** en la ventana). Se aplican en cada revisión y no hay
que tocar el código ni reconstruir el .exe.

Una regla se escribe así:

    reglas:
      - observacion: "En anexos, incluir matriz de consistencia"
        donde: documento
        si_no_aparece: ["matriz de consistencia"]
        severidad: Error

      - observacion: "Incrementar número de citas (antecedentes)"
        donde: "Antecedentes del proyecto"
        minimo_citas: 5

Condiciones disponibles: `si_aparece` (una lista de textos; si alguno aparece, se observa),
`si_no_aparece` (si ninguno aparece, se observa), `minimo_citas` y `minimo_palabras`.
El campo `donde` es "documento" o el nombre exacto de una sección del esquema. Con `activa: false`
se desactiva una regla sin borrarla.

Cuando la condición apunta a un texto concreto, la observación sale con su número de línea, igual
que las demás.

## 6. Configurar el borrador de tesis

Las reglas del borrador vienen vacías porque su plantilla es distinta a la del proyecto:

1. Copia la plantilla oficial del borrador como `plantillas/borrador.docx`.
2. Ejecuta: `python generar_reglas.py plantillas/borrador.docx --tipo borrador`
3. Se crea `reglas/borrador.yaml` con las secciones, subsecciones, márgenes, fuente e interlineado
   leídos de la plantilla. Ábrelo, compáralo con el reglamento y completa lo marcado COMPLETAR.
4. Prueba con un borrador que ya hayas revisado a mano.
5. Vuelve a construir el .exe para que lo incluya, o copia ese `borrador.yaml` a la carpeta
   `reglas` que está junto al ejecutable.

## 7. Qué revisa

- **Estructura**: secciones del esquema, orden, secciones vacías, subsecciones, tablas llenas.
- **Plantilla**: texto guía entre paréntesis que el tesista no borró.
- **Formato**: tamaño de papel, márgenes, encabezado institucional, fuente, tamaño, interlineado,
  justificado.
- **Extensión**: páginas, palabras del título, cantidad de palabras clave.
- **Citas APA 7**: citas sin referencia, año que no coincide, referencias no citadas, orden
  alfabético, sangría francesa, "et al." sin punto, citas numéricas mezcladas, DOI antiguo.
- **Ortografía y gramática**, más tipografía (doble espacio, espacio antes de coma).

### Página y línea

Si el documento tiene activada la numeración de líneas de Word, la plantilla oficial del proyecto
la trae, se usan esos mismos números: los que el tesista ve en el margen de su documento. Si la
desactivó, se cuentan las líneas desde el inicio de cada página y así se indica. Las observaciones
del documento completo (márgenes, sección faltante) y las de dentro de tablas no llevan línea;
ahí la columna Ubicación trae el texto exacto para buscarlo con Ctrl+B.

## 8. Ajustes

| Situación | Qué hacer |
|---|---|
| Cambió el reglamento | editar `reglas/proyecto.yaml` o `reglas/borrador.yaml` |
| Marca un término técnico o un apellido como error | agregarlo a `permitidas.txt`, uno por línea |
| Un tesista titula una sección de otra forma válida | agregar el alias en el .yaml |

### Los alias son lo que más mejora el revisor

El programa reconoce una sección por su título. Si no lo reconoce, reporta que falta esa sección y,
en el caso de Referencias, no puede cruzar las citas del texto con la lista de referencias: en ese
caso lo dice explícitamente en el reporte, no lo calla.

Ya tolera títulos más largos que el esperado, "Referencias bibliográficas" se reconoce como
"Referencias", pero no adivina sinónimos. Cuando veas un título nuevo, agrégalo en minúscula y sin
tildes:

    - {nombre: "Referencias", alias: ["referencias bibliograficas", "bibliografia", "fuentes consultadas"]}

Cada alias que agregas evita un falso "falta la sección" para siempre. Es la forma concreta de
"entrenar" este programa: no aprende solo, aprende de lo que tú escribes en estos archivos.

Esos archivos están junto al ejecutable y se editan con el Bloc de notas. No hay que reconstruir
el .exe para cambiarlos.

## 9. Archivos del proyecto

    app.py                   la aplicación (interfaz)
    revisor.py               el motor de revisión; también funciona por línea de comandos
    reglas_revisor.py        aplica las reglas propias del revisor
    mis_reglas.yaml          esas reglas, editables
    reporte_resumen.py       arma la hoja de revisión corta
    reporte_word.py          arma el reporte detallado
    localizador.py           calcula la página y línea de cada observación
    generar_reglas.py        crea un archivo de reglas a partir de una plantilla
    reglas/                  reglas de cada tipo de documento
    plantillas/              plantillas oficiales, de ahí salen las marcas guía a detectar
    dic/                     diccionario español (Perú) para cuando no hay Java
    permitidas.txt           palabras que el corrector no debe marcar
    registro.csv             datos de los expedientes
    ejemplos/                un proyecto correcto y otro con errores para probar
    construir_exe.bat        construye el ejecutable
    abrir_app.bat            abre la aplicación sin construir el .exe (requiere Python)
    instalar.bat             instala solo las dependencias
    RevisorTesisFINESI.spec  receta de PyInstaller

## 10. Línea de comandos (opcional)

    python revisor.py --tipo proyecto archivo.docx
    python revisor.py --tipo proyecto carpeta/ --revisor "Tu Nombre"
    python revisor.py --tipo proyecto archivo.docx --formato detallado
    python revisor.py --tipo proyecto archivo.docx --codigo P26-183522A --n-revision 1 \
        --asesor "ELQUI YEYE PARI C." --extra "En anexos, incluir matriz de consistencia"

## 11. Límites conocidos

- Solo .docx. Un .doc se abre en Word y se guarda como .docx. Los PDF no se revisan.
- El conteo de páginas y líneas sale de LibreOffice y puede diferir un poco de lo que muestra Word.
- La detección de citas usa patrones: apellidos compuestos poco comunes o instituciones con nombre
  largo pueden dar falsos positivos. Por eso "referencia no citada" es advertencia y no error.
- El corrector marca términos técnicos correctos hasta que se agregan a `permitidas.txt`.
- No evalúa contenido ni coherencia metodológica.
