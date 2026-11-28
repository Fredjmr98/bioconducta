# Columnas: guía del editor

Este es el procedimiento para publicar una columna. Todo se hace desde la carpeta principal del sitio.

## Preparación (una sola vez)

1. Instala Python 3 si no lo tienes.
2. Instala la librería de lectura de Word:  `pip install python-docx`
   (opcional: `pip install pillow`, para reducir imágenes muy grandes).
3. Abre `tools/convertir_columna.py` y revisa el bloque «configuración» del principio:
   dirección del sitio (`SITIO_URL`), licencia (`LICENCIA`), rango de palabras, etc.

## Publicar una columna

1. **Recibe el documento.** La persona autora te envía el `.docx` hecho con `columnas/plantilla-columna.docx`.
   Si usó Google Docs, debe descargarlo con *Archivo → Descargar → Microsoft Word (.docx)*.
2. **Revisa lo formal:**
   `python tools/convertir_columna.py entrega.docx --solo-revisar`
   El informe lista **errores** (impiden publicar: campos vacíos, citas sin referencia, textos de la plantilla sin
   sustituir, ecuaciones del editor de Word) y **avisos** (longitud, referencias sin citar, notas al pie...).
   No escribe nada. Devuelve el documento a la persona autora con lo que haya que corregir.
3. **Edición de contenido.** Comenta el texto con la persona autora (comentarios de Word o de Docs) hasta aceptarlo.
   Conserva en una carpeta propia el envío original, tus comentarios y la versión aceptada.
4. **Convierte:**
   `python tools/convertir_columna.py entrega.docx --recibida 2026-11-03 --aceptada 2026-11-20 --publicada 2026-11-27`
   Crea `columnas/<titulo>.html`, copia las figuras a `columnas/<titulo>/`, numera la columna y actualiza
   `columnas.html` y `columnas/catalogo.json`. Si no pones `--publicada`, usa la fecha de hoy.
   Puedes elegir el nombre del archivo con `--slug nombre-corto`.
5. **Prueba de página.** Abre la página en el navegador, revísala y envíale el enlace (o el archivo) a la persona
   autora para que apruebe la versión final.
6. **Sube los cambios a GitHub** (`columnas/`, `columnas.html`).

## Corregir una columna ya publicada

Corrige el `.docx`, y ejecuta de nuevo el comando con `--actualizar` (conserva el número y las fechas).
Añade al final de la columna una nota de corrección fechada, como indica la política editorial.

## Formato que entiende el conversor

| En el documento                                    | En la página                              |
|----------------------------------------------------|-------------------------------------------|
| `Título:`, `Autores:`, `Resumen:`… (primeras líneas) | Encabezado, autoría, resumen, declaraciones |
| Título 1 / Título 2 / Título 3                       | Apartados (h2 / h3 / h4)                  |
| `[@brusca2023]`, `[@brusca2023, p. 44]`, `[@a; @b]` | Cita (año enlazado a la referencia)       |
| `Brusca ([-@brusca2023])`                           | Cita narrativa (solo el año)              |
| `[clave] Apellido, N. (Año)…` bajo «Referencias»     | Lista de referencias (APA, con clave)     |
| Imagen en su párrafo + «Figura N. Texto»            | Figura con leyenda                        |
| «Tabla N. Texto» + tabla                             | Tabla con leyenda                         |
| `$…$`, `$$…$$`, `\begin{equation}…\end{equation}`    | Ecuaciones (MathJax)                      |
| Párrafo que empieza con `> ` o estilo «Cita»        | Cita textual larga                        |
| Listas con viñetas o numeradas (hasta 4 niveles)     | `<ul>` / `<ol>`                           |

No se convierten: notas al pie, ecuaciones del editor de Word, cuadros de texto, formatos de imagen distintos de
PNG/JPG/GIF/SVG/WebP. El conversor avisa si los encuentra.

## La página de ejemplo

`columnas/ejemplo-de-columna.html` se generó con `tools/ejemplos/ejemplo-entrega.docx` y la opción `--ejemplo`
(lleva un aviso y no aparece en el listado). Puedes borrar la página y la carpeta `columnas/ejemplo-de-columna/`
cuando ya no la necesites.
