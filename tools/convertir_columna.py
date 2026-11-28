#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Convierte una entrega (.docx) de la sección Columnas en una página del sitio.

Uso (desde la carpeta principal del sitio):
    pip install python-docx            # solo la primera vez
    python tools/convertir_columna.py entrega.docx
    python tools/convertir_columna.py entrega.docx --recibida 2026-11-03 --aceptada 2026-11-20 --publicada 2026-11-27
    python tools/convertir_columna.py entrega.docx --solo-revisar       # solo informe, no escribe nada
    python tools/convertir_columna.py entrega.docx --actualizar         # reemplaza una columna ya publicada

Qué hace:
  1. Lee los metadatos y el texto del documento (títulos, cursivas, listas, tablas, figuras, hipervínculos).
  2. Convierte las citas [@clave] y la lista de referencias al formato de citations.js.
  3. Revisa las normas formales (citas sin referencia, campos vacíos, longitud...) e imprime un informe.
  4. Escribe columnas/<slug>.html, copia las imágenes y actualiza columnas/catalogo.json y columnas.html.
"""
import argparse, datetime, html, io, json, re, sys, unicodedata
from pathlib import Path

try:
    from docx import Document
    from docx.oxml.ns import qn
except ImportError:
    sys.exit("Falta python-docx. Instálalo con:  pip install python-docx")

# ------------------------------------------------------------------ configuración
SITIO_URL = "https://bioconducta.blog"      # dirección del sitio (para «Cómo citar»)
SECCION = "Bioprâxis. Columnas"            # nombre que aparece en «Cómo citar»
LICENCIA = None                            # p. ej. "CC BY 4.0"; si es None no se muestra ninguna
CONJUNCION = "y"                           # «Brusca y Giribet (2023)»
LONGITUD = (800, 2500)                     # palabras del cuerpo (sin referencias); solo avisa
RESUMEN_MAX = 150                          # palabras
N_CLAVES = (3, 6)                          # palabras clave
ANCHO_MAX_IMG = 1800                       # px; las imágenes más anchas se reducen (si hay Pillow)
MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
         "septiembre", "octubre", "noviembre", "diciembre"]

RAIZ = Path(__file__).resolve().parent.parent

# ------------------------------------------------------------------ utilidades
def sin_tildes(s):
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")

def norm(s):
    return sin_tildes(s).lower().strip()

def esc(s):
    return html.escape(s, quote=False)

def attr(s):
    return html.escape(s, quote=True)

def fecha_larga(iso):
    d = datetime.date.fromisoformat(iso)
    return f"{d.day} de {MESES[d.month - 1]} de {d.year}"

def slugify(s):
    s = re.sub(r"[^a-z0-9]+", "-", sin_tildes(s).lower()).strip("-")
    return s[:70].strip("-") or "columna"

def palabras(s):
    return len(re.findall(r"\w+", s, flags=re.UNICODE))

def clave_ok(k):
    return sin_tildes(k).lower()

class Informe:
    def __init__(self):
        self.errores, self.avisos = [], []
    def error(self, m): self.errores.append(m)
    def aviso(self, m): self.avisos.append(m)

# ------------------------------------------------------------------ lectura del .docx
W_P, W_TBL, W_R, W_T = qn("w:p"), qn("w:tbl"), qn("w:r"), qn("w:t")
W_HYPER, W_INS, W_DEL, W_SDT = qn("w:hyperlink"), qn("w:ins"), qn("w:del"), qn("w:sdt")
M_OMATH = "{http://schemas.openxmlformats.org/officeDocument/2006/math}oMath"
M_OMATHPARA = "{http://schemas.openxmlformats.org/officeDocument/2006/math}oMathPara"

def iter_bloques(padre):
    for h in padre:
        if h.tag == W_P:
            yield "p", h
        elif h.tag == W_TBL:
            yield "tbl", h
        elif h.tag == W_SDT:
            c = h.find(qn("w:sdtContent"))
            if c is not None:
                yield from iter_bloques(c)

def val_on(el):
    if el is None:
        return False
    v = el.get(qn("w:val"))
    return v not in ("0", "false", "off")

class Seg:
    __slots__ = ("t", "b", "i", "sup", "sub", "href")
    def __init__(self, t, b=False, i=False, sup=False, sub=False, href=None):
        self.t, self.b, self.i, self.sup, self.sub, self.href = t, b, i, sup, sub, href
    def clave(self):
        return (self.b, self.i, self.sup, self.sub, self.href)

class Doc:
    def __init__(self, ruta, inf):
        self.doc = Document(ruta)
        self.inf = inf
        self.estilos = {s.style_id: (s.name or "").lower() for s in self.doc.styles}
        self.numeros = self._mapa_numeracion()
        self.rels = self.doc.part.rels
        self.cambios_marcados = False
        self.pie_notas = False

    def _mapa_numeracion(self):
        try:
            raiz = self.doc.part.numbering_part.element
        except Exception:
            return {}
        abstractos = {a.get(qn("w:abstractNumId")): a for a in raiz.findall(qn("w:abstractNum"))}
        mapa = {}
        for num in raiz.findall(qn("w:num")):
            aid = num.find(qn("w:abstractNumId"))
            niveles = {}
            a = abstractos.get(aid.get(qn("w:val"))) if aid is not None else None
            if a is not None:
                for l in a.findall(qn("w:lvl")):
                    f = l.find(qn("w:numFmt"))
                    niveles[int(l.get(qn("w:ilvl")))] = f.get(qn("w:val")) if f is not None else "decimal"
            mapa[num.get(qn("w:numId"))] = niveles
        return mapa

    # ---- propiedades de párrafo
    def estilo(self, p):
        ppr = p.find(qn("w:pPr"))
        if ppr is not None:
            ps = ppr.find(qn("w:pStyle"))
            if ps is not None:
                return self.estilos.get(ps.get(qn("w:val")), ps.get(qn("w:val"), "").lower())
        return "normal"

    def lista(self, p):
        """Devuelve (tipo, nivel) si el párrafo es un elemento de lista, o None."""
        est = self.estilo(p)
        ppr = p.find(qn("w:pPr"))
        numpr = ppr.find(qn("w:numPr")) if ppr is not None else None
        if numpr is not None:
            nid = numpr.find(qn("w:numId"))
            il = numpr.find(qn("w:ilvl"))
            nid = nid.get(qn("w:val")) if nid is not None else None
            nivel = int(il.get(qn("w:val"))) if il is not None else 0
            if nid and nid != "0":
                fmt = self.numeros.get(nid, {}).get(nivel, "decimal")
                return ("ul" if fmt in ("bullet", "none") else "ol", nivel)
            return None
        m = re.match(r"list (bullet|number)(?: (\d))?$", est)
        if m:
            return ("ul" if m.group(1) == "bullet" else "ol", int(m.group(2)) - 1 if m.group(2) else 0)
        return None

    # ---- segmentos de texto con formato
    def segmentos(self, p):
        segs = []
        def run(r, href):
            rpr = r.find(qn("w:rPr"))
            b = i = sup = sub = False
            if rpr is not None:
                b, i = val_on(rpr.find(qn("w:b"))), val_on(rpr.find(qn("w:i")))
                va = rpr.find(qn("w:vertAlign"))
                if va is not None:
                    sup, sub = va.get(qn("w:val")) == "superscript", va.get(qn("w:val")) == "subscript"
                rs = rpr.find(qn("w:rStyle"))
                if rs is not None:
                    nom = self.estilos.get(rs.get(qn("w:val")), "")
                    b = b or nom == "strong"
                    i = i or nom == "emphasis"
            for h in r:
                if h.tag == W_T:
                    segs.append(Seg(h.text or "", b, i, sup, sub, href))
                elif h.tag == qn("w:tab"):
                    segs.append(Seg(" ", b, i, sup, sub, href))
                elif h.tag == qn("w:br") and h.get(qn("w:type")) in (None, "textWrapping"):
                    segs.append(Seg("\n", b, i, sup, sub, href))
                elif h.tag == qn("w:noBreakHyphen"):
                    segs.append(Seg("-", b, i, sup, sub, href))
                elif h.tag == qn("w:footnoteReference") or h.tag == qn("w:endnoteReference"):
                    self.pie_notas = True
        def recorrer(el, href=None):
            for h in el:
                if h.tag == W_R:
                    run(h, href)
                elif h.tag == W_HYPER:
                    rid = h.get(qn("r:id"))
                    destino = None
                    if rid and rid in self.rels and self.rels[rid].is_external:
                        destino = self.rels[rid].target_ref
                    elif h.get(qn("w:anchor")):
                        destino = "#" + h.get(qn("w:anchor"))
                    recorrer(h, destino)
                elif h.tag == W_INS:
                    self.cambios_marcados = True
                    recorrer(h, href)
                elif h.tag == W_DEL:
                    self.cambios_marcados = True
                elif h.tag == W_SDT:
                    c = h.find(qn("w:sdtContent"))
                    if c is not None:
                        recorrer(c, href)
                elif h.tag in (qn("w:smartTag"), qn("w:customXml"), qn("w:fldSimple")):
                    recorrer(h, href)
        recorrer(p)
        # unir segmentos vecinos con el mismo formato
        out = []
        for s in segs:
            if out and out[-1].clave() == s.clave():
                out[-1].t += s.t
            else:
                out.append(Seg(s.t, s.b, s.i, s.sup, s.sub, s.href))
        return out

    def banderas(self, p):
        return {
            "dibujos": p.findall(".//" + qn("w:drawing")),
            "omml": bool(list(p.iter(M_OMATH)) or list(p.iter(M_OMATHPARA))),
            "cuadro": bool(list(p.iter(qn("w:txbxContent")))),
        }

def texto(segs):
    return "".join(s.t for s in segs)

def quitar_prefijo(segs, n):
    out, resto = [], n
    for s in segs:
        if resto >= len(s.t):
            resto -= len(s.t)
            continue
        out.append(Seg(s.t[resto:], s.b, s.i, s.sup, s.sub, s.href))
        resto = 0
    return out

def a_html(segs):
    partes = []
    for s in segs:
        t = esc(s.t).replace("\n", "<br>")
        if not t:
            continue
        if s.sub: t = f"<sub>{t}</sub>"
        if s.sup: t = f"<sup>{t}</sup>"
        if s.b: t = f"<b>{t}</b>"
        if s.i: t = f"<i>{t}</i>"
        if s.href: t = f'<a href="{attr(s.href)}">{t}</a>'
        partes.append(t)
    return "".join(partes).strip()

# ------------------------------------------------------------------ citas y referencias
CITA_RE = re.compile(r"\[((?:-?@[^\]\[;<>]+)(?:\s*;\s*-?@[^\]\[;<>]+)*)\]")
XREF_RE = re.compile(r"\b(Figuras?|Tablas?)(\s+\d+(?:\.\d+)?)")

class Contexto:
    def __init__(self):
        self.citadas = []          # claves en orden de aparición
        self.hay_mates = False

def procesar_citas(h, ctx, inf):
    def sub(m):
        partes = [x.strip() for x in m.group(1).split(";")]
        claves, loc, solo_anio = [], None, False
        for i, parte in enumerate(partes):
            mm = re.match(r"(-)?@([^\s,]+)\s*(?:,\s*(.+))?$", parte)
            if not mm:
                inf.aviso(f"Cita mal formada: [{m.group(1)}]")
                return m.group(0)
            if mm.group(1):
                solo_anio = True
            k = clave_ok(mm.group(2))
            if not re.fullmatch(r"[a-z0-9_:\-]+", k):
                inf.aviso(f"La clave «{mm.group(2)}» tiene caracteres no permitidos (usa letras, números, - _ :).")
            claves.append(k)
            ctx.citadas.append(k)
            if mm.group(3):
                if i != len(partes) - 1:
                    inf.aviso(f"El localizador «{mm.group(3)}» solo se admite al final del grupo de citas [{m.group(1)}].")
                loc = mm.group(3)
        extra = (f' data-loc="{attr(html.unescape(loc))}"' if loc else "") + (' data-only="year"' if solo_anio else "")
        return f'<span class="cite" data-ref="{",".join(claves)}"{extra}></span>'
    out = CITA_RE.sub(sub, h)
    if ("[@" in out or "[-@" in out) and not any(a.startswith("Cita mal formada") for a in inf.avisos):
        inf.aviso("Hay una cita que no se reconoció (¿corchetes o formato dentro de la cita?). Revisa el texto cerca de «[@».")
    return out

def marcar_xref(h):
    return XREF_RE.sub(r'<span class="xref">\1\2</span>', h)

def analizar_referencia(cuerpo_txt, clave, inf):
    """Devuelve (data-author, data-year) a partir del texto APA de la referencia."""
    m = re.search(r"\((\d{4}[a-z]?|s\.\s?f\.|en prensa)\)", cuerpo_txt)
    anio = None
    if m:
        anio = m.group(1)
    else:
        mm = re.search(r"(\d{4}[a-z]?)$", clave)
        anio = mm.group(1) if mm else "s. f."
        inf.aviso(f"Referencia «{clave}»: no se encontró el año entre paréntesis; se usó «{anio}».")
    previo = cuerpo_txt[:m.start()] if m else cuerpo_txt
    apellidos = [a.strip() for a, _ in re.findall(r"([^,&]+?),\s*((?:[A-ZÁÉÍÓÚÑÜ]\.[\s\-]?)+)", previo)]
    apellidos = [re.sub(r"^(?:y|&)\s+", "", a).strip() for a in apellidos]
    if not apellidos:
        org = previo.strip().rstrip(".").strip()
        apellidos = [org] if org else [clave]
        if not org:
            inf.aviso(f"Referencia «{clave}»: no se pudo leer el autor.")
    if len(apellidos) == 1:
        aut = apellidos[0]
    elif len(apellidos) == 2:
        aut = f"{apellidos[0]} {CONJUNCION} {apellidos[1]}"
    else:
        aut = f"{apellidos[0]} et al."
    return aut, anio

# ------------------------------------------------------------------ metadatos
ETIQUETAS = {
    "titulo": "titulo", "subtitulo": "subtitulo", "autor": "autores", "autores": "autores",
    "afiliacion": "afiliacion", "afiliaciones": "afiliacion", "correo": "correo",
    "orcid": "orcid", "resumen": "resumen", "palabras clave": "claves",
    "conflicto de interes": "conflictos", "conflictos de interes": "conflictos",
    "uso de ia": "ia", "uso de herramientas de ia": "ia",
    "agradecimiento": "agradecimientos", "agradecimientos": "agradecimientos",
}
INICIO_RE = re.compile(r"^[\s\-—=*_]*comienza (?:tu columna )?aqu[ií][\s\-—=*_]*$", re.I)

OBLIGATORIOS = ("titulo", "autores", "resumen", "claves", "conflictos", "ia")
NOMBRES = {"titulo": "Título", "autores": "Autores", "resumen": "Resumen", "claves": "Palabras clave",
           "conflictos": "Conflictos de interés", "ia": "Uso de herramientas de IA"}

def leer_etiqueta(t):
    m = re.match(r"^\s*([^:\n]{2,60}):\s*(.*)$", t, re.S)
    if not m:
        return None
    et = norm(re.sub(r"\(.*?\)", "", m.group(1))).strip()
    return (ETIQUETAS[et], m.group(2).strip()) if et in ETIQUETAS else None

def es_marcador(v):
    v = v.strip()
    return v.startswith("[") and v.endswith("]")

def nombre_visible(a):
    """'Apellido, Nombre' -> 'Nombre Apellido'."""
    if "," in a:
        ap, no = [x.strip() for x in a.split(",", 1)]
        return f"{no} {ap}".strip()
    return a.strip()

def nombre_apa(a):
    """'Apellido, Nombre Segundo' -> 'Apellido, N. S.'"""
    if "," not in a:
        return a.strip()
    ap, no = [x.strip() for x in a.split(",", 1)]
    ini = " ".join(p[0].upper() + "." for p in re.split(r"[\s]+", no) if p)
    ini = re.sub(r"(?<=\w\.)-", "-", ini)
    return f"{ap}, {ini}".strip().rstrip(",")

# ------------------------------------------------------------------ conversión principal
def convertir(ruta, args, inf):
    d = Doc(ruta, inf)
    ctx = Contexto()
    bloques = list(iter_bloques(d.doc.element.body))

    # --- 1. dónde empieza la columna
    inicio = 0
    for n, (tipo, el) in enumerate(bloques):
        if tipo == "p" and INICIO_RE.match(texto(d.segmentos(el))):
            inicio = n + 1
            break
    else:
        for n, (tipo, el) in enumerate(bloques):
            if tipo == "p":
                lab = leer_etiqueta(texto(d.segmentos(el)))
                if lab and lab[0] == "titulo":
                    inicio = n
                    break
    bloques = bloques[inicio:]

    # --- 2. metadatos
    meta, marcados = {}, set()
    n = 0
    while n < len(bloques):
        tipo, el = bloques[n]
        if tipo != "p":
            break
        t = texto(d.segmentos(el)).strip()
        if not t:
            n += 1
            continue
        lab = leer_etiqueta(t)
        if not lab:
            break
        clave, valor = lab
        meta[clave] = "" if es_marcador(valor) else valor
        if es_marcador(valor) and clave in OBLIGATORIOS:
            inf.error(f"El campo «{NOMBRES[clave]}» todavía tiene el texto de ejemplo de la plantilla.")
            marcados.add(clave)
        n += 1
    cuerpo_bloques = bloques[n:]

    titulo = meta.get("titulo", "").strip()
    autores = [a.strip() for a in meta.get("autores", "").split(";") if a.strip()]
    resumen = meta.get("resumen", "").strip()
    claves_kw = [k.strip() for k in meta.get("claves", "").split(";") if k.strip()]
    for campo in OBLIGATORIOS:
        if not meta.get(campo, "").strip() and campo not in marcados:
            inf.error(f"Falta el campo «{NOMBRES[campo]}».")
    if not meta.get("afiliacion", "").strip():
        inf.aviso("No hay afiliación. Se publicará la columna sin ella.")
    if autores and not all("," in a for a in autores):
        inf.aviso("Los autores deben escribirse «Apellido(s), Nombre(s)» separados por «;». Revisa el campo Autores.")
    if resumen and palabras(resumen) > RESUMEN_MAX:
        inf.aviso(f"El resumen tiene {palabras(resumen)} palabras (máximo sugerido: {RESUMEN_MAX}).")
    if claves_kw and not (N_CLAVES[0] <= len(claves_kw) <= N_CLAVES[1]):
        inf.aviso(f"Hay {len(claves_kw)} palabras clave (se sugieren {N_CLAVES[0]}–{N_CLAVES[1]}).")
    if meta.get("correo"):
        pass  # solo para el editor; no se publica

    # --- 3. cuerpo
    html_cuerpo, refs, refs_dup = [], {}, set()
    texto_cuerpo = []
    primera = True
    en_refs = False
    carpeta_img = None
    n_fig = 0
    pila_listas = []          # [(tipo, nivel)]
    cita_bloque = []          # párrafos de una cita larga en curso
    img_pend = []             # imágenes a guardar [(nombre, bytes)]
    sobrantes = []            # párrafos de la plantilla sin sustituir

    def cerrar_listas(hasta=0):
        while len(pila_listas) > hasta:
            tipo, _ = pila_listas.pop()
            html_cuerpo.append(f"</li></{tipo}>")

    def cerrar_cita():
        if cita_bloque:
            html_cuerpo.append("<blockquote>" + "".join(f"<p>{x}</p>" for x in cita_bloque) + "</blockquote>")
            cita_bloque.clear()

    def inline(segs, xref=True):
        h = a_html(segs)
        h = procesar_citas(h, ctx, inf)
        return marcar_xref(h) if xref else h

    def detectar_mates(t):
        if re.search(r"(?<!\\)\$[^$\n]+?(?<!\\)\$|\\\(|\\\[|\\begin\{", t):
            ctx.hay_mates = True
        if len(re.findall(r"(?<!\\)\$", t)) % 2:
            inf.aviso(f"Un número impar de signos «$» cerca de «{t[:50]}…». Si es un precio, escribe \\$.")

    i = 0
    while i < len(cuerpo_bloques):
        tipo, el = cuerpo_bloques[i]
        i += 1
        if tipo == "tbl":
            cerrar_listas(); cerrar_cita()
            filas = []
            for tr in el.findall(qn("w:tr")):
                celdas = []
                for tc in tr.findall(qn("w:tc")):
                    span = tc.find(qn("w:tcPr") + "/" + qn("w:gridSpan"))
                    ncol = int(span.get(qn("w:val"))) if span is not None else 1
                    partes = []
                    for t2, e2 in iter_bloques(tc):
                        if t2 == "p":
                            h = inline(d.segmentos(e2), xref=False)
                            if h:
                                partes.append(h)
                    celdas.append(("<br>".join(partes), ncol))
                filas.append(celdas)
            leyenda = ""
            if html_cuerpo and html_cuerpo[-1].startswith("<!--TABLA-LEYENDA:"):
                leyenda = html_cuerpo.pop()[len("<!--TABLA-LEYENDA:"):-3]
            t = ['<div class="table-wrap"><table>']
            if leyenda:
                t.append(f"<caption>{leyenda}</caption>")
            for k, fila in enumerate(filas):
                tag = "th" if k == 0 and len(filas) > 1 else "td"
                if k == 0 and tag == "th": t.append("<thead>")
                if k == 1 and filas and len(filas) > 1: t.append("<tbody>")
                t.append("<tr>" + "".join(
                    f"<{tag}{f' colspan={nc}' if nc > 1 else ''}>{c}</{tag}>" for c, nc in fila) + "</tr>")
                if k == 0 and tag == "th": t.append("</thead>")
            if len(filas) > 1: t.append("</tbody>")
            t.append("</table></div>")
            html_cuerpo.append("".join(t))
            continue

        # ---- párrafo
        bandera = d.banderas(el)
        if bandera["omml"]:
            inf.error("Hay una ecuación creada con el editor de ecuaciones de Word/Docs. Escríbela en LaTeX "
                      "(entre $ $ o $$ $$) en su lugar; si no, se pierde al convertir.")
        if bandera["cuadro"]:
            inf.aviso("Hay un cuadro de texto o forma en el documento; su contenido no se convierte.")
        segs = d.segmentos(el)
        t = texto(segs).strip()
        est = d.estilo(el)

        # figuras (imagen en su propio párrafo)
        if bandera["dibujos"]:
            cerrar_listas(); cerrar_cita()
            for dib in bandera["dibujos"]:
                blip = dib.find(".//" + qn("a:blip"))
                if blip is None:
                    continue
                rid = blip.get(qn("r:embed"))
                parte = d.doc.part.related_parts.get(rid)
                if parte is None:
                    continue
                ext = {"image/png": "png", "image/jpeg": "jpg", "image/gif": "gif",
                       "image/svg+xml": "svg", "image/webp": "webp"}.get(parte.content_type)
                if not ext:
                    inf.error(f"Una imagen está en un formato no admitido ({parte.content_type}). Usa PNG o JPG.")
                    continue
                n_fig += 1
                nombre = f"fig-{n_fig}.{ext}"
                img_pend.append((nombre, parte.blob))
                pr = dib.find(".//" + qn("wp:docPr"))
                alt = (pr.get("descr") if pr is not None and pr.get("descr") else "") or ""
                leyenda_html, leyenda_txt = "", ""
                if i < len(cuerpo_bloques) and cuerpo_bloques[i][0] == "p":
                    sg = d.segmentos(cuerpo_bloques[i][1])
                    tt = texto(sg).strip()
                    mm = re.match(r"^(Figura|Fig\.)\s*(\d+(?:\.\d+)?)\s*[.:\-–—]?\s*", tt)
                    if mm:
                        i += 1
                        resto = quitar_prefijo([Seg(s.t, s.b, s.i, s.sup, s.sub, s.href) for s in sg],
                                               len(texto(sg)) - len(texto(sg).lstrip()) + mm.end())
                        leyenda_html = f"<figcaption><b>Figura {mm.group(2)}</b> {inline(resto)}</figcaption>"
                        leyenda_txt = texto(resto).strip()
                if not leyenda_html:
                    inf.aviso(f"La imagen {n_fig} no tiene leyenda. Escribe debajo «Figura N. Descripción…».")
                if not alt:
                    alt = leyenda_txt or f"Figura {n_fig}"
                html_cuerpo.append(f'<figure><img src="{{SLUG}}/{nombre}" alt="{attr(alt)}" loading="lazy">{leyenda_html}</figure>')
            continue

        if not t:
            continue
        if re.match(r"^\[(escribe|describe|presenta|desarrolla|pega|resume|enuncia|cierra|anade|anota|indica|incluye|lista)\b", norm(t)):
            sobrantes.append(t[:50])
            continue
        if est == "title":
            inf.aviso("Se ignoró un párrafo con estilo «Título» dentro del cuerpo; usa «Título 1» para los apartados.")
            continue

        # encabezados
        mh = re.match(r"heading (\d)", est)
        if mh:
            cerrar_listas(); cerrar_cita()
            nivel = int(mh.group(1))
            h_txt = a_html(segs)
            if norm(t) in ("referencias", "referencias bibliograficas", "bibliografia"):
                en_refs = True
                continue
            en_refs = False
            html_cuerpo.append(f"<h{min(nivel + 1, 4)}>{h_txt}</h{min(nivel + 1, 4)}>")
            continue

        # referencias
        if en_refs:
            mr = re.match(r"^\[@?([^\]\s]+)\]\s*", t)
            if not mr:
                inf.error(f"Referencia sin clave al inicio (escribe «[clave] Apellido, N. (Año)…»): «{t[:60]}…»")
                continue
            k = clave_ok(mr.group(1))
            ini = len(texto(segs)) - len(texto(segs).lstrip())
            resto = quitar_prefijo(segs, ini + mr.end())
            cuerpo_ref = a_html(resto)
            if k in refs:
                refs_dup.add(k)
                inf.error(f"La referencia «{k}» está repetida.")
            aut, anio = analizar_referencia(texto(resto), k, inf)
            refs[k] = (aut, anio, cuerpo_ref)
            continue

        texto_cuerpo.append(t)
        detectar_mates(t)

        # ecuaciones en párrafo propio
        if re.match(r"^\$\$.*\$\$$", t, re.S) or re.match(r"^\\begin\{[a-z*]+\}", t):
            cerrar_listas(); cerrar_cita()
            html_cuerpo.append(f'<div class="math">{esc(t)}</div>')
            primera = False
            continue

        # citas largas
        es_cita = est in ("quote", "intense quote", "block text") or t.startswith("> ")
        if es_cita:
            cerrar_listas()
            if t.startswith("> "):
                segs = quitar_prefijo(segs, len(texto(segs)) - len(texto(segs).lstrip()) + 2)
            cita_bloque.append(inline(segs))
            continue
        cerrar_cita()

        # listas
        li = d.lista(el)
        if li:
            tipo_l, nivel = li
            nivel = min(nivel, 3)
            if not pila_listas:
                pila_listas.append((tipo_l, nivel))
                html_cuerpo.append(f"<{tipo_l}><li>{inline(segs)}")
            else:
                actual = pila_listas[-1][1]
                if nivel > actual:
                    pila_listas.append((tipo_l, nivel))
                    html_cuerpo.append(f"<{tipo_l}><li>{inline(segs)}")
                else:
                    while len(pila_listas) > 1 and pila_listas[-1][1] > nivel:
                        tp, _ = pila_listas.pop()
                        html_cuerpo.append(f"</li></{tp}>")
                    html_cuerpo.append(f"</li><li>{inline(segs)}")
            continue
        cerrar_listas()

        # leyenda de tabla (párrafo «Tabla N.» justo antes de una tabla)
        if i < len(cuerpo_bloques) and cuerpo_bloques[i][0] == "tbl" and re.match(r"^Tabla\s+\d", t):
            html_cuerpo.append("<!--TABLA-LEYENDA:" + inline(segs, xref=False) + "-->")
            continue

        clase = ' class="lead"' if primera else ""
        primera = False
        html_cuerpo.append(f"<p{clase}>{inline(segs)}</p>")
    cerrar_listas(); cerrar_cita()

    # --- 4. comprobaciones
    if sobrantes:
        inf.error(f"Quedan {len(sobrantes)} texto(s) de ejemplo de la plantilla entre corchetes, p. ej. «{sobrantes[0]}…». Sustitúyelos o bórralos.")
    n_pal = sum(palabras(x) for x in texto_cuerpo)
    if not texto_cuerpo:
        inf.error("El documento no tiene texto en el cuerpo.")
    elif not (LONGITUD[0] <= n_pal <= LONGITUD[1]):
        inf.aviso(f"El cuerpo tiene {n_pal} palabras (rango sugerido: {LONGITUD[0]}–{LONGITUD[1]}).")
    for k in sorted(set(ctx.citadas)):
        if k not in refs:
            inf.error(f"La cita «{k}» no tiene referencia en la lista de referencias.")
    for k in refs:
        if k not in ctx.citadas:
            inf.aviso(f"La referencia «{k}» no se cita en el texto.")
    if not ctx.citadas and not refs:
        inf.aviso("El texto no tiene citas ni referencias.")
    if d.pie_notas:
        inf.aviso("El documento tiene notas al pie; no se convierten. Pásalas al texto o a las referencias.")
    if d.cambios_marcados:
        inf.aviso("El documento tiene cambios sin resolver (control de cambios/sugerencias). Se tomó el texto con las "
                  "inserciones aceptadas y las eliminaciones aplicadas; confírmalo con la persona autora.")

    return {
        "titulo": titulo, "subtitulo": meta.get("subtitulo", "").strip(), "autores": autores,
        "afiliacion": [x.strip() for x in meta.get("afiliacion", "").split(";") if x.strip()],
        "orcid": [x.strip() for x in meta.get("orcid", "").split(";") if x.strip()],
        "resumen": resumen, "claves": claves_kw, "conflictos": meta.get("conflictos", "").strip(),
        "ia": meta.get("ia", "").strip(), "agradecimientos": meta.get("agradecimientos", "").strip(),
        "cuerpo": "\n".join(html_cuerpo), "refs": refs, "hay_mates": ctx.hay_mates, "n_pal": n_pal,
        "imagenes": img_pend,
    }

# ------------------------------------------------------------------ salida
def autoria_html(r):
    afs, ors = r["afiliacion"], r["orcid"]
    emparejar_af = len(afs) == len(r["autores"]) and len(afs) > 1
    out = []
    for k, a in enumerate(r["autores"]):
        linea = f'<span class="author"><b>{esc(nombre_visible(a))}</b>'
        if len(ors) == len(r["autores"]) and ors[k]:
            o = ors[k].strip()
            url = o if o.startswith("http") else f"https://orcid.org/{o}"
            linea += f' <a class="orcid" href="{attr(url)}">ORCID</a>'
        linea += "</span>"
        if emparejar_af:
            linea += f'<span class="affil">{esc(afs[k])}</span>'
        out.append(linea)
    if afs and not emparejar_af:
        out.append(f'<span class="affil">{esc("; ".join(afs))}</span>')
    return "".join(out)

def referencias_html(refs):
    if not refs:
        return ""
    def orden(k):
        return norm(re.sub(r"<[^>]+>", "", refs[k][2]))
    filas = []
    for k in sorted(refs, key=lambda k: (norm(refs[k][0]), refs[k][1], k)):
        aut, anio, cuerpo = refs[k]
        filas.append(f'<div class="ref" id="ref-{k}" data-author="{attr(aut)}" data-year="{attr(anio)}">{cuerpo}</div>')
    return '<section class="references"><h2>Referencias bibliográficas</h2>\n' + "\n".join(filas) + "\n</section>"

def declaraciones_html(r):
    partes = []
    if r["conflictos"]:
        partes.append(f"<h4>Conflictos de interés</h4><p>{esc(r['conflictos'])}</p>")
    if r["ia"]:
        partes.append(f"<h4>Uso de herramientas de inteligencia artificial</h4><p>{esc(r['ia'])}</p>")
    if r["agradecimientos"]:
        partes.append(f"<h4>Agradecimientos</h4><p>{esc(r['agradecimientos'])}</p>")
    if LICENCIA:
        partes.append(f"<h4>Licencia</h4><p>{esc(LICENCIA)}</p>")
    if not partes:
        return ""
    return '<section class="box box--outline declaraciones"><p class="box-title">Declaraciones</p>' + "".join(partes) + "</section>"

def como_citar_html(r, numero, slug, anio):
    if not r["autores"]:
        return ""
    apa = [nombre_apa(a) for a in r["autores"]]
    if len(apa) == 1:
        aut = apa[0]
    elif len(apa) <= 20:
        aut = ", ".join(apa[:-1]) + f", & {apa[-1]}"
    else:
        aut = ", ".join(apa[:19]) + f", … {apa[-1]}"
    num = f", <i>n.º {numero}</i>" if numero else ""
    url = f"{SITIO_URL}/columnas/{slug}.html"
    return (f'<section class="cite-as box"><p class="box-title">Cómo citar</p>'
            f'<p>{esc(aut)} ({anio}). {esc(r["titulo"])}. <i>{esc(SECCION)}</i>{num}. <a href="{attr(url)}">{esc(url)}</a></p></section>')

def actualizar_listado(raiz, catalogo):
    ruta = raiz / "columnas.html"
    s = ruta.read_text(encoding="utf-8")
    ini = s.index("<!-- COLUMNAS:INICIO")
    fin = s.index("<!-- COLUMNAS:FIN -->")
    cabecera = s[ini:s.index("-->", ini) + 3]
    if not catalogo:
        bloque = '<p class="vacio">Aún no hay columnas publicadas.</p>'
    else:
        filas = []
        for c in sorted(catalogo, key=lambda c: -c["numero"]):
            nombres = ", ".join(nombre_visible(a) for a in c["autores"])
            res = c["resumen"]
            if len(res) > 230:
                res = res[:230].rsplit(" ", 1)[0].rstrip(",;:.") + "…"
            filas.append(
                f'  <li><span class="num">{c["numero"]}</span><span><a class="title" href="columnas/{c["slug"]}.html">{esc(c["titulo"])}</a>'
                f'<span class="desc">{esc(res)}</span>'
                f'<span class="meta">{esc(nombres)} · {fecha_larga(c["publicada"])}</span></span></li>')
        bloque = '<ol class="toc">\n' + "\n".join(filas) + "\n</ol>"
    ruta.write_text(s[:ini] + cabecera + "\n" + bloque + "\n" + s[fin:], encoding="utf-8")

def reducir(blob, nombre):
    try:
        from PIL import Image
    except ImportError:
        return blob
    try:
        im = Image.open(io.BytesIO(blob))
        if im.width <= ANCHO_MAX_IMG or nombre.endswith(("svg", "gif")):
            return blob
        im = im.resize((ANCHO_MAX_IMG, round(im.height * ANCHO_MAX_IMG / im.width)), Image.LANCZOS)
        out = io.BytesIO()
        im.save(out, format="PNG" if nombre.endswith("png") else "JPEG", optimize=True, **({} if nombre.endswith("png") else {"quality": 88}))
        return out.getvalue()
    except Exception:
        return blob

def main():
    ap = argparse.ArgumentParser(description="Convierte una entrega .docx en una página de Columnas.")
    ap.add_argument("docx")
    ap.add_argument("--slug", help="nombre del archivo (por defecto, sale del título)")
    ap.add_argument("--recibida"); ap.add_argument("--aceptada"); ap.add_argument("--publicada")
    ap.add_argument("--solo-revisar", action="store_true", help="muestra el informe y no escribe nada")
    ap.add_argument("--actualizar", action="store_true", help="reemplaza una columna ya publicada (conserva su número)")
    ap.add_argument("--ejemplo", action="store_true", help="página de ejemplo: aviso visible, noindex y sin entrar en el listado")
    args = ap.parse_args()

    for nom in ("recibida", "aceptada", "publicada"):
        v = getattr(args, nom)
        if v:
            try: datetime.date.fromisoformat(v)
            except ValueError: sys.exit(f"--{nom}: usa el formato AAAA-MM-DD (por ejemplo 2026-11-27)")
    inf = Informe()
    r = convertir(args.docx, args, inf)

    slug = args.slug or slugify(r["titulo"] or Path(args.docx).stem)
    destino = RAIZ / "columnas" / f"{slug}.html"
    cat_ruta = RAIZ / "columnas" / "catalogo.json"
    catalogo = json.loads(cat_ruta.read_text(encoding="utf-8")) if cat_ruta.exists() else []
    previo = next((c for c in catalogo if c["slug"] == slug), None)
    if destino.exists() and not args.actualizar and not args.ejemplo:
        inf.error(f"Ya existe columnas/{slug}.html. Usa --actualizar para reemplazarla o --slug para otro nombre.")

    publicada = args.publicada or (previo["publicada"] if previo else datetime.date.today().isoformat())
    numero = None if args.ejemplo else (previo["numero"] if previo else 1 + max([c["numero"] for c in catalogo], default=0))

    # ---- informe
    print(f"\nColumna: {r['titulo'] or '(sin título)'}")
    print(f"Autores: {', '.join(nombre_visible(a) for a in r['autores']) or '(sin autores)'}")
    print(f"Cuerpo: {r['n_pal']} palabras · Referencias: {len(r['refs'])} · Figuras: {len(r['imagenes'])}")
    for m in inf.avisos: print("  AVISO:", m)
    for m in inf.errores: print("  ERROR:", m)
    if inf.errores:
        print(f"\nNo se escribió nada: hay {len(inf.errores)} error(es) que corregir en el documento.")
        sys.exit(1)
    if args.solo_revisar:
        print("\nRevisión terminada (no se escribió nada).")
        return

    # ---- página
    plantilla = (RAIZ / "tools" / "plantilla-columna.html").read_text(encoding="utf-8")
    anio = publicada[:4]
    fechas = []
    if args.recibida: fechas.append(f"Recibida: {fecha_larga(args.recibida)}")
    if args.aceptada: fechas.append(f"Aceptada: {fecha_larga(args.aceptada)}")
    fechas.append(f"Publicada: {fecha_larga(publicada)}")
    autoria = autoria_html(r) + '<span class="dates">' + "<br>".join(fechas) + "</span>"
    cuerpo = r["cuerpo"].replace("{SLUG}", slug)
    valores = {
        "titulo": esc(r["titulo"]), "titulo_txt": attr(r["titulo"]),
        "descripcion": attr(r["resumen"][:300]),
        "autores_txt": attr(", ".join(nombre_visible(a) for a in r["autores"])),
        "meta_autores": "".join(f'<meta name="citation_author" content="{attr(a)}">\n' for a in r["autores"]),
        "fecha_iso": publicada,
        "robots": '<meta name="robots" content="noindex">\n' if args.ejemplo else "",
        "math_script": '<script src="../math.js"></script>\n' if r["hay_mates"] else "",
        "kicker": "Ejemplo de columna" if args.ejemplo else (f"Columna n.º {numero}"),
        "subtitulo": f'<p class="deck">{esc(r["subtitulo"])}</p>' if r["subtitulo"] else "",
        "aviso": ('<div class="aviso"><b>Ejemplo.</b> Esta página muestra cómo se ve una columna. '
                  'No es una columna real y no figura en el listado.</div>\n') if args.ejemplo else "",
        "autoria": autoria, "resumen": esc(r["resumen"]),
        "palabras_clave": esc("; ".join(r["claves"])),
        "cuerpo": cuerpo, "declaraciones": declaraciones_html(r),
        "referencias": referencias_html(r["refs"]),
        "como_citar": como_citar_html(r, numero, slug, anio),
    }
    pagina = plantilla
    for k, v in valores.items():
        pagina = pagina.replace("{{" + k + "}}", v)
    sobran = re.findall(r"\{\{\w+\}\}", pagina)
    if sobran:
        sys.exit(f"La plantilla tiene marcas sin rellenar: {sobran}")

    destino.parent.mkdir(exist_ok=True)
    destino.write_text(pagina, encoding="utf-8")
    if r["imagenes"]:
        carpeta = RAIZ / "columnas" / slug
        carpeta.mkdir(exist_ok=True)
        for nombre, blob in r["imagenes"]:
            (carpeta / nombre).write_bytes(reducir(blob, nombre))

    if not args.ejemplo:
        entrada = {"numero": numero, "slug": slug, "titulo": r["titulo"], "autores": r["autores"],
                   "resumen": r["resumen"], "palabras_clave": r["claves"],
                   "recibida": args.recibida or (previo or {}).get("recibida"),
                   "aceptada": args.aceptada or (previo or {}).get("aceptada"), "publicada": publicada}
        catalogo = [c for c in catalogo if c["slug"] != slug] + [entrada]
        cat_ruta.write_text(json.dumps(sorted(catalogo, key=lambda c: c["numero"]), ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        actualizar_listado(RAIZ, catalogo)
    print(f"\nListo: columnas/{slug}.html" + ("" if args.ejemplo else " (y se actualizó columnas.html)"))
    print("Revisa la página en el navegador antes de subirla a GitHub.")

if __name__ == "__main__":
    main()
