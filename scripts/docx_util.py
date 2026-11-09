"""
Utilidades de formato para el documento del PIA de Lumina.

Centraliza portada, estilos, tablas, figuras, cajas de nota y numeración de
página. El objetivo es que el documento final tenga una presentación
consistente y que el contenido se escriba sin repetir XML de Word.
"""

from __future__ import annotations

import re
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor
from PIL import Image, ImageChops

# ---------------------------------------------------------------------------
# Paleta y rutas
# ---------------------------------------------------------------------------

CARBON = "171512"
CARBON_SUAVE = "2A241D"
DORADO = "C6A15B"
DORADO_CLARO = "E3C98F"
CREMA = "F7F2E9"
CREMA_OSCURA = "EFE7D8"
LINEA = "D8D2C4"
GRIS = "6F675B"
BLANCO = "FFFDF9"
ROJO = "A12622"

RGB_CARBON = RGBColor(0x17, 0x15, 0x12)
RGB_DORADO = RGBColor(0xC6, 0xA1, 0x5B)
RGB_GRIS = RGBColor(0x6F, 0x67, 0x5B)
RGB_BLANCO = RGBColor(0xFF, 0xFF, 0xFF)

BASE = Path(__file__).resolve().parent.parent
DOCS = BASE / "docs"
ASSETS = DOCS / "assets"
CAPTURAS = BASE / "capturas"
CAPTURAS_DOC = ASSETS / "capturas"

FUENTE_TEXTO = "Arial"
FUENTE_TITULO = "Georgia"
FUENTE_CODIGO = "Consolas"

ANCHO_UTIL = 6.55  # pulgadas disponibles con márgenes de 2.3 cm

_figura_actual = 0
_tabla_actual = 0


# ---------------------------------------------------------------------------
# Preparación de imágenes
# ---------------------------------------------------------------------------

def _recortar_margen_blanco(imagen: Image.Image) -> Image.Image:
    """Recorta el área blanca inferior de las capturas del admin.

    Las pantallas de Django suelen terminar mucho antes que el viewport, así que
    el archivo original arrastra cientos de píxeles vacíos. Recortarlos hace que
    la figura ocupe solo su contenido y que el documento respire mejor.
    """
    rgb = imagen.convert("RGB")
    bbox = ImageChops.difference(rgb, Image.new("RGB", rgb.size, (255, 255, 255))).getbbox()
    if not bbox:
        return rgb
    _, _, derecha, abajo = bbox
    margen = 24
    return rgb.crop((0, 0, min(rgb.width, derecha + margen), min(rgb.height, abajo + margen)))


def preparar_capturas() -> None:
    """Reduce las capturas para que el .docx no crezca de más sin perder nitidez."""
    CAPTURAS_DOC.mkdir(parents=True, exist_ok=True)
    for origen in sorted(CAPTURAS.glob("*.png")):
        destino = CAPTURAS_DOC / origen.name
        if destino.exists() and destino.stat().st_mtime >= origen.stat().st_mtime:
            continue
        imagen = Image.open(origen)
        if imagen.mode not in ("RGB", "L"):
            fondo = Image.new("RGB", imagen.size, (255, 255, 255))
            fondo.paste(imagen, mask=imagen.split()[-1] if imagen.mode == "RGBA" else None)
            imagen = fondo
        if origen.name.startswith(("admin_", "rbac_")):
            imagen = _recortar_margen_blanco(imagen)
        imagen.thumbnail((1440, 5000), Image.LANCZOS)
        imagen.save(destino, "PNG", optimize=True)
    print(f"Capturas preparadas para el documento: {len(list(CAPTURAS_DOC.glob('*.png')))}")


# ---------------------------------------------------------------------------
# Estilos generales
# ---------------------------------------------------------------------------

def _fuente(run, nombre: str, tamano: float | None = None, negrita: bool | None = None,
            cursiva: bool | None = None, color: str | RGBColor | None = None) -> None:
    run.font.name = nombre
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    for atributo in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rfonts.set(qn(atributo), nombre)
    if tamano is not None:
        run.font.size = Pt(tamano)
    if negrita is not None:
        run.font.bold = negrita
    if cursiva is not None:
        run.font.italic = cursiva
    if color is not None:
        run.font.color.rgb = RGBColor.from_string(color) if isinstance(color, str) else color


def configurar_estilos(doc: Document) -> None:
    seccion = doc.sections[0]
    seccion.page_width = Inches(8.5)
    seccion.page_height = Inches(11)
    seccion.top_margin = Cm(2.3)
    seccion.bottom_margin = Cm(2.1)
    seccion.left_margin = Cm(2.3)
    seccion.right_margin = Cm(2.3)
    seccion.header_distance = Cm(1.0)
    seccion.footer_distance = Cm(0.9)

    normal = doc.styles["Normal"]
    normal.font.name = FUENTE_TEXTO
    normal.font.size = Pt(10.5)
    normal.font.color.rgb = RGBColor(0x2C, 0x28, 0x23)
    normal.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    normal.paragraph_format.line_spacing = 1.15
    normal.paragraph_format.space_after = Pt(6)
    rpr = normal.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    for atributo in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rfonts.set(qn(atributo), FUENTE_TEXTO)

    estilos_titulo = {
        "Heading 1": (17, True, RGB_CARBON, 20, 8),
        "Heading 2": (13.5, True, RGB_CARBON, 14, 5),
        "Heading 3": (11.5, True, RGB_DORADO, 11, 4),
    }
    for nombre, (tamano, negrita, color, antes, despues) in estilos_titulo.items():
        estilo = doc.styles[nombre]
        estilo.font.name = FUENTE_TEXTO
        estilo.font.size = Pt(tamano)
        estilo.font.bold = negrita
        estilo.font.color.rgb = color
        estilo.paragraph_format.space_before = Pt(antes)
        estilo.paragraph_format.space_after = Pt(despues)
        estilo.paragraph_format.keep_with_next = True
        estilo.paragraph_format.line_spacing = 1.05
        estilo.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT

    # Estilo propio para pies de figura y tabla.
    from docx.enum.style import WD_STYLE_TYPE
    if "Pie Lumina" not in [s.name for s in doc.styles]:
        pie = doc.styles.add_style("Pie Lumina", WD_STYLE_TYPE.PARAGRAPH)
    else:
        pie = doc.styles["Pie Lumina"]
    pie.font.name = FUENTE_TEXTO
    pie.font.size = Pt(8.5)
    pie.font.italic = True
    pie.font.color.rgb = RGB_GRIS
    pie.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pie.paragraph_format.space_before = Pt(3)
    pie.paragraph_format.space_after = Pt(10)
    pie.paragraph_format.line_spacing = 1.0

    # Los estilos Header/Footer heredan tabulaciones centrales que rompen la
    # alineación. Se reemplazan por una única tabulación derecha al margen.
    for nombre in ("Header", "Footer"):
        estilo = doc.styles[nombre]
        tabulaciones = estilo.paragraph_format.tab_stops
        tabulaciones.clear_all()
        tabulaciones.add_tab_stop(Inches(ANCHO_UTIL), WD_TAB_ALIGNMENT.RIGHT)


# ---------------------------------------------------------------------------
# Texto enriquecido
# ---------------------------------------------------------------------------

_PATRON = re.compile(r"(\*\*.+?\*\*|\*[^*]+\*|`.+?`)")


def _agregar_runs(par, texto: str, tamano: float | None = None) -> None:
    for parte in _PATRON.split(texto):
        if not parte:
            continue
        if parte.startswith("**") and parte.endswith("**"):
            run = par.add_run(parte[2:-2])
            run.bold = True
            _fuente(run, FUENTE_TEXTO, tamano)
        elif parte.startswith("*") and parte.endswith("*") and len(parte) > 2:
            run = par.add_run(parte[1:-1])
            run.italic = True
            _fuente(run, FUENTE_TEXTO, tamano)
        elif parte.startswith("`") and parte.endswith("`"):
            run = par.add_run(parte[1:-1])
            _fuente(run, FUENTE_CODIGO, (tamano or 10.5) - 1)
            run.font.color.rgb = RGBColor(0x7A, 0x45, 0x12)
        else:
            run = par.add_run(parte)
            _fuente(run, FUENTE_TEXTO, tamano)


def par(doc, texto: str = "", *, negrita: bool = False, cursiva: bool = False,
        tamano: float | None = None, alineacion=WD_ALIGN_PARAGRAPH.JUSTIFY,
        antes: float | None = None, despues: float | None = None,
        color: str | None = None, estilo: str | None = None, sangria: float | None = None):
    p = doc.add_paragraph(style=estilo) if estilo else doc.add_paragraph()
    _agregar_runs(p, texto, tamano)
    if negrita or cursiva or color:
        for run in p.runs:
            if negrita:
                run.bold = True
            if cursiva:
                run.italic = True
            if color:
                run.font.color.rgb = RGBColor.from_string(color)
    p.paragraph_format.alignment = alineacion
    if antes is not None:
        p.paragraph_format.space_before = Pt(antes)
    if despues is not None:
        p.paragraph_format.space_after = Pt(despues)
    if sangria is not None:
        p.paragraph_format.left_indent = Cm(sangria)
    return p


def vinetas(doc, elementos: list[str], nivel: int = 0) -> None:
    estilo = "List Bullet" if nivel == 0 else "List Bullet 2"
    for elemento in elementos:
        p = doc.add_paragraph(style=estilo)
        _agregar_runs(p, elemento)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.line_spacing = 1.12
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT


def pasos(doc, elementos: list[str]) -> None:
    for indice, elemento in enumerate(elementos, 1):
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.8)
        p.paragraph_format.first_line_indent = Cm(-0.55)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run(f"{indice}.  ")
        run.bold = True
        run.font.color.rgb = RGB_DORADO
        _fuente(run, FUENTE_TEXTO, 10.5)
        _agregar_runs(p, elemento)


def espacio(doc, puntos: float = 6) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.line_spacing = 1.0
    run = p.add_run(" ")
    run.font.size = Pt(puntos)


# ---------------------------------------------------------------------------
# Encabezados
# ---------------------------------------------------------------------------

def _borde_inferior(par, color: str = DORADO, grosor: int = 12) -> None:
    ppr = par._p.get_or_add_pPr()
    bordes = ppr.find(qn("w:pBdr"))
    if bordes is None:
        bordes = OxmlElement("w:pBdr")
        ppr.append(bordes)
    borde = OxmlElement("w:bottom")
    borde.set(qn("w:val"), "single")
    borde.set(qn("w:sz"), str(grosor))
    borde.set(qn("w:space"), "4")
    borde.set(qn("w:color"), color)
    bordes.append(borde)


def h1(doc, texto: str, *, salto: bool = True):
    if salto:
        doc.add_page_break()
    p = doc.add_heading(texto, level=1)
    _borde_inferior(p, DORADO, 10)
    return p


def h2(doc, texto: str):
    return doc.add_heading(texto, level=2)


def h3(doc, texto: str):
    return doc.add_heading(texto, level=3)


# ---------------------------------------------------------------------------
# Tablas
# ---------------------------------------------------------------------------

def _sombra_celda(celda, color: str) -> None:
    tcpr = celda._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color)
    tcpr.append(shd)


def _bordes_tabla(tabla, color: str = LINEA, grosor: int = 4) -> None:
    tblpr = tabla._tbl.tblPr
    bordes = OxmlElement("w:tblBorders")
    for borde in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{borde}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(grosor))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        bordes.append(el)
    tblpr.append(bordes)


def _margenes_celda(tabla, arriba: int = 40, abajo: int = 40,
                    izquierda: int = 80, derecha: int = 80) -> None:
    tblpr = tabla._tbl.tblPr
    margenes = OxmlElement("w:tblCellMar")
    for nombre, valor in (("top", arriba), ("left", izquierda),
                          ("bottom", abajo), ("right", derecha)):
        el = OxmlElement(f"w:{nombre}")
        el.set(qn("w:w"), str(valor))
        el.set(qn("w:type"), "dxa")
        margenes.append(el)
    tblpr.append(margenes)


def _repetir_encabezado(fila) -> None:
    trpr = fila._tr.get_or_add_trPr()
    encabezado = OxmlElement("w:tblHeader")
    encabezado.set(qn("w:val"), "true")
    trpr.append(encabezado)


def _no_partir_fila(fila) -> None:
    trpr = fila._tr.get_or_add_trPr()
    cant = OxmlElement("w:cantSplit")
    trpr.append(cant)


def _texto_celda(celda, texto: str, *, negrita: bool = False, tamano: float = 8.5,
                 color: str | None = None, alineacion=WD_ALIGN_PARAGRAPH.LEFT) -> None:
    celda.text = ""
    p = celda.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.line_spacing = 1.0
    p.paragraph_format.alignment = alineacion
    _agregar_runs(p, str(texto), tamano)
    for run in p.runs:
        if negrita:
            run.bold = True
        if color:
            run.font.color.rgb = RGBColor.from_string(color)


def tabla(doc, encabezados: list[str], filas: list[list[str]], anchos: list[float] | None = None,
          titulo: str | None = None, tamano: float = 8.5,
          primera_negrita: bool = False, centrar: tuple[int, ...] = (),
          fuente: str = "Elaboración propia (2026).") -> None:
    global _tabla_actual
    if titulo:
        _tabla_actual += 1
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.keep_with_next = True
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run(f"Tabla {_tabla_actual}. ")
        run.bold = True
        _fuente(run, FUENTE_TEXTO, 8.5, color=RGB_CARBON)
        run2 = p.add_run(titulo)
        _fuente(run2, FUENTE_TEXTO, 8.5, cursiva=True, color=RGB_CARBON)

    tabla_obj = doc.add_table(rows=1, cols=len(encabezados))
    tabla_obj.style = "Table Grid"
    tabla_obj.alignment = WD_TABLE_ALIGNMENT.CENTER
    tabla_obj.autofit = False
    tabla_obj.allow_autofit = False
    _bordes_tabla(tabla_obj)
    _margenes_celda(tabla_obj)

    if anchos is None:
        anchos = [ANCHO_UTIL / len(encabezados)] * len(encabezados)
    else:
        total = sum(anchos)
        anchos = [ancho * ANCHO_UTIL / total for ancho in anchos]

    for indice, texto in enumerate(encabezados):
        celda = tabla_obj.rows[0].cells[indice]
        celda.width = Inches(anchos[indice])
        _sombra_celda(celda, CARBON)
        _texto_celda(celda, texto, negrita=True, tamano=tamano, color=BLANCO,
                     alineacion=WD_ALIGN_PARAGRAPH.CENTER)
        celda.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    _repetir_encabezado(tabla_obj.rows[0])
    _no_partir_fila(tabla_obj.rows[0])

    for numero_fila, fila in enumerate(filas):
        celdas = tabla_obj.add_row().cells
        _no_partir_fila(tabla_obj.rows[-1])
        for indice, valor in enumerate(fila):
            celda = celdas[indice]
            celda.width = Inches(anchos[indice])
            alineacion = WD_ALIGN_PARAGRAPH.CENTER if indice in centrar else WD_ALIGN_PARAGRAPH.LEFT
            _texto_celda(
                celda, valor, tamano=tamano,
                negrita=primera_negrita and indice == 0,
                alineacion=alineacion,
            )
            celda.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            if numero_fila % 2 == 1:
                _sombra_celda(celda, CREMA)

    # Las tablas cortas no deben partirse entre páginas.
    if len(filas) <= 10:
        for fila in tabla_obj.rows:
            for celda in fila.cells:
                for parrafo in celda.paragraphs:
                    parrafo.paragraph_format.keep_with_next = True

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(f"Fuente: {fuente}")
    _fuente(run, FUENTE_TEXTO, 7.5, cursiva=True, color=RGB_GRIS)


# ---------------------------------------------------------------------------
# Figuras
# ---------------------------------------------------------------------------

def figura(doc, archivo: str, titulo: str, *, ancho: float = 6.3,
           max_alto: float = 7.2, fuente: str = "Elaboración propia (2026).") -> None:
    global _figura_actual
    ruta = CAPTURAS_DOC / archivo
    if not ruta.exists():
        ruta = CAPTURAS / archivo
    if not ruta.exists():
        raise FileNotFoundError(f"No se encontró la captura: {ruta}")

    with Image.open(ruta) as imagen:
        ancho_px, alto_px = imagen.size
    alto = ancho * alto_px / ancho_px
    if alto > max_alto:
        ancho = max_alto * ancho_px / alto_px

    doc.add_picture(str(ruta), width=Inches(ancho))
    p_imagen = doc.paragraphs[-1]
    p_imagen.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_imagen.paragraph_format.space_before = Pt(6)
    p_imagen.paragraph_format.space_after = Pt(2)
    p_imagen.paragraph_format.keep_with_next = True

    _figura_actual += 1
    p = doc.add_paragraph(style="Pie Lumina")
    run = p.add_run(f"Figura {_figura_actual}. ")
    run.bold = True
    run.italic = False
    p.add_run(f"{titulo}. Fuente: {fuente}")
    p.paragraph_format.keep_together = True


# ---------------------------------------------------------------------------
# Cajas de nota y código
# ---------------------------------------------------------------------------

def caja(doc, titulo: str, texto: str, *, color_fondo: str = CREMA,
         color_borde: str = DORADO, tamano: float = 9.2) -> None:
    tabla_obj = doc.add_table(rows=1, cols=1)
    tabla_obj.alignment = WD_TABLE_ALIGNMENT.CENTER
    tabla_obj.autofit = False
    celda = tabla_obj.rows[0].cells[0]
    celda.width = Inches(ANCHO_UTIL)
    _sombra_celda(celda, color_fondo)
    _margenes_celda(tabla_obj, arriba=110, abajo=110, izquierda=160, derecha=160)

    tblpr = tabla_obj._tbl.tblPr
    bordes = OxmlElement("w:tblBorders")
    for borde in ("top", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{borde}")
        el.set(qn("w:val"), "nil")
        bordes.append(el)
    izquierdo = OxmlElement("w:left")
    izquierdo.set(qn("w:val"), "single")
    izquierdo.set(qn("w:sz"), "18")
    izquierdo.set(qn("w:space"), "0")
    izquierdo.set(qn("w:color"), color_borde)
    bordes.append(izquierdo)
    tblpr.append(bordes)
    _no_partir_fila(tabla_obj.rows[0])

    celda.text = ""
    p = celda.paragraphs[0]
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run(titulo)
    _fuente(run, FUENTE_TEXTO, tamano, negrita=True, color=RGB_CARBON)
    p2 = celda.add_paragraph()
    p2.paragraph_format.space_after = Pt(0)
    p2.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _agregar_runs(p2, texto, tamano)
    espacio(doc, 6)


def codigo(doc, texto: str, *, titulo: str | None = None) -> None:
    if titulo:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(titulo)
        _fuente(run, FUENTE_TEXTO, 8.5, negrita=True, color=RGB_GRIS)

    tabla_obj = doc.add_table(rows=1, cols=1)
    tabla_obj.alignment = WD_TABLE_ALIGNMENT.CENTER
    tabla_obj.autofit = False
    celda = tabla_obj.rows[0].cells[0]
    celda.width = Inches(ANCHO_UTIL)
    _sombra_celda(celda, "F4F1EA")
    _margenes_celda(tabla_obj, arriba=90, abajo=90, izquierda=140, derecha=140)

    tblpr = tabla_obj._tbl.tblPr
    bordes = OxmlElement("w:tblBorders")
    for borde in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{borde}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "3")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), LINEA)
        bordes.append(el)
    tblpr.append(bordes)
    _no_partir_fila(tabla_obj.rows[0])

    celda.text = ""
    lineas = texto.strip("\n").split("\n")
    for indice, linea in enumerate(lineas):
        p = celda.paragraphs[0] if indice == 0 else celda.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.0
        run = p.add_run(linea if linea else " ")
        _fuente(run, FUENTE_CODIGO, 8)
        run.font.color.rgb = RGBColor(0x2C, 0x28, 0x23)
    espacio(doc, 6)


# ---------------------------------------------------------------------------
# Portada, índice y encabezados de página
# ---------------------------------------------------------------------------

def _tabla_sin_bordes(tabla_obj) -> None:
    tblpr = tabla_obj._tbl.tblPr
    bordes = OxmlElement("w:tblBorders")
    for borde in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{borde}")
        el.set(qn("w:val"), "nil")
        bordes.append(el)
    tblpr.append(bordes)


def _fijar_anchos_tabla(tabla_obj, anchos: list[float]) -> None:
    """Fija ancho total, cuadrícula y celdas para que Word no reinterprete la tabla."""
    total = int(round(sum(anchos) * 1440))
    tblpr = tabla_obj._tbl.tblPr

    tblw = tblpr.find(qn("w:tblW"))
    if tblw is None:
        tblw = OxmlElement("w:tblW")
        tblpr.append(tblw)
    tblw.set(qn("w:w"), str(total))
    tblw.set(qn("w:type"), "dxa")

    indent = tblpr.find(qn("w:tblInd"))
    if indent is None:
        indent = OxmlElement("w:tblInd")
        tblpr.append(indent)
    indent.set(qn("w:w"), "0")
    indent.set(qn("w:type"), "dxa")

    for gridcol, ancho in zip(tabla_obj._tbl.tblGrid.findall(qn("w:gridCol")), anchos):
        gridcol.set(qn("w:w"), str(int(round(ancho * 1440))))

    for fila in tabla_obj.rows:
        for celda, ancho in zip(fila.cells, anchos):
            tcpr = celda._tc.get_or_add_tcPr()
            tcw = tcpr.find(qn("w:tcW"))
            if tcw is None:
                tcw = OxmlElement("w:tcW")
                tcpr.append(tcw)
            tcw.set(qn("w:w"), str(int(round(ancho * 1440))))
            tcw.set(qn("w:type"), "dxa")


def portada(doc, integrantes: list[tuple[str, str]], programa: str = "420",
            semestre: str = "5° semestre", grupo: str = "53",
            docente: str = "Juan Gabriel López Solórzano",
            fecha: str = "13 de noviembre de 2026") -> None:
    """Portada institucional en una hoja, con logos y datos editables."""
    logos = doc.add_table(rows=1, cols=2)
    logos.alignment = WD_TABLE_ALIGNMENT.CENTER
    logos.autofit = False
    _tabla_sin_bordes(logos)
    _margenes_celda(logos, 0, 0, 0, 0)
    celda_uanl, celda_facpya = logos.rows[0].cells
    celda_uanl.width = Inches(3.1)
    celda_facpya.width = Inches(3.1)
    for celda in (celda_uanl, celda_facpya):
        celda.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        celda.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        celda.paragraphs[0].paragraph_format.space_after = Pt(0)

    run_uanl = celda_uanl.paragraphs[0].add_run()
    run_uanl.add_picture(str(ASSETS / "logo_uanl.png"), width=Inches(1.75))
    run_facpya = celda_facpya.paragraphs[0].add_run()
    run_facpya.add_picture(str(ASSETS / "logo_facpya.png"), width=Inches(2.12))
    _fijar_anchos_tabla(logos, [3.1, 3.1])

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(1)
    run = p.add_run("UNIVERSIDAD AUTÓNOMA DE NUEVO LEÓN")
    _fuente(run, FUENTE_TEXTO, 12, negrita=True, color=RGB_CARBON)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(6)
    run = p.add_run("FACULTAD DE CONTADURÍA PÚBLICA Y ADMINISTRACIÓN")
    _fuente(run, FUENTE_TEXTO, 9.5, negrita=True, color=RGB_CARBON)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(12)
    _borde_inferior(p, DORADO, 12)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run("FUNDAMENTOS DE DESARROLLO WEB")
    _fuente(run, FUENTE_TEXTO, 10, negrita=True, color=RGB_GRIS)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(20)
    run = p.add_run("PRODUCTO INTEGRADOR DE APRENDIZAJE")
    _fuente(run, FUENTE_TEXTO, 9.5, color=RGB_GRIS)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.0
    run = p.add_run("LUMINA")
    _fuente(run, FUENTE_TITULO, 42, negrita=True, color=RGB_CARBON)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(6)
    run = p.add_run("Salón & Barbería")
    _fuente(run, FUENTE_TITULO, 15, cursiva=True, color=RGB_DORADO)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(14)
    p.paragraph_format.left_indent = Inches(0.6)
    p.paragraph_format.right_indent = Inches(0.6)
    run = p.add_run("Sitio web y sistema de gestión para un salón de belleza y barbería")
    _fuente(run, FUENTE_TEXTO, 10.5, color=RGB_GRIS)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(3)
    run = p.add_run("EQUIPO")
    _fuente(run, FUENTE_TEXTO, 8, negrita=True, color=RGB_DORADO)

    tabla_equipo = doc.add_table(rows=1, cols=2)
    tabla_equipo.alignment = WD_TABLE_ALIGNMENT.CENTER
    tabla_equipo.autofit = False
    _bordes_tabla(tabla_equipo, LINEA, 4)
    _margenes_celda(tabla_equipo, 20, 20, 70, 70)
    encabezados = ("Integrante", "Matrícula")
    for indice, texto in enumerate(encabezados):
        celda = tabla_equipo.rows[0].cells[indice]
        celda.width = Inches(4.9 if indice == 0 else 1.5)
        _sombra_celda(celda, CARBON)
        _texto_celda(celda, texto, negrita=True, tamano=8, color=BLANCO,
                     alineacion=WD_ALIGN_PARAGRAPH.CENTER)
    for indice, (nombre, matricula) in enumerate(integrantes):
        celdas = tabla_equipo.add_row().cells
        celdas[0].width = Inches(4.9)
        celdas[1].width = Inches(1.5)
        _texto_celda(celdas[0], nombre, tamano=8.5)
        _texto_celda(celdas[1], matricula, tamano=8.5,
                     alineacion=WD_ALIGN_PARAGRAPH.CENTER)
        celdas[0].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        celdas[1].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    _fijar_anchos_tabla(tabla_equipo, [4.9, 1.5])

    # Párrafo separador: sin él, Word fusiona ambas tablas en una sola.
    espacio(doc, 7)

    # La tabla de datos institucionales va sin bordes visibles.
    tabla_datos = doc.add_table(rows=4, cols=2)
    tabla_datos.alignment = WD_TABLE_ALIGNMENT.CENTER
    tabla_datos.autofit = False
    _tabla_sin_bordes(tabla_datos)
    _margenes_celda(tabla_datos, 10, 10, 0, 60)
    datos = [
        ("Programa educativo:", f"{programa} · {semestre} · Grupo {grupo}"),
        ("Docente:", docente),
        ("Periodo:", "Agosto – diciembre de 2026"),
        ("Lugar y fecha:", f"San Nicolás de los Garza, Nuevo León · {fecha}"),
    ]
    for fila, (etiqueta, valor) in enumerate(datos):
        celdas = tabla_datos.rows[fila].cells
        celdas[0].width = Inches(1.55)
        celdas[1].width = Inches(4.85)
        p0 = celdas[0].paragraphs[0]
        p0.paragraph_format.space_after = Pt(1)
        run = p0.add_run(etiqueta)
        _fuente(run, FUENTE_TEXTO, 8, negrita=True, color=RGB_GRIS)
        p1 = celdas[1].paragraphs[0]
        p1.paragraph_format.space_after = Pt(1)
        run = p1.add_run(valor)
        _fuente(run, FUENTE_TEXTO, 8.5, color=RGB_CARBON)
    _fijar_anchos_tabla(tabla_datos, [1.55, 4.85])


def indice(doc) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(3)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run("Índice")
    _fuente(run, FUENTE_TITULO, 22, negrita=True, color=RGB_CARBON)
    _borde_inferior(p, DORADO, 12)

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(10)
    run = p.add_run("Contenido general del documento")
    _fuente(run, FUENTE_TEXTO, 9, cursiva=True, color=RGB_GRIS)

    p = doc.add_paragraph()
    _campo(p, r'TOC \o "1-3" \h \z \u',
           "El índice se actualizará automáticamente al abrir el documento.")


def _campo(par, instruccion: str, texto_inicial: str = "1") -> None:
    """Inserta un campo de Word con cada instrucción en su propio run."""
    partes = (
        ("begin", None),
        ("instr", instruccion),
        ("separate", None),
        ("texto", texto_inicial),
        ("end", None),
    )
    for tipo, valor in partes:
        run = par.add_run()
        if tipo == "instr":
            el = OxmlElement("w:instrText")
            el.set(qn("xml:space"), "preserve")
            el.text = valor
        elif tipo == "texto":
            el = OxmlElement("w:t")
            el.text = valor
        else:
            el = OxmlElement("w:fldChar")
            el.set(qn("w:fldCharType"), tipo)
            if tipo == "begin":
                el.set(qn("w:dirty"), "true")
        run._r.append(el)


def encabezado_y_pie(seccion, texto_izquierda: str, texto_derecha: str) -> None:
    seccion.header.is_linked_to_previous = False
    seccion.footer.is_linked_to_previous = False

    encabezado = seccion.header
    p = encabezado.paragraphs[0]
    p.text = ""
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run(texto_izquierda)
    _fuente(run, FUENTE_TEXTO, 7.5, negrita=True, color=RGB_GRIS)
    run = p.add_run("\t" + texto_derecha)
    _fuente(run, FUENTE_TEXTO, 7.5, color=RGB_GRIS)
    _borde_inferior(p, DORADO, 6)

    pie = seccion.footer
    p = pie.paragraphs[0]
    p.text = ""
    run = p.add_run("UANL · FACPYA · Lumina")
    _fuente(run, FUENTE_TEXTO, 7.5, color=RGB_GRIS)
    run = p.add_run("\tPágina ")
    _fuente(run, FUENTE_TEXTO, 7.5, color=RGB_GRIS)
    _campo(p, "PAGE", "1")
    run = p.add_run(" de ")
    _fuente(run, FUENTE_TEXTO, 7.5, color=RGB_GRIS)
    _campo(p, "SECTIONPAGES", "1")
    for run in p.runs:
        _fuente(run, FUENTE_TEXTO, 7.5, color=RGB_GRIS)


def iniciar_cuerpo(doc) -> None:
    """Crea la sección del cuerpo (sin encabezado en portada) y numera desde 1."""
    seccion = doc.add_section(WD_SECTION.NEW_PAGE)
    seccion.top_margin = Cm(2.3)
    seccion.bottom_margin = Cm(2.1)
    seccion.left_margin = Cm(2.3)
    seccion.right_margin = Cm(2.3)
    encabezado_y_pie(
        seccion,
        "LUMINA · PRODUCTO INTEGRADOR DE APRENDIZAJE",
        "FUNDAMENTOS DE DESARROLLO WEB",
    )

    # Reinicia la numeración de página en la sección del cuerpo.
    sectpr = seccion._sectPr
    pg_num = sectpr.find(qn("w:pgNumType"))
    if pg_num is None:
        pg_num = OxmlElement("w:pgNumType")
        sectpr.append(pg_num)
    pg_num.set(qn("w:start"), "1")


def metadatos(doc, titulo: str, autor: str, asunto: str, palabras: str) -> None:
    props = doc.core_properties
    props.title = titulo
    props.author = autor
    props.subject = asunto
    props.keywords = palabras
    props.comments = "Proyecto integrador de Fundamentos de Desarrollo Web, UANL-FACPYA."
    props.category = "Producto Integrador de Aprendizaje"
