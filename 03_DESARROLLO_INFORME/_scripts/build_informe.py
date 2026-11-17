# -*- coding: utf-8 -*-
"""Genera el borrador del Informe Final de Prácticas (INVERSIONES JL) en DOCX.
Uso: python3 build_informe.py <salida.docx> <logo.png> <sunat.png> [paginas.json]
"""
import sys, json, re
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER, WD_COLOR_INDEX, WD_BREAK
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT, LOGO, SUNAT = sys.argv[1], sys.argv[2], sys.argv[3]
PAGES = json.load(open(sys.argv[4])) if len(sys.argv) > 4 else {}

FONT = "Arial"
doc = Document()

# ---------- estilos base ----------
st = doc.styles["Normal"]
st.font.name = FONT; st.font.size = Pt(12)
st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
pf = st.paragraph_format
pf.line_spacing = 1.5; pf.space_after = Pt(6); pf.space_before = Pt(0)
pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
for name, size in (("Heading 1", 12), ("Heading 2", 12), ("Heading 3", 12), ("Title", 12)):
    h = doc.styles[name]
    h.font.name = FONT; h.font.size = Pt(size); h.font.bold = True
    h.font.color.rgb = RGBColor(0, 0, 0); h.font.italic = False
    rpr = h.element.get_or_add_rPr(); rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts"); rpr.append(rf)
    for a in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rf.set(qn(a), FONT)
    for a in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
        if rf.get(qn(a)) is not None:
            del rf.attrib[qn(a)]
    h.paragraph_format.space_before = Pt(12); h.paragraph_format.space_after = Pt(6)
    h.paragraph_format.line_spacing = 1.5; h.paragraph_format.keep_with_next = True
    h.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT

def setup_section(sec):
    sec.page_width = Cm(21.0); sec.page_height = Cm(29.7)
    for m in ("top_margin", "bottom_margin", "left_margin", "right_margin"):
        setattr(sec, m, Cm(2.54))

setup_section(doc.sections[0])

# ---------- utilidades ----------
TOC = []      # (nivel, texto)
TABLES = []   # titulos de tablas
FIGS = []
counters = {"tabla": 0, "figura": 0}

def add_runs(p, text, bold=False, italic=False, size=None):
    """[[texto]] = resaltado amarillo (pendiente). **x** = negrita. *x* = cursiva."""
    parts = re.split(r"(\[\[.*?\]\]|\*\*.*?\*\*|\*[^*]+?\*)", text)
    for part in parts:
        if not part:
            continue
        b, i, hl = bold, italic, False
        if part.startswith("[[") and part.endswith("]]"):
            part = part[2:-2]; hl = True
        elif part.startswith("**") and part.endswith("**"):
            part = part[2:-2]; b = True
        elif part.startswith("*") and part.endswith("*") and len(part) > 2:
            part = part[1:-1]; i = True
        r = p.add_run(part); r.bold = b; r.italic = i
        if size: r.font.size = Pt(size)
        if hl: r.font.highlight_color = WD_COLOR_INDEX.YELLOW
    return p

def P(text, align=None, bold=False, italic=False, indent=True, size=None, after=None, spacing=None):
    p = doc.add_paragraph()
    add_runs(p, text, bold, italic, size)
    if align is not None: p.alignment = align
    if indent and align is None: p.paragraph_format.first_line_indent = Cm(1.25)
    if after is not None: p.paragraph_format.space_after = Pt(after)
    if spacing is not None: p.paragraph_format.line_spacing = spacing
    return p

def B(text, level=0):
    p = doc.add_paragraph(style="List Bullet")
    add_runs(p, text)
    p.paragraph_format.left_indent = Cm(1.0 + 0.6 * level)
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    for r in p.runs:
        r.font.name = FONT; r.font.size = Pt(12)
    return p

def H(level, text):
    h = doc.add_heading(text, level=level)
    TOC.append((level, text))
    return h

def shade(cell, color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd"); shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), color)
    tcPr.append(shd)

def apa_table(title, header, rows, widths=None, note=None, size=10, header_fill="D9D9D9", label="Tabla"):
    key = "tabla" if label == "Tabla" else "figura"
    counters[key] += 1
    n = counters[key]
    (TABLES if label == "Tabla" else FIGS).append(f"{label} {n}. {title}")
    p = doc.add_paragraph(); p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.space_after = Pt(0); p.paragraph_format.keep_with_next = True
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r = p.add_run(f"{label} {n}"); r.bold = True
    p2 = doc.add_paragraph(); p2.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p2.paragraph_format.keep_with_next = True; p2.paragraph_format.space_after = Pt(4)
    add_runs(p2, title, italic=True)
    t = doc.add_table(rows=1 if header else 0, cols=len(header) if header else len(rows[0]))
    t.style = "Table Grid"; t.alignment = WD_TABLE_ALIGNMENT.CENTER
    if header:
        for i, htxt in enumerate(header):
            c = t.rows[0].cells[i]; c.text = ""
            pp = c.paragraphs[0]; add_runs(pp, htxt, bold=True, size=size)
            pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pp.paragraph_format.line_spacing = 1.0; pp.paragraph_format.first_line_indent = Cm(0)
            pp.paragraph_format.space_after = Pt(0)
            shade(c, header_fill)
    for row in rows:
        cells = t.add_row().cells
        for i, val in enumerate(row):
            c = cells[i]; c.text = ""
            lines = val.split("\n") if isinstance(val, str) else [str(val)]
            for k, line in enumerate(lines):
                pp = c.paragraphs[0] if k == 0 else c.add_paragraph()
                add_runs(pp, line, size=size)
                pp.alignment = WD_ALIGN_PARAGRAPH.LEFT
                pp.paragraph_format.line_spacing = 1.0; pp.paragraph_format.first_line_indent = Cm(0)
                pp.paragraph_format.space_after = Pt(0)
    if widths:
        for row in t.rows:
            for i, w in enumerate(widths):
                row.cells[i].width = Cm(w)
    if note:
        pn = doc.add_paragraph(); pn.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pn.paragraph_format.line_spacing = 1.0; pn.paragraph_format.space_before = Pt(4)
        pn.paragraph_format.space_after = Pt(10)
        r = pn.add_run("Nota. "); r.italic = True; r.font.size = Pt(10)
        add_runs(pn, note, size=10)
    else:
        doc.add_paragraph().paragraph_format.space_after = Pt(4)
    return t

def page_number_footer(sec, fmt, start):
    sec.footer.is_linked_to_previous = False
    fp = sec.footer.paragraphs[0]; fp.text = ""
    fp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    if fmt:
        r = fp.add_run()
        for typ, txt in (("begin", None), (None, " PAGE "), ("end", None)):
            if typ:
                f = OxmlElement("w:fldChar"); f.set(qn("w:fldCharType"), typ); r._r.append(f)
            else:
                it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = txt; r._r.append(it)
        r.font.name = FONT; r.font.size = Pt(11)
    sectPr = sec._sectPr
    pg = sectPr.find(qn("w:pgNumType"))
    if pg is None:
        pg = OxmlElement("w:pgNumType"); sectPr.append(pg)
    if fmt: pg.set(qn("w:fmt"), fmt)
    if start: pg.set(qn("w:start"), str(start))

def new_section():
    sec = doc.add_section(WD_SECTION.NEW_PAGE); setup_section(sec); return sec

def toc_entry(paragraph_text, page, level, bold=False):
    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.left_indent = Cm(0.6 * (level - 1))
    p.paragraph_format.line_spacing = 1.15; p.paragraph_format.space_after = Pt(2)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.tab_stops.add_tab_stop(Cm(15.9), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
    r = p.add_run(f"{paragraph_text}\t{page}"); r.bold = bold; r.font.size = Pt(11)
    return p

def page_break():
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

# =====================================================================
# CARÁTULA (modelo de la plantilla de la Escuela)
# =====================================================================
sec0 = doc.sections[0]
page_number_footer(sec0, None, None)
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.add_run().add_picture(LOGO, width=Cm(11.5))
for txt, sz in (("FACULTAD DE CIENCIAS EMPRESARIALES", 14),
                ("ESCUELA PROFESIONAL DE ADMINISTRACIÓN Y NEGOCIOS INTERNACIONALES", 13)):
    P(txt, align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=sz, after=2)
P("", align=WD_ALIGN_PARAGRAPH.CENTER, after=0)
P("INFORME DE PRÁCTICA PREPROFESIONAL TERMINAL II", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=13, after=2)
P("INVERSIONES JL", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=14, after=8)
box = doc.add_table(rows=1, cols=1); box.style = "Table Grid"; box.alignment = WD_TABLE_ALIGNMENT.CENTER
c = box.rows[0].cells[0]; c.width = Cm(14)
lines = [("INFORME ACADÉMICO:", True, 12),
         ("INFORME FINAL DE PRÁCTICAS PREPROFESIONALES TERMINAL II", True, 12),
         ("[[PROPUESTA DE MEJORA DEL CONTROL DE EXISTENCIAS Y DE LA REPOSICIÓN DE PRODUCTOS EN EL ÁREA DE GESTIÓN COMERCIAL Y ADMINISTRATIVA]]", True, 12),
         ("REALIZADO EN INVERSIONES JL", True, 12)]
for k, (t_, b_, s_) in enumerate(lines):
    pp = c.paragraphs[0] if k == 0 else c.add_paragraph()
    pp.alignment = WD_ALIGN_PARAGRAPH.CENTER; pp.paragraph_format.first_line_indent = Cm(0)
    pp.paragraph_format.space_after = Pt(6)
    add_runs(pp, t_, bold=b_, size=s_)
P("", after=0)
for lab, val in (("Practicante:", "Melendrez Jaramillo, Jorge Junior"),
                 ("Docente:", "Alarcón Ato, José Antonio"),
                 ("Jefe inmediato:", "García Guevara, Lea Lucía – Propietaria")):
    P(lab, align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, after=0)
    P(val, align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, italic=True, after=6)
P("2026", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, after=0)
P("Piura – Perú", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, after=0)

# =====================================================================
# ÍNDICE (sección con numeración romana, inicia en ii)
# =====================================================================
sec1 = new_section(); page_number_footer(sec1, "lowerRoman", 2)
P("ÍNDICE", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True)
TOC_ANCHOR = doc.add_paragraph()   # se reemplaza al final
TOC_ANCHOR_T = doc.add_paragraph()

# =====================================================================
# CUERPO (numeración arábiga desde 1)
# =====================================================================
sec2 = new_section(); page_number_footer(sec2, "decimal", 1)

# ---------------- I. RESUMEN EJECUTIVO ----------------
H(1, "I. RESUMEN EJECUTIVO")
P("El presente informe describe la Práctica Preprofesional Terminal II que desarrollé en INVERSIONES JL, nombre comercial del negocio de la persona natural Lea Lucía García Guevara (RUC 10719941023), dedicado a la venta al por menor de artículos de ferretería, pinturas y materiales de construcción en el distrito de Veintiséis de Octubre, Piura. Las prácticas se iniciaron el 20 de marzo de 2026, en el área de gestión comercial y administrativa, con funciones de atención al cliente, apoyo en ventas, manejo y balance diario de caja, revisión de productos, identificación de productos con bajo nivel de existencias y comunicación de necesidades de reposición.")
P("La problemática principal identificada es una reposición reactiva de mercadería: aunque el sistema de ventas Rapifac registra el stock, este se encuentra desactualizado; la revisión se realiza de forma visual en los estantes y un producto se considera “bajo” recién cuando se agota. Como consecuencia, se presenta en promedio un caso diario de producto solicitado y no disponible, con mayor incidencia en cemento y conexiones de tubería; dado que los pedidos se realizan una vez por semana y demoran cuatro días en promedio, un producto agotado puede permanecer sin stock entre 4 y 10 días hábiles, periodo en el cual parte de los clientes acude a las dos ferreterías competidoras cercanas. De manera complementaria, se observó que las diferencias del cierre diario de caja (sobrantes de S/ 50 a S/ 100 y faltantes ocasionales) no se registran ni se analizan.")
P("Como alternativa de solución se propone un sistema simple de control de existencias y reposición, sin inversión en software adicional, que comprende: la clasificación ABC de los productos, la definición de stock mínimo y nivel máximo para los productos de clase A, la actualización y verificación periódica del stock en Rapifac, una revisión programada previa al pedido semanal, una lista estandarizada de reposición y un registro de la demanda no atendida. Asimismo, se propone un formato de arqueo de caja que registre y explique las diferencias. Se concluye que la propuesta es viable por su bajo costo y por basarse en recursos que la empresa ya posee, y se recomienda implementarla de forma gradual, empezando por los productos de mayor rotación, para medir su efecto en la reducción de quiebres de stock.")

# ---------------- II. ASPECTOS GENERALES ----------------
H(1, "II. ASPECTOS GENERALES DE LA EMPRESA")
H(2, "2.1. Razón social")
P("De acuerdo con la consulta RUC de la Superintendencia Nacional de Aduanas y de Administración Tributaria [SUNAT] (2026), el contribuyente es una **persona natural con negocio** cuyo titular es **García Guevara, Lea Lucía**, y su **nombre comercial** es **INVERSIONES JL**, denominación que figura también en el letrero del local y en los comprobantes de pago. Por tratarse de una persona natural, el negocio no tiene una razón social de persona jurídica; en los documentos legales figura el nombre de su titular.")
H(2, "2.2. Ubicación geográfica")
P("El establecimiento se ubica en el A.H. Luis Paredes Maceda, Mz. E, Lote 9, distrito de Veintiséis de Octubre, provincia y departamento de Piura, Perú. La empresa no cuenta con sucursales ni con almacenes independientes: la mercadería se exhibe y almacena en el mismo local de venta. En la consulta RUC no figura consignado un domicilio fiscal (SUNAT, 2026).")
H(2, "2.3. Objeto social")
P("Al ser una persona natural con negocio, no existe un estatuto que defina un objeto social. En su lugar, se toma como referencia la actividad económica principal registrada en SUNAT (2026): **CIIU 4752 – Venta al por menor de artículos de ferretería, pinturas y productos de vidrio en comercios especializados**. En la misma consulta se registra la actividad de comercio exterior como “sin actividad”, por lo que la empresa no realiza operaciones directas de importación ni exportación.")
H(2, "2.4. RUC")
P("El Registro Único de Contribuyentes es **10719941023**, con estado **activo** y condición **habido**. El contribuyente está inscrito como emisor electrónico desde el 10 de febrero de 2025 y emite facturas y boletas electrónicas; su sistema de contabilidad es manual (SUNAT, 2026).")
H(2, "2.5. Reseña histórica")
P("Según la consulta RUC, la fecha de inscripción y de inicio de actividades es el 9 de octubre de 2017 (SUNAT, 2026), año desde el cual funciona la ferretería. El hecho más relevante de su evolución reciente fue la incorporación de la facturación electrónica en febrero de 2025, que se realiza mediante el sistema Rapifac. Desde su inicio, el negocio no ha cambiado de local ni ha incorporado nuevas líneas de productos, y es dirigido por su propietaria junto con su esposo. [[Si la propietaria lo autoriza, conviene añadir una o dos líneas sobre cómo empezó el negocio y sus principales logros.]]")
P("El negocio forma parte del segmento de micro y pequeñas empresas, que en 2024 representaba el 99,3 % de las empresas formales del Perú (Ministerio de la Producción [PRODUCE], 2025). Estas empresas se categorizan según sus ventas anuales (Decreto Supremo N.° 013-2013-PRODUCE, 2013); en este informe no se precisa la categoría de INVERSIONES JL porque no se cuenta con información de ventas autorizada para su publicación.")
H(2, "2.6. Actividades que realiza la empresa")
P("INVERSIONES JL comercializa al por menor, al contado, entre 300 y 400 artículos distintos, agrupados en las siguientes líneas:")
apa_table("Líneas de productos comercializados por INVERSIONES JL",
          ["Línea", "Ejemplos de productos"],
          [["Materiales de construcción", "Cemento, autoperforantes para calaminón"],
           ["Gasfitería y fontanería", "Tubos y conexiones de tubería"],
           ["Materiales eléctricos", "Tomacorrientes, interruptores y accesorios"],
           ["Pinturas", "Pinturas y complementos"],
           ["Herramientas y otros", "Herramientas manuales y otros artículos de ferretería"]],
          widths=[5.5, 10], note="Elaboración propia con información de la empresa. Los productos de mayor venta son tubos, cemento, autoperforantes para calaminón, tomacorrientes e interruptores.")
P("Sus clientes son principalmente familias de la zona, maestros de obra, gasfiteros y electricistas, con mayor afluencia entre las 8:00 y las 10:00 y entre las 14:00 y las 15:00. Los pagos se reciben en efectivo y mediante la billetera digital Yape; no se otorgan créditos ni se prestan servicios adicionales como reparto a domicilio. De este modo, la empresa contribuye al abastecimiento de materiales para la construcción, el mantenimiento y las reparaciones de viviendas en su entorno.")

# ---------------- III. ORGANIZACIÓN ----------------
H(1, "III. ORGANIZACIÓN DE LA EMPRESA")
H(2, "3.1. Estructura organizacional de la empresa")
P("INVERSIONES JL presenta una estructura simple, característica de los negocios pequeños, en la que la autoridad se concentra en el propietario, existen pocos niveles jerárquicos y escasa formalización (Robbins et al., 2025). No existen áreas o departamentos independientes: las funciones comerciales y administrativas las realizan sus tres integrantes, como se muestra en la Figura 1.")
fig = apa_table("Organigrama de INVERSIONES JL", None,
          [["", "PROPIETARIA\nLea Lucía García Guevara\n(Jefa inmediata)", ""],
           ["", "↓", ""],
           ["", "JEFE\n(esposo de la propietaria; apoyo en la dirección del negocio)", ""],
           ["", "↓", ""],
           ["", "PRACTICANTE\nGestión comercial y administrativa", ""]],
          widths=[3.0, 9.0, 3.0], note="Elaboración propia. La relación jefe–practicante se representa como apoyo a la dirección del negocio; las decisiones de compra las toman la propietaria y el jefe.",
          label="Figura", size=10)
for i, row in enumerate(fig.rows):
    for j, cell in enumerate(row.cells):
        for pp in cell.paragraphs: pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if j != 1 or i in (1, 3):
            tcPr = cell._tc.get_or_add_tcPr(); b = OxmlElement("w:tcBorders")
            for side in ("top", "left", "bottom", "right"):
                e = OxmlElement(f"w:{side}"); e.set(qn("w:val"), "nil"); b.append(e)
            tcPr.append(b)
        elif j == 1:
            shade(cell, "DEEAF6")
H(2, "3.2. Descripción de las funciones de cada una de las áreas de la empresa")
P("Al no existir áreas formales, las funciones se describen por integrante:")
apa_table("Funciones de los integrantes de INVERSIONES JL",
          ["Integrante", "Funciones principales"],
          [["Propietaria", "Dirige el negocio; decide, junto con el jefe, qué productos comprar y en qué cantidad; recibe diariamente el dinero y el resultado del balance de caja; atiende y vende; aplica descuentos."],
           ["Jefe", "Decide las compras junto con la propietaria y apoya en la conducción del negocio. [[Precisar, si se desea, otras funciones habituales.]]"],
           ["Practicante", "Atiende a los clientes, registra las ventas en Rapifac y cobra; realiza el balance diario de caja; revisa los productos disponibles; comunica por WhatsApp los productos agotados; recibe, verifica e ingresa al sistema la mercadería."]],
          widths=[3.5, 12], note="Elaboración propia a partir de la observación y la información de la empresa.")
H(2, "3.3. Funciones empresariales en el área donde se desempeña el practicante")
P("Me desempeño en el área de gestión comercial y administrativa, de lunes a sábado, de 7:30 a 16:30. Mis funciones se agrupan en tres procesos:")
P("**a) Atención al cliente y apoyo en ventas.** Cuando un cliente solicita productos, registro en Rapifac cada artículo pedido; si el cliente no sabe exactamente qué necesita, le hago preguntas para entender su necesidad (por ejemplo, el tipo de tubo o de interruptor). Luego le indico el precio de cada artículo y el total, consulto si desea boleta, recibo el pago en efectivo o por Yape y preparo y entrego la mercadería. Si un producto no está disponible, se lo comunico y le ofrezco alternativas. Puedo aplicar descuentos, que registro en el sistema, porque Rapifac muestra el precio de compra y permite verificar que el precio final no quede por debajo del costo.")
P("**b) Manejo y balance de caja.** La caja se inicia cada día con un fondo de S/ 100. No se realizan pagos ni egresos desde la caja. Al finalizar el día, realizo el balance: descuento el fondo inicial, sumo el efectivo y los pagos recibidos por Yape, y comparo el total con el reporte de ventas que genera Rapifac en Excel. Entrego el resultado y el dinero diariamente a la propietaria.")
P("**c) Revisión de existencias y reposición.** Reviso de forma visual los productos en los estantes cuando tengo dudas sobre su disponibilidad; cuando observo que un producto se agotó, lo comunico a los jefes por WhatsApp, y ellos deciden qué se compra y en qué cantidad. Los pedidos se realizan una vez por semana a proveedores de Piura y de Lima; cuando llega la mercadería, la recibo, la verifico y la ingreso en Rapifac.")
P("Respecto de la evolución de mis funciones, durante las primeras semanas me dediqué a conocer los productos; desde la segunda semana asumí el manejo de la caja y el balance diario. Las principales dificultades iniciales fueron aprender los precios y comprender qué productos buscaban los clientes, lo cual superé con la práctica diaria y el apoyo del sistema de ventas.")

# ---------------- IV. PLAN ESTRATÉGICO ----------------
H(1, "IV. DESARROLLO DEL PLAN ESTRATÉGICO DE LA EMPRESA")
P("[[INVERSIONES JL no cuenta con una visión, una misión ni valores formalmente redactados (confirmar).]] Por ello, conforme lo indica la estructura del informe, se presenta una propuesta elaborada a partir de la actividad real del negocio. La visión describe lo que la organización quiere llegar a ser y la misión expresa su razón de ser, sus clientes y sus productos (David et al., 2023).")
H(2, "4.1. Visión")
P("*Propuesta:* “Ser reconocida al 2030 como la ferretería de confianza del distrito Veintiséis de Octubre, por tener siempre disponibles los productos que necesitan las familias y los maestros de obra, y por brindar una atención rápida y bien orientada”.")
P("Análisis: la visión propuesta es alcanzable para el tamaño del negocio, se limita a su ámbito geográfico y está vinculada al principal reto identificado, que es la disponibilidad de productos.")
H(2, "4.2. Misión")
P("*Propuesta:* “Somos una ferretería que abastece de materiales de construcción, gasfitería, electricidad, pinturas y herramientas a familias, maestros de obra, gasfiteros y electricistas de nuestra zona, ofreciendo orientación en la compra, precios justos y comprobantes electrónicos, con un equipo comprometido con el servicio”.")
P("Análisis: la misión identifica clientes, productos, mercado y propuesta de valor, que son los componentes que recomienda incluir David et al. (2023).")
H(2, "4.3. Valores esenciales")
P("A partir del comportamiento observado en el negocio, se proponen los siguientes valores:")
B("**Honestidad:** registrar todas las ventas en el sistema, con o sin boleta, y comunicar con transparencia los faltantes de caja.")
B("**Orientación al cliente:** preguntar para entender la necesidad del cliente y ofrecer alternativas cuando un producto no está disponible.")
B("**Responsabilidad:** cumplir el cierre diario de caja y la verificación de la mercadería recibida.")
B("**Precio justo:** otorgar descuentos razonables, verificando el costo del producto en el sistema.")
B("**Trabajo en equipo:** coordinación permanente entre la propietaria, el jefe y el practicante.")
H(2, "4.4. Diagnóstico externo e interno de la empresa")
H(3, "4.4.1. Matriz FODA: Fortalezas, oportunidades, debilidades y amenazas")
P("El análisis FODA permite evaluar los factores internos (fortalezas y debilidades) y externos (oportunidades y amenazas) que inciden en la empresa (David et al., 2023). Para INVERSIONES JL se elaboró a partir de la observación directa durante las prácticas, la consulta RUC y fuentes secundarias:")
apa_table("Matriz FODA de INVERSIONES JL",
          ["Fortalezas (internas)", "Debilidades (internas)"],
          [["F1. Formalidad tributaria: RUC activo y habido; emisión de boletas y facturas electrónicas.\nF2. Sistema Rapifac con precios, costos y control de stock disponible.\nF3. Variedad de 300 a 400 artículos en líneas complementarias.\nF4. Atención orientada: se pregunta al cliente y se ofrecen alternativas.\nF5. Ventas al contado: sin riesgo de cobranza.\nF6. Trayectoria desde 2017 y proveedores recomendados por una ferretería de la familia.",
            "D1. Stock del sistema desactualizado.\nD2. Sin stock mínimo: se repone cuando el producto ya se agotó.\nD3. Revisión de existencias visual y no programada.\nD4. Diferencias de caja sin registro ni análisis.\nD5. Procedimientos no documentados.\nD6. Sin servicios complementarios (por ejemplo, reparto)."]],
          widths=[7.9, 7.9], note="Elaboración propia.")
apa_table("Matriz FODA de INVERSIONES JL (continuación)",
          ["Oportunidades (externas)", "Amenazas (externas)"],
          [["O1. Uso creciente de billeteras digitales como Yape, que facilitan los cobros y el registro de operaciones (Barrantes y Alzamora, 2023).\nO2. Clientes recurrentes por oficio (maestros de obra, gasfiteros, electricistas).\nO3. Herramientas digitales de bajo costo para comunicarse con clientes y proveedores.\nO4. Red familiar con otra ferretería para referir proveedores y coordinar compras.\nO5. Proveedores disponibles en Piura y Lima.",
            "A1. Dos ferreterías competidoras cercanas.\nA2. Ante un faltante, los clientes tienden a cambiar de tienda, sobre todo cuando hay competidores cercanos (Azeem et al., 2019; Fitzsimons, 2000).\nA3. Tiempo de entrega de cuatro días en promedio y dependencia de proveedores externos.\nA4. Riesgos del transporte de mercadería desde Lima (demoras o daños)."]],
          widths=[7.9, 7.9], note="Elaboración propia con base en fuentes citadas.")
H(2, "4.5. Objetivos estratégicos")
P("Los objetivos estratégicos deben ser claros, medibles y coherentes con la visión (David et al., 2023). Se proponen los siguientes, con un horizonte de seis meses desde su implementación:")
apa_table("Objetivos estratégicos propuestos",
          ["N.°", "Objetivo", "Indicador", "Línea base"],
          [["OE1", "Reducir los casos de producto agotado en los artículos de mayor rotación.", "Casos de demanda no atendida por semana", "≈ 6 por semana (1 diario)"],
           ["OE2", "Mantener actualizado el stock del sistema Rapifac.", "% de artículos clase A cuyo stock del sistema coincide con el conteo físico", "Desactualizado (no medido)"],
           ["OE3", "Registrar y explicar todas las diferencias de caja.", "% de días con diferencia registrada y analizada", "0 %"],
           ["OE4", "Fortalecer la relación con clientes por oficio.", "N.° de clientes frecuentes registrados", "Sin registro"]],
          widths=[1.2, 6.0, 5.0, 3.6], note="Elaboración propia. Las líneas base proceden de la observación y la información del practicante.")
H(2, "4.6. Planteamiento de Estrategias")
P("A partir del cruce de los factores de la matriz FODA (David et al., 2023), se plantean las siguientes estrategias:")
apa_table("Estrategias derivadas de la matriz FODA",
          ["Tipo", "Estrategia", "Objetivo"],
          [["FO (F2, O3)", "Usar el módulo de stock de Rapifac y WhatsApp para estandarizar la lista de reposición semanal.", "OE1, OE2"],
           ["FO (F4, O2)", "Registrar a los clientes por oficio y comunicarles la llegada de productos o novedades.", "OE4"],
           ["DO (D2, O4)", "Definir stock mínimo y nivel máximo de los productos clase A y coordinar compras con la ferretería familiar ante urgencias.", "OE1"],
           ["FA (F3, A2)", "Asegurar la disponibilidad de los productos de mayor rotación para no perder clientes frente a la competencia.", "OE1"],
           ["DA (D4, A1)", "Implementar un formato de arqueo de caja que registre y explique las diferencias.", "OE3"]],
          widths=[2.6, 10.2, 3.0], note="Elaboración propia.")

# ---------------- V. GESTIÓN INTERNA ----------------
H(1, "V. GESTIÓN INTERNA, PRODUCTIVIDAD Y ANÁLISIS FINANCIERO")
H(2, "5.1. Gestión de Recursos y Sostenibilidad")
H(3, "5.1.1. Evaluación integral de recursos disponibles (Materiales, humanos, tecnológicos y financieros)")
apa_table("Recursos disponibles en el área de gestión comercial y administrativa",
          ["Tipo", "Recursos", "Evaluación"],
          [["Materiales", "Local de venta con estantería y exhibición; mercadería de 300 a 400 artículos.", "Adecuados; el mismo local funciona como almacén."],
           ["Humanos", "Propietaria, jefe y practicante.", "Equipo reducido; la capacitación es práctica, en el puesto de trabajo."],
           ["Tecnológicos", "Computadora con el sistema Rapifac (ventas, precios, costos, stock y comprobantes electrónicos), reportes en Excel, Yape y WhatsApp.", "Subutilizados: el control de stock del sistema no se mantiene actualizado."],
           ["Financieros", "Ventas al contado; fondo diario de caja de S/ 100.", "Liquidez diaria; no se otorga crédito."]],
          widths=[2.6, 7.2, 6.0], note="Elaboración propia. No se incluyen montos de ventas ni presupuestos por tratarse de información reservada de la empresa.")
H(3, "5.1.2. Identificación de áreas de mejora para la generación de valor y responsabilidad social empresarial")
P("El recurso con mayor potencial no aprovechado es el sistema Rapifac: ya permite controlar el stock, pero sus datos no coinciden con las existencias reales. La exactitud de los registros de inventario es una condición para que cualquier sistema de reposición funcione (Heizer et al., 2023). Una segunda área de mejora es el control de caja, porque los sobrantes y faltantes no se documentan, lo que debilita las actividades de control y la información para la toma de decisiones que propone el marco de control interno (Committee of Sponsoring Organizations of the Treadway Commission [COSO], 2013).")
P("En cuanto a la responsabilidad social, la empresa puede generar valor para su comunidad asegurando la disponibilidad de los materiales que los vecinos necesitan para mejorar sus viviendas, orientando con honestidad la compra (sin vender productos innecesarios), entregando comprobantes electrónicos y evitando pérdidas de mercadería por un mal almacenamiento.")
H(2, "5.2. Optimización de la Productividad Operativa")
H(3, "5.2.1. Evaluación crítica de los procesos actuales")
P("Se analizó el proceso de **control de existencias y reposición**, por su relación directa con mis funciones y con la pérdida de ventas. La Tabla 8 resume el proceso actual y sus puntos críticos.")
apa_table("Proceso actual de control de existencias y reposición",
          ["Paso", "Actividad actual", "Punto crítico (cuello de botella, demora o reproceso)"],
          [["1", "Venta registrada en Rapifac.", "El stock del sistema baja, pero no coincide con el físico porque no se verifica."],
           ["2", "Revisión visual de estantes, solo cuando hay dudas.", "No hay revisión programada; el faltante se descubre frente al cliente."],
           ["3", "Se identifica el producto “bajo” cuando ya no hay.", "No existe stock mínimo: la detección es tardía."],
           ["4", "Aviso por WhatsApp a los jefes.", "Mensajes sueltos, sin formato ni cantidades sugeridas."],
           ["5", "Los jefes deciden la compra; pedido semanal.", "Un faltante detectado después del pedido espera hasta la siguiente semana."],
           ["6", "Entrega del proveedor (≈ 4 días); recepción, verificación e ingreso en Rapifac por el practicante.", "Producto agotado entre 4 y 10 días hábiles."]],
          widths=[1.2, 6.8, 7.8], note="Elaboración propia a partir de la observación durante las prácticas.")
P("El efecto se observa en la atención: en promedio, una vez al día un cliente solicita un producto que no está disponible, con mayor frecuencia cemento y conexiones de tubería. Algunos clientes aceptan una alternativa y otros no, lo que representa ventas perdidas. Adicionalmente, en el **proceso de cierre de caja** se observan sobrantes frecuentes de S/ 50 a S/ 100 y faltantes ocasionales que no se registran ni se investigan; dado que todas las ventas y los descuentos se registran en Rapifac y que el fondo inicial se descuenta, la causa aún no está identificada.")
H(3, "5.2.2. Propuestas de mejora continua para la eficiencia organizacional")
B("**Revisión programada de existencias:** realizar, el día anterior al pedido semanal, una revisión de los productos de clase A comparando el estante con el stock de Rapifac. Ahorra tiempo porque se revisan solo los artículos clave y no todo el local.")
B("**Lista estandarizada de reposición:** enviar por WhatsApp una sola lista semanal con el producto, el stock actual, el stock mínimo y la cantidad sugerida (Anexo 3). Elimina los mensajes dispersos y agiliza la decisión de compra de los jefes.")
B("**Ajuste de stock en el sistema:** registrar en Rapifac las diferencias encontradas en la revisión, para que el sistema vuelva a ser confiable (Heizer et al., 2023).")
B("**Arqueo de caja con registro de diferencias:** usar un formato diario (Anexo 4) que separe el efectivo, el Yape y las ventas por medio de pago según Rapifac, y que anote la diferencia y su posible explicación. Al cabo de algunas semanas, permitirá identificar la causa de los sobrantes.")
H(2, "5.3. Evaluación Financiera para la Toma de Decisiones")
H(3, "5.3.1. Análisis de la información financiera vinculada a las operaciones internacionales")
P("INVERSIONES JL no realiza operaciones internacionales (SUNAT, 2026), por lo que no maneja presupuestos de importación ni exportación. Por ello, se analiza la estructura de costos de las operaciones del área en términos cualitativos, ya que la información de montos es reservada:")
apa_table("Estructura de costos del área de gestión comercial y administrativa",
          ["Tipo de costo", "Concepto", "Comportamiento"],
          [["Variable", "Costo de la mercadería vendida", "Varía con el volumen de ventas; es el principal costo del negocio."],
           ["Variable", "Pérdidas por faltantes de caja", "Ocasionales; no se registran."],
           ["Fijo", "Servicio del sistema de facturación electrónica (Rapifac)", "Periódico. [[Monto no informado.]]"],
           ["Fijo", "Servicios del local (energía eléctrica, internet, telefonía)", "Periódico. [[Monto no informado.]]"],
           ["Costo de oportunidad", "Ventas perdidas por productos agotados", "≈ 1 caso diario; no se mide su valor."]],
          widths=[3.2, 6.6, 6.0], note="Elaboración propia. Se incluyen solo conceptos observados; no se presentan montos sin autorización de la empresa.")
H(3, "5.3.2. Conclusiones sobre la viabilidad económica y toma de decisiones (con actitud ética y crítica)")
P("Las operaciones actuales son sostenibles porque las ventas son al contado y el negocio no asume riesgo de cobranza. Sin embargo, la reposición reactiva genera un costo oculto: cada venta perdida reduce el margen y puede trasladar clientes a la competencia. Las mejoras propuestas en el punto 5.2 son económicamente viables porque no requieren comprar software ni equipos: utilizan Rapifac, Excel y WhatsApp, que ya existen, y demandan principalmente tiempo del personal, que puede programarse en las horas de menor afluencia (fuera de 8:00–10:00 y 14:00–15:00). Desde una perspectiva ética, no se estiman montos de ahorro sin datos reales; por ello, la propuesta incluye un registro de demanda no atendida que permitirá cuantificar, con datos de la propia empresa, el valor de las ventas recuperadas.")

# ---------------- VI. LOGÍSTICA ----------------
H(1, "VI. LOGÍSTICA INTERNACIONAL Y CALIDAD DE LA OFERTA EXPORTABLE")
P("Según la consulta RUC, la actividad de comercio exterior de la empresa es “sin actividad” (SUNAT, 2026). Por tanto, INVERSIONES JL no exporta, no tiene oferta exportable ni realiza trámites aduaneros. En concordancia con ello, este capítulo analiza la cadena de suministro real del negocio, que tiene un vínculo internacional indirecto: parte de los productos que comercializa son de origen chino y los adquiere a un distribuidor peruano que los importa.")
H(2, "6.1. Diagnóstico de la Logística Internacional")
H(3, "6.1.1. Evaluación de las estrategias de distribución física internacional de la organización")
P("La cadena de suministro comprende los flujos de productos, información y dinero entre proveedores, el negocio y el cliente final (Chopra, 2025). En INVERSIONES JL, la cadena es la siguiente: fabricantes nacionales y extranjeros → distribuidores e importadores de Lima y proveedores de Piura → INVERSIONES JL → clientes finales. Los proveedores fueron seleccionados por recomendación de una ferretería de la familia, los pedidos se realizan una vez por semana y la mercadería llega en un promedio de cuatro días. Al comprar en el mercado nacional, la empresa no negocia Incoterms ni contrata transporte internacional: esos riesgos y costos los asume el distribuidor importador, y se reflejan indirectamente en el precio de compra.")
H(3, "6.1.2. Métodos de aseguramiento de la calidad de la oferta exportable")
P("La empresa no cuenta con certificaciones de calidad. El control de calidad se realiza en la **recepción de mercadería**: recibo los productos, verifico que las cantidades y los productos coincidan con lo solicitado y con el comprobante del proveedor, y los ingreso en Rapifac. Para fortalecer este control se propone una lista de verificación en la recepción (cantidad, estado del producto y del empaque, coincidencia con el comprobante) y el registro de incidencias para reclamar al proveedor.")
H(2, "6.2. Optimización de la Cadena de Suministro")
H(3, "6.2.1. Propuestas para la reducción de costos logísticos y optimización de tiempos de entrega")
B("**Pedido semanal planificado:** basar el pedido en la lista de reposición y en los niveles mínimos, para consolidar en un solo pedido lo que hoy se pide de forma reactiva.")
B("**Seleccionar al proveedor según urgencia:** atender las urgencias de productos clase A con proveedores de Piura y los pedidos programados con proveedores de Lima.")
B("**Compras coordinadas con la ferretería familiar** para productos de baja rotación, de modo que se reduzcan los pedidos pequeños.")
H(3, "6.2.2. Criterios de responsabilidad y mitigación de riesgos en el tránsito internacional")
apa_table("Riesgos de abastecimiento y plan de contingencia",
          ["Riesgo", "Efecto", "Medida de contingencia"],
          [["Demora del proveedor de Lima", "Productos agotados por más tiempo", "Stock de seguridad en productos clase A; proveedor alterno en Piura."],
           ["Mercadería dañada o incompleta", "Pérdida económica; faltantes", "Verificación en la recepción y registro de la incidencia para el reclamo."],
           ["Desabastecimiento del importador", "Falta de productos de origen extranjero", "Identificar productos sustitutos de origen nacional."],
           ["Error en el ingreso al sistema", "Stock incorrecto", "Ingresar con el comprobante en mano y revisar el stock de los productos recibidos."]],
          widths=[4.2, 4.6, 7.0], note="Elaboración propia.")

# ---------------- VII. MARKETING ----------------
H(1, "VII. EXPANSIÓN GLOBAL: MARKETING Y NEGOCIACIÓN INTERNACIONAL")
P("INVERSIONES JL no tiene clientes en el exterior. Por ello, se analizan las estrategias de marketing y negociación que efectivamente se aplican con los clientes locales, y se identifican mejoras.")
H(2, "7.1. Estrategias de Marketing Internacional")
H(3, "7.1.1. Diseño e implementación de estrategias para el fortalecimiento de relaciones cliente-organización")
P("El marketing busca crear valor para el cliente y construir relaciones duraderas basadas en su satisfacción (Kotler et al., 2022). Actualmente, la empresa aplica tácticas de atención personalizada: se pregunta al cliente para entender su necesidad, se ofrecen alternativas cuando un producto no está disponible y se otorgan descuentos razonables. Como propuestas se plantean:")
B("Registrar a los clientes por oficio (maestros de obra, gasfiteros, electricistas) y avisarles por WhatsApp de la llegada de los productos que buscaban.")
B("Registrar la demanda no atendida para incorporar al surtido los productos solicitados con frecuencia.")
B("Asegurar la disponibilidad de los productos más vendidos, ya que un faltante puede llevar al cliente a cambiar de tienda (Azeem et al., 2019).")
H(3, "7.1.2. Análisis del impacto del liderazgo y el trabajo en equipo en el despliegue de la estrategia comercial")
P("El liderazgo en la empresa lo ejerce la propietaria, quien delega en mí funciones de confianza como el manejo de la caja desde la segunda semana de prácticas. Esta delegación, junto con la comunicación diaria por WhatsApp y la entrega diaria del balance, permitió que me integrara rápidamente al equipo. Cuando ocurrió un faltante de caja, lo comuniqué de inmediato y fue comprendido, lo que refleja un ambiente de confianza. En las pymes comerciales, los cambios tienden a ser poco planificados cuando sus líderes no cuentan con suficiente preparación (Caller Luna et al., 2022), y la gestión empresarial influye en su desarrollo (Luciano Alipio et al., 2023). Por ello, las propuestas de este informe se plantean de forma sencilla y gradual, para que puedan ser asumidas por el equipo.")
H(2, "7.2. Ejecución de Planes de Negociación Global")
H(3, "7.2.1. Diseño y ejecución de tácticas para concretar transacciones económicas internacionales")
P("La empresa no ha participado en negociaciones internacionales; la negociación con los proveedores la realizan los jefes. En mi función, la negociación ocurre con los clientes al momento de la venta: cuando un cliente solicita un descuento, verifico en Rapifac el precio de compra del producto y ofrezco un descuento que mantenga un margen para el negocio, generalmente cuando la compra es de mayor volumen. Esta táctica permite cerrar la venta sin perder rentabilidad.")
H(3, "7.2.2. Aplicación de la comunicación efectiva y criterios de equidad en el cierre de acuerdos comerciales")
P("La comunicación efectiva se aplica informando con claridad el precio de cada producto y el total antes del pago, y diciendo con honestidad cuando un producto no está disponible. La equidad se asegura porque los descuentos se registran en el sistema y se aplican con el mismo criterio a todos los clientes, buscando un beneficio mutuo: el cliente obtiene un mejor precio y la empresa concreta la venta con margen positivo.")

# ---------------- VIII. PROPUESTA ESTRATÉGICA ----------------
H(1, "VIII. PROPUESTA ESTRATÉGICA")
H(2, "8.1. Descripción del Problema Crítico identificado")
P("**Problema:** reposición reactiva de mercadería por falta de un stock mínimo y por un stock desactualizado en el sistema, lo que genera quiebres de stock en productos de alta rotación.")
apa_table("Cuantificación del problema crítico",
          ["Aspecto", "Situación actual", "Fuente"],
          [["Frecuencia de productos no disponibles", "≈ 1 caso por día (≈ 6 por semana; ≈ 26 por mes de 26 días hábiles)", "Observación del practicante"],
           ["Productos más afectados", "Cemento y conexiones de tubería", "Observación del practicante"],
           ["Tiempo sin producto", "De 4 días hábiles (si el faltante se detecta antes del pedido) a 10 días hábiles (si se detecta justo después del pedido semanal)", "Pedido semanal + 4 días de entrega"],
           ["Clientes perdidos", "Parte de los clientes no acepta la alternativa y acude a uno de los 2 competidores cercanos", "Observación del practicante"],
           ["Exactitud del stock en el sistema", "Desactualizado", "Observación del practicante"]],
          widths=[4.2, 7.6, 4.0], note="Elaboración propia. El valor monetario de las ventas perdidas no se estima por falta de datos de precios y márgenes autorizados; se propone medirlo con el registro de demanda no atendida.")
P("**Causas identificadas:** (a) no se ha definido un stock mínimo por producto; (b) el stock de Rapifac no se verifica ni se ajusta, por lo que no sirve para decidir; (c) la revisión de existencias es visual y ocasional; y (d) la comunicación de faltantes es informal. **Consecuencias:** ventas perdidas, clientes que se trasladan a la competencia (Fitzsimons, 2000) y compras de urgencia. Estudios en negocios minoristas que gestionaban su inventario de forma manual y por experiencia del personal muestran que la clasificación ABC permite concentrar el control en pocos productos que concentran la mayor parte del valor de las ventas (Pradita y Widodo, 2024).")
H(2, "8.2. Diseño de la Solución y Plan de Implementación")
P("**Objetivo de la propuesta:** reducir los quiebres de stock de los productos de mayor rotación mediante un sistema simple de control de existencias y reposición. La propuesta se basa en la clasificación ABC y en el modelo de revisión periódica, adecuado para negocios que realizan pedidos en intervalos fijos (Heizer et al., 2023; Jacobs y Chase, 2024).")
P("**Componente 1 – Clasificación ABC.** Con el reporte de ventas de Rapifac de los últimos tres meses, ordenar los productos según su valor de ventas: la clase A agrupa los pocos productos que concentran la mayor parte de las ventas y recibe control semanal; las clases B y C, control mensual (Heizer et al., 2023).")
P("**Componente 2 – Stock mínimo y nivel máximo para la clase A.** Como el pedido es semanal, en la revisión se compara el stock con un nivel mínimo y se pide hasta un nivel máximo:")
B("Stock mínimo = demanda diaria × tiempo de entrega + stock de seguridad.")
B("Nivel máximo = demanda diaria × (días entre pedidos + tiempo de entrega) + stock de seguridad (Jacobs y Chase, 2024).")
P("A modo ilustrativo, con la estimación de venta de [[unas 20 bolsas de cemento por semana]] (≈ 3,3 bolsas por día hábil), un tiempo de entrega de 4 días y un stock de seguridad equivalente a 2 días de venta, el cálculo es el siguiente:")
apa_table("Cálculo ilustrativo de niveles de inventario para el cemento",
          ["Parámetro", "Valor", "Cálculo"],
          [["Demanda diaria (d)", "3,3 bolsas", "20 bolsas ÷ 6 días"],
           ["Tiempo de entrega (L)", "4 días", "Dato de la empresa"],
           ["Días entre pedidos (R)", "6 días", "Pedido semanal"],
           ["Stock de seguridad (SS)", "≈ 7 bolsas", "3,3 × 2 días"],
           ["Stock mínimo", "≈ 20 bolsas", "3,3 × 4 + 7"],
           ["Nivel máximo", "≈ 40 bolsas", "3,3 × (6 + 4) + 7"],
           ["Cantidad a pedir", "Nivel máximo − stock actual", "Ej.: si hay 15 bolsas, pedir ≈ 25"]],
          widths=[5.0, 4.4, 6.4], note="Elaboración propia. Valores aproximados para ilustrar el método; deben recalcularse con las ventas reales registradas en Rapifac.")
P("**Componente 3 – Stock confiable.** Realizar un conteo físico inicial de los productos clase A, ajustar el stock en Rapifac y verificarlo cada semana; los productos B y C se cuentan por grupos cada mes, en las horas de menor afluencia (Heizer et al., 2023).")
P("**Componente 4 – Lista de reposición y registro de demanda no atendida.** Usar el formato del Anexo 3 para el pedido semanal y anotar cada producto solicitado que no se tenga, lo que permitirá medir el problema y su reducción.")
apa_table("Plan de implementación de la propuesta",
          ["Semana", "Actividad", "Responsable", "Recursos"],
          [["1", "Presentar la propuesta a la propietaria y al jefe; obtener la aprobación.", "Practicante", "Informe y formatos"],
           ["1–2", "Exportar las ventas de Rapifac y realizar la clasificación ABC.", "Practicante", "Rapifac, Excel"],
           ["2", "Conteo físico de los productos clase A y ajuste del stock en Rapifac.", "Practicante y propietaria", "Rapifac, formato de conteo"],
           ["3", "Calcular el stock mínimo y el nivel máximo de los productos clase A.", "Practicante", "Excel"],
           ["3 en adelante", "Revisión semanal previa al pedido y lista estandarizada de reposición.", "Practicante; deciden los jefes", "Rapifac, WhatsApp, Anexo 3"],
           ["3 en adelante", "Registro de demanda no atendida y arqueo de caja con diferencias.", "Practicante", "Anexos 3 y 4"],
           ["8", "Evaluar los indicadores (casos de producto agotado y exactitud del stock) y ajustar los niveles.", "Practicante y propietaria", "Registros"]],
          widths=[2.2, 7.6, 3.4, 2.6], note="Elaboración propia. Propuesta no implementada a la fecha del informe; no requiere inversión en software adicional.")
P("**Resultados esperados:** disminuir los casos semanales de producto agotado en la clase A respecto de la línea base (≈ 6 por semana), lograr que el stock del sistema coincida con el físico en los productos clave y contar con datos propios para estimar el valor de las ventas recuperadas. Asimismo, el registro de las diferencias de caja permitirá identificar su causa y fortalecer el control interno del negocio (COSO, 2013; León León et al., 2023).")

# ---------------- IX. CONCLUSIONES ----------------
H(1, "IX. CONCLUSIONES")
B("INVERSIONES JL es una ferretería minorista formal (persona natural con negocio, RUC activo y habido) que opera desde 2017 con una estructura simple de tres integrantes y que no realiza operaciones de comercio exterior; por ello, los capítulos referidos a logística y negociación internacional se adaptaron a su cadena de suministro y a su relación con los clientes locales.")
B("El problema crítico es la reposición reactiva: sin stock mínimo y con el stock del sistema desactualizado, se produce en promedio un caso diario de producto no disponible, que puede durar entre 4 y 10 días hábiles y que genera ventas perdidas frente a dos competidores cercanos.")
B("La propuesta de control de existencias (clasificación ABC, stock mínimo y nivel máximo, revisión programada y lista de reposición) es viable porque utiliza recursos existentes y no requiere inversión en software adicional.")
B("El control de caja se realiza diariamente, pero las diferencias no se registran ni se analizan, lo que impide identificar su causa.")
B("**Lecciones aprendidas:** aprendí que la atención al cliente empieza por entender su necesidad; que la información de un sistema solo es útil si se mantiene actualizada; y que, en una empresa pequeña, la confianza y la comunicación diaria son la base del trabajo en equipo. Estas prácticas me permitieron aplicar conocimientos de gestión comercial, operaciones y control interno de mi formación en Administración y Negocios Internacionales.")

# ---------------- X. RECOMENDACIONES ----------------
H(1, "X. RECOMENDACIONES")
B("Implementar de forma gradual el sistema de control de existencias, empezando por los productos de mayor rotación (cemento, conexiones, tubos y material eléctrico).")
B("Mantener actualizado el stock de Rapifac mediante conteos semanales de la clase A y mensuales de las clases B y C.")
B("Usar el formato de arqueo de caja para registrar las diferencias diarias y revisar sus causas al final de cada mes.")
B("Documentar en un procedimiento escrito sencillo la atención al cliente, el cierre de caja y la recepción de mercadería, para facilitar la capacitación de futuros colaboradores.")
B("Evaluar, una vez que el control de stock sea confiable, servicios complementarios como la atención de pedidos por WhatsApp para clientes frecuentes.")

# ---------------- XI. REFERENCIAS ----------------
H(1, "XI. REFERENCIAS")
REFS = [
 "Aquiño, L. (2024). Influencia del control interno en la gestión administrativa de un fondo editorial de Lima 2023. *Revista Latinoamericana de Ciencias Sociales – Relacis, 2*(2), 11–28.",
 "Azeem, M. M., Baker, D., Villano, R. A., Mounter, S. y Griffith, G. (2019). Response to stockout in grocery stores: A small city case in a changing competitive environment. *Journal of Retailing and Consumer Services, 49*, 242–252. https://doi.org/10.1016/j.jretconser.2019.04.001",
 "Barrantes, L. y Alzamora, M. (2023). La clave está en la inclusión financiera. *Revista Moneda*, (194), 25–31.",
 "Caller Luna, J., Lizarbe Castro, M. y González Ponce de León, F. (2022). Cambio organizacional en pymes de Lima. *Revista Venezolana de Gerencia, 27*(Especial 8), 914–925. https://doi.org/10.52080/rvgluz.27.8.12",
 "Chopra, S. (2025). *Supply chain management: Strategy, planning, and operation* (8.ª ed.). Pearson.",
 "Committee of Sponsoring Organizations of the Treadway Commission. (2013). *Internal control – Integrated framework: Executive summary*. COSO.",
 "David, F. R., David, F. R. y David, M. E. (2023). *Strategic management: A competitive advantage approach, concepts and cases* (18.ª ed.). Pearson.",
 "Decreto Supremo N.° 013-2013-PRODUCE. Texto Único Ordenado de la Ley de Impulso al Desarrollo Productivo y al Crecimiento Empresarial. (2013, 28 de diciembre). Diario Oficial El Peruano.",
 "Fitzsimons, G. J. (2000). Consumer response to stockouts. *Journal of Consumer Research, 27*(2), 249–266. https://doi.org/10.1086/314323",
 "Heizer, J., Render, B. y Munson, C. (2023). *Operations management: Sustainability and supply chain management* (14.ª ed.). Pearson.",
 "Jacobs, F. R. y Chase, R. B. (2024). *Operations and supply chain management* (17.ª ed.). McGraw Hill.",
 "Kotler, P., Keller, K. L. y Chernev, A. (2022). *Marketing management* (16.ª ed.). Pearson.",
 "León León, C. F., Maguiña Rivero, O. F. y León Cavero, F. R. (2023). Control interno para la mejora de la rentabilidad empresarial en Nuevo Chimbote-Perú. *Revista de Ciencias Sociales, 29*(Especial 7), 66–76. https://doi.org/10.31876/rcs.v29i.40447",
 "Luciano Alipio, R. A., Sotomayor Chahuaylla, J. A., García Juárez, H. D. y Peláez Camacho, H. Y. (2023). Gestión empresarial en el desarrollo de las MYPES en zonas mineras del Perú. *Revista Venezolana de Gerencia, 28*(103), 1174–1189. https://doi.org/10.52080/rvgluz.28.103.16",
 "Ministerio de la Producción. (2025). *Las MIPYME en cifras 2024*. https://www.producempresarial.pe/las-mipyme-en-cifras-2024/",
 "Mubiru, K. (2024). A stochastic goal-based economic order quantity model for multi-item inventory. *International Journal of Research in Industrial Engineering, 13*(4), 399–413. https://doi.org/10.22105/riej.2024.430542.1432",
 "Pradita, S. P. y Widodo, K. H. (2024). Product inventory control at a local brand forum outlet using ABC analysis. *Agrointek, 18*(2), 246–255. https://doi.org/10.21107/agrointek.v18i2.18819",
 "Robbins, S. P., Coulter, M. y Long, L. (2025). *Management* (16.ª ed.). Pearson.",
 "Superintendencia Nacional de Aduanas y de Administración Tributaria. (2026). *Consulta RUC: 10719941023* [Base de datos]. Recuperado el 17 de noviembre de 2026 de https://e-consultaruc.sunat.gob.pe/",
 "Yaipen Ayasta, Y. P. y Gonzales Rentería, Y. G. (2024). Caracterización del control interno de las micro y pequeñas empresas del sector comercio del Perú: Caso ferretería y multiservicios San Rafael S.A.C. – Casma, 2023. *Divulgación Científica de Investigación y Emprendimiento*, (2). https://doi.org/10.54798/BYUI7468",
]
for ref in REFS:
    p = doc.add_paragraph(); add_runs(p, ref)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.left_indent = Cm(1.27); p.paragraph_format.first_line_indent = Cm(-1.27)

# =====================================================================
# ANEXOS (sin numeración de página)
# =====================================================================
sec3 = new_section(); page_number_footer(sec3, None, None)
H(1, "XII. ANEXOS")
P("**Anexo 1.** Consulta RUC de SUNAT (17/11/2026)", indent=False)
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.add_run().add_picture(SUNAT, width=Cm(14.5))
P("*Nota.* Captura de la consulta RUC proporcionada por el practicante.", indent=False, size=10)
page_break()
P("**Anexo 2.** Evidencias fotográficas del practicante en el centro de prácticas", indent=False)
P("[[Insertar entre 4 y 6 fotografías: fachada del local con el letrero INVERSIONES JL; atención al cliente; registro de ventas en Rapifac (sin datos sensibles); revisión de estantes; recepción de mercadería. Cada foto con una breve descripción.]]", indent=False)
P("**Anexo 3.** Formato propuesto: lista semanal de reposición y registro de demanda no atendida", indent=False)
apa_table("Formato de lista semanal de reposición",
          ["Producto", "Clase", "Stock sistema", "Stock físico", "Stock mínimo", "Nivel máximo", "Cantidad sugerida"],
          [["", "", "", "", "", "", ""] for _ in range(4)], size=9,
          note="Elaboración propia. Se envía a los jefes un día antes del pedido semanal.")
apa_table("Formato de registro de demanda no atendida",
          ["Fecha", "Producto solicitado", "Cantidad", "¿Aceptó alternativa? (Sí/No)", "Observación"],
          [["", "", "", "", ""] for _ in range(4)], size=9, note="Elaboración propia.")
P("**Anexo 4.** Formato propuesto de arqueo diario de caja", indent=False)
apa_table("Formato de arqueo diario de caja",
          ["Concepto", "Monto (S/)"],
          [["(A) Efectivo contado al cierre", ""], ["(B) Fondo inicial", "100,00"], ["(C) Efectivo neto de ventas = A − B", ""],
           ["(D) Cobros por Yape del día (según aplicativo)", ""], ["(E) Total recaudado = C + D", ""],
           ["(F) Ventas del día según reporte de Rapifac", ""], ["(G) Diferencia = E − F (sobrante + / faltante −)", ""],
           ["Explicación de la diferencia y firma de quien entrega y de quien recibe", ""]],
          widths=[11.0, 4.8], size=10, note="Elaboración propia con base en el procedimiento actual de cierre de caja.")
P("**Anexo 5.** Instrumento de recojo de información", indent=False)
P("Guía de entrevista estructurada aplicada en cinco bloques (identificación legal, antecedentes, productos, clientes, proveedores, ventas, caja, inventarios, reposición y funciones del practicante), complementada con la observación directa durante las prácticas.", indent=False)
P("**Anexo 6.** Autorización de la empresa para el uso de su información", indent=False)
P("[[Adjuntar la carta de autorización firmada por la propietaria.]]", indent=False)

# =====================================================================
# Completar ÍNDICE con páginas (2.ª pasada)
# =====================================================================
def fill_toc():
    body = TOC_ANCHOR._p
    parent = body.getparent()
    idx = parent.index(body)
    new_ps = []
    def mk(text, page, level, bold=False):
        p = toc_entry(text, page, level, bold)
        new_ps.append(p._p)
    for level, text in TOC:
        pg = PAGES.get(text, "")
        mk(text, pg, level, bold=(level == 1))
    # índice de tablas y figuras
    mk("", "", 1)
    mk("ÍNDICE DE TABLAS", "", 1, bold=True)
    for t in TABLES:
        mk(t, PAGES.get(t, ""), 2)
    mk("", "", 1)
    mk("ÍNDICE DE FIGURAS", "", 1, bold=True)
    for f in FIGS:
        mk(f, PAGES.get(f, ""), 2)
    for k, np_ in enumerate(new_ps):
        parent.remove(np_)
        parent.insert(idx + k, np_)
    parent.remove(body); parent.remove(TOC_ANCHOR_T._p)
fill_toc()
doc.save(OUT)
json.dump({"toc": TOC, "tables": TABLES, "figs": FIGS}, open(OUT + ".meta.json", "w"), ensure_ascii=False, indent=1)
print("OK", OUT, len(TOC), "títulos,", len(TABLES), "tablas")
