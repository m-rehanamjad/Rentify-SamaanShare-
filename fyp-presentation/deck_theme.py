"""Shared layout primitives for the SamaanShare FYDP-I deck.

Colours and type follow the product's own design system (src/app/globals.css):
deep navy #142B3B ink, emerald #047857 brand, mint #E6F4EE surface, #F8FAFC page.
"""

from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

# ---------------------------------------------------------------- palette
NAVY = RGBColor(0x14, 0x2B, 0x3B)
NAVY_DEEP = RGBColor(0x0C, 0x1C, 0x27)
NAVY_SOFT = RGBColor(0x1E, 0x3D, 0x50)
EMERALD = RGBColor(0x04, 0x78, 0x57)
EMERALD_DK = RGBColor(0x06, 0x5F, 0x46)
EMERALD_BR = RGBColor(0x05, 0x96, 0x69)
MINT = RGBColor(0xE6, 0xF4, 0xEE)
INK = RGBColor(0x14, 0x2B, 0x3B)
MUTED = RGBColor(0x53, 0x62, 0x73)
LINE = RGBColor(0xD5, 0xDF, 0xE7)
PAPER = RGBColor(0xF8, 0xFA, 0xFC)
SLATE = RGBColor(0xF1, 0xF5, 0xF9)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
AMBER = RGBColor(0xB4, 0x5A, 0x09)
AMBER_BR = RGBColor(0xD9, 0x77, 0x06)
AMBER_SOFT = RGBColor(0xFD, 0xF3, 0xE3)
RED = RGBColor(0xB9, 0x1C, 0x1C)
RED_SOFT = RGBColor(0xFD, 0xEC, 0xEC)

HEAD = "Segoe UI"
BODY = "Segoe UI"

# ---------------------------------------------------------------- geometry
SW, SH = 13.3333, 7.5
M = 0.78                 # left content margin
CW = SW - 2 * M          # content width  (11.773)
TOP = 1.62               # first content row
BOTTOM = 6.82            # last content row


def hexw(color):
    return str(color)


# ---------------------------------------------------------------- low level
def _style_run(run, size, color, bold, italic, font, spacing=None, caps=False):
    f = run.font
    f.size = Pt(size)
    f.bold = bold
    f.italic = italic
    f.name = font
    f.color.rgb = color
    rPr = run._r.get_or_add_rPr()
    if spacing is not None:
        rPr.set("spc", str(int(round(spacing * 100))))
    if caps:
        rPr.set("cap", "all")
    # make sure the same face is used for east-asian / complex script ranges
    for tag in ("a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {})
            rPr.append(el)
        el.set("typeface", font)


def _rich(paragraph, text, size, color, bold, italic, font, spacing, caps):
    """Write text into a paragraph, honouring **bold** and [[muted]] segments."""
    if isinstance(text, (list, tuple)):
        for piece in text:
            if isinstance(piece, str):
                _rich(paragraph, piece, size, color, bold, italic, font, spacing, caps)
            else:
                t, over = piece
                _rich(
                    paragraph, t,
                    over.get("size", size),
                    over.get("color", color),
                    over.get("bold", bold),
                    over.get("italic", italic),
                    over.get("font", font),
                    over.get("spacing", spacing),
                    over.get("caps", caps),
                )
        return

    parts = str(text).split("**")
    for i, part in enumerate(parts):
        if part == "":
            continue
        run = paragraph.add_run()
        run.text = part
        _style_run(run, size, color, bold or (i % 2 == 1), italic, font, spacing, caps)


def textbox(slide, x, y, w, h, text="", size=13, color=INK, bold=False, italic=False,
            align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, font=BODY, spacing=None,
            line=None, space_after=0, margins=0.0, caps=False, wrap=True):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = wrap
    tf.margin_left = tf.margin_right = Inches(margins)
    tf.margin_top = tf.margin_bottom = Inches(margins * 0.6)
    tf.vertical_anchor = anchor
    tf.auto_size = None
    p = tf.paragraphs[0]
    p.alignment = align
    if line:
        p.line_spacing = line
    if space_after:
        p.space_after = Pt(space_after)
    if text != "":
        _rich(p, text, size, color, bold, italic, font, spacing, caps)
    return box


def paragraph(shape_or_tf, text, size=13, color=INK, bold=False, italic=False,
              align=PP_ALIGN.LEFT, space_before=0, space_after=0, line=None,
              font=BODY, spacing=None, caps=False, first=False):
    tf = shape_or_tf.text_frame if hasattr(shape_or_tf, "text_frame") else shape_or_tf
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.alignment = align
    if space_before:
        p.space_before = Pt(space_before)
    if space_after:
        p.space_after = Pt(space_after)
    if line:
        p.line_spacing = line
    _rich(p, text, size, color, bold, italic, font, spacing, caps)
    return p


def bullets(slide, x, y, w, h, items, size=13, color=INK, gap=7, line=1.18,
            bullet_color=EMERALD, marker="\u2013", indent=0.24, anchor=MSO_ANCHOR.TOP):
    """items: list of str | (str, level) | dict(t=..., sub=True, marker=..., color=...)"""
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    tf.auto_size = None
    for i, item in enumerate(items):
        sub = False
        mk = marker
        col = color
        sz = size
        bold_all = False
        if isinstance(item, dict):
            text = item["t"]
            sub = item.get("sub", False)
            mk = item.get("marker", "·" if sub else marker)
            col = item.get("color", color)
            sz = item.get("size", size)
            bold_all = item.get("bold", False)
        elif isinstance(item, tuple):
            text, sub = item[0], item[1]
            mk = "·" if sub else marker
        else:
            text = item
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(gap)
        p.line_spacing = line
        lvl = 1 if sub else 0
        marL = indent * (1 + lvl)
        pPr = p._p.get_or_add_pPr()
        pPr.set("marL", str(int(marL * 914400)))
        pPr.set("indent", str(-int((indent if not sub else 0.18) * 914400)))
        pPr.append(pPr.makeelement(qn("a:buNone"), {}))
        # glyph
        r = p.add_run()
        r.text = mk + "\u2002"
        _style_run(r, sz, bullet_color if not sub else MUTED, True, False, BODY)
        _rich(p, text, sz, col, bold_all, False, BODY, None, False)
    return box


# ---------------------------------------------------------------- shapes
def bg(slide, color=WHITE):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(SW), Inches(SH))
    shp.fill.solid()
    shp.fill.fore_color.rgb = color
    shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def rect(slide, x, y, w, h, fill=WHITE, line=None, lw=1.0, radius=None,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE, dash=None):
    shp = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if radius is not None and shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        shp.adjustments[0] = radius
    if fill is None:
        shp.fill.background()
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = Pt(lw)
        if dash:
            from pptx.enum.dml import MSO_LINE_DASH_STYLE
            shp.line.dash_style = dash
    shp.shadow.inherit = False
    shp.text_frame.word_wrap = True
    return shp


def card(slide, x, y, w, h, fill=WHITE, line=LINE, lw=1.0, radius=0.045):
    return rect(slide, x, y, w, h, fill=fill, line=line, lw=lw, radius=radius)


def pill(slide, x, y, w, h, text, fill=MINT, color=EMERALD_DK, size=10, bold=True,
         line=None, lw=0.75, spacing=0.4):
    shp = rect(slide, x, y, w, h, fill=fill, line=line, lw=lw, radius=0.5)
    tf = shp.text_frame
    tf.margin_left = tf.margin_right = Inches(0.04)
    tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    _rich(p, text, size, color, bold, False, BODY, spacing, False)
    return shp


def hline(slide, x, y, w, color=LINE, lw=1.0):
    shp = slide.shapes.add_connector(1, Inches(x), Inches(y), Inches(x + w), Inches(y))
    shp.line.color.rgb = color
    shp.line.width = Pt(lw)
    return shp


def arrow_right(slide, x, y, w, h, color=LINE):
    shp = rect(slide, x, y, w, h, fill=color, line=None, shape=MSO_SHAPE.RIGHT_ARROW)
    return shp


def arrow_down(slide, x, y, w, h, color=LINE):
    return rect(slide, x, y, w, h, fill=color, line=None, shape=MSO_SHAPE.DOWN_ARROW)


def circle_num(slide, x, y, d, label, fill=EMERALD, color=WHITE, size=12, bold=True):
    shp = rect(slide, x, y, d, d, fill=fill, line=None, shape=MSO_SHAPE.OVAL)
    tf = shp.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    _rich(p, label, size, color, bold, False, BODY, None, False)
    return shp


# ---------------------------------------------------------------- tables
def _cell_border(cell, color=LINE, w=0.75, edges=("L", "R", "T", "B")):
    tcPr = cell._tc.get_or_add_tcPr()
    for e in reversed([e for e in ("L", "R", "T", "B") if e in edges]):
        tag = "a:ln" + e
        for el in tcPr.findall(qn(tag)):
            tcPr.remove(el)
        ln = tcPr.makeelement(qn(tag), {
            "w": str(int(w * 12700)), "cap": "flat", "cmpd": "sng", "algn": "ctr"})
        fill = ln.makeelement(qn("a:solidFill"), {})
        clr = fill.makeelement(qn("a:srgbClr"), {"val": hexw(color)})
        fill.append(clr)
        ln.append(fill)
        ln.append(ln.makeelement(qn("a:prstDash"), {"val": "solid"}))
        tcPr.insert(0, ln)


def table(slide, x, y, w, rows, widths, row_h=0.36, header_h=None, size=10.5,
          hsize=10, header_fill=NAVY, header_color=WHITE, zebra=None,
          border=LINE, aligns=None, bold_first=False, first_color=INK,
          valign=MSO_ANCHOR.MIDDLE, pad=0.09, header_align=None, bold_rows=(),
          colors=None):
    """rows[0] is the header. widths are inches and should sum to w."""
    nrows, ncols = len(rows), len(rows[0])
    header_h = header_h or row_h
    total_h = header_h + row_h * (nrows - 1)
    gf = slide.shapes.add_table(nrows, ncols, Inches(x), Inches(y), Inches(w), Inches(total_h))
    tbl = gf.table
    tbl.first_row = False
    tbl.horz_banding = False
    for i, cw in enumerate(widths):
        tbl.columns[i].width = Inches(cw)
    tbl.rows[0].height = Inches(header_h)
    for r in range(1, nrows):
        tbl.rows[r].height = Inches(row_h)

    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            cell = tbl.cell(r, c)
            cell.margin_left = Inches(pad)
            cell.margin_right = Inches(pad * 0.7)
            cell.margin_top = Inches(0.035)
            cell.margin_bottom = Inches(0.035)
            cell.vertical_anchor = valign
            is_head = (r == 0)
            fill = header_fill if is_head else (
                zebra if (zebra and r % 2 == 0) else WHITE)
            cell.fill.solid()
            cell.fill.fore_color.rgb = fill
            if border is not None:
                _cell_border(cell, color=border, w=0.75,
                             edges=("T", "B") if not is_head else ("T", "B"))
            tf = cell.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            if is_head:
                al = (header_align[c] if header_align else
                      (aligns[c] if aligns else PP_ALIGN.LEFT))
            else:
                al = aligns[c] if aligns else PP_ALIGN.LEFT
            p.alignment = al
            p.line_spacing = 1.02
            col = header_color if is_head else first_color
            bold = is_head or (bold_first and c == 0) or (r in bold_rows)
            if colors and (r, c) in colors:
                col = colors[(r, c)]
            _rich(p, val, hsize if is_head else size, col, bold, False, BODY,
                  None, is_head)
    return gf


# ---------------------------------------------------------------- page furniture
def footer(slide, page, note="SamaanShare  \u00b7  Group 23FYDP24  \u00b7  FYDP-I Final Presentation"):
    hline(slide, M, 6.98, CW, color=LINE, lw=0.75)
    textbox(slide, M, 7.05, 9.0, 0.28, note, size=8.5, color=MUTED, spacing=0.3)
    textbox(slide, SW - M - 1.0, 7.05, 1.0, 0.28, f"{page:02d}", size=8.5,
            color=MUTED, align=PP_ALIGN.RIGHT, bold=True)


_PAGE = [0]


def pg():
    _PAGE[0] += 1
    return _PAGE[0]


def content_slide(prs, kicker, title, page=None, title_size=25):
    page = pg() if page is None else page
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s, WHITE)
    textbox(s, M, 0.40, 9.5, 0.24, kicker.upper(), size=9, color=EMERALD,
            bold=True, spacing=1.7)
    # keep every title on a single line: shrink to fit the content width
    fit = 800.0 / (0.505 * max(1, len(title)))
    textbox(s, M, 0.63, CW, 0.62, title, size=min(title_size, fit), color=NAVY,
            bold=True, line=1.0)
    rect(s, M, 1.32, 0.62, 0.045, fill=EMERALD, line=None, shape=MSO_SHAPE.RECTANGLE)
    hline(s, M + 0.72, 1.343, CW - 0.72, color=LINE, lw=0.75)
    footer(s, page)
    return s


def lead(slide, x, y, w, text, size=12.5, color=MUTED, h=0.5):
    return textbox(slide, x, y, w, h, text, size=size, color=color, line=1.22)


def caption(slide, x, y, w, text, size=9.5, color=MUTED, bold=False):
    return textbox(slide, x, y, w, 0.3, text, size=size, color=color, bold=bold)


def set_alpha(shape, pct):
    """Apply transparency (0 = opaque, 100 = invisible) to a solid fill."""
    from pptx.oxml.ns import qn as _qn
    srgb = shape.fill.fore_color._xFill.find(_qn("a:srgbClr"))
    if srgb is None:
        return
    alpha = srgb.makeelement(_qn("a:alpha"), {"val": str(int((100 - pct) * 1000))})
    srgb.append(alpha)


def picture(slide, path, x, y, w=None, h=None):
    kw = {}
    if w:
        kw["width"] = Inches(w)
    if h:
        kw["height"] = Inches(h)
    return slide.shapes.add_picture(path, Inches(x), Inches(y), **kw)
