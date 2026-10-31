"""Formatting helpers for the FYDP-1 methodology baseline document.

Kept separate from the content so the builder file stays readable.
Everything here is presentation only: page setup, styles, tables, callouts.
"""

from __future__ import annotations

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

NAVY = RGBColor(0x1F, 0x38, 0x64)
INK = RGBColor(0x25, 0x2A, 0x33)
GREY = RGBColor(0x5B, 0x66, 0x74)
LIGHT_RULE = "C9D2D9"
NAVY_HEX = "1F3864"
BAND_HEX = "F4F6F8"
AMBER_FILL = "FCF6E8"
AMBER_EDGE = "C79A3B"
BLUE_FILL = "EDF2F8"
BLUE_EDGE = "7C9CBF"

P_SUCCESSORS = (
    "w:shd", "w:tabs", "w:suppressAutoHyphens", "w:kinsoku", "w:wordWrap",
    "w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN",
    "w:bidi", "w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind",
    "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap", "w:jc",
    "w:textDirection", "w:textAlignment", "w:textboxTightWrap", "w:outlineLvl",
    "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr", "w:pPrChange",
)

BODY_FONT = "Georgia"
HEAD_FONT = "Calibri"
TABLE_FONT = "Calibri"

TEXT_WIDTH_CM = 15.5


# --------------------------------------------------------------------------
# low level xml helpers


def _el(tag, **attrs):
    element = OxmlElement(tag)
    for key, value in attrs.items():
        element.set(qn(f"w:{key}"), value)
    return element


TC_SHD_SUCCESSORS = (
    "w:noWrap", "w:tcMar", "w:textDirection", "w:tcFitText", "w:vAlign",
    "w:hideMark", "w:headers", "w:cellIns", "w:cellDel", "w:cellMerge",
    "w:tcPrChange",
)


def shade(cell, hex_fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    for old in tc_pr.findall(qn("w:shd")):
        tc_pr.remove(old)
    tc_pr.insert_element_before(
        _el("w:shd", val="clear", color="auto", fill=hex_fill), *TC_SHD_SUCCESSORS
    )


def cell_margins(cell, top=60, bottom=60, left=100, right=100) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    for old in tc_pr.findall(qn("w:tcMar")):
        tc_pr.remove(old)
    mar = _el("w:tcMar")
    for edge, value in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        mar.append(_el(f"w:{edge}", w=str(value), type="dxa"))
    tc_pr.insert_element_before(
        mar, "w:textDirection", "w:tcFitText", "w:vAlign", "w:hideMark", "w:headers",
        "w:cellIns", "w:cellDel", "w:cellMerge", "w:tcPrChange",
    )


TBL_SUCCESSORS = (
    "w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook", "w:tblCaption",
    "w:tblDescription", "w:tblPrChange",
)


def _put_borders(tbl_pr, borders) -> None:
    for old in tbl_pr.findall(qn("w:tblBorders")):
        tbl_pr.remove(old)
    tbl_pr.insert_element_before(borders, *TBL_SUCCESSORS)


def table_borders(table, color: str = LIGHT_RULE, size: int = 6) -> None:
    borders = _el("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        borders.append(
            _el(f"w:{edge}", val="single", sz=str(size), space="0", color=color)
        )
    _put_borders(table._tbl.tblPr, borders)


def no_borders(table) -> None:
    borders = _el("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        borders.append(_el(f"w:{edge}", val="none", sz="0", space="0", color="auto"))
    _put_borders(table._tbl.tblPr, borders)


def left_accent(table, color: str, size: int = 18) -> None:
    borders = _el("w:tblBorders")
    for edge in ("top", "bottom", "right"):
        borders.append(_el(f"w:{edge}", val="none", sz="0", space="0", color="auto"))
    borders.append(_el("w:left", val="single", sz=str(size), space="0", color=color))
    # re-order into the schema sequence: top, left, bottom, right
    ordered = _el("w:tblBorders")
    for edge in ("top", "left", "bottom", "right"):
        ordered.append(borders.find(qn(f"w:{edge}")))
    _put_borders(table._tbl.tblPr, ordered)


def keep_row_together(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    tr_pr.append(_el("w:cantSplit"))


def repeat_header(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    tr_pr.append(_el("w:tblHeader"))


def set_col_widths(table, widths_cm) -> None:
    table.autofit = False
    for row in table.rows:
        for idx, width in enumerate(widths_cm):
            row.cells[idx].width = Cm(width)
    for idx, width in enumerate(widths_cm):
        table.columns[idx].width = Cm(width)


def add_field(paragraph, instr: str, placeholder: str = "1") -> None:
    run = paragraph.add_run()
    begin = _el("w:fldChar", fldCharType="begin")
    run._r.append(begin)
    run2 = paragraph.add_run()
    instr_el = OxmlElement("w:instrText")
    instr_el.set(qn("xml:space"), "preserve")
    instr_el.text = instr
    run2._r.append(instr_el)
    run3 = paragraph.add_run()
    run3._r.append(_el("w:fldChar", fldCharType="separate"))
    run4 = paragraph.add_run(placeholder)
    run5 = paragraph.add_run()
    run5._r.append(_el("w:fldChar", fldCharType="end"))


def update_fields_on_open(document: Document) -> None:
    settings = document.settings.element
    for old in settings.findall(qn("w:updateFields")):
        settings.remove(old)
    settings.insert_element_before(
        _el("w:updateFields", val="true"),
        "w:hdrShapeDefaults", "w:footnotePr", "w:endnotePr", "w:compat",
        "w:docVars", "w:rsids", "w:mathPr", "w:attachedSchema", "w:themeFontLang",
        "w:clrSchemeMapping", "w:doNotIncludeSubdocsInStats",
        "w:doNotAutoCompressPictures", "w:forceUpgrade", "w:captions",
        "w:shapeDefaults", "w:decimalSymbol", "w:listSeparator",
    )


# --------------------------------------------------------------------------
# document setup


def new_document() -> Document:
    document = Document()

    section = document.sections[0]
    section.page_height = Cm(29.7)
    section.page_width = Cm(21.0)
    section.top_margin = Cm(2.4)
    section.bottom_margin = Cm(2.2)
    section.left_margin = Cm(3.0)
    section.right_margin = Cm(2.5)
    section.header_distance = Cm(1.2)
    section.footer_distance = Cm(1.1)

    styles = document.styles

    normal = styles["Normal"]
    normal.font.name = BODY_FONT
    normal.font.size = Pt(10.5)
    normal.font.color.rgb = INK
    normal.paragraph_format.space_after = Pt(7)
    normal.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    normal.paragraph_format.line_spacing = 1.28
    rpr = normal.element.get_or_add_rPr()
    fonts = rpr.find(qn("w:rFonts"))
    if fonts is None:
        fonts = _el("w:rFonts")
        rpr.append(fonts)
    for attr in ("ascii", "hAnsi", "cs", "eastAsia"):
        fonts.set(qn(f"w:{attr}"), BODY_FONT)

    def heading_style(name, size, color, before, after, bold=True):
        style = styles[name]
        style.font.name = HEAD_FONT
        style.font.size = Pt(size)
        style.font.bold = bold
        style.font.color.rgb = color
        style.paragraph_format.space_before = Pt(before)
        style.paragraph_format.space_after = Pt(after)
        style.paragraph_format.keep_with_next = True
        style.paragraph_format.line_spacing = 1.0
        srpr = style.element.get_or_add_rPr()
        sfonts = srpr.find(qn("w:rFonts"))
        if sfonts is None:
            sfonts = _el("w:rFonts")
            srpr.append(sfonts)
        for attr in ("ascii", "hAnsi", "cs", "eastAsia"):
            sfonts.set(qn(f"w:{attr}"), HEAD_FONT)
        return style

    heading_style("Heading 1", 15.5, NAVY, 20, 8)
    heading_style("Heading 2", 12.5, NAVY, 15, 5)
    heading_style("Heading 3", 10.8, RGBColor(0x33, 0x4E, 0x68), 11, 4)

    for list_name in ("List Bullet", "List Number"):
        style = styles[list_name]
        style.font.name = BODY_FONT
        style.font.size = Pt(10.5)
        style.paragraph_format.space_after = Pt(4)
        style.paragraph_format.line_spacing = 1.24

    return document


def setup_header_footer(document: Document, header_left: str, header_right: str) -> None:
    section = document.sections[0]
    section.different_first_page_header_footer = True

    header = section.header
    para = header.paragraphs[0]
    para.text = ""
    para.paragraph_format.tab_stops.add_tab_stop(
        Cm(TEXT_WIDTH_CM), WD_TAB_ALIGNMENT.RIGHT
    )
    para.paragraph_format.space_after = Pt(2)
    run = para.add_run(header_left)
    run.font.size = Pt(8)
    run.font.name = HEAD_FONT
    run.font.color.rgb = GREY
    run = para.add_run("\t" + header_right)
    run.font.size = Pt(8)
    run.font.name = HEAD_FONT
    run.font.color.rgb = GREY

    bottom = OxmlElement("w:pBdr")
    bottom.append(_el("w:bottom", val="single", sz="4", space="4", color=LIGHT_RULE))
    para._p.get_or_add_pPr().insert_element_before(bottom, *P_SUCCESSORS)

    footer = section.footer
    fpara = footer.paragraphs[0]
    fpara.text = ""
    fpara.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = fpara.add_run("Page ")
    run.font.size = Pt(8)
    run.font.name = HEAD_FONT
    run.font.color.rgb = GREY
    add_field(fpara, "PAGE")
    for r in fpara.runs:
        r.font.size = Pt(8)
        r.font.name = HEAD_FONT
        r.font.color.rgb = GREY


# --------------------------------------------------------------------------
# content helpers


def md_runs(paragraph, text: str, size=None, italic=False, color=None,
            bold_all=False, font=None) -> None:
    """Adds runs, treating **pairs** of asterisks as bold markers."""
    for idx, chunk in enumerate(text.split("**")):
        if not chunk:
            continue
        run = paragraph.add_run(chunk)
        if size:
            run.font.size = Pt(size)
        run.italic = italic
        if bold_all or idx % 2 == 1:
            run.bold = True
        if color:
            run.font.color.rgb = color
        if font:
            run.font.name = font


def para(document, text: str = "", style=None, size=None, italic=False, bold=False,
         color=None, align=None, space_after=None, space_before=None, indent=None,
         line_spacing=None):
    p = document.add_paragraph(style=style)
    if text:
        md_runs(p, text, size=size, italic=italic, color=color, bold_all=bold)
    if align == "center":
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == "justify":
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    elif align == "right":
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    if space_after is not None:
        p.paragraph_format.space_after = Pt(space_after)
    if space_before is not None:
        p.paragraph_format.space_before = Pt(space_before)
    if indent is not None:
        p.paragraph_format.left_indent = Cm(indent)
    if line_spacing is not None:
        p.paragraph_format.line_spacing = line_spacing
    return p


def rich(document, parts, style=None, align="justify", space_after=None,
         space_before=None, indent=None, size=None, line_spacing=None):
    """parts: list of (text, {'bold':..,'italic':..,'color':..}) or plain strings."""
    p = document.add_paragraph(style=style)
    for part in parts:
        if isinstance(part, str):
            text, fmt = part, {}
        else:
            text, fmt = part
        run = p.add_run(text)
        if fmt.get("bold"):
            run.bold = True
        if fmt.get("italic"):
            run.italic = True
        if fmt.get("color"):
            run.font.color.rgb = fmt["color"]
        if fmt.get("size") or size:
            run.font.size = Pt(fmt.get("size") or size)
        if fmt.get("font"):
            run.font.name = fmt["font"]
    if align == "justify":
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    elif align == "center":
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if space_after is not None:
        p.paragraph_format.space_after = Pt(space_after)
    if space_before is not None:
        p.paragraph_format.space_before = Pt(space_before)
    if indent is not None:
        p.paragraph_format.left_indent = Cm(indent)
    if line_spacing is not None:
        p.paragraph_format.line_spacing = line_spacing
    return p


def bullets(document, items, size=None) -> None:
    for item in items:
        if isinstance(item, tuple):
            lead, rest = item
            p = document.add_paragraph(style="List Bullet")
            run = p.add_run(lead)
            run.bold = True
            if size:
                run.font.size = Pt(size)
            run2 = p.add_run(" " + rest)
            if size:
                run2.font.size = Pt(size)
        else:
            p = document.add_paragraph(item, style="List Bullet")
            if size:
                for run in p.runs:
                    run.font.size = Pt(size)


def numbered(document, items) -> None:
    for item in items:
        p = document.add_paragraph(style="List Number")
        p.paragraph_format.space_after = Pt(4)
        p.add_run(item)


def rule(document, color: str = LIGHT_RULE, space_before=2, space_after=8) -> None:
    p = document.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    borders = OxmlElement("w:pBdr")
    borders.append(_el("w:bottom", val="single", sz="6", space="1", color=color))
    p._p.get_or_add_pPr().insert_element_before(borders, *P_SUCCESSORS)


def caption(document, label: str, text: str) -> None:
    p = document.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(12)
    p.paragraph_format.keep_with_next = False
    run = p.add_run(label + "  ")
    run.font.size = Pt(8.5)
    run.font.bold = True
    run.font.name = HEAD_FONT
    run.font.color.rgb = NAVY
    run = p.add_run(text)
    run.font.size = Pt(8.5)
    run.font.name = HEAD_FONT
    run.font.color.rgb = GREY
    run.italic = True


def figure(document, path: str, width_cm: float, label: str, text: str) -> None:
    p = document.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.keep_with_next = True
    p.add_run().add_picture(path, width=Cm(width_cm))
    cpara = document.add_paragraph()
    cpara.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cpara.paragraph_format.space_after = Pt(13)
    run = cpara.add_run(label + "  ")
    run.font.size = Pt(8.5)
    run.bold = True
    run.font.name = HEAD_FONT
    run.font.color.rgb = NAVY
    run = cpara.add_run(text)
    run.font.size = Pt(8.5)
    run.font.name = HEAD_FONT
    run.font.color.rgb = GREY
    run.italic = True


def table(document, headers, rows, widths, label=None, caption_text=None,
          font_size=8.6, header_size=8.6, align_center_cols=(), zebra=True):
    if label and caption_text:
        cpara = document.add_paragraph()
        cpara.paragraph_format.space_before = Pt(8)
        cpara.paragraph_format.space_after = Pt(4)
        cpara.paragraph_format.keep_with_next = True
        run = cpara.add_run(label + "  ")
        run.font.size = Pt(8.5)
        run.bold = True
        run.font.name = HEAD_FONT
        run.font.color.rgb = NAVY
        run = cpara.add_run(caption_text)
        run.font.size = Pt(8.5)
        run.font.name = HEAD_FONT
        run.font.color.rgb = GREY
        run.italic = True

    tbl = document.add_table(rows=1, cols=len(headers))
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_borders(tbl)

    hdr = tbl.rows[0]
    keep_row_together(hdr)
    repeat_header(hdr)
    for idx, head in enumerate(headers):
        cell = hdr.cells[idx]
        cell.text = ""
        shade(cell, NAVY_HEX)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        cell_margins(cell, top=70, bottom=70)
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.0
        run = p.add_run(head)
        run.font.size = Pt(header_size)
        run.font.bold = True
        run.font.name = TABLE_FONT
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        if idx in align_center_cols:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    for r_idx, row in enumerate(rows):
        cells = tbl.add_row()
        keep_row_together(cells)
        for c_idx, value in enumerate(row):
            cell = cells.cells[c_idx]
            cell.text = ""
            cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
            cell_margins(cell, top=55, bottom=55)
            if zebra and r_idx % 2 == 1:
                shade(cell, BAND_HEX)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.12
            lines = str(value).split("\n")
            for l_idx, line in enumerate(lines):
                if l_idx:
                    p = cell.add_paragraph()
                    p.paragraph_format.space_after = Pt(0)
                    p.paragraph_format.line_spacing = 1.12
                md_runs(p, line, size=font_size, font=TABLE_FONT)
                if c_idx in align_center_cols:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    set_col_widths(tbl, widths)

    if label and caption_text:
        spacer = document.add_paragraph()
        spacer.paragraph_format.space_before = Pt(0)
        spacer.paragraph_format.space_after = Pt(10)
        for run in spacer.runs:
            run.font.size = Pt(2)
    else:
        document.add_paragraph().paragraph_format.space_after = Pt(6)

    return tbl


def callout(document, title: str, body: str, fill: str = AMBER_FILL,
            edge: str = AMBER_EDGE, title_color=NAVY, caps: bool = True):
    tbl = document.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    left_accent(tbl, edge, size=18)
    keep_row_together(tbl.rows[0])
    cell = tbl.rows[0].cells[0]
    shade(cell, fill)
    cell_margins(cell, top=130, bottom=130, left=170, right=150)

    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(3)
    run = p.add_run(title)
    run.font.size = Pt(8.8)
    run.font.bold = True
    run.font.name = HEAD_FONT
    run.font.color.rgb = title_color
    run.font.all_caps = caps

    for idx, chunk in enumerate(body.split("\n\n")):
        p = cell.add_paragraph()
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.26
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        run = p.add_run(chunk)
        run.font.size = Pt(9.2)
        run.font.name = BODY_FONT
        run.font.color.rgb = INK

    set_col_widths(tbl, [TEXT_WIDTH_CM])
    spacer = document.add_paragraph()
    spacer.paragraph_format.space_after = Pt(2)
    spacer.paragraph_format.space_before = Pt(0)
    for run in spacer.runs:
        run.font.size = Pt(2)


def reflection(document, text: str, extra: str | None = None) -> None:
    title = "Group reflection"
    body = text
    if extra:
        body = text + "\n\n" + extra
    callout(
        document,
        title + "  ·  written by the group, not by AI",
        body,
        fill=BLUE_FILL,
        edge=BLUE_EDGE,
        caps=False,
    )


def two_col_table(document, rows, widths=(4.6, 10.9), font_size=9.0):
    tbl = document.add_table(rows=0, cols=2)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_borders(tbl)
    for idx, (key, value) in enumerate(rows):
        cells = tbl.add_row()
        keep_row_together(cells)
        for c_idx, text in enumerate((key, value)):
            cell = cells.cells[c_idx]
            cell.text = ""
            cell_margins(cell, top=60, bottom=60)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
            if c_idx == 0:
                shade(cell, BAND_HEX)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.15
            run = p.add_run(text)
            run.font.size = Pt(font_size)
            run.font.name = TABLE_FONT
            run.bold = c_idx == 0
    set_col_widths(tbl, list(widths))
    return tbl


def page_break(document) -> None:
    from docx.enum.text import WD_BREAK

    p = document.add_paragraph()
    p.paragraph_format.space_after = Pt(0)
    p.add_run().add_break(WD_BREAK.PAGE)


__all__ = [
    "Document",
    "new_document",
    "setup_header_footer",
    "para",
    "rich",
    "bullets",
    "numbered",
    "rule",
    "caption",
    "figure",
    "table",
    "callout",
    "reflection",
    "two_col_table",
    "page_break",
    "update_fields_on_open",
    "add_field",
    "NAVY",
    "INK",
    "GREY",
    "TEXT_WIDTH_CM",
    "WD_SECTION",
]
