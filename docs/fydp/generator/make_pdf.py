"""Render a reading copy of the FYDP-1 deliverable as PDF.

The Word file is the submission; this script exists so the same text can be
read on a phone or in a browser without Word. It is built from the .docx
itself - mammoth converts the document body to HTML, xhtml2pdf lays that HTML
out on A4 - so the two files cannot drift apart.

Formatting is approximate. Cell shading, italic and small-caps runs are
flattened to plain text, page breaks are not preserved, and the sandbox has
no Georgia, so a bundled DejaVu serif is used instead. The words are exact.

    python3 make_pdf.py            # writes deliverables/..._FYDP1.pdf
"""

from __future__ import annotations

import os
import re
import shutil
import sys
import tempfile
from io import BytesIO

import mammoth
from PIL import Image
from xhtml2pdf import pisa

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.normpath(os.path.join(HERE, "..", "..", ".."))
DOCX = os.path.join(
    ROOT, "deliverables", "23FYDP24_SamaanShare_MethodologyBaseline_FYDP1.docx"
)
PDF = os.path.join(
    ROOT, "deliverables", "23FYDP24_SamaanShare_MethodologyBaseline_FYDP1.pdf"
)

FONT_DIRS = [
    "/usr/share/fonts/truetype/dejavu",
    "/usr/share/fonts/dejavu",
    "/Library/Fonts",
    "/System/Library/Fonts",
]

PT_PER_CM = 72 / 2.54
CONTENT_WIDTH_PT = 17.0 * PT_PER_CM  # A4 21 cm less 2 cm margins each side
MAX_FIGURE_HEIGHT_PT = 19.0 * PT_PER_CM

CSS = """
@font-face { font-family: Body; src: url("DejaVuSerif.ttf"); }
@font-face { font-family: BodyBold; src: url("DejaVuSerif-Bold.ttf"); }

@page {
    size: A4;
    margin: 1.9cm 2.0cm 1.7cm 2.0cm;
    @bottom-center {
        content: counter(page);
        font-family: Body;
        font-size: 8pt;
        color: #6B7280;
    }
}
body {
    font-family: Body;
    font-size: 9.3pt;
    line-height: 1.34;
    color: #252A33;
}
p { margin: 0 0 4.5pt 0; }
strong, b { font-family: BodyBold; }
h1, h2, h3, h4 { font-family: BodyBold; color: #1F3864; page-break-after: avoid; }
h1 { font-size: 15pt; margin: 14pt 0 7pt 0; }
h2 { font-size: 12pt; margin: 12pt 0 5pt 0; }
h3 { font-size: 10.3pt; margin: 10pt 0 4pt 0; }
h4 { font-size: 9.6pt; margin: 8pt 0 3pt 0; }
table {
    width: 100%;
    margin: 5pt 0 9pt 0;
    font-size: 7.6pt;
    line-height: 1.26;
}
td, th {
    border: 0.5pt solid #C9CED6;
    padding: 2.6pt 3.4pt;
    vertical-align: top;
}
td p { margin: 0; }
p.figure { text-align: center; margin: 8pt 0 10pt 0; page-break-inside: avoid; }
"""

STYLE_MAP = """
p[style-name='Caption'] => p.caption
p[style-name='Title'] => h1.title
"""

_saved_images: dict[str, tuple[int, int]] = {}


def convert_image(image):
    """Write each embedded image next to the fonts and remember its size."""
    with image.open() as stream:
        data = stream.read()
    ext = {"image/png": "png", "image/jpeg": "jpg", "image/gif": "gif"}.get(
        image.content_type, "png"
    )
    name = f"figure{len(_saved_images) + 1}.{ext}"
    path = os.path.join(_workdir, name)
    with open(path, "wb") as fh:
        fh.write(data)
    try:
        with Image.open(path) as im:
            _saved_images[name] = im.size
    except Exception:
        _saved_images[name] = (1000, 600)
    return {"src": name}


def size_images(html: str) -> str:
    """Give every <img> an explicit point size that fits the text column."""

    def repl(match: re.Match[str]) -> str:
        src = match.group(1)
        width, height = _saved_images.get(src, (1000, 600))
        draw_w = CONTENT_WIDTH_PT * 0.92
        draw_h = draw_w * height / width
        if draw_h > MAX_FIGURE_HEIGHT_PT:
            draw_h = MAX_FIGURE_HEIGHT_PT
            draw_w = draw_h * width / height
        return (
            f'<img src="{src}" width="{draw_w:.0f}pt" height="{draw_h:.0f}pt">'
        )

    return re.sub(r'<img\s+src="([^"]+)"\s*/?>', repl, html)


def find_font_dir() -> str | None:
    for font_dir in FONT_DIRS:
        if os.path.exists(os.path.join(font_dir, "DejaVuSerif.ttf")):
            return font_dir
    return None


_workdir = tempfile.mkdtemp(prefix="fydp-pdf-")


def main() -> int:
    if not os.path.exists(DOCX):
        print("missing:", DOCX)
        return 1

    font_dir = find_font_dir()
    if font_dir:
        for name in ("DejaVuSerif.ttf", "DejaVuSerif-Bold.ttf"):
            src = os.path.join(font_dir, name)
            if os.path.exists(src):
                shutil.copy2(src, _workdir)

    with open(DOCX, "rb") as fh:
        result = mammoth.convert_to_html(
            fh, style_map=STYLE_MAP, convert_image=mammoth.images.img_element(convert_image)
        )

    for message in result.messages[:5]:
        print("mammoth:", str(message)[:160])

    body = size_images(result.value)
    html = (
        "<html><head><meta charset='utf-8'><style>{css}</style></head>"
        "<body>{body}</body></html>"
    ).format(css=CSS, body=body)

    # xhtml2pdf only reads resource files beneath the current directory, so the
    # fonts and figures are written into a scratch directory and read from there.
    previous = os.getcwd()
    buffer = BytesIO()
    try:
        os.chdir(_workdir)
        status = pisa.CreatePDF(BytesIO(html.encode("utf-8")), dest=buffer, encoding="utf-8")
    finally:
        os.chdir(previous)

    if status.err:
        print("pdf errors:", status.err)
        return 1

    data = buffer.getvalue()
    with open(PDF, "wb") as fh:
        fh.write(data)

    try:
        import pypdfium2 as pdfium

        pages = len(pdfium.PdfDocument(PDF))
    except Exception:  # pragma: no cover - optional dependency
        pages = "?"

    shutil.rmtree(_workdir, ignore_errors=True)
    print(f"wrote {PDF} ({len(data) / 1024:.0f} KB, {pages} pages, {len(_saved_images)} figures)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
