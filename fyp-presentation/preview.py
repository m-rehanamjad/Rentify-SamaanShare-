"""Approximate raster preview of the generated deck.

No Office/LibreOffice renderer is available in this sandbox, so this walks the
shape tree and paints geometry + wrapped text with PIL. Metrics are DejaVu rather
than Segoe UI, so treat it as a layout check, not a pixel-accurate proof.
"""

import os
import sys
import math

from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Emu

SCALE = 120            # px per inch  -> 1600 x 900
FONT_DIR = "/usr/share/fonts/truetype/dejavu"
REG = os.path.join(FONT_DIR, "DejaVuSans.ttf")
BOLD = os.path.join(FONT_DIR, "DejaVuSans-Bold.ttf")
FALLBACK = os.path.join(FONT_DIR, "DejaVuSans.ttf")

_cache = {}


def font(size_pt, bold):
    key = (round(size_pt, 1), bold)
    if key not in _cache:
        path = BOLD if bold else REG
        _cache[key] = ImageFont.truetype(path, int(round(size_pt * SCALE / 72.0)))
    return _cache[key]


def px(emu):
    return int(round(Emu(emu).inches * SCALE))


def solid_rgb(fill):
    try:
        if fill.type is not None and fill.type == 1:  # MSO_FILL.SOLID
            c = fill.fore_color.rgb
            return (c[0], c[1], c[2])
    except Exception:
        pass
    return None


def line_rgb(shape):
    try:
        if shape.line.fill.type == 1:
            c = shape.line.color.rgb
            return (c[0], c[1], c[2])
    except Exception:
        pass
    return None


def wrap_runs(paragraph, max_w):
    """Return list of lines; each line is a list of (text, font, colour)."""
    runs = []
    for r in paragraph.runs:
        size = r.font.size.pt if r.font.size else 13
        bold = bool(r.font.bold)
        try:
            col = r.font.color.rgb
            colour = (col[0], col[1], col[2])
        except Exception:
            colour = (20, 43, 59)
        runs.append((r.text, font(size, bold), colour, size))
    if not runs:
        return []

    lines, cur, cur_w = [], [], 0.0
    space_w = None
    for text, f, colour, size in runs:
        space_w = f.getlength(" ")
        # split keeping spaces
        tokens = []
        buf = ""
        for ch in text:
            if ch == " ":
                if buf:
                    tokens.append(buf)
                    buf = ""
                tokens.append(" ")
            else:
                buf += ch
        if buf:
            tokens.append(buf)
        for tok in tokens:
            w = f.getlength(tok)
            if tok != " " and cur_w + w > max_w and cur:
                lines.append(cur)
                cur, cur_w = [], 0.0
            if tok == " " and not cur:
                continue
            cur.append((tok, f, colour, size))
            cur_w += w
    if cur:
        lines.append(cur)
    return lines


def draw_text_frame(d, shape, box):
    tf = shape.text_frame
    x0, y0, x1, y1 = box
    ml = px(tf.margin_left or 0)
    mr = px(tf.margin_right or 0)
    mt = px(tf.margin_top or 0)
    mb = px(tf.margin_bottom or 0)
    max_w = max(4, (x1 - x0) - ml - mr)
    blocks = []
    total_h = 0
    for p in tf.paragraphs:
        lines = wrap_runs(p, max_w)
        p_size = max([r[3] for line in lines for r in line], default=13)
        lh = p_size * 1.22 * SCALE / 72.0
        gap = 0
        if p.space_after is not None:
            gap = p.space_after.pt * SCALE / 72.0
        total_h += lh * max(1, len(lines)) + gap
        blocks.append((p, lines, lh, gap))

    anchor = str(tf.vertical_anchor)
    y = y0 + mt
    if "MIDDLE" in anchor:
        y = y0 + ((y1 - y0) - total_h) / 2
    elif "BOTTOM" in anchor:
        y = y1 - mb - total_h

    for p, lines, lh, gap in blocks:
        align = str(p.alignment)
        for line in lines:
            w = sum(f.getlength(t) for t, f, c, s in line)
            if "CENTER" in align:
                x = x0 + ml + (max_w - w) / 2
            elif "RIGHT" in align:
                x = x0 + (x1 - x0) - mr - w
            else:
                x = x0 + ml
            for text, f, colour, size in line:
                d.text((x, y), text, font=f, fill=colour)
                x += f.getlength(text)
            y += lh
        y += gap
    return y


def paste_picture(canvas, shape):
    try:
        img = Image.open(shape.image.blob and __import__("io").BytesIO(shape.image.blob))
        img = img.convert("RGBA")
        w, h = px(shape.width), px(shape.height)
        img = img.resize((max(1, w), max(1, h)))
        canvas.paste(img, (px(shape.left), px(shape.top)), img)
    except Exception as exc:
        print("  ! picture failed", exc)


def draw_table(canvas, d, shape):
    tbl = shape.table
    x0, y0 = px(shape.left), px(shape.top)
    y = y0
    for r, row in enumerate(tbl.rows):
        h = px(row.height)
        x = x0
        for c, col in enumerate(tbl.columns):
            w = px(col.width)
            cell = tbl.cell(r, c)
            fill = solid_rgb(cell.fill)
            if fill:
                d.rectangle([x, y, x + w, y + h], fill=fill)
            d.rectangle([x, y, x + w, y + h], outline=(213, 223, 231))
            # text
            for p in cell.text_frame.paragraphs:
                lines = wrap_runs(p, w - 2 * px(cell.margin_left or 0))
                p_size = max([run[3] for line in lines for run in line], default=10)
                lh = p_size * 1.2 * SCALE / 72.0
                ty = y + (h - lh * len(lines)) / 2
                for line in lines:
                    lw = sum(f.getlength(t) for t, f, c2, s in line)
                    align = str(p.alignment or cell.text_frame.paragraphs[0].alignment)
                    align = str(cell.text_frame.paragraphs[0].alignment)
                    tx = x + px(cell.margin_left or 0)
                    if "RIGHT" in align:
                        tx = x + w - px(cell.margin_right or 0) - lw
                    for text, f, colour, size in line:
                        d.text((tx, ty), text, font=f, fill=colour)
                        tx += f.getlength(text)
                    ty += lh
            x += w
        y += h


def render(pptx_path, outdir):
    prs = Presentation(pptx_path)
    W = int(prs.slide_width.inches * SCALE)
    H = int(prs.slide_height.inches * SCALE)
    os.makedirs(outdir, exist_ok=True)
    paths = []
    for idx, slide in enumerate(prs.slides, 1):
        canvas = Image.new("RGB", (W, H), "white")
        d = ImageDraw.Draw(canvas)
        for shape in slide.shapes:
            st = shape.shape_type
            box = (px(shape.left), px(shape.top),
                   px(shape.left + shape.width), px(shape.top + shape.height))
            if st == MSO_SHAPE_TYPE.PICTURE:
                paste_picture(canvas, shape)
                d = ImageDraw.Draw(canvas)
                continue
            if shape.has_table:
                draw_table(canvas, d, shape)
                continue
            if st == MSO_SHAPE_TYPE.LINE:
                d.line([box[0], box[1], box[2], box[3]], fill=line_rgb(shape) or (200, 200, 200),
                       width=max(1, int(round(shape.line.width.pt * SCALE / 72.0))) if shape.line.width else 1)
                continue
            fill = solid_rgb(shape.fill) if shape.fill.type is not None else None
            outline = line_rgb(shape)
            name = ""
            try:
                name = str(shape.auto_shape_type)
            except Exception:
                pass
            if "OVAL" in name:
                d.ellipse(box, fill=fill, outline=outline)
            elif "ROUNDED_RECTANGLE" in name:
                r = 10
                try:
                    r = int(shape.adjustments[0] * min(box[2] - box[0], box[3] - box[1]))
                except Exception:
                    pass
                d.rounded_rectangle(box, radius=max(2, r), fill=fill, outline=outline)
            elif "RIGHT_ARROW" in name:
                x0, y0, x1, y1 = box
                mid = (y0 + y1) / 2
                d.polygon([(x0, mid - (y1 - y0) / 4), (x1 - (x1 - x0) * 0.4, mid - (y1 - y0) / 4),
                           (x1 - (x1 - x0) * 0.4, y0), (x1, mid),
                           (x1 - (x1 - x0) * 0.4, y1), (x1 - (x1 - x0) * 0.4, mid + (y1 - y0) / 4),
                           (x0, mid + (y1 - y0) / 4)], fill=fill or (220, 220, 220))
            elif "DOWN_ARROW" in name:
                x0, y0, x1, y1 = box
                mid = (x0 + x1) / 2
                d.polygon([(mid - (x1 - x0) / 4, y0), (mid + (x1 - x0) / 4, y0),
                           (mid + (x1 - x0) / 4, y1 - (y1 - y0) * 0.4), (x1, y1 - (y1 - y0) * 0.4),
                           (mid, y1), (x0, y1 - (y1 - y0) * 0.4),
                           (mid - (x1 - x0) / 4, y1 - (y1 - y0) * 0.4)], fill=fill or (220, 220, 220))
            else:
                if fill or outline:
                    d.rectangle(box, fill=fill, outline=outline)
            if shape.has_text_frame and shape.text_frame.text.strip():
                draw_text_frame(d, shape, box)
        path = os.path.join(outdir, f"slide-{idx:02d}.png")
        canvas.save(path)
        paths.append(path)
    return paths


if __name__ == "__main__":
    src = sys.argv[1]
    out = sys.argv[2]
    ps = render(src, out)
    print(f"rendered {len(ps)} slides to {out}")
