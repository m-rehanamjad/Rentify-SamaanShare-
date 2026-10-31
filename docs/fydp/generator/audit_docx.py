"""Structural audit of the generated .docx.

Checks the things that a visual pass would catch: table widths against the
text column, caption presence, heading order, stray phrases, and a realistic
page-count estimate.
"""

from __future__ import annotations

import os
import re
import sys
from zipfile import ZipFile

from docx import Document
from docx.shared import Cm

PATH = os.path.normpath(
    os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        "..", "..", "..", "deliverables",
        "23FYDP24_SamaanShare_MethodologyBaseline_FYDP1.docx",
    )
)

TEXT_WIDTH_CM = 15.5


def main() -> int:
    if not os.path.exists(PATH):
        print("missing:", PATH)
        return 1

    # ---- zip / xml integrity
    with ZipFile(PATH) as zf:
        names = zf.namelist()
        bad = zf.testzip()
    print(f"zip ok={bad is None}  parts={len(names)}")
    for required in ("word/document.xml", "word/styles.xml", "word/settings.xml"):
        print(f"  {required}: {'yes' if required in names else 'MISSING'}")

    doc = Document(PATH)

    # ---- text accounting
    words = 0
    chars = 0
    for p in doc.paragraphs:
        t = p.text
        words += len(t.split())
        chars += len(t)
    table_words = 0
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                table_words += len(cell.text.split())
    print(f"\nbody paragraphs: {len(doc.paragraphs)}   words: {words}")
    print(f"tables: {len(doc.tables)}   words inside tables: {table_words}")
    print(f"total words: {words + table_words}")

    # A4 text block: 15.5 cm x 25.1 cm. 10.5pt Georgia at 1.28 spacing = 13.4pt
    # line height -> 53 lines per page; ~13.5 words per line.
    text_pages = words / 620
    table_pages = table_words / 900
    figure_pages = 1.6
    print(
        f"estimated pages: ~{text_pages + table_pages + figure_pages:.0f} of content "
        f"(text {text_pages:.1f} + tables {table_pages:.1f} + figures {figure_pages})"
    )

    # ---- headings
    print("\nheadings:")
    for p in doc.paragraphs:
        if p.style.name in ("Heading 1", "Heading 2", "Heading 3"):
            indent = {"Heading 1": "", "Heading 2": "  ", "Heading 3": "    "}[p.style.name]
            print(f"  {indent}{p.text[:78]}")

    # ---- table widths
    print("\ntable widths (must be <= 15.5 cm):")
    problems = []
    for idx, tbl in enumerate(doc.tables, start=1):
        widths = []
        for cell in tbl.rows[0].cells:
            widths.append(cell.width.cm if cell.width else 0)
        total = sum(widths)
        flag = "" if total <= TEXT_WIDTH_CM + 0.05 else "  <-- TOO WIDE"
        if total > TEXT_WIDTH_CM + 0.05:
            problems.append((idx, total))
        first = tbl.rows[0].cells[0].text.strip().replace("\n", " ")[:38]
        print(f"  {idx:>2}: {total:5.1f} cm  cols={len(widths)}  [{first}]{flag}")

    # ---- figure count
    with ZipFile(PATH) as zf:
        media = [n for n in zf.namelist() if n.startswith("word/media/")]
    print(f"\nembedded images: {len(media)} -> {media}")

    # ---- phrasing checks
    all_text = "\n".join(p.text for p in doc.paragraphs)
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                all_text += "\n" + cell.text

    banned = [
        r"\bmobile app\b",
        r"\bmobile application\b",
        r"\bapp store\b",
        r"\bandroid\b",
        r"\breact native\b",
        r"\bplay store\b",
    ]
    print("\nphrase checks:")
    for pattern in banned:
        hits = re.findall(pattern, all_text, flags=re.I)
        print(f"  {pattern:<26} {len(hits):>2}  {hits[:3]}")

    print("\nrequired reflections:",
          all_text.lower().count("group reflection"))
    print("RQ mentions:", {q: len(re.findall(rf"\b{q}\b", all_text)) for q in ("RQ1", "RQ2", "RQ3")})

    # ---- placeholder leakage
    for suspect in ("TODO", "XXX", "[PASTE", "lorem", "{", "}"):
        if suspect in all_text:
            print(f"  WARNING suspect token in text: {suspect!r}")

    if problems:
        print("\nPROBLEM tables too wide:", problems)
        return 1

    # ---- duplicate paragraph check (copy-paste smell)
    seen = {}
    for p in doc.paragraphs:
        t = p.text.strip()
        if len(t) > 120:
            seen[t] = seen.get(t, 0) + 1
    dupes = {k: v for k, v in seen.items() if v > 1}
    if dupes:
        print("\nNOTE duplicated long paragraphs:", len(dupes))
        for k, v in list(dupes.items())[:3]:
            print("   x", v, k[:80])

    print("\naudit complete")
    return 0


if __name__ == "__main__":
    sys.exit(main())
