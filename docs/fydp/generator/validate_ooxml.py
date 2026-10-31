"""Checks that generated property elements appear in the order OOXML requires.

Word is tolerant about many things and not about others; an out-of-order
`w:shd` inside `w:tcPr` is the sort of thing that turns a submission into
"Word found unreadable content". This validates document, headers, footers
and settings against the schema sequences.
"""

from __future__ import annotations

import os
import sys
from zipfile import ZipFile

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

SEQUENCES = {
    "tblPr": [
        "tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize",
        "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing", "tblInd",
        "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook", "tblCaption",
        "tblDescription", "tblPrChange",
    ],
    "tcPr": [
        "cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd",
        "noWrap", "tcMar", "textDirection", "tcFitText", "vAlign", "hideMark",
        "headers", "cellIns", "cellDel", "cellMerge", "tcPrChange",
    ],
    "trPr": [
        "cnfStyle", "divId", "gridBefore", "gridAfter", "wBefore", "wAfter",
        "cantSplit", "trHeight", "tblHeader", "tblCellSpacing", "jc", "hidden",
        "ins", "del", "trPrChange",
    ],
    "pPr": [
        "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr",
        "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs",
        "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct",
        "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
        "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents",
        "suppressOverlap", "jc", "textDirection", "textAlignment",
        "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr",
        "pPrChange",
    ],
    "rPr": [
        "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike",
        "dstrike", "outline", "shadow", "emboss", "imprint", "noProof",
        "snapToGrid", "vanish", "webHidden", "color", "spacing", "w", "kern",
        "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd",
        "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout",
        "specVanish", "oMath",
    ],
    "settings": None,
    "tblBorders": ["top", "start", "left", "bottom", "end", "right", "insideH", "insideV"],
    "tcMar": ["top", "start", "left", "bottom", "end", "right"],
}

SETTINGS_ORDER = [
    "writeProtection", "view", "zoom", "removePersonalInformation",
    "removeDateAndTime", "doNotDisplayPageBoundaries", "displayBackgroundShape",
    "printPostScriptOverText", "printFractionalCharacterWidth",
    "printFormsData", "embedTrueTypeFonts", "embedSystemFonts",
    "saveSubsetFonts", "saveFormsData", "mirrorMargins",
    "alignBordersAndEdges", "bordersDoNotSurroundHeader",
    "bordersDoNotSurroundFooter", "gutterAtTop", "hideSpellingErrors",
    "hideGrammaticalErrors", "activeWritingStyle", "proofState", "formsDesign",
    "attachedTemplate", "linkStyles", "stylePaneFormatFilter",
    "stylePaneSortMethod", "documentType", "mailMerge", "revisionView",
    "trackRevisions", "doNotTrackMoves", "doNotTrackFormatting",
    "documentProtection", "autoFormatOverride", "styleLockTheme",
    "styleLockQFSet", "defaultTabStop", "autoHyphenation",
    "consecutiveHyphenLimit", "hyphenationZone", "doNotHyphenateCaps",
    "showEnvelope", "summaryLength", "clickAndTypeStyle", "defaultTableStyle",
    "evenAndOddHeaders", "bookFoldRevPrinting", "bookFoldPrinting",
    "bookFoldPrintingSheets", "drawingGridHorizontalSpacing",
    "drawingGridVerticalSpacing", "displayHorizontalDrawingGridEvery",
    "displayVerticalDrawingGridEvery", "doNotUseMarginsForDrawingGridOrigin",
    "drawingGridHorizontalOrigin", "drawingGridVerticalOrigin",
    "doNotShadeFormData", "noPunctuationKerning", "characterSpacingControl",
    "printTwoOnOne", "strictFirstAndLastChars", "noLineBreaksAfter",
    "noLineBreaksBefore", "savePreviewPicture", "doNotValidateAgainstSchema",
    "saveInvalidXml", "ignoreMixedContent", "alwaysShowPlaceholderText",
    "doNotDemarcateInvalidXml", "saveXmlDataOnly", "useXSLTWhenSaving",
    "saveThroughXslt", "showXMLTags", "alwaysMergeEmptyNamespace",
    "updateFields", "hdrShapeDefaults", "footnotePr", "endnotePr", "compat",
    "docVars", "rsids", "mathPr", "attachedSchema", "themeFontLang",
    "clrSchemeMapping", "doNotIncludeSubdocsInStats", "doNotAutoCompressPictures",
    "forceUpgrade", "captions", "readModeInkLockDown", "smartTagType",
    "schemaLibrary", "shapeDefaults", "doNotEmbedSmartTags", "decimalSymbol",
    "listSeparator",
]


def check_part(xml: bytes, name: str) -> list[str]:
    root = etree.fromstring(xml)
    problems: list[str] = []

    for seq_name, order in SEQUENCES.items():
        if order is None:
            continue
        for el in root.iter(f"{{{W}}}{seq_name}"):
            kids = [etree.QName(k).localname for k in el]
            unknown = [k for k in kids if k not in order]
            if unknown:
                continue
            positions = [order.index(k) for k in kids]
            if positions != sorted(positions):
                problems.append(f"{name}: <{seq_name}> out of order -> {kids}")
            if len(set(kids)) != len(kids):
                dupes = {k for k in kids if kids.count(k) > 1}
                problems.append(f"{name}: <{seq_name}> duplicate {sorted(dupes)}")

    if name.endswith("settings.xml"):
        kids = [etree.QName(k).localname for k in root]
        positions = [SETTINGS_ORDER.index(k) for k in kids if k in SETTINGS_ORDER]
        if positions != sorted(positions):
            problems.append(f"settings.xml: children out of order -> {kids}")
        if len(kids) != len(set(kids)):
            problems.append(f"settings.xml: duplicate children -> {kids}")

    return problems


def main() -> int:
    path = os.path.normpath(
        os.path.join(
            os.path.dirname(os.path.abspath(__file__)),
            "..", "..", "..", "deliverables",
            "23FYDP24_SamaanShare_MethodologyBaseline_FYDP1.docx",
        )
    )
    if not os.path.exists(path):
        print("missing", path)
        return 1

    all_problems: list[str] = []
    with ZipFile(path) as zf:
        targets = [
            n for n in zf.namelist()
            if n.startswith("word/") and n.endswith(".xml")
            and any(t in n for t in ("document", "header", "footer", "settings", "styles"))
        ]
        for part in targets:
            all_problems += check_part(zf.read(part), part)

    if all_problems:
        print(f"{len(all_problems)} problem(s):")
        for p in all_problems[:20]:
            print("  -", p)
        return 1

    print(f"ok: {len(targets)} parts, no ordering or duplicate problems")
    return 0


if __name__ == "__main__":
    sys.exit(main())
