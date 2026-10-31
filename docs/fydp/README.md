# FYDP-1 Proposed Methodology Baseline

The submitted document is:

```
deliverables/23FYDP24_SamaanShare_MethodologyBaseline_FYDP1.docx
```

It covers Heading 3 (Problem Statement, Objectives and Research Questions) and
Heading 4 (Proposed Methodology: 4.1 System Architecture, 4.2 Dataset and Data
Collection Strategy, 4.3 Proposed Algorithm / Core Technique, 4.4 Traceability
Matrix), with five group reflections, a reference list, and three appendices.

## Before you submit — five things to change

1. **AI tool names.** Appendix A names *Claude* as the drafting tool and
   *ChatGPT* as the verifier, matching the disclosure in Chapter 2. If the group
   actually used different tools for this chapter, edit the table and the
   paragraph beneath it.
2. **Pilot numbers.** Section 4.2.3 assumes 30–50 completed rentals, 60–100
   listings and 5–10 claims. Change these to whatever the group is actually
   promising before they become a target you are measured against.
3. **Supervisor sign-off.** The title page says "Submitted to FYDP Coordination
   Committee". Add the sign-off line the department requires once
   Dr. Sarwat Iqbal has reviewed the draft.
4. **Contents page.** Open the file in Word, press Ctrl+A then F9, and choose
   "Update entire table". Page numbers are fields and are blank until refreshed.
5. **Group reflections.** These were written to sound like a group speaking
   honestly. Read them once and change anything that is not true of you — an
   examiner can tell when a reflection has been borrowed.

## Rebuilding the document

Everything in `generator/` regenerates the Word file from scratch:

```bash
python3 -m venv .venv && .venv/bin/pip install python-docx matplotlib
cd docs/fydp/generator
.venv/bin/python make_figures.py     # rewrites architecture.png, deposit_flow.png
.venv/bin/python build_docx.py       # rewrites the .docx
.venv/bin/python validate_ooxml.py   # checks element order Word cares about
.venv/bin/python audit_docx.py       # word counts, table widths, phrasing checks
```

Text lives in `build_docx.py`; layout, styles and table helpers live in
`docx_kit.py`. Figures are drawn in `make_figures.py` at 300 DPI, above the
department's 150 DPI minimum.

## What is deliberately in the document

- **Every figure and number is checked against the repository**, not invented:
  27 database tables, 24 migrations, 52 test files, 7 categories and 30
  subcategories, deposit thresholds of PKR 25,000 and 100,000, score weights
  0.45 / 0.25 / 0.20 / 0.10, windows of 48 hours, 7 days and 14 days.
- **The traceability matrix has 21 rows**, each with a real source from the
  approved literature review and a value of RQ1, RQ2 or RQ3 — no free text.
- **Appendix B records the five places where this chapter departs from the
  original proposal** (single database, browser-based application rather than an
  installed phone build, custodial deposit, three objectives instead of seven,
  manual payment verification). Say this before a supervisor finds it.
- **Appendix C states which parts are built and which are planned**, so nothing
  reads as finished when it is not.

## Length

The body (Heading 3 and 4.1–4.3) runs to roughly 12 pages with its figures and
tables; the department's guidance is 6–9 excluding the matrix and references. If
your supervisor enforces the limit strictly, the easiest cuts are 4.1.1 (the three
paragraphs duplicate Figure 1), the ten numbered steps in 4.1.3 (Figure 2 already
tells that story) and 4.2.5 (five paragraphs to three). Do not cut 4.3.3: the two
worked examples in Table 9 are what make the trust score checkable.
