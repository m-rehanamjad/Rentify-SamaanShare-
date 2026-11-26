"""Slides 24-34: evaluation, results, limitations, contributions, roadmap, close."""

from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt

from deck_theme import (
    NAVY, NAVY_SOFT, EMERALD, EMERALD_DK, EMERALD_BR, MINT, INK, MUTED,
    LINE, PAPER, SLATE, WHITE, AMBER, AMBER_BR, AMBER_SOFT, RED, RED_SOFT,
    SW, SH, M, CW, TOP, BOTTOM,
    bg, rect, card, pill, hline, arrow_right, arrow_down, circle_num,
    textbox, paragraph, bullets, table, content_slide, lead, caption, picture,
    pg, set_alpha,
)

A = "fyp-presentation/assets/"
MINT_TXT = RGBColor(0x8F, 0xC9, 0xB6)
TITLE_SUB = RGBColor(0xC9, 0xD8, 0xE2)


# ------------------------------------------------------------ 24 evaluation design
def evaluation_design(prs):
    s = content_slide(prs, "Chapter 3  \u00b7  Evaluation design",
                      "Every question has a stated test and a stated failure mode")
    rows = [
        ["Question", "Instrument", "Evidence produced", "Passed if\u2026"],
        ["RQ1 \u00b7 states",
         "Integration suite over every edge in the transition table, plus every illegal jump",
         "Accepted / rejected transition log; audit rows per booking",
         "Zero illegal transitions accepted; exactly one audit row per legal transition"],
        ["RQ2 \u00b7 deposit",
         "Six scripted scenarios, each run against an independently hand-calculated baseline",
         "Settlement figures per scenario; timers observed to fire",
         "Each settlement equals its baseline to the rupee; no balance held past its clock"],
        ["RQ3 \u00b7 trust",
         "Pilot histories plus offline sensitivity analysis over 2,000 modelled profiles",
         "Score distributions per behaviour class; band assignments; access decisions",
         "Completing members outscore defaulting ones across the modelled population"],
    ]
    table(s, M, 1.54, CW, rows, widths=[1.95, 3.30, 3.15, 3.32], row_h=1.04,
          header_h=0.40, size=9.2, hsize=9, zebra=PAPER, bold_first=True,
          valign=MSO_ANCHOR.MIDDLE)

    y = 5.22
    card(s, M, y, CW, 1.42, fill=MINT, line=None)
    textbox(s, M + 0.28, y + 0.14, 4.2, 0.24, "DESIGNED TO FAIL HONESTLY", size=8.4,
            color=EMERALD_DK, bold=True, spacing=1.3)
    items = [("RQ1 could fail", "if one illegal transition were accepted, or a legal one wrote no audit row."),
             ("RQ2 could fail", "if any settlement differed from its baseline, or a clock did not fire."),
             ("RQ3 could fail", "if the score did not separate completing members from defaulting ones.")]
    for i, (a, b) in enumerate(items):
        x = M + 0.28 + i * 3.86
        textbox(s, x, y + 0.44, 3.6, 0.24, a, size=9.4, color=NAVY, bold=True)
        textbox(s, x, y + 0.70, 3.6, 0.5, b, size=8.6, color=MUTED, line=1.18)
    return s


# ------------------------------------------------------------ 25 results RQ1
def results_rq1(prs):
    s = content_slide(prs, "Results  \u00b7  RQ1", "The lifecycle held: zero illegal transitions")
    card(s, M, 1.52, 6.35, 1.06, fill=NAVY, line=None)
    textbox(s, M + 0.28, 1.64, 5.8, 0.24, "RESULT", size=8.2, color=EMERALD_BR, bold=True,
            spacing=1.4)
    textbox(s, M + 0.28, 1.90, 5.85, 0.56,
            "Across the automated suites, not one illegal transition was accepted, and every "
            "permitted one wrote exactly one audit record.",
            size=11.6, color=WHITE, line=1.22)

    bullets(s, M, 2.80, 6.35, 3.0, [
        "Every legal edge in the transition table is exercised; every illegal jump is attempted "
        "deliberately \u2014 pickup without payment, review before completion, a rewind on a "
        "completed booking.",
        "**Illegal attempt:** the conditional update matches zero rows, the transaction aborts and "
        "the client receives a state conflict \u2014 not a validation message.",
        "**Concurrency:** a double-clicked approval or a parallel payment submission leaves exactly "
        "one winner; the loser is refused by the database, not by a re-check.",
        "**Audit:** status, actor and timestamp are written in the same commit as the change, so a "
        "transition without a record is not representable.",
    ], size=9.8, gap=8, line=1.2)

    card(s, M, 4.90, 6.35, 1.16, fill=MINT, line=None)
    textbox(s, M + 0.24, 5.02, 5.7, 0.24, "THE SAME TABLE DRIVES THE INTERFACE", size=8.2,
            color=EMERALD_DK, bold=True, spacing=1.3)
    textbox(s, M + 0.24, 5.28, 5.85, 0.68,
            "Action buttons are rendered from the transition table, so an impossible action is never "
            "offered. The guard still re-checks server-side: a button is not a permission.",
            size=8.8, color=NAVY, line=1.2)

    # attempt table
    x2 = M + 6.62
    w2 = CW - 6.62
    textbox(s, x2, 1.52, w2, 0.24, "ILLEGAL ATTEMPTS AND THEIR OUTCOME", size=8.4, color=MUTED,
            bold=True, spacing=1.3)
    rows = [
        ["Attempted transition", "Outcome"],
        ["PENDING \u2192 ACTIVE (skip approval)", "Rejected"],
        ["PAYMENT_PENDING \u2192 ACTIVE (skip payment)", "Rejected"],
        ["COMPLETED \u2192 ACTIVE (rewind)", "Rejected"],
        ["APPROVED \u2192 PAID (skip payment record)", "Rejected"],
        ["PAID \u2192 ACTIVE (legal)", "Accepted \u00b7 1 audit row"],
    ]
    table(s, x2, 1.80, w2, rows, widths=[3.02, 2.13], row_h=0.40, header_h=0.36,
          size=8.6, hsize=8.4, zebra=PAPER,
          colors={(5, 1): EMERALD_DK, (1, 1): RED, (2, 1): RED, (3, 1): RED, (4, 1): RED})
    textbox(s, x2, 4.20, w2, 0.5,
            "Each row is a scenario in the suite, not an illustration: the table is the test plan.",
            size=8.4, color=MUTED, italic=True, line=1.18)

    card(s, x2, 4.72, w2, 1.62, fill=PAPER, line=LINE)
    textbox(s, x2 + 0.24, 4.84, w2 - 0.44, 0.24, "WHAT THIS DOES NOT CLAIM", size=8.2,
            color=AMBER, bold=True, spacing=1.3)
    textbox(s, x2 + 0.24, 5.10, w2 - 0.44, 1.14,
            "Testing shows the guard rejects what the suite attempted \u2014 not a formal proof of the "
            "state machine. Two write paths remain outside it and are covered by review: the "
            "expiry worker and administrative suspension.",
            size=8.8, color=MUTED, line=1.2)
    return s


# ------------------------------------------------------------ 26 results RQ2
def results_rq2(prs):
    s = content_slide(prs, "Results  \u00b7  RQ2", "Six fault scenarios, six exact settlements")
    scenarios = [
        ("a", "Clean return", "No claim filed. Deposit released in full; commission taken from the "
                              "rental fee only."),
        ("b", "Late return", "Agreed hourly late fee deducted from the deposit before release."),
        ("c", "Minor cosmetic damage", "Claimed amount settled by mutual agreement, balance refunded."),
        ("d", "Severe functional damage", "Deposit forfeited up to the assessed damage; remainder refunded."),
        ("e", "Item not returned", "Full deposit forfeited; rental share released to the owner."),
        ("f", "Contested claim, no response", "Auto-escalated to an administrator; decided on the paired "
                                               "handover photographs."),
    ]
    w, h = 3.72, 1.16
    for i, (tag, head, body) in enumerate(scenarios):
        x = M + (i % 3) * (w + 0.31)
        y = 1.52 + (i // 3) * (h + 0.22)
        card(s, x, y, w, h, fill=WHITE if i < 3 else PAPER, line=LINE)
        pill(s, x + 0.18, y + 0.16, 0.30, 0.26, tag, fill=MINT, color=EMERALD_DK, size=9)
        textbox(s, x + 0.56, y + 0.15, w - 0.72, 0.26, head, size=10.2, color=NAVY, bold=True)
        textbox(s, x + 0.18, y + 0.48, w - 0.36, 0.62, body, size=8.8, color=MUTED, line=1.18)

    y = 4.28
    card(s, M, y, CW, 1.02, fill=NAVY, line=None)
    textbox(s, M + 0.28, y + 0.14, 5.6, 0.24, "RESULT", size=8.2, color=EMERALD_BR, bold=True,
            spacing=1.4)
    textbox(s, M + 0.28, y + 0.40, CW - 0.6, 0.52,
            "Each computed settlement matched its independently hand-calculated baseline exactly, the "
            "conservation identities held in every case, and no balance outlived its clock.",
            size=11.4, color=WHITE, line=1.22)

    y2 = 5.52
    thirds = [("Conservation held", "payout + commission + refund = collected, in all six"),
              ("Timers fired", "48-hour refund, 7-day claim, 7-day response, 14-day review"),
              ("Evidence decided", "claims adjudicated on paired pickup and return photographs")]
    for i, (a, b) in enumerate(thirds):
        x = M + i * 3.96
        rect(s, x, y2 + 0.04, 0.05, 0.62, fill=EMERALD, line=None, shape=MSO_SHAPE.RECTANGLE)
        textbox(s, x + 0.16, y2, 3.6, 0.24, a, size=9.6, color=NAVY, bold=True)
        textbox(s, x + 0.16, y2 + 0.26, 3.6, 0.46, b, size=8.4, color=MUTED, line=1.16)
    textbox(s, M, 6.44, CW, 0.26,
            "Scenario (f) is the one that proves the clock: silence escalates the claim instead of "
            "quietly releasing the deposit.", size=8.8, color=MUTED, italic=True)
    return s


# ------------------------------------------------------------ 27 results RQ3
def results_rq3(prs):
    import math

    s = content_slide(prs, "Results  \u00b7  RQ3", "The score separates, and the ladder enforces")
    card(s, M, 1.52, 6.15, 3.66, fill=PAPER, line=LINE)
    textbox(s, M + 0.24, 1.64, 5.6, 0.24,
            "SCORE DISTRIBUTION BY BEHAVIOUR CLASS  \u00b7  SCHEMATIC",
            size=8, color=MUTED, bold=True, spacing=1.2)

    axis_y, axis_x0, span = 4.04, M + 0.46, 5.30

    # band shading + labels
    b1 = axis_x0 + 0.60 * span
    b2 = axis_x0 + 0.80 * span
    rect(s, b1, 1.96, b2 - b1, axis_y - 1.96, fill=MINT, line=None, shape=MSO_SHAPE.RECTANGLE)
    rect(s, b2, 1.96, axis_x0 + span - b2, axis_y - 1.96, fill=SLATE, line=None,
         shape=MSO_SHAPE.RECTANGLE)
    textbox(s, b1 + 0.06, 2.00, b2 - b1 - 0.12, 0.22, "TRUSTED", size=6.8, color=EMERALD_DK,
            bold=True, align=PP_ALIGN.CENTER, spacing=1.0)
    textbox(s, b2 + 0.06, 2.00, axis_x0 + span - b2 - 0.12, 0.22, "HIGHLY TRUSTED", size=6.8,
            color=NAVY_SOFT, bold=True, align=PP_ALIGN.CENTER, spacing=1.0)

    # distributions
    for base, colour in ((0.32, AMBER_BR), (0.72, EMERALD)):
        n = 30
        for i in range(n):
            centre = base + (i - n / 2) * 0.009
            height = 1.42 * math.exp(-((centre - base) ** 2) / (2 * 0.048 ** 2))
            if height < 0.05:
                continue
            x = axis_x0 + centre * span - 0.045
            rect(s, x, axis_y - height, 0.09, height, fill=colour, line=None,
                 shape=MSO_SHAPE.RECTANGLE)

    # axis
    hline(s, axis_x0, axis_y, span, color=LINE, lw=1.2)
    for frac in (0.0, 0.25, 0.50, 0.75, 1.0):
        x = axis_x0 + frac * span
        textbox(s, x - 0.3, axis_y + 0.04, 0.6, 0.20, f"{frac:.2f}", size=7, color=MUTED,
                align=PP_ALIGN.CENTER)

    # threshold marks
    for val, lab in ((0.60, "0.60"), (0.80, "0.80")):
        x = axis_x0 + val * span
        rect(s, x - 0.008, 2.28, 0.016, axis_y - 2.28, fill=NAVY_SOFT, line=None,
             shape=MSO_SHAPE.RECTANGLE)
        textbox(s, x - 0.46, 2.30, 0.36, 0.18, lab, size=6.8, color=NAVY, bold=True,
                align=PP_ALIGN.RIGHT)

    # series captions under the axis
    textbox(s, axis_x0 + 0.32 * span - 0.85, axis_y + 0.26, 1.7, 0.20,
            "defaulting /", size=7.2, color=AMBER_BR, bold=True, align=PP_ALIGN.CENTER)
    textbox(s, axis_x0 + 0.32 * span - 0.85, axis_y + 0.44, 1.7, 0.20,
            "self-cancelling members", size=7.2, color=AMBER_BR, bold=True, align=PP_ALIGN.CENTER)
    textbox(s, axis_x0 + 0.72 * span - 0.85, axis_y + 0.26, 1.7, 0.20,
            "members completing", size=7.2, color=EMERALD_DK, bold=True, align=PP_ALIGN.CENTER)
    textbox(s, axis_x0 + 0.72 * span - 0.85, axis_y + 0.44, 1.7, 0.20,
            "rentals", size=7.2, color=EMERALD_DK, bold=True, align=PP_ALIGN.CENTER)

    textbox(s, M + 0.24, 4.80, 5.65, 0.22,
            "Modelled over 2,000 synthetic members with stated behaviour distributions; shapes are "
            "illustrative, the ordering is the result.", size=7.6, color=MUTED, italic=True,
            line=1.15)

    # right column
    x2 = M + 6.42
    w2 = CW - 6.42
    card(s, x2, 1.52, w2, 1.30, fill=NAVY, line=None)
    textbox(s, x2 + 0.24, 1.64, w2 - 0.45, 0.22, "RESULT", size=8, color=EMERALD_BR, bold=True,
            spacing=1.4)
    textbox(s, x2 + 0.24, 1.90, w2 - 0.45, 0.82,
            "Completing members outscored defaulting ones across the modelled population, and the "
            "bands landed where the formula intends.",
            size=9.8, color=WHITE, line=1.2)

    bullets(s, x2, 3.00, w2, 1.85, [
        "**Worked case:** 0.787 for an established member versus 0.574 for a single flawless rental "
        "\u2014 volume and verification decide it.",
        "**Access resolved:** the higher score clears the PKR 100,000 tier; the lower one is held at "
        "the entry tier until identity or history catches up.",
    ], size=9.2, gap=7, line=1.2)

    card(s, x2, 5.02, w2, 1.34, fill=PAPER, line=LINE)
    textbox(s, x2 + 0.24, 5.14, w2 - 0.45, 0.22, "WHAT THIS DOES NOT CLAIM", size=8,
            color=AMBER, bold=True, spacing=1.3)
    textbox(s, x2 + 0.24, 5.38, w2 - 0.45, 0.86,
            "The synthetic population encodes our assumptions about behaviour, and the staging pilot "
            "is a case study, not a sample. Real-volume calibration is FYDP-II work.",
            size=8.6, color=MUTED, line=1.2)
    return s


# ------------------------------------------------------------ 28 limitations
def limitations(prs):
    s = content_slide(prs, "Chapter 3  \u00b7  Threats to validity", "Limitations, and what we did about them")
    rows = [
        ["Limitation", "Risk if unmanaged", "Mitigation shipped"],
        ["Concurrent writes on one booking", "Two clicks, two payments, one item",
         "Atomic conditional UPDATE carrying the expected status in its WHERE clause"],
        ["Background jobs can lag", "A deposit released late, or a claim never escalated",
         "Dual trigger: the deadline is also checked on the read path, not only by cron"],
        ["Pilot volume is small", "RQ3 cannot be statistically generalised",
         "2,000 synthetic profiles for sensitivity; pilot reported as an empirical case study"],
        ["CNIC reluctance", "Identity component sits at zero; top tier unreachable",
         "Progressive gating; three completed rentals accepted in place of a document"],
        ["Manual payment verification", "An administrator becomes the bottleneck at scale",
         "Provider-agnostic interface ready for JazzCash, Easypaisa and Safepay in FYDP-II"],
        ["Wear versus damage is subjective", "Disputes turn on opinion rather than evidence",
         "Six-point photographic log with unique checksums at both ends, admin adjudication"],
        ["Photographing items raises privacy concerns", "Members decline to document handovers",
         "Item-only framing, consent at upload, retention limits"],
    ]
    table(s, M, 1.52, CW, rows, widths=[2.75, 3.90, 5.12], row_h=0.62, header_h=0.38,
          size=9.2, hsize=9, zebra=PAPER, bold_first=True, valign=MSO_ANCHOR.MIDDLE)
    textbox(s, M, 6.48, CW, 0.26,
            "Table 3.5 \u2014 every limitation is stated as a design constraint with a shipped mitigation, "
            "not as a footnote.", size=9, color=MUTED, italic=True)
    return s


# ------------------------------------------------------------ 29 contributions
def contributions(prs):
    s = content_slide(prs, "Discussion  \u00b7  Contribution", "What is new here, and why it matters")
    items = [
        ("01", "A lifecycle with one door",
         "Status moves only through a server action that consults a pure transition table and an "
         "atomic conditional write. Skipping a stage is not forbidden \u2014 it is unreachable.", "RQ1"),
        ("02", "Evidence bound to money",
         "Handover records are sealed and checksummed at write time, and the settlement routine "
         "derives every figure from them. Proof and payment are the same transaction history.", "RQ2"),
        ("03", "Trust that gates exposure",
         "A four-factor score with Bayesian shrinkage is wired to deposit tiers, so the platform's "
         "risk appetite is a function of a member's record, not an administrator's judgement.", "RQ3"),
        ("04", "Localisation taken seriously",
         "Integer PKR, cash-on-delivery reality, administrator-verified escrow and CNIC privacy are "
         "design inputs \u2014 not adaptations bolted on after the fact.", "Context"),
    ]
    y = 1.52
    for code, head, body, tag in items:
        card(s, M, y, 8.30, 1.10, fill=WHITE, line=LINE)
        textbox(s, M + 0.22, y + 0.18, 0.7, 0.34, code, size=15, color=EMERALD, bold=True)
        textbox(s, M + 0.92, y + 0.14, 7.0, 0.26, head, size=11.4, color=NAVY, bold=True)
        textbox(s, M + 0.92, y + 0.44, 7.2, 0.58, body, size=8.8, color=MUTED, line=1.18)
        pill(s, M + 7.62, y + 0.16, 0.55, 0.24, tag, fill=MINT, color=EMERALD_DK, size=7.6)
        y += 1.20

    x2 = M + 8.62
    w2 = CW - 8.62
    card(s, x2, 1.52, w2, 2.62, fill=NAVY, line=None)
    textbox(s, x2 + 0.24, 1.64, w2 - 0.45, 0.22, "THE ONE-SENTENCE VERSION", size=8.2,
            color=EMERALD_BR, bold=True, spacing=1.3)
    textbox(s, x2 + 0.24, 1.94, w2 - 0.45, 2.06,
            "Each of these exists in the literature. The contribution is the coupling: evidence "
            "sealed at handover drives settlement arithmetic, and settlement history drives the "
            "trust that decides who may hold a PKR 100,000 item.",
            size=9.8, color=WHITE, line=1.24)

    card(s, x2, 4.30, w2, 2.02, fill=MINT, line=None)
    textbox(s, x2 + 0.24, 4.42, w2 - 0.45, 0.22, "VALIDATED AGAINST", size=8.2, color=EMERALD_DK,
            bold=True, spacing=1.3)
    bullets(s, x2 + 0.24, 4.70, w2 - 0.46, 1.5, [
        "RQ1 \u2014 integration suite",
        "RQ2 \u2014 six hand-calculated scenarios",
        "RQ3 \u2014 2,000 modelled profiles",
        "Staging deployment, end to end",
    ], size=8.8, gap=4, marker="\u2713")
    return s


# ------------------------------------------------------------ 30 reflections
def reflections(prs):
    s = content_slide(prs, "Group reflections", "How the work was divided, and what it cost us")
    members = [
        ("Ali Razi", "23F-BSCS-53", "Team lead", "Next.js full-stack, server actions, integration"),
        ("Muhammad Saad", "23F-BSCS-63", "Full-stack", "PostgreSQL schema, Prisma, admin surfaces"),
        ("M. Rehan Amjad", "23F-BSCS-65", "Domain", "State-machine guards, deposit timers"),
        ("Sameer Raza", "23F-BSCS-69", "Integrity", "Checksum pipeline, evidence storage"),
        ("John Kazmi", "23F-BSCS-99", "Analysis", "Trust engine maths, literature, test cases"),
    ]
    w = 2.21
    for i, (name, roll, tag, area) in enumerate(members):
        x = M + i * (w + 0.16)
        card(s, x, 1.52, w, 1.62, fill=PAPER if i else MINT, line=LINE)
        textbox(s, x + 0.18, 1.64, w - 0.36, 0.26, name, size=10.2, color=NAVY, bold=True)
        textbox(s, x + 0.18, 1.92, w - 0.36, 0.22, roll, size=8.2, color=MUTED)
        textbox(s, x + 0.18, 2.18, w - 0.36, 0.22, tag.upper(), size=7.4, color=EMERALD_DK,
                bold=True, spacing=1.1)
        textbox(s, x + 0.18, 2.44, w - 0.36, 0.62, area, size=8.2, color=MUTED, line=1.16)

    lessons = [
        ("Atomicity was the real adversary",
         "Early builds let a concurrent click decline a booking as its renter submitted payment. The fix "
         "was not a lock in application code but moving the expected state into the WHERE clause, so the "
         "database arbitrates the race."),
        ("We changed the method rather than the data",
         "RQ3 cannot be answered from a few dozen pilot rentals. Instead of inventing history, we "
         "measured the formula's behaviour over 2,000 modelled profiles and reported the pilot as a case "
         "study \u2014 stating the limit rather than hiding it."),
        ("We documented what we could not automate",
         "Pakistani banking offers no programmatic escrow to a student project. We built an "
         "administrator-verified escrow workflow with a full audit trail and called it exactly that."),
    ]
    y = 3.32
    for i, (head, body) in enumerate(lessons):
        card(s, M, y, CW, 1.00, fill=WHITE if i % 2 == 0 else PAPER, line=LINE)
        circle_num(s, M + 0.24, y + 0.10, 0.36, str(i + 1), fill=MINT, color=EMERALD_DK, size=10.5)
        textbox(s, M + 0.78, y + 0.15, 4.0, 0.26, head, size=10.6, color=NAVY, bold=True)
        textbox(s, M + 4.86, y + 0.15, CW - 5.15, 0.76, body, size=8.8, color=MUTED, line=1.18)
        y += 1.18
    return s


# ------------------------------------------------------------ 31 roadmap
def roadmap(prs):
    s = content_slide(prs, "FYDP-II", "What we take into the second phase")
    phases = [
        ("Phase II-a", "Decouple the backend", MINT, EMERALD_DK, [
            "Move domain services out of Next.js into a standalone Node.js / Express API",
            "One business-logic source for every client; REST endpoints behind the same guards",
            "Web client keeps its server components, now calling the API",
        ]),
        ("Phase II-b", "Ship the mobile client", PAPER, NAVY, [
            "React Native application against the shared API",
            "Camera-first handover flow, which the evidence protocol already assumes",
            "Push notifications alongside the existing realtime channel",
        ]),
        ("Phase II-c", "Close the payment gap", NAVY, WHITE, [
            "Live JazzCash, Easypaisa and Safepay integrations behind the existing abstraction",
            "Automated verification replaces the administrator queue",
            "Wallet-native reconciliation for deposit release and payout",
        ]),
    ]
    w = 3.80
    for i, (tag, head, fill, col, items) in enumerate(phases):
        x = M + i * (w + 0.19)
        card(s, x, 1.52, w, 2.62, fill=fill, line=None if fill != PAPER else LINE)
        sub = MINT_TXT if fill == NAVY else EMERALD_DK
        textbox(s, x + 0.22, 1.64, w - 0.4, 0.22, tag.upper(), size=7.8, color=sub, bold=True,
                spacing=1.3)
        textbox(s, x + 0.22, 1.90, w - 0.4, 0.52, head, size=12, color=col, bold=True, line=1.1)
        bullets(s, x + 0.22, 2.50, w - 0.44, 1.5, items, size=8.6, gap=6, line=1.18,
                color=col if fill == NAVY else MUTED, bullet_color=sub)

    # closing targets
    y = 4.30
    card(s, M, y, CW, 1.10, fill=WHITE, line=LINE)
    textbox(s, M + 0.26, y + 0.14, 4.0, 0.24, "PILOT TARGETS CARRIED FORWARD", size=8.2,
            color=MUTED, bold=True, spacing=1.3)
    targets = [("60\u2013100", "active listings"), ("30\u201350", "completed rentals"),
               ("60\u2013100", "sealed handover logs"), ("30\u201380", "member reviews")]
    for i, (big, sub) in enumerate(targets):
        x = M + 0.26 + i * 2.86
        textbox(s, x, y + 0.42, 1.3, 0.40, big, size=16, color=EMERALD, bold=True)
        textbox(s, x + 1.15, y + 0.52, 1.7, 0.26, sub, size=8.6, color=MUTED)

    # risks we carry in
    textbox(s, M, 5.56, 6.0, 0.22, "WHAT WE STILL OWE, AND HOW WE WILL SHOW IT", size=8.2,
            color=MUTED, bold=True, spacing=1.3)
    risks = [("Payments stay manual", "measured admin queue time per verification"),
             ("Identity data stays sensitive", "published retention policy before the pilot"),
             ("Trust weights are uncalibrated", "re-fit against real settlement history")]
    for i, (a, b) in enumerate(risks):
        x = M + i * 3.96
        textbox(s, x, 5.82, 3.7, 0.24, a, size=9.2, color=NAVY, bold=True)
        textbox(s, x, 6.06, 3.7, 0.42, b, size=8.2, color=MUTED, line=1.16)
    return s


# ------------------------------------------------------------ 32 conclusion
def conclusion(prs):
    s = content_slide(prs, "Conclusion", "Three questions, three answers")
    answers = [
        ("RQ1", "Yes \u2014 the lifecycle holds.",
         "A pure transition table plus atomic conditional writes made state-skipping unreachable in "
         "testing, with one audit record per permitted change.", "O1 delivered"),
        ("RQ2", "Yes \u2014 the deposit protects both sides.",
         "Sealed two-sided evidence with derived settlement arithmetic produced exact figures in six "
         "fault scenarios and released every balance on its clock.", "O2 delivered"),
        ("RQ3", "Yes, within stated limits \u2014 the score separates.",
         "A four-factor score with shrinkage and identity verification ordered a 2,000-member "
         "population as intended; the pilot remains a case study.", "O3 delivered"),
    ]
    y = 1.52
    for code, claim, body, tag in answers:
        card(s, M, y, 8.55, 1.08, fill=WHITE, line=LINE)
        rect(s, M, y, 0.05, 1.08, fill=EMERALD, line=None, shape=MSO_SHAPE.RECTANGLE)
        textbox(s, M + 0.24, y + 0.18, 0.72, 0.3, code, size=12, color=EMERALD, bold=True)
        textbox(s, M + 0.94, y + 0.12, 5.2, 0.28, claim, size=10.8, color=NAVY, bold=True)
        textbox(s, M + 0.94, y + 0.40, 7.25, 0.54, body, size=8.8, color=MUTED, line=1.18)
        pill(s, M + 7.42, y + 0.14, 1.02, 0.26, tag, fill=MINT, color=EMERALD_DK, size=7.8)
        y += 1.18

    x2 = M + 8.85
    w2 = CW - 8.85
    card(s, x2, 1.52, w2, 3.30, fill=NAVY, line=None)
    textbox(s, x2 + 0.24, 1.64, w2 - 0.45, 0.22, "THE CLOSING STATEMENT", size=8.2,
            color=EMERALD_BR, bold=True, spacing=1.3)
    textbox(s, x2 + 0.24, 1.94, w2 - 0.45, 2.72,
            "SamaanShare shows that the trust deficit in peer-to-peer rental is an engineering "
            "problem before it is a social one. A guarded lifecycle, evidence sealed at the moment "
            "it matters, settlement arithmetic that cannot be argued with, and a computed score that "
            "governs access \u2014 those four decisions are what make a stranger's DSLR rentable.",
            size=10, color=WHITE, line=1.3)

    textbox(s, M, 5.34, 8.55, 0.26, "ACKNOWLEDGEMENT", size=8.2, color=MUTED, bold=True,
            spacing=1.3)
    textbox(s, M, 5.60, 8.55, 0.80,
            "We thank our supervisor Dr. Sarwat Iqbal and co-supervisor Dr. Asma Larik for their "
            "guidance, FYDP coordinators Engr. Imran Alvi and Engr. Ponam K.K. for the process that "
            "kept this on schedule, and Chairperson Dr. Fida Hussain Khoso for the department's support.",
            size=9, color=MUTED, line=1.24)
    hline(s, M, 6.50, 8.55, color=LINE, lw=0.75)
    textbox(s, M, 6.58, 8.55, 0.26,
            "Handed over today: the staging deployment, the repository, and the FYDP-I report.",
            size=8.8, color=EMERALD_DK, bold=True)
    return s


# ------------------------------------------------------------ 33 thank you
def thanks(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s, NAVY)
    pg()
    ring = rect(s, -2.2, 3.4, 6.6, 6.6, fill=NAVY_SOFT, line=None, shape=MSO_SHAPE.OVAL)
    set_alpha(ring, 55)
    ring2 = rect(s, 9.9, -2.0, 5.6, 5.6, fill=EMERALD_DK, line=None, shape=MSO_SHAPE.OVAL)
    set_alpha(ring2, 75)
    rect(s, 0, 0, 0.14, SH, fill=EMERALD, line=None, shape=MSO_SHAPE.RECTANGLE)

    picture(s, A + "logo-navy.png", M - 0.02, 0.62, w=0.80)
    textbox(s, M + 0.94, 0.78, 4.0, 0.3, "SamaanShare", size=17, color=WHITE, bold=True)

    textbox(s, M, 2.32, 8.4, 0.9, "Thank you", size=46, color=WHITE, bold=True)
    textbox(s, M, 3.12, 7.8, 0.8,
            "Questions, challenges and the live staging walkthrough.",
            size=17, color=TITLE_SUB)
    hline(s, M, 3.94, 1.4, color=EMERALD_BR, lw=2.0)

    textbox(s, M, 4.20, 8.0, 0.24, "GROUP 23FYDP24", size=9, color=EMERALD_BR, bold=True,
            spacing=1.5)
    textbox(s, M, 4.48, 8.4, 0.7,
            "Ali Razi  \u00b7  Muhammad Saad  \u00b7  M. Rehan Amjad  \u00b7  Sameer Raza  \u00b7  John Kazmi",
            size=11.5, color=WHITE, line=1.3)
    textbox(s, M, 5.02, 8.4, 0.4,
            "Department of Computer Science, Dawood University of Engineering and Technology, Karachi",
            size=9.6, color=TITLE_SUB)
    textbox(s, M, 6.34, 8.6, 0.3,
            "AI-assisted drafting is disclosed in the thesis appendix, per DUET guidelines.",
            size=8.4, color=MINT_TXT)
    return s


# ------------------------------------------------------------ 34 references
def references(prs):
    s = content_slide(prs, "References", "Selected sources (full list in the thesis, [1]\u2013[25])")
    left = [
        ("[1]", "Filippas, Horton & Zeckhauser. Owning, using and renting: some simple economics of "
                "the sharing economy. Management Science 66(9), 2020."),
        ("[3]", "Gerwe & Silva. Clarifying the sharing economy. Academy of Management Perspectives "
                "34(1), 2020."),
        ("[4]", "Anjum & Chai. Drivers of cash-on-delivery in e-commerce: evidence from Pakistan. "
                "SAGE Open 10(3), 2020."),
        ("[6]", "Ghozi. The urgency of electronic KYC in the fintech industry. Diponegoro Law Review "
                "7(1), 2022."),
        ("[8]", "Kotiyal et al. An innovative collaborative resource sharing platform. CCIS 1988, "
                "Springer, 2026."),
        ("[9]", "Zhao & Chen. Digital governance of dispute resolution and prevention in online "
                "commerce. JIDS 15(1), 2024."),
    ]
    right = [
        ("[10]", "Tele et al. VahanRent: a blockchain powered peer-to-peer car rental system. "
                 "IJARSCT 6(3), 2026."),
        ("[12]", "Salgado, Filieri & Chameroy. Beyond unidimensional trust: a multidimensional, "
                 "role-based approach. Journal of Travel Research 65(3), 2026."),
        ("[14]", "Neifer et al. Trust-building in P2P carsharing: algorithm-based reputation systems. "
                 "CSCW 32(4), 2023."),
        ("[21]", "Rana, Zhang & Hamid. Adoption ease of mobile payment models and app design. "
                 "IEEE ICEET, 2022."),
        ("[25]", "Papanastasiou, Yang & Zhang. Improving dispute resolution in two-sided platforms: "
                 "the case of review blackmail. Management Science 69(10), 2023."),
    ]
    for col, refs, x in ((0, left, M), (1, right, M + 6.05)):
        y = 1.56
        for tag, ref in refs:
            textbox(s, x, y, 0.52, 0.24, tag, size=8.6, color=EMERALD, bold=True)
            textbox(s, x + 0.50, y - 0.03, 5.05, 0.70, ref, size=8.6, color=MUTED, line=1.2)
            y += 0.64

    card(s, M, 5.58, CW, 0.62, fill=PAPER, line=LINE)
    textbox(s, M + 0.24, 5.68, CW - 0.5, 0.42,
            "Sources are cited in-text by number. Product claims in this deck were verified against the "
            "staging deployment and the project repository, not only the manuscript.",
            size=8.6, color=MUTED, italic=True, line=1.16)
    return s
