"""Slides 13-23: journey, the three core techniques, traceability, implementation, walkthrough."""

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


# ------------------------------------------------------------ 13 journey
def journey(prs):
    s = content_slide(prs, "Chapter 3  \u00b7  Operational lifecycle",
                      "The ten stages of one rental, end to end")
    steps = [
        ("1", "Listing published", "Taxonomy, city, rates, deposit, up to 10 photos", "PUBLISHED"),
        ("2", "Booking requested", "Dates chosen; a hold is placed in the same transaction", "PENDING"),
        ("3", "Owner approves", "Prices frozen onto the booking", "APPROVED"),
        ("4", "Payment submitted", "Rent + deposit sent; reference and receipt uploaded", "PAYMENT_PENDING"),
        ("5", "Admin verifies", "Commission rate frozen at verification", "PAID"),
        ("6", "Pickup sealed", "Condition, notes, up to 6 photos", "ACTIVE"),
        ("7", "Return sealed", "Matching photographic standard applied", "COMPLETED"),
        ("8", "Claim window", "Owner files within 7 days; renter replies in 7", "CLAIMED"),
        ("9", "Admin settlement", "Figures derived, payout and refund recorded", "SETTLED"),
        ("10", "Reciprocal reviews", "Sealed until both submit or 14 days pass", "REVIEWED"),
    ]
    w, h, gap = 2.21, 1.30, 0.16
    for i, (num, head, body, state) in enumerate(steps):
        col, row = i % 5, i // 5
        x = M + col * (w + gap)
        y = 1.52 + row * (h + 0.30)
        card(s, x, y, w, h, fill=WHITE if row == 0 else PAPER, line=LINE)
        circle_num(s, x + 0.14, y + 0.13, 0.30, num, fill=MINT, color=EMERALD_DK, size=9.5)
        textbox(s, x + 0.50, y + 0.13, w - 0.62, 0.26, head, size=9.6, color=NAVY, bold=True)
        textbox(s, x + 0.14, y + 0.48, w - 0.28, 0.52, body, size=8.2, color=MUTED, line=1.15)
        textbox(s, x + 0.14, y + h - 0.26, w - 0.28, 0.20, state, size=7.2,
                color=EMERALD, bold=True)
        if col < 4:
            arrow_right(s, x + w + 0.02, y + 0.56, 0.12, 0.14, color=LINE)

    # clocks band
    y = 4.86
    card(s, M, y, CW, 1.34, fill=NAVY, line=None)
    textbox(s, M + 0.28, y + 0.14, 4.0, 0.24, "TIME-BOUNDED BY DESIGN \u2014 NO OPEN-ENDED HOLD",
            size=8.4, color=EMERALD_BR, bold=True, spacing=1.3)
    clocks = [("48 h", "deposit refund clock from a sealed return"),
              ("7 days", "owner's claim window after completion"),
              ("7 days", "renter's response window on a claim"),
              ("14 days", "reciprocal review release")]
    for i, (big, sub) in enumerate(clocks):
        x = M + 0.28 + i * 2.86
        textbox(s, x, y + 0.44, 1.5, 0.40, big, size=17, color=WHITE, bold=True)
        textbox(s, x, y + 0.86, 2.6, 0.36, sub, size=8.4, color=MINT_TXT, line=1.14)

    textbox(s, M, 6.34, CW, 0.3,
            "Steps 8 and 9 are conditional: a clean return jumps straight from the return log to "
            "settlement and release.", size=9, color=MUTED, italic=True)
    return s


# ------------------------------------------------------------ 14 state machine
def state_machine(prs):
    s = content_slide(prs, "Core technique 01  \u00b7  RQ1",
                      "A lifecycle that cannot skip a step")
    flow = ["PENDING", "APPROVED", "PAYMENT_PENDING", "PAID", "ACTIVE", "COMPLETED", "REVIEWED"]
    w, gap = 1.42, 0.30
    for i, st in enumerate(flow):
        x = M + i * (w + gap)
        pill(s, x, 1.58, w, 0.38, st, fill=MINT, color=EMERALD_DK, size=9.2, spacing=0.2)
        if i < len(flow) - 1:
            arrow_right(s, x + w + 0.04, 1.70, 0.22, 0.14, color=EMERALD_BR)

    textbox(s, M, 2.18, 4.9, 0.20, "TERMINAL EXITS \u2014 FROM ANY OPEN STATE",
            size=7.8, color=MUTED, bold=True, spacing=1.0)
    exits = [("DECLINED", "the owner refuses the request"),
             ("CANCELLED", "either side withdraws before pickup"),
             ("EXPIRED", "a worker closes an abandoned booking")]
    for i, (st, why) in enumerate(exits):
        x = M + i * 2.73
        card(s, x, 2.42, 2.55, 0.62, fill=PAPER, line=LINE)
        textbox(s, x + 0.16, 2.49, 1.6, 0.22, st, size=9, color=NAVY, bold=True)
        textbox(s, x + 0.16, 2.71, 2.24, 0.26, why, size=7.8, color=MUTED, line=1.1)
    textbox(s, M + 8.32, 2.50, CW - 8.32, 0.46,
            "No transition out of a terminal state exists \u2014 a declined booking cannot be revived.",
            size=8.2, color=MUTED, italic=True, line=1.18)

    # code card
    card(s, M, 3.28, 4.55, 1.66, fill=NAVY, line=None)
    textbox(s, M + 0.20, 3.38, 4.1, 0.22, "THE UPDATE THAT CANNOT RACE", size=8.2,
            color=EMERALD_BR, bold=True, spacing=1.3)
    code = [
        ("UPDATE bookings", WHITE),
        ("   SET status = 'ACTIVE'", WHITE),
        (" WHERE id    = :id", WHITE),
        ("   AND status = 'PAID';", WHITE),
        ("-- 0 rows matched: another", MINT_TXT),
        ("-- request already moved it.", MINT_TXT),
        ("-- Abort; report a conflict.", MINT_TXT),
    ]
    tb = textbox(s, M + 0.20, 3.62, 4.2, 1.26, "", size=8.2, color=WHITE, line=1.16)
    tf = tb.text_frame
    for i, (line_s, colr) in enumerate(code):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.line_spacing = 1.14
        run = p.add_run()
        run.text = line_s
        run.font.size = Pt(8.0)
        run.font.name = "Consolas"
        run.font.color.rgb = colr

    # two explanation cards
    x2 = M + 4.75
    card(s, x2, 3.28, 3.45, 1.66, fill=MINT, line=None)
    textbox(s, x2 + 0.20, 3.38, 3.0, 0.22, "WHY NOT A SMART CONTRACT", size=8.2,
            color=EMERALD_DK, bold=True, spacing=1.2)
    textbox(s, x2 + 0.20, 3.62, 3.05, 1.20,
            "We kept the automated lifecycle of [10] but executed it with relational locks: the same "
            "guarantee without gas, wallet custody or block latency.", size=8.4, color=NAVY, line=1.2)

    x3 = M + 8.40
    w3 = CW - 8.40
    card(s, x3, 3.28, w3, 1.66, fill=WHITE, line=LINE)
    textbox(s, x3 + 0.20, 3.38, w3 - 0.4, 0.22, "CONCURRENCY", size=8.2, color=AMBER, bold=True,
            spacing=1.3)
    textbox(s, x3 + 0.20, 3.62, w3 - 0.40, 1.20,
            "Two clicks, one payment: the condition lives in the WHERE clause, so the database "
            "arbitrates and the loser is refused, not re-checked.", size=8.4, color=MUTED, line=1.2)

    # evidence band
    y = 5.24
    card(s, M, y, CW, 0.92, fill=PAPER, line=LINE)
    textbox(s, M + 0.26, y + 0.10, 4.6, 0.20, "ONE PURE TRANSITION TABLE, NO I/O \u2014 WHAT IT BOUGHT",
            size=7.8, color=MUTED, bold=True, spacing=1.0)
    items = [("0", "illegal transitions accepted by the suite"),
             ("1", "audit row per permitted transition"),
             ("5", "guards in every server action")]
    for i, (big, sub) in enumerate(items):
        x = M + 0.26 + i * 3.86
        textbox(s, x, y + 0.38, 0.5, 0.40, big, size=18, color=EMERALD, bold=True)
        textbox(s, x + 0.42, y + 0.48, 3.2, 0.36, sub, size=8.4, color=NAVY, line=1.12)
    textbox(s, M, 6.38, CW, 0.24,
            "The table also decides which buttons the interface offers \u2014 the panel never sees an "
            "action the server would refuse.", size=8.8, color=MUTED, italic=True)
    return s


# ------------------------------------------------------------ 15 handover
def handover(prs):
    s = content_slide(prs, "Core technique 02  \u00b7  RQ2",
                      "Evidence that is sealed at the moment it matters")
    cols = [
        ("Pickup log", "Written by the owner as the item leaves", [
            "Condition rating: as expected / minor wear / damaged",
            "Free-text notes on pre-existing marks",
            "Up to six photographs, taken with the camera only",
        ]),
        ("Return log", "Written by the owner as the item comes back", [
            "Same rating scale, same photographic standard",
            "Filed in the same transaction as COMPLETED",
            "Starts the 48-hour deposit refund clock",
        ]),
    ]
    w = 3.62
    for i, (head, sub, items) in enumerate(cols):
        x = M + i * (w + 0.26)
        card(s, x, 1.58, w, 2.28, fill=PAPER, line=LINE)
        textbox(s, x + 0.22, 1.70, w - 0.4, 0.28, head, size=12.5, color=NAVY, bold=True)
        textbox(s, x + 0.22, 1.98, w - 0.4, 0.26, sub, size=8.6, color=EMERALD_DK, italic=True)
        bullets(s, x + 0.22, 2.32, w - 0.44, 1.4, items, size=9.2, gap=5, line=1.16)

    # sealed rules
    x3 = M + 2 * (w + 0.26)
    w3 = CW - 2 * (w + 0.26)
    card(s, x3, 1.58, w3, 2.28, fill=NAVY, line=None)
    textbox(s, x3 + 0.22, 1.70, w3 - 0.4, 0.24, "SEALED, NOT MERELY SAVED", size=8.2,
            color=EMERALD_BR, bold=True, spacing=1.3)
    bullets(s, x3 + 0.22, 2.02, w3 - 0.44, 1.7, [
        "Written in the **same transaction** as the status change",
        "Triggers refuse **update or delete** \u2014 no edit path exists",
        "Checksum read server-side, **unique per table**",
        "A reused pickup photo is refused at write time",
    ], size=8.6, gap=5, color=WHITE, bullet_color=EMERALD_BR, marker="\u2713", line=1.14)

    # integrity note
    card(s, M, 4.02, CW, 0.96, fill=MINT, line=None)
    textbox(s, M + 0.26, 4.12, 3.0, 0.22, "THE DESIGN PRINCIPLE", size=8.2, color=EMERALD_DK,
            bold=True, spacing=1.3)
    textbox(s, M + 0.26, 4.36, CW - 0.55, 0.56,
            "A record its author can revise afterwards is a claim, not evidence \u2014 and the moment "
            "it matters is exactly the moment they would want to revise it. So the platform seals on "
            "submission and never offers a second write.", size=10.2, color=NAVY, line=1.22)

    # clock timeline
    y = 5.10
    textbox(s, M, y, 6.0, 0.24, "FOUR CLOCKS, NO INDEFINITE HOLD", size=8.4, color=MUTED,
            bold=True, spacing=1.3)
    hline(s, M + 0.1, y + 0.86, CW - 1.6, color=LINE, lw=1.4)
    marks = [("Return sealed", "t = 0", EMERALD),
             ("Deposit released", "+48 h if no claim", EMERALD_BR),
             ("Claim filed", "within 7 days", AMBER_BR),
             ("Renter responds", "7 days, else escalate", AMBER_BR),
             ("Reviews open", "14-day window", NAVY_SOFT)]
    span = (CW - 1.6)
    for i, (head, sub, col) in enumerate(marks):
        x = M + 0.1 + i * (span / 4.0)
        rect(s, x - 0.055, y + 0.78, 0.13, 0.13, fill=col, line=None, shape=MSO_SHAPE.OVAL)
        lx = min(max(x - 0.1, M), SW - M - 2.0)
        textbox(s, lx, y + 0.98, 2.0, 0.24, head, size=8.8, color=NAVY, bold=True)
        textbox(s, lx, y + 1.22, 2.0, 0.24, sub, size=8, color=MUTED)
    return s


# ------------------------------------------------------------ 16 escrow + settlement
def escrow(prs):
    s = content_slide(prs, "Core technique 02  \u00b7  RQ2",
                      "Deposit states and settlement arithmetic")
    textbox(s, M, 1.48, 6.2, 0.24, "THE DEPOSIT RUNS ITS OWN STATE MACHINE", size=8.4,
            color=MUTED, bold=True, spacing=1.3)
    row1 = ["PENDING", "HELD", "CLAIMED"]
    row2 = ["RELEASED", "FORFEITED", "SETTLED"]
    for i, st in enumerate(row1):
        x = M + i * 1.92
        pill(s, x, 1.78, 1.66, 0.36, st, fill=MINT, color=EMERALD_DK, size=9.5)
        if i < 2:
            arrow_right(s, x + 1.70, 1.89, 0.16, 0.14, color=EMERALD_BR)
    pill(s, M, 2.66, 1.66, 0.36, "RELEASED", fill=WHITE, color=EMERALD_DK, size=9.5, line=LINE)
    pill(s, M + 1.84, 2.66, 1.66, 0.36, "FORFEITED", fill=WHITE, color=EMERALD_DK, size=9.5,
         line=LINE)
    arrow_right(s, M + 3.62, 2.77, 0.30, 0.14, color=EMERALD_BR)
    pill(s, M + 4.04, 2.66, 1.52, 0.36, "SETTLED", fill=MINT, color=EMERALD_DK, size=9.5)
    textbox(s, M, 2.42, 6.0, 0.22, "A CLAIM RESOLVES ONE WAY OR THE OTHER \u2014 BOTH CLOSE ON SETTLED",
            size=7.6, color=MUTED, bold=True, spacing=0.8)
    textbox(s, M + 4.04, 3.10, 4.0, 0.24, "money leaves the platform here", size=7.8,
            color=MUTED, italic=True)

    # notes
    bullets(s, M, 3.48, 5.9, 1.5, [
        "The rental fee and the deposit are **different pots** \u2014 commission is computed from the "
        "rental amount alone, so it structurally cannot take a cut of a deposit.",
        "Every figure is **derived** by the settlement routine, never accepted from a form.",
        "Payout and refund are recorded as **separate acts**: they go to different people and "
        "nothing makes one wait on the other.",
    ], size=9.4, gap=6, line=1.18)

    # settlement table
    rows = [
        ["Settlement line", "Clean return", "Damage upheld"],
        ["Rent collected", "PKR 4,000", "PKR 4,000"],
        ["Deposit collected", "PKR 30,000", "PKR 30,000"],
        ["Commission applied", "0% (launch rate)", "7.5% (frozen)"],
        ["Platform commission", "PKR 0", "PKR 300"],
        ["Owner rental share", "PKR 4,000", "PKR 3,700"],
        ["Damage compensation", "PKR 0", "PKR 6,000"],
        ["Owner payout", "PKR 4,000", "PKR 9,700"],
        ["Deposit refunded", "PKR 30,000", "PKR 24,000"],
    ]
    table(s, M + 6.12, 1.48, CW - 6.12, rows, widths=[2.05, 1.72, 1.66],
          row_h=0.303, header_h=0.34, size=9.2, hsize=8.8, zebra=PAPER,
          aligns=[PP_ALIGN.LEFT, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT],
          bold_rows=(7, 8))
    textbox(s, M + 6.12, 4.50, CW - 6.12, 0.44,
            "Table 3.2 \u2014 both columns sum to the PKR 34,000 collected. The check identity is "
            "asserted by the code, not by the accountant.",
            size=8.4, color=MUTED, italic=True, line=1.18)

    # conservation band
    y = 5.22
    card(s, M, y, CW, 1.16, fill=PAPER, line=LINE)
    textbox(s, M + 0.28, y + 0.14, 3.4, 0.24, "TWO CONSERVATION RULES, BY CONSTRUCTION", size=8.4,
            color=MUTED, bold=True, spacing=1.3)
    rules = [("commission + owner rental = rental fee", "the platform cannot profit from rent it did not earn"),
             ("damage compensation + deposit refund = deposit", "the deposit cannot grow or evaporate")]
    for i, (a, b) in enumerate(rules):
        x = M + 0.28 + i * 5.62
        textbox(s, x, y + 0.44, 5.3, 0.26, a, size=10.4, color=NAVY, bold=True)
        textbox(s, x, y + 0.72, 5.3, 0.26, b, size=8.8, color=MUTED)
    return s


# ------------------------------------------------------------ 17 trust formula
def trust_formula(prs):
    s = content_slide(prs, "Core technique 03  \u00b7  RQ3",
                      "A trust score that resists gaming")
    card(s, M, 1.50, CW, 0.96, fill=NAVY, line=None)
    textbox(s, M + 0.30, 1.62, 4.0, 0.24, "COMPOSITE SCORE  \u00b7  S \u2208 [0, 1]", size=8.2,
            color=EMERALD_BR, bold=True, spacing=1.4)
    textbox(s, M + 0.30, 1.88, CW - 0.6, 0.46,
            "S = 0.45\u00b7C_rep  +  0.25\u00b7C_exp  +  0.20\u00b7C_rel  +  0.10\u00b7C_ver",
            size=19, color=WHITE, bold=True)

    comps = [
        ("0.45", "Reputation", "C_rep",
         "( \u03a3 r\u1d62 + 5 \u00d7 3.5 ) / (N + 5) / 5",
         "Bayesian shrinkage toward a neutral 3.5 with 5 pseudo-counts: one five-star rating cannot "
         "outrank fifty good ones."),
        ("0.25", "Experience", "C_exp",
         "min( N_completed / 10 , 1 )",
         "Saturates at ten completed rentals. New members default to 1.0 so inexperience is not "
         "punished twice."),
        ("0.20", "Reliability", "C_rel",
         "N_finished / (N_finished + N_self_cancelled)",
         "Counts only cancellations the member chose, and only on rentals that had begun."),
        ("0.10", "Verification", "C_ver",
         "1.0  \u00b7  0.4  \u00b7  0.0",
         "CNIC verified \u00b7 email and phone only \u00b7 neither. Identity is a component, not a gate "
         "for everything."),
    ]
    w = 2.83
    for i, (wt, name, sym, formula, why) in enumerate(comps):
        x = M + i * (w + 0.20)
        card(s, x, 2.66, w, 2.94, fill=WHITE, line=LINE)
        rect(s, x, 2.66, w, 0.05, fill=EMERALD if i == 0 else EMERALD_BR, line=None,
             shape=MSO_SHAPE.RECTANGLE)
        textbox(s, x + 0.20, 2.80, 0.9, 0.34, wt, size=17, color=EMERALD, bold=True)
        textbox(s, x + 0.20, 3.18, w - 0.4, 0.24, name.upper(), size=8.6, color=NAVY, bold=True,
                spacing=1.2)
        textbox(s, x + 0.20, 3.44, w - 0.4, 0.24, sym, size=9.4, color=EMERALD_DK, italic=True)
        card(s, x + 0.20, 3.72, w - 0.40, 0.62, fill=PAPER, line=None)
        textbox(s, x + 0.28, 3.80, w - 0.56, 0.48, formula, size=8.2,
                color=NAVY_SOFT, line=1.16)
        textbox(s, x + 0.20, 4.44, w - 0.40, 1.04, why, size=8.4, color=MUTED, line=1.2)

    textbox(s, M, 5.76, CW, 0.26,
            "Weights sum to 1. The score is recomputed from live records, and a suspended member's "
            "history stops contributing to the aggregates.", size=9.2, color=MUTED, italic=True)
    textbox(s, M, 6.08, CW, 0.26,
            "Why not a star average? Because a single perfect rating and a long, consistent history "
            "would score the same \u2014 and the second is the thing we are trying to measure.",
            size=9.2, color=MUTED, italic=True)
    return s


# ------------------------------------------------------------ 18 trust example
def trust_example(prs):
    s = content_slide(prs, "Core technique 03  \u00b7  RQ3",
                      "From score to badge to deposit tier")
    rows = [
        ["Dimension", "Member A", "Member B"],
        ["Ratings received", "4.8 \u00d7 6 owners, 4.5 \u00d7 2 renters", "5.0 \u00d7 1 owner"],
        ["Reputation C_rep", "0.813", "0.688"],
        ["Experience C_exp", "0.600", "0.100"],
        ["Reliability C_rel", "0.857", "1.000"],
        ["Verification C_ver", "1.000 (CNIC)", "0.400 (email)"],
        ["Composite S", "0.787", "0.574"],
        ["Band", "Trusted", "No badge"],
        ["PKR 100,000 deposit", "Permitted", "Denied"],
    ]
    table(s, M, 1.56, 6.05, rows, widths=[1.85, 2.30, 1.90], row_h=0.375, header_h=0.40,
          size=9.4, hsize=9, zebra=PAPER, bold_first=True, bold_rows=(6,),
          aligns=[PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.LEFT])
    textbox(s, M, 5.06, 6.05, 0.44,
            "Table 3.3 \u2014 Member B has a flawless single rating and still scores lower, because "
            "volume, reliability and identity carry weight.", size=8.6, color=MUTED, italic=True,
            line=1.18)

    # bands
    x2 = M + 6.35
    w2 = CW - 6.35
    textbox(s, x2, 1.48, w2, 0.24, "TRUST BANDS", size=8.4, color=MUTED, bold=True, spacing=1.3)
    bands = [("New Member", "no computed score yet", PAPER, NAVY_SOFT, LINE),
             ("Trusted", "0.60 \u2264 S < 0.80", MINT, EMERALD_DK, None),
             ("Highly Trusted", "S \u2265 0.80 with CNIC verified", NAVY, WHITE, None)]
    for i, (name, rule, fill, col, ln) in enumerate(bands):
        y = 1.76 + i * 0.50
        card(s, x2, y, w2, 0.42, fill=fill, line=ln)
        textbox(s, x2 + 0.18, y + 0.09, 1.9, 0.24, name, size=9.6, color=col, bold=True)
        textbox(s, x2 + 2.10, y + 0.10, w2 - 2.28, 0.24, rule, size=8.4,
                color=col if col != WHITE else MINT_TXT)

    # ladder
    textbox(s, x2, 3.32, w2, 0.24, "DEPOSIT ACCESS LADDER", size=8.4, color=MUTED, bold=True,
            spacing=1.3)
    tiers = [
        ("PKR 100,000 +", "CNIC verified, or 3 completed rentals", NAVY, WHITE, MINT_TXT),
        ("PKR 25,000 \u2013 99,999", "active account in good standing", MINT, EMERALD_DK, EMERALD_DK),
        ("Under PKR 25,000", "email and phone confirmed", PAPER, NAVY_SOFT, MUTED),
    ]
    for i, (amt, rule, fill, col, sub) in enumerate(tiers):
        y = 3.60 + i * 0.86
        card(s, x2, y, w2, 0.76, fill=fill, line=None if fill != PAPER else LINE)
        textbox(s, x2 + 0.18, y + 0.10, w2 - 0.36, 0.26, amt, size=10.4, color=col, bold=True)
        textbox(s, x2 + 0.18, y + 0.38, w2 - 0.36, 0.32, rule, size=8.4, color=sub, line=1.14)

    textbox(s, x2, 6.18, w2, 0.36,
            "Gating is progressive: nobody is asked for a CNIC to rent a drill.",
            size=8.8, color=MUTED, italic=True, line=1.16)
    return s


# ------------------------------------------------------------ 19 traceability
def traceability(prs):
    s = content_slide(prs, "Chapter 3  \u00b7  Traceability",
                      "Every design decision has a source and a question")
    rows = [
        ["Decision", "Type", "Source", "RQ"],
        ["One responsive web application rather than separate installed builds", "Architecture", "[21]", "RQ1"],
        ["Server-side lifecycle guard as the only path to a status change", "Architecture", "[10], [8]", "RQ1"],
        ["One account model serving renter, owner and administrator roles", "Architecture", "[24]", "RQ1"],
        ["Centralised administrator panel rather than on-chain arbitration", "Architecture", "[24], [23]", "RQ2"],
        ["A single relational store, so evidence and money commit together", "Architecture", "[8]", "RQ2"],
        ["Payment providers behind one interface, manual verification first", "Architecture", "[5], [4]", "RQ2"],
        ["Sealed pickup and return records with photographs and checksums", "Algorithm", "[9], [25]", "RQ2"],
        ["Deposit as its own state machine with a 48-hour return window", "Algorithm", "[8], [20]", "RQ2"],
        ["Four-component score with shrinkage and identity verification", "Algorithm", "[12], [13]", "RQ3"],
        ["Fixed taxonomy of 7 categories, 30 subcategories, 3 cities", "Dataset", "[2], [3]", "RQ1"],
    ]
    table(s, M, 1.54, CW, rows, widths=[7.05, 1.55, 1.75, 1.42], row_h=0.375, header_h=0.38,
          size=9.4, hsize=9, zebra=PAPER,
          aligns=[PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.LEFT])
    textbox(s, M, 6.10, CW, 0.26,
            "Table 3.4 maps all 21 decisions across architecture, data, algorithms and datasets \u2014 "
            "each to a peer-reviewed source and to RQ1, RQ2 or RQ3.",
            size=9, color=MUTED, italic=True)
    card(s, M, 6.36, CW, 0.42, fill=MINT, line=None)
    textbox(s, M + 0.24, 6.44, CW - 0.5, 0.26,
            "The matrix is also a defence: if a decision cannot name its source, it does not ship.",
            size=9.6, color=EMERALD_DK, bold=True)
    return s


# ------------------------------------------------------------ 20 implementation
def implementation(prs):
    s = content_slide(prs, "Chapter 3  \u00b7  Implementation",
                      "What is actually deployed")
    stack = [
        ("Client", "Next.js 15 App Router, React 19 Server Components, Tailwind"),
        ("Application", "Server actions, Auth.js v5, Zod payload validation"),
        ("Domain", "Pure TypeScript modules for lifecycle, escrow, scoring"),
        ("Data", "PostgreSQL + Prisma, 27 tables, integer PKR, private CNIC store"),
        ("Services", "Cloudinary media + checksums, Pusher Channels, cron workers"),
        ("Quality", "Vitest, Playwright against a production build, axe-core scans"),
    ]
    y = 1.54
    for name, detail in stack:
        card(s, M, y, 6.55, 0.66, fill=WHITE, line=LINE)
        textbox(s, M + 0.22, y + 0.20, 1.35, 0.26, name, size=9.6, color=EMERALD_DK, bold=True)
        textbox(s, M + 1.62, y + 0.14, 4.75, 0.46, detail, size=9.2, color=MUTED, line=1.18)
        y += 0.745

    # metric tiles
    x2 = M + 6.80
    w2 = CW - 6.80
    tiles = [("27", "database tables in one relational store"),
             ("24", "applied migrations, run during deploy"),
             ("808", "unit cases across 52 test files"),
             ("6", "deposit states and six scripted fault cases"),
             ("7 / 30", "categories and subcategories at launch"),
             ("3", "launch cities: Karachi, Lahore, Islamabad")]
    tw = (w2 - 0.18) / 2
    for i, (big, sub) in enumerate(tiles):
        x = x2 + (i % 2) * (tw + 0.18)
        yy = 1.54 + (i // 2) * 1.12
        card(s, x, yy, tw, 1.00, fill=PAPER, line=LINE)
        textbox(s, x + 0.18, yy + 0.12, tw - 0.3, 0.42, big, size=20, color=EMERALD, bold=True)
        textbox(s, x + 0.18, yy + 0.56, tw - 0.32, 0.4, sub, size=8.4, color=MUTED, line=1.14)

    card(s, M, 6.32, CW, 0.46, fill=NAVY, line=None)
    textbox(s, M + 0.24, 6.42, CW - 0.5, 0.26,
            "Deployed to staging on Vercel with Neon PostgreSQL \u00b7 migrations gated in the build "
            "\u00b7 end-to-end journey runs register \u2192 review in two browser contexts",
            size=9, color=WHITE)
    return s


# ------------------------------------------------------------ 21-23 walkthrough
def _shot(s, img, x, y, w, head, points, accent=EMERALD):
    card(s, x - 0.06, y - 0.06, w + 0.12, w * 0.578 + 0.12, fill=None, line=LINE, lw=0.75)
    picture(s, A + img, x, y, w=w)
    yy = y + w * 0.578 + 0.16
    textbox(s, x, yy, w, 0.26, head, size=11, color=NAVY, bold=True)
    bullets(s, x, yy + 0.30, w, 0.7, points, size=8.8, gap=3, line=1.14,
            bullet_color=accent, marker="\u00b7")


def walkthrough_a(prs):
    s = content_slide(prs, "Walkthrough  \u00b7  I", "Discovery, pricing and the booking record")
    _shot(s, "mock-browse.png", M, 1.58, 5.72,
          "Browse: rental terms, not sale prices",
          ["Categories, city and PKR-per-day rate on every card; verified-owner badge from the score.",
           "Search and filters run server-side; a listing cannot be published without a deposit amount."])
    _shot(s, "mock-booking.png", M + 5.96, 1.58, 5.72,
          "Booking: the terms are frozen at approval",
          ["Price and deposit are snapshotted onto the booking, so later listing edits cannot reach it.",
           "The status rail is the state machine made visible; payment awaits admin verification."],
          accent=AMBER_BR)
    textbox(s, M, 6.42, CW, 0.26,
            "Illustrative renderings of the deployed screens \u2014 the live walkthrough follows this "
            "presentation.", size=8.8, color=MUTED, italic=True)
    return s


def walkthrough_b(prs):
    s = content_slide(prs, "Walkthrough  \u00b7  II", "Sealing evidence and arbitrating a claim")
    _shot(s, "mock-handover.png", M, 1.58, 5.72,
          "Handover: camera-only, checksummed, sealed",
          ["Six photographs, a condition rating and notes; the record is immutable on write.",
           "Submitting it is what moves the booking to COMPLETED and starts the 48-hour clock."])
    _shot(s, "mock-admin.png", M + 5.96, 1.58, 5.72,
          "Admin: verify money, adjudicate damage",
          ["Manual payment verification with a full audit trail \u2014 the honest baseline for PKR rails.",
           "Claims compare pickup and return photos; the settlement figure is computed, not typed."],
          accent=AMBER_BR)
    textbox(s, M, 6.42, CW, 0.26,
            "The admin role is a privilege elevation on the same account \u2014 every write names the "
            "actor, the target and the reason.", size=8.8, color=MUTED, italic=True)
    return s


def walkthrough_c(prs):
    s = content_slide(prs, "Walkthrough  \u00b7  III", "What a member's standing looks like")
    _shot(s, "mock-trust.png", M, 1.58, 6.10,
          "One page: score, band and ladder",
          ["Composite score, band and identity state on one page, with the components behind it.",
           "The ladder shows what unlocks the next deposit tier before a member needs it."])
    x2 = M + 6.55
    w2 = CW - 6.55
    facts = [("Score, not sentiment", "Every figure traces to booking events: completions, self-cancellations, ratings, verification."),
             ("Reputation split by role", "Standing as an owner is never averaged with standing as a renter."),
             ("Badges earned, not bought", "Trusted and Highly Trusted are computed thresholds; CNIC is required for the top band."),
             ("Access scales with evidence", "Deposit tiers gate exposure without asking a first-time renter for documents.")]
    y = 1.58
    for i, (head, body) in enumerate(facts):
        card(s, x2, y, w2, 1.12, fill=WHITE if i % 2 else PAPER, line=LINE)
        textbox(s, x2 + 0.20, y + 0.14, w2 - 0.40, 0.26, head, size=10.2, color=NAVY, bold=True)
        textbox(s, x2 + 0.20, y + 0.44, w2 - 0.40, 0.62, body, size=8.6, color=MUTED, line=1.18)
        y += 1.22
    return s
