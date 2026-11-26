"""Slides 1-12: framing, problem, literature, objectives, method, architecture."""

from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt

from deck_theme import (
    NAVY, NAVY_DEEP, NAVY_SOFT, EMERALD, EMERALD_DK, EMERALD_BR, MINT, INK, MUTED,
    LINE, PAPER, SLATE, WHITE, AMBER, AMBER_BR, AMBER_SOFT, RED, RED_SOFT,
    SW, SH, M, CW, TOP, BOTTOM,
    bg, rect, card, pill, hline, arrow_right, arrow_down, circle_num,
    textbox, paragraph, bullets, table, content_slide, lead, caption, picture,
    pg, set_alpha,
)

ASSETS = "fyp-presentation/assets/"


# ------------------------------------------------------------------ 1 title
def title_slide(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s, NAVY)
    pg()

    # decorative arcs
    ring = rect(s, 9.35, -1.5, 6.4, 6.4, fill=NAVY_SOFT, line=None, shape=MSO_SHAPE.OVAL)
    set_alpha(ring, 55)
    ring2 = rect(s, 10.6, 4.05, 4.6, 4.6, fill=EMERALD_DK, line=None, shape=MSO_SHAPE.OVAL)
    set_alpha(ring2, 78)
    rect(s, 0, 0, 0.14, SH, fill=EMERALD, line=None, shape=MSO_SHAPE.RECTANGLE)

    picture(s, ASSETS + "logo-navy.png", M - 0.02, 0.55, w=0.86)
    textbox(s, M + 0.98, 0.72, 4.2, 0.34, "SamaanShare", size=19, color=WHITE, bold=True,
            spacing=0.3)
    textbox(s, M + 0.99, 1.06, 5.4, 0.26,
            "Staging deployment  \u00b7  Next.js 15  \u00b7  PostgreSQL",
            size=9.2, color=RGB_MINT_TEXT, spacing=0.5)

    textbox(s, SW - M - 5.6, 0.72, 5.6, 0.28,
            "FINAL YEAR DESIGN PROJECT  \u00b7  PHASE I  \u00b7  FINAL DEFENSE",
            size=9.5, color=RGB_MINT_TEXT, bold=True, align=PP_ALIGN.RIGHT, spacing=1.6)

    textbox(s, M, 1.92, 9.0, 0.9, "SamaanShare", size=50, color=WHITE, bold=True, line=1.0)
    textbox(s, M, 2.78, 8.3, 1.35,
            "A peer-to-peer household-item rental platform with sealed deposit "
            "protection and computed trust scoring",
            size=18, color=RGB_TITLE_SUB, line=1.28)
    hline(s, M, 4.16, 1.5, color=EMERALD_BR, lw=2.0)

    # team block
    textbox(s, M, 4.36, 3.4, 0.24, "GROUP 23FYDP24", size=9.5, color=EMERALD_BR,
            bold=True, spacing=1.5)
    members = [
        ("Ali Razi", "23F-BSCS-53", "Team Lead"),
        ("Muhammad Saad", "23F-BSCS-63", ""),
        ("M. Rehan Amjad", "23F-BSCS-65", ""),
        ("Sameer Raza", "23F-BSCS-69", ""),
        ("John Kazmi", "23F-BSCS-99", ""),
    ]
    col_w, row_h = 3.62, 0.30
    for i, (name, roll, role) in enumerate(members):
        col, row = i % 2, i // 2
        x = M + col * col_w
        y = 4.66 + row * row_h
        textbox(s, x, y, 1.64, 0.26, name, size=10.8, color=WHITE, bold=True)
        textbox(s, x + 1.62, y + 0.035, 1.98, 0.24,
                roll + (f"  \u00b7  {role}" if role else ""),
                size=8.6, color=RGB_TITLE_SUB)

    # supervisors
    textbox(s, M + 2 * col_w, 4.36, 3.2, 0.24, "SUPERVISION", size=9.5, color=EMERALD_BR,
            bold=True, spacing=1.5)
    textbox(s, M + 2 * col_w, 4.66, 3.4, 0.26, "Dr. Sarwat Iqbal", size=11.5, color=WHITE,
            bold=True)
    textbox(s, M + 2 * col_w, 4.94, 3.4, 0.24, "Supervisor", size=9.5, color=RGB_TITLE_SUB)
    textbox(s, M + 2 * col_w, 5.26, 3.4, 0.26, "Dr. Asma Larik", size=11.5, color=WHITE,
            bold=True)
    textbox(s, M + 2 * col_w, 5.54, 3.4, 0.24, "Co-supervisor", size=9.5, color=RGB_TITLE_SUB)

    # institution
    hline(s, M, 6.28, CW * 0.66, color=NAVY_SOFT, lw=0.9)
    textbox(s, M, 6.44, 8.4, 0.26,
            "Department of Computer Science  \u00b7  Faculty of Information and Computing Sciences",
            size=10, color=RGB_TITLE_SUB, spacing=0.3)
    textbox(s, M, 6.72, 8.8, 0.26,
            "Dawood University of Engineering and Technology, Karachi  \u00b7  November 2026",
            size=10, color=RGB_MINT_TEXT, spacing=0.3)
    textbox(s, M, 7.10, 6.0, 0.24, "FYDP coordinators: Engr. Imran Alvi  \u00b7  Engr. Ponam K.K.",
            size=8.5, color=RGB_TITLE_SUB)
    return s


from pptx.dml.color import RGBColor  # noqa: E402

RGB_TITLE_SUB = RGBColor(0xC9, 0xD8, 0xE2)
RGB_MINT_TEXT = RGBColor(0x8F, 0xC9, 0xB6)


# ------------------------------------------------------------------ 2 agenda
def agenda(prs):
    s = content_slide(prs, "Agenda", "What this presentation covers")
    items = [
        ("01", "Problem & context", "Why household assets sit idle in Pakistan, and why classifieds cannot fix it."),
        ("02", "Literature & gap", "25 studies across five themes; the blind spot they share."),
        ("03", "Aims & research questions", "Three objectives, three testable questions."),
        ("04", "Method & core techniques", "Six-layer architecture, state machine, sealed escrow, trust engine."),
        ("05", "Implementation & results", "What is deployed, and the evidence each question required."),
        ("06", "Limitations & FYDP-II", "What we could not do, and what comes next."),
    ]
    w, h = 5.72, 1.42
    for i, (num, head, desc) in enumerate(items):
        col, row = i % 2, i // 2
        x = M + col * (w + 0.33)
        y = TOP + row * (h + 0.24)
        card(s, x, y, w, h, fill=PAPER, line=LINE)
        textbox(s, x + 0.26, y + 0.24, 0.8, 0.5, num, size=22, color=EMERALD, bold=True)
        textbox(s, x + 1.02, y + 0.22, w - 1.3, 0.3, head, size=14.5, color=NAVY, bold=True)
        textbox(s, x + 1.02, y + 0.60, w - 1.28, 0.66, desc, size=10.5, color=MUTED, line=1.2)
    textbox(s, M, 6.52, CW, 0.28,
            "Roughly 20 minutes  \u00b7  a live walkthrough of the staging deployment follows the summary",
            size=10, color=MUTED, italic=True)
    return s


# ------------------------------------------------------------------ 3 background
def background(prs):
    s = content_slide(prs, "Chapter 1  \u00b7  Background", "A rental market exists; the trust to use it does not")
    lead(s, M, 1.52, CW - 0.2,
         "Between the five of us we own one DSLR, two industrial drill kits, two projectors and a "
         "bridal sherwani. All of it has sat unused for almost three years \u2014 and none of us would "
         "hand any of it to a stranger from a classifieds board.", size=12.5)

    stats = [
        ("~90%", "of Pakistani e-commerce still settles cash-on-delivery [4]",
         "card rails cannot be assumed"),
        ("3 yrs", "idle: the group's own camera, drills, projectors, sherwani",
         "high cost, near-zero utilisation"),
        ("0", "rental safeguards on the boards people use",
         "no booking record, deposit or ID check"),
    ]
    w = 3.72
    for i, (big, mid, small) in enumerate(stats):
        x = M + i * (w + 0.34)
        card(s, x, 2.28, w, 1.24, fill=MINT, line=None)
        textbox(s, x + 0.24, 2.40, 1.35, 0.6, big, size=27, color=EMERALD_DK, bold=True)
        textbox(s, x + 1.5, 2.42, w - 1.72, 0.62, mid, size=9.8, color=NAVY, line=1.16)
        textbox(s, x + 1.5, 3.14, w - 1.72, 0.26, small, size=8.8, color=EMERALD_DK, bold=True)

    # two explanatory columns
    colw = 5.72
    textbox(s, M, 3.78, colw, 0.28, "WHY OWNERSHIP FAILS FOR OCCASIONAL NEEDS", size=9,
            color=MUTED, bold=True, spacing=1.4)
    bullets(s, M, 4.10, colw, 1.9, [
        "A **six-figure** asset serves a **three-day** need \u2014 a wedding, a shoot, a deadline.",
        "Storage, servicing and depreciation land on one household; the item earns nothing between uses.",
        "Buying is the only option the market offers, so the need is either over-served or abandoned.",
    ], size=11.5, gap=7)

    x2 = M + colw + 0.36
    textbox(s, x2, 3.78, colw, 0.28, "WHY SHARING FAILS WITHOUT INSTITUTIONS", size=9,
            color=MUTED, bold=True, spacing=1.4)
    bullets(s, x2, 4.10, colw, 1.9, [
        "An owner cannot verify who is asking, and has no way to recover a damaged item [7].",
        "A renter will not pay an advance deposit to an unverified stranger \u2014 with no receipt, no proof.",
        "Nothing records condition at handover, so every dispute is one person's word against another's.",
    ], size=11.5, gap=7, bullet_color=AMBER)

    bar = rect(s, M, 5.98, CW, 0.72, fill=NAVY, line=None, radius=0.10)
    textbox(s, M + 0.3, 6.14, CW - 0.6, 0.44,
            "Economic theory says sharing platforms create welfare by cutting transaction costs [1] \u2014 "
            "but only where trust is produced. That is an engineering problem, not a marketing one.",
            size=11.5, color=WHITE, line=1.2)
    return s


# ------------------------------------------------------------------ 4 problem
def problem(prs):
    s = content_slide(prs, "Chapter 1  \u00b7  Problem statement",
                      "Three failures compound into one trust deficit")
    cards = [
        ("No rental-native marketplace",
         "OLX and Facebook groups were built for sales. There is no booking record, no availability "
         "calendar, no agreed rental term \u2014 so nothing that happens next is on the record.",
         ["sale-first listing model", "no lifecycle to skip or verify", "verbal terms only"]),
        ("The deposit deadlock",
         "The owner needs collateral for a damaged item. The renter will not send an unsecured "
         "advance to a stranger. Neither side can prove the item's condition at handover.",
         ["unsecured advance", "damage unprovable", "no release deadline"]),
        ("No reputation that means anything",
         "Five-star averages are thin and easily gamed; a first-time defaulter and a reliable member "
         "look identical. Without identity assurance there is nothing to weight them by.",
         ["gameable ratings", "no identity signal", "no access tiers"]),
    ]
    w = 3.72
    for i, (head, body, chips) in enumerate(cards):
        x = M + i * (w + 0.34)
        card(s, x, TOP, w, 3.30, fill=WHITE, line=LINE)
        rect(s, x, TOP, w, 0.055, fill=EMERALD if i != 1 else AMBER_BR, line=None,
             shape=MSO_SHAPE.RECTANGLE)
        circle_num(s, x + 0.24, TOP + 0.26, 0.38, str(i + 1),
                   fill=MINT, color=EMERALD_DK, size=12.5)
        textbox(s, x + 0.74, TOP + 0.26, w - 0.96, 0.6, head, size=13.5, color=NAVY,
                bold=True, line=1.1)
        textbox(s, x + 0.24, TOP + 1.02, w - 0.48, 1.4, body, size=10.3, color=MUTED, line=1.26)
        for j, chip in enumerate(chips):
            pill(s, x + 0.24, TOP + 2.26 + j * 0.30, w - 0.48, 0.26, chip,
                 fill=PAPER, color=NAVY_SOFT, size=9, bold=False, line=LINE)

    # consequence flow
    y = 5.30
    boxes = [("Owners delist", "one bad rental outweighs a year of income"),
             ("Renters lose access", "the affordable option never appears"),
             ("Assets sit idle", "capital locked in cupboards, not in use")]
    bw = 3.3
    for i, (head, sub) in enumerate(boxes):
        x = M + i * (bw + 0.72)
        card(s, x, y, bw, 0.92, fill=NAVY if i == 2 else PAPER, line=None if i == 2 else LINE)
        textbox(s, x + 0.2, y + 0.16, bw - 0.4, 0.28, head, size=12, bold=True,
                color=WHITE if i == 2 else NAVY)
        textbox(s, x + 0.2, y + 0.46, bw - 0.4, 0.36, sub, size=9.3,
                color=RGB_MINT_TEXT if i == 2 else MUTED, line=1.15)
        if i < 2:
            arrow_right(s, x + bw + 0.16, y + 0.34, 0.4, 0.24, color=LINE)
    textbox(s, M, 6.40, CW, 0.3,
            "A persistent trust deficit keeps a mutually beneficial sharing economy closed in Pakistan \u2014 "
            "Chapter 1, 1.3", size=9.5, color=MUTED, italic=True)
    return s


from pptx.dml.color import RGBColor  # noqa: E402

RGB_MINT_TEXT = RGBColor(0x8F, 0xC9, 0xB6)


# ------------------------------------------------------------------ 5 existing
def existing(prs):
    s = content_slide(prs, "Chapter 1 & 2  \u00b7  Competitive context",
                      "Why the options available today fall short")
    rows = [
        ["Option", "What it does provide", "Why it is not enough for this market"],
        ["Classifieds & Facebook groups",
         "Reach, price discovery, zero listing cost",
         "Built for outright sale: no booking record, no deposit, no condition evidence, no identity"],
        ["Global rental marketplaces",
         "Institutional escrow, card rails, dispute teams, strong trust UX",
         "Assume card payments and high baseline institutional trust; not built for household goods in PKR"],
        ["Blockchain rental prototypes [10], [22], [23]",
         "Deterministic lifecycle, automated escrow in smart contracts",
         "Wallet handling, gas fees and confirmation latency deter non-technical users [11]"],
        ["Academic trust & dispute models [9], [12], [13], [25]",
         "Validated multidimensional trust, review-blackmail theory",
         "Conceptual or survey-based: never bound to an escrow engine or a gated access ladder"],
    ]
    table(s, M, TOP, CW, rows, widths=[3.35, 3.95, 4.47], row_h=0.86, header_h=0.42,
          size=10.2, hsize=9.6, zebra=PAPER, bold_first=True)

    y = 5.62
    card(s, M, y, CW, 1.02, fill=MINT, line=None)
    textbox(s, M + 0.3, y + 0.14, 3.3, 0.26, "SAMANSHARE'S POSITION", size=8.8,
            color=EMERALD_DK, bold=True, spacing=1.4)
    textbox(s, M + 0.3, y + 0.40, CW - 0.6, 0.5,
            "Deterministic lifecycle  \u00b7  evidence-sealed escrow with bounded timers  \u00b7  "
            "computed trust gating access \u2014 engineered for a cash-first, PKR-denominated market.",
            size=11.5, color=NAVY, line=1.2)
    return s


# ------------------------------------------------------------------ 6 lit themes
def literature_themes(prs):
    s = content_slide(prs, "Chapter 2  \u00b7  Literature review",
                      "Five themes, twenty-five studies, one blind spot")
    themes = [
        ("01", "Microeconomics of sharing", "Rental thresholds, welfare gains, what counts as true P2P sharing.",
         "[1]  [3]  [11]"),
        ("02", "Computational trust", "Multidimensional reputation outperforms star averages; trust is role-specific.",
         "[12]  [13]  [14]  [15]"),
        ("03", "Identity assurance", "e-KYC, MFA and the convenience-versus-security trade-off.",
         "[6]  [16]  [17]  [18]"),
        ("04", "Localised finance", "Cash-on-delivery, wallet adoption and inclusion in Pakistan.",
         "[4]  [5]  [19]  [20]  [21]"),
        ("05", "Dispute & escrow mechanics", "State separation, arbitration models, review blackmail.",
         "[8]  [9]  [22]  [25]"),
    ]
    w = 2.24
    for i, (num, head, body, refs) in enumerate(themes):
        x = M + i * (w + 0.14)
        card(s, x, TOP, w, 2.92, fill=PAPER, line=LINE)
        textbox(s, x + 0.2, TOP + 0.18, 1.0, 0.34, num, size=15, color=EMERALD, bold=True)
        textbox(s, x + 0.2, TOP + 0.58, w - 0.4, 0.72, head, size=11.5, color=NAVY, bold=True,
                line=1.12)
        textbox(s, x + 0.2, TOP + 1.34, w - 0.4, 1.10, body, size=9.6, color=MUTED, line=1.24)
        textbox(s, x + 0.2, TOP + 2.50, w - 0.4, 0.3, refs, size=8.2, color=EMERALD_DK, bold=True)

    # takeaways
    y = 4.86
    card(s, M, y, CW, 1.52, fill=WHITE, line=LINE)
    textbox(s, M + 0.28, y + 0.16, 4.0, 0.26, "WHAT THE LITERATURE AGREES ON", size=8.8,
            color=MUTED, bold=True, spacing=1.4)
    points = [
        ("Renting raises utilisation", "below a per-use price threshold [1], [3]"),
        ("Trust must be multidimensional", "role-aware, not a star average [12], [14]"),
        ("Escrow must be separated from fees", "and released on a clock [8], [22]"),
        ("Feedback needs anti-retaliation design", "reciprocal release [25]"),
    ]
    for i, (a, b) in enumerate(points):
        x = M + 0.28 + (i % 2) * 5.68
        yy = y + 0.50 + (i // 2) * 0.46
        textbox(s, x, yy, 5.5, 0.34,
                [(a, {"bold": True, "color": NAVY}), "  " + b],
                size=9.8, color=MUTED, line=1.1)
    return s


# ------------------------------------------------------------------ 7 lit table
def literature_table(prs):
    s = content_slide(prs, "Chapter 2  \u00b7  Key studies",
                      "The studies that shaped each design decision")
    rows = [
        ["Study", "Focus", "Method", "Finding \u2192 what we took from it"],
        ["Filippas et al. (2020) [1]", "Economics of sharing", "Analytical model",
         "Renting beats owning below a use threshold; platforms create welfare by cutting transaction cost \u2192 justifies P2P rental of durables"],
        ["Gerwe & Silva (2020) [3]", "Sharing-economy taxonomy", "Review, 120+ platforms",
         "Separates true P2P from pseudo-sharing \u2192 our marketplace boundary: peer-to-peer, not a rental shop"],
        ["Tele et al. (2026) [10]", "VahanRent: blockchain car rental", "System build, Polygon",
         "Automated lifecycle and escrow are achievable \u2192 we kept the state design, dropped the chain and its friction"],
        ["Salgado et al. (2026) [12]", "Multidimensional trust", "Survey, PLS-SEM",
         "Trust is role- and dimension-specific \u2192 four components, ratings split by renter/owner direction"],
        ["Anjum & Chai (2020) [4]", "Cash-on-delivery in Pakistan", "Empirical study",
         "Cash-first buyers dominate local e-commerce \u2192 manual, auditable payment path before any gateway"],
        ["Papanastasiou et al. (2023) [25]", "Review blackmail on platforms", "Game-theoretic model",
         "Retaliatory reviews distort feedback \u2192 14-day window with reciprocal, sealed release"],
    ]
    table(s, M, TOP, CW, rows, widths=[2.35, 2.25, 2.05, 5.12], row_h=0.70, header_h=0.40,
          size=9.4, hsize=9.2, zebra=PAPER, bold_first=True)
    textbox(s, M, 6.52, CW, 0.26,
            "Table 2.1 compares all 25 reviewed studies across methodology, dataset, findings, limitations "
            "and relevance to this gap.", size=9.2, color=MUTED, italic=True)
    return s


# ------------------------------------------------------------------ 8 gap
def gap(prs):
    s = content_slide(prs, "Chapter 2  \u00b7  Research gap",
                      "Everyone studied the pieces. Nobody assembled them.")
    colw = 3.78
    panels = [
        ("What has been studied", EMERALD, [
            "Rental thresholds and welfare gains of sharing [1], [3]",
            "Decentralised carsharing and on-chain escrow [10], [22], [23]",
            "Multidimensional trust and algorithmic reputation [12], [13], [14]",
            "Dispute prevention and review-blackmail models [9], [25]",
        ]),
        ("What remains missing", AMBER_BR, [
            "Work assumes card rails, Stripe-style escrow and high baseline institutional trust.",
            "Blockchain answers add wallet and gas complexity that deter non-technical users [11].",
            "No architecture blends **local mobile wallets** with **state-guarded, evidence-sealed escrow** for physical household goods in a cash-dominant market [4], [5].",
            "Handover evidence is never bound to **settlement arithmetic** \u2014 so owners cannot prove damage and renters face indefinite holds.",
        ]),
    ]
    for i, (head, accent, items) in enumerate(panels):
        x = M + i * (colw + 0.30)
        card(s, x, TOP, colw, 2.92, fill=PAPER, line=LINE)
        rect(s, x, TOP, 0.05, 2.92, fill=accent, line=None, shape=MSO_SHAPE.RECTANGLE)
        textbox(s, x + 0.26, TOP + 0.18, colw - 0.5, 0.28, head, size=12.5, color=NAVY,
                bold=True)
        bullets(s, x + 0.26, TOP + 0.54, colw - 0.52, 2.3, items, size=9.4, gap=5,
                bullet_color=accent, line=1.16)

    # response panel
    x3 = M + 2 * (colw + 0.30)
    w3 = CW - 2 * (colw + 0.30)
    card(s, x3, TOP, w3, 2.92, fill=NAVY, line=None)
    textbox(s, x3 + 0.26, TOP + 0.18, w3 - 0.5, 0.28, "How SamaanShare responds", size=12.5,
            color=WHITE, bold=True)
    resp = [
        ("Guarded lifecycle", "server actions + atomic conditional SQL"),
        ("Sealed handover, bounded escrow", "checksums, triggers, 48-hour clock"),
        ("Four-factor trust engine", "with a deposit access ladder"),
    ]
    for i, (a, b) in enumerate(resp):
        y = TOP + 0.62 + i * 0.74
        circle_num(s, x3 + 0.26, y, 0.30, str(i + 1), fill=EMERALD_BR, color=WHITE, size=10.5)
        textbox(s, x3 + 0.66, y - 0.02, w3 - 0.9, 0.26, a, size=10.5, color=WHITE, bold=True)
        textbox(s, x3 + 0.66, y + 0.24, w3 - 0.9, 0.3, b, size=9, color=RGB_MINT_TEXT, line=1.1)

    # gap statement
    y = 5.02
    card(s, M, y, CW, 1.56, fill=WHITE, line=LINE)
    textbox(s, M + 0.3, y + 0.16, 2.2, 0.26, "THE GAP, IN ONE LINE", size=8.8, color=EMERALD_DK,
            bold=True, spacing=1.4)
    textbox(s, M + 0.3, y + 0.44, CW - 0.6, 0.98,
            "No published system architecture combines a deterministic rental lifecycle, a "
            "checksummed two-sided handover record, time-bounded deposit settlement arithmetic and a "
            "computed trust score into one platform designed for cash-dominant, low-trust emerging markets.",
            size=12.2, color=NAVY, line=1.26)
    return s


from pptx.dml.color import RGBColor  # noqa: E402

RGB_MINT_TEXT = RGBColor(0x8F, 0xC9, 0xB6)


# ------------------------------------------------------------------ 9 objectives
def objectives(prs):
    s = content_slide(prs, "Chapter 1  \u00b7  Aims & objectives",
                      "One aim, three SMART objectives")
    card(s, M, 1.50, CW, 0.90, fill=NAVY, line=None)
    textbox(s, M + 0.3, 1.60, 2.0, 0.24, "OVERARCHING AIM", size=8.8, color=EMERALD_BR,
            bold=True, spacing=1.5)
    textbox(s, M + 0.3, 1.86, CW - 0.6, 0.46,
            "Design, implement and assess an integrated P2P household-item rental platform with "
            "deterministic lifecycle enforcement, evidence-sealed deposit management and "
            "multidimensional computational trust scoring.",
            size=11.6, color=WHITE, line=1.2)

    rows = [
        ("O1", "Build the guarded rental application",
         "Every rental from listing to verified return runs through one tracked lifecycle that cannot "
         "skip a step.",
         "RQ1", "Guarded lifecycle", "integration suite"),
        ("O2", "Make the deposit safe for both sides",
         "A security deposit backed by sealed records of both handovers and a time-bounded claim "
         "process, settled on evidence rather than assertion.",
         "RQ2", "Sealed escrow", "six scenarios"),
        ("O3", "Compute and gate on trust",
         "A computed trust score plus a deposit-based access ladder that measures how reliably a "
         "member keeps their side of a rental.",
         "RQ3", "Trust engine", "2,000 profiles"),
    ]
    y = 2.66
    for code, head, body, rq, tag, tested in rows:
        card(s, M, y, CW, 1.24, fill=WHITE, line=LINE)
        rect(s, M, y, 0.055, 1.24, fill=EMERALD, line=None, shape=MSO_SHAPE.RECTANGLE)
        textbox(s, M + 0.28, y + 0.24, 0.72, 0.4, code, size=17, color=EMERALD, bold=True)
        textbox(s, M + 1.08, y + 0.18, 3.9, 0.3, head, size=12.8, color=NAVY, bold=True)
        textbox(s, M + 1.08, y + 0.50, 5.5, 0.62, body, size=10, color=MUTED, line=1.22)
        pill(s, M + 7.05, y + 0.26, 0.95, 0.28, rq, fill=MINT, color=EMERALD_DK, size=10)
        pill(s, M + 8.12, y + 0.26, 2.05, 0.28, tag, fill=PAPER, color=NAVY_SOFT, size=9.5,
             bold=False, line=LINE)
        textbox(s, M + 10.42, y + 0.20, 1.25, 0.24, "TESTED BY", size=7.6, color=MUTED,
                bold=True, spacing=1.2)
        textbox(s, M + 10.42, y + 0.44, 1.35, 0.50, tested, size=8.6, color=EMERALD_DK,
                bold=True, line=1.15)
        y += 1.36

    textbox(s, M, 6.86 - 0.06, CW, 0.24,
            "Each objective maps one-to-one to a research question, so every deliverable is also a testable claim.",
            size=9.2, color=MUTED, italic=True)
    return s


# ------------------------------------------------------------------ 10 RQs
def research_questions(prs):
    s = content_slide(prs, "Chapter 1  \u00b7  Research questions",
                      "Three questions, each with a stated evidence base")
    rows = [
        ("RQ1", "How can the rental application be designed so that every listing, approval, payment, "
                "handover and review is recorded and cannot skip a step?",
         "Automated integration suite attempting every illegal transition against the live Next.js app",
         "Zero accepted illegal transitions \u00b7 exactly one audit record per permitted transition"),
        ("RQ2", "How can the deposit, a sealed handover record and a time-bounded claim process be "
                "combined to protect the owner when an item returns damaged \u2014 without holding the "
                "renter's money indefinitely?",
         "Six scripted deposit scenarios, each against an independently hand-calculated settlement",
         "Settlements match ground truth exactly \u00b7 no balance locked past its stated deadline"),
        ("RQ3", "How precisely can a trust score computed from completed rentals, ratings, "
                "cancellations and identity verification separate reliable members from defaulting ones?",
         "Pilot member histories plus an offline sensitivity analysis over 2,000 modelled profiles",
         "Measurable score discrimination between completing and cancelling members"),
    ]
    y = TOP
    for code, q, ev, crit in rows:
        card(s, M, y, CW, 1.52, fill=PAPER, line=LINE)
        textbox(s, M + 0.26, y + 0.22, 0.85, 0.36, code, size=14.5, color=EMERALD, bold=True)
        textbox(s, M + 1.12, y + 0.16, 7.30, 0.80, q, size=10.6, color=NAVY, line=1.24)
        textbox(s, M + 1.12, y + 1.00, 6.9, 0.34,
                [("Evidence  ", {"bold": True, "color": EMERALD_DK, "size": 8.6}), ev],
                size=9.2, color=MUTED, line=1.15)
        card(s, M + 8.62, y + 0.16, 2.95, 1.20, fill=WHITE, line=LINE)
        textbox(s, M + 8.78, y + 0.26, 2.66, 0.22, "SUCCESS CRITERION", size=7.4, color=MUTED,
                bold=True, spacing=1.2)
        textbox(s, M + 8.78, y + 0.50, 2.66, 0.80, crit, size=8.8, color=EMERALD_DK, line=1.2)
        y += 1.62
    return s


# ------------------------------------------------------------------ 11 method
def methodology(prs):
    s = content_slide(prs, "Chapter 3  \u00b7  Methodology",
                      "Build, then measure \u2014 one framework per technique")
    textbox(s, M, 1.50, 6.0, 0.26, "RESEARCH DESIGN", size=8.8, color=MUTED, bold=True,
            spacing=1.4)
    steps = [("Design", "Lifecycle, escrow and score derived from the literature gap"),
             ("Build", "Six decoupled layers, deployed to staging as one web application"),
             ("Evaluate", "Scripted scenarios, integration suites and a modelled population")]
    y = 1.80
    for i, (head, body) in enumerate(steps):
        card(s, M, y, 5.85, 0.78, fill=WHITE, line=LINE)
        circle_num(s, M + 0.20, y + 0.20, 0.38, str(i + 1), fill=MINT, color=EMERALD_DK, size=11)
        textbox(s, M + 0.72, y + 0.12, 1.3, 0.26, head, size=11.5, color=NAVY, bold=True)
        textbox(s, M + 1.90, y + 0.13, 3.8, 0.52, body, size=9.3, color=MUTED, line=1.18)
        if i < 2:
            arrow_down(s, M + 2.55, y + 0.78, 0.18, 0.16, color=LINE)
        y += 0.94

    # six-point framework
    textbox(s, M, 4.72, 6.0, 0.26, "SIX-POINT FRAMEWORK APPLIED TO EACH TECHNIQUE", size=8.8,
            color=MUTED, bold=True, spacing=1.4)
    fw = [("Technique", "what it is"), ("Mechanism", "how it works"),
          ("Literature", "basis in Ch. 2"), ("Adaptation", "what changes locally"),
          ("Evaluation", "how we judged it"), ("Question", "which RQ it serves")]
    for i, (a, b) in enumerate(fw):
        x = M + (i % 3) * 1.97
        yy = 5.02 + (i // 3) * 0.78
        card(s, x, yy, 1.86, 0.66, fill=MINT, line=None)
        textbox(s, x + 0.14, yy + 0.10, 1.6, 0.24, a, size=9.6, color=EMERALD_DK, bold=True)
        textbox(s, x + 0.14, yy + 0.34, 1.6, 0.26, b, size=8.2, color=NAVY_SOFT)

    # right column: evidence base
    x2 = M + 6.20
    w2 = CW - 6.20
    card(s, x2, 1.50, w2, 5.12, fill=PAPER, line=LINE)
    textbox(s, x2 + 0.28, 1.64, w2 - 0.5, 0.26, "EVIDENCE AND QUALITY CONTROLS", size=8.8,
            color=MUTED, bold=True, spacing=1.4)
    ev = [
        ("Reference taxonomy", "7 categories, 30 subcategories, 3 launch cities"),
        ("Pilot records", "live staging transactions: timestamps, transitions, rulings"),
        ("Scripted scenarios", "six end-to-end deposit cases with hand-calculated baselines"),
        ("Modelled population", "2,000 synthetic profiles with stated behaviour distributions"),
    ]
    yy = 1.98
    for a, b in ev:
        textbox(s, x2 + 0.28, yy, w2 - 0.55, 0.24, a, size=10, color=NAVY, bold=True)
        textbox(s, x2 + 0.28, yy + 0.24, w2 - 0.55, 0.30, b, size=9, color=MUTED, line=1.16)
        yy += 0.66

    hline(s, x2 + 0.28, 4.72, w2 - 0.56, color=LINE, lw=0.75)
    textbox(s, x2 + 0.28, 4.82, w2 - 0.5, 0.26, "NON-NEGOTIABLE INVARIANTS", size=8.8,
            color=MUTED, bold=True, spacing=1.4)
    bullets(s, x2 + 0.28, 5.12, w2 - 0.56, 1.4, [
        "One payment row per booking; at most one pickup and one return log",
        "Every evidence photo carries SHA-256 read server-side; no duplicate hashes",
        "Money stored as integer PKR \u2014 floats rejected at the ORM layer",
        "Payout + commission + refund must equal the amount collected",
    ], size=8.8, gap=4, marker="\u2713")
    return s


# ------------------------------------------------------------------ 12 architecture
def architecture(prs):
    s = content_slide(prs, "Chapter 3  \u00b7  System architecture",
                      "Six decoupled layers, one guarded path to the database")
    layers = [
        ("1", "User", "Renter, owner and administrator share one account; admin is a privilege elevation, not a second login",
         "Three roles, one identity"),
        ("2", "Client", "Next.js App Router with React Server Components; responsive on phone, tablet and desktop without an install",
         "No app-store friction"),
        ("3", "Application", "Every mutation is a server action; each runs a mandatory five-stage guard before touching data",
         "Session \u00b7 role \u00b7 ownership \u00b7 state \u00b7 payload"),
        ("4", "Domain rule", "Pure TypeScript modules: booking lifecycle, deposit states, claims, handover, scoring, review release",
         "No driver, no session, no ambient clock"),
        ("5", "Platform services", "Provider-agnostic payments, cron timers, Cloudinary storage, Pusher realtime, notifications",
         "Swappable external boundaries"),
        ("6", "Data", "PostgreSQL through Prisma: 27 tables, 24 migrations, integer PKR, private CNIC storage",
         "One store, so evidence and money commit together"),
    ]
    y = 1.50
    for num, name, body, note in layers:
        h = 0.70
        card(s, M, y, 7.55, h, fill=WHITE if num not in "36" else PAPER, line=LINE)
        circle_num(s, M + 0.20, y + 0.17, 0.36, num, fill=EMERALD if num == "3" else MINT,
                   color=WHITE if num == "3" else EMERALD_DK, size=10.5)
        textbox(s, M + 0.68, y + 0.10, 1.55, 0.26, name + " layer", size=10.4, color=NAVY,
                bold=True)
        textbox(s, M + 2.28, y + 0.09, 5.15, 0.56, body, size=8.8, color=MUTED, line=1.14)
        y += h + 0.045

    x2 = M + 7.88
    w2 = CW - 7.88
    card(s, x2, 1.50, w2, 2.30, fill=MINT, line=None)
    textbox(s, x2 + 0.24, 1.62, w2 - 0.45, 0.26, "THE FIVE-STAGE GUARD", size=8.6,
            color=EMERALD_DK, bold=True, spacing=1.3)
    stages = ["Active session", "Role permission", "Record ownership",
              "Status \u2260 expected \u2192 reject", "Zod payload validation"]
    for i, st in enumerate(stages):
        yy = 1.94 + i * 0.35
        textbox(s, x2 + 0.24, yy, 0.3, 0.24, f"{i + 1}", size=8.6, color=EMERALD_DK, bold=True)
        textbox(s, x2 + 0.50, yy, w2 - 0.7, 0.24, st, size=8.8, color=NAVY)

    card(s, x2, 4.00, w2, 1.32, fill=WHITE, line=LINE)
    textbox(s, x2 + 0.24, 4.12, w2 - 0.45, 0.26, "WHY THE SPLIT MATTERS", size=8.6, color=MUTED,
            bold=True, spacing=1.3)
    textbox(s, x2 + 0.24, 4.38, w2 - 0.45, 0.86,
            "Business invariants in layer 4 can be exercised without a database, a session or a clock "
            "\u2014 which is what makes the illegal-transition suite possible.",
            size=9, color=NAVY, line=1.24)

    card(s, x2, 5.50, w2, 1.10, fill=NAVY, line=None)
    textbox(s, x2 + 0.24, 5.62, w2 - 0.45, 0.26, "STACK", size=8.2, color=EMERALD_BR, bold=True,
            spacing=1.3)
    textbox(s, x2 + 0.24, 5.88, w2 - 0.45, 0.62,
            "Next.js 15  \u00b7  React 19  \u00b7  TypeScript  \u00b7  PostgreSQL  \u00b7  Prisma  \u00b7  "
            "Auth.js v5  \u00b7  Cloudinary  \u00b7  Pusher", size=8.8, color=WHITE, line=1.3)
    return s
