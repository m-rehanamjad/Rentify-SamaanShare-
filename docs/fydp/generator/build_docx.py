"""Builds the FYDP-1 Proposed Methodology Baseline as a Word document.

    python3 build_docx.py

Writes 23FYDP24_SamaanShare_MethodologyBaseline_FYDP1.docx into /deliverables.
"""

from __future__ import annotations

import os

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt
from docx.oxml.ns import qn

from docx_kit import (
    GREY,
    INK,
    NAVY,
    bullets,
    callout,
    figure,
    new_document,
    numbered,
    page_break,
    para,
    reflection,
    rule,
    setup_header_footer,
    table,
    two_col_table,
    update_fields_on_open,
)

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.normpath(
    os.path.join(
        HERE, "..", "..", "..", "deliverables",
        "23FYDP24_SamaanShare_MethodologyBaseline_FYDP1.docx",
    )
)

GROUP_ID = "23FYDP24"
MEMBERS = [
    ("Ali Razi", "23F-BSCS-53"),
    ("Muhammad Saad", "23F-BSCS-63"),
    ("Rehan Amjad", "23F-BSCS-65"),
    ("Sameer Raza", "23F-BSCS-69"),
    ("John Kazmi", "23F-BSCS-99"),
]


def lead_para(doc, lead: str, text: str) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(5)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    run = p.add_run(lead + " ")
    run.bold = True
    run.font.color.rgb = NAVY
    p.add_run(text)


# --------------------------------------------------------------------------
# front matter


def title_page(doc):
    para(doc, "", space_after=2)
    p = para(
        doc,
        "DAWOOD UNIVERSITY OF ENGINEERING AND TECHNOLOGY",
        size=11.5, bold=True, color=NAVY, align="center", space_after=2,
    )
    for run in p.runs:
        run.font.name = "Calibri"
    para(doc, "KARACHI, PAKISTAN", size=9.5, color=GREY, align="center", space_after=14)
    rule(doc, space_before=0, space_after=14)
    para(
        doc,
        "Faculty of Information and Computing Sciences  ·  Department of Computer Science",
        size=9, color=GREY, align="center", space_after=30,
    )

    para(
        doc, "FINAL YEAR DESIGN PROJECT  ·  FYDP-1",
        size=9.5, bold=True, color=GREY, align="center", space_after=10,
    )
    p = para(doc, "Proposed Methodology", size=27, bold=True, color=NAVY, align="center", space_after=0)
    for run in p.runs:
        run.font.name = "Georgia"
    p = para(doc, "Baseline", size=27, bold=True, color=NAVY, align="center", space_after=16)
    for run in p.runs:
        run.font.name = "Georgia"

    para(
        doc,
        "Problem Statement, Objectives, Research Questions\nand the Core Technical Methodology",
        size=11, italic=True, color=GREY, align="center", space_after=26,
    )

    rule(doc, space_before=0, space_after=18)

    p = para(doc, "SamaanShare", size=17, bold=True, color=INK, align="center", space_after=4)
    for run in p.runs:
        run.font.name = "Georgia"
    para(
        doc,
        "A Peer-to-Peer Rental Platform with Secure Deposit and Trust Scoring",
        size=11, color=INK, align="center", space_after=34,
    )

    two_col_table(
        doc,
        [
            ("Group ID", GROUP_ID),
            ("Programme", "BS Computer Science  ·  Batch 2023  ·  Semester 23-Fall"),
            ("Supervisor", "Dr. Sarwat Iqbal"),
            ("Submitted to", "FYDP Coordination Committee, Department of Computer Science"),
            ("Date", "November 2026"),
        ],
        widths=(3.4, 8.2),
        font_size=9.2,
    )

    para(doc, "", space_after=18)
    para(doc, "Submitted by", size=8.6, bold=True, color=GREY, align="center", space_after=5)
    names = "   ·   ".join(f"{name} ({roll})" for name, roll in MEMBERS)
    para(doc, names, size=9, color=INK, align="center", space_after=22)

    para(
        doc,
        "Prerequisite completed: Chapter 2, Preliminary Literature Review, approved by the supervisor.\n"
        "This chapter will be extended in FYDP-2 with implementation results and evaluation feedback.",
        size=8, italic=True, color=GREY, align="center", space_after=0,
    )


def contents_page(doc):
    heading = doc.add_paragraph(style="Heading 1")
    heading.paragraph_format.space_before = Pt(0)
    heading.add_run("Contents")

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(8)
    run = p.add_run(
        "Page numbers are fields. If they appear as placeholders, press Ctrl+A then F9 to refresh."
    )
    run.font.size = Pt(7.8)
    run.italic = True
    run.font.color.rgb = GREY

    toc_para = doc.add_paragraph()
    toc_para.paragraph_format.space_after = Pt(0)
    run = toc_para.add_run()
    run._r.append(run._r.makeelement(qn("w:fldChar"), {qn("w:fldCharType"): "begin"}))
    run = toc_para.add_run()
    instr = run._r.makeelement(qn("w:instrText"), {qn("xml:space"): "preserve"})
    instr.text = r'TOC \o "1-2" \h \z \u'
    run._r.append(instr)
    run = toc_para.add_run()
    run._r.append(run._r.makeelement(qn("w:fldChar"), {qn("w:fldCharType"): "separate"}))
    run = toc_para.add_run("Right-click here and choose Update Field to build the contents list.")
    run.font.size = Pt(9)
    run.font.color.rgb = GREY
    run.italic = True
    run = toc_para.add_run()
    run._r.append(run._r.makeelement(qn("w:fldChar"), {qn("w:fldCharType"): "end"}))

    para(doc, "Figures", size=9.5, bold=True, color=NAVY, space_before=16, space_after=3)
    for line in (
        "Figure 1  Layered architecture of the proposed platform",
        "Figure 2  The deposit, handover and claim path",
    ):
        para(doc, line, size=8.8, space_after=2, indent=0.4)

    para(doc, "Tables", size=9.5, bold=True, color=NAVY, space_before=10, space_after=3)
    for line in (
        "Table 1  Where each sentence of the problem statement comes from",
        "Table 2  The three objectives and what each one commits us to",
        "Table 3  Objectives and research questions in strict one-to-one order",
        "Table 4  What the methodology has to deliver, question by question",
        "Table 5  Architectural decisions and their justification in Chapter 2",
        "Table 6  The data the platform holds, and which question it serves",
        "Table 7  Risks to the data plan and the fallback for each",
        "Table 8  Settlement arithmetic for one booking",
        "Table 9  The trust score worked through for two members",
        "Table 10  How each technique will be evaluated",
        "Table 11  Traceability matrix: decision, source, research question",
        "Table 12  AI tool usage record for this chapter",
        "Table 13  Decisions revised since the literature review",
        "Table 14  Implementation status at the time of writing",
    ):
        para(doc, line, size=8.8, space_after=2, indent=0.4)


def note_page(doc):
    h = doc.add_paragraph(style="Heading 1")
    h.paragraph_format.space_before = Pt(0)
    h.add_run("A note before you read")

    para(
        doc,
        "This document covers Heading 3 and Heading 4 of the FYDP-1 methodology baseline and nothing "
        "else. It states direction rather than final depth: tools, development process, evaluation "
        "results, assumptions and limitations belong to FYDP-2, once there is a running system to "
        "describe honestly.",
    )
    para(
        doc,
        "The platform described here is not a proposal on paper. SamaanShare is built to a pre-release "
        "state and deployed to a staging server, and the figures in this document — the number of database "
        "tables, the time windows, the score weights, the deposit thresholds — are read from the working "
        "codebase rather than estimated for the report. Where something is not yet built, the text says so "
        "in plain words instead of implying that it is.",
    )
    para(
        doc,
        "One choice of words matters throughout. SamaanShare is a single application that runs in a "
        "browser on a laptop, a tablet or a phone, so we have deliberately not described it as something "
        "a user installs. That would be a different product with a different cost and a different release "
        "process, and Appendix B records why the original plan changed.",
    )
    callout(
        doc,
        "How this chapter was written",
        "Every section was drafted by the group from our own reading of Chapter 2, then passed through "
        "the two-tool rule: one AI tool to improve and structure the draft, and a second, different tool "
        "to check it for technical errors, unsupported claims and citation drift. The record of what each "
        "tool did, and what we changed because of it, is in Appendix A. The five reflections are the "
        "exception; they were written by us without drafting help, which is why they read less smoothly "
        "than the rest.",
    )


# --------------------------------------------------------------------------
# heading 3


def section_three(doc):
    doc.add_paragraph("3.  Problem Statement, Objectives and Research Questions", style="Heading 1")

    para(
        doc,
        "Chapter 2 closed with a gap rather than a topic. The sharing-economy literature explains why "
        "people rent instead of buying, and it explains how a platform manufactures trust between "
        "strangers, but it does not join the two for a market like ours — where the deposit has to be "
        "collected in rupees from someone who has never met the owner, and no card network stands behind "
        "it. The three parts below turn that gap into a problem statement, then into three objectives, "
        "then into three questions that answer those objectives one for one.",
    )
    para(
        doc,
        "The order is not decorative. A question that cannot be traced to an objective, or an objective "
        "that cannot be traced to a sentence of the problem statement, is exactly what the department's "
        "rubric penalises. So each part is written to be checked against the part before it, and the "
        "checks are shown rather than asserted.",
    )

    doc.add_paragraph("3.1  Problem Statement", style="Heading 2")
    para(
        doc,
        "Three sentences follow, and each one has a job: state the situation and who it affects, narrow "
        "it to the part we can act on, and explain why it persists. That last part is what makes it a "
        "research problem rather than an observation.",
    )

    callout(
        doc,
        "Problem statement",
        "Pakistan has no marketplace built for renting everyday household goods between people, so a "
        "camera, a generator or a wedding outfit that is used a few times a year sits idle in one home "
        "while somebody two streets away pays full price for the same thing.\n\n"
        "What people use instead — classified sites, Facebook groups, a message to a cousin — can carry a "
        "sale but not a rental: there is no booking record, no proof of the condition an item left in and "
        "came back in, and no deposit, so the owner carries every rupee of the risk and most owners "
        "quietly decide not to bother.\n\n"
        "The reason that risk is refused rather than priced is trust, and it is specific to this market: "
        "Pakistani shoppers still prefer cash on delivery because they do not trust online sellers to "
        "keep their word, and owners of expensive goods will not list them at all until the platform "
        "itself has checked who the other party is — which is why a market that would earn money for "
        "owners and save it for renters stays closed.",
        fill="EDF2F8",
        edge="7C9CBF",
    )

    para(
        doc,
        "The middle sentence does the narrowing. We are not proposing to fix online trust in Pakistan, "
        "and we are not proposing a general classifieds platform. The slice we take is one physical item "
        "rented between two individuals, in three cities, changing hands in person or by courier, with the "
        "money for it held by the platform until the handover is closed.",
    )

    table(
        doc,
        ["Part of the statement", "What it claims", "Where it comes from in Chapter 2"],
        [
            [
                "First sentence",
                "Expensive household items sit idle while occasional use is expensive; no rental-specific "
                "marketplace serves this.",
                "Section 4 (gap analysis), opening paragraph. The economics are Filippas et al. (2020) on "
                "the rental threshold; the closest local system analogue is Ragavarthini et al. (2025).",
            ],
            [
                "Second sentence",
                "Classified platforms carry a sale but not a rental: no lifecycle record, no condition "
                "evidence, no deposit.",
                "Gerwe and Silva (2020) on real peer platforms versus listing boards; Kotiyal et al. (2026) "
                "on separating rent from a refundable deposit; Zhao and Chen (2024) on handover records.",
            ],
            [
                "Third sentence",
                "The binding constraint is trust, and it is a property of this market rather than of "
                "sharing platforms in general.",
                "Anjum and Chai (2020) on cash-on-delivery dominance; Li and Wang (2020) on provider-side "
                "trust; Malik et al. (2025) on mobile-wallet trust; Ghozi (2022) on identity checks as the "
                "accepted remedy.",
            ],
        ],
        widths=[3.0, 6.0, 6.5],
        label="Table 1",
        caption_text="Where each sentence of the problem statement comes from. This mapping is the evidence "
        "for the department's first golden rule: nothing is asserted without a row of Chapter 2 behind it.",
    )

    para(
        doc,
        "Two things are deliberately absent. There is no claim that rental platforms are new — Fat Llama "
        "and Rent the Runway did it elsewhere, and Chapter 2 says so — and no claim that SamaanShare ends "
        "fraud. It makes fraud expensive and traceable, which is a smaller promise we can actually keep.",
    )

    doc.add_paragraph("3.2  Objectives", style="Heading 2")
    para(
        doc,
        "Exactly three objectives follow, one for each job the problem statement leaves open: record the "
        "rental, protect the money, and give a stranger something to judge. A fourth candidate — the "
        "optional fraud detector carried over from the project proposal — was dropped rather than merged, "
        "because at this project's transaction volume there is no evidential path for it.",
    )

    table(
        doc,
        ["", "Objective", "What it commits us to in practice"],
        [
            [
                "O1",
                "To design and develop a web-based peer-to-peer rental platform for household items in "
                "which every rental passes through a tracked and recorded lifecycle, from listing to "
                "confirmed return.",
                "Three roles on one account (renter, owner, administrator); a booking that moves through "
                "named states and cannot skip any of them; an evidence record at both handovers; one "
                "codebase serving a laptop, a tablet and a phone browser.",
            ],
            [
                "O2",
                "To design and implement a security deposit mechanism, backed by a sealed record of both "
                "handovers and a time-bounded claim process, so that a damaged or missing item can be "
                "settled on evidence rather than on argument.",
                "Deposit collected with the rent and held by the platform; a 48-hour window for its "
                "return; 7 days to file a claim and 7 days to answer one; an administrator ruling with a "
                "written reason; settlement figures that are calculated and cannot be typed in by hand.",
            ],
            [
                "O3",
                "To develop and evaluate a computed trust score, combined with a deposit-based access "
                "ladder, that measures how reliably a member keeps their side of a rental.",
                "A score built from four weighted components rather than the star average; badges that are "
                "earned and never negative; high-value listings gated behind identity or a real track "
                "record; an evaluation against labelled history, where the label is what happened.",
            ],
        ],
        widths=[1.0, 6.2, 8.3],
        label="Table 2",
        caption_text="The three objectives and what each one commits us to. Each begins with an infinitive "
        "verb and maps back to one sentence of the problem statement.",
        align_center_cols=(0,),
    )

    para(
        doc,
        "The three are not equally measurable, and it is worth saying where. O1 can be checked by "
        "inspection: a rental either skips a state or it does not, and a test can prove which. O2 is "
        "arithmetic: any booking has one correct set of settlement figures. O3 is the hard one, because "
        "\"reliably\" is a claim about a population, and a population is what a project this size "
        "struggles to produce. Section 4.2 says how much evidence we expect and what we will say if we get "
        "less.",
    )

    doc.add_paragraph("3.3  Research Questions", style="Heading 2")
    para(
        doc,
        "Each objective is restated below as a question, in the same order. The restatement is not "
        "cosmetic: an objective is something we intend to do, and a question is something whose answer "
        "could turn out to be no. Writing them this way forced us to name, in the third column, the "
        "evidence that would settle each one.",
    )

    table(
        doc,
        ["Objective", "Research question", "Evidence that would answer it"],
        [
            [
                "O1",
                "**RQ1.** How can a web-based peer-to-peer rental platform be designed so that every "
                "rental — listing, approval, payment, pickup, return and review — is recorded and cannot "
                "skip a step?",
                "A working platform plus a systematic test: every transition the rules forbid is attempted "
                "and refused, and every permitted one produces exactly one record. The answer is a "
                "demonstration, not a statistic.",
            ],
            [
                "O2",
                "**RQ2.** How can a security deposit, a sealed handover record and a time-bounded claim "
                "process be combined so that an owner is protected when an item comes back damaged, while "
                "the renter's money is not held indefinitely?",
                "Six worked deposit scenarios — clean return, late return, minor damage, major damage, item "
                "not returned, disputed and unanswered — each producing the expected settlement figures, "
                "plus a check that no scenario leaves money held past its stated window.",
            ],
            [
                "O3",
                "**RQ3.** How accurately can a trust score computed from completed rentals, review "
                "ratings, cancellations and identity verification separate members who complete rentals "
                "from those who do not?",
                "A labelled set of member histories, scored once and judged on how well the score ranks "
                "reliable members above unreliable ones; reported as a discrimination measure at the "
                "published threshold, with the sample size stated beside it.",
            ],
        ],
        widths=[1.4, 7.4, 6.7],
        label="Table 3",
        caption_text="Objectives and research questions in strict one-to-one order, with the evidence each "
        "question expects. RQ1 is answered by demonstration, RQ2 by arithmetic, RQ3 by measurement.",
        align_center_cols=(0,),
    )

    para(
        doc,
        "The third column is the reason for the wording. RQ3 could easily have read \"is the trust score "
        "effective?\", which sounds fine and can never be answered, because no result would contradict it. "
        "Asking instead how accurately the score separates two groups forces a number, a threshold and a "
        "sample size onto the table — including the uncomfortable possibility that the sample is too small "
        "to say anything firm.",
    )

    doc.add_paragraph("3.4  Reflection on Heading 3", style="Heading 2")
    reflection(
        doc,
        "The problem matters to us for a reason we can point at: between the five of us we own a DSLR, a "
        "drill set, two projectors and a bridal outfit that together have been used maybe fifteen times in "
        "three years, and not one of us would have handed any of it to a stranger from an online group "
        "without a deposit and a CNIC copy. We are least confident about Objective 3, because a trust "
        "score only means something once the platform has finished rentals behind it, and finished rentals "
        "are the one thing we cannot write ourselves. For the same reason RQ3 is the question that depends "
        "most on access we do not yet have confirmed: we can build and unit-test the formula today, but "
        "measuring whether it separates reliable members from unreliable ones needs real transaction "
        "history that we will only have if the pilot is taken up.",
        extra=(
            "We are also aware that Objective 2 is a compromise rather than a clean solution. Fully "
            "automated escrow is not available to a student project in Pakistan, so the platform holds the "
            "money itself and an administrator stands in the middle of every dispute; that is a design "
            "with a human bottleneck in it, and we have written it down as a limitation instead of hiding "
            "it. Our plan for FYDP-2 is to publish the rules and the formulas, run them against whatever "
            "history the pilot produces, and report the limits of that evidence plainly rather than "
            "claiming more than we measured."
        ),
    )


# --------------------------------------------------------------------------
# heading 4


def section_four_intro(doc):
    doc.add_paragraph("4.  Proposed Methodology", style="Heading 1")

    para(
        doc,
        "This heading carries the four sub-sections the department asks for: the architecture, the data, "
        "the core technique, and the matrix that ties the first three back to Section 3. Each is built the "
        "same way — state the decision, explain it, then show the row of Chapter 2 that justifies it — and "
        "each ends with a reflection. The platform is at pre-release on a staging deployment, so most of "
        "what follows describes something that exists; where it does not, the text says so.",
    )

    doc.add_paragraph("Opening: what the methodology has to deliver", style="Heading 2")
    para(
        doc,
        "Based on our research questions, the methodology we propose must do three connected things. It "
        "must record a rental as a sequence of states that cannot be skipped (RQ1); it must decide who "
        "bears the loss when an item comes back damaged, inside a time limit stated in advance (RQ2); and "
        "it must turn ordinary transaction history into a signal a stranger can read before agreeing to a "
        "rental (RQ3). Our literature review shows that this combination is where the prior work stops "
        "short. Kotiyal et al. (2026) and Ahmadi et al. (2023) both model the deposit as programmatic "
        "states running beside the rental, but neither pairs it with a photographic record of the two "
        "handovers, which is the only evidence a physical item can produce. Papanastasiou et al. (2023) "
        "shows that raw reviews cannot carry reputation alone, and Stevens and Bossauer (2020) shows "
        "users trust a calculated signal more than the reviews it was calculated from, but neither turns "
        "that signal into a rule about who may rent something expensive. Salgado et al. (2026) supplies "
        "the multidimensional view of trust both of them lack.",
    )
    para(
        doc,
        "Table 4 is the bridge between those questions and the four sub-sections that follow. Nothing in "
        "4.1 to 4.3 was added without appearing in it first.",
    )

    table(
        doc,
        ["Question", "What the methodology must therefore contain", "Answered in"],
        [
            [
                "RQ1",
                "One authoritative record of the rental, written on the server, with named states, a guard "
                "on every transition, and no path that lets a party skip or reverse a step.",
                "4.1, 4.3.1",
            ],
            [
                "RQ2",
                "A deposit that is a first-class record rather than a number on a booking; two sealed "
                "handover records with photographic evidence; a claim process with published windows; and "
                "settlement figures that are derived rather than entered.",
                "4.2, 4.3.2",
            ],
            [
                "RQ3",
                "A scoring rule whose inputs are ordinary transaction history, a published weight for each "
                "input, a stated minimum below which no score is produced, and an evaluation that can fail.",
                "4.2, 4.3.3",
            ],
        ],
        widths=[2.0, 9.4, 4.1],
        label="Table 4",
        caption_text="What the methodology has to deliver, question by question. This is the required "
        "bridge paragraph expanded into a checkable table.",
        align_center_cols=(0,),
    )


def section_41(doc):
    doc.add_paragraph("4.1  System Architecture", style="Heading 2")

    para(
        doc,
        "The architecture is layered, and the layering is the point rather than a drawing convention. "
        "Everything that decides something — may this booking move to ACTIVE, is this deposit overdue, what "
        "badge does this member hold — sits in a layer with no database connection and no session, and is "
        "therefore testable on its own. Everything that touches the database sits below it, and everything "
        "a person clicks sits above. When a rule is wrong we can fix it and prove it fixed without "
        "standing up the whole application.",
    )

    doc.add_paragraph("4.1.1  What the architecture has to carry", style="Heading 3")
    lead_para(
        doc,
        "One record of the rental.",
        "A booking is the only place where the state of an agreement lives. Photographs, payments, claims, "
        "messages and reviews all hang off it, so any later question — what condition did it leave in, who "
        "agreed to what, when did the money move — is answered by reading rows attached to one booking.",
    )
    lead_para(
        doc,
        "Evidence written once.",
        "A handover record is sealed when it is created: condition, notes and photographs cannot be edited "
        "afterwards, and a second record of the same kind for the same booking is refused by the database. "
        "A record its author can revise later is not evidence, and the moment it would matter is exactly "
        "the moment they would want to revise it.",
    )
    lead_para(
        doc,
        "Decisions that leave a trace.",
        "Anything an administrator does — verifying an identity, rejecting a payment, ruling on a claim, "
        "settling a booking — writes a row with the actor, the reason and the time. The platform's answer "
        "to \"why was this deposit withheld\" has to be a record, not a memory.",
    )

    doc.add_paragraph("4.1.2  The layers", style="Heading 3")
    para(
        doc,
        "Figure 1 shows the six layers and what sits in each. The arrows run downward because that is the "
        "direction of a request: a person acts in the browser, the application layer authenticates and "
        "validates, the domain layer decides, the services carry out side effects, and the data layer "
        "stores the result.",
    )
    figure(
        doc,
        os.path.join(HERE, "architecture.png"),
        14.6,
        "Figure 1",
        "Layered architecture of the proposed platform. Each band is a deployment or testing boundary, not "
        "just a grouping: the domain rule layer has no database access and is covered by unit tests "
        "(drawn by the group).",
    )

    para(
        doc,
        "**User layer.** Three roles share one account model. A renter browses and books; an owner lists, "
        "approves and hands over; an administrator verifies identities, moderates listings and settles "
        "disputes. These are not separate applications — most people are both renter and owner — and "
        "Chapter 2's finding that trust must be modelled per role is why ratings and trust components are "
        "kept separate by direction (Salgado et al., 2026).",
    )
    para(
        doc,
        "**Client layer.** A responsive interface, rendered on the server. The same pages serve a laptop "
        "in an office and a phone in a rickshaw; the layout adapts and there is nothing to install. This "
        "departs from the original plan, and the reason is in Appendix B.",
    )
    para(
        doc,
        "**Application layer.** Every write is a server action, and every server action runs the same five "
        "checks in the same order: is there a session, does the role permit this, is the caller a party to "
        "this record, does the current status allow the change, and does the payload pass its schema. Rate "
        "limits sit on top for the actions that cost something. One order means a missing check is a bug "
        "in one place rather than a flaw repeated six times.",
    )
    para(
        doc,
        "**Domain rule layer.** The six rule modules — booking lifecycle, trust score, access tiers, "
        "deposit state, claim rules, handover protocol — plus review release. Each is pure: values in, "
        "decision out, no Prisma, no session, no clock read from the environment. This is the layer that "
        "answers RQ1 and RQ2, and the one with the thickest test coverage, because a rule that cannot be "
        "tested in isolation tends not to be tested at all.",
    )
    para(
        doc,
        "**Platform services.** Payments sit behind a provider-agnostic interface, so booking code does "
        "not know whether money arrived by transfer, card or wallet. Only the manual path is live today: "
        "the renter transfers the rent and deposit, uploads the reference, and an administrator verifies "
        "receipt. JazzCash, Easypaisa and Safepay are defined in the same interface and are the next "
        "integration (Malik et al., 2025). Realtime messaging carries chat, price offers and notifications; "
        "scheduled jobs expire stale bookings, escalate unanswered claims and release withheld reviews.",
    )
    para(
        doc,
        "**Data layer.** One PostgreSQL database, twenty-seven tables and twenty-four migrations, reached "
        "through Prisma. Money is stored as integers in whole rupees and every state change runs inside a "
        "transaction, so a booking cannot be half-moved. Photographs of items, handovers and claims live in "
        "Cloudinary, and each stored file carries a checksum recorded with its row — a photograph reused "
        "from another rental is refused at write time. Identity documents are stored separately and are "
        "never publicly addressable.",
    )

    doc.add_paragraph("4.1.3  The path of one rental", style="Heading 3")
    para(
        doc,
        "Figure 2 follows the same objects through time rather than through layers, which is the view an "
        "examiner usually wants. Read the sequence below against it.",
    )
    figure(
        doc,
        os.path.join(HERE, "deposit_flow.png"),
        14.6,
        "Figure 2",
        "The deposit, handover and claim path. The upper row is the normal rental; the lower row is the "
        "same rental when the item comes back damaged, rejoining at settlement (drawn by the group).",
    )
    numbered(
        doc,
        [
            "An owner publishes a listing: category, city, prices by day, week and month, a deposit they "
            "choose, and up to ten photographs.",
            "A renter requests dates. The booking is created as PENDING and the dates are held, so the item "
            "cannot be double-booked while the owner decides.",
            "The owner approves or declines. Approval copies the prices from the listing at that moment; "
            "they are not read again later.",
            "The renter pays rent and deposit to the platform, uploads the reference, and the booking moves "
            "to PAYMENT_PENDING.",
            "An administrator verifies that the money arrived. The commission rate in force at that "
            "instant is frozen onto the payment row.",
            "The owner records the pickup: condition, a short note, up to six photographs. The record is "
            "sealed, the booking becomes ACTIVE, and the renter is asked to agree or dispute — but is "
            "never blocked from proceeding if they say nothing.",
            "The owner records the return the same way. The booking becomes COMPLETED and the 48-hour "
            "deposit clock starts.",
            "Within seven days the owner may file a claim against the deposit. The renter has seven days "
            "to accept or dispute it; silence escalates it to an administrator rather than accepting it.",
            "Both parties review each other. Reviews stay hidden until both are in, or until the "
            "fourteen-day window closes, and only then do they move the rating averages.",
            "An administrator settles: the owner is paid the rent less commission plus any upheld "
            "compensation, and the rest of the deposit returns to the renter.",
        ]
    )

    doc.add_paragraph("4.1.4  Decisions and the literature behind them", style="Heading 3")
    para(
        doc,
        "Table 5 lists the choices a reader could reasonably disagree with: the choice, the reason, the "
        "source in Chapter 2, and the question it serves. Choices that are simply conventional — HTTPS, "
        "password hashing, parameterised queries — are left out, because citing a source for them would be "
        "padding.",
    )

    table(
        doc,
        ["Decision", "What we chose", "Why, from Chapter 2", "Source", "RQ"],
        [
            [
                "One web application, not separate installed builds",
                "A single Next.js codebase, rendered on the server, used in any browser.",
                "Adoption here turns on screen-flow simplicity and latency far more than on the delivery "
                "platform; a browser build removes an install step and a store review from the critical "
                "path.",
                "Rana et al. (2022)",
                "RQ1",
            ],
            [
                "Server-side lifecycle guard",
                "Status changes happen only through one guarded server path, never from the client.",
                "Automated booking lifecycles remove friction, but state authority has to sit with the "
                "platform rather than with either party.",
                "Tele et al. (2026); Kotiyal et al. (2026)",
                "RQ1",
            ],
            [
                "Deposit as its own record",
                "A payment row carries rent and deposit separately, with its own status machine.",
                "Splitting the rental charge from the refundable deposit is what makes conditional release "
                "possible; both cited systems model them as distinct states.",
                "Kotiyal et al. (2026); Ahmadi et al. (2023)",
                "RQ2",
            ],
            [
                "Administrator, not on-chain arbitration",
                "A trained administrator rules on disputes; each ruling is logged with a reason.",
                "Administrative mediation is the mechanism users trust, and smart contracts are too rigid "
                "to judge physical damage nobody photographed.",
                "Waris and Kim (2023); Košťál et al. (2025)",
                "RQ2",
            ],
            [
                "One relational database",
                "PostgreSQL for every table, including listings.",
                "Evidence and money must commit together. A deposit state in one store and a claim in "
                "another is a reconciliation problem we would have invented for ourselves.",
                "Kotiyal et al. (2026)",
                "RQ2",
            ],
            [
                "Sealed double handover record",
                "A pickup and a return record, each with condition, notes, photographs and a checksum.",
                "Dispute prevention beats dispute resolution: a photograph taken at the door settles in one "
                "reading what a written argument never settles.",
                "Zhao and Chen (2024); Papanastasiou et al. (2023)",
                "RQ2",
            ],
            [
                "Computed score, not the star average",
                "Four weighted components, with volume, cancellation and identity in the calculation.",
                "A single averaged rating is mathematically weak and inflates; a calculated signal is both "
                "more informative and more trusted than the reviews behind it.",
                "Salgado et al. (2026); Stevens and Bossauer (2020)",
                "RQ3",
            ],
            [
                "Deposit-band access ladder",
                "Requirements rise with the deposit: confirmation at PKR 25,000, identity or history at "
                "PKR 100,000.",
                "Heavy verification at registration drives users away, while providers withhold valuable "
                "goods without it; gating by what is at stake satisfies both findings.",
                "Ferraz and Ferraz (2021); Li and Wang (2020)",
                "RQ3",
            ],
            [
                "Private identity-document store",
                "CNIC images go to a private store, are reviewed by an administrator, and are never served "
                "from a public URL.",
                "A government identity check is the accepted remedy for marketplace fraud, and a verified "
                "badge is what neutralises first-rental stranger risk.",
                "Ghozi (2022); Bird (2026)",
                "RQ3",
            ],
        ],
        widths=[3.5, 3.6, 4.9, 2.4, 1.1],
        label="Table 5",
        caption_text="Architectural decisions and their justification in Chapter 2. The last column is kept "
        "to RQ1, RQ2 or RQ3 for the same reason the traceability matrix is.",
        font_size=8.1,
        align_center_cols=(4,),
    )

    doc.add_paragraph("4.1.5  What we chose not to build", style="Heading 3")
    lead_para(
        doc,
        "No blockchain or smart contract.",
        "Tele et al. (2026) and Ahmadi et al. (2023) make a good case for cryptographic escrow, and Košťál "
        "et al. (2025) show what it costs: latency that makes a mainstream marketplace unusable, and no "
        "handling for the real dispute, which is a scratched lens rather than a failed transaction. A "
        "wallet would also have added a fifth registration step to a market where the fourth already hurts.",
    )
    lead_para(
        doc,
        "No installable phone build in this phase.",
        "The same interface already works in a phone browser, so a separate build would double the release "
        "surface without changing what a user can do. The architecture does not obstruct it: every screen "
        "is served from server actions a native client could call.",
    )
    lead_para(
        doc,
        "No GPS tagging and no insurance.",
        "Hardware tagging is a logistics problem, not a software one, and insurance here is priced per item "
        "in a way a student platform cannot negotiate. Chapter 2's nearest analogue, Fat Llama, sells "
        "insurance as its differentiator; we are explicit that we do not, and the deposit stands in its "
        "place.",
    )

    doc.add_paragraph("4.1.6  Reflection on the architecture", style="Heading 2")
    reflection(
        doc,
        "The component we are least confident about is the set of scheduled jobs, because they are the "
        "only part of the design that can fail without anybody noticing. The deposit clock, the claim "
        "escalation and the review release all assume a job runs at the right hour; if the schedule "
        "silently stops, a renter's deposit sits past its window and no error is raised anywhere. Our "
        "fallback is partly in place already: every one of those deadlines is also checked whenever the "
        "relevant page is opened, so a missed schedule delays an action rather than losing it. If it "
        "proves unreliable in the pilot we will move the checks fully into the read path and keep the "
        "scheduled job only as an optimisation.",
        extra=(
            "The question this architecture serves most directly is RQ1, since the lifecycle guard is the "
            "spine of the system. RQ2 is served nearly as well — the evidence chain and the deposit state "
            "machine sit on the same booking record — but RQ3 is served only indirectly, through the "
            "history the lifecycle produces. That asymmetry is real: starting again, we would separate the "
            "trust service from the booking transaction so it could be rebuilt without touching a money "
            "path."
        ),
    )


def section_42(doc):
    doc.add_paragraph("4.2  Dataset and Data Collection Strategy", style="Heading 2")

    para(
        doc,
        "A software project of this kind does not have a dataset in the sense a machine-learning project "
        "does. There is no corpus to download and no benchmark to beat. What it has is a set of records the "
        "platform itself produces as people use it, a small reference taxonomy, and the scenario set we "
        "build in order to test the rules. This sub-section describes all three, including how much of the "
        "third we realistically expect.",
    )

    doc.add_paragraph("4.2.1  The data the platform holds", style="Heading 3")
    table(
        doc,
        ["Data", "Where it comes from", "When captured", "Serves", "Expected volume"],
        [
            [
                "Category and city taxonomy",
                "Written by the group and seeded into every environment, because no listing can exist "
                "without a category.",
                "Before launch",
                "RQ1",
                "7 categories, 30 subcategories, 3 cities",
            ],
            [
                "Listings",
                "Owners: title, description, condition, prices, deposit, up to 10 photographs, "
                "availability.",
                "At listing time",
                "RQ1, RQ3",
                "60–100 in the pilot",
            ],
            [
                "Identity documents",
                "Members: front and back CNIC images, reviewed by an administrator.",
                "On registration",
                "RQ3",
                "One pair per verified member",
            ],
            [
                "Bookings and payments",
                "Generated by the lifecycle: dates, amounts, deposit, provider, reference, frozen "
                "commission rate.",
                "During every rental",
                "RQ1, RQ2",
                "30–50 complete rentals",
            ],
            [
                "Handover evidence",
                "Both parties: condition, notes, up to 6 photographs per handover, checksums.",
                "At pickup and return",
                "RQ2",
                "2 records and up to 12 images per rental",
            ],
            [
                "Claims and rulings",
                "Owners file, renters answer, administrators rule with a written reason.",
                "Within 7 days of completion",
                "RQ2",
                "5–10, most resolved by agreement",
            ],
            [
                "Reviews",
                "Both directions, released reciprocally inside a 14-day window.",
                "After completion",
                "RQ3",
                "Up to 2 per rental",
            ],
            [
                "Behavioural history",
                "Completions, self-cancellations and lateness — derived, never asked for.",
                "Continuously",
                "RQ3",
                "One row per member",
            ],
        ],
        widths=[3.2, 4.5, 3.4, 1.5, 2.9],
        label="Table 6",
        caption_text="The data the platform holds, and the question each element serves. Nothing is "
        "collected for its own sake; every row exists because a rule in Section 4.3 reads it.",
        font_size=8.1,
        align_center_cols=(3,),
    )

    doc.add_paragraph("4.2.2  Where the data comes from", style="Heading 3")
    para(
        doc,
        "There are three sources, and which is which matters, because they carry different levels of "
        "confidence.",
    )
    lead_para(
        doc,
        "Reference data, which we already have.",
        "The category tree and the launch cities are written by us and seeded before launch, so this part "
        "is complete today rather than pending. It is small on purpose: seven categories with thirty "
        "subcategories between them, and three cities. Ragavarthini et al. (2025) found that community "
        "rental platforms live or die on local density, so a narrow taxonomy that matches how people "
        "describe the thing they want to borrow beats a broad one that sorts everything into \"other\".",
    )
    lead_para(
        doc,
        "Primary data, which the pilot produces.",
        "Every registration, listing, booking, handover, claim and review is collected through ordinary use "
        "of the platform, with consent taken at registration. This is the only source that can answer RQ3, "
        "because a trust score is a statement about behaviour and behaviour cannot be simulated into "
        "existence. It is also the source we control least, which is the subject of 4.2.3 and 4.2.6.",
    )
    lead_para(
        doc,
        "The scenario set, which we build.",
        "Six scripted rentals run through the platform in a test environment: a clean return, a late "
        "return, minor damage, major damage, an item that never comes back, and a disputed claim nobody "
        "answers. Each produces a full set of records — both handovers, a payment, a claim, a ruling and a "
        "settlement — and each has a hand-calculated expected outcome. This is the evidence base for RQ2, "
        "and it is complete on the day we build it.",
    )

    doc.add_paragraph("4.2.3  How much data, and why that number is small", style="Heading 3")
    para(
        doc,
        "It would be easy, and dishonest, to write that the platform will generate \"thousands of "
        "transactions\". A final-year project launching in three cities with no marketing budget produces "
        "tens, not thousands. Our planning assumption is thirty to fifty completed rentals over the pilot, "
        "drawn from the five of us, classmates, family and word of mouth, with every rental contributing "
        "two sealed handover records, a payment, two reviews and a member history.",
    )
    para(
        doc,
        "That figure is comfortable for RQ1 and RQ2 and uncomfortable for RQ3, and the reason is worth "
        "stating precisely. The trust score is not a trained model. It has no parameters that learn from "
        "data: the four weights and the two thresholds are fixed in advance from the literature, and the "
        "code is a pure function of six inputs. What RQ3 needs is therefore not a training set but an "
        "evaluation set — member histories already labelled by what happened, on which we can ask whether "
        "the score ranks reliable members above unreliable ones. Thirty to fifty rentals yields between "
        "thirty and eighty labelled cases, enough to describe behaviour and too few to support a confident "
        "general claim. We will report it as a case study with its sample size attached, not as a "
        "validation study.",
    )
    para(
        doc,
        "To get at the behaviour of the formula itself — whether it is stable, whether it rewards the "
        "right things, where the thresholds bite — we will also run it over a simulated population of two "
        "thousand members, with behaviour drawn from distributions stated in the report and known ground "
        "truth. That is a sensitivity analysis of a formula, not evidence about Pakistani renters, and it "
        "will be labelled as such wherever it appears. Presenting simulated results as if they were "
        "observed is the easiest way to fail the technical soundness criterion.",
    )

    doc.add_paragraph("4.2.4  Labels and ground truth", style="Heading 3")
    para(
        doc,
        "For RQ3 the label on each case is what happened, not what anyone thought of the member. A "
        "completed rental with no upheld claim is a positive case; a rental the member cancelled "
        "themselves, or one that ended in an upheld claim against them, is a negative one. Using events "
        "rather than opinions removes the circularity of grading a reputation score against reputations. "
        "For RQ2 the labels are arithmetic: every scenario has an expected settlement calculated by hand "
        "before the test runs, a second reviewer who did not write the scenario checks the calculation, and "
        "the test passes only if the system produces that split exactly.",
    )

    doc.add_paragraph("4.2.5  Quality control", style="Heading 3")
    lead_para(
        doc,
        "The database refuses bad states.",
        "One handover per type per booking, one payment per booking, one claim per booking, and a trigger "
        "that refuses a payment whose amounts no longer match the booking's terms. These are constraints "
        "rather than application checks, because a constraint cannot be forgotten by whoever writes the "
        "next action.",
    )
    lead_para(
        doc,
        "Money is stored in whole rupees.",
        "Every amount is an integer and the settlement function throws on a fractional value. Rounding is "
        "not a formatting detail when the remainder has to land somewhere, so the split is calculated such "
        "that payout, commission and refund always add back to the amount collected.",
    )
    lead_para(
        doc,
        "Evidence is checksummed and sealed.",
        "Each photograph carries the storage provider's checksum, recorded with the row; a file already "
        "attached to another booking is refused. That closes the cheapest way to manufacture evidence — "
        "reusing a picture of an undamaged item from an earlier rental — and no code path updates a "
        "handover after it is written.",
    )
    lead_para(
        doc,
        "Personal data stays private.",
        "CNIC images are never served from a public URL; they are read through an authenticated route that "
        "checks for the administrator role. A cleanup job removes abandoned uploads and a member can ask "
        "for their documents to be deleted.",
    )
    lead_para(
        doc,
        "The rules are covered by tests.",
        "Fifty-two test files sit beside the code, concentrated on the pure rule modules and written "
        "against properties where possible — that no combination of inputs can produce a negative trust "
        "badge, and that a gate can always be cleared by somebody.",
    )

    doc.add_paragraph("4.2.6  Risks and fallbacks", style="Heading 3")
    para(
        doc,
        "Four things can go wrong with this data plan, and pretending otherwise would be the weakest part "
        "of the chapter. Each risk has a fallback we can actually execute.",
    )
    table(
        doc,
        ["Risk", "Likely?", "Effect if it happens", "Fallback"],
        [
            [
                "The pilot produces fewer than twenty completed rentals.",
                "Moderate",
                "RQ3 cannot be answered on real data; the score can be described but not measured.",
                "Report the formula's behaviour on the simulated population, present the real cases as a "
                "case study, and state plainly that the sample is too small for a general claim.",
            ],
            [
                "Pilot users decline to upload CNIC images.",
                "Moderate",
                "The identity component is empty for most members and the high-value tier is unreachable.",
                "The CNIC gate is a configuration switch and turns off where documents cannot be stored. "
                "The top tier already accepts three completed rentals in place of verification.",
            ],
            [
                "Wallet payments cannot be verified automatically.",
                "High",
                "Every payment needs a human to check a reference, which caps how fast the pilot grows.",
                "This is the current design rather than a risk to it: manual verification with an audit "
                "trail is live, and the payment interface keeps wallet providers behind one boundary.",
            ],
            [
                "Photographs of items or people raise a privacy objection.",
                "Low",
                "Members refuse to document handovers, removing the evidence RQ2 depends on.",
                "Item-only photography, no faces, consent text at upload and a retention rule. Where a "
                "member refuses, the record still captures condition and notes.",
            ],
            [
                "A suspended member's history skews the score.",
                "Low",
                "An account that has been stood down keeps contributing to the aggregates it earned.",
                "Suspension is recorded as an administrative action with a reason, the score is recomputed "
                "from live rows only, and the audit trail keeps the case reviewable.",
            ],
        ],
        widths=[4.3, 1.7, 4.3, 5.2],
        label="Table 7",
        caption_text="Risks to the data plan and the fallback for each. The likelihood column is our own "
        "judgement and is meant to be argued with.",
        font_size=8.2,
        align_center_cols=(1,),
    )

    para(
        doc,
        "One thing we deliberately did not do is scrape listings from OLX or Facebook groups to inflate "
        "the dataset. Those are sales rather than rentals, so they do not carry the fields the score reads; "
        "the terms of both platforms forbid it; and a dataset assembled that way would answer a different "
        "question while looking like it answered ours.",
    )

    doc.add_paragraph("4.2.7  Reflection on the data plan", style="Heading 2")
    reflection(
        doc,
        "We have confirmed access to the reference taxonomy and to the scenario set, because both are "
        "ours. We have not confirmed the volume of real transaction history, and we are not going to write "
        "that we have: the pilot depends on people outside the group deciding to rent a camera from a "
        "classmate's website, and no amount of planning makes that certain. Our fallback is the simulated "
        "population described above, used only to show how the formula behaves and never presented as "
        "evidence about real users.",
        extra=(
            "The other thing worth admitting is that our data plan got smaller as the project went on. The "
            "proposal promised an optional fraud-detection model trained on user behaviour; with thirty to "
            "fifty rentals there is nothing to train it on, and a model fitted to that much data would "
            "report an accuracy figure that meant nothing. Dropping it was the right call, but it cost us "
            "a feature we had already announced, and that is recorded in Appendix B rather than quietly "
            "removed."
        ),
    )


def section_43(doc):
    doc.add_paragraph("4.3  Proposed Algorithm / Core Technique", style="Heading 2")

    para(
        doc,
        "Three techniques carry the project, and they layer on each other rather than sitting side by "
        "side. The lifecycle decides what happened and when; the evidence chain decides what the item "
        "looked like at each of those moments and turns that into money; the trust score reads the history "
        "the first two produce and turns it into a judgement a stranger can use. None of the three is a "
        "novel algorithm in the machine-learning sense. The contribution is in how they are joined, and in "
        "the rules invented to make the join safe.",
    )

    doc.add_paragraph("4.3.1  Technique one: the guarded rental lifecycle", style="Heading 3")
    para(
        doc,
        "A booking moves through six states — PENDING, APPROVED, PAYMENT_PENDING, ACTIVE, COMPLETED, "
        "REVIEWED — with three terminal side-states: DECLINED, CANCELLED, EXPIRED. Each transition has "
        "exactly one party entitled to make it and exactly one state it may be made from. The owner "
        "approves; the renter pays; the owner records the pickup; the owner records the return; both "
        "review.",
    )
    para(
        doc,
        "The guard is a conditional update rather than a read-then-write. The status the caller believes "
        "the booking is in becomes part of the WHERE clause, so if the row has moved in between — because "
        "the other party acted, or because the same user double-clicked — the update matches nothing and "
        "the action reports a concurrent change instead of silently overwriting. This matters more than it "
        "sounds: several of these transitions move money, and \"apply the change twice\" is how a deposit "
        "is held twice or released twice.",
    )
    para(
        doc,
        "The lifecycle is the substrate for everything else. A handover record can only be filed against "
        "the transition it belongs to, so a return record cannot be attached to a booking that never "
        "started — the cheapest way to manufacture evidence about an item that was never collected. A "
        "review can only follow a completed rental. A claim can only be filed while the deposit is still "
        "held.",
    )
    para(
        doc,
        "**Where it comes from.** Kotiyal et al. (2026) specify the rental and the deposit as separate "
        "state flows and are the direct blueprint for the split; Tele et al. (2026) show that an automated "
        "lifecycle removes friction, with booking states that map closely onto ours. **What we changed.** "
        "Both put the authoritative state somewhere the parties cannot edit — a smart contract, in their "
        "case. We put it in a relational row with a locking read, for the reasons Košťál et al. (2025) "
        "give: the disputes we expect are about a scratched lens, which no contract can adjudicate, and "
        "chain latency would be felt on every screen.",
    )

    doc.add_paragraph("4.3.2  Technique two: the evidence chain and the deposit", style="Heading 3")
    para(
        doc,
        "This is the part that answers RQ2, and it has four moving parts: the handover record, the deposit "
        "clock, the claim process, and the settlement arithmetic.",
    )
    lead_para(
        doc,
        "The handover record.",
        "At each of the two moments an item changes hands, the person handing it over records the condition "
        "against what was expected — as expected, minor wear, or damaged — a note of up to a thousand "
        "characters, and up to six photographs. The record is sealed on write and a second record of the "
        "same type is refused by a unique constraint. The other party is asked to agree or dispute and "
        "their answer is stored, but their silence never blocks the rental: a confirmation that can freeze "
        "somebody else's deposit is a weapon, not a safeguard.",
    )
    lead_para(
        doc,
        "The deposit clock.",
        "The deposit has six states: none, not due, returned, due, overdue and claimed. The clock starts "
        "when the rental completes and runs for forty-eight hours — long enough for a bank transfer between "
        "two people, short enough that a renter is not left guessing. \"Returned but unconfirmed\" is one "
        "state with a nullable timestamp rather than two, because a renter who has not opened the app is "
        "the normal case and must not read as a complaint.",
    )
    lead_para(
        doc,
        "The claim.",
        "A claim is bounded by the deposit, because that is the only obligation the platform has standing "
        "to describe. It may be filed within seven days of completion — damage is usually found when the "
        "item is next used, not at the door — and the renter has seven days to accept or dispute it. The "
        "same constant serves both windows on purpose: a pause longer than the response window would hold a "
        "deposit after the renter's chance to speak had passed, and a shorter one would restart the clock "
        "while they were still entitled to answer. Silence is never acceptance; an unanswered claim "
        "escalates to an administrator.",
    )
    lead_para(
        doc,
        "The settlement.",
        "Three figures are derived and none is accepted from the administrator. The commission rate frozen "
        "onto the payment at verification is applied to the rental; any upheld claim is taken out of the "
        "deposit and added to the owner's payout; the rest of the deposit returns to the renter. The "
        "administrator decides that a booking settles, never for how much. An administrator who believes a "
        "figure is wrong has to correct the record it derives from, which leaves a trace; typing a "
        "different number into a payout field would not.",
    )

    table(
        doc,
        ["Line", "Clean return, 0% commission", "Damage upheld, 7.5% commission"],
        [
            ["Rent collected from renter", "PKR 4,000", "PKR 4,000"],
            ["Deposit collected from renter", "PKR 30,000", "PKR 30,000"],
            ["Commission rate applied", "0% (launch rate)", "7.5% (frozen at verification)"],
            ["Platform commission", "PKR 0", "PKR 300"],
            ["Owner's share of the rent", "PKR 4,000", "PKR 3,700"],
            ["Damage compensation upheld", "PKR 0", "PKR 6,000"],
            ["**Owner payout**", "**PKR 4,000**", "**PKR 9,700**"],
            ["**Deposit returned to renter**", "**PKR 30,000**", "**PKR 24,000**"],
            ["**Check: payout + commission + refund**", "**PKR 34,000 = collected**", "**PKR 34,000 = collected**"],
        ],
        widths=[6.3, 4.6, 4.6],
        label="Table 8",
        caption_text="Settlement arithmetic for one booking, in the two cases that matter. The bottom row is "
        "the invariant the code asserts: the three outputs always add back to the amount collected.",
        font_size=8.6,
    )

    para(
        doc,
        "That last row is why the arithmetic is a function rather than a form. A settlement that pays out "
        "more than was collected is not a rounding annoyance; it is a hole in the ledger, and in a system "
        "where a human verifies every payment by hand, a hole is exactly what nobody notices until it is "
        "large.",
    )
    para(
        doc,
        "**Where it comes from.** Kotiyal et al. (2026) supply the split between rental charge and "
        "refundable deposit; Ahmadi et al. (2023) contribute the deterministic payment states and the idea "
        "that a refund is triggered by an event rather than by a person; Papanastasiou et al. (2023) is why "
        "reviews and deposit release stay on separate tracks, since a review that can be traded against a "
        "deposit is the blackmail they model; Zhao and Chen (2024) is why the record is photographic; and "
        "Waris and Kim (2023) is why a human administrator makes the final call. **What we changed.** None "
        "of the cited works pairs a sealed, checksummed photographic record with the deposit state machine. "
        "Kotiyal models the states without evidence; Zhao argues for evidence without modelling the money. "
        "Joining them is the contribution this project claims, and it is a small one stated precisely "
        "rather than a large one stated vaguely.",
    )

    doc.add_paragraph("4.3.3  Technique three: the trust score and the access ladder", style="Heading 3")
    para(
        doc,
        "The trust score exists because a star average does not answer the question a renter actually has, "
        "which is not \"was this person nice\" but \"is this person going to bring my camera back\". A "
        "score that is mostly the star average restates the rating in a second unit and adds nothing but "
        "false precision. Ours combines four components, each of which a rating cannot carry.",
    )
    lead_para(
        doc,
        "Reputation (weight 0.45).",
        "Ratings received in both directions, weighted by how many there are on each side, then shrunk "
        "towards a neutral prior of 3.5 with a weight of five imaginary reviews. Shrinkage stops one "
        "perfect five-star review from outranking fifty strong ones, and 3.5 rather than 3 is chosen "
        "because a genuinely average rental here earns four stars; a prior of 3 would drag every honest new "
        "record downwards and read as a penalty for being new.",
    )
    lead_para(
        doc,
        "Experience (weight 0.25).",
        "Completed rentals divided by ten, capped at one. It saturates on purpose: the difference between "
        "two rentals and ten is large, between forty and fifty is nothing, and a linear scale would let "
        "trust be bought with volume on a marketplace where an owner can list ten cheap items.",
    )
    lead_para(
        doc,
        "Reliability (weight 0.20).",
        "The share of finished business the member did not cancel themselves. Being cancelled on is not "
        "evidence about you, so it is not counted, and with no history at all this component is one rather "
        "than zero — starting everyone at zero would punish inexperience twice, once here and once through "
        "experience.",
    )
    lead_para(
        doc,
        "Verification (weight 0.10).",
        "One for a verified identity, 0.4 for a confirmed email address, zero otherwise. The smallest "
        "weight of the four, because an identity says who somebody is rather than how they behave; its "
        "real force is a gate rather than a weight, as the bands below show.",
    )
    para(
        doc,
        "Two rules sit on top of the arithmetic, and both protect people rather than rank them. A member "
        "with no completed rental is not scored at all: the score is empty and the interface says \"new to "
        "SamaanShare\", because \"not yet established\" and \"established as unreliable\" are opposite "
        "claims and a number cannot say the first. And there are no negative badges: below 0.60 a member "
        "simply has no badge, while the ratings and the rental count stay on the page. A published \"low "
        "trust\" label assembled from proxies is an accusation, and a cancelled booking might have been the "
        "other party's fault.",
    )
    para(
        doc,
        "Bands are trusted at 0.60 and highly trusted at 0.80, with the top band additionally requiring a "
        "verified identity. That requirement is a condition rather than a heavier weight because "
        "arithmetic would not achieve it: the other three components carry ninety per cent of the weight, "
        "so a flawless unverified record reaches about 0.94 and clears any threshold worth setting. Saying "
        "it as a condition says what is meant.",
    )

    table(
        doc,
        ["Component", "Member A — established", "Member B — one rental"],
        [
            ["Ratings received", "4.8 from 6 owners, 4.5 from 2 renters", "5.0 from 1 owner"],
            ["Reputation (0.45)", "4.725 weighted, 4.254 after shrinkage → 0.813", "3.750 after shrinkage → 0.688"],
            ["Experience (0.25)", "6 completed ÷ 10 → 0.600", "1 completed ÷ 10 → 0.100"],
            ["Reliability (0.20)", "6 of 7 finished → 0.857", "1 of 1 finished → 1.000"],
            ["Verification (0.10)", "identity verified → 1.000", "email only → 0.400"],
            ["**Score**", "**0.787**", "**0.574**"],
            ["**Band**", "**Trusted** (0.60 or above, short of 0.80)", "**No badge**"],
            ["Access to a PKR 100,000-deposit listing", "Allowed — verified identity", "Refused — no identity and fewer than 3 rentals"],
        ],
        widths=[4.5, 5.6, 5.4],
        label="Table 9",
        caption_text="The trust score worked through for two members. Member B has a perfect rating and still "
        "earns nothing, which is the property the design is protecting.",
        font_size=8.2,
    )

    para(
        doc,
        "The access ladder applies the same logic to a different question: not \"how trustworthy is this "
        "person\" but \"how much should we ask of them before they take something valuable\". Listings with "
        "no deposit are open to any member with a confirmed email address. From PKR 25,000 to PKR 99,999 "
        "the same requirement holds, and the tier is kept in the code as the place to add one. At PKR "
        "100,000 and above, a renter needs either a verified identity or three completed rentals.",
    )
    para(
        doc,
        "The \"or\" is the load-bearing part. Requiring verification alone would be stricter on paper and "
        "worse in practice, because identity verification is granted by an administrator out of band: at "
        "launch nobody has it, and every high-value listing would be unbookable by everyone. Three "
        "completed rentals is different evidence for the same thing — an account with a history it would "
        "lose — and it is measured against what the owner themselves declared at risk.",
    )
    para(
        doc,
        "**Where it comes from.** Salgado et al. (2026) is why the score is multidimensional rather than an "
        "average; Stevens and Bossauer (2020) is why a computed signal is worth publishing at all; Neifer "
        "et al. (2023) supplies the pattern of a reputation score built from handover events; Ferraz and "
        "Ferraz (2021) is why access is progressive rather than uniform; and Li and Wang (2020) is why "
        "requirements rise with the value at stake. **What we changed, and what it costs.** Salgado et al. "
        "measure competence, integrity and reliability with survey instruments. We cannot survey every "
        "rental, so we approximate those three dimensions with observable events: ratings in both "
        "directions, completions, self-cancellations and identity. The two components that would map most "
        "directly onto their dimensions — whether the item matched its listing, and whether the handover "
        "was punctual — are not in the score yet, because the platform does not record them as separate "
        "fields. That is the first thing we would add with a second year.",
    )

    doc.add_paragraph("4.3.4  How the three techniques will be evaluated", style="Heading 3")
    para(
        doc,
        "Evaluation proper belongs to FYDP-2, but the plan has to exist now, because a technique whose "
        "success cannot be recognised is not a technique.",
    )
    table(
        doc,
        ["RQ", "What is tested", "Measure", "Where it runs"],
        [
            [
                "RQ1",
                "Every transition the rules forbid is attempted and must be refused; every permitted one "
                "must produce exactly one record and one notification.",
                "Number of illegal transitions accepted (target zero) and permitted transitions that fail "
                "to record (target zero).",
                "Unit tests on the lifecycle rules, then end-to-end tests through the interface.",
            ],
            [
                "RQ2",
                "The six deposit scenarios, each run end to end from booking to settlement.",
                "Exact match between the hand-calculated settlement and the system's three output figures; "
                "no scenario holding money past a stated window.",
                "Scenario set in a staging database, with the calculation reviewed independently.",
            ],
            [
                "RQ3",
                "How well the score ranks members by outcome across the labelled histories.",
                "Discrimination at the 0.60 threshold — the probability that a completing member outscores "
                "a cancelling one — reported with its sample size, and the same measure on the simulated "
                "population reported separately.",
                "Offline evaluation over exported histories and the simulation harness.",
            ],
        ],
        widths=[1.3, 6.2, 5.3, 2.7],
        label="Table 10",
        caption_text="How each technique will be evaluated. Each measure can come out badly, which is the "
        "test of whether it is a measure at all.",
        font_size=8.2,
        align_center_cols=(0,),
    )

    doc.add_paragraph("4.3.5  Reflection on the technique", style="Heading 2")
    reflection(
        doc,
        "If the trust score turns out not to separate members at all — and with thirty to fifty rentals "
        "that is a genuinely possible result — our backup is to stop publishing a score and publish the "
        "counts instead: rentals completed, cancellations, reviews received, identity verified. That is a "
        "weaker product but a more honest one, and it is still grounded in the literature, since Stevens "
        "and Bossauer (2020) found users trust concrete verified facts at least as much as a composite "
        "number. The deposit ladder does not depend on the score, so it survives either way.",
        extra=(
            "The risk we think about more often is the opposite: that the score works well enough to be "
            "believed and is then gamed. Listing ten cheap items to accumulate completions is the obvious "
            "route, and saturating at ten only slows it down rather than closing it. We do not have a good "
            "answer yet. Our plan is to measure the pattern in the pilot data and, if it appears, to weight "
            "completions by the value of the item involved — which needs a field we do not currently store, "
            "so it is a FYDP-2 change rather than something we can claim now."
        ),
    )


def section_44(doc):
    doc.add_paragraph("4.4  Traceability Matrix", style="Heading 2")

    para(
        doc,
        "The matrix below has one row for every decision a reader could argue with in Sections 4.1 to 4.3. "
        "Each row names the decision, its type, the source in Chapter 2 that justifies it, and the research "
        "question it serves. The last column contains nothing but RQ1, RQ2 or RQ3, because a free-text "
        "answer there is how a matrix looks complete while telling nobody anything.",
    )

    rows = [
        ["Single responsive web application rather than separate installed builds", "Architecture", "Rana et al. (2022)", "RQ1"],
        ["Server-side lifecycle guard as the only path for a status change", "Architecture", "Tele et al. (2026); Kotiyal et al. (2026)", "RQ1"],
        ["One account model serving renter, owner and administrator roles", "Architecture", "Waris and Kim (2023)", "RQ1"],
        ["Centralised administrator panel rather than on-chain arbitration", "Architecture", "Waris and Kim (2023); Košťál et al. (2025)", "RQ2"],
        ["A single relational store so evidence and money commit together", "Architecture", "Kotiyal et al. (2026)", "RQ2"],
        ["Payment providers kept behind one interface, manual verification first", "Architecture", "Malik et al. (2025); Anjum and Chai (2020)", "RQ2"],
        ["Private identity-document store with administrator review", "Data collection", "Ghozi (2022); Bird (2026)", "RQ3"],
        ["Fixed taxonomy of 7 categories, 30 subcategories and 3 launch cities", "Dataset", "Ragavarthini et al. (2025); Gerwe and Silva (2020)", "RQ1"],
        ["Rental history generated by the platform as the RQ3 evaluation set", "Dataset", "Salgado et al. (2026); Neifer et al. (2023)", "RQ3"],
        ["Six-scenario scripted set with hand-calculated expected settlements", "Dataset", "Ahmadi et al. (2023)", "RQ2"],
        ["Behavioural events rather than opinions used as ground-truth labels", "Dataset", "Stevens and Bossauer (2020)", "RQ3"],
        ["Sealed pickup and return records with photographs and checksums", "Algorithm", "Zhao and Chen (2024); Papanastasiou et al. (2023)", "RQ2"],
        ["Deposit as its own status machine with a 48-hour return window", "Algorithm", "Kotiyal et al. (2026); Bashir et al. (2023)", "RQ2"],
        ["7-day claim window and 7-day response window from one constant", "Algorithm", "Ahmadi et al. (2023)", "RQ2"],
        ["Silence escalates to a human instead of accepting the claim", "Algorithm", "Papanastasiou et al. (2023); Zhao and Chen (2024)", "RQ2"],
        ["Settlement figures derived, with the commission rate frozen at verification", "Algorithm", "Kotiyal et al. (2026)", "RQ2"],
        ["Four-component trust score with shrinkage towards a neutral prior", "Algorithm", "Salgado et al. (2026); Stevens and Bossauer (2020)", "RQ3"],
        ["No score for a new member and no negative badge for anyone", "Algorithm", "Stevens and Bossauer (2020); Salgado et al. (2026)", "RQ3"],
        ["Deposit-band access ladder at PKR 25,000 and PKR 100,000", "Algorithm", "Ferraz and Ferraz (2021); Li and Wang (2020)", "RQ3"],
        ["Reciprocal review release inside a 14-day window", "Algorithm", "Papanastasiou et al. (2023)", "RQ3"],
        ["Local wallet providers defined now, integrated after the pilot", "Other", "Malik et al. (2025); Ali et al. (2022)", "RQ2"],
    ]

    table(
        doc,
        ["Methodology decision", "Type", "Justifying source (Chapter 2)", "Serves"],
        rows,
        widths=[7.4, 2.3, 4.6, 1.2],
        label="Table 11",
        caption_text="Traceability matrix: every methodology decision, the source that justifies it, and the "
        "research question it serves. Twenty-one rows, no blanks, no free text in the last column.",
        font_size=8.2,
        align_center_cols=(3,),
    )

    para(
        doc,
        "Two checks were run over the finished table. The first is coverage: every decision named in 4.1 "
        "to 4.3 appears as a row, and every source can be found in the reference list rather than in a "
        "paper we only read the abstract of. The second is balance: RQ1 is served by four rows, RQ2 by "
        "ten and RQ3 by seven. That is not a plan, it is a description — most of the technical work here "
        "is about money changing hands safely — and it is a warning for FYDP-2, where RQ1 and RQ3 will "
        "need more attention than they got.",
    )
    para(
        doc,
        "The weakest row is the sixth, the payment-provider interface. It is justified by two Pakistani "
        "studies that establish why local wallets matter and one that explains why card payment never took "
        "hold, but none of them is a study of a provider-agnostic abstraction; the row is really supported "
        "by engineering judgement with a citation standing next to it. What would strengthen it is a "
        "published integration report from a Pakistani marketplace describing how it moved from manual "
        "verification to a live wallet API — which is the document we will be writing ourselves after the "
        "pilot.",
    )

    doc.add_paragraph("4.4.1  Reflection on the matrix", style="Heading 2")
    reflection(
        doc,
        "Building this table was the most useful thing we did in the four weeks, and also the least "
        "comfortable. Three decisions we were proud of turned out to have citations supporting a nearby "
        "claim rather than the one we were making — the access ladder in particular leaned on Ferraz and "
        "Ferraz for a point about registration, when our use of it is about deposit value. We kept the "
        "decision and fixed the wording, but the exercise showed how easily a reference can be made to look "
        "like support for something it does not say. If the matrix is checked against the sources and not "
        "just read, that is exactly what we want it to catch.",
    )


# --------------------------------------------------------------------------
# references and appendices


REFERENCES = [
    "Ahmadi, K., Esmaili, M., & Khorsandi, S. (2023). A P2P file sharing market based on blockchain and IPFS with dispute resolution mechanism. In *IEEE International Conference on AI and Brain Things*. https://doi.org/10.1109/AIBThings58340.2023.10292453",
    "Ali, M. Y., et al. (2022). An integrated framework for mobile payment in Pakistan: Drivers, barriers, and facilitators of usage behavior. *Journal of Financial Services Marketing*. https://doi.org/10.1057/s41264-022-00160-5",
    "Anjum, S., & Chai, J. (2020). Drivers of cash-on-delivery method of payment in e-commerce shopping: Evidence from Pakistan. *SAGE Open, 10*(3). https://doi.org/10.1177/2158244020917392",
    "Bashir, R. N., et al. (2023). The role of customization, security, and flexibility in enhancing trust and continuance intentions of digital payment services: Insights from Pakistan. *iRASD Journal of Economics, 5*(2), 112–130. https://doi.org/10.30545/iraos.v5i2.441",
    "Bird, Y. (2026). Seal of approval? Trust signals and cultural distance in a global peer-to-peer platform market. *Organization Studies, 47*(3), 511–536. https://doi.org/10.1177/01708406241309995",
    "Ferraz, A., & Ferraz, C. (2021). Digital identity challenge: The security and convenience dilemma. In *Anais do Seminário Integrado de Software e Hardware (SEMISH)*. https://doi.org/10.5753/semish.2021.15829",
    "Filippas, A., Horton, J. J., & Zeckhauser, R. J. (2020). Owning, using, and renting: Some simple economics of the sharing economy. *Management Science, 66*(9), 4152–4172. https://doi.org/10.1287/mnsc.2019.3396",
    "Gerwe, O., & Silva, R. (2020). Clarifying the sharing economy: Conceptualization, typology, antecedents, and effects. *Academy of Management Perspectives, 34*(1), 141–166. https://doi.org/10.5465/amp.2017.0010",
    "Ghozi, A. (2022). The urgency of Electronic Know Your Customer (e-KYC): How electronic customer identification works to prevent money laundering in the fintech industry. *Diponegoro Law Review, 7*(1), 34–52. https://doi.org/10.14710/dilrev.7.1.2022.34-52",
    "Kotiyal, A., Negi, A., Bhardwaj, S., Jindal, P. K., & Puneet. (2026). An innovative collaborative resource sharing platform. In *Communications in Computer and Information Science*. https://doi.org/10.1007/978-3-032-11488-4_26",
    "Košťál, K., Mastilak, L., Morháč, D., & Valach, A. (2025). Beyond Airbnb: Exploring decentralized short-term rentals with Ethereum and IOTA. In *IEEE International Conference on Decentralized Applications and Systems (DAPPS)*. https://doi.org/10.1109/DAPPS65174.2025.00016",
    "Li, L., & Wang, W. (2020). The effects of online trust-building mechanisms on trust in the sharing economy: The perspective of providers. *Sustainability, 12*(5), 1717. https://doi.org/10.3390/su12051717",
    "Malik, M., Khan, L., Maqbool, H. N., & Anwar, M. (2025). Mobile wallets and digital payments: Assessing the impact on financial inclusion and consumer trust in Pakistan. *Review of Applied Management and Social Sciences, 8*(3), 1349–1362. https://doi.org/10.47067/ramss.v8i3.557",
    "Neifer, T., Bossauer, P., Pakusch, C., & Lawo, D. (2023). Trust-building in peer-to-peer carsharing: Design case study for algorithm-based reputation systems. *Computer Supported Cooperative Work*. https://doi.org/10.1007/s10606-022-09461-4",
    "Papanastasiou, Y., Yang, S. A., & Zhang, A. H. (2023). Improving dispute resolution in two-sided platforms: The case of review blackmail. *Management Science, 69*(10), 6021–6037. https://doi.org/10.1287/mnsc.2022.4655",
    "Ragavarthini, V., et al. (2025). Kaarya Setu: A community rental marketplace for sustainable resource utilization. In *IEEE International Conference on Intelligent Data Communication Technologies and Internet of Things*. https://doi.org/10.1109/ICIDCA66325.2025.11280491",
    "Rana, M. W., Zhang, S., & Hamid, I. (2022). An empirical study of models and APPs design of the adoption ease of mobile payments. In *IEEE International Conference on Emerging Technologies*. https://doi.org/10.1109/ICEET56123.2022.00115",
    "Salgado, S., Filieri, R., & Chameroy, F. (2026). Beyond unidimensional trust and user roles: A multidimensional role-based approach to trust. *Journal of Travel Research, 65*(3), 894–915. https://doi.org/10.1177/00472875241305628",
    "Stevens, G., & Bossauer, P. (2020). Who do you trust: Peers or technology? A conjoint analysis about computational reputation mechanisms. In *Proceedings of the 18th European Conference on Computer Supported Cooperative Work*. https://doi.org/10.18420/ecscw2020_ep01",
    "Tele, A. A., Kongari, D., Kane, O., Raut, S., & Lokhande, S. (2026). VahanRent — A blockchain powered peer-to-peer car rental system. *International Journal of Advanced Research in Science, Communication and Technology, 6*(3), 100–111. https://doi.org/10.48175/ijarsct-30910",
    "Waris, M., & Kim, S. I. (2023). A study of safety and security enhancement measures of sharing accommodation platform 'Airbnb'. *Journal of Digital Media & Culture Technology, 3*(2), 81–91. https://doi.org/10.29056/jdmct.2023.12.02",
    "Zhao, Y., & Chen, H. (2024). Enhancing access to digital justice: Digital governance of dispute resolution and dispute prevention in online commercial activities. *Journal of International Dispute Settlement*. https://doi.org/10.1093/jids/head024",
]


def references(doc):
    doc.add_paragraph("References", style="Heading 1")
    para(
        doc,
        "The list takes the approved Chapter 2 list and keeps the works this chapter actually cites: "
        "twenty-two sources, all peer reviewed, all published in 2020 or later, all managed in Mendeley "
        "and cited here in APA 7th edition.",
        size=9.5,
        italic=True,
    )
    for entry in REFERENCES:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.75)
        p.paragraph_format.first_line_indent = Cm(-0.75)
        p.paragraph_format.space_after = Pt(5)
        p.paragraph_format.line_spacing = 1.15
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        for idx, chunk in enumerate(entry.split("*")):
            if not chunk:
                continue
            run = p.add_run(chunk)
            run.font.size = Pt(9)
            run.italic = idx % 2 == 1


def appendices(doc):
    page_break(doc)
    doc.add_paragraph("Appendix A  AI tool usage record", style="Heading 1")
    para(
        doc,
        "The department requires one tool for generation and a different tool for verification, with the "
        "division of labour recorded. This table states what each tool did on this chapter, what it was "
        "asked to check, and what we changed as a result. The five group reflections are excluded from the "
        "table because no tool drafted or revised them.",
    )
    table(
        doc,
        ["Tool", "Role", "Used on", "What it was asked to do"],
        [
            [
                "Claude",
                "Generation / improvement",
                "3.1, 3.2, 3.3, 4.1, 4.2, 4.3, 4.4",
                "Tighten sentences we had already written; convert rough notes into single-sentence "
                "objectives; restructure the architecture description once the diagram was fixed; assemble "
                "matrix rows from a list of decisions we supplied. It was instructed not to add "
                "components, sources or numbers we had not given it.",
            ],
            [
                "ChatGPT",
                "Independent verification",
                "3.1, 3.2, 3.3, 4.2, 4.3, 4.4",
                "Check that the problem statement follows from the gap statement; that each objective maps "
                "to a sentence of the problem statement; that the dataset size is plausible for the method "
                "described; that the score formula is arithmetically correct and correctly described; and "
                "that no decision in 4.1 to 4.3 is missing from the matrix.",
            ],
        ],
        widths=[2.2, 3.0, 4.3, 6.0],
        label="Table 12",
        caption_text="AI tool usage record for this chapter. Both tools were used under the supervisor's "
        "guidance; every citation was checked against the original paper by a group member.",
        font_size=8.4,
    )

    para(
        doc,
        "What the verification pass actually caught, in the order it mattered: it flagged that RQ3 as "
        "first written could not be answered by a trained model at the volume the pilot would produce, "
        "which is why the score is presented as a fixed rule with an evaluation set rather than a learned "
        "model; it caught a sentence in 4.3.2 that implied the platform already refunds wallet payments, "
        "which it does not; it confirmed the shrinkage arithmetic in Table 9; and it listed four decisions "
        "from 4.1 that were missing from the first draft of the matrix. All four were added. Where the two "
        "tools disagreed on wording we kept our own version, and where the second tool asserted something "
        "about a paper we checked the paper — twice its summary of a cited finding was broader than what "
        "the paper said.",
    )

    page_break(doc)
    doc.add_paragraph("Appendix B  Decisions revised since the literature review", style="Heading 1")
    para(
        doc,
        "Chapter 2 described a system that, in five places, is not the system we now propose. Each change "
        "was made for a stated reason during design work, and each is recorded here rather than left for a "
        "reader to notice as an inconsistency.",
    )
    table(
        doc,
        ["Chapter 2 said", "This chapter proposes", "Why it changed"],
        [
            [
                "MongoDB for listings and user profiles, MySQL or PostgreSQL for financial records.",
                "A single PostgreSQL schema for every table.",
                "Evidence and money have to commit in one transaction. A handover record in one store and "
                "the deposit state in another is a reconciliation problem with no owner, and Prisma gives "
                "one migration history instead of two.",
            ],
            [
                "A separate installed application for phones, with iOS planned for a later release.",
                "One responsive web application used in any browser, with an installed build deferred.",
                "A second build doubles the release surface without changing what a user can do, and the "
                "adoption evidence points at screen-flow simplicity rather than at the delivery platform.",
            ],
            [
                "Deposits held by the owner, with the platform keeping a manual record.",
                "The deposit is collected by the platform and returned to the renter at settlement.",
                "Holding it on the owner's side made the renter's protection depend on the owner's honesty, "
                "which is the problem the project exists to solve.",
            ],
            [
                "Seven objectives, including an optional fraud-detection model at 80% accuracy.",
                "Three objectives, with the fraud model dropped.",
                "A model trained on thirty to fifty rentals would report an accuracy figure that meant "
                "nothing. Three objectives also match the department's one-to-one mapping requirement.",
            ],
            [
                "Automated escrow through a payment gateway.",
                "Manual verification of each payment by an administrator, behind a provider-agnostic "
                "interface.",
                "Fully automated escrow is not available without a merchant agreement in Pakistan. The "
                "interface lets a real integration replace the manual step without touching booking code.",
            ],
        ],
        widths=[4.6, 4.6, 6.3],
        label="Table 13",
        caption_text="Decisions revised since the literature review, with the reason for each. A baseline "
        "that silently contradicts the chapter before it is worse than one that explains itself.",
        font_size=8.3,
    )

    doc.add_paragraph("Appendix C  Implementation status at the time of writing", style="Heading 1")
    para(
        doc,
        "Included so a reader can tell which parts of this baseline describe working software and which "
        "describe a plan. The figures are read from the repository, not estimated.",
    )
    table(
        doc,
        ["Area", "Status", "Notes"],
        [
            ["Authentication, roles and email confirmation", "Built", "Auth.js v5, credentials and Google, password reset, session invalidation"],
            ["Listings, search, categories, availability", "Built", "Up to ten images per listing, signed uploads, filters and sorting"],
            ["Booking lifecycle and notifications", "Built", "Nine states including declined, cancelled and expired; guarded transitions"],
            ["Payments and manual verification", "Built", "Rent and deposit held separately; reference upload; administrator verification"],
            ["Handover records and evidence", "Built", "Sealed records, six photographs per handover, checksum on every file"],
            ["Deposits, claims and settlement", "Built", "48-hour window, 7-day claim and response windows, derived settlement"],
            ["Trust score and access ladder", "Built", "Pure functions with unit tests; bands and thresholds as described here"],
            ["Identity documents and review", "Built", "Private store, administrator approval, audit trail"],
            ["Admin panel", "Built", "Users, listings, bookings, payments, claims, reports, analytics, settlements"],
            ["Local wallet integrations", "Planned", "JazzCash and Easypaisa defined in the payment interface, not live"],
            ["Installed phone build", "Planned", "The web application already runs in a phone browser"],
            ["Production launch", "Blocked", "Waiting on a domain; the staging deployment is live"],
        ],
        widths=[5.4, 2.2, 7.9],
        label="Table 14",
        caption_text="Implementation status at the time of writing. Twenty-seven database tables, twenty-four "
        "migrations and fifty-two test files sit in the repository behind this table.",
        font_size=8.3,
        align_center_cols=(1,),
    )

    para(doc, "", space_after=6)
    rule(doc)
    para(
        doc,
        "End of FYDP-1 Proposed Methodology Baseline  ·  Group 23FYDP24  ·  Department of Computer Science, "
        "DUET Karachi",
        size=8,
        italic=True,
        color=GREY,
        align="center",
    )


def build():
    doc = new_document()
    props = doc.core_properties
    props.title = "SamaanShare - FYDP-1 Proposed Methodology Baseline"
    props.subject = "Problem Statement, Objectives, Research Questions and Core Methodology"
    props.author = "Group 23FYDP24, Department of Computer Science, DUET Karachi"
    props.keywords = "FYDP-1; SamaanShare; peer-to-peer rental; security deposit; trust score"
    props.comments = (
        "Prepared for the FYDP Coordination Committee. Prerequisite: approved Chapter 2 "
        "(Preliminary Literature Review)."
    )
    props.category = "Final Year Design Project"
    setup_header_footer(doc, "SamaanShare  ·  FYDP-1 Proposed Methodology Baseline", GROUP_ID)
    update_fields_on_open(doc)

    title_page(doc)
    page_break(doc)
    contents_page(doc)
    page_break(doc)
    note_page(doc)
    page_break(doc)

    section_three(doc)
    page_break(doc)

    section_four_intro(doc)
    section_41(doc)
    section_42(doc)
    section_43(doc)
    section_44(doc)

    page_break(doc)
    references(doc)

    appendices(doc)

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    doc.save(OUT)
    print("saved:", OUT)


if __name__ == "__main__":
    build()
