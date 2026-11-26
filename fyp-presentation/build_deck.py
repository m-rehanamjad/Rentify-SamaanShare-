"""Build the SamaanShare FYDP-I final presentation.

    python3 fyp-presentation/build_deck.py
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from pptx import Presentation
from pptx.util import Inches

import deck_part1 as p1
import deck_part2 as p2
import deck_part3 as p3

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "SamaanShare_FYDP1_Final_Presentation.pptx")



NOTES = [
    # 01 title
    "Open with the group, then one sentence: SamaanShare is a peer-to-peer rental platform for household goods, "
    "and the engineering problem we took on is trust. Supervisors and coordinators acknowledged by name. "
    "About 20 minutes, then a live walkthrough of staging.",
    # 02 agenda
    "Six sections, in the order a defence should follow: the problem, what the literature already settles, "
    "our three objectives and questions, the method and the three techniques, what is deployed and what the "
    "evaluation returned, then limits and the plan for phase two.",
    # 03 background
    "Lead with our own shelf: one DSLR, two drill kits, two projectors, a bridal sherwani, unused for about three "
    "years. None of us would hand any of it to a stranger from a classifieds board. The 90% cash-on-delivery "
    "figure is Anjum and Chai (2020), and it is why card rails cannot be assumed. Economic theory says platforms "
    "create welfare by cutting transaction cost, but only where trust is produced.",
    # 04 problem
    "Three failures, each of which alone would be tolerable: no rental-native marketplace, the deposit deadlock, "
    "and a reputation signal that cannot separate a first-time defaulter from a reliable member. Together they "
    "produce one outcome: owners delist, renters lose access, assets stay idle.",
    # 05 existing
    "Be fair to the alternatives. Classifieds give reach but no rental primitives; global marketplaces solved "
    "trust with institutional escrow and card rails we do not have; blockchain prototypes automate the lifecycle "
    "but add wallet and gas friction; academic trust models validate multidimensional reputation but stop before "
    "an escrow engine. Our position is the intersection.",
    # 06 literature themes
    "Twenty-five peer-reviewed studies, 2019 to 2026, organised into five themes. The point of the slide is that "
    "each theme has consensus inside it, and the four consensus statements at the bottom are the ones we built on.",
    # 07 lit table
    "Do not read the table. Pick two rows: Filippas gives the economic justification for renting durables, and "
    "Salgado gives the reason our score has four components rather than one star average. Full comparison in Table 2.1.",
    # 08 gap
    "The gap is not a missing technique; it is a missing assembly. Prior work assumes cards, Stripe-style escrow and "
    "high baseline trust, or adds a blockchain whose friction excludes the exact users we are building for. Nobody "
    "binds checksummed handover evidence to settlement arithmetic in a cash-dominant market.",
    # 09 objectives
    "One aim, three SMART objectives, each mapped to a question and to the instrument that tests it. This is the "
    "promise the rest of the deck has to keep.",
    # 10 RQs
    "Each question names its evidence base and its success criterion in advance. That is deliberate: the answer "
    "could have come out negative, and the criteria were written before the results.",
    # 11 methodology
    "Design, build, evaluate. Each technique is presented through the same six points: what it is, how it works, "
    "its literature basis, what changed for this domain, how we judged it, and which question it serves. The right "
    "column is the evidence base plus the invariants the database enforces.",
    # 12 architecture
    "Six layers, strictly decoupled. The one that matters for the defence is layer four: pure TypeScript domain "
    "modules with no database driver, no session lookup and no ambient clock read. That is what makes the illegal "
    "transition suite possible. Every mutation enters through layer three and passes five guards.",
    # 13 journey
    "Ten stages, two of them conditional. Note the four clocks: 48 hours for the refund after a sealed return, seven "
    "days for a claim, seven days to respond, fourteen days for reciprocal reviews. Nothing is held indefinitely.",
    # 14 state machine
    "One table of legal edges, consulted server-side only. The conditional update carries the expected status in its "
    "WHERE clause, so if another request already moved the record, zero rows match and the transaction aborts. We kept "
    "the automated lifecycle from the literature but used relational locks instead of a chain: no gas, no wallet custody.",
    # 15 handover
    "This is the design principle of the whole deposit mechanism: a record its author can revise is a claim, not "
    "evidence. So the log is written in the same transaction as the status change, triggers refuse updates and "
    "deletes, and the checksum is read server-side and is unique per table.",
    # 16 escrow + settlement
    "The deposit is its own state machine, and the settlement routine derives every figure rather than accepting one. "
    "Both columns sum to the 34,000 collected, and that identity is asserted by code, not by an accountant. Commission "
    "is computed from the rental amount alone, so it structurally cannot touch the deposit.",
    # 17 formula
    "Four components, weights sum to one. The reputation term is shrunk toward a neutral 3.5 with five pseudo-counts, "
    "so a single five-star rating cannot outrank fifty good ones. Experience saturates at ten rentals and defaults to "
    "1.0 for new members, and verification is a graded component rather than a hard gate.",
    # 18 example
    "Member B has a flawless single rating and still scores 0.574 because volume, reliability and identity carry weight. "
    "That is the intended behaviour. The ladder then decides exposure: CNIC or three completed rentals for the top tier, "
    "email and phone for the entry tier.",
    # 19 traceability
    "Twenty-one decisions, each mapped to a peer-reviewed source and a question. We treat this as a defence in itself: "
    "if a decision cannot name its source, it does not ship.",
    # 20 implementation
    "What is live: 27 tables and 24 migrations in one relational store, 808 unit cases across 52 test files, Playwright "
    "end to end against a production build with axe scans, and migrations gated in the deploy so a build cannot serve a "
    "schema it does not match.",
    # 21 walkthrough A
    "These are renderings of the deployed screens, not screenshots; the live walkthrough follows. Two points: the listing "
    "cannot be published without a deposit, and the price breakdown is snapshotted onto the booking at approval, so a "
    "later listing edit cannot reach an agreed rental.",
    # 22 walkthrough B
    "The handover form is camera-only and seals on submission. On the admin side, payments are verified manually with an "
    "audit trail, and a damage claim is decided on the paired pickup and return photographs, with the figure computed "
    "from the claim and the deposit rather than typed in.",
    # 23 walkthrough C
    "The public profile is where the score becomes legible to a stranger: the composite, the band, the identity state, "
    "and the ladder showing what unlocks the next tier. Reputation is split by role, so standing as an owner is never "
    "averaged with standing as a renter.",
    # 24 evaluation design
    "Every question has an instrument, an artefact and a pass condition, written before we ran anything. The green band "
    "states how each could have failed, which is the honest way to present an evaluation.",
    # 25 results RQ1
    "Zero illegal transitions accepted; one audit row per permitted transition. The table on the right is the test plan, "
    "not an illustration. Be explicit about the limit: this is testing, not formal verification, and two write paths - "
    "the expiry worker and administrative suspension - remain outside the table and are covered by review.",
    # 26 results RQ2
    "Six scenarios, each with a hand-calculated baseline worked out before the code ran. Every settlement matched to the "
    "rupee, both conservation identities held, and no balance outlived its clock. Scenario (f) is the important one: "
    "silence escalates a claim rather than quietly releasing the deposit.",
    # 27 results RQ3
    "The synthetic population is how we answered RQ3 without inventing history: 2,000 modelled members with stated "
    "behaviour distributions. Completing members outscore defaulting ones, and the bands land where the formula intends. "
    "Say plainly what it does not prove: the shapes are illustrative and the pilot is a case study, not a sample.",
    # 28 limitations
    "Seven limitations, each stated as a design constraint with something shipped against it. The two the panel usually "
    "push on are manual payment verification and CNIC reluctance; both have a concrete answer here.",
    # 29 contributions
    "Each piece exists in the literature. The contribution is the coupling: evidence sealed at handover drives the "
    "settlement arithmetic, and settlement history drives the trust that decides who may hold a 100,000-rupee item.",
    # 30 reflections
    "Division of work and the three things that cost us most: state atomicity, the RQ3 method change, and the decision "
    "to document manual escrow as a limitation instead of claiming automation we did not have.",
    # 31 roadmap
    "Phase two: decouple the backend into a standalone Node API, ship React Native on that API, and close the payment gap "
    "with live JazzCash, Easypaisa and Safepay integrations. The pilot targets are the volumes we need for RQ3 to become "
    "a measurement rather than a model.",
    # 32 conclusion
    "Three questions, three answers, with the limit on RQ3 stated in the same breath. Then the closing line: the trust "
    "deficit is an engineering problem before it is a social one. Thank the supervisors, coordinators and chairperson.",
    # 33 thanks
    "Stop talking. Invite the panel to the live staging walkthrough and hold the demo script ready: register, request, "
    "approve, pay, collect, return, review.",
    # 34 references
    "Leave this up while questions are asked. The AI-tool disclosure is in the thesis appendix, per DUET guidelines: "
    "Claude for drafting and editing, ChatGPT for verification passes.",
]

def build():
    prs = Presentation()
    prs.slide_width = Inches(13.3333)
    prs.slide_height = Inches(7.5)

    p1.title_slide(prs)
    p1.agenda(prs)
    p1.background(prs)
    p1.problem(prs)
    p1.existing(prs)
    p1.literature_themes(prs)
    p1.literature_table(prs)
    p1.gap(prs)
    p1.objectives(prs)
    p1.research_questions(prs)

    p1.methodology(prs)
    p1.architecture(prs)
    p2.journey(prs)
    p2.state_machine(prs)
    p2.handover(prs)
    p2.escrow(prs)
    p2.trust_formula(prs)
    p2.trust_example(prs)
    p2.traceability(prs)
    p2.implementation(prs)

    p2.walkthrough_a(prs)
    p2.walkthrough_b(prs)
    p2.walkthrough_c(prs)

    p3.evaluation_design(prs)
    p3.results_rq1(prs)
    p3.results_rq2(prs)
    p3.results_rq3(prs)
    p3.limitations(prs)
    p3.contributions(prs)
    p3.reflections(prs)
    p3.roadmap(prs)
    p3.conclusion(prs)
    p3.thanks(prs)
    p3.references(prs)

    for slide, note in zip(prs.slides, NOTES):
        slide.notes_slide.notes_text_frame.text = note

    prs.save(OUT)
    print(f"saved {OUT}  ({len(prs.slides.__iter__.__self__._sldIdLst)} slides)")
    return OUT


if __name__ == "__main__":
    build()
