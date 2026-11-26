# FYDP-I final presentation — SamaanShare

`SamaanShare_FYDP1_Final_Presentation.pptx` — 34 slides, 16:9, with speaker notes on every slide.

## How it was built

The deck is generated, not hand-assembled, so edits are made in code and rebuilt:

```bash
python3 fyp-presentation/build_deck.py          # writes the .pptx next to this file
python3 fyp-presentation/preview.py fyp-presentation/SamaanShare_FYDP1_Final_Presentation.pptx /tmp/preview
```

| File | Purpose |
| --- | --- |
| `deck_theme.py` | Palette, type scale, layout primitives (cards, pills, tables, footer, title auto-fit) |
| `deck_part1.py` | Slides 1–12: title, agenda, problem, literature, gap, objectives, questions, method, architecture |
| `deck_part2.py` | Slides 13–23: rental journey, the three core techniques, traceability, implementation, walkthrough |
| `deck_part3.py` | Slides 24–34: evaluation design, results per research question, limitations, contributions, reflections, roadmap, close |
| `build_deck.py` | Orchestrates the build and attaches the speaker notes |
| `preview.py` | Approximate PIL rasteriser used to check layout (no Office renderer in this sandbox) |

## Design

Colours and type follow the product's own design system in `src/app/globals.css`: deep navy
`#142B3B` for ink, emerald `#047857` for brand, mint `#E6F4EE` for surfaces, `#F8FAFC` page,
`#536273` muted text. Fonts are set to Segoe UI (with `Inter` available in the app itself).

## Content sources

Everything factual is grounded in the FYDP-I report (`docs/` mirrors it) and this repository:

- Chapters 1–3 of the thesis: problem statement, objectives, RQ evidence bases, Table 2.1,
  Table 3.2 (settlement arithmetic), Table 3.3 (trust worked examples), Table 3.4 (traceability),
  Table 3.5 (limitations).
- Repository facts: 27 Prisma models, 24 migrations, 9 booking statuses, 6 deposit states,
  52 test files / 808 unit cases, Playwright + axe, Pusher, Cloudinary, Vercel/Neon staging.

## Note on the images

`assets/mock-*.png` are **AI-generated illustrations** of the deployed screens (browse, booking,
handover log, admin queue, trust profile), used because screenshots need to be captured from the
live staging deployment. Replace them with real screenshots before the final submission if the
department requires them — keep the same 16:9 framing and the layout will hold.

`assets/logo-*.png` are recoloured copies of `public/images/logo-mark-light.png` with the grey
plate swapped for the slide background so the mark sits cleanly on navy, paper and white.

## Delivery order

33 slides are content; slide 32 is the conclusion, 33 the "thank you", 34 the references left up
during questions. The walkthrough (21–23) is where the live staging demo takes over.
