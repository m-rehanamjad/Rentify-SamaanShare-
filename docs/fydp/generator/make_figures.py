"""Builds the two figures used in the FYDP-1 Proposed Methodology Baseline.

Figure 1  architecture.png  - layered system architecture
Figure 2  deposit_flow.png  - deposit, handover and claim flow

Run:  python3 make_figures.py
Both files are written at 300 DPI, above the department's 150 DPI minimum.
"""

from __future__ import annotations

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyArrowPatch, FancyBboxPatch, Rectangle

OUT_DIR = "."

INK = "#1F2933"
MUTED = "#5B6674"
RULE = "#9AA5B1"

STYLES = {
    "users": ("#EDF1F6", "#8C97A6"),
    "client": ("#E7EFF7", "#6C8EBF"),
    "app": ("#DAE6F2", "#4C7CA9"),
    "domain": ("#FCF1DC", "#BE8A2C"),
    "service": ("#E5EFE8", "#5D8A69"),
    "data": ("#ECEDEE", "#7B7B7B"),
    "ops": ("#F6F7F9", "#AEB6BF"),
}


def band(ax, y_top, height, label, sub, fill, edge, dashed=False):
    """A full-width band with a label column on the left."""
    rect = FancyBboxPatch(
        (1.5, y_top - height),
        97.0,
        height,
        boxstyle="round,pad=0,rounding_size=0.7",
        linewidth=1.0 if not dashed else 0.9,
        edgecolor=edge,
        facecolor=fill,
        linestyle="--" if dashed else "-",
        zorder=1,
    )
    ax.add_patch(rect)
    ax.plot(
        [22.3, 22.3],
        [y_top - height + 0.9, y_top - 0.9],
        color=edge,
        linewidth=0.8,
        alpha=0.55,
        zorder=2,
    )
    ty = y_top - height / 2 + (1.3 if sub else 0)
    ax.text(
        2.8,
        ty,
        label,
        fontsize=6.4,
        fontweight="bold",
        color=INK,
        va="center",
        ha="left",
        zorder=3,
        linespacing=1.3,
    )
    if sub:
        ax.text(
            2.8,
            ty - 2.9,
            sub,
            fontsize=5.5,
            color=MUTED,
            va="center",
            ha="left",
            style="italic",
            zorder=3,
            linespacing=1.2,
        )


def boxes(ax, y_top, height, cols, pad_x=1.0, pad_y=1.35):
    """Returns centre-x and box geometry for `cols` evenly spaced cells."""
    left, right = 23.0, 97.0
    gap = pad_x
    width = (right - left - gap * (cols - 1)) / cols
    out = []
    for i in range(cols):
        x = left + i * (width + gap)
        out.append(
            dict(
                x=x,
                y=y_top - height + pad_y,
                w=width,
                h=height - 2 * pad_y,
                cx=x + width / 2,
                cy=y_top - height / 2,
            )
        )
    return out


def cell(ax, geom, title, sub, fill="#FFFFFF", edge="#8C97A6", title_size=6.9):
    ax.add_patch(
        FancyBboxPatch(
            (geom["x"], geom["y"]),
            geom["w"],
            geom["h"],
            boxstyle="round,pad=0,rounding_size=0.45",
            linewidth=0.75,
            edgecolor=edge,
            facecolor=fill,
            zorder=4,
        )
    )
    cx = geom["cx"]
    if sub:
        ax.text(
            cx,
            geom["cy"] + 1.05,
            title,
            fontsize=title_size,
            fontweight="bold",
            color=INK,
            ha="center",
            va="center",
            zorder=5,
        )
        ax.text(
            cx,
            geom["cy"] - 1.35,
            sub,
            fontsize=5.9,
            color=MUTED,
            ha="center",
            va="center",
            zorder=5,
            linespacing=1.2,
        )
    else:
        ax.text(
            cx,
            geom["cy"],
            title,
            fontsize=title_size,
            fontweight="bold",
            color=INK,
            ha="center",
            va="center",
            zorder=5,
        )


def arrow(ax, x, y_from, y_to, style="-", color="#6B7280"):
    ax.add_patch(
        FancyArrowPatch(
            (x, y_from),
            (x, y_to),
            arrowstyle="-|>",
            mutation_scale=7,
            linewidth=0.9,
            linestyle=style,
            color=color,
            shrinkA=0,
            shrinkB=0,
            zorder=6,
        )
    )


def figure_architecture(path: str) -> None:
    fig = plt.figure(figsize=(6.3, 6.55), dpi=300)
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis("off")

    gaps = 2.0
    plan = [
        ("USER LAYER", None, 9.0, 3),
        ("CLIENT LAYER", "browser: desktop,\ntablet, phone", 11.0, 3),
        ("APPLICATION LAYER", "Next.js 15 on Vercel", 11.0, 3),
        ("DOMAIN RULE LAYER", "pure TypeScript,\nunit tested", 20.0, 3),
        ("PLATFORM SERVICES", "shared services", 11.0, 4),
        ("DATA LAYER", "relational store, ORM", 11.0, 3),
    ]
    ops_h = 8.0

    total = sum(h for _, _, h, _ in plan) + ops_h + gaps * len(plan)
    y = 50 + total / 2

    contents = [
        [
            ("Renter", "browses, books, pays"),
            ("Owner", "lists, approves, hands over"),
            ("Administrator", "verifies, mediates, settles"),
        ],
        [
            ("Marketplace and search", "listing, category, city pages"),
            ("Booking and payment", "timeline and evidence upload"),
            ("Handover and review", "photos, condition, ratings"),
        ],
        [
            ("Server actions", "every mutation, one path"),
            ("Session and role check", "Auth.js, USER or ADMIN"),
            ("Validation and limits", "Zod schemas, rate limits"),
        ],
        [
            ("Booking lifecycle guard", "PENDING to REVIEWED"),
            ("Trust score", "four weighted components"),
            ("Access tiers", "deposit bands 25k / 100k"),
            ("Deposit and claim rules", "48-hour window, 7-day claim"),
            ("Handover protocol", "sealed record, 6 photos"),
            ("Review release", "reciprocal, 14 days"),
        ],
        [
            ("Payments", "manual transfer, card"),
            ("Realtime", "chat and notifications"),
            ("Email", "verification links"),
            ("Scheduled jobs", "expiry, escalation"),
        ],
        [
            ("PostgreSQL + Prisma", "27 tables, ACID, migrations"),
            ("Cloudinary", "images, CNIC, checksums"),
            ("Audit trail", "one row per admin decision"),
        ],
    ]

    prev_bottom = None
    for idx, (label, sub, height, cols) in enumerate(plan):
        key = ["users", "client", "app", "domain", "service", "data"][idx]
        fill, edge = STYLES[key]
        top = y
        band(ax, top, height, label, sub, fill, edge)

        if idx == 3:
            # two rows of three inside the domain band
            row_h = (height - 1.4) / 2
            for r in range(2):
                row_top = top - 0.7 - r * row_h
                geoms = boxes(ax, row_top, row_h, 3, pad_y=1.0)
                for c, (t, s) in enumerate(contents[idx][r * 3 : r * 3 + 3]):
                    cell(ax, geoms[c], t, s, fill="#FFFDF7", edge=edge)
        else:
            geoms = boxes(ax, top, height, cols)
            for c, (t, s) in enumerate(contents[idx]):
                cell(ax, geoms[c], t, s, fill="#FFFFFF", edge=edge)

        if prev_bottom is not None:
            arrow(ax, 50, prev_bottom, top)

        prev_bottom = top - height
        y = top - height - gaps

    # operations strip
    fill, edge = STYLES["ops"]
    band(ax, y, ops_h, "OPERATIONS\nAND QUALITY", None, fill, edge, dashed=True)
    ax.text(
        60.0,
        y - ops_h / 2,
        "Vercel hosting  |  Sentry error tracking  |  Vitest unit tests  |  "
        "Playwright end-to-end tests",
        fontsize=6.1,
        color=INK,
        ha="center",
        va="center",
        zorder=5,
    )
    arrow(ax, 50, prev_bottom, y, style=(0, (2, 2)))

    fig.savefig(path, dpi=300, facecolor="white")
    plt.close(fig)


def figure_deposit(path: str) -> None:
    fig = plt.figure(figsize=(6.3, 2.85), dpi=300)
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, 100)
    ax.set_ylim(14, 100)
    ax.axis("off")

    BOX_FILL = "#FFFFFF"
    BOX_EDGE = "#8C97A6"

    def flow_box(x, y, w, h, title, sub, fill=BOX_FILL, edge=BOX_EDGE, t_size=6.4):
        ax.add_patch(
            FancyBboxPatch(
                (x, y),
                w,
                h,
                boxstyle="round,pad=0,rounding_size=0.6",
                linewidth=0.8,
                edgecolor=edge,
                facecolor=fill,
                zorder=4,
            )
        )
        ax.text(
            x + w / 2,
            y + h * 0.65,
            title,
            fontsize=t_size,
            fontweight="bold",
            color=INK,
            ha="center",
            va="center",
            zorder=5,
            linespacing=1.15,
        )
        ax.text(
            x + w / 2,
            y + h * 0.27,
            sub,
            fontsize=5.6,
            color=MUTED,
            ha="center",
            va="center",
            zorder=5,
            linespacing=1.25,
        )

    def seg(pts, style="-", color="#6B7280", head=True):
        for i in range(len(pts) - 1):
            (x1, y1), (x2, y2) = pts[i], pts[i + 1]
            last = i == len(pts) - 2
            ax.add_patch(
                FancyArrowPatch(
                    (x1, y1),
                    (x2, y2),
                    arrowstyle="-|>" if (head and last) else "-",
                    mutation_scale=6.5,
                    linewidth=0.85,
                    linestyle=style,
                    color=color,
                    shrinkA=0,
                    shrinkB=0,
                    zorder=6,
                )
            )

    def heading(x, y, text):
        ax.text(
            x,
            y,
            text,
            fontsize=6.2,
            fontweight="bold",
            color=MUTED,
            ha="left",
            va="center",
            zorder=5,
            style="italic",
        )

    heading(2.0, 95.0, "Normal rental")

    row = [
        ("Booking\napproved", "owner accepts", BOX_FILL),
        ("Payment\nverified", "deposit held by\nthe platform", "#E5EFE8"),
        ("Pickup record\nsealed", "6 photos, condition,\nchecksum", "#FCF1DC"),
        ("Item in use", "booking ACTIVE", BOX_FILL),
        ("Return record\nsealed", "owner records\ncondition", "#FCF1DC"),
        ("48-hour\ndeposit window", "clock starts\nat return", "#E7EFF7"),
    ]
    w, h, gap = 13.9, 18.0, 2.2
    y_top = 88.0
    cx = []
    for i, (t, sub, f) in enumerate(row):
        x = 2.0 + i * (w + gap)
        flow_box(x, y_top - h, w, h, t, sub, fill=f)
        cx.append(x + w / 2)
        if i:
            seg([(x - gap + 0.1, y_top - h / 2), (x - 0.1, y_top - h / 2)])

    # settlement box, bottom right
    sx, sy, sw, sh = 76.5, 20.0, 19.9, 26.0
    flow_box(
        sx,
        sy,
        sw,
        sh,
        "Settlement",
        "owner is paid the rent\nplus any upheld claim;\nthe rest of the deposit\ngoes back to the renter",
        fill="#E5EFE8",
    )
    seg([(cx[5], y_top - h), (cx[5], sy + sh + 0.1)])
    ax.text(
        cx[5] - 1.2,
        58.0,
        "nothing claimed",
        fontsize=5.4,
        color=MUTED,
        ha="right",
        va="center",
        style="italic",
        zorder=6,
    )

    # exception chain, bottom left
    heading(2.0, 53.0, "Exception: the item comes back damaged")
    chain = [
        ("Claim filed", "within 7 days of\nreturn, bounded by\nthe deposit"),
        ("Renter replies", "7 days; silence\nescalates to an\nadministrator"),
        ("Admin ruling", "written reason kept\nin the audit log"),
    ]
    bw, bh = 22.5, 24.0
    by = 20.0
    bx = [2.0, 27.5, 53.0]
    for i, (t, sub) in enumerate(chain):
        flow_box(bx[i], by, bw, bh, t, sub)
        if i:
            seg([(bx[i] - 2.4, by + bh / 2), (bx[i] - 0.1, by + bh / 2)])
    seg([(bx[2] + bw + 0.2, by + bh / 2), (sx - 0.1, by + bh / 2)])

    # dashed branch from the return record down and across into the claim chain
    seg(
        [
            (cx[4], y_top - h - 0.1),
            (cx[4], 47.5),
            (bx[0] + bw / 2, 47.5),
            (bx[0] + bw / 2, by + bh + 0.1),
        ],
        style=(0, (2, 2)),
    )

    fig.savefig(path, dpi=300, facecolor="white")
    plt.close(fig)


if __name__ == "__main__":
    figure_architecture(f"{OUT_DIR}/architecture.png")
    figure_deposit(f"{OUT_DIR}/deposit_flow.png")
    print("wrote architecture.png and deposit_flow.png")
