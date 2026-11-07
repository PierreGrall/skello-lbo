"""Charts for the memo, drawn only from outputs/analysis.json."""
from __future__ import annotations

import json

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt

from engine import ROOT, load_inputs

OUT = ROOT / "memo" / "figures"
INK, MUTED, RULE, SURFACE = "#0b0b0b", "#52514e", "#d9d8d3", "#ffffff"
BASE, DOWN, PLAN = "#2a78d6", "#eb6834", "#1baf7a"
NEUTRAL = "#8f8e88"

plt.rcParams.update({
    "font.family": "Inter", "font.size": 8, "axes.edgecolor": RULE, "axes.linewidth": 0.8, "axes.labelcolor": MUTED,
    "xtick.color": MUTED, "ytick.color": MUTED, "xtick.major.size": 0, "ytick.major.size": 0, "axes.titlesize": 8,
    "figure.facecolor": SURFACE, "axes.facecolor": SURFACE, "pdf.fonttype": 42,
})


def clean(ax, grid="y"):
    for side in ("top", "right"):
        ax.spines[side].set_visible(False)
    if grid:
        ax.grid(axis=grid, color=RULE, linewidth=0.5)
        ax.set_axisbelow(True)


def required_arr(a):
    """Iso-return lines: the pairs of exit ARR and EV/ARR multiple that deliver a given MOIC."""
    fig, ax = plt.subplots(figsize=(6.7, 2.45))
    req = a["required_arr"]
    xs = req["multiples"]
    shades = {"2.0": "#b9b8b1", "2.5": "#7d7c76", "3.0": INK}
    for target in ("2.0", "2.5", "3.0"):
        ys = req["targets"][target]
        ax.plot(xs, ys, color=shades[target], linewidth=1.6 if target == "3.0" else 1.2, solid_capstyle="round")
        ax.annotate(f"{float(target):.1f}x", (xs[-1], ys[-1]), xytext=(5, 0), textcoords="offset points", ha="left", va="center",
                    color=INK if target == "3.0" else MUTED, fontweight="bold" if target == "3.0" else "normal")
    for name, colour, off in (("Downside", DOWN, (7, 0)), ("Base", BASE, (7, -9)), ("Plan", PLAN, (7, 0))):
        s = a["scenarios"][name]
        m = s["exits"]["Dec-30"]["ev_arr"]
        y = s["rows"]["arr_total"][4]
        ax.scatter([m], [y], s=38, color=colour, edgecolor=SURFACE, linewidth=1.2, zorder=5)
        ax.annotate(f"{name}: {s['exits']['Dec-30']['moic']:.1f}x", (m, y), xytext=off, textcoords="offset points", color=INK, fontweight="bold", va="center")
    base = a["scenarios"]["Base"]
    nb = base["next_buyer"]
    ax.scatter([nb["ev_arr"]], [base["rows"]["arr_total"][4]], s=34, facecolor=SURFACE, edgecolor=BASE, linewidth=1.3, zorder=5)
    ax.annotate(f"Base, sponsor buyer: {nb['moic']:.1f}x", (nb["ev_arr"], base["rows"]["arr_total"][4]), xytext=(-7, 0), textcoords="offset points", ha="right", va="center", color=MUTED)
    m0 = base["entry"]["m_entry"]
    ax.scatter([m0], [50], s=22, marker="s", color=NEUTRAL, zorder=4)
    ax.annotate(f"Entry: {m0:.1f}x, €50M", (m0, 50), xytext=(7, 0), textcoords="offset points", color=MUTED, va="center")
    ax.set_xlim(2.2, 8.5)
    ax.set_ylim(35, 300)
    ax.set_xticks(range(3, 9))
    ax.set_xticklabels([f"{x}.0x" for x in range(3, 9)])
    ax.set_xlabel("Exit EV / ARR")
    ax.set_ylabel("ARR at Dec-30 (€M)")
    clean(ax)
    fig.tight_layout(pad=0.4)
    fig.savefig(OUT / "required_arr.pdf")
    plt.close(fig)


def arr_paths(a):
    inp = load_inputs()
    fig, ax = plt.subplots(figsize=(3.25, 2.1))
    x = [0, 0.5, 1.5, 2.5, 3.5, 4.5]
    for name, colour in (("Plan", PLAN), ("Base", BASE), ("Downside", DOWN)):
        y = [inp["deal"]["arr_entry"]] + a["scenarios"][name]["rows"]["arr_total"][:5]
        ax.plot(x, y, color=colour, linewidth=1.8, solid_capstyle="round")
        ax.annotate(f"{name} {y[-1]:.0f}", (x[-1], y[-1]), xytext=(4, 0), textcoords="offset points", va="center", color=INK, fontsize=7.5)
    ax.scatter([2.0, 4.5], [100, 200], s=26, marker="D", facecolor=SURFACE, edgecolor=INK, linewidth=1, zorder=6)
    ax.annotate("Reported targets", (2.0, 100), xytext=(-4, 12), textcoords="offset points", color=MUTED, fontsize=7, ha="right")
    ax.set_xticks([0, 1.5, 3.5, 4.5])
    ax.set_xticklabels(["Jun-26", "Dec-27", "Dec-29", "Dec-30"])
    ax.set_xlim(-0.1, 6.0)
    ax.set_ylim(0, 230)
    ax.set_ylabel("ARR (€M)")
    clean(ax)
    fig.tight_layout(pad=0.3)
    fig.savefig(OUT / "arr_paths.pdf")
    plt.close(fig)


def bridge(a):
    b = a["scenarios"]["Base"]["bridge"]
    steps = [("Entry\nequity", b["entry_equity"], "total"), ("Organic\nARR", b["organic"], "delta"), ("Acquired\nARR", b["acquired"], "delta"),
             ("Multiple", b["multiple"], "delta"), ("Net\ndebt", b["net_debt"], "delta"), ("Exit\ncosts", b["costs"], "delta"), ("Exit\nequity", b["exit_equity"], "total")]
    fig, ax = plt.subplots(figsize=(3.3, 2.1))
    running = 0.0
    for i, (label, value, kind) in enumerate(steps):
        if kind == "total":
            bottom, height, colour = 0, value, "#3d3c38"
            running = value
        else:
            bottom = running if value >= 0 else running + value
            height, colour = abs(value), BASE if value >= 0 else DOWN
            running += value
        ax.bar(i, height, bottom=bottom, width=0.62, color=colour)
        text = f"{value:,.0f}" if kind == "total" else f"{value:+,.0f}"
        ax.text(i, bottom + height + 12, text, ha="center", va="bottom", fontsize=7, color=INK)
    ax.set_xticks(range(len(steps)))
    ax.set_xticklabels([s[0] for s in steps], fontsize=6.6)
    ax.set_ylim(0, 1090)
    ax.set_yticks([])
    ax.spines["left"].set_visible(False)
    clean(ax, grid=None)
    ax.spines["left"].set_visible(False)
    fig.tight_layout(pad=0.3)
    fig.savefig(OUT / "bridge.pdf")
    plt.close(fig)


def tornado(a):
    t = a["tornado"]
    rows = [r for r in t["rows"]][::-1]
    base = t["base"]
    fig, ax = plt.subplots(figsize=(6.7, 2.7))
    for i, r in enumerate(rows):
        lo, hi = min(r["low"], r["high"]), max(r["low"], r["high"])
        if lo < base - 0.005:
            ax.barh(i, base - lo, left=lo, height=0.56, color=DOWN)
        if hi > base:
            ax.barh(i, hi - base, left=base, height=0.56, color=BASE)
        ax.text(lo - 0.015, i, f"{lo:.2f}x", ha="right", va="center", fontsize=7, color=INK)
        if hi >= base + 0.005:
            ax.text(hi + 0.015, i, f"{hi:.2f}x", ha="left", va="center", fontsize=7, color=INK)
    ax.axvline(base, color=INK, linewidth=0.9)
    ax.set_yticks(range(len(rows)))
    ax.set_yticklabels([r["label"].replace("EUR ", "€") for r in rows], color=INK, fontsize=7.4)
    ax.set_xlim(1.6, 2.65)
    ax.set_xticks([])
    for side in ("top", "right", "bottom", "left"):
        ax.spines[side].set_visible(False)
    ax.text(base, len(rows) - 0.3, f"Base {base:.2f}x", ha="center", va="bottom", fontsize=7.4, color=INK, fontweight="bold")
    fig.tight_layout(pad=0.3)
    fig.savefig(OUT / "tornado.pdf")
    plt.close(fig)


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    a = json.loads((ROOT / "outputs" / "analysis.json").read_text())
    required_arr(a)
    arr_paths(a)
    bridge(a)
    tornado(a)


if __name__ == "__main__":
    main()
