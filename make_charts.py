"""
make_charts.py
==============
Builds a three-chart pack from the data extracted by extract_annual_report.py.

Each chart answers one question about Stora Enso's 2025 performance:

  1. chart1_margin_gap.png   Which segments are furthest from the 10% margin target?
  2. chart2_ebit_bridge.png  What moved group adjusted EBIT from 2024 to 2025?
  3. chart3_sensitivity.png  What would move Packaging Solutions' margin most?

Run after the extraction:
    pip install matplotlib
    python make_charts.py                     # reads output/, writes output/charts/
"""

import argparse
from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd

# ---------------------------------------------------------------------------
# Style: one accent colour for the story, grey for everything else
# ---------------------------------------------------------------------------
ACCENT = "#1F4E3D"        # Packaging Solutions / positive
ACCENT_LIGHT = "#7FB685"
NEUTRAL = "#9AA5A0"
NEGATIVE = "#B23A48"
TARGET = "#C9A227"

plt.rcParams.update({
    "figure.dpi": 150,
    "font.family": "DejaVu Sans",
    "font.size": 11,
    "axes.titlesize": 14,
    "axes.titleweight": "bold",
    "axes.spines.top": False,
    "axes.spines.right": False,
    "axes.grid": True,
    "grid.alpha": 0.25,
})

SOURCE = "Source: Stora Enso Annual Report 2025 (p. 51, 53-54, 60). Figures in EUR million. Unofficial analysis."
MARGIN_TARGET = 0.10      # group target from 2026 (AR2025 p. 64)


def footnote(fig, text=SOURCE):
    fig.text(0.01, 0.01, text, fontsize=7.5, color="#666666")


# ---------------------------------------------------------------------------
# Chart 1: margin by segment vs the 10% target
# ---------------------------------------------------------------------------
def chart_margin_gap(seg, out):
    df = seg[(seg["year"] == 2025) & (~seg["segment"].isin(["Eliminations", "Other"]))].copy()
    df["margin"] = df["adj_ebit"] / df["sales"]
    df = df.sort_values("margin")

    colors = [ACCENT if s == "Packaging Solutions" else NEUTRAL for s in df["segment"]]

    fig, ax = plt.subplots(figsize=(9, 5))
    bars = ax.barh(df["segment"], df["margin"] * 100, color=colors)

    ax.axvline(MARGIN_TARGET * 100, color=TARGET, linestyle="--", linewidth=1.5)
    ax.text(MARGIN_TARGET * 100 + 0.15, -0.45, "Group target >10%", color=TARGET, fontsize=9)

    for bar, value in zip(bars, df["margin"] * 100):
        offset = 0.15 if value >= 0 else -0.15
        ax.text(value + offset, bar.get_y() + bar.get_height() / 2, f"{value:.1f}%",
                va="center", ha="left" if value >= 0 else "right", fontsize=10)

    ax.set_title("No segment reaches the 10% margin target;\nPackaging Solutions is 8.6 points short", loc="left")
    ax.set_xlabel("Adjusted EBIT margin 2025 (%)")
    ax.set_xlim(-1.5, 12)
    ax.grid(axis="y", visible=False)
    footnote(fig, SOURCE + " Segment Other excluded (not an operating segment).")
    fig.tight_layout(rect=[0, 0.03, 1, 1])
    fig.savefig(out / "chart1_margin_gap.png", bbox_inches="tight")
    plt.close(fig)


# ---------------------------------------------------------------------------
# Chart 2: what moved group adjusted EBIT from 2024 to 2025
# ---------------------------------------------------------------------------
def chart_ebit_bridge(seg, group, out):
    wide = seg.pivot(index="segment", columns="year", values="adj_ebit")
    wide["change"] = wide[2025] - wide[2024]
    changes = wide["change"].sort_values(ascending=False)

    start = group.loc[group["year"] == 2024, "adj_ebit"].iloc[0]
    end = group.loc[group["year"] == 2025, "adj_ebit"].iloc[0]

    labels = ["2024"] + list(changes.index) + ["2025"]
    values = [start] + list(changes.values) + [end]

    # running total tells each floating bar where to start
    bottoms, running = [], start
    for i, v in enumerate(values):
        if i == 0 or i == len(values) - 1:
            bottoms.append(0)                      # the two total bars sit on the axis
        else:
            bottoms.append(running if v >= 0 else running + v)
            running += v

    colors = []
    for i, v in enumerate(values):
        if i == 0 or i == len(values) - 1:
            colors.append(ACCENT)
        else:
            colors.append(ACCENT_LIGHT if v >= 0 else NEGATIVE)

    fig, ax = plt.subplots(figsize=(11, 5.5))
    heights = [abs(v) if 0 < i < len(values) - 1 else v for i, v in enumerate(values)]
    ax.bar(labels, heights, bottom=bottoms, color=colors, width=0.62)

    for i, (label, value) in enumerate(zip(labels, values)):
        y = bottoms[i] + abs(value) + 12 if 0 < i < len(values) - 1 else value + 12
        text = f"{value:,.0f}" if i in (0, len(values) - 1) else f"{value:+,.0f}"
        ax.text(i, y, text, ha="center", fontsize=10)

    ax.set_title("Biomaterials wiped out the gains elsewhere:\ngroup adjusted EBIT fell EUR 70m in 2025", loc="left")
    ax.set_ylabel("Adjusted EBIT (EUR million)")
    ax.set_ylim(0, max(start, end) + 140)
    ax.grid(axis="x", visible=False)
    plt.setp(ax.get_xticklabels(), rotation=20, ha="right")
    footnote(fig)
    fig.tight_layout(rect=[0, 0.03, 1, 1])
    fig.savefig(out / "chart2_ebit_bridge.png", bbox_inches="tight")
    plt.close(fig)


# ---------------------------------------------------------------------------
# Chart 3: price vs volume sensitivity for Packaging Solutions
# ---------------------------------------------------------------------------
def chart_sensitivity(seg, sens, out):
    row = seg[(seg["segment"] == "Packaging Solutions") & (seg["year"] == 2025)].iloc[0]
    s = sens[sens["segment"] == "Packaging Solutions"].iloc[0]

    sales, ebit = row["sales"], row["adj_ebit"]
    per_point_price = s["price_10pct"] / 10        # EUR m per 1% price change
    per_point_volume = s["volume_10pct"] / 10

    changes = range(-10, 11)
    price_line = [ebit + c * per_point_price for c in changes]
    volume_line = [ebit + c * per_point_volume for c in changes]
    target_ebit = [MARGIN_TARGET * sales * (1 + c / 100) for c in changes]

    fig, ax = plt.subplots(figsize=(9.5, 5.5))
    ax.plot(changes, price_line, color=ACCENT, linewidth=2.5, label=f"Price change (EUR {per_point_price:.1f}m per 1%)")
    ax.plot(changes, volume_line, color=NEUTRAL, linewidth=2.5, linestyle="--",
            label=f"Volume change (EUR {per_point_volume:.1f}m per 1%)")
    ax.plot(changes, target_ebit, color=TARGET, linewidth=1.5, linestyle=":", label="EBIT needed for a 10% margin")

    ax.axhline(ebit, color="#333333", linewidth=1)
    ax.text(-9.8, ebit + 6, f"2025 actual: EUR {ebit:,.0f}m (1.4% margin)", fontsize=9)

    # where the price line crosses the target line
    breakeven = (MARGIN_TARGET * sales - ebit) / (per_point_price - MARGIN_TARGET * sales / 100)
    ax.plot([breakeven], [ebit + breakeven * per_point_price], "o", color=ACCENT, markersize=7)
    ax.annotate(f"+{breakeven:.1f}% price needed\nto reach a 10% margin",
                xy=(breakeven, ebit + breakeven * per_point_price),
                xytext=(0.5, ebit + 60),
                fontsize=9, arrowprops=dict(arrowstyle="->", color="#333333", lw=1))
    ax.set_ylim(min(price_line) - 12, max(target_ebit) + 15)

    ax.set_title("Price moves Packaging Solutions' profit three times as hard as volume,\n"
                 "but the unit still needs a ~10% price rise to hit target", loc="left")
    ax.set_xlabel("Change vs 2025 (%)")
    ax.set_ylabel("Adjusted EBIT (EUR million)")
    ax.legend(frameon=False, loc="upper left", fontsize=9)
    footnote(fig, SOURCE + " Assumes linear effects and unchanged costs.")
    fig.tight_layout(rect=[0, 0.03, 1, 1])
    fig.savefig(out / "chart3_sensitivity.png", bbox_inches="tight")
    plt.close(fig)

    return breakeven, per_point_price, per_point_volume


# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(description="Build the chart pack from extracted data")
    parser.add_argument("--data", default="output", help="folder holding the extracted CSV files")
    parser.add_argument("--out", default="charts", help="folder for the PNG files")
    args = parser.parse_args()

    data, out = Path(args.data), Path(args.out)
    out.mkdir(parents=True, exist_ok=True)

    seg = pd.read_csv(data / "fact_segment.csv")
    group = pd.read_csv(data / "fact_group.csv")
    sens = pd.read_csv(data / "fact_sensitivity.csv")

    chart_margin_gap(seg, out)
    chart_ebit_bridge(seg, group, out)
    breakeven, price_pt, volume_pt = chart_sensitivity(seg, sens, out)

    ps = seg[(seg["segment"] == "Packaging Solutions") & (seg["year"] == 2025)].iloc[0]
    print("Chart pack written to", out)
    print(f"  1. Margin gap      Packaging Solutions {ps['adj_ebit'] / ps['sales']:.1%} "
          f"vs {MARGIN_TARGET:.0%} target")
    print(f"  2. EBIT bridge     group {group.loc[group.year == 2024, 'adj_ebit'].iloc[0]:,.0f} "
          f"-> {group.loc[group.year == 2025, 'adj_ebit'].iloc[0]:,.0f} EUR m")
    print(f"  3. Sensitivity     price EUR {price_pt:.1f}m per 1% vs volume EUR {volume_pt:.1f}m; "
          f"+{breakeven:.1f}% price needed for a 10% margin")


if __name__ == "__main__":
    main()
