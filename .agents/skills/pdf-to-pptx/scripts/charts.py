"""Phase 2: Data-driven 300 DPI transparent-background chart generator."""

from __future__ import annotations

import argparse
import json
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np


THEMES_PATH = Path(__file__).resolve().parent.parent / "references" / "themes.json"


def load_palette(theme_key: str, brand_path: Path | None) -> tuple[list[str], str, str]:
    """Load chart color palette, dark text color, and font family."""
    themes = json.loads(THEMES_PATH.read_text(encoding="utf-8"))
    theme = themes.get(theme_key, themes["classic_navy_teal"])
    if brand_path and brand_path.exists():
        brand = json.loads(brand_path.read_text(encoding="utf-8"))
        theme = {**theme, **brand}
    palette = theme.get("chart_palette", ["#1E3A8A", "#0D9488", "#E05A47", "#D97706"])
    dark = theme.get("dark", "#0F172A")
    font = theme.get("font_body", "DejaVu Sans")
    return palette, dark, font


def style_axes(ax: plt.Axes, dark: str) -> None:
    """Apply clean executive styling: no top/right spines, subtle grid."""
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_color("#CBD5E1")
    ax.spines["bottom"].set_color("#CBD5E1")
    ax.tick_params(colors=dark, labelsize=11)
    ax.yaxis.grid(True, linestyle="--", alpha=0.35, color="#94A3B8")
    ax.set_axisbelow(True)


def render_chart(spec: dict, out_path: Path, palette: list[str], dark: str, font: str) -> None:
    """Render a single chart from a chart_plan spec."""
    plt.rcParams["font.sans-serif"] = [font, "DejaVu Sans", "Arial"]
    chart_type = spec.get("type", "bar")
    categories = spec.get("categories", [])
    series = spec.get("series", [])
    fmt = spec.get("value_format", "{:.1f}")
    title = spec.get("title", "")

    fig, ax = plt.subplots(figsize=(7.2, 4.6), dpi=300)
    fig.patch.set_alpha(0.0)
    ax.patch.set_alpha(0.0)

    if chart_type == "donut":
        values = series[0]["values"] if series else []
        colors = [palette[i % len(palette)] for i in range(len(values))]
        wedges, _, autotexts = ax.pie(
            values,
            labels=categories,
            autopct="%1.1f%%",
            startangle=90,
            colors=colors,
            pctdistance=0.78,
            wedgeprops={"width": 0.42, "edgecolor": "white", "linewidth": 2},
            textprops={"fontsize": 11, "color": dark},
        )
        for at in autotexts:
            at.set_color("white")
            at.set_fontweight("bold")
            at.set_fontsize(10)
        ax.axis("equal")

    elif chart_type == "line":
        style_axes(ax, dark)
        x = np.arange(len(categories))
        for idx, s in enumerate(series):
            color = palette[idx % len(palette)]
            vals = s["values"]
            ax.plot(x, vals, marker="o", linewidth=2.6, markersize=7, color=color, label=s.get("name", ""))
            for xi, yi in zip(x, vals):
                ax.annotate(
                    fmt.format(yi),
                    (xi, yi),
                    textcoords="offset points",
                    xytext=(0, 8),
                    ha="center",
                    fontsize=10,
                    fontweight="bold",
                    color=dark,
                )
        ax.set_xticks(x)
        ax.set_xticklabels(categories)
        if len(series) > 1:
            ax.legend(frameon=False, fontsize=10, loc="upper left")

    elif chart_type == "barh":
        style_axes(ax, dark)
        ax.yaxis.grid(False)
        ax.xaxis.grid(True, linestyle="--", alpha=0.35, color="#94A3B8")
        y = np.arange(len(categories))
        n_series = max(len(series), 1)
        height = 0.72 / n_series
        for idx, s in enumerate(series):
            offset = (idx - (n_series - 1) / 2) * height
            color = palette[idx % len(palette)]
            bars = ax.barh(y + offset, s["values"], height=height, color=color, label=s.get("name", ""))
            for bar in bars:
                w = bar.get_width()
                ax.text(
                    w,
                    bar.get_y() + bar.get_height() / 2,
                    f" {fmt.format(w)}",
                    va="center",
                    ha="left",
                    fontsize=10,
                    fontweight="bold",
                    color=dark,
                )
        ax.set_yticks(y)
        ax.set_yticklabels(categories)
        ax.invert_yaxis()
        if len(series) > 1:
            ax.legend(frameon=False, fontsize=10, loc="lower right")

    else:  # vertical bar / grouped bar
        style_axes(ax, dark)
        x = np.arange(len(categories))
        n_series = max(len(series), 1)
        width = 0.72 / n_series
        for idx, s in enumerate(series):
            offset = (idx - (n_series - 1) / 2) * width
            color = palette[idx % len(palette)]
            bars = ax.bar(x + offset, s["values"], width=width, color=color, label=s.get("name", ""))
            for bar in bars:
                h = bar.get_height()
                ax.text(
                    bar.get_x() + bar.get_width() / 2,
                    h,
                    fmt.format(h),
                    va="bottom",
                    ha="center",
                    fontsize=10,
                    fontweight="bold",
                    color=dark,
                )
        ax.set_xticks(x)
        ax.set_xticklabels(categories)
        if len(series) > 1:
            ax.legend(frameon=False, fontsize=10, loc="upper left")

    if title:
        ax.set_title(title, fontsize=13, fontweight="bold", color=dark, pad=14)
    if spec.get("xlabel") and chart_type != "donut":
        ax.set_xlabel(spec["xlabel"], fontsize=11, color=dark, labelpad=8)
    if spec.get("ylabel") and chart_type != "donut":
        ax.set_ylabel(spec["ylabel"], fontsize=11, color=dark, labelpad=8)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(out_path, dpi=300, transparent=True, bbox_inches="tight", pad_inches=0.18)
    plt.close(fig)


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate 300 DPI charts from chart_plan JSON.")
    parser.add_argument("--deck-dir", required=True, help="Target deck directory")
    parser.add_argument("--plan", required=True, help="Path to chart_plan JSON (relative to deck-dir or absolute)")
    args = parser.parse_args()

    deck_dir = Path(args.deck_dir).resolve()
    plan_path = Path(args.plan) if Path(args.plan).is_absolute() else deck_dir / args.plan
    plan = json.loads(plan_path.read_text(encoding="utf-8"))

    brand_file = plan.get("brand_file", "brand.json")
    brand_path = deck_dir / brand_file if brand_file else None
    palette, dark, font = load_palette(plan.get("theme", "classic_navy_teal"), brand_path)

    generated: list[str] = []
    for chart in plan.get("charts", []):
        out_rel = chart["output_file"]
        out_path = Path(out_rel) if Path(out_rel).is_absolute() else deck_dir / out_rel
        render_chart(chart, out_path, palette, dark, font)
        generated.append(str(out_rel))

    print(json.dumps({"status": "ok", "charts_generated": len(generated), "files": generated}))


if __name__ == "__main__":
    main()
