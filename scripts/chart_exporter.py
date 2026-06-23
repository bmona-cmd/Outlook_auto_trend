"""
chart_exporter.py  —  Generates PNG chart images from Weekend_Cases_Tracker.xlsx.

Produces:
  output/charts/saturday_charts.png   — Vertical / Technology / Delivery breakdown for Saturday
  output/charts/sunday_charts.png     — Same for Sunday
  output/charts/monthly_summary.png   — Monthly totals + breakdown trend

Called by email_report.py before sending the report email.
"""

from __future__ import annotations

from pathlib import Path
from datetime import datetime
import os

import pandas as pd


# ── paths ────────────────────────────────────────────────────────────────────

BASE_DIR    = Path(__file__).resolve().parent.parent
OUTPUT_FILE = BASE_DIR / "output" / "Weekend_Cases_Tracker.xlsx"
CHARTS_DIR  = BASE_DIR / "output" / "charts"
MPL_DIR     = BASE_DIR / "output" / ".matplotlib"
MPL_DIR.mkdir(parents=True, exist_ok=True)
os.environ.setdefault("MPLCONFIGDIR", str(MPL_DIR))

import matplotlib
matplotlib.use("Agg")           # headless — no display needed
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
from matplotlib.gridspec import GridSpec

# HPE brand-adjacent palette
PALETTE = [
    "#01A982", "#FF8300", "#C140FF", "#00739D",
    "#FEC901", "#FF3C3C", "#00B388", "#5F249F",
]

SHEET_MAP = {"Sat": "Saturday", "Sun": "Sunday"}


# ── helpers ──────────────────────────────────────────────────────────────────

def _load_sheet(sheet: str) -> pd.DataFrame:
    """Return a DataFrame for the given sheet, or empty DataFrame."""
    if not OUTPUT_FILE.exists():
        return pd.DataFrame()
    try:
        df = pd.read_excel(OUTPUT_FILE, sheet_name=sheet, engine="openpyxl")
        return df if not df.empty else pd.DataFrame()
    except Exception:
        return pd.DataFrame()


def _value_counts(df: pd.DataFrame, col: str) -> pd.Series:
    if df.empty or col not in df.columns:
        return pd.Series(dtype=int)
    return (
        df[col]
        .dropna()
        .astype(str)
        .str.strip()
        .replace("", pd.NA)
        .dropna()
        .value_counts()
        .sort_index()
    )


def _parse_dates(series: pd.Series) -> pd.Series:
    return pd.to_datetime(series, errors="coerce", format="mixed")


def _bar(ax, series: pd.Series, title: str, color_list: list):
    """Draw a compact horizontal bar chart."""
    if series.empty:
        ax.text(0.5, 0.5, "No data", ha="center", va="center",
                transform=ax.transAxes, color="#888", fontsize=10)
        ax.set_title(title, fontsize=11, fontweight="bold", pad=6)
        ax.axis("off")
        return

    colors = [color_list[i % len(color_list)] for i in range(len(series))]
    bars = ax.barh(series.index, series.values, color=colors, height=0.55, edgecolor="none")

    for bar, val in zip(bars, series.values):
        ax.text(bar.get_width() + 0.1, bar.get_y() + bar.get_height() / 2,
                str(val), va="center", ha="left", fontsize=9, color="#333")

    ax.set_title(title, fontsize=11, fontweight="bold", pad=6)
    ax.xaxis.set_major_locator(mticker.MaxNLocator(integer=True))
    ax.set_xlabel("Cases", fontsize=8, color="#555")
    ax.tick_params(axis="y", labelsize=9)
    ax.tick_params(axis="x", labelsize=8)
    ax.spines[["top", "right"]].set_visible(False)
    ax.set_xlim(0, max(series.values) * 1.25)


# ── per-day chart ─────────────────────────────────────────────────────────────

def _day_chart(sheet: str, day_name: str) -> Path | None:
    """Generate a 3-panel chart for one weekend day. Returns the saved PNG path."""
    CHARTS_DIR.mkdir(parents=True, exist_ok=True)
    df = _load_sheet(sheet)

    fig = plt.figure(figsize=(14, 5))
    fig.patch.set_facecolor("#F8F9FA")

    date_str = ""
    if not df.empty and "Date" in df.columns:
        dates = _parse_dates(df["Date"]).dropna()
        if not dates.empty:
            date_str = f"  —  {dates.max().strftime('%d %b %Y')}"

    fig.suptitle(
        f"{day_name} Cases Summary{date_str}",
        fontsize=14, fontweight="bold", color="#1A1A2E", y=1.02
    )

    gs = GridSpec(1, 3, figure=fig, wspace=0.45)
    axes = [fig.add_subplot(gs[0, i]) for i in range(3)]

    cols_titles = [
        ("Vertical",           "By Vertical"),
        ("Technology",         "By Technology"),
        ("Case Delivery Type", "By Delivery Type"),
    ]

    for ax, (col, title) in zip(axes, cols_titles):
        _bar(ax, _value_counts(df, col), title, PALETTE)

    # row count badge
    total = len(df) if not df.empty else 0
    fig.text(
        0.99, 0.98, f"Total cases: {total}",
        ha="right", va="top", fontsize=9,
        color="#555", style="italic"
    )

    fig.subplots_adjust(left=0.08, right=0.95, bottom=0.18, top=0.82, wspace=0.45)
    out = CHARTS_DIR / f"{sheet.lower()}_charts.png"
    fig.savefig(out, dpi=130, bbox_inches="tight", facecolor=fig.get_facecolor())
    plt.close(fig)
    return out


# ── monthly summary chart ─────────────────────────────────────────────────────

def _monthly_chart() -> Path | None:
    """
    Generate a stacked-bar monthly summary across both sheets.
    X-axis = month labels, bars stacked by Vertical.
    """
    CHARTS_DIR.mkdir(parents=True, exist_ok=True)

    frames = []
    for sheet in ("Sat", "Sun"):
        df = _load_sheet(sheet)
        if not df.empty:
            frames.append(df)

    if not frames:
        return None

    combined = pd.concat(frames, ignore_index=True)
    if "Date" not in combined.columns:
        return None

    combined["_date"] = _parse_dates(combined["Date"])
    combined = combined[combined["_date"].notna()].copy()
    if combined.empty:
        return None

    combined["_month"] = combined["_date"].dt.to_period("M")
    combined["_month_label"] = combined["_date"].dt.strftime("%b %Y")

    month_order = (
        combined[["_month", "_month_label"]]
        .drop_duplicates()
        .sort_values("_month")["_month_label"]
        .tolist()
    )

    # ── figure: 2 rows — top: monthly totals bar, bottom: vertical stacked ───
    fig, (ax_top, ax_bot) = plt.subplots(2, 1, figsize=(max(8, len(month_order) * 1.4 + 2), 9))
    fig.patch.set_facecolor("#F8F9FA")
    fig.suptitle("Monthly Cases Summary", fontsize=14, fontweight="bold", color="#1A1A2E")

    # top — total per month
    totals = [int((combined["_month_label"] == m).sum()) for m in month_order]
    bars = ax_top.bar(month_order, totals, color=PALETTE[0], width=0.55, edgecolor="none")
    for bar, val in zip(bars, totals):
        ax_top.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.2,
                    str(val), ha="center", va="bottom", fontsize=9, fontweight="bold")
    ax_top.set_title("Total Cases per Month", fontsize=11, fontweight="bold")
    ax_top.set_ylabel("Cases", fontsize=9)
    ax_top.tick_params(axis="x", rotation=30, labelsize=9)
    ax_top.spines[["top", "right"]].set_visible(False)
    ax_top.yaxis.set_major_locator(mticker.MaxNLocator(integer=True))

    # bottom — stacked by Vertical
    if "Vertical" in combined.columns:
        pivot = (
            combined.assign(Vertical=combined["Vertical"].fillna("Unknown").astype(str).str.strip())
            .groupby(["_month_label", "Vertical"])
            .size()
            .unstack(fill_value=0)
            .reindex(index=month_order, fill_value=0)
        )
        bottom = [0] * len(month_order)
        for i, col in enumerate(pivot.columns):
            vals = pivot[col].tolist()
            ax_bot.bar(month_order, vals, bottom=bottom, label=col,
                       color=PALETTE[i % len(PALETTE)], width=0.55, edgecolor="none")
            bottom = [b + v for b, v in zip(bottom, vals)]

        ax_bot.set_title("Cases by Vertical (Monthly)", fontsize=11, fontweight="bold")
        ax_bot.set_ylabel("Cases", fontsize=9)
        ax_bot.tick_params(axis="x", rotation=30, labelsize=9)
        ax_bot.spines[["top", "right"]].set_visible(False)
        ax_bot.yaxis.set_major_locator(mticker.MaxNLocator(integer=True))
        ax_bot.legend(loc="upper left", fontsize=8, framealpha=0.7, ncol=2)
    else:
        ax_bot.axis("off")

    fig.tight_layout()
    out = CHARTS_DIR / "monthly_summary.png"
    fig.savefig(out, dpi=130, bbox_inches="tight", facecolor=fig.get_facecolor())
    plt.close(fig)
    return out


# ── public API ────────────────────────────────────────────────────────────────

def generate_all_charts() -> list[Path]:
    """
    Generate all chart images and return a list of existing PNG paths.
    Safe to call even if the Excel file is missing — returns [].
    """
    paths = []
    try:
        sat = _day_chart("Sat", "Saturday")
        if sat and sat.exists():
            paths.append(sat)
    except Exception as e:
        print(f"Chart export warning (Saturday): {e}")

    try:
        sun = _day_chart("Sun", "Sunday")
        if sun and sun.exists():
            paths.append(sun)
    except Exception as e:
        print(f"Chart export warning (Sunday): {e}")

    try:
        mon = _monthly_chart()
        if mon and mon.exists():
            paths.append(mon)
    except Exception as e:
        print(f"Chart export warning (Monthly): {e}")

    return paths
