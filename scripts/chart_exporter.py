"""
chart_exporter.py - Generates email-ready PNG charts from Weekend_Cases_Tracker.xlsx.

The exported images mirror the website Charts tab:
  output/charts/weekly_charts.png  - latest weekend, split by Sat/Sun
  output/charts/monthly_charts.png - current month, split by Week 1-5

Called by email_report.py before sending the report email.
"""

from __future__ import annotations

from datetime import date, datetime
from pathlib import Path
import os
import re

import pandas as pd


BASE_DIR = Path(__file__).resolve().parent.parent
OUTPUT_FILE = BASE_DIR / "output" / "Weekend_Cases_Tracker.xlsx"
CHARTS_DIR = BASE_DIR / "output" / "charts"
MPL_DIR = BASE_DIR / "output" / ".matplotlib"
MPL_DIR.mkdir(parents=True, exist_ok=True)
os.environ.setdefault("MPLCONFIGDIR", str(MPL_DIR))

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker


DEFAULT_VERTICAL = ["EMEA", "Cable", "Content", "Enterprise", "Telco", "Software"]
DEFAULT_TECHNOLOGY = ["Routing", "Switching", "Security", "Software"]
DEFAULT_DELIVERY = ["Dispatch P1", "Dispatch P2", "Handover"]

PALETTE = {
    "Cable": "#4472C4",
    "Content": "#FF0000",
    "EMEA": "#70AD47",
    "Emea": "#70AD47",
    "Enterprise": "#FFC000",
    "Software": "#5B9BD5",
    "Telco": "#7030A0",
    "Routing": "#4472C4",
    "Switching": "#ED7D31",
    "Security": "#70AD47",
    "Dispatch P1": "#C00000",
    "Dispatch P2": "#ED7D31",
    "Handover": "#70AD47",
}
DEFAULT_COLORS = [
    "#4472C4",
    "#ED7D31",
    "#A9D18E",
    "#FFC000",
    "#7030A0",
    "#70AD47",
    "#5B9BD5",
    "#C55A11",
]


def _parse_one(value):
    if value is None:
        return pd.NaT
    try:
        if pd.isna(value):
            return pd.NaT
    except Exception:
        pass
    if isinstance(value, (datetime, date)):
        return pd.Timestamp(value)

    text = str(value).strip()
    if not text or text.lower() in ("nat", "none", "nan"):
        return pd.NaT

    padded = re.sub(
        r"^(\d{1})([-/])([A-Za-z]+)([-/])(\d{2,4})$",
        lambda m: f"0{m.group(1)}{m.group(2)}{m.group(3)}{m.group(4)}{m.group(5)}",
        text,
    )
    for candidate in (padded, text):
        for fmt in ("%d-%b-%y", "%d-%b-%Y", "%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
            try:
                return pd.Timestamp(datetime.strptime(candidate, fmt))
            except ValueError:
                pass

    return pd.Timestamp(pd.to_datetime(text, dayfirst=True, errors="coerce"))


def _parse_col(series: pd.Series) -> pd.Series:
    if pd.api.types.is_datetime64_any_dtype(series):
        return pd.to_datetime(series, utc=False, errors="coerce").dt.tz_localize(None)
    return pd.Series([_parse_one(v) for v in series], index=series.index, dtype="datetime64[ns]")


ACRONYMS = {"Emea": "EMEA", "Cfts": "CFTS", "Bngl": "BNGL"}
ALIASES_LOWER = {
    "entfin": "Enterprise",
    "cloud": "Software",
    "new dispatch p1": "Dispatch P1",
    "new dispatch p2": "Dispatch P2",
    "handover in": "Handover",
    "handover-in": "Handover",
}


def _fix_case(value):
    raw = str(value).replace("\xa0", "").strip()
    if not raw or raw.lower() in ("nan", "none", "nat"):
        return ""
    normalized = " ".join(raw.lower().split())
    alias = ALIASES_LOWER.get(normalized)
    if alias:
        return alias
    if normalized.startswith("new ") and "dispatch" in normalized:
        return normalized[4:].strip().title()
    titled = raw.title()
    return ACRONYMS.get(titled, titled)


def _load_data() -> pd.DataFrame:
    if not OUTPUT_FILE.exists():
        return pd.DataFrame()

    frames = []
    for sheet_name in ("Sat", "Sun"):
        try:
            df = pd.read_excel(OUTPUT_FILE, sheet_name=sheet_name, engine="openpyxl")
        except Exception:
            continue
        if df.empty:
            continue

        df["_sheet"] = sheet_name
        df["Date"] = _parse_col(df["Date"]) if "Date" in df.columns else pd.NaT
        for col in ("Vertical", "Technology", "Case Delivery Type"):
            if col in df.columns:
                df[col] = df[col].astype(str).str.strip().apply(_fix_case)
        frames.append(df)

    if not frames:
        return pd.DataFrame()

    combined = pd.concat(frames, ignore_index=True)
    combined = combined[combined["Date"].notna()].copy()
    if combined.empty:
        return combined

    combined["weekday"] = combined["Date"].dt.weekday
    combined["day_name"] = combined["weekday"].map({5: "Sat", 6: "Sun"}).fillna(
        combined["Date"].dt.strftime("%a")
    )
    combined["weekend_start"] = combined["Date"] - pd.to_timedelta(
        (combined["weekday"] - 5) % 7, unit="d"
    )
    combined["week_of_month"] = (combined["Date"].dt.day - 1) // 7 + 1
    combined["week_label"] = "Week " + combined["week_of_month"].astype(str)
    return combined


def _build_timeseries(df, group_col, category_col, defaults, labels):
    if category_col not in df.columns or df.empty:
        return {
            "labels": labels,
            "series": {k: [0] * len(labels) for k in defaults},
            "totals": [0] * len(labels),
        }

    df2 = df[[group_col, category_col]].copy()
    df2 = df2[df2[category_col].astype(str).str.strip() != ""]
    grouped = df2.groupby([group_col, category_col]).size().unstack(fill_value=0)

    categories = list(defaults)
    for cat in grouped.columns.astype(str):
        if cat not in categories:
            categories.append(cat)

    grouped = grouped.reindex(index=labels, columns=categories, fill_value=0)
    return {
        "labels": labels,
        "series": {
            str(c): [int(grouped.at[label, c]) for label in labels]
            for c in categories
        },
        "totals": [int(x) for x in grouped.sum(axis=1)],
    }


def _chart_payload():
    combined = _load_data()
    if combined.empty:
        return None

    latest_weekend = combined["weekend_start"].max()
    weekly_df = combined[combined["weekend_start"] == latest_weekend].copy()

    now = datetime.now()
    monthly_df = combined[
        (combined["Date"].dt.month == now.month)
        & (combined["Date"].dt.year == now.year)
    ].copy()

    weekly_labels = ["Sat", "Sun"]
    monthly_labels = ["Week 1", "Week 2", "Week 3", "Week 4", "Week 5"]

    return {
        "weekly": {
            "title": "Weekly Cases",
            "subtitle": "Latest weekend split by Sat/Sun",
            "vertical": _build_timeseries(weekly_df, "day_name", "Vertical", DEFAULT_VERTICAL, weekly_labels),
            "technology": _build_timeseries(
                weekly_df, "day_name", "Technology", DEFAULT_TECHNOLOGY, weekly_labels
            ),
            "delivery_type": _build_timeseries(
                weekly_df, "day_name", "Case Delivery Type", DEFAULT_DELIVERY, weekly_labels
            ),
        },
        "monthly": {
            "title": "Monthly Cases",
            "subtitle": f"{now.strftime('%B %Y')} split by week",
            "vertical": _build_timeseries(
                monthly_df, "week_label", "Vertical", DEFAULT_VERTICAL, monthly_labels
            ),
            "technology": _build_timeseries(
                monthly_df, "week_label", "Technology", DEFAULT_TECHNOLOGY, monthly_labels
            ),
            "delivery_type": _build_timeseries(
                monthly_df, "week_label", "Case Delivery Type", DEFAULT_DELIVERY, monthly_labels
            ),
        },
    }


def _draw_line_chart(ax, chart_data, title):
    labels = chart_data["labels"]
    plotted = False

    for idx, (category, values) in enumerate(chart_data["series"].items()):
        if not any(values):
            continue
        color = PALETTE.get(category, DEFAULT_COLORS[idx % len(DEFAULT_COLORS)])
        ax.plot(
            labels,
            values,
            marker="o",
            linewidth=2.5,
            markersize=6,
            color=color,
            label=category,
        )
        plotted = True

    ax.set_title(title, fontsize=12, fontweight="bold", pad=8)
    ax.set_xlabel("Period", fontsize=9)
    ax.set_ylabel("Cases", fontsize=9)
    ax.yaxis.set_major_locator(mticker.MaxNLocator(integer=True))
    ax.grid(True, axis="y", color="#E5E7EB", linewidth=0.8)
    ax.spines[["top", "right"]].set_visible(False)
    ax.tick_params(axis="x", labelsize=9)
    ax.tick_params(axis="y", labelsize=9)
    ax.set_ylim(bottom=0)

    if plotted:
        ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.18), ncol=3, fontsize=8, frameon=False)
    else:
        ax.text(0.5, 0.5, "No data", ha="center", va="center", transform=ax.transAxes, color="#6B7280")


def _render_period_image(period_key: str, period_data: dict) -> Path:
    CHARTS_DIR.mkdir(parents=True, exist_ok=True)

    fig, axes = plt.subplots(3, 1, figsize=(12, 12))
    fig.patch.set_facecolor("#F8F9FA")
    fig.suptitle(period_data["title"], fontsize=17, fontweight="bold", color="#111827", y=0.985)
    fig.text(0.5, 0.955, period_data["subtitle"], ha="center", fontsize=10, color="#6B7280")

    _draw_line_chart(axes[0], period_data["vertical"], "Cases by Vertical")
    _draw_line_chart(axes[1], period_data["technology"], "Cases by Technology")
    _draw_line_chart(axes[2], period_data["delivery_type"], "Cases by Delivery Type")

    fig.subplots_adjust(left=0.08, right=0.97, top=0.91, bottom=0.08, hspace=0.58)
    out = CHARTS_DIR / f"{period_key}_charts.png"
    fig.savefig(out, dpi=140, bbox_inches="tight", facecolor=fig.get_facecolor())
    plt.close(fig)
    return out


def generate_all_charts() -> list[Path]:
    """
    Generate the same Weekly and Monthly chart views shown in the website.
    Returns existing PNG paths for email attachment.
    """
    payload = _chart_payload()
    if not payload:
        return []

    paths = []
    for key in ("weekly", "monthly"):
        try:
            path = _render_period_image(key, payload[key])
            if path.exists():
                paths.append(path)
        except Exception as exc:
            print(f"Chart export warning ({key}): {exc}")
    return paths
