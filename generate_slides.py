#!/usr/bin/env python3
"""
generate_slides.py  — v2
Redesigned to match Analysis_Story.pdf visual quality.

8 slides  |  Storytelling with Data  |  Situation -> Complication -> Resolution
Big Idea: America's longevity problem is not a budget problem — it is upstream of the budget.
"""

import os, io
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# ── Slide dimensions ──────────────────────────────────────────────────────────
W = Inches(13.333)
H = Inches(7.5)

# ── Color palette (matches Analysis_Story.pdf) ────────────────────────────────
C_NAVY    = "#1B2A45"   # dark slide background
C_RED     = "#C0392B"   # United States / bad metric
C_TEAL    = "#2BA88A"   # Japan / positive metric
C_WHITE   = "#FFFFFF"
C_DARK    = "#0F1F35"   # near-black headers on white slides
C_BODY    = "#4B5563"   # body text on white slides
C_SECTION = "#2BA88A"   # section label colour (= teal)
C_SUB     = "#94A3B8"   # muted subtext on dark slides
C_MUTED   = "#5D7A8F"   # "how much" in takeaway headline
C_MIDGRAY = "#D1D5DB"

# Chart-specific palette
CH_US    = "#C0392B"
CH_JP    = "#2BA88A"
CH_SP    = "#D4980A"   # Spain gold
CH_GRAY  = "#AAAAAA"   # all-country dots
CH_TREND = "#CCCCCC"   # dashed trend line
CH_BAR   = "#D1D5DB"   # peer country bars


# ── Core helpers ──────────────────────────────────────────────────────────────

def _rgb(h):
    h = h.lstrip("#")
    return RGBColor(int(h[:2], 16), int(h[2:4], 16), int(h[4:], 16))


def _blank(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def _bg(slide, color):
    f = slide.background.fill
    f.solid()
    f.fore_color.rgb = _rgb(color)


def _rect(slide, l, t, w, h, fill, line=None):
    shape = slide.shapes.add_shape(1, l, t, w, h)
    shape.fill.solid()
    shape.fill.fore_color.rgb = _rgb(fill)
    if line:
        shape.line.color.rgb = _rgb(line)
        shape.line.width = Pt(0.75)
    else:
        shape.line.fill.background()
    return shape


def _oval(slide, l, t, w, h, fill):
    shape = slide.shapes.add_shape(9, l, t, w, h)
    shape.fill.solid()
    shape.fill.fore_color.rgb = _rgb(fill)
    shape.line.fill.background()
    return shape


def _txt(slide, text, l, t, w, h, size=12, bold=False, italic=False,
         color=C_DARK, align=PP_ALIGN.LEFT):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    for i, line in enumerate(text.split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run()
        r.text = line
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.italic = italic
        r.font.color.rgb = _rgb(color)
    return box


def _mixed_box(slide, l, t, w, h):
    """Return a text frame ready for mixed-colour paragraphs."""
    box = slide.shapes.add_textbox(l, t, w, h)
    box.text_frame.word_wrap = True
    return box.text_frame


def _para(tf, parts, size, bold=True, space_before=0, space_after=0, align=PP_ALIGN.LEFT):
    """Append a paragraph with mixed-colour runs.  parts = [(text, color), ...]"""
    if len(tf.paragraphs) == 1 and not tf.paragraphs[0].runs:
        p = tf.paragraphs[0]
    else:
        p = tf.add_paragraph()
    p.alignment = align
    p.space_before = Pt(space_before)
    p.space_after = Pt(space_after)
    for text, color in parts:
        if text:
            r = p.add_run()
            r.text = text
            r.font.size = Pt(size)
            r.font.bold = bold
            r.font.color.rgb = _rgb(color)
    return p


def _to_img(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight",
                facecolor="white", edgecolor="none")
    buf.seek(0)
    plt.close(fig)
    return buf


def _pic(slide, img, l, t, w, h):
    slide.shapes.add_picture(img, l, t, w, h)


def _mpl_clean(ax, xlabel=None, ylabel=None, x_grid=False):
    ax.set_facecolor("white")
    ax.figure.patch.set_facecolor("white")
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_color("#E5E7EB")
    ax.spines["bottom"].set_color("#E5E7EB")
    ax.tick_params(colors="#9CA3AF", labelsize=9, length=2)
    if xlabel:
        ax.set_xlabel(xlabel, fontsize=9, color="#9CA3AF", labelpad=6)
    if ylabel:
        ax.set_ylabel(ylabel, fontsize=9, color="#9CA3AF", labelpad=6)
    if x_grid:
        ax.grid(axis="x", color="#F3F4F6", linewidth=0.8, zorder=0)


def _section(slide, text):
    _txt(slide, text, Inches(0.45), Inches(0.27), Inches(9.0), Inches(0.32),
         size=9, bold=True, color=C_SECTION)


def _footer(slide, text):
    _txt(slide, text, Inches(0.45), Inches(7.17), Inches(12.4), Inches(0.28),
         size=7.5, italic=True, color="#9CA3AF")


def _dot_legend(slide, x, label, color):
    """Small filled square + label for chart legends."""
    _rect(slide, x, Inches(1.82), Inches(0.2), Inches(0.2), color)
    _txt(slide, label, x + Inches(0.28), Inches(1.79), Inches(1.8), Inches(0.3),
         size=9.5, bold=True, color=color)


# ── Slide 1 — Cover ───────────────────────────────────────────────────────────

def slide_cover(prs):
    s = _blank(prs)
    _bg(s, C_NAVY)

    _txt(s, "A DATA STORY  ·  HEALTH SPENDING vs. LONGEVITY",
         Inches(0.45), Inches(0.38), Inches(12.0), Inches(0.32),
         size=9, bold=True, color=C_SUB)

    # ── Big mixed-colour headline ──
    tf = _mixed_box(s, Inches(0.45), Inches(1.2), Inches(12.3), Inches(3.1))
    _para(tf, [("America’s longevity problem ", C_WHITE),
               ("isn’t a", C_RED)],
          size=50, bold=True, space_after=0)
    _para(tf, [("budget problem.", C_RED)],
          size=50, bold=True, space_after=0)
    _para(tf, [("It’s ", C_WHITE),
               ("upstream", C_TEAL),
               (" of the budget.", C_WHITE)],
          size=50, bold=True)

    # ── Body ──
    _txt(s,
         "Longer lives track education, lifestyle, demographics and how care is priced —\n"
         "and on every one of these, the country that spends the most falls behind.\n"
         "Here is what the highest spenders are doing differently.",
         Inches(0.45), Inches(4.9), Inches(9.8), Inches(1.3),
         size=13.5, color="#B8C8D8")

    # ── Thin rule ──
    _rect(s, Inches(0.45), Inches(6.42), Inches(5.5), Inches(0.025), "#2B4B75")

    # ── Bottom tag ──
    _oval(s, Inches(0.45), Inches(6.62), Inches(0.22), Inches(0.22), C_RED)
    _txt(s, "United States", Inches(0.76), Inches(6.59), Inches(2.0), Inches(0.32),
         size=11, bold=True, color=C_RED)
    _txt(s, "vs. the world’s 14 highest-income health spenders",
         Inches(2.85), Inches(6.59), Inches(8.5), Inches(0.32),
         size=11, color=C_SUB)

    _footer(s,
            "Built on the attached Power BI analysis (World Bank WDI). "
            "New factors: OECD · WHO · World Bank.")


# ── Slide 2 — THE SETUP ───────────────────────────────────────────────────────

def _chart_setup(df):
    sub = df.dropna(subset=["health_expenditure_ppp", "life_expectancy"]).copy()
    lx = np.log10(sub["health_expenditure_ppp"])
    le = sub["life_expectancy"]

    fig, ax = plt.subplots(figsize=(7.6, 5.0))

    # All countries — light gray dots
    ax.scatter(lx, le, color=CH_GRAY, alpha=0.30, s=18, zorder=2)

    # Trend line — dashed gray
    coef = np.polyfit(lx, le, 1)
    xfit = np.linspace(lx.min(), lx.max(), 300)
    ax.plot(xfit, np.poly1d(coef)(xfit), color=CH_TREND,
            lw=1.6, linestyle="--", zorder=1)

    def _highlight(country, color, s, label_txt, dx_pt, dy_pt):
        row = sub[sub["country"] == country]
        if row.empty:
            return
        cx = float(np.log10(row["health_expenditure_ppp"].iloc[0]))
        cy = float(row["life_expectancy"].iloc[0])
        ax.scatter([cx], [cy], color=color, s=s, zorder=5)
        ax.annotate(label_txt, xy=(cx, cy),
                    xytext=(dx_pt, dy_pt), textcoords="offset points",
                    fontsize=8.5, color=color, fontweight="bold",
                    arrowprops=dict(arrowstyle="->", color=color, lw=1.1))

    _highlight("Spain", CH_SP, 80,
               "Spain\n$3,050 · 82 yrs", -60, 18)
    _highlight("Japan", CH_JP, 100,
               "Japan\n$3,640 · 83 yrs", 8, 12)
    _highlight("United States", CH_US, 110,
               "United States\n$8,431 · 78 yrs", 6, -38)

    ax.set_xticks([np.log10(100), np.log10(1000), np.log10(10000)])
    ax.set_xticklabels(["$100", "$1,000", "$10,000"])
    ax.set_xlim(1.65, 4.18)
    ax.set_ylim(49, 88)
    ax.text(0.015, 0.025,
            "Each gray dot = one of 189 countries · 2000–2023 averages",
            transform=ax.transAxes, fontsize=7.5, color="#AAAAAA", style="italic")
    _mpl_clean(ax,
               xlabel="Avg. health spending per person, PPP (log scale)",
               ylabel="Avg. life expectancy (years)")
    fig.tight_layout(pad=0.4)
    return fig


def slide_setup(prs, df):
    s = _blank(prs)
    _bg(s, C_WHITE)
    _section(s, "THE SETUP")

    tf = _mixed_box(s, Inches(0.45), Inches(0.58), Inches(12.3), Inches(1.05))
    _para(tf, [("The biggest spender on Earth buys ", C_DARK),
               ("one of the shortest lives", C_RED),
               (" among rich nations", C_DARK)],
          size=27, bold=True)

    _pic(s, _to_img(_chart_setup(df)),
         Inches(0.3), Inches(1.65), Inches(8.75), Inches(5.6))

    rx, rw = Inches(9.2), Inches(3.95)

    tf2 = _mixed_box(s, rx, Inches(1.9), rw, Inches(0.75))
    _para(tf2, [("Spending buys longevity — until it doesn’t.", C_DARK)],
          size=14, bold=True)

    tf3 = _mixed_box(s, rx, Inches(2.75), rw, Inches(2.6))
    _para(tf3,
          [("Above roughly $4,000 per person the curve flattens. Yet the ", C_BODY),
           ("U.S. spends ~$8,400", C_RED),
           (" and still lands ", C_BODY),
           ("5 years below Japan", C_TEAL),
           (", which spends less than half as much.", C_BODY)],
          size=12, bold=False)
    _para(tf3, [("", C_BODY)], size=5, bold=False)
    _para(tf3,
          [("So the question isn’t how much to spend. "
            "It’s what those dollars sit on top of.", C_DARK)],
          size=12, bold=True)

    _footer(s,
            "Source: World Bank, World Development Indicators "
            "(health expenditure & life expectancy, 2000–2023 averages). "
            "Peer comparison restricted to the world’s highest-spending countries.")


# ── Slide 3 — FACTOR 1: EDUCATION QUALITY ────────────────────────────────────

def _chart_education(df):
    hi = df[df["health_expenditure_ppp"] >= 4000].dropna(
        subset=["pisa_2022_score", "life_expectancy"]).copy()

    fig, ax = plt.subplots(figsize=(7.6, 5.0))

    other = hi[~hi["country"].isin(["United States", "Japan"])]
    ax.scatter(other["pisa_2022_score"], other["life_expectancy"],
               color=CH_GRAY, alpha=0.55, s=55, zorder=2)

    for _, row in other.iterrows():
        ax.annotate(row["country"],
                    xy=(row["pisa_2022_score"], row["life_expectancy"]),
                    xytext=(4, 0), textcoords="offset points",
                    fontsize=7.5, color="#888888")

    us = hi[hi["country"] == "United States"]
    if not us.empty:
        ux = float(us["pisa_2022_score"].iloc[0])
        uy = float(us["life_expectancy"].iloc[0])
        ax.axvline(x=ux, color=CH_US, lw=0.9, linestyle="--", alpha=0.45, zorder=1)
        ax.scatter([ux], [uy], color=CH_US, s=120, zorder=5)
        ax.annotate("United States", xy=(ux, uy),
                    xytext=(5, -14), textcoords="offset points",
                    fontsize=9, color=CH_US, fontweight="bold")

    jp = hi[hi["country"] == "Japan"]
    if not jp.empty:
        jx = float(jp["pisa_2022_score"].iloc[0])
        jy = float(jp["life_expectancy"].iloc[0])
        ax.scatter([jx], [jy], color=CH_JP, s=120, zorder=5)
        ax.annotate("Japan", xy=(jx, jy),
                    xytext=(6, 2), textcoords="offset points",
                    fontsize=9, color=CH_JP, fontweight="bold")

    ax.set_xlim(461, 542)
    ax.set_ylim(77.0, 84.5)
    _mpl_clean(ax,
               xlabel="Education quality — PISA 2022 mean score (math · reading · science)",
               ylabel="Avg. life expectancy (years)")
    fig.tight_layout(pad=0.4)
    return fig


def slide_education(prs, df):
    s = _blank(prs)
    _bg(s, C_WHITE)
    _section(s, "FACTOR 1 — EDUCATION QUALITY")

    tf = _mixed_box(s, Inches(0.45), Inches(0.58), Inches(12.3), Inches(1.05))
    _para(tf, [("Countries that ", C_DARK),
               ("out-educate", C_TEAL),
               (" the U.S. tend to ", C_DARK),
               ("outlive", C_TEAL),
               (" it", C_DARK)],
          size=27, bold=True)

    _pic(s, _to_img(_chart_education(df)),
         Inches(0.3), Inches(1.65), Inches(8.75), Inches(5.6))

    _dot_legend(s, Inches(9.2), "United States", CH_US)
    _dot_legend(s, Inches(11.15), "Japan", CH_JP)

    rx, rw = Inches(9.2), Inches(3.95)

    tf2 = _mixed_box(s, rx, Inches(2.3), rw, Inches(0.85))
    _para(tf2, [("Japan tops the peer group at 533;\n"
                 "the U.S. sits mid-pack at 489.", C_TEAL)],
          size=13.5, bold=True)

    _txt(s,
         "Skills predict the behaviors — diet,\nscreening, adherence — that compound\ninto longer lives. The high-education,\nhigh-longevity countries cluster toward\nthe upper right; the U.S. does not.",
         rx, Inches(3.3), rw, Inches(2.5), size=12, color=C_BODY)

    _footer(s,
            "Source: OECD PISA 2022 (mean of mathematics, reading & science). "
            "Life expectancy: World Bank WDI, 2000–2023 avg.")


# ── Slide 4 — FACTOR 2: LIFESTYLE ────────────────────────────────────────────

def _chart_lifestyle(df):
    order = ["United States", "Ireland", "Canada", "Germany", "Belgium",
             "Luxembourg", "Norway", "Spain", "Austria", "Netherlands",
             "Sweden", "Denmark", "Switzerland", "France", "Japan"]
    sub = (df[df["country"].isin(order)]
             .dropna(subset=["obesity_rate"])
             .set_index("country")
             .reindex(order)
             .reset_index())
    sub = sub.sort_values("obesity_rate", ascending=True)   # ascending for barh

    colors = [CH_US if c == "United States" else
              (CH_JP if c == "Japan" else CH_BAR)
              for c in sub["country"]]

    fig, ax = plt.subplots(figsize=(7.6, 5.2))
    bars = ax.barh(sub["country"], sub["obesity_rate"],
                   color=colors, height=0.68, zorder=3)

    for bar, val, cntry in zip(bars, sub["obesity_rate"], sub["country"]):
        fc = CH_US if cntry == "United States" else (CH_JP if cntry == "Japan" else "#6B7280")
        bw = "bold" if cntry in ("United States", "Japan") else "normal"
        ax.text(bar.get_width() + 0.4,
                bar.get_y() + bar.get_height() / 2,
                f"{val:.1f}%", va="center", fontsize=8.5,
                color=fc, fontweight=bw)

    ax.set_xlim(0, 53)
    ax.xaxis.set_major_formatter(
        mticker.FuncFormatter(lambda x, _: f"{x:.0f}%"))
    ax.set_facecolor("white")
    ax.figure.patch.set_facecolor("white")
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_visible(False)
    ax.spines["bottom"].set_color("#E5E7EB")
    ax.tick_params(colors="#9CA3AF", labelsize=9, length=0)
    ax.set_xlabel("Adults with obesity, BMI ≥ 30 (%)", fontsize=9, color="#9CA3AF")
    ax.grid(axis="x", color="#F3F4F6", linewidth=0.8, zorder=0)
    fig.tight_layout(pad=0.4)
    return fig


def slide_lifestyle(prs, df):
    s = _blank(prs)
    _bg(s, C_WHITE)
    _section(s, "FACTOR 2 — LIFESTYLE")

    tf = _mixed_box(s, Inches(0.45), Inches(0.58), Inches(12.3), Inches(1.05))
    _para(tf, [("The gap is written in waistlines: ", C_DARK),
               ("43% obese vs. 5%", C_RED)],
          size=27, bold=True)

    _pic(s, _to_img(_chart_lifestyle(df)),
         Inches(0.3), Inches(1.6), Inches(8.75), Inches(5.65))

    rx, rw = Inches(9.2), Inches(3.95)

    tf2 = _mixed_box(s, rx, Inches(2.0), rw, Inches(0.95))
    _para(tf2, [("Nearly 1 in 2 American adults", C_RED),
                (" has obesity —\n", C_DARK),
                ("almost 9× Japan’s rate.", C_TEAL)],
          size=13.5, bold=True)

    _txt(s,
         "Obesity drives diabetes, heart disease\nand several cancers — the conditions\nthat pull life expectancy down. No amount\nof treatment spending offsets a\npopulation-wide risk this large.",
         rx, Inches(3.2), rw, Inches(2.2), size=12, color=C_BODY)

    _txt(s,
         "This single factor does more to explain\nthe U.S. shortfall than the health budget does.",
         rx, Inches(5.5), rw, Inches(0.9), size=12, bold=True, color=C_DARK)

    _footer(s,
            "Source: WHO Global Health Observatory — "
            "prevalence of obesity among adults, BMI ≥ 30 (%), 2022.")


# ── Slide 5 — FACTOR 3: DEMOGRAPHICS ─────────────────────────────────────────

def _chart_demographics(df):
    peers = ["United States", "Japan", "France", "Spain", "Sweden",
             "Germany", "Canada", "Austria", "Belgium", "Netherlands",
             "Switzerland", "Denmark", "Norway", "Luxembourg", "Ireland"]
    sub = df[df["country"].isin(peers)].dropna(
        subset=["pop_65_plus_pct", "life_expectancy", "population_millions"]).copy()

    fig, ax = plt.subplots(figsize=(7.6, 5.0))

    for _, row in sub[~sub["country"].isin(["United States", "Japan"])].iterrows():
        bsz = (row["population_millions"] ** 0.5) * 22
        ax.scatter([row["pop_65_plus_pct"]], [row["life_expectancy"]],
                   color=CH_GRAY, alpha=0.40, s=bsz, zorder=2)
        ax.annotate(row["country"],
                    xy=(row["pop_65_plus_pct"], row["life_expectancy"]),
                    xytext=(4, 0), textcoords="offset points",
                    fontsize=7.5, color="#888888")

    for country, color in [("United States", CH_US), ("Japan", CH_JP)]:
        r = sub[sub["country"] == country]
        if r.empty:
            continue
        bsz = (float(r["population_millions"].iloc[0]) ** 0.5) * 22
        ax.scatter([float(r["pop_65_plus_pct"].iloc[0])],
                   [float(r["life_expectancy"].iloc[0])],
                   color=color, s=bsz, zorder=5, alpha=0.85)
        dy = -15 if country == "United States" else 5
        ax.annotate(country,
                    xy=(float(r["pop_65_plus_pct"].iloc[0]),
                        float(r["life_expectancy"].iloc[0])),
                    xytext=(5, dy), textcoords="offset points",
                    fontsize=9, color=color, fontweight="bold")

    ax.set_xlim(13, 32)
    ax.set_ylim(77.5, 85.0)
    _mpl_clean(ax,
               xlabel="Share of population aged 65+ (%, 2024)  —  bubble size = population",
               ylabel="Avg. life expectancy (years)")
    fig.tight_layout(pad=0.4)
    return fig


def slide_demographics(prs, df):
    s = _blank(prs)
    _bg(s, C_WHITE)
    _section(s, "FACTOR 3 — POPULATION & DEMOGRAPHICS")

    tf = _mixed_box(s, Inches(0.45), Inches(0.58), Inches(12.3), Inches(1.05))
    _para(tf, [("Ageing isn’t the excuse — ", C_DARK),
               ("Japan is the oldest, and lives the longest", C_TEAL)],
          size=27, bold=True)

    _pic(s, _to_img(_chart_demographics(df)),
         Inches(0.3), Inches(1.65), Inches(8.75), Inches(5.6))

    _dot_legend(s, Inches(9.2), "United States", CH_US)
    _dot_legend(s, Inches(11.15), "Japan", CH_JP)

    rx, rw = Inches(9.2), Inches(3.95)

    _txt(s,
         "Some argue demographics explain the gap — that the U.S. is simply younger.\nThe age structure says otherwise.",
         rx, Inches(2.35), rw, Inches(1.1), size=12, color=C_BODY)

    tf3 = _mixed_box(s, rx, Inches(3.55), rw, Inches(1.9))
    _para(tf3,
          [("The U.S. is the ", C_BODY),
           ("youngest", C_RED),
           (" of these peers\n(18% over 65) yet sits ", C_BODY),
           ("lowest", C_RED),
           (". Japan is the\n", C_BODY),
           ("oldest", C_TEAL),
           (" (30%) and ", C_BODY),
           ("highest", C_TEAL), (".", C_BODY)],
          size=12, bold=False)

    _txt(s,
         "A younger population should make\nlongevity easier, not harder.",
         rx, Inches(5.55), rw, Inches(0.8), size=12, bold=True, color=C_DARK)

    _footer(s,
            "Source: UN World Population Prospects / World Bank — "
            "population aged 65+ (% of total), 2024; population, 2023. "
            "Life expectancy: World Bank WDI.")


# ── Slide 6 — FACTOR 4: HEALTHCARE COSTS ─────────────────────────────────────

COSTS_2022 = {
    "United States": 12586, "Switzerland": 10575, "Norway": 9797,
    "Germany": 8470, "Ireland": 8288, "Luxembourg": 8203,
    "Austria": 7996, "Netherlands": 7788, "Belgium": 7359,
    "Denmark": 7333, "Sweden": 7229, "Canada": 7066,
    "France": 6643, "Japan": 5846, "Spain": 4916,
}


def _chart_costs():
    countries = list(COSTS_2022.keys())[::-1]   # ascending for barh
    values = [COSTS_2022[c] for c in countries]
    colors = [CH_US if c == "United States" else
              (CH_JP if c == "Japan" else CH_BAR) for c in countries]

    fig, ax = plt.subplots(figsize=(7.6, 5.2))
    bars = ax.barh(countries, values, color=colors, height=0.68, zorder=3)

    for bar, val, cntry in zip(bars, values, countries):
        fc = CH_US if cntry == "United States" else (CH_JP if cntry == "Japan" else "#6B7280")
        bw = "bold" if cntry in ("United States", "Japan") else "normal"
        ax.text(bar.get_width() + 80,
                bar.get_y() + bar.get_height() / 2,
                f"${val:,.0f}", va="center", fontsize=8.5,
                color=fc, fontweight=bw)

    ax.set_xlim(0, 16500)
    ax.xaxis.set_major_formatter(
        mticker.FuncFormatter(lambda x, _: f"${x/1000:.0f}K" if x > 0 else "$0"))
    ax.set_facecolor("white")
    ax.figure.patch.set_facecolor("white")
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_visible(False)
    ax.spines["bottom"].set_color("#E5E7EB")
    ax.tick_params(colors="#9CA3AF", labelsize=9, length=0)
    ax.set_xlabel("Health spending per person, PPP (current international $, 2022)",
                  fontsize=9, color="#9CA3AF")
    ax.grid(axis="x", color="#F3F4F6", linewidth=0.8, zorder=0)
    fig.tight_layout(pad=0.4)
    return fig


def slide_costs(prs):
    s = _blank(prs)
    _bg(s, C_WHITE)
    _section(s, "FACTOR 4 — HEALTHCARE COSTS")

    tf = _mixed_box(s, Inches(0.45), Inches(0.55), Inches(12.3), Inches(1.1))
    _para(tf, [("America’s bill is high because its ", C_DARK),
               ("prices", C_RED),
               (" are — not because it uses more care", C_DARK)],
          size=27, bold=True)

    _pic(s, _to_img(_chart_costs()),
         Inches(0.3), Inches(1.6), Inches(8.75), Inches(5.65))

    rx, rw = Inches(9.2), Inches(3.95)

    tf2 = _mixed_box(s, rx, Inches(2.0), rw, Inches(0.8))
    _para(tf2, [("The U.S. spends ~$12,600 per person", C_RED),
                (" — about ", C_BODY),
                ("2.2× Japan.", C_TEAL)],
          size=13.5, bold=True)

    tf3 = _mixed_box(s, rx, Inches(3.0), rw, Inches(2.4))
    _para(tf3,
          [("It isn’t that Americans see doctors more. "
            "It’s unit price: U.S. medical prices run about ", C_BODY),
           ("43% above the OECD average", C_RED),
           (" (price level 143 vs. 100), led by drugs and administration.", C_BODY)],
          size=12, bold=False)
    _para(tf3, [("", C_BODY)], size=7, bold=False)
    _para(tf3,
          [("Same care, higher price tag — dollars that "
            "never reach longer life.", C_DARK)],
          size=12, bold=True)

    _footer(s,
            "Source: World Bank WDI (spend per person, PPP, 2022). "
            "Price level: OECD, Society at a Glance 2024 / "
            "Health at a Glance (OECD avg = 100).")


# ── Slide 7 — THE PATTERN ─────────────────────────────────────────────────────

def slide_pattern(prs):
    s = _blank(prs)
    _bg(s, C_WHITE)
    _section(s, "THE PATTERN")

    tf = _mixed_box(s, Inches(0.45), Inches(0.52), Inches(12.3), Inches(1.0))
    _para(tf, [("One country, ", C_DARK),
               ("four red flags", C_RED),
               (" — and a budget that can’t fix them", C_DARK)],
          size=27, bold=True)

    # ── Column headers ──
    _txt(s, "UNITED STATES", Inches(5.35), Inches(1.72), Inches(3.2), Inches(0.38),
         size=10.5, bold=True, color=C_RED, align=PP_ALIGN.CENTER)
    _txt(s, "JAPAN", Inches(9.55), Inches(1.72), Inches(2.6), Inches(0.38),
         size=10.5, bold=True, color=C_TEAL, align=PP_ALIGN.CENTER)
    _rect(s, Inches(0.45), Inches(2.16), Inches(12.4), Inches(0.022), "#E5E7EB")

    rows = [
        ("Avg. life expectancy",   "78 yrs",  C_DARK,  "83 yrs",  C_TEAL),
        ("Health spend / person",  "$12,600", C_RED,   "$5,300",  C_DARK),
        ("Education — PISA 2022", "489",  C_DARK,  "533",     C_TEAL),
        ("Adult obesity",          "42.9%",   C_RED,   "4.9%",    C_DARK),
        ("Population aged 65+",    "17.9%",   C_DARK,  "29.8%",   C_DARK),
        ("Medical price level (OECD = 100)", "143", C_RED, "≈100", C_DARK),
    ]
    rh = Inches(0.66)
    for i, (metric, us_v, us_c, jp_v, jp_c) in enumerate(rows):
        top = Inches(2.18) + rh * i
        if i % 2 == 0:
            _rect(s, Inches(0.4), top, Inches(12.5), rh, "#F9FAFB")
        _txt(s, metric,
             Inches(0.55), top + Inches(0.16), Inches(4.4), Inches(0.38),
             size=12.5, color="#374151")
        _txt(s, us_v,
             Inches(5.35), top + Inches(0.1), Inches(3.2), Inches(0.48),
             size=21, bold=True, color=us_c, align=PP_ALIGN.CENTER)
        _txt(s, jp_v,
             Inches(9.55), top + Inches(0.1), Inches(2.6), Inches(0.48),
             size=21, bold=True, color=jp_c, align=PP_ALIGN.CENTER)

    _rect(s, Inches(0.45), Inches(6.60), Inches(12.4), Inches(0.022), "#E5E7EB")
    _txt(s,
         "The U.S. leads only where leading hurts — spending and prices — "
         "and trails on the upstream factors that actually make lives longer.",
         Inches(0.55), Inches(6.68), Inches(12.2), Inches(0.52),
         size=12, bold=True, color=C_DARK)

    _footer(s,
            "Sources: World Bank WDI; OECD PISA 2022; WHO GHO 2022; "
            "UN/World Bank 2024; OECD price levels. Figures rounded.")


# ── Slide 8 — THE TAKEAWAY ────────────────────────────────────────────────────

def slide_takeaway(prs):
    s = _blank(prs)
    _bg(s, C_NAVY)

    _txt(s, "THE TAKEAWAY",
         Inches(0.45), Inches(0.32), Inches(4.0), Inches(0.32),
         size=9, bold=True, color=C_SUB)

    tf = _mixed_box(s, Inches(0.45), Inches(0.82), Inches(12.3), Inches(2.15))
    _para(tf, [("Stop asking ", C_WHITE),
               ("how much", C_MUTED),
               (" to spend.", C_WHITE)],
          size=44, bold=True, space_after=0)
    _para(tf, [("Fix what the money ", C_WHITE),
               ("sits on.", C_TEAL)],
          size=44, bold=True)

    recs = [
        ("Measure outcomes per dollar, not dollars.",
         "Benchmark health systems on life expectancy gained per dollar — "
         "the metric on which the U.S. ranks last among peers."),
        ("Move marginal dollars upstream.",
         "Shift spending toward obesity prevention and education equity — "
         "the levers where Japan and Spain win on far smaller budgets."),
        ("Attack price, not just volume.",
         "Target the ~43% price premium — drug costs and administration — "
         "before adding another dollar of treatment spend."),
    ]

    for i, (title, body) in enumerate(recs):
        top = Inches(3.2) + Inches(1.3) * i
        circ = _oval(s, Inches(0.45), top + Inches(0.0), Inches(0.46), Inches(0.46), C_RED)
        # Number inside circle — use shape text frame
        circ.text_frame.text = str(i + 1)
        circ.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        run = circ.text_frame.paragraphs[0].runs[0]
        run.font.size = Pt(14)
        run.font.bold = True
        run.font.color.rgb = _rgb(C_WHITE)
        from pptx.enum.text import MSO_ANCHOR
        circ.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

        _txt(s, title,
             Inches(1.1), top, Inches(11.8), Inches(0.44),
             size=15, bold=True, color=C_WHITE)
        _txt(s, body,
             Inches(1.1), top + Inches(0.46), Inches(11.8), Inches(0.76),
             size=11.5, color=C_SUB)

    _footer(s,
            "Built on the attached Power BI analysis (World Bank WDI) "
            "· added factors: OECD, WHO, World Bank.")


# ── Entry point ───────────────────────────────────────────────────────────────

def main():
    base = os.path.dirname(os.path.abspath(__file__))
    df   = pd.read_csv(os.path.join(base, "data", "health_data.csv"))
    out  = os.path.join(base, "health_expenditure_analysis.pptx")

    prs = Presentation()
    prs.slide_width  = W
    prs.slide_height = H

    steps = [
        ("Cover",         lambda: slide_cover(prs)),
        ("The Setup",     lambda: slide_setup(prs, df)),
        ("Education",     lambda: slide_education(prs, df)),
        ("Lifestyle",     lambda: slide_lifestyle(prs, df)),
        ("Demographics",  lambda: slide_demographics(prs, df)),
        ("Costs",         lambda: slide_costs(prs)),
        ("Pattern",       lambda: slide_pattern(prs)),
        ("Takeaway",      lambda: slide_takeaway(prs)),
    ]

    for label, fn in steps:
        print(f"  {label} ...")
        fn()

    prs.save(out)
    print(f"Saved: {out}")


if __name__ == "__main__":
    main()
