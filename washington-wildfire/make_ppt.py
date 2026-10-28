#!/usr/bin/env python3
"""
Generate a concise, professional presentation deck (.pptx) for the
Washington State Wildfire Preparedness Dashboard, with REAL screenshots
of the running dashboard embedded.

Narrative focus: target user -> problem -> solution -> how it supports
Washington State wildfire preparedness planning.

All figures are read from data/dashboard_data.json (real DNR data).
Screenshots are read from assets/ (captured by shoot.py from the live app).

Usage:
    .venv/bin/python make_ppt.py
Output:
    Washington_Wildfire_Dashboard.pptx
"""

import json
import os
from collections import Counter, defaultdict

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from PIL import Image

# ---------------------------------------------------------------- palette
INK     = RGBColor(0x1E, 0x29, 0x3B)
MUTED   = RGBColor(0x64, 0x74, 0x8B)
ACCENT  = RGBColor(0xB9, 0x1C, 0x1C)
ACCENT2 = RGBColor(0xEA, 0x58, 0x0C)
DARK1   = RGBColor(0x7F, 0x1D, 0x1D)
DARK2   = RGBColor(0x9A, 0x34, 0x12)
WHITE   = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT   = RGBColor(0xF1, 0xF5, 0xF9)
LINE    = RGBColor(0xE2, 0xE8, 0xF0)
SAND    = RGBColor(0xFD, 0xBA, 0x74)
CREAM   = RGBColor(0xFD, 0xE8, 0xE8)
FONT    = "Calibri"

ASSETS = "assets"

# ---------------------------------------------------------------- data
with open("data/dashboard_data.json", encoding="utf-8") as f:
    D = json.load(f)
META, RECS = D["meta"], D["records"]
MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]

total_inc = len(RECS)
total_ac = sum(r["a"] for r in RECS)
cty = defaultdict(lambda: {"inc": 0, "ac": 0.0, "region": "", "causes": Counter()})
mon = Counter(); dated = 0
cause = Counter(); years = Counter()
for r in RECS:
    o = cty[r["c"]]; o["inc"] += 1; o["ac"] += r["a"]; o["region"] = r["r"]; o["causes"][r["u"]] += 1
    if r["m"]:
        mon[r["m"]] += 1; dated += 1
    cause[r["u"]] += 1
    if r["y"]:
        years[r["y"]] += 1

top_counties = sorted(cty.items(), key=lambda kv: -kv[1]["ac"])[:5]
top_causes = cause.most_common(3)
top_months = mon.most_common(3)
unknown = cause.get("Unknown", 0)
yr_min, yr_max = min(years), max(years)


def f(n):
    return f"{int(round(n)):,}"


def pct(a, b):
    return round(100 * a / b) if b else 0


# ---------------------------------------------------------------- deck setup
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]


def slide():
    return prs.slides.add_slide(BLANK)


def rect(s, x, y, w, h, fill, line=None):
    shp = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shp.fill.solid(); shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line; shp.line.width = Pt(0.75)
    shp.shadow.inherit = False
    return shp


def textbox(s, x, y, w, h, anchor=MSO_ANCHOR.TOP):
    tb = s.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = Inches(0.05); tf.margin_right = Inches(0.05)
    tf.margin_top = Inches(0.02); tf.margin_bottom = Inches(0.02)
    return tf


def para(tf, text, size=18, color=INK, bold=False, align=PP_ALIGN.LEFT,
         first=False, space_after=6, bullet=False):
    p = tf.paragraphs[0] if first and not tf.paragraphs[0].runs else tf.add_paragraph()
    p.alignment = align
    p.space_after = Pt(space_after)
    if bullet:
        text = "•  " + text
    r = p.add_run(); r.text = text
    fnt = r.font
    fnt.name = FONT; fnt.size = Pt(size); fnt.bold = bold; fnt.color.rgb = color
    return p


def title_bar(s, kicker, title):
    rect(s, 0, 0, SW, Inches(1.1), DARK1)
    rect(s, 0, Inches(1.1), SW, Inches(0.06), ACCENT2)
    tf = textbox(s, Inches(0.6), Inches(0.12), Inches(12.1), Inches(0.92), MSO_ANCHOR.MIDDLE)
    para(tf, kicker, size=12, color=SAND, bold=True, first=True, space_after=2)
    para(tf, title, size=27, color=WHITE, bold=True, space_after=0)


def footer(s, n):
    tf = textbox(s, Inches(0.6), Inches(7.04), Inches(11), Inches(0.36))
    para(tf, "Washington State Wildfire Preparedness Dashboard  ·  Source: WA DNR Fire Statistics 2008–Present  ·  Historical patterns, not a prediction",
         size=9, color=MUTED, first=True, space_after=0)
    tf2 = textbox(s, Inches(12.3), Inches(7.04), Inches(0.7), Inches(0.36))
    para(tf2, str(n), size=9, color=MUTED, first=True, align=PP_ALIGN.RIGHT, space_after=0)


def add_shot(s, path, x, y, max_w, max_h, border=True):
    """Place an image scaled to fit (max_w x max_h), centered in that box."""
    with Image.open(path) as im:
        iw, ih = im.size
    ar = iw / ih
    box_ar = max_w / max_h
    if ar > box_ar:
        w = max_w; h = int(max_w / ar)
    else:
        h = max_h; w = int(max_h * ar)
    px = x + (max_w - w) // 2
    py = y + (max_h - h) // 2
    if border:
        rect(s, Emu(int(px)) - Emu(9000), Emu(int(py)) - Emu(9000),
             Emu(int(w)) + Emu(18000), Emu(int(h)) + Emu(18000), WHITE, LINE)
    s.shapes.add_picture(path, Emu(int(px)), Emu(int(py)), Emu(int(w)), Emu(int(h)))


def caption(s, x, y, w, text):
    tf = textbox(s, x, y, w, Inches(0.9))
    para(tf, text, size=12.5, color=INK, first=True, space_after=0)


# ================================================================ 1. Title
s = slide()
rect(s, 0, 0, SW, SH, DARK1)
rect(s, 0, 0, SW, SH, DARK2)
rect(s, 0, Inches(4.55), SW, Inches(0.08), ACCENT2)
tf = textbox(s, Inches(0.9), Inches(1.55), Inches(11.5), Inches(2.9))
para(tf, "WASHINGTON STATE", size=20, color=SAND, bold=True, first=True, space_after=4)
para(tf, "Wildfire Preparedness Dashboard", size=43, color=WHITE, bold=True, space_after=10)
para(tf, "Turning decades of DNR fire data into actionable statewide preparedness planning.",
     size=18, color=CREAM, space_after=0)
tf2 = textbox(s, Inches(0.9), Inches(4.85), Inches(11.5), Inches(1.6))
para(tf2, f"{f(total_inc)} real fire records   ·   {yr_min}–{yr_max}   ·   {len(cty)} counties   ·   {f(total_ac)} acres",
     size=16, color=WHITE, bold=True, first=True, space_after=6)
para(tf2, "For Washington State emergency-management & wildfire-planning officials",
     size=13, color=SAND, space_after=0)

# ================================================================ 2. User & problem
s = slide(); title_bar(s, "WHO & WHY", "Our user and the problem")
# user card
rect(s, Inches(0.6), Inches(1.5), Inches(5.9), Inches(1.65), LIGHT, LINE)
rect(s, Inches(0.6), Inches(1.5), Inches(0.12), Inches(1.65), ACCENT)
tf = textbox(s, Inches(0.9), Inches(1.62), Inches(5.5), Inches(1.45), MSO_ANCHOR.MIDDLE)
para(tf, "TARGET USER", size=12, color=ACCENT, bold=True, first=True, space_after=4)
para(tf, "Washington State emergency-management and wildfire-planning officials", size=17, color=INK, bold=True, space_after=4)
para(tf, "deciding where and when to focus limited preparedness resources.", size=13, color=MUTED, space_after=0)

tf2 = textbox(s, Inches(0.6), Inches(3.4), Inches(5.9), Inches(3.4))
para(tf2, "The problem", size=18, color=ACCENT, bold=True, first=True, space_after=8)
for t in [
    "Longer, more intense wildfire seasons strain state resources.",
    "Decades of DNR records exist, but are hard to turn into decisions.",
    "Raw data mixes false alarms and blank fields, hiding real patterns.",
    "Officials lack a clear, trustworthy view of where activity concentrates.",
]:
    para(tf2, t, size=14.5, color=INK, bullet=True, space_after=9)

# right: impact stats
x0 = Inches(6.9); w = Inches(2.9); h = Inches(1.55); gap = Inches(0.25)
def stat(x, y, v, l):
    rect(s, x, y, w, h, WHITE, LINE)
    tf = textbox(s, x + Inches(0.1), y + Inches(0.1), w - Inches(0.2), h - Inches(0.2), MSO_ANCHOR.MIDDLE)
    para(tf, v, size=30, color=ACCENT, bold=True, first=True, align=PP_ALIGN.CENTER, space_after=2)
    para(tf, l, size=12, color=MUTED, align=PP_ALIGN.CENTER, space_after=0)
stat(x0, Inches(1.5), f(total_inc), "Real fires analyzed")
stat(x0 + w + gap, Inches(1.5), f(total_ac), "Acres burned")
stat(x0, Inches(1.5) + h + gap, f(META['false_alarms_excluded']), "False alarms removed")
stat(x0 + w + gap, Inches(1.5) + h + gap, f"{len(cty)}", "Counties covered")
tf3 = textbox(s, Inches(6.9), Inches(4.9), Inches(5.8), Inches(1.9))
para(tf3, "Every figure is computed directly from the real DNR records — false alarms excluded, blanks shown as 'Unknown', nothing invented.",
     size=13, color=MUTED, first=True, space_after=0)
footer(s, 2)

# ================================================================ 3. Solution overview
s = slide(); title_bar(s, "OUR SOLUTION", "One dashboard, four decision views")
feats = [
    ("Statewide fire map", "Every real fire located and clustered, with cause, acres and year."),
    ("Historical trends", "Incidents and acres burned by year and month."),
    ("Fire cause analysis", "Dominant causes across counties and regions."),
    ("Preparedness priorities", "Top counties to assess, with specific recommendations."),
    ("Live filters", "County, region, year and cause update every view together."),
    ("Transparent limits", "Sources, cleaning rules and caveats stated plainly."),
]
cols = 3; cw = Inches(3.95); ch = Inches(2.15); gx = Inches(0.25); gy = Inches(0.35)
x_start = Inches(0.6); y_start = Inches(1.5)
for i, (h1, body) in enumerate(feats):
    r, c = divmod(i, cols)
    x = x_start + c * (cw + gx); y = y_start + r * (ch + gy)
    rect(s, x, y, cw, ch, WHITE, LINE)
    rect(s, x, y, Inches(0.09), ch, ACCENT)
    tf = textbox(s, x + Inches(0.22), y + Inches(0.18), cw - Inches(0.35), ch - Inches(0.35))
    para(tf, h1, size=16, color=ACCENT, bold=True, first=True, space_after=6)
    para(tf, body, size=13, color=INK, space_after=0)
footer(s, 3)

# ================================================================ 4. Screenshot: map
s = slide(); title_bar(s, "LIVE DASHBOARD", "Statewide fire map")
add_shot(s, f"{ASSETS}/shot_map.png", Inches(0.6), Inches(1.45), Inches(8.6), Inches(5.1))
tf = textbox(s, Inches(9.45), Inches(1.6), Inches(3.4), Inches(5))
para(tf, "What it shows", size=15, color=ACCENT, bold=True, first=True, space_after=8)
for t in [
    f"{f(total_inc)} real (non-false-alarm) fires, clustered for readability.",
    "Clusters reveal where activity concentrates — the dry eastern counties.",
    "Click any cluster to drill into county, cause, acres and year.",
]:
    para(tf, t, size=13.5, color=INK, bullet=True, space_after=10)
para(tf, "Planning value: an at-a-glance map of the state's historical fire footprint.",
     size=13, color=MUTED, space_after=0)
footer(s, 4)

# ================================================================ 5. Screenshot: trends
s = slide(); title_bar(s, "LIVE DASHBOARD", "Historical trends")
add_shot(s, f"{ASSETS}/shot_trends.png", Inches(0.6), Inches(1.45), Inches(8.6), Inches(5.1))
tf = textbox(s, Inches(9.45), Inches(1.6), Inches(3.4), Inches(5))
para(tf, "What it shows", size=15, color=ACCENT, bold=True, first=True, space_after=8)
for t in [
    f"Incidents & acres by year ({yr_min}–{yr_max}) and by month.",
    f"Clear summer peak: {MONTHS[top_months[0][0]-1]}, {MONTHS[top_months[1][0]-1]} and {MONTHS[top_months[2][0]-1]} drive ~{pct(sum(c for _,c in top_months), dated)}% of incidents.",
    "Toggle between incident counts and acres burned.",
]:
    para(tf, t, size=13.5, color=INK, bullet=True, space_after=10)
para(tf, "Planning value: shows when to stage crews, equipment and burn bans.",
     size=13, color=MUTED, space_after=0)
footer(s, 5)

# ================================================================ 6. Screenshot: priorities
s = slide(); title_bar(s, "FROM DATA TO DECISIONS", "State Preparedness Priorities")
add_shot(s, f"{ASSETS}/shot_priorities.png", Inches(0.6), Inches(1.5), Inches(12.1), Inches(3.9))
tf = textbox(s, Inches(0.6), Inches(5.65), Inches(12.1), Inches(1.2))
para(tf, f"Top counties by historical burned acreage — {top_counties[0][0].title()} (~{pct(top_counties[0][1]['ac'], total_ac)}% of all acreage), "
         f"{top_counties[1][0].title()}, {top_counties[2][0].title()}. Each card pairs the evidence (incidents, acres, causes, peak months) "
         f"with a specific, data-backed recommendation — explicitly labeled historical, not a prediction.",
     size=13, color=INK, first=True, space_after=0)
footer(s, 6)

# ================================================================ 7. How it supports planning
s = slide(); title_bar(s, "SUPPORTING PREPAREDNESS", "How it helps Washington State plan")
recs = [
    ("Target high-burden counties",
     f"The top 5 counties hold ~{pct(sum(o['ac'] for _,o in top_counties), total_ac)}% of burned acreage — focus mitigation, mutual-aid and pre-positioning there."),
    ("Stage resources before peak season",
     f"~{pct(sum(c for _,c in top_months), dated)}% of incidents occur {MONTHS[top_months[0][0]-1]}–{MONTHS[top_months[2][0]-1]}; align staffing, readiness and burn bans ahead of summer."),
    ("Match prevention to local causes",
     f"Leading causes differ by county ({top_causes[0][0]}, {top_causes[1][0]}, {top_causes[2][0]}); tailor permits and outreach to each area's drivers."),
    ("Keep decisions evidence-based",
     f"Only ~{pct(unknown, total_inc)}% of causes are Unknown; every recommendation cites its supporting data and stays historical, not predictive."),
]
y = Inches(1.5)
for h1, body in recs:
    rect(s, Inches(0.6), y, Inches(12.1), Inches(1.15), WHITE, LINE)
    rect(s, Inches(0.6), y, Inches(0.12), Inches(1.15), ACCENT2)
    tf = textbox(s, Inches(0.9), y + Inches(0.1), Inches(11.6), Inches(0.95), MSO_ANCHOR.MIDDLE)
    para(tf, h1, size=16, color=ACCENT, bold=True, first=True, space_after=3)
    para(tf, body, size=13, color=INK, space_after=0)
    y += Inches(1.3)
footer(s, 7)

# ================================================================ 8. Closing
s = slide()
rect(s, 0, 0, SW, SH, DARK1)
rect(s, 0, Inches(3.45), SW, Inches(0.08), ACCENT2)
tf = textbox(s, Inches(0.9), Inches(2.1), Inches(11.5), Inches(2))
para(tf, "Data-driven wildfire preparedness", size=36, color=WHITE, bold=True, first=True, space_after=10)
para(tf, "A clean, transparent, filter-driven view of Washington's wildfire history — built to help the state prepare, not to predict.",
     size=17, color=CREAM, space_after=0)
tf2 = textbox(s, Inches(0.9), Inches(3.95), Inches(11.5), Inches(1.3))
para(tf2, f"{f(total_inc)} fires  ·  {f(total_ac)} acres  ·  {len(cty)} counties  ·  {yr_min}–{yr_max}",
     size=16, color=SAND, bold=True, first=True, space_after=10)
para(tf2, "Limitations: historical patterns only; acreage skewed by a few large fires; some records lack a full year or cause.",
     size=12, color=CREAM, space_after=0)

out = "Washington_Wildfire_Dashboard.pptx"
prs.save(out)
print("Saved", out, "with", len(prs.slides._sldIdLst), "slides")
