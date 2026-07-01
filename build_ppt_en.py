# -*- coding: utf-8 -*-
"""
Generate the ENGLISH CFO deck: Finance & Controlling Transformation (10-15 min).
Run:    python3 build_ppt_en.py
Output: CFO_Finance_Controlling_Transformation_EN.pptx  (with speaker notes)
Deps:   python-pptx  (pip install python-pptx)

Design: clean corporate / Office template -- light body, slim navy header,
        hairline dividers, consistent footer with slide numbers, Calibri type.
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# ---------- Corporate palette ----------
NAVY  = RGBColor(0x1F, 0x3A, 0x5F)   # deep navy (primary)
STEEL = RGBColor(0x2E, 0x6D, 0xA4)   # steel blue (accent)
SLATE = RGBColor(0x5A, 0x6B, 0x7B)   # slate
TEAL  = RGBColor(0x17, 0x8F, 0x80)   # positive
RED   = RGBColor(0xB0, 0x3A, 0x2E)   # muted red (pain/risk)
GOLD  = RGBColor(0xC8, 0xA2, 0x4B)   # gold accent
INK   = RGBColor(0x33, 0x38, 0x3F)   # body text
MIST  = RGBColor(0xF2, 0xF5, 0xF8)   # light card/bg
LINE  = RGBColor(0xD9, 0xE0, 0xE7)   # hairline
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BGTINT= RGBColor(0xFB, 0xFC, 0xFD)   # near-white slide bg

TITLE_FONT = "Calibri Light"
BODY_FONT  = "Calibri"

prs = Presentation()
prs.slide_width  = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]
SW, SH = prs.slide_width, prs.slide_height

FOOTER = "Finance & Controlling Transformation"
CONF   = "Confidential  ·  Prepared for the CFO"


# ---------- helpers ----------
def add_slide():
    return prs.slides.add_slide(BLANK)

def rect(slide, x, y, w, h, fill, line=None, line_w=1.0):
    sp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid(); sp.fill.fore_color.rgb = fill
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line; sp.line.width = Pt(line_w)
    sp.shadow.inherit = False
    return sp

def txt(slide, x, y, w, h, runs, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
        space_after=6, line_spacing=1.05, font=BODY_FONT):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame; tf.word_wrap = True; tf.vertical_anchor = anchor
    for i, para in enumerate(runs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align; p.space_after = Pt(space_after); p.line_spacing = line_spacing
        for (t, sz, col, bold) in para:
            r = p.add_run(); r.text = t
            r.font.size = Pt(sz); r.font.color.rgb = col; r.font.bold = bold
            r.font.name = font
    return tb

def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text

def bg(slide, color=BGTINT):
    rect(slide, 0, 0, SW, SH, color)

def footer(slide, idx):
    rect(slide, Inches(0.6), Inches(7.02), Inches(12.13), Pt(0.75), LINE)
    txt(slide, Inches(0.6), Inches(7.08), Inches(5.0), Inches(0.32),
        [[(FOOTER, 9, SLATE, False)]])
    txt(slide, Inches(4.0), Inches(7.08), Inches(5.33), Inches(0.32),
        [[(CONF, 9, SLATE, False)]], align=PP_ALIGN.CENTER)
    txt(slide, Inches(11.4), Inches(7.08), Inches(1.33), Inches(0.32),
        [[(f"{idx:02d}", 9, SLATE, True)]], align=PP_ALIGN.RIGHT)

def header(slide, kicker, title, idx):
    rect(slide, 0, 0, SW, Inches(1.12), NAVY)
    rect(slide, 0, Inches(1.12), SW, Pt(2.5), GOLD)
    txt(slide, Inches(0.6), Inches(0.16), Inches(11.8), Inches(0.3),
        [[(kicker.upper(), 11, GOLD, True)]])
    txt(slide, Inches(0.6), Inches(0.44), Inches(12.1), Inches(0.62),
        [[(title, 25, WHITE, False)]], anchor=MSO_ANCHOR.MIDDLE, font=TITLE_FONT)
    footer(slide, idx)

def bullets(slide, x, y, w, h, items, size=15, gap=8):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame; tf.word_wrap = True
    for i, (t, lvl, col) in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(gap); p.line_spacing = 1.08
        mark = "▪   " if lvl == 0 else "–   "
        r = p.add_run(); r.text = ("" if lvl == 0 else "        ") + mark + t
        r.font.size = Pt(size if lvl == 0 else size-1.5)
        r.font.color.rgb = col; r.font.bold = (lvl == 0); r.font.name = BODY_FONT
    return tb

def styled_table(slide, rows, x, y, w, h, col_w, head_fill=NAVY, zebra=True,
                 hi_col=None, hi_color=TEAL, first_col_accent=STEEL,
                 head_sz=13, body_sz=12):
    t = slide.shapes.add_table(len(rows), len(rows[0]), x, y, w, h).table
    for c, cw in enumerate(col_w):
        t.columns[c].width = cw
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            cell = t.cell(r, c); cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.margin_left = Inches(0.12); cell.margin_right = Inches(0.1)
            cell.margin_top = Inches(0.04); cell.margin_bottom = Inches(0.04)
            tf = cell.text_frame; tf.word_wrap = True
            p = tf.paragraphs[0]; run = p.add_run(); run.text = val
            run.font.name = BODY_FONT
            if r == 0:
                cell.fill.solid(); cell.fill.fore_color.rgb = head_fill
                run.font.color.rgb = WHITE; run.font.bold = True; run.font.size = Pt(head_sz)
            else:
                cell.fill.solid()
                cell.fill.fore_color.rgb = MIST if (zebra and r % 2 == 0) else WHITE
                run.font.size = Pt(body_sz); run.font.color.rgb = INK
                if c == 0 and first_col_accent:
                    run.font.bold = True; run.font.color.rgb = first_col_accent
                if hi_col is not None and c == hi_col:
                    run.font.bold = True; run.font.color.rgb = hi_color
    return t


# =====================================================================
# 1) Title
# =====================================================================
s = add_slide()
rect(s, 0, 0, SW, SH, NAVY)
rect(s, 0, 0, Inches(0.22), SH, GOLD)
txt(s, Inches(0.95), Inches(1.45), Inches(11.5), Inches(0.4),
    [[("CONTROLLER  ·  BUSINESS INTELLIGENCE & REPORTING", 14, GOLD, True)]])
txt(s, Inches(0.95), Inches(2.15), Inches(11.6), Inches(2.1),
    [[("From Reporting Factory to Business Partner", 40, WHITE, False)],
     [("A Finance & Controlling Transformation Proposal", 24, RGBColor(0xC9,0xD6,0xE4), False)]],
    line_spacing=1.15, font=TITLE_FONT)
rect(s, Inches(0.98), Inches(4.35), Inches(3.4), Pt(2), GOLD)
txt(s, Inches(0.95), Inches(4.65), Inches(11.6), Inches(1.1),
    [[("Guiding idea:  ", 16, GOLD, True),
      ("free Finance from producing reports, so it can help you make decisions —", 16, WHITE, False)],
     [("faster, more reliable, more forward-looking, at low cost and aligned with Group.", 16, WHITE, False)]],
    line_spacing=1.25)
txt(s, Inches(0.95), Inches(6.55), Inches(11.6), Inches(0.4),
    [[("Audience: CFO     |     Duration: 10–15 minutes", 12, RGBColor(0x9F,0xB4,0xC9), False)]])
notes(s, "Opening: Thank you for the opportunity. In the next 12 minutes I will cover three things — how I read the situation, what I would do first, and a low-cost roadmap that stays aligned with Group. My goal in one sentence: free Finance from producing reports so it can help you make decisions.")

# =====================================================================
# 2) Diagnosis
# =====================================================================
s = add_slide(); bg(s)
header(s, "1 · Diagnosis", "The issue is not effort — it is the lack of a platform and standards", 2)
cards = [
    ("First report only on Day 4; close is too slow",
     "Root cause: SAP export → Excel rework → manual consolidation; no automated pipeline"),
    ("7 entities define figures differently",
     "Root cause: no common KPI / account definitions or master-data governance; no single source of truth"),
    ("Every follow-up question means redoing the work",
     "Root cause: reports are static Excel with no self-service drill-down"),
    ("Forecast / plan do not reconcile with actuals",
     "Root cause: plan, forecast and actuals live in separate spreadsheets; no unified data model"),
]
x0=Inches(0.6); y0=Inches(1.55); cw=Inches(6.0); ch=Inches(1.35); gx=Inches(0.25); gy=Inches(0.22)
for i,(head,root) in enumerate(cards):
    cx=x0+(cw+gx)*(i%2); cy=y0+(ch+gy)*(i//2)
    rect(s, cx, cy, cw, ch, WHITE, line=LINE)
    rect(s, cx, cy, Pt(5), ch, RED)
    txt(s, cx+Inches(0.22), cy+Inches(0.14), cw-Inches(0.4), Inches(0.55),
        [[("▲  "+head, 14.5, RED, True)]])
    txt(s, cx+Inches(0.22), cy+Inches(0.68), cw-Inches(0.4), Inches(0.6),
        [[(root, 12, SLATE, False)]], line_spacing=1.05)
txt(s, Inches(0.6), Inches(4.72), Inches(12), Inches(0.5),
    [[("In one line:  ", 15, NAVY, True),
      ("today we are slow, inconsistent, reactive — and backward-looking.", 15, INK, False)]])
notes(s, "The bottom line first: I translated management's four complaints into four root causes. Note the red text on the right is all about systems and standards — not people not trying. In particular, 'inconsistent figures' is fundamentally a people-and-process issue: each entity defines the same metric differently. Getting the root cause right is what makes the roadmap effective.")

# =====================================================================
# 3) Principles
# =====================================================================
s = add_slide(); bg(s)
header(s, "2 · Approach", "Five guiding principles", 3)
principles = [
    ("1 · Single source of truth", "SAP is the one system of record; every report comes from the same data model"),
    ("2 · Define first, then automate", "Align KPI definitions before tooling — otherwise you automate errors"),
    ("3 · People analyze, machines move data", "Hand manual consolidation to the system; people focus on insight"),
    ("4 · Self-service", "Management drills down on its own; follow-ups no longer create manual work"),
    ("5 · Pragmatic, small steps", "Start with existing licenses for quick wins; trade results for budget"),
]
y=Inches(1.55)
for head,desc in principles:
    rect(s, Inches(0.6), y, Inches(3.9), Inches(0.84), NAVY)
    txt(s, Inches(0.78), y+Inches(0.14), Inches(3.6), Inches(0.6),
        [[(head, 15, WHITE, True)]], anchor=MSO_ANCHOR.MIDDLE)
    rect(s, Inches(4.65), y, Inches(8.05), Inches(0.84), WHITE, line=LINE)
    txt(s, Inches(4.85), y+Inches(0.14), Inches(7.7), Inches(0.6),
        [[(desc, 13.5, INK, False)]], anchor=MSO_ANCHOR.MIDDLE)
    y+=Inches(1.0)
notes(s, "These five are my working method. The one I stress most is #2: define first, then automate — if definitions are not aligned, automation just produces inconsistent numbers faster. And #5 is your reassurance: I will not ask for a big budget and a big project up front; I will use tools the company already owns to create results, then use those results to earn the next investment.")

# =====================================================================
# 4) Vision
# =====================================================================
s = add_slide(); bg(s)
header(s, "3 · Vision", "From a reporting factory to a business decision partner", 4)
rect(s, Inches(0.8), Inches(2.2), Inches(5.0), Inches(3.1), WHITE, line=LINE)
rect(s, Inches(0.8), Inches(2.2), Inches(5.0), Inches(0.72), RED)
txt(s, Inches(0.8), Inches(2.31), Inches(5.0), Inches(0.55),
    [[("TODAY · Reporting Factory", 16, WHITE, True)]], align=PP_ALIGN.CENTER)
bullets(s, Inches(1.15), Inches(3.15), Inches(4.4), Inches(2.0),
    [("First report on Day 4", 0, INK), ("Heavy manual consolidation", 0, INK),
     ("Inconsistent · reactive", 0, INK), ("Backward-looking", 0, INK)], size=15, gap=12)
txt(s, Inches(5.95), Inches(3.3), Inches(1.4), Inches(1.0),
    [[("➜", 48, STEEL, True)]], align=PP_ALIGN.CENTER)
rect(s, Inches(7.55), Inches(2.2), Inches(5.0), Inches(3.1), WHITE, line=LINE)
rect(s, Inches(7.55), Inches(2.2), Inches(5.0), Inches(0.72), TEAL)
txt(s, Inches(7.55), Inches(2.31), Inches(5.0), Inches(0.55),
    [[("TARGET · Business Partner", 16, WHITE, True)]], align=PP_ALIGN.CENTER)
bullets(s, Inches(7.9), Inches(3.15), Inches(4.4), Inches(2.0),
    [("Near-real-time · one truth", 0, NAVY), ("Automated pipeline", 0, NAVY),
     ("Self-service · proactive", 0, NAVY), ("Forward-looking · rolling forecast", 0, NAVY)], size=15, gap=12)
notes(s, "One picture for the direction: left is where we are, right is the target. This transformation is not about buying a tool; it is about shifting Finance from moving numbers to explaining the business and looking ahead. It also answers the three words management asked for: faster, more reliable, more forward-looking.")

# =====================================================================
# 5) Priorities
# =====================================================================
s = add_slide(); bg(s)
header(s, "4 · Priorities", "What comes first, and why", 5)
items = [
    ("First: align KPI / account definitions (a data dictionary)", 0, NAVY),
    ("Zero cost, highest value; it is the foundation for automation and real-time. Without it, automation only produces errors faster.", 1, SLATE),
    ("Next: automate the current consolidation (Power Query)", 0, NAVY),
    ("Immediate impact — first report from Day 4 to Day 2, with no IT and no spend.", 1, SLATE),
    ("Then: a central data model = single source of truth", 0, NAVY),
    ("Map each entity's local accounts to Group accounts; end 'whose spreadsheet wins'.", 1, SLATE),
    ("Finally: integrate planning & forecasting + forward analytics", 0, NAVY),
    ("Actuals / budget / forecast in one view; rolling forecast and scenario modelling.", 1, SLATE),
]
bullets(s, Inches(0.7), Inches(1.6), Inches(12), Inches(5), items, size=16, gap=10)
notes(s, "The sequencing logic: I order by value, speed to impact, and low dependency. The first move is aligning definitions because it costs nothing yet underpins everything; the second is automating the existing Excel so you see reports get faster within 90 days. Fix 'fast and reliable' first, then move to 'forward-looking'.")

# =====================================================================
# 6) Roadmap
# =====================================================================
s = add_slide(); bg(s)
header(s, "5 · Roadmap", "Stabilize → Standardize → Automate → Elevate (four phases)", 6)
rows = [
    ("Phase", "Key actions", "Deliverable / milestone"),
    ("0 · Assess (~30 d)", "Map the process; inventory reports & licenses; lock core KPIs", "Assessment + quick-win list"),
    ("1 · Stabilize (1–3 m)", "KPI data dictionary + entity sign-off; Power Query auto-consolidation", "First report: Day 4 → Day 2"),
    ("2 · Standardize (3–6 m)", "Central data model; local→Group account mapping; master-data governance", "Single source of truth + drill-down"),
    ("3 · Automate (6–12 m)", "End-to-end automation; daily/near-real-time; integrate plan & forecast", "First report → Day 1 / near-real-time"),
    ("4 · Elevate (12 m +)", "Rolling forecast, driver-based models, scenarios, anomaly alerts", "Forward-looking + predictive analytics"),
]
styled_table(s, rows, Inches(0.6), Inches(1.5), Inches(12.13), Inches(4.9),
             [Inches(2.55), Inches(6.0), Inches(3.58)], head_sz=13.5, body_sz=12)
notes(s, "This is the backbone of the proposal. Four phases, each with a measurable milestone — watch the single thread 'time to first report': Day 4 to Day 2 to Day 1 to near-real-time. I will not rip and replace; each phase delivers a visible result.")

# =====================================================================
# 7) Platform
# =====================================================================
s = add_slide(); bg(s)
header(s, "6 · Platform", "Intelligent Shared Platform — layered architecture", 7)
layers = [
    ("Consumption", "BI dashboards · self-service · planning & forecasting · advanced analytics / AI", STEEL),
    ("Semantic / KPI", "Unified KPI definitions (the data dictionary built into the model)", TEAL),
    ("Central data", "Data warehouse = single source of truth (unified master data · Group account mapping)", NAVY),
    ("Integration", "Automated ELT (Power Query → dataflows → SAP Datasphere / BW)", SLATE),
    ("Source", "SAP (ERP · the one system of record) + other sources", RGBColor(0x44,0x4A,0x52)),
]
y=Inches(1.5)
for name,desc,col in layers:
    rect(s, Inches(2.3), y, Inches(3.0), Inches(0.84), col)
    txt(s, Inches(2.4), y+Inches(0.14), Inches(2.8), Inches(0.6),
        [[(name, 15, WHITE, True)]], anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
    rect(s, Inches(5.5), y, Inches(6.9), Inches(0.84), WHITE, line=LINE)
    txt(s, Inches(5.7), y+Inches(0.14), Inches(6.6), Inches(0.6),
        [[(desc, 12.5, INK, False)]], anchor=MSO_ANCHOR.MIDDLE)
    y+=Inches(0.96)
rect(s, Inches(0.6), Inches(1.5), Inches(1.5), Inches(4.6), GOLD)
txt(s, Inches(0.62), Inches(1.5), Inches(1.46), Inches(4.6),
    [[("Governance", 14, WHITE, True)],
     [("across all layers", 10, WHITE, False)],
     [("Data ownership", 10.5, WHITE, False)],
     [("KPI standards", 10.5, WHITE, False)],
     [("Security & access", 10.5, WHITE, False)],
     [("Change mgmt", 10.5, WHITE, False)]],
    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, space_after=9)
notes(s, "The platform is built layer by layer: SAP at the base is the one record, then automated integration, the central warehouse (single truth), the KPI semantic layer, and only at the top the dashboards and forecasts. The gold bar on the left is governance running through every layer — ownership, standards, access, change. The roadmap is simply building this bottom-up.")

# =====================================================================
# 8) Case A
# =====================================================================
s = add_slide(); bg(s)
header(s, "7 · Case A · One truth", "One truth in practice: the same 'Net Sales', 7 entities aligned", 8)
txt(s, Inches(0.6), Inches(1.35), Inches(12.1), Inches(0.5),
    [[("How: a KPI data dictionary locks 4 dimensions (tax / add-deduct items / account mapping / recognition timing) + entity sign-off + built into the model", 12, NAVY, True)]])
rows = [
    ("Entity", "As reported (k)", "Aligned basis (k)", "Reason for change"),
    ("Medical – East", "10,000", "9,400", "Deduct cash discounts & returns"),
    ("Medical – North", "9,500", "9,100", "Exclude freight income"),
    ("Production", "10,500", "9,000", "Exclude VAT (tax-inclusive split)"),
    ("Service", "9,000", "9,300", "Add omitted service revenue"),
    ("Trading", "9,800", "9,400", "Shipment basis → billing basis"),
]
styled_table(s, rows, Inches(0.6), Inches(1.95), Inches(12.13), Inches(3.5),
             [Inches(2.7), Inches(2.6), Inches(2.9), Inches(3.93)],
             hi_col=2, hi_color=TEAL, head_sz=13, body_sz=12.5)
txt(s, Inches(0.6), Inches(5.7), Inches(12.1), Inches(1.0),
    [[("Value:  ", 14, NAVY, True),
      ("differences shift from 'method differences' to 'business differences' — your follow-up moves from "
       "'is this number right?' to 'why is this business performing this way?', with no manual reconciliation.", 13.5, INK, False)]],
    line_spacing=1.15)
notes(s, "This is the most persuasive slide. The same 'Net Sales' produced seven different numbers, off by millions. After the data dictionary locks the four dimensions, entities sign off, and it is built into the model, we recompute — the green column is the aligned result. The key value: differences between entities are now real business differences, so your questions no longer trigger a round of manual reconciliation.")

# =====================================================================
# 9) Case B
# =====================================================================
s = add_slide(); bg(s)
header(s, "8 · Case B · Pipeline", "Mapping & split pipeline: dirty data in → clean table out", 9)
steps = [
    ("1 · Staging", "SAP raw export\n(tax-incl / freight mixed in / returns positive)", RED),
    ("2 · Transform", "join config tables → split VAT ÷1.13\nsplit service · drop freight · apply +/- sign", STEEL),
    ("3 · Clean fact", "long / tidy format\nGroup accounts + EUR", TEAL),
]
x=Inches(0.6)
for i,(h,d,col) in enumerate(steps):
    rect(s, x, Inches(1.55), Inches(3.7), Inches(1.75), col)
    txt(s, x+Inches(0.2), Inches(1.7), Inches(3.3), Inches(0.5),
        [[(h, 16, WHITE, True)]])
    txt(s, x+Inches(0.2), Inches(2.24), Inches(3.35), Inches(0.95),
        [[(ln, 12, WHITE, False)] for ln in d.split("\n")], line_spacing=1.05)
    if i<2:
        txt(s, x+Inches(3.75), Inches(1.95), Inches(0.5), Inches(0.8),
            [[("➜", 28, NAVY, True)]], align=PP_ALIGN.CENTER)
    x+=Inches(4.15)
txt(s, Inches(0.6), Inches(3.6), Inches(12), Inches(0.4),
    [[("Production entity (tax-inclusive 13% + freight mixed in) — result:", 14, NAVY, True)]])
res = [
    ("Naive sum (wrong)", "11,300,000 + 113,000 = 11,413,000", RED),
    ("Aligned Net Sales (right)", "10,000,000 − 1,000,000 − 200,000 − 50,000 = 8,750,000 CNY", TEAL),
    ("Built-in checks", "VAT reconstruction ✓   total conservation ✓   unmapped-account alert ✓", NAVY),
]
y=Inches(4.05)
for h,d,col in res:
    rect(s, Inches(0.6), y, Inches(3.4), Inches(0.62), col)
    txt(s, Inches(0.72), y+Inches(0.1), Inches(3.2), Inches(0.45),
        [[(h, 12, WHITE, True)]], anchor=MSO_ANCHOR.MIDDLE)
    rect(s, Inches(4.1), y, Inches(8.6), Inches(0.62), WHITE, line=LINE)
    txt(s, Inches(4.25), y+Inches(0.1), Inches(8.3), Inches(0.45),
        [[(d, 12, INK, False)]], anchor=MSO_ANCHOR.MIDDLE)
    y+=Inches(0.72)
txt(s, Inches(0.6), Inches(6.4), Inches(12.1), Inches(0.5),
    [[("Rules live in config tables, logic in the pipeline: change a basis by editing one row, reuse across all entities, every figure auditable.", 12.5, NAVY, True)]])
notes(s, "The previous slide is the 'why'; this is the 'how'. A three-layer pipeline: dirty data in, automatically split VAT, split service, drop freight, apply +/- signs, and out comes a clean tidy table. For the production entity: a naive sum overstates by 2.6m; the aligned Net Sales is 8.75m. The key is 'config-driven' — to change a basis, business edits one row of a table, no IT needed, and every figure traces back to the SAP document. I have this as runnable SQL and Excel.")

# =====================================================================
# 10) Real-time
# =====================================================================
s = add_slide(); bg(s)
header(s, "9 · Real-time", "How to achieve 'real-time' reporting", 10)
txt(s, Inches(0.6), Inches(1.45), Inches(12.1), Inches(0.75),
    [[("A mature view:  ", 15, GOLD, True),
      ("management decisions need 'right-time' — mostly daily / near-real-time, not per-second streaming. Spend where it matters.", 15, INK, False)]],
    line_spacing=1.15)
items = [
    ("Automated scheduled extraction + incremental refresh, replacing manual export (start with daily refresh)", 0, NAVY),
    ("Pre-built data model; reports read the model directly — open in seconds, drill down freely", 0, NAVY),
    ("Critical metrics (cash / orders) handled separately via DirectQuery / near-real-time connections", 0, NAVY),
    ("SAP data access (light to heavy):", 0, NAVY),
    ("OData / CDS Views  →  connector table extraction  →  reuse Group SAP BW / SAP Analytics Cloud", 1, SLATE),
]
bullets(s, Inches(0.7), Inches(2.5), Inches(12), Inches(4), items, size=16, gap=13)
notes(s, "On 'how to be real-time', I first give you a mature judgment: don't chase per-second streaming — it is expensive and unnecessary. Management reporting needs the right time; for most, a daily refresh is enough. The method: automated extraction replaces manual export, a pre-built model makes reports open instantly and drillable; only cash and orders get near-real-time. Choose the SAP access method based on Group's existing assets.")

# =====================================================================
# 11) Constraints
# =====================================================================
s = add_slide(); bg(s)
header(s, "10 · Delivery under constraints", "Limited budget + limited IT + Group alignment — how", 11)
items = [
    ("Use tools we have already paid for", 0, NAVY),
    ("Microsoft 365 — Power Query / Power BI / Power Automate — near-zero incremental cost for Phases 1–2", 1, SLATE),
    ("Do the zero-cost, high-value things first", 0, NAVY),
    ("The KPI data dictionary and process standardization cost nothing yet fix inconsistency", 1, SLATE),
    ("Low-code, minimal customization", 0, NAVY),
    ("Reduce IT dependency: business-led with IT guardrails", 1, SLATE),
    ("Reuse Group assets", 0, NAVY),
    ("Connect to the Group chart of accounts / SAP BW / SAC — cheaper and compliant by design", 1, SLATE),
    ("Trade results for budget", 0, NAVY),
    ("Deliver 90-day quick wins, then use 'Day 4 → Day 2, X hours saved' to request the next investment", 1, SLATE),
]
bullets(s, Inches(0.7), Inches(1.55), Inches(12), Inches(5.4), items, size=15, gap=8)
notes(s, "This slide directly addresses the 'this will cost a lot' concern. The core idea: start with Microsoft tools the company already owns, do the free-but-high-value standardization first, keep it low-code with minimal IT dependency, and reuse whatever Group already has. Most importantly — use 90-day results to earn the next budget, so investment always follows value.")

# =====================================================================
# 12) Growth
# =====================================================================
s = add_slide(); bg(s)
header(s, "11 · Staying ahead", "The role is evolving — how I keep up", 12)
items = [
    ("Build a 'Finance + Data' T-shaped skill set", 0, NAVY),
    ("Deepen data modelling, BI and analytics (Power BI / SQL / SAC certifications)", 1, SLATE),
    ("Track the trends", 0, NAVY),
    ("FP&A automation, analytics and AI in finance; join the Group Controlling network and external communities", 1, SLATE),
    ("Build lasting partnerships", 0, NAVY),
    ("Work alongside IT and the business, not in isolation", 1, SLATE),
    ("Keep a 'product mindset'", 0, NAVY),
    ("Treat reports as a product to iterate, not a one-off deliverable", 1, SLATE),
]
bullets(s, Inches(0.7), Inches(1.7), Inches(12), Inches(5), items, size=16, gap=11)
notes(s, "This role is evolving fast. My response: keep building myself into a T-shaped professional who understands both finance and data, stay close to FP&A automation and AI trends, build lasting partnerships with IT and the business, and use a product mindset to keep iterating the reports — they are always v1, v2, v3, never finished after one delivery.")

# =====================================================================
# 13) First 90 days + ask
# =====================================================================
s = add_slide(); bg(s)
header(s, "12 · First 90 days", "90-day commitment + the ask", 13)
rows = [
    ("30 days", "Situation assessment + a unified KPI data dictionary (entity-confirmed)"),
    ("60 days", "Automate the monthly consolidation; first report Day 4 → Day 2"),
    ("90 days", "Launch the first self-service management dashboard (single version of truth)"),
]
y=Inches(1.6)
for d,t in rows:
    rect(s, Inches(0.7), y, Inches(2.0), Inches(0.98), TEAL)
    txt(s, Inches(0.7), y+Inches(0.24), Inches(2.0), Inches(0.5),
        [[(d, 20, WHITE, True)]], align=PP_ALIGN.CENTER)
    rect(s, Inches(2.9), y, Inches(9.8), Inches(0.98), WHITE, line=LINE)
    txt(s, Inches(3.1), y+Inches(0.24), Inches(9.4), Inches(0.5),
        [[(t, 15, INK, False)]], anchor=MSO_ANCHOR.MIDDLE)
    y+=Inches(1.12)
rect(s, Inches(0.7), Inches(5.15), Inches(12.0), Inches(1.5), NAVY)
txt(s, Inches(0.95), Inches(5.34), Inches(11.6), Inches(1.15),
    [[("THE ASK", 15, GOLD, True)],
     [("Please authorize me to lead the KPI alignment and nominate one data contact per entity.", 18, WHITE, True)]],
    line_spacing=1.25, space_after=8)
notes(s, "Finally, a low-risk, verifiable commitment: exactly what I deliver at 30, 60 and 90 days. I have only one request today — authorize me to lead the KPI alignment and have each entity nominate a single data contact. It costs nothing, yet it starts the whole transformation. Thank you; I look forward to joining the team.")

# =====================================================================
# 14) Closing
# =====================================================================
s = add_slide()
rect(s, 0, 0, SW, SH, NAVY)
rect(s, 0, 0, Inches(0.22), SH, GOLD)
txt(s, Inches(0.95), Inches(2.4), Inches(11.6), Inches(1.0),
    [[("Free Finance from reports, so it can help you decide.", 32, WHITE, False)]], font=TITLE_FONT)
rect(s, Inches(0.98), Inches(3.5), Inches(3.4), Pt(2), GOLD)
txt(s, Inches(0.95), Inches(3.75), Inches(11.6), Inches(1.0),
    [[("Faster · More reliable · More forward-looking     |     Low cost · Group-aligned", 19, GOLD, True)]])
txt(s, Inches(0.95), Inches(6.5), Inches(11.6), Inches(0.5),
    [[("Thank you  ·  Happy to take your questions", 15, RGBColor(0x9F,0xB4,0xC9), False)]])
notes(s, "Closing: nail the message one more time — free Finance from reports so it can help you decide. Leave time for questions. Be ready for three likely ones: how do you ensure data quality, what if entities do not cooperate, and how does this avoid conflict with Group projects.")

prs.save("CFO_Finance_Controlling_Transformation_EN.pptx")
print("Saved: CFO_Finance_Controlling_Transformation_EN.pptx  |  slides:",
      len(prs.slides._sldIdLst))
