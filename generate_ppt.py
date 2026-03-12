"""
generate_ppt.py
---------------
Generates a professional 19-slide PowerPoint presentation:
"India's Private Investment Puzzle: Still Not Picking Up?"

Run:
    pip install -r requirements.txt
    python generate_ppt.py
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn

# ── Color palette ────────────────────────────────────────────────────────────
NAVY       = RGBColor(0x1B, 0x2A, 0x4A)   # title bars / headers
GREEN      = RGBColor(0x27, 0xAE, 0x60)   # New Economy / positive
RED        = RGBColor(0xE7, 0x4C, 0x3C)   # Old Economy / negative
AMBER      = RGBColor(0xF3, 0x9C, 0x12)   # On the Brink
DARK_GRAY  = RGBColor(0x33, 0x33, 0x33)   # body text
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_GRAY = RGBColor(0xF5, 0xF5, 0xF5)   # slide background
MID_GRAY   = RGBColor(0xCC, 0xCC, 0xCC)   # table borders
LIGHT_NAVY = RGBColor(0x2E, 0x4A, 0x7A)   # secondary nav shade

# ── Slide dimensions (16:9 widescreen) ───────────────────────────────────────
W = Inches(13.33)
H = Inches(7.5)

# ── Typography ────────────────────────────────────────────────────────────────
FONT = "Calibri"

# ─────────────────────────────────────────────────────────────────────────────
# Low-level helpers
# ─────────────────────────────────────────────────────────────────────────────

def add_rect(slide, x, y, w, h, fill_color, border_color=None, border_pt=0):
    """Add a filled rectangle (shape) to the slide."""
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        Inches(x), Inches(y), Inches(w), Inches(h)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if border_color and border_pt:
        shape.line.color.rgb = border_color
        shape.line.width = Pt(border_pt)
    else:
        shape.line.fill.background()
    return shape


def add_textbox(slide, x, y, w, h, text, font_size=14, bold=False,
                italic=False, color=DARK_GRAY, align=PP_ALIGN.LEFT,
                wrap=True, font_name=FONT):
    """Add a text box to the slide."""
    txBox = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = txBox.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return txBox


def add_label_on_rect(slide, x, y, w, h, fill_color, text, font_size=13,
                      bold=True, text_color=WHITE, align=PP_ALIGN.CENTER):
    """Add a colored rectangle with centered text on top."""
    add_rect(slide, x, y, w, h, fill_color)
    add_textbox(slide, x, y, w, h, text, font_size=font_size,
                bold=bold, color=text_color, align=align)


def title_bar(slide, title_text, bar_h=0.72):
    """Full-width dark-navy title bar at the top of a content slide."""
    add_rect(slide, 0, 0, 13.33, bar_h, NAVY)
    add_textbox(slide, 0.25, 0.05, 12.8, bar_h - 0.08,
                title_text, font_size=22, bold=True,
                color=WHITE, align=PP_ALIGN.LEFT)


def bullet_frame(slide, x, y, w, h, items, font_size=13,
                 color=DARK_GRAY, indent=0.25, level_indent=0.4,
                 line_spacing_pt=None):
    """
    Add a text box with multiple bullet lines.
    Each item is either a str or a (str, level) tuple.
    """
    txBox = slide.shapes.add_textbox(
        Inches(x), Inches(y), Inches(w), Inches(h))
    tf = txBox.text_frame
    tf.word_wrap = True

    for idx, item in enumerate(items):
        if isinstance(item, tuple):
            text, level = item
        else:
            text, level = item, 0

        if idx == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()

        p.level = level
        left_indent = Inches(indent + level * level_indent)
        p.space_before = Pt(2)
        if line_spacing_pt:
            p.line_spacing = Pt(line_spacing_pt)

        # hanging indent via XML
        pPr = p._pPr
        if pPr is None:
            pPr = p._p.get_or_add_pPr()
        pPr.set('indent', str(int(-Inches(0.2))))
        pPr.set('marL', str(int(left_indent)))

        run = p.add_run()
        run.text = text
        run.font.name = FONT
        run.font.size = Pt(font_size)
        run.font.color.rgb = color
    return txBox


def add_table(slide, x, y, w, rows, cols, headers, data,
              header_fill=NAVY, header_text_color=WHITE,
              row_fill=WHITE, alt_fill=LIGHT_GRAY,
              font_size=11, header_font_size=12):
    """Add a formatted table to the slide."""
    tbl = slide.shapes.add_table(
        rows, cols,
        Inches(x), Inches(y), Inches(w),
        Inches(0.38 * rows)
    ).table

    # Column widths equal
    col_w = Inches(w / cols)
    for c in range(cols):
        tbl.columns[c].width = col_w

    # Headers
    for c, hdr in enumerate(headers):
        cell = tbl.cell(0, c)
        cell.fill.solid()
        cell.fill.fore_color.rgb = header_fill
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = hdr
        run.font.name = FONT
        run.font.size = Pt(header_font_size)
        run.font.bold = True
        run.font.color.rgb = header_text_color

    # Data rows
    for r, row_data in enumerate(data):
        fill = alt_fill if r % 2 == 0 else row_fill
        for c, val in enumerate(row_data):
            cell = tbl.cell(r + 1, c)
            cell.fill.solid()
            cell.fill.fore_color.rgb = fill
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT if c == 0 else PP_ALIGN.CENTER
            run = p.add_run()
            run.text = str(val)
            run.font.name = FONT
            run.font.size = Pt(font_size)
            run.font.color.rgb = DARK_GRAY

    return tbl


def blank_slide(prs):
    """Add a completely blank slide."""
    blank_layout = prs.slide_layouts[6]  # Blank layout
    return prs.slides.add_slide(blank_layout)


def set_slide_bg(slide, color=WHITE):
    """Set solid background color on a slide."""
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = color


# ─────────────────────────────────────────────────────────────────────────────
# Individual Slide Builders
# ─────────────────────────────────────────────────────────────────────────────

def slide_01_title(prs):
    """Slide 1 – Title Slide"""
    slide = blank_slide(prs)
    set_slide_bg(slide, NAVY)

    # Decorative accent stripe (green)
    add_rect(slide, 0, 5.9, 13.33, 0.18, GREEN)

    # Main title box
    add_textbox(
        slide, 0.7, 1.2, 11.9, 2.2,
        "India's Private Investment Puzzle:\nStill Not Picking Up?",
        font_size=38, bold=True, color=WHITE,
        align=PP_ALIGN.CENTER
    )

    # Divider line
    add_rect(slide, 1.5, 3.5, 10.33, 0.06, AMBER)

    # Subtitle
    add_textbox(
        slide, 0.7, 3.7, 11.9, 0.6,
        "A Macroeconomic Analysis",
        font_size=22, bold=False, color=RGBColor(0xCC, 0xDD, 0xFF),
        align=PP_ALIGN.CENTER
    )

    # Presenters box
    add_textbox(
        slide, 0.7, 4.5, 11.9, 1.1,
        "Group D3:  Brunda Mary C P  •  Darshan A S  •  Dheemanth D  •  Dheeraj R  •  Likhith Yadav S  •  Gella Chidvilas Aryan",
        font_size=13, bold=False, color=RGBColor(0xAA, 0xBB, 0xDD),
        align=PP_ALIGN.CENTER
    )


def slide_02_paradox(prs):
    """Slide 2 – The Macroeconomic Paradox"""
    slide = blank_slide(prs)
    set_slide_bg(slide, WHITE)
    title_bar(slide, "The Macroeconomic Paradox")

    # Subheader
    add_textbox(
        slide, 0.3, 0.8, 12.7, 0.45,
        "Theory (Mankiw): Investment (I) is the most volatile component of GDP  →  AD = C + I + G + NX",
        font_size=12, italic=True, color=LIGHT_NAVY
    )

    # Paradox statement
    add_rect(slide, 0.3, 1.35, 12.7, 0.55, RGBColor(0xEB, 0xF5, 0xFB))
    add_textbox(
        slide, 0.5, 1.38, 12.3, 0.5,
        "The Paradox:  Government spending (G) is driving demand, but the Private Investment multiplier is lagging.",
        font_size=13, bold=True, color=NAVY
    )

    # Dashboard cards
    cards = [
        ("Real GDP Growth", "7.4%", "Fastest-growing major economy", GREEN),
        ("Public Capex Surge", "↑ 40% YoY", "In H1 – government-led push", LIGHT_NAVY),
        ("Aggregate Investment (GFCF)", "30%–33.8%", "of GDP – stagnating band", AMBER),
        ("The Execution Gap", "21.5% vs ▼", "RBI projected rise; NIPFP shows actual project completions contracted in 2024-25", RED),
    ]

    card_x = [0.3, 3.55, 6.8, 9.8]
    card_w = 3.0
    for i, (label, number, detail, color) in enumerate(cards):
        cx = card_x[i]
        # Card background
        add_rect(slide, cx, 2.05, card_w, 2.8,
                 RGBColor(0xF8, 0xF9, 0xFA),
                 border_color=color, border_pt=2)
        # Top stripe
        add_rect(slide, cx, 2.05, card_w, 0.30, color)
        # Label
        add_textbox(slide, cx + 0.08, 2.07, card_w - 0.16, 0.28,
                    label, font_size=10, bold=True, color=WHITE)
        # Big number
        add_textbox(slide, cx + 0.08, 2.42, card_w - 0.16, 0.7,
                    number, font_size=22, bold=True, color=color,
                    align=PP_ALIGN.CENTER)
        # Detail
        add_textbox(slide, cx + 0.1, 3.15, card_w - 0.2, 1.6,
                    detail, font_size=10, color=DARK_GRAY)

    # Bottom note
    add_textbox(
        slide, 0.3, 5.1, 12.7, 0.4,
        "Sources: RBI Monetary Policy Report, NIPFP Investment Tracker, MoSPI National Accounts",
        font_size=9, italic=True, color=MID_GRAY
    )


def slide_03_two_economies(prs):
    """Slide 3 – A 'Tale of Two Economies'"""
    slide = blank_slide(prs)
    set_slide_bg(slide, WHITE)
    title_bar(slide, 'A "Tale of Two Economies"')

    # Left panel – New Economy
    add_rect(slide, 0.25, 0.85, 6.15, 5.9, RGBColor(0xEA, 0xF7, 0xEE))
    add_rect(slide, 0.25, 0.85, 6.15, 0.5, GREEN)
    add_textbox(slide, 0.35, 0.87, 5.9, 0.46,
                "🟢  NEW ECONOMY  —  Full Throttle",
                font_size=14, bold=True, color=WHITE)

    new_items = [
        "Sectors:",
        "  • Semiconductors",
        "  • Electric Vehicles (EVs)",
        "  • Renewable Energy",
        "  • Electronics & Contract Mfg",
        "",
        "Trajectory:",
        "  Aggressive greenfield expansions; deeply integrating",
        "  into global supply chains via China+1 realignment.",
        "  MEC ranges: 15%–25% — clears every hurdle rate.",
    ]
    add_textbox(slide, 0.45, 1.45, 5.8, 5.0,
                "\n".join(new_items), font_size=12, color=DARK_GRAY)

    # Right panel – Old Economy
    add_rect(slide, 6.93, 0.85, 6.15, 5.9, RGBColor(0xFD, 0xED, 0xEC))
    add_rect(slide, 6.93, 0.85, 6.15, 0.5, RED)
    add_textbox(slide, 7.03, 0.87, 5.9, 0.46,
                "🔴  OLD ECONOMY  —  Rational Caution",
                font_size=14, bold=True, color=WHITE)

    old_items = [
        "Sectors:",
        "  • FMCG (HUL, Dabur, Parle-G)",
        "  • Basic Metals (JSW Steel)",
        "  • Textiles (Welspun India)",
        "  • Core Infrastructure",
        "",
        "Trajectory:",
        "  Deferring capacity additions; focusing on",
        "  maximising existing plant efficiency.",
        "  MEC ranges: 8%–10% — barely clears hurdle rate.",
    ]
    add_textbox(slide, 7.03, 1.45, 5.8, 5.0,
                "\n".join(old_items), font_size=12, color=DARK_GRAY)

    # VS badge in the middle
    add_rect(slide, 6.28, 3.1, 0.77, 0.77, AMBER)
    add_textbox(slide, 6.28, 3.1, 0.77, 0.77,
                "VS", font_size=18, bold=True, color=WHITE,
                align=PP_ALIGN.CENTER)


def slide_04_theory(prs):
    """Slide 4 – The Investment Decision (Theory)"""
    slide = blank_slide(prs)
    set_slide_bg(slide, WHITE)
    title_bar(slide, "The Investment Decision — Theory")

    # Keynesian rule box
    add_rect(slide, 0.3, 0.85, 12.7, 1.35, RGBColor(0xEB, 0xF5, 0xFB))
    add_textbox(slide, 0.5, 0.9, 12.3, 0.35,
                "Keynesian Investment Rule", font_size=13, bold=True, color=NAVY)
    add_textbox(slide, 0.5, 1.25, 12.3, 0.85,
                "A firm invests ONLY if   MEC > Real Interest Rate (r)\n"
                "Fisher Equation:   r  =  i  −  π      (Real Rate = Nominal Rate − Inflation)",
                font_size=13, color=DARK_GRAY)

    # Section header
    add_textbox(slide, 0.3, 2.35, 8.0, 0.38,
                "India's Current Hurdle Rate", font_size=15, bold=True, color=NAVY)

    # Rate cards
    rate_data = [
        ("Nominal Repo Rate  (i)", "5.25%", "(after cumulative 125 bps cuts)", LIGHT_NAVY),
        ("Current CPI Inflation  (π)", "1.7%", "(remarkably subdued)", GREEN),
        ("Effective Real Rate  (r)", "3.55%", "(= 5.25% − 1.7%  — remains restrictive)", RED),
    ]
    for idx, (label, val, note, col) in enumerate(rate_data):
        ry = 2.82 + idx * 0.9
        add_rect(slide, 0.3, ry, 12.7, 0.78, RGBColor(0xF8, 0xF9, 0xFA),
                 border_color=col, border_pt=1)
        add_rect(slide, 0.3, ry, 0.18, 0.78, col)
        add_textbox(slide, 0.6, ry + 0.06, 6.5, 0.35,
                    label, font_size=13, bold=True, color=DARK_GRAY)
        add_textbox(slide, 7.2, ry + 0.04, 2.0, 0.4,
                    val, font_size=20, bold=True, color=col,
                    align=PP_ALIGN.CENTER)
        add_textbox(slide, 9.3, ry + 0.1, 3.5, 0.55,
                    note, font_size=10, italic=True, color=DARK_GRAY)

    # Macro takeaway
    add_rect(slide, 0.3, 5.58, 12.7, 0.72, NAVY)
    add_textbox(slide, 0.5, 5.62, 12.3, 0.65,
                "📌  Macro Takeaway:  Because inflation is so low, real borrowing costs remain highly restrictive despite RBI rate cuts.",
                font_size=12, bold=True, color=WHITE)


def slide_05_corporate_reality(prs):
    """Slide 5 – The Investment Decision (Corporate Reality)"""
    slide = blank_slide(prs)
    set_slide_bg(slide, WHITE)
    title_bar(slide, "The Investment Decision — Corporate Reality")

    add_textbox(slide, 0.3, 0.82, 12.7, 0.38,
                "Go / No-Go Decision  |  Hurdle Rate = 3.55% real",
                font_size=13, italic=True, color=LIGHT_NAVY)

    # LEFT – New Economy
    add_rect(slide, 0.25, 1.3, 6.15, 5.35, RGBColor(0xEA, 0xF7, 0xEE))
    add_rect(slide, 0.25, 1.3, 6.15, 0.45, GREEN)
    add_textbox(slide, 0.35, 1.32, 5.9, 0.42,
                "🟢  NEW ECONOMY  —  MEC > Hurdle Rate",
                font_size=13, bold=True, color=WHITE)
    new_lines = [
        "MEC Range:  15% – 25%",
        "",
        "Tata Electronics (iPhone mfg.)",
        "  → MEC: 18–22%  ✓ Clears hurdle",
        "",
        "Adani Green Energy (solar)",
        "  → IRRs: 14–16%  ✓ Clears hurdle",
        "",
        "Both easily exceed the 3.55% real cost of capital.",
        "Investment is proceeding at full speed.",
    ]
    add_textbox(slide, 0.45, 1.85, 5.8, 4.6,
                "\n".join(new_lines), font_size=12, color=DARK_GRAY)

    # RIGHT – Old Economy
    add_rect(slide, 6.93, 1.3, 6.15, 5.35, RGBColor(0xFD, 0xED, 0xEC))
    add_rect(slide, 6.93, 1.3, 6.15, 0.45, RED)
    add_textbox(slide, 7.03, 1.32, 5.9, 0.42,
                "🔴  OLD ECONOMY  —  MEC ≈ Hurdle Rate",
                font_size=13, bold=True, color=WHITE)
    old_lines = [
        "MEC Range:  8% – 10%",
        "",
        "JSW Steel — Dolvi Phase 3",
        "  → MEC: ~9–10% at current prices",
        "  → Prompted deferral of expansion",
        "",
        "ITC — Food Processing Plant",
        "  → Deferred ₹3,000 crore plant",
        "  → Failed to clear cost-of-capital test",
        "",
        "Rational managerial discipline — not panic.",
    ]
    add_textbox(slide, 7.03, 1.85, 5.8, 4.6,
                "\n".join(old_lines), font_size=12, color=DARK_GRAY)


def slide_06_multiplier(prs):
    """Slide 6 – The Broken Keynesian Multiplier"""
    slide = blank_slide(prs)
    set_slide_bg(slide, WHITE)
    title_bar(slide, "The Broken Keynesian Multiplier")

    # Theory box
    add_rect(slide, 0.3, 0.85, 12.7, 0.65, RGBColor(0xEB, 0xF5, 0xFB))
    add_textbox(slide, 0.5, 0.88, 12.3, 0.6,
                "Theory:  Multiplier (k) = 1 / (1 − MPC).  Massive public infrastructure spending should cascade into a broad consumption & investment boom.",
                font_size=12, color=DARK_GRAY)

    add_rect(slide, 0.3, 1.6, 12.7, 0.42, RED)
    add_textbox(slide, 0.5, 1.63, 12.3, 0.38,
                "Reality:  The multiplier is severely dampened by three key domestic leakages.",
                font_size=13, bold=True, color=WHITE)

    # Three leakage cards
    leakages = [
        ("Leakage 1 — Debt", NAVY,
         "Household debt surged to a record\n41.3% of GDP.\n\nIncome is leaking into EMI\npayments before reaching\nthe real economy."),
        ("Leakage 2 — Savings", AMBER,
         "Net household financial savings\ncollapsed to a decadal low of\n~5.3% of GDP.\n\nHouseholds are borrowing to\nconsume, not to invest."),
        ("Leakage 3 — Imports", RED,
         "~80% of solar equipment &\nmajority of semiconductor\ncomponents are imported.\n\nMultiplier leaks abroad rather\nthan circulating domestically."),
    ]
    for i, (label, col, detail) in enumerate(leakages):
        lx = 0.3 + i * 4.35
        add_rect(slide, lx, 2.15, 4.0, 3.6, RGBColor(0xF8, 0xF9, 0xFA),
                 border_color=col, border_pt=2)
        add_rect(slide, lx, 2.15, 4.0, 0.42, col)
        add_textbox(slide, lx + 0.1, 2.17, 3.8, 0.38,
                    label, font_size=12, bold=True, color=WHITE)
        add_textbox(slide, lx + 0.15, 2.65, 3.7, 3.0,
                    detail, font_size=11, color=DARK_GRAY)

    # Bottom note
    add_rect(slide, 0.3, 5.88, 12.7, 0.62, RGBColor(0xFE, 0xF9, 0xE7))
    add_textbox(slide, 0.5, 5.91, 12.3, 0.55,
                "Net result: Even as public Capex soars, the demand signal reaching firms is far weaker than the headline spending figure implies.",
                font_size=11, italic=True, color=DARK_GRAY)


def slide_07_demand_shock(prs):
    """Slide 7 – The Corporate Demand Shock"""
    slide = blank_slide(prs)
    set_slide_bg(slide, WHITE)
    title_bar(slide, "The Corporate Demand Shock")

    add_rect(slide, 0.3, 0.82, 12.7, 0.45, RGBColor(0xFD, 0xED, 0xEC))
    add_textbox(slide, 0.5, 0.85, 12.3, 0.4,
                "Fragile, debt-driven consumption cannot send strong demand signals for new factory builds.",
                font_size=12, bold=True, color=RED)

    add_textbox(slide, 0.3, 1.42, 6.0, 0.38,
                "Corporate Execution in FY25-26", font_size=14, bold=True, color=NAVY)

    rows_data = [
        ("FMCG Deferrals",
         "HUL & Dabur flagged 'weak mass-market demand'. Rural volume growth stuck at 3–4%. "
         "HUL deferred a planned personal care unit."),
        ("Inventory Stress",
         "Parle-G chose NOT to add new production lines in FY25 due to insufficient volume "
         "certainty in rural markets."),
        ("Auto Sector Hesitation",
         "Hero MotoCorp delayed its ₹2,000 crore Rajasthan greenfield plant by 18 months "
         "due to sluggish rural recovery."),
    ]

    headers = ["Sector / Company", "What Happened"]
    data = rows_data
    add_table(slide, 0.3, 1.9, 12.7, 4, 2, headers, data,
              font_size=11, header_font_size=12)

    add_textbox(
        slide, 0.3, 5.55, 12.7, 0.5,
        "Common thread: Weak rural demand signals — not financial constraints — are holding back Old Economy investment.",
        font_size=11, italic=True, color=LIGHT_NAVY
    )


def slide_08_cu_theory(prs):
    """Slide 8 – The Capacity Utilization Trap (Theory)"""
    slide = blank_slide(prs)
    set_slide_bg(slide, WHITE)
    title_bar(slide, "The Capacity Utilization Trap — Theory")

    # Gauge visual (text-based)
    add_rect(slide, 0.3, 0.85, 4.8, 4.5, RGBColor(0xF8, 0xF9, 0xFA),
             border_color=MID_GRAY, border_pt=1)
    add_textbox(slide, 0.35, 0.92, 4.6, 0.38,
                "Capacity Gauge", font_size=13, bold=True, color=NAVY,
                align=PP_ALIGN.CENTER)

    # Bar chart approximation
    add_rect(slide, 0.9, 1.42, 3.7, 0.55, RGBColor(0xE8, 0xE8, 0xE8))  # bg
    add_rect(slide, 0.9, 1.42, 3.7 * 0.743, 0.55, AMBER)                 # 74.3%
    add_textbox(slide, 0.9, 2.03, 3.7, 0.4,
                "Current: 74.3%  (AMBER zone)", font_size=12, bold=True,
                color=AMBER, align=PP_ALIGN.CENTER)

    # Trigger zone marker
    add_rect(slide, 0.9, 2.52, 3.7, 0.55, RGBColor(0xE8, 0xE8, 0xE8))
    add_rect(slide, 0.9, 2.52, 3.7 * 0.80, 0.55, GREEN)
    add_textbox(slide, 0.9, 3.13, 3.7, 0.4,
                "Trigger Zone: 78%–80%  (GREEN zone)", font_size=11,
                color=GREEN, align=PP_ALIGN.CENTER)

    add_textbox(slide, 0.35, 3.62, 4.6, 1.5,
                "Traditional manufacturing is trapped in a waiting game —\n"
                "rational firms will not invest until the gauge turns green.",
                font_size=11, italic=True, color=DARK_GRAY,
                align=PP_ALIGN.CENTER)

    # Theory panel
    add_rect(slide, 5.4, 0.85, 7.65, 4.5, RGBColor(0xEB, 0xF5, 0xFB),
             border_color=LIGHT_NAVY, border_pt=1)
    add_textbox(slide, 5.55, 0.9, 7.3, 0.38,
                "Theory  —  Operating in the SRAS Zone", font_size=13,
                bold=True, color=NAVY)

    theory_lines = [
        "• Firms operate along the SRAS curve by adding",
        "  shifts / labour to existing factories.",
        "",
        "• Financially irrational to deploy capital to shift",
        "  the LRAS curve until existing capacity is exhausted.",
        "",
        "Key Numbers  (RBI OBICUS Survey):",
        "",
        "  Current Aggregate CU:         74.3%",
        "  Macro Threshold (capex boom):  78%–80%",
        "",
        "  Below the threshold → firms rationally defer.",
        "  Above the threshold → broad-based capex boom.",
    ]
    add_textbox(slide, 5.55, 1.35, 7.3, 3.85,
                "\n".join(theory_lines), font_size=12, color=DARK_GRAY)

    # Bottom source
    add_textbox(slide, 0.3, 5.52, 12.7, 0.38,
                "Source: RBI Order Books, Inventories and Capacity Utilisation Survey (OBICUS), Q2 FY26",
                font_size=9, italic=True, color=MID_GRAY)


def slide_09_cu_corporate(prs):
    """Slide 9 – The Capacity Utilization Trap (Corporate Reality)"""
    slide = blank_slide(prs)
    set_slide_bg(slide, WHITE)
    title_bar(slide, "The Capacity Utilization Trap — Corporate Reality")

    add_textbox(slide, 0.3, 0.82, 12.7, 0.38,
                "Traffic-light framework: Where are India's major manufacturers right now?",
                font_size=12, italic=True, color=LIGHT_NAVY)

    # Three traffic-light cards
    cards_data = [
        (RED, "🔴  HOLDING BACK", "Utilization < 75%",
         "JSW Steel:\n"
         "  Running at 72–74%.\n"
         "  Deferred ₹40,000+ crore capacity expansion.\n"
         "  Awaiting demand clarity before committing."),
        (AMBER, "🟡  ON THE BRINK", "Utilization ~78%",
         "Welspun India:\n"
         "  Operating MMF textile units at 78%.\n"
         "  Holding off on Phase 2 expansion.\n"
         "  Trigger: Sustained 80%+ for two quarters."),
        (GREEN, "🟢  TRIGGER FIRED", "Utilization > 80%",
         "Maruti Suzuki:\n"
         "  Manesar & Gujarat plants crossed 85%.\n"
         "  → Triggered ₹11,000 crore Kharkhoda plant.\n\n"
         "Tata Motors:\n"
         "  EV lines in Pune above 80%.\n"
         "  → New battery-pack capacity immediately commissioned."),
    ]

    for i, (col, header, sub, detail) in enumerate(cards_data):
        cx = 0.25 + i * 4.37
        add_rect(slide, cx, 1.35, 4.1, 5.2, RGBColor(0xF8, 0xF9, 0xFA),
                 border_color=col, border_pt=3)
        add_rect(slide, cx, 1.35, 4.1, 0.55, col)
        add_textbox(slide, cx + 0.1, 1.37, 3.9, 0.28,
                    header, font_size=12, bold=True, color=WHITE)
        add_textbox(slide, cx + 0.1, 1.66, 3.9, 0.28,
                    sub, font_size=10, italic=True, color=WHITE)
        add_textbox(slide, cx + 0.15, 2.02, 3.8, 4.35,
                    detail, font_size=11, color=DARK_GRAY)


def slide_10_pli(prs):
    """Slide 10 – Supply-Side Interventions & PLI"""
    slide = blank_slide(prs)
    set_slide_bg(slide, WHITE)
    title_bar(slide, "Supply-Side Interventions & PLI")

    # Theory
    add_rect(slide, 0.3, 0.85, 12.7, 1.0, RGBColor(0xEB, 0xF5, 0xFB))
    add_textbox(slide, 0.5, 0.88, 12.3, 0.4,
                "Theory  —  Supply-Side Economics & Market Failure", font_size=13, bold=True, color=NAVY)
    add_textbox(slide, 0.5, 1.28, 12.3, 0.5,
                "When the free market fails to clear the hurdle rate for critical technologies, the State acts as a 'risk absorber' — artificially boosting MEC to crowd-in private investment.",
                font_size=12, color=DARK_GRAY)

    # PLI headline
    add_rect(slide, 0.3, 2.0, 12.7, 0.55, NAVY)
    add_textbox(slide, 0.5, 2.03, 12.3, 0.5,
                "The Policy Anchor:  Production Linked Incentive (PLI) Scheme",
                font_size=15, bold=True, color=WHITE)

    # PLI details
    pli_points = [
        "Total Outlay:  ₹1.91 lakh crore  across  14 strategic sectors",
        "Mechanism:  Moves from input subsidies → links financial rewards directly to incremental sales & production",
        "Effect on MEC:  Artificially raises the effective return on new capacity, making the Go / No-Go calculus viable in critical tech sectors",
        "Target sectors:  Semiconductors • EVs • Batteries • Specialty Chemicals • Pharmaceuticals • Mobile Phones • White Goods • Solar PV • Textiles",
    ]
    for idx, pt in enumerate(pli_points):
        ry = 2.68 + idx * 0.72
        add_rect(slide, 0.3, ry, 12.7, 0.62, RGBColor(0xF8, 0xF9, 0xFA),
                 border_color=LIGHT_NAVY, border_pt=1)
        add_rect(slide, 0.3, ry, 0.15, 0.62, GREEN)
        add_textbox(slide, 0.55, ry + 0.06, 12.2, 0.5,
                    pt, font_size=12, color=DARK_GRAY)

    add_textbox(slide, 0.3, 5.65, 12.7, 0.38,
                "Source: DPIIT PLI Annual Report FY25; Ministry of Finance Budget Documents",
                font_size=9, italic=True, color=MID_GRAY)


def slide_11_sectoral(prs):
    """Slide 11 – Sectoral Divergence in Execution"""
    slide = blank_slide(prs)
    set_slide_bg(slide, WHITE)
    title_bar(slide, "Sectoral Divergence in Execution")

    # LEFT – Winners
    add_rect(slide, 0.25, 0.85, 6.1, 5.9, RGBColor(0xEA, 0xF7, 0xEE))
    add_rect(slide, 0.25, 0.85, 6.1, 0.5, GREEN)
    add_textbox(slide, 0.35, 0.87, 5.9, 0.46,
                "🟢  THE BOOMING WINNERS  —  Electronics",
                font_size=13, bold=True, color=WHITE)

    win_lines = [
        "Domestic electronics production:",
        "  ↑ 146%  to  ₹5.45 lakh crore",
        "",
        "Mobile phone imports:",
        "  ↓ 77%  (import-substitution working)",
        "",
        "Tata Electronics & Powerchip:",
        "  $11 billion semiconductor fab — Dholera",
        "",
        "Micron Technology:",
        "  ₹22,000 crore ATMP facility — Sanand",
    ]
    add_textbox(slide, 0.45, 1.45, 5.7, 5.0,
                "\n".join(win_lines), font_size=12, color=DARK_GRAY)

    # RIGHT – Laggards
    add_rect(slide, 6.95, 0.85, 6.1, 5.9, RGBColor(0xFD, 0xED, 0xEC))
    add_rect(slide, 6.95, 0.85, 6.1, 0.5, RED)
    add_textbox(slide, 7.05, 0.87, 5.9, 0.46,
                "🔴  THE LAGGING TRADITIONAL SECTORS",
                font_size=13, bold=True, color=WHITE)

    lag_lines = [
        "Index of Eight Core Industries:",
        "  Near-zero growth in October 2025",
        "",
        "PLI for MMF Textiles:",
        "  Only ₹1,100 crore invested",
        "  vs ₹10,683 crore target  (10% uptake)",
        "",
        "Core infrastructure:",
        "  Capacity utilization stuck below 75%",
        "  across cement, steel, power sectors",
    ]
    add_textbox(slide, 7.05, 1.45, 5.7, 5.0,
                "\n".join(lag_lines), font_size=12, color=DARK_GRAY)


def slide_12_balance_sheets(prs):
    """Slide 12 – Corporate Balance Sheets vs. Crowding Out"""
    slide = blank_slide(prs)
    set_slide_bg(slide, WHITE)
    title_bar(slide, "Corporate Balance Sheets vs. Crowding Out")

    # Twin Balance Sheet Advantage
    add_rect(slide, 0.3, 0.85, 12.7, 1.05, RGBColor(0xEA, 0xF7, 0xEE))
    add_textbox(slide, 0.5, 0.88, 12.3, 0.38,
                '"Twin Balance Sheet Advantage"', font_size=14, bold=True, color=GREEN)
    add_textbox(slide, 0.5, 1.24, 12.3, 0.6,
                "Unlike the 2010s (NPA crisis era), corporate India is highly deleveraged with record profitability. Internal financial ability to invest is exceptionally strong.",
                font_size=12, color=DARK_GRAY)

    # IS-LM Theory box
    add_rect(slide, 0.3, 2.05, 5.9, 4.2, RGBColor(0xEB, 0xF5, 0xFB),
             border_color=LIGHT_NAVY, border_pt=1)
    add_rect(slide, 0.3, 2.05, 5.9, 0.42, LIGHT_NAVY)
    add_textbox(slide, 0.45, 2.07, 5.6, 0.38,
                "IS-LM Theory  —  Crowding Out Mechanism", font_size=12,
                bold=True, color=WHITE)
    theory_steps = [
        "1.  Massive government infrastructure spending",
        "    shifts IS curve outward  →",
        "2.  Sovereign borrowing absorbs domestic savings",
        "    →",
        "3.  Upward pressure on equilibrium interest rates",
        "    →",
        "4.  'Crowding out' private investment",
        "",
        "   r*  ↑  →  MEC of marginal projects",
        "   no longer clears hurdle  →  deferral",
    ]
    add_textbox(slide, 0.45, 2.55, 5.6, 3.55,
                "\n".join(theory_steps), font_size=11, color=DARK_GRAY)

    # Reality / nuance
    add_rect(slide, 6.5, 2.05, 6.6, 4.2, RGBColor(0xFE, 0xF9, 0xE7),
             border_color=AMBER, border_pt=1)
    add_rect(slide, 6.5, 2.05, 6.6, 0.42, AMBER)
    add_textbox(slide, 6.65, 2.07, 6.3, 0.38,
                "India's Reality  —  Partial Crowding Out", font_size=12,
                bold=True, color=WHITE)
    reality_points = [
        "Large corporates have bypassed the squeeze:",
        "  → Record profit retention funds capex internally",
        "  → Access to corporate bond markets & private credit",
        "",
        "MSMEs face the full crowding-out effect:",
        "  → Dependent on commercial bank lending",
        "  → Cost of debt remains above their MEC",
        "  → Expansions are being delayed",
        "",
        "Net verdict: Crowding out is real but sector-specific.",
    ]
    add_textbox(slide, 6.65, 2.55, 6.3, 3.55,
                "\n".join(reality_points), font_size=11, color=DARK_GRAY)


def slide_13_private_credit(prs):
    """Slide 13 – Bypassing the Squeeze (Private Credit)"""
    slide = blank_slide(prs)
    set_slide_bg(slide, WHITE)
    title_bar(slide, "Bypassing the Squeeze — Private Credit")

    add_textbox(slide, 0.3, 0.82, 12.7, 0.38,
                "Two Diverging Paths in the Capital Market",
                font_size=13, bold=True, color=NAVY)

    # LEFT – Large Corporations
    add_rect(slide, 0.25, 1.3, 6.1, 5.25, RGBColor(0xEA, 0xF7, 0xEE))
    add_rect(slide, 0.25, 1.3, 6.1, 0.48, GREEN)
    add_textbox(slide, 0.35, 1.32, 5.9, 0.44,
                "🟢  LARGE CORPORATIONS  —  Immune",
                font_size=13, bold=True, color=WHITE)
    large_lines = [
        "Bypassing domestic banking squeeze via:",
        "",
        "  Corporate Bond Market:",
        "    Direct placement at competitive spreads",
        "",
        "  Private Credit Market:",
        "    Surged 35%  →  US$12.4 billion in CY 2025",
        "    (Blackstone, KKR, Ares deploying actively)",
        "",
        "  Internal Accruals:",
        "    Record FY25 profits fund capex internally",
        "",
        "Effect: Effectively insulated from rate environment.",
    ]
    add_textbox(slide, 0.45, 1.88, 5.7, 4.5,
                "\n".join(large_lines), font_size=11, color=DARK_GRAY)

    # RIGHT – MSMEs
    add_rect(slide, 6.93, 1.3, 6.1, 5.25, RGBColor(0xFD, 0xED, 0xEC))
    add_rect(slide, 6.93, 1.3, 6.1, 0.48, RED)
    add_textbox(slide, 7.03, 1.32, 5.9, 0.44,
                "🔴  MSMEs & TRADITIONAL MFG  —  Priced Out",
                font_size=13, bold=True, color=WHITE)
    msme_lines = [
        "Rely entirely on:",
        "",
        "  Commercial Bank Loans",
        "    Spread above repo; effective rate ~9–10%",
        "",
        "  NBFCs",
        "    Even higher effective borrowing costs",
        "",
        "The Crowding-Out Math:",
        "  Borrowing cost:  ~9–10%",
        "  Old Economy MEC: ~8–10%",
        "  → Margin of safety = zero or negative",
        "",
        "Factory expansions are being deferred.",
    ]
    add_textbox(slide, 7.03, 1.88, 5.7, 4.5,
                "\n".join(msme_lines), font_size=11, color=DARK_GRAY)


def slide_14_energy(prs):
    """Slide 14 – The Energy-Digital Mismatch"""
    slide = blank_slide(prs)
    set_slide_bg(slide, WHITE)
    title_bar(slide, "The Energy-Digital Mismatch")

    # Theory box
    add_rect(slide, 0.3, 0.85, 12.7, 0.85, RGBColor(0xEB, 0xF5, 0xFB))
    add_textbox(slide, 0.5, 0.88, 12.3, 0.4,
                "Theory  —  Complementary Capital", font_size=13, bold=True, color=NAVY)
    add_textbox(slide, 0.5, 1.26, 12.3, 0.38,
                "To sustainably shift LRAS rightward via high-tech manufacturing, proportionate expansion in energy infrastructure is essential.",
                font_size=11, color=DARK_GRAY)

    # Challenge box
    add_rect(slide, 0.3, 1.82, 12.7, 0.88, RGBColor(0xFD, 0xED, 0xEC))
    add_textbox(slide, 0.5, 1.85, 12.3, 0.38,
                "The Challenge", font_size=13, bold=True, color=RED)
    add_textbox(slide, 0.5, 2.2, 12.3, 0.45,
                '"New Economy" anchors ($11 billion fabs & AI data centers) are massively energy-intensive. Without synchronized power boom, operational costs rise — destroying MEC.',
                font_size=11, color=DARK_GRAY)

    # Policy Response
    add_textbox(slide, 0.3, 2.85, 5.5, 0.38,
                "Policy Response — Energy Security Roadmap", font_size=13, bold=True, color=NAVY)

    policy_items = [
        (NAVY, "SHANTI Act 2025",
         "₹20,000 crore Nuclear Energy Mission for Small Modular Reactors (SMRs).\n"
         "Target: 100 GW nuclear capacity by 2047. Low-carbon baseload for fabs & data centres."),
        (GREEN, "CCUS Initiative",
         "₹20,000 crore for Carbon Capture & Utilisation infrastructure in heavy industries.\n"
         "Focus: Green steel production — makes India's steel export-competitive."),
        (AMBER, "Renewable Push",
         "450 GW renewable target by 2030. Solar + storage to power EV charging networks.\n"
         "Adani Green, NTPC, and Torrent Power all commissioning gigawatt-scale projects."),
    ]
    for idx, (col, title, detail) in enumerate(policy_items):
        ry = 3.32 + idx * 0.88
        add_rect(slide, 0.3, ry, 12.7, 0.78, RGBColor(0xF8, 0xF9, 0xFA),
                 border_color=col, border_pt=2)
        add_rect(slide, 0.3, ry, 0.18, 0.78, col)
        add_textbox(slide, 0.6, ry + 0.04, 2.8, 0.34,
                    title, font_size=12, bold=True, color=col)
        add_textbox(slide, 3.5, ry + 0.06, 9.4, 0.65,
                    detail, font_size=10, color=DARK_GRAY)


def slide_15_green_infra(prs):
    """Slide 15 – Corporate Action in Green Infrastructure"""
    slide = blank_slide(prs)
    set_slide_bg(slide, WHITE)
    title_bar(slide, "Corporate Action in Green Infrastructure")

    add_textbox(slide, 0.3, 0.82, 12.7, 0.38,
                "How Private Capital is Responding to the Energy-Digital Opportunity",
                font_size=13, italic=True, color=LIGHT_NAVY)

    companies = [
        (GREEN, "Larsen & Toubro (L&T)",
         "Constructing India's largest green hydrogen plant at Indian Oil's Panipat refinery.\n"
         "Positions L&T as the dominant EPC contractor for the energy transition — recurring revenue model."),
        (RGBColor(0x1A, 0x5C, 0x8A), "Adani Green Energy",
         "Deploying ₹2.4 lakh crore capex pipeline to feed digital & industrial power demand.\n"
         "Locked-in long-term Power Purchase Agreements (PPAs) yield IRRs of 14–16%,\n"
         "comfortably beating the 3.55% real cost of capital."),
        (AMBER, "NTPC Green Energy",
         "100 GW renewable capacity target. Hybrid solar-wind-storage complexes.\n"
         "Offtake guaranteed by PSU and industrial consumers — virtually zero merchant risk."),
        (RGBColor(0x6C, 0x3A, 0x8F), "Tata Power",
         "Pumped hydro storage + rooftop solar + EV charging network.\n"
         "Complementary capex: each MW of storage unlocks additional renewable MEC."),
    ]

    for idx, (col, name, detail) in enumerate(companies):
        ry = 1.35 + idx * 1.3
        add_rect(slide, 0.3, ry, 12.7, 1.15, RGBColor(0xF8, 0xF9, 0xFA),
                 border_color=col, border_pt=2)
        add_rect(slide, 0.3, ry, 0.22, 1.15, col)
        add_textbox(slide, 0.65, ry + 0.08, 3.8, 0.38,
                    name, font_size=13, bold=True, color=col)
        add_textbox(slide, 0.65, ry + 0.46, 12.1, 0.65,
                    detail, font_size=11, color=DARK_GRAY)

    add_textbox(slide, 0.3, 6.6, 12.7, 0.38,
                "Net result: Energy-transition investments achieve IRRs of 14–16% — well above the 3.55% hurdle rate.",
                font_size=11, bold=True, color=GREEN)


def slide_16_fdi(prs):
    """Slide 16 – Global Competitive Positioning"""
    slide = blank_slide(prs)
    set_slide_bg(slide, WHITE)
    title_bar(slide, "Global Competitive Positioning")

    # Theory
    add_rect(slide, 0.3, 0.85, 12.7, 1.0, RGBColor(0xEB, 0xF5, 0xFB))
    add_textbox(slide, 0.5, 0.88, 12.3, 0.38,
                "Theory  —  Open Economy Macro & International Risk Premiums", font_size=13, bold=True, color=NAVY)
    add_textbox(slide, 0.5, 1.24, 12.3, 0.55,
                "When international risk premiums shift (e.g., US-China trade fragmentation), global capital reallocates to stable economies.\n"
                "Competitiveness = 'Strategic Indispensability' in global supply chains.",
                font_size=11, color=DARK_GRAY)

    # Key Numbers
    add_textbox(slide, 0.3, 2.0, 5.0, 0.38,
                "Key Numbers", font_size=14, bold=True, color=NAVY)

    metrics = [
        ("Global Greenfield FDI  (H1 2025)", "↑ 7%",
         "Project values surged globally even as overall FDI dropped.", GREEN),
        ("India Gross FDI Inflows  (FY25)", "USD 81.04 Bn",
         "Resilient inflows despite global headwinds.", LIGHT_NAVY),
        ("China+1 Beneficiary", "#1 Ranked",
         "India capturing largest share of supply-chain realignment.", AMBER),
    ]

    for idx, (label, val, note, col) in enumerate(metrics):
        ry = 2.5 + idx * 1.05
        add_rect(slide, 0.3, ry, 12.7, 0.88, RGBColor(0xF8, 0xF9, 0xFA),
                 border_color=col, border_pt=2)
        add_rect(slide, 0.3, ry, 0.18, 0.88, col)
        add_textbox(slide, 0.6, ry + 0.05, 6.0, 0.38,
                    label, font_size=12, bold=True, color=DARK_GRAY)
        add_textbox(slide, 6.7, ry + 0.04, 2.5, 0.42,
                    val, font_size=18, bold=True, color=col,
                    align=PP_ALIGN.CENTER)
        add_textbox(slide, 9.3, ry + 0.1, 3.6, 0.65,
                    note, font_size=10, italic=True, color=DARK_GRAY)

    add_textbox(slide, 0.3, 5.72, 12.7, 0.5,
                "Source: UNCTAD World Investment Report 2025; RBI FDI Statistics; Financial Times fDi Markets database",
                font_size=9, italic=True, color=MID_GRAY)


def slide_17_china_plus_one(prs):
    """Slide 17 – Capturing the 'China+1' Shift"""
    slide = blank_slide(prs)
    set_slide_bg(slide, WHITE)
    title_bar(slide, 'Capturing the "China+1" Shift')

    add_textbox(slide, 0.3, 0.82, 12.7, 0.38,
                "India's window of strategic opportunity in global supply-chain realignment",
                font_size=12, italic=True, color=LIGHT_NAVY)

    data_points = [
        (GREEN, "The Apple / iPhone Shift",
         "Apple successfully moved 14% of its global iPhone production to India in 2025.\n"
         "Tata Electronics & Foxconn now assembling for global export — not just domestic consumption."),
        (LIGHT_NAVY, "Contract Manufacturing Surge",
         "Samsung, Foxconn, Jabil, and Pegatron are aggressively sourcing locally for export-grade electronics.\n"
         "India's electronics export value crossed $29 billion in FY25 — up 24% YoY."),
        (AMBER, "Semiconductor Market Projection",
         "India Semiconductor Mission projects domestic semiconductor market to hit ₹9.7 trillion by 2030.\n"
         "Tata–Powerchip $11 billion fab & Micron ₹22,000 crore ATMP already under construction."),
        (RGBColor(0x5D, 0x6D, 0x7E), "Defence & Aerospace",
         "India's defence exports surged to ₹21,083 crore in FY24 (2x in three years).\n"
         "HAL, BEL, and private OEMs integrating into US & European supply chains post-Ukraine conflict."),
    ]

    for idx, (col, title, detail) in enumerate(data_points):
        ry = 1.3 + idx * 1.3
        add_rect(slide, 0.3, ry, 12.7, 1.12, RGBColor(0xF8, 0xF9, 0xFA),
                 border_color=col, border_pt=2)
        add_rect(slide, 0.3, ry, 0.22, 1.12, col)
        add_textbox(slide, 0.65, ry + 0.08, 12.1, 0.35,
                    title, font_size=13, bold=True, color=col)
        add_textbox(slide, 0.65, ry + 0.46, 12.1, 0.62,
                    detail, font_size=11, color=DARK_GRAY)


def slide_18_verdict(prs):
    """Slide 18 – The Final Verdict"""
    slide = blank_slide(prs)
    set_slide_bg(slide, WHITE)
    title_bar(slide, "The Final Verdict")

    # Question
    add_rect(slide, 0.3, 0.85, 12.7, 0.6, NAVY)
    add_textbox(slide, 0.5, 0.88, 12.3, 0.55,
                "Is Private Investment Still Not Picking Up?",
                font_size=18, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    # Answer
    add_rect(slide, 0.3, 1.55, 12.7, 0.62, GREEN)
    add_textbox(slide, 0.5, 1.58, 12.3, 0.56,
                "NO — It is NOT structurally stagnant. The economy is selectively booming while traditional sectors exercise rational managerial discipline.",
                font_size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    # Two-column summary
    # Green box
    add_rect(slide, 0.25, 2.3, 6.1, 4.3, RGBColor(0xEA, 0xF7, 0xEE))
    add_rect(slide, 0.25, 2.3, 6.1, 0.52, GREEN)
    add_textbox(slide, 0.35, 2.32, 5.9, 0.48,
                "🟢  'New Economy' IS BOOMING", font_size=14, bold=True, color=WHITE)
    new_verdict = [
        "Sectors: High-tech, EV, Renewables, Semiconductors",
        "",
        "MEC Range: 15%–25%  →  clears 3.55% hurdle easily",
        "",
        "PLI scheme acts as government risk-absorber,",
        "making marginal projects commercially viable.",
        "",
        "FDI + China+1 supply chain integration providing",
        "additional demand certainty.",
        "",
        "Greenfield investments proceeding at full pace.",
    ]
    add_textbox(slide, 0.4, 2.92, 5.8, 3.5,
                "\n".join(new_verdict), font_size=11, color=DARK_GRAY)

    # Red box
    add_rect(slide, 6.95, 2.3, 6.1, 4.3, RGBColor(0xFD, 0xED, 0xEC))
    add_rect(slide, 6.95, 2.3, 6.1, 0.52, RED)
    add_textbox(slide, 7.05, 2.32, 5.9, 0.48,
                "🔴  'Old Economy' IS WAITING", font_size=14, bold=True, color=WHITE)
    old_verdict = [
        "Sectors: FMCG, Steel, Textiles, Core Infra",
        "",
        "MEC Range: 8%–10%  →  borderline viable",
        "",
        "Trapped by 74.3% capacity utilization —",
        "below the 78–80% trigger threshold.",
        "",
        "Broken domestic multiplier:",
        "  • Household debt at 41.3% of GDP",
        "  • Rural wage growth subdued",
        "",
        "Rational deferral — not corporate failure.",
    ]
    add_textbox(slide, 7.1, 2.92, 5.8, 3.5,
                "\n".join(old_verdict), font_size=11, color=DARK_GRAY)


def slide_19_triggers(prs):
    """Slide 19 – Triggers & Corporate Action Plan (FY26-27)"""
    slide = blank_slide(prs)
    set_slide_bg(slide, WHITE)
    title_bar(slide, "Triggers & Corporate Action Plan (FY26-27)")

    add_textbox(slide, 0.3, 0.82, 12.7, 0.4,
                "🚀  The Launchpad:  Corporate India is fueled, positioned, and awaiting the final macroeconomic demand signal.",
                font_size=12, bold=True, color=NAVY)

    # Final Triggers
    add_rect(slide, 0.3, 1.32, 6.1, 3.8, RGBColor(0xEB, 0xF5, 0xFB),
             border_color=LIGHT_NAVY, border_pt=1)
    add_rect(slide, 0.3, 1.32, 6.1, 0.44, LIGHT_NAVY)
    add_textbox(slide, 0.45, 1.34, 5.8, 0.4,
                "The Final Triggers", font_size=13, bold=True, color=WHITE)

    trigger_lines = [
        "Trigger 1 — Demand:",
        "  Aggregate demand must push capacity utilization",
        "  consistently past the 80% SRAS boundary.",
        "  (Current: 74.3%  →  Gap: ~6 percentage points)",
        "",
        "Trigger 2 — Interest Rates:",
        "  Real interest rates must normalize to make MSME",
        "  borrowing commercially viable.",
        "  (Current: 3.55%  →  Target: sub-2.5% real)",
    ]
    add_textbox(slide, 0.45, 1.85, 5.8, 3.15,
                "\n".join(trigger_lines), font_size=11, color=DARK_GRAY)

    # Action Plan
    add_rect(slide, 6.75, 1.32, 6.3, 3.8, RGBColor(0xEA, 0xF7, 0xEE),
             border_color=GREEN, border_pt=1)
    add_rect(slide, 6.75, 1.32, 6.3, 0.44, GREEN)
    add_textbox(slide, 6.9, 1.34, 6.0, 0.4,
                "Action Plan for Corporate Planners", font_size=13,
                bold=True, color=WHITE)

    action_lines = [
        "Action 1 — Stop evaluating purely for domestic",
        "  consumption demand.",
        "  Highest ROIs are currently in integrating domestic",
        "  plants into China+1 global supply chains.",
        "",
        "Action 2 — Prepare for the Turn.",
        "  Corporate balance sheets at a decadal best.",
        "  Firms possess the capital — they are simply stationed",
        "  on the launchpad awaiting the final",
        "  macroeconomic demand signal.",
    ]
    add_textbox(slide, 6.9, 1.85, 6.0, 3.15,
                "\n".join(action_lines), font_size=11, color=DARK_GRAY)

    # Bottom conclusion bar
    add_rect(slide, 0.3, 5.28, 12.7, 1.0, NAVY)
    add_textbox(
        slide, 0.5, 5.33, 12.3, 0.9,
        "Conclusion:  India's investment story is not broken — it is bifurcated.\n"
        "The macroeconomic triggers to unlock broad-based private capex are identifiable, measurable, and closer than the headlines suggest.",
        font_size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER
    )


# ─────────────────────────────────────────────────────────────────────────────
# Main entry point
# ─────────────────────────────────────────────────────────────────────────────

def build_presentation(output_path="India_Private_Investment_Puzzle.pptx"):
    prs = Presentation()

    # Set slide dimensions to 16:9 widescreen
    prs.slide_width = W
    prs.slide_height = H

    slide_builders = [
        slide_01_title,
        slide_02_paradox,
        slide_03_two_economies,
        slide_04_theory,
        slide_05_corporate_reality,
        slide_06_multiplier,
        slide_07_demand_shock,
        slide_08_cu_theory,
        slide_09_cu_corporate,
        slide_10_pli,
        slide_11_sectoral,
        slide_12_balance_sheets,
        slide_13_private_credit,
        slide_14_energy,
        slide_15_green_infra,
        slide_16_fdi,
        slide_17_china_plus_one,
        slide_18_verdict,
        slide_19_triggers,
    ]

    print(f"Building presentation with {len(slide_builders)} slides...")
    for i, builder in enumerate(slide_builders, 1):
        builder(prs)
        print(f"  ✓ Slide {i:02d} — {builder.__doc__.strip().split(chr(10))[0]}")

    prs.save(output_path)
    print(f"\n✅  Saved: {output_path}")
    return output_path


if __name__ == "__main__":
    build_presentation()
