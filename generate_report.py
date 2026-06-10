from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

doc = Document()

# ── Page margins ──────────────────────────────────────────────────
for section in doc.sections:
    section.top_margin    = Cm(2.5)
    section.bottom_margin = Cm(2.5)
    section.left_margin   = Cm(3.0)
    section.right_margin  = Cm(2.5)

# ── Colour palette ────────────────────────────────────────────────
BLACK   = RGBColor(0x11, 0x11, 0x11)
GOLD    = RGBColor(0xE8, 0xC5, 0x47)
MID     = RGBColor(0x44, 0x44, 0x44)
LIGHT   = RGBColor(0x88, 0x88, 0x88)
RED     = RGBColor(0xC0, 0x39, 0x2B)
GREEN   = RGBColor(0x1A, 0x7A, 0x3C)
AMBER   = RGBColor(0x99, 0x6B, 0x00)
RULE    = RGBColor(0xCC, 0xCC, 0xCC)
TBL_HDR = RGBColor(0x1A, 0x1A, 0x1A)
TBL_ALT = RGBColor(0xF7, 0xF7, 0xF7)

# ── Helper: paragraph shading ─────────────────────────────────────
def shade_cell(cell, hex_str):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement('w:shd')
    shd.set(qn('w:val'),   'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'),  hex_str)
    tcPr.append(shd)

def set_cell_border(cell, **kwargs):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        if side in kwargs:
            tag = OxmlElement(f'w:{side}')
            tag.set(qn('w:val'),   kwargs[side].get('val',   'single'))
            tag.set(qn('w:sz'),    kwargs[side].get('sz',    '4'))
            tag.set(qn('w:space'), kwargs[side].get('space', '0'))
            tag.set(qn('w:color'), kwargs[side].get('color', 'auto'))
            tcBorders.append(tag)
    tcPr.append(tcBorders)

def add_run(para, text, bold=False, italic=False, size=11, color=None, underline=False):
    run = para.add_run(text)
    run.bold      = bold
    run.italic    = italic
    run.underline = underline
    run.font.size = Pt(size)
    run.font.color.rgb = color or BLACK
    return run

def add_heading(text, level=1, space_before=18, space_after=8):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after  = Pt(space_after)
    if level == 1:
        run = p.add_run(text.upper())
        run.bold = True
        run.font.size  = Pt(10)
        run.font.color.rgb = LIGHT
        run.font.all_caps  = True
        # rule below
        pPr  = p._p.get_or_add_pPr()
        pBdr = OxmlElement('w:pBdr')
        bot  = OxmlElement('w:bottom')
        bot.set(qn('w:val'),   'single')
        bot.set(qn('w:sz'),    '4')
        bot.set(qn('w:space'), '4')
        bot.set(qn('w:color'), 'CCCCCC')
        pBdr.append(bot)
        pPr.append(pBdr)
    elif level == 2:
        run = p.add_run(text)
        run.bold = True
        run.font.size  = Pt(13)
        run.font.color.rgb = BLACK
    elif level == 3:
        run = p.add_run(text)
        run.bold = True
        run.font.size  = Pt(11)
        run.font.color.rgb = MID
    return p

def add_body(text, space_after=6):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    r = p.add_run(text)
    r.font.size = Pt(10.5)
    r.font.color.rgb = MID
    return p

def add_bullet(text, bold_prefix=None):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.space_after  = Pt(3)
    p.paragraph_format.left_indent  = Cm(0.6)
    if bold_prefix:
        rb = p.add_run(bold_prefix)
        rb.bold = True
        rb.font.size = Pt(10.5)
        rb.font.color.rgb = BLACK
        r = p.add_run(text)
    else:
        r = p.add_run(text)
    r.font.size = Pt(10.5)
    r.font.color.rgb = MID
    return p

def add_callout(text, label=None, color_hex='F7F7F0', left_color='111111'):
    """A shaded callout paragraph."""
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
    cell = tbl.cell(0, 0)
    shade_cell(cell, color_hex)
    cell.width = Inches(5.8)
    # Left border accent
    set_cell_border(cell,
        left  ={'val':'single','sz':'16','color':left_color},
        top   ={'val':'none'},
        right ={'val':'none'},
        bottom={'val':'none'},
    )
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after  = Pt(6)
    if label:
        rb = p.add_run(label + '  ')
        rb.bold = True
        rb.font.size = Pt(10)
        rb.font.color.rgb = BLACK
    r = p.add_run(text)
    r.font.size = Pt(10)
    r.font.color.rgb = MID
    doc.add_paragraph().paragraph_format.space_after = Pt(4)
    return tbl

def simple_table(headers, rows, col_widths=None):
    tbl = doc.add_table(rows=1 + len(rows), cols=len(headers))
    tbl.style = 'Table Grid'
    tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
    # Header row
    hrow = tbl.rows[0]
    for i, h in enumerate(headers):
        cell = hrow.cells[i]
        shade_cell(cell, '111111')
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after  = Pt(4)
        r = p.add_run(h.upper())
        r.bold = True
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    # Data rows
    for ri, row in enumerate(rows):
        drow = tbl.rows[ri + 1]
        fill = 'FFFFFF' if ri % 2 == 0 else 'F7F7F7'
        for ci, val in enumerate(row):
            cell = drow.cells[ci]
            shade_cell(cell, fill)
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(3)
            p.paragraph_format.space_after  = Pt(3)
            if isinstance(val, tuple):
                text, bold, color = val
                r = p.add_run(text)
                r.bold = bold
                r.font.size = Pt(9.5)
                r.font.color.rgb = color or MID
            else:
                r = p.add_run(str(val))
                r.font.size = Pt(9.5)
                r.font.color.rgb = MID
    if col_widths:
        for i, w in enumerate(col_widths):
            for row in tbl.rows:
                row.cells[i].width = Cm(w)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return tbl

# ════════════════════════════════════════════════════════════════════
# COVER PAGE
# ════════════════════════════════════════════════════════════════════
p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(60)
p.paragraph_format.space_after  = Pt(4)
r = p.add_run('ONE LDN')
r.bold = True
r.font.size = Pt(11)
r.font.color.rgb = LIGHT

p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(8)
r = p.add_run('Shop Sales Dashboard')
r.bold = True
r.font.size = Pt(28)
r.font.color.rgb = BLACK

p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(4)
r = p.add_run('Redesign Planning Report')
r.font.size = Pt(16)
r.font.color.rgb = MID

# Gold rule
p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(20)
p.paragraph_format.space_after  = Pt(20)
pPr  = p._p.get_or_add_pPr()
pBdr = OxmlElement('w:pBdr')
bot  = OxmlElement('w:bottom')
bot.set(qn('w:val'),   'single')
bot.set(qn('w:sz'),    '12')
bot.set(qn('w:space'), '0')
bot.set(qn('w:color'), 'E8C547')
pBdr.append(bot)
pPr.append(pBdr)

meta = [
    ('Date',         'June 2026'),
    ('Version',      '1.0 — Draft for review'),
    ('Project',      'shop-sales · github.com/ONE-LDN/shop-sales'),
    ('Prepared for', 'ONE LDN Management'),
]
for label, val in meta:
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(3)
    rb = p.add_run(label + ':  ')
    rb.bold = True
    rb.font.size = Pt(10)
    rb.font.color.rgb = MID
    rv = p.add_run(val)
    rv.font.size = Pt(10)
    rv.font.color.rgb = LIGHT

doc.add_page_break()

# ════════════════════════════════════════════════════════════════════
# 1. EXECUTIVE SUMMARY
# ════════════════════════════════════════════════════════════════════
add_heading('1. Executive Summary', level=1, space_before=0)

add_callout(
    'The ONE LDN shop sales dashboard will provide a week-on-week view of stock levels and '
    'sales, serving as the central management system and reporting tool for all shop sales activity.',
    color_hex='F7F7F0', left_color='E8C547'
)

add_body(
    'ONE LDN operates a café and a merchandise shop within the gym. Sales are processed through '
    'WodBoard; stock arrives via supplier deliveries logged by PDF invoice or manual entry. '
    'This report documents the planning and design of a dashboard that gives management a single, '
    'reliable place to answer every operational question about the shop — without duplicating '
    'information or requiring data to be tracked in more than one place.'
)

add_heading('What We Measure', level=3, space_before=14, space_after=4)
add_body(
    'Three inputs drive everything in the dashboard:'
)
simple_table(
    ['Input', 'Source', 'What it enables'],
    [
        ('Total sales',      'WodBoard CSV — uploaded weekly',            'What sold, revenue, velocity, reorder triggers'),
        ('Total deliveries', 'Supplier invoices (PDF or manual entry)',   'Stock received, COGS, reconciliation opening balance'),
        ('Total stock',      'Physical count entered in the dashboard',   'Actual vs theoretical variance, accurate reorder baseline'),
    ],
    col_widths=[3.5, 5.5, 7.0]
)

add_heading('What the Dashboard Outputs', level=3, space_before=14, space_after=4)
add_body(
    'From those three inputs, the dashboard produces four outputs — one per tab — with no overlap '
    'between them. Each tab answers exactly one question:'
)
simple_table(
    ['Tab', 'Question', 'Outcome for the business'],
    [
        ('Weekly Report',        'What sold?',               'Week-on-week sales by product and category; stock health at a glance'),
        ('Stock Reconciliation', 'What do we have?',         'Physical stock vs theoretical; variance tracking; accurate reorder baseline'),
        ('Order Prediction',     'What needs reordering?',   'Velocity-driven days-of-stock and suggested order quantities per product'),
        ('Finance',              'How are we performing?',   'Revenue trends over time; gross profit % by product and month'),
    ],
    col_widths=[4.0, 4.5, 7.5]
)

add_callout(
    'The connection between tabs is deliberate and sequential: what sold (Weekly Report) '
    'directly drives what needs reordering (Order Prediction) via sales velocity. '
    'What we have (Stock Reconciliation) provides the physical count that makes the '
    'reorder calculation accurate. Finance draws on all three to give the full picture.',
    color_hex='F0F7F0', left_color='1A7A3C'
)

add_heading('What This Replaced', level=3, space_before=14, space_after=4)
add_body(
    'The previous dashboard had five tabs built incrementally. Over time it accumulated '
    'overlapping views, a default screen that didn\'t reflect daily use, and no physical '
    'stock count workflow — despite the required database infrastructure already existing. '
    'The redesign retires three of the five tabs and consolidates their content with no '
    'loss of functionality.'
)

doc.add_paragraph()

# ════════════════════════════════════════════════════════════════════
# 2. BACKGROUND
# ════════════════════════════════════════════════════════════════════
add_heading('2. Background', level=1)

add_heading('The Business', level=2, space_before=10, space_after=6)
add_body(
    'ONE LDN\'s shop operates across two distinct product areas with different inventory '
    'characteristics:'
)
add_bullet('Café — coffees, protein shakes, energy drinks, bars and snacks. High-frequency, '
           'ordered on a two-week cycle from suppliers including Muscle Finesse. Products turn '
           'over weekly; stock runs out quickly if not monitored.')
add_bullet('Merchandise — branded apparel and accessories. Longer lead times, ordered on an '
           'eight-week cycle from suppliers including Tapstitch. Lower velocity but higher '
           'unit value; over-ordering ties up cash.')
add_body(
    'Sales flow from WodBoard into the dashboard via weekly CSV upload. Deliveries are logged '
    'by uploading supplier PDF invoices (parsed automatically) or by manual entry. All data '
    'is stored in Supabase.'
)

add_heading('Problems with the Previous Design', level=2, space_before=10, space_after=6)
add_body(
    'The original dashboard grew tab by tab as requirements were added. By the time the '
    'redesign was initiated it had five tabs, several of which served overlapping purposes:'
)
problems = [
    'Overview and Top Sellers both showed sales rankings in different formats — the same question answered twice.',
    'Stock History combined delivery reconciliation with gross profit reporting, serving two different audiences in one view.',
    'The default Overview tab showed all-time data rather than the current week, making it less useful for the most common use case.',
    'No physical stock count workflow existed, even though a Supabase table (shop_stock_takes) had been created for exactly this purpose.',
    'Stock urgency information (days of stock remaining, Order Now flags) required navigating to the Order Prediction tab — it was not visible from the default screen.',
]
for p in problems:
    add_bullet(p)

doc.add_paragraph()

# ════════════════════════════════════════════════════════════════════
# 3. WHAT WE MEASURE AND WHY
# ════════════════════════════════════════════════════════════════════
add_heading('3. What We Measure and Why', level=1)

add_body(
    'The dashboard is built around three tracked quantities and the four outcomes they produce. '
    'The design principle is that each quantity is measured once, in the right place, with no '
    'duplication across tabs.'
)

add_heading('The Three Inputs', level=2, space_before=10, space_after=6)

add_heading('Total Sales', level=3, space_before=10, space_after=4)
add_body(
    'Every transaction is captured in the WodBoard CSV: product, quantity, revenue, date. '
    'This is the primary data source. It feeds the Weekly Report (what sold), drives the '
    'velocity calculation used by Order Prediction (how fast is it selling), and contributes '
    'the "sold" column in Stock Reconciliation.'
)

add_heading('Total Deliveries', level=3, space_before=10, space_after=4)
add_body(
    'Every delivery is logged with product, quantity received, date, and unit price. This feeds '
    'the Stock Reconciliation "received" column, provides the cost base for gross profit '
    'calculation in Finance, and resets the estimated stock baseline used by Order Prediction.'
)

add_heading('Total Stock', level=3, space_before=10, space_after=4)
add_body(
    'The physical count entered in Stock Reconciliation. Without this, the system can only '
    'estimate stock from deliveries minus sales — which accumulates error over time. The count '
    'grounds the system in reality and makes Order Prediction significantly more accurate. '
    'A monthly count cadence is planned rather than weekly.'
)

add_heading('The Four Outputs', level=2, space_before=10, space_after=6)

add_heading('Weekly Report — What Sold', level=3, space_before=10, space_after=4)
add_body(
    'A week-on-week view of sales by product and category. The primary screen for anyone '
    'opening the dashboard. Shows units sold, revenue, and the percentage change versus the '
    'previous week for each category. Also surfaces days of stock remaining and a stock health '
    'summary so that urgent reorder situations are visible without navigating away.'
)
add_body('Measured: units sold, revenue per product and category, week-on-week change, days of stock.')

add_heading('Stock Reconciliation — What We Have', level=3, space_before=10, space_after=4)
add_body(
    'The bridge between the theoretical stock position (deliveries minus sales) and the physical '
    'reality. Each week shows the opening balance, units received, units sold, the theoretical '
    'closing stock, and — when a physical count has been entered — the variance. '
    'A monthly variance summary aggregates shrinkage or surplus across all counted weeks.'
)
add_body('Measured: opening balance, deliveries in, sold out, theoretical stock, physical count, variance.')

add_heading('Order Prediction — What Needs Reordering', level=3, space_before=10, space_after=4)
add_body(
    'A velocity-driven reorder view. For each product, the system calculates average weekly '
    'sales over the last four weeks, divides the estimated stock by that velocity to arrive '
    'at days of stock remaining, and compares it to the order cycle threshold. Products '
    'falling below the threshold are flagged Order Now or Order Soon. A suggested order '
    'quantity covers the next full cycle.'
)
add_body(
    'This tab is informed by both sales data (velocity) and reconciliation data (stock on hand). '
    'The more accurate the physical count, the more reliable the reorder suggestion.'
)
add_body('Measured: velocity (avg units/week), estimated stock, days of cover, suggested order quantity, urgency status.')

add_heading('Finance — Trends and Margin', level=3, space_before=10, space_after=4)
add_body(
    'A management-level view for monthly review. Shows revenue by category over time (all '
    'historical weeks) and gross profit percentage by product for a selected month. This is '
    'the only tab concerned with cost and margin — GP% is not shown elsewhere, avoiding any '
    'risk of it being used operationally before cost prices are fully populated.'
)
add_body('Measured: weekly revenue by category (trend), revenue, COGS, and GP% by product (monthly).')

doc.add_paragraph()

# ════════════════════════════════════════════════════════════════════
# 4. THE PREVIOUS DESIGN — FIVE-TAB STRUCTURE
# ════════════════════════════════════════════════════════════════════
add_heading('4. The Previous Design — Five-Tab Structure', level=1)

simple_table(
    ['Tab', 'Purpose', 'Issues'],
    [
        ('Overview',
         'KPI cards, weekly revenue bar chart, category donut, top-10 product table',
         'Mixed all-time and current-period data; overlapped with Top Sellers; stock alerts '
         'placed here rather than in Order Prediction'),
        ('Top Sellers',
         'Expandable brand and collection hierarchy tables for Café and Merch',
         'Same question as Overview ("what\'s selling?") in a different format — required '
         'visiting two tabs for the same information'),
        ('Trends',
         'Stacked line charts of weekly revenue and units by category',
         'Useful content isolated on its own tab; not integrated with financial review'),
        ('Order Prediction',
         'Velocity-based reorder table with covers bar and urgency flags',
         'Strong and correctly isolated — no structural issues; tab position was wrong '
         '(should follow reconciliation, not precede it)'),
        ('Stock History',
         'Monthly delivery reconciliation plus gross profit % by product',
         'Two audiences in one view: operational (stock flow) and financial (GP%). '
         'No physical count entry despite the database table existing'),
    ],
    col_widths=[3.0, 5.5, 7.5]
)

add_heading('What the Previous Design Did Well', level=2, space_before=10, space_after=6)
strong = [
    'The Order Prediction velocity engine — covers calculation, urgency flags, tab badge — was well-conceived and ahead of most entry-level commercial tools.',
    'PDF invoice parsing for Muscle Finesse and Tapstitch deliveries, reducing manual data entry at the point of goods receipt.',
    'Separate order cycle thresholds for café (two weeks) and merch (eight weeks), correctly modelling the different inventory dynamics of each product type.',
    'A clean product master table (category, brand, cost price, pack size, minimum stock) providing a reliable data foundation.',
    'Correct exclusion of in-house products (house-made shakes, coffee) from stock tracking — the system only predicted stock for items that are actually ordered.',
]
for s in strong:
    add_bullet(s)

doc.add_paragraph()

# ════════════════════════════════════════════════════════════════════
# 5. THE REDESIGN — FOUR-TAB STRUCTURE
# ════════════════════════════════════════════════════════════════════
add_heading('5. The Redesign — Four-Tab Structure', level=1)

add_body(
    'The redesign consolidates five tabs into four, each serving a single purpose with no '
    'overlap. Three tabs are replaced or merged; Order Prediction is carried over unchanged '
    'with a corrected position in the tab sequence.'
)

simple_table(
    ['#', 'Tab', 'Replaces', 'New or changed?'],
    [
        ('1', 'Weekly Report',        'Overview + Top Sellers',             'New — week navigation, health strip, days-left column'),
        ('2', 'Stock Reconciliation', 'Stock History (operational rows)',    'New — physical count entry, monthly variance summary'),
        ('3', 'Order Prediction',     'Order Prediction',                   'Unchanged — position corrected to follow reconciliation'),
        ('4', 'Finance',              'Trends + Stock History (GP% data)',   'Merged — revenue chart and GP table in one monthly view'),
    ],
    col_widths=[0.8, 4.0, 5.2, 6.0]
)

add_heading('Tab 1 — Weekly Report', level=2, space_before=12, space_after=6)
add_body('Default landing screen. Primary audience: anyone opening the dashboard to understand the current week.')
wr_items = [
    'Week navigation (arrows + dropdown) — moves between any week in the dataset.',
    'Stock health strip — four count badges (Order Now / Order Soon / Healthy / Excess), each linking to Order Prediction pre-filtered to that status.',
    'Per-category KPI cards — units, revenue, and week-on-week % change for Coffee, Shakes, Bars & Snacks, Functional Drinks, Merch.',
    'Product table — grouped by brand, expandable to product level, sortable by units or revenue. Columns: Product, Category, Units Sold, Revenue, Days Left, Current Stock.',
    'Category filter (All Café, sub-categories, All Merch).',
]
for item in wr_items:
    add_bullet(item)

add_heading('Tab 2 — Stock Reconciliation', level=2, space_before=12, space_after=6)
add_body('Physical stock reality check. Intended for monthly use, not weekly. Covers bought-in products only.')
sr_items = [
    'Week navigation to review any historical week.',
    'Per-product flow table: Opening Balance → + Deliveries → − Sold → = Theoretical → Closing Count → Variance.',
    'Opening balance chains automatically from the previous week\'s entered count.',
    'Closing count entry — click any cell to open a popover input with an optional notes field.',
    'Monthly variance summary — aggregates all weeks with entered counts; shows total variance per product and month total.',
]
for item in sr_items:
    add_bullet(item)
add_callout(
    'The database table for physical counts (shop_stock_takes) already existed in Supabase '
    'with zero rows. This tab connects the UI to that table for the first time.',
    color_hex='F0F7F0', left_color='1A7A3C'
)

add_heading('Tab 3 — Order Prediction', level=2, space_before=12, space_after=6)
add_body('Velocity-driven reorder decisions. Unchanged from the previous design; position corrected to follow Stock Reconciliation.')
op_items = [
    'Velocity = average units sold per week over the last four weeks.',
    'Estimated stock = last delivery quantity minus units sold since that delivery.',
    'Days of cover = estimated stock ÷ weekly velocity × 7.',
    'Urgency flags: Order Now (café < 7 days, merch < 14 days), Order Soon, Sufficient.',
    'Suggested order quantity covering the next full cycle (2 weeks for café, 8 weeks for merch).',
    'Live tab badge showing the current Order Now count.',
]
for item in op_items:
    add_bullet(item)

add_heading('Tab 4 — Finance', level=2, space_before=12, space_after=6)
add_body('Monthly management review. Revenue trend and gross profit — not for daily operational use.')
fi_items = [
    'Weekly revenue by category — line chart, all historical weeks, per-category toggle.',
    'Gross profit table by product — month picker, columns: Product, Units Sold, Revenue, COGS, GP%.',
    'Month-total GP% in the table footer.',
]
for item in fi_items:
    add_bullet(item)

doc.add_paragraph()

# ════════════════════════════════════════════════════════════════════
# 6. PROTOTYPE ITERATIONS
# ════════════════════════════════════════════════════════════════════
add_heading('6. Prototype Iterations', level=1)

add_body(
    'Two prototype files were produced to validate layout and interactions before committing '
    'to building against live data. Both use static mock data.'
)

add_heading('Iteration 1 — Early Layout', level=2, space_before=10, space_after=6)
add_body(
    'Established the visual language: dark theme matching the ONE LDN aesthetic, category '
    'colour coding (red for coffee, gold for shakes, green for functional drinks, blue for '
    'merch), KPI card format, and the covers bar in Order Prediction. No tab structure was '
    'finalised at this stage.'
)

add_heading('Iteration 2 — Structure Prototype', level=2, space_before=10, space_after=6)
add_body(
    'The four-tab structure prototype, built to validate layout and key interactions. '
    'All core flows were made interactive using static JavaScript:'
)
iter2_items = [
    'Week-by-week navigation with prev/next arrows and a dropdown.',
    'Brand-expandable product table with sort and category filter controls.',
    'Stock reconciliation table with popover count-entry: click a cell, enter a quantity, add a note, save — triggers a flash confirmation and chains the count to the following week\'s opening balance.',
    'Finance tab: line chart with per-category legend toggles; GP table with month picker.',
    'Order Prediction: covers bar, urgency flags, status filter, Order button per product row.',
]
for item in iter2_items:
    add_bullet(item)

add_heading('Audit and Revisions', level=2, space_before=10, space_after=6)
add_body(
    'The structure prototype was formally audited against the benchmark metrics and the '
    'stated objective. Four high-priority issues were identified and resolved before the '
    'prototype was signed off.'
)
hi_items = [
    ('Tab order transposed. ',
     'Stock Reconciliation was at position 3, Order Prediction at position 2 — reversing '
     'the intended flow (stock reality should inform reorder decisions). Corrected.'),
    ('Days Left not visible on the default screen. ',
     'Days of stock remaining was only shown inside Order Prediction. Added as a colour-coded '
     'column in the Weekly Report product table so it is visible without navigating away.'),
    ('No stock health summary on the primary view. ',
     'Order Now / Order Soon counts required a tab switch. Added as a four-badge health strip '
     'above the KPI cards, with each badge linking directly to Order Prediction filtered to '
     'that status.'),
    ('Monthly variance summary missing. ',
     'Stock Reconciliation showed variance per week but not aggregated by month. Added a '
     'summary table below the weekly view, covering all months with entered counts.'),
]
for bold, rest in hi_items:
    add_bullet(rest, bold_prefix=bold)

add_body('Items identified but deferred to the next iteration:')
med_items = [
    'Sell-through rate (units sold ÷ delivered + opening stock) — not yet calculated.',
    'Overstocked / excess flag in Order Prediction for products above cycle-length thresholds.',
    'Total-week revenue headline above the category KPI cards.',
    'Merch variant grouping (size, colourway) in the product table.',
]
for item in med_items:
    add_bullet(item)

doc.add_paragraph()

# ════════════════════════════════════════════════════════════════════
# 7. KEY DESIGN DECISIONS
# ════════════════════════════════════════════════════════════════════
add_heading('7. Key Design Decisions', level=1)

decisions = [
    (
        'Week-by-week navigation, not a rolling-period selector',
        'The original dashboard had a rolling-period control (4 / 8 / 12 weeks, or a month). '
        'The redesign replaces this on the primary tab with direct week navigation. '
        'The most common operational question is about a specific week, not an average period. '
        'Rolling-period analysis belongs on the Finance tab.'
    ),
    (
        'Stock health as count badges, not a chart',
        'A bar chart was considered for the stock health panel. Badges were chosen because '
        'stock health status is categorical — a product is either low, healthy, or excess — '
        'not a continuous value that benefits from a chart axis. The badges also serve as '
        'direct navigation shortcuts to the filtered Order Prediction view.'
    ),
    (
        'Monthly physical count cadence',
        'Daily or weekly counts are not practical for this business. The design accommodates '
        'this by allowing counts to be entered for any week at any time, with the monthly '
        'variance summary only aggregating weeks where a count was actually recorded. '
        'The theoretical stock estimate remains available in all other weeks.'
    ),
    (
        'Order Prediction carried over unchanged',
        'The velocity and reorder engine was the strongest part of the original design — '
        'correctly modelled, well-executed, and already trusted by the team. The only '
        'change was correcting its position to tab 3 so that Stock Reconciliation (the '
        'stock reality check) precedes it logically.'
    ),
    (
        'Café and merch as segment toggles, not separate tabs',
        'Earlier iterations had separate tabs for each product area. The redesign uses '
        'a toggle within each tab. The overall stock health picture matters first; '
        'the café/merch breakdown is a filter applied on top of it.'
    ),
    (
        'In-house products excluded from stock tracking',
        'Coffee and house-made shakes are not purchased stock — they are prepared in-house '
        'from consumables. Tracking them as inventory units would produce meaningless reorder '
        'suggestions. The exclusion logic from Order Prediction is applied consistently '
        'to Stock Reconciliation.'
    ),
]
for title, body in decisions:
    add_heading(title, level=3, space_before=10, space_after=4)
    add_body(body)

doc.add_paragraph()

# ════════════════════════════════════════════════════════════════════
# 8. NEXT STEPS
# ════════════════════════════════════════════════════════════════════
add_heading('8. Next Steps', level=1)

add_heading('Prototype to Production', level=2, space_before=10, space_after=6)
immediate = [
    'Implement the four-tab structure in the live dashboard, migrating existing logic into the new tab functions.',
    'Connect Stock Reconciliation to the live shop_stock_takes Supabase table.',
    'Populate cost prices in the product lookup table to unlock gross profit reporting on the Finance tab.',
    'Validate the reconciliation opening-balance chain against real delivery and sales data.',
]
for item in immediate:
    add_bullet(item)

add_heading('Short-Term', level=2, space_before=10, space_after=6)
short = [
    'Sell-through rate column in the Weekly Report product table.',
    'Overstocked / excess flag in Order Prediction for products above cycle-length thresholds.',
    'Total-week revenue headline above the category KPI cards.',
    'Merch variant grouping (size, colourway) in the Weekly Report product table.',
]
for item in short:
    add_bullet(item)

add_heading('Longer Term', level=2, space_before=10, space_after=6)
longer = [
    'Gross margin headline KPI on the Finance tab.',
    'Slow-movers callout: products with zero sales this week but positive estimated stock.',
    'Performance: parallelise Supabase data loading on start-up (currently 9+ sequential calls).',
]
for item in longer:
    add_bullet(item)

doc.add_paragraph()

# ════════════════════════════════════════════════════════════════════
# FOOTER
# ════════════════════════════════════════════════════════════════════
p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(24)
pPr  = p._p.get_or_add_pPr()
pBdr = OxmlElement('w:pBdr')
top  = OxmlElement('w:top')
top.set(qn('w:val'),   'single')
top.set(qn('w:sz'),    '4')
top.set(qn('w:space'), '4')
top.set(qn('w:color'), 'CCCCCC')
pBdr.append(top)
pPr.append(pBdr)
r = p.add_run('ONE LDN Shop Dashboard — Redesign Planning Report   ·   June 2026   ·   Internal document')
r.font.size = Pt(9)
r.font.color.rgb = LIGHT

# ── Save ─────────────────────────────────────────────────────────
doc.save('/home/user/shop-sales/shop-dashboard-redesign-report.docx')
print("Done: shop-dashboard-redesign-report.docx")
