from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

doc = Document()

# ── Page margins ──────────────────────────────────────────────────────────────
for section in doc.sections:
    section.top_margin    = Cm(2)
    section.bottom_margin = Cm(2)
    section.left_margin   = Cm(2.5)
    section.right_margin  = Cm(2.5)

# ── Colour palette ────────────────────────────────────────────────────────────
NAVY  = RGBColor(0x0A, 0x16, 0x28)
GOLD  = RGBColor(0xC9, 0xA8, 0x4C)
RED   = RGBColor(0xC0, 0x39, 0x2B)
GREEN = RGBColor(0x27, 0xAE, 0x60)
BLUE  = RGBColor(0x1E, 0x40, 0xAF)
GREY  = RGBColor(0xF4, 0xF6, 0xF9)
MUTED = RGBColor(0x64, 0x74, 0x8B)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

# ── Helpers ───────────────────────────────────────────────────────────────────
def set_cell_bg(cell, hex_color):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement('w:shd')
    shd.set(qn('w:val'),   'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'),  hex_color)
    tcPr.append(shd)

def set_row_height(row, height_cm):
    tr   = row._tr
    trPr = tr.get_or_add_trPr()
    trH  = OxmlElement('w:trHeight')
    trH.set(qn('w:val'),  str(int(height_cm * 567)))
    trH.set(qn('w:hRule'), 'atLeast')
    trPr.append(trH)

def cell_text(cell, text, bold=False, size=10, color=None, align=WD_ALIGN_PARAGRAPH.LEFT, italic=False):
    cell.text = ''
    p  = cell.paragraphs[0]
    p.alignment = align
    run = p.add_run(text)
    run.bold   = bold
    run.italic = italic
    run.font.size = Pt(size)
    if color:
        run.font.color.rgb = color

def add_heading(doc, text, level=1, color=NAVY):
    p = doc.add_paragraph()
    if level == 0:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(text)
        run.bold = True
        run.font.size = Pt(24)
        run.font.color.rgb = WHITE
        # navy background via table trick — just style the run
        run.font.color.rgb = GOLD
    elif level == 1:
        run = p.add_run(text)
        run.bold = True
        run.font.size = Pt(18)
        run.font.color.rgb = color
        p.paragraph_format.space_before = Pt(18)
        p.paragraph_format.space_after  = Pt(6)
    elif level == 2:
        run = p.add_run(text)
        run.bold = True
        run.font.size = Pt(14)
        run.font.color.rgb = color
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after  = Pt(4)
    elif level == 3:
        run = p.add_run(text)
        run.bold = True
        run.font.size = Pt(11)
        run.font.color.rgb = color
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after  = Pt(2)
    return p

def add_body(doc, text, size=10, color=None, bold=False, italic=False, indent=False):
    p   = doc.add_paragraph()
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.bold   = bold
    run.italic = italic
    if color:
        run.font.color.rgb = color
    p.paragraph_format.space_after = Pt(4)
    if indent:
        p.paragraph_format.left_indent = Cm(0.6)
    return p

def add_bullet(doc, text, size=10):
    p = doc.add_paragraph(style='List Bullet')
    run = p.add_run(text)
    run.font.size = Pt(size)
    p.paragraph_format.space_after = Pt(2)
    return p

def add_kpi_table(doc, kpis):
    """kpis = list of (label, value, sub) tuples"""
    cols  = min(3, len(kpis))
    rows  = -(-len(kpis) // cols)
    table = doc.add_table(rows=rows, cols=cols)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    idx = 0
    for r in range(rows):
        for c in range(cols):
            if idx >= len(kpis): break
            cell  = table.cell(r, c)
            label, val, sub = kpis[idx]
            set_cell_bg(cell, '0A1628')
            cell.text = ''
            lp  = cell.add_paragraph()
            lr  = lp.add_run(label.upper())
            lr.font.size = Pt(7)
            lr.font.color.rgb = RGBColor(0xC9, 0xA8, 0x4C)
            lr.bold = True
            vp  = cell.add_paragraph()
            vr  = vp.add_run(val)
            vr.font.size = Pt(18)
            vr.font.color.rgb = WHITE
            vr.bold = True
            sp  = cell.add_paragraph()
            sr  = sp.add_run(sub)
            sr.font.size = Pt(8)
            sr.font.color.rgb = RGBColor(0xA0, 0xAE, 0xC0)
            idx += 1
    doc.add_paragraph()

def add_data_table(doc, headers, rows_data, col_widths=None):
    table = doc.add_table(rows=1 + len(rows_data), cols=len(headers))
    table.style = 'Table Grid'
    # header row
    hdr = table.rows[0]
    for i, h in enumerate(headers):
        cell_text(hdr.cells[i], h, bold=True, size=9, color=WHITE)
        set_cell_bg(hdr.cells[i], '0A1628')
    # data rows
    for ri, row_data in enumerate(rows_data):
        row = table.rows[ri + 1]
        bg  = 'F8FAFC' if ri % 2 == 0 else 'FFFFFF'
        for ci, val in enumerate(row_data):
            cell_text(row.cells[ci], str(val), size=9)
            set_cell_bg(row.cells[ci], bg)
    if col_widths:
        for ci, w in enumerate(col_widths):
            for row in table.rows:
                row.cells[ci].width = Cm(w)
    doc.add_paragraph()

def add_pipeline_stage(doc, num, title, timing, description, tags):
    table = doc.add_table(rows=1, cols=2)
    table.style = 'Table Grid'
    # number cell
    nc = table.cell(0, 0)
    nc.width = Cm(1.2)
    set_cell_bg(nc, '0A1628')
    cell_text(nc, str(num), bold=True, size=16, color=GOLD, align=WD_ALIGN_PARAGRAPH.CENTER)
    nc.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    # content cell
    cc = table.cell(0, 1)
    set_cell_bg(cc, 'FAFBFC')
    cc.text = ''
    tp = cc.add_paragraph()
    tr = tp.add_run(title)
    tr.bold = True
    tr.font.size = Pt(11)
    tr.font.color.rgb = NAVY
    if timing:
        tr2 = tp.add_run(f'  {timing}')
        tr2.font.size = Pt(8)
        tr2.font.color.rgb = GOLD
    dp = cc.add_paragraph()
    dr = dp.add_run(description)
    dr.font.size = Pt(9)
    dr.font.color.rgb = MUTED
    if tags:
        tp2 = cc.add_paragraph()
        tr3 = tp2.add_run('  '.join(f'[{t}]' for t in tags))
        tr3.font.size = Pt(8)
        tr3.font.color.rgb = RGBColor(0x5B, 0x21, 0xB6)
        tr3.bold = True
    doc.add_paragraph().paragraph_format.space_after = Pt(2)

def add_term_table(doc, terms):
    """terms = list of (es, en, desc)"""
    table = doc.add_table(rows=len(terms), cols=3)
    table.style = 'Table Grid'
    hdrs = ['Spanish Term', 'English', 'Meaning']
    hr = doc.add_table(rows=1, cols=3)
    hr.style = 'Table Grid'
    for i, h in enumerate(hdrs):
        cell_text(hr.rows[0].cells[i], h, bold=True, size=9, color=WHITE)
        set_cell_bg(hr.rows[0].cells[i], '0A1628')
    for ri, (es, en, desc) in enumerate(terms):
        row = table.rows[ri]
        bg  = 'F0F4FF' if ri % 2 == 0 else 'FFFFFF'
        cell_text(row.cells[0], es, bold=True, size=9, color=NAVY)
        set_cell_bg(row.cells[0], bg)
        cell_text(row.cells[1], en, size=9, italic=True, color=MUTED)
        set_cell_bg(row.cells[1], bg)
        cell_text(row.cells[2], desc, size=9)
        set_cell_bg(row.cells[2], bg)
    doc.add_paragraph()

# ══════════════════════════════════════════════════════════════════════════════
# COVER PAGE
# ══════════════════════════════════════════════════════════════════════════════
cover_tbl = doc.add_table(rows=1, cols=1)
cover_tbl.style = 'Table Grid'
cc = cover_tbl.cell(0, 0)
set_cell_bg(cc, '0A1628')
cc.text = ''
p = cc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run('HUSPY')
r.bold = True; r.font.size = Pt(11); r.font.color.rgb = GOLD

p2 = cc.add_paragraph()
p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
r2 = p2.add_run('Spanish Real Estate Market Analysis')
r2.bold = True; r2.font.size = Pt(26); r2.font.color.rgb = WHITE

p3 = cc.add_paragraph()
p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
r3 = p3.add_run('Málaga Operations Director — Market Intelligence Report')
r3.font.size = Pt(13); r3.font.color.rgb = RGBColor(0xA0, 0xAE, 0xC0)

p4 = cc.add_paragraph()
p4.alignment = WD_ALIGN_PARAGRAPH.CENTER
r4 = p4.add_run('March 2026  ·  Spain National + Málaga Province + Costa del Sol')
r4.font.size = Pt(9); r4.font.color.rgb = RGBColor(0x64, 0x74, 0x8B)

doc.add_paragraph()

# ══════════════════════════════════════════════════════════════════════════════
# SECTION A — DATA ANALYSIS
# ══════════════════════════════════════════════════════════════════════════════
add_heading(doc, 'SECTION A — DATA ANALYSIS', 1)
add_body(doc, 'Key performance indicators derived from INE, Registro de la Propiedad, Banco de España, Tinsa, and Idealista Research through Q1 2026.', italic=True, color=MUTED)

add_heading(doc, 'A1. National KPIs — Spain 2025', 2)
add_kpi_table(doc, [
    ('Transactions 2025', '620,000', '+8% vs 2024'),
    ('Avg. Price/m² National', '€2,311', '+11.2% YoY'),
    ('Avg. Price/m² Málaga', '€3,180', '#1 in Andalusia'),
    ('Foreign Buyer Share', '15.1%', '93,000+ transactions'),
    ('Mortgage Approvals', '480,000', '+12% YoY'),
    ('Avg. Mortgage Size', '€148,000', '+6% vs 2024'),
])

add_heading(doc, 'A2. 15-Year Market Cycle (2010–2025)', 2)
add_body(doc, 'Spain experienced one of Europe\'s most dramatic boom-bust-recovery cycles. Three distinct phases define the market:')
add_bullet(doc, 'CRISIS (2010–2014): Prices fell 43% from 2007 peak. Mortgage new lending collapsed from €220B (2006) to €35B (2012). Unemployment hit 26.9%.')
add_bullet(doc, 'RECOVERY (2015–2019): Gradual price recovery +2–5%/yr. Foreign buyers became structural demand floor. Tourism-linked coastal markets led the rebound.')
add_bullet(doc, 'ACCELERATION (2020–2025): COVID dip (−2%) quickly reversed. Low ECB rates + pent-up demand + supply shortage drove prices up 40–50% in prime markets. Málaga: +68% in 5 years.')

add_heading(doc, 'A3. Pricing Dynamics: Listing vs. Sale Price', 2)
add_data_table(doc,
    ['Market', 'Avg Listing Price/m²', 'Avg Sale Price/m²', 'Discount', 'Avg Days on Market'],
    [
        ['Spain National', '€2,500', '€2,311', '7.6%', '65 days'],
        ['Málaga Province', '€3,420', '€3,180', '7.0%', '45 days'],
        ['Marbella', '€5,200', '€4,961', '4.6%', '38 days'],
        ['Málaga City', '€3,100', '€2,950', '4.8%', '32 days'],
        ['Costa del Sol (avg)', '€3,600', '€3,350', '6.9%', '42 days'],
    ],
    col_widths=[4, 3.5, 3.5, 2.5, 3.5]
)
add_body(doc, 'Key insight: Málaga\'s discount gap (7.0%) is narrowing vs. the national average as demand exceeds supply. In Marbella and Málaga city, properties in prime locations sell at or above asking price within days of listing.', italic=True, color=BLUE)

add_heading(doc, 'A4. Transaction Volume by Segment', 2)
add_data_table(doc,
    ['Year', 'Total Transactions', 'Resale', 'New Build', 'Málaga Province', 'Foreign Buyers'],
    [
        ['2010', '491,000', '360,000', '131,000', '28,400', '8.2%'],
        ['2015', '354,000', '290,000', '64,000', '22,100', '11.0%'],
        ['2019', '501,000', '415,000', '86,000', '30,200', '12.6%'],
        ['2021', '565,000', '472,000', '93,000', '33,500', '14.1%'],
        ['2023', '574,000', '480,000', '94,000', '35,385', '14.8%'],
        ['2024', '574,000', '481,000', '93,000', '35,400', '15.0%'],
        ['2025', '620,000', '518,000', '102,000', '38,200*', '15.1%'],
    ],
    col_widths=[2, 3.5, 2.5, 2.5, 3.5, 3]
)

add_heading(doc, 'A5. Mortgage Market Data', 2)
add_data_table(doc,
    ['Metric', '2019', '2022', '2024', '2025 Est.'],
    [
        ['New mortgage lending', '€52B', '€50B', '€54B', '€60B'],
        ['Avg. mortgage size', '€139k', '€145k', '€148k', '€152k'],
        ['Fixed rate share', '40%', '70%', '74%', '72%'],
        ['Variable rate share', '60%', '30%', '26%', '28%'],
        ['Euribor 12m (year-end)', '−0.27%', '3.02%', '2.52%', '~2.3%'],
        ['Mortgage broker penetration', '25%', '28%', '30%', '32%'],
    ],
    col_widths=[5, 3, 3, 3, 3]
)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION B — MARKET KNOWLEDGE
# ══════════════════════════════════════════════════════════════════════════════
doc.add_page_break()
add_heading(doc, 'SECTION B — KNOWLEDGE OF THE SPANISH REAL ESTATE MARKET', 1)

add_heading(doc, 'B1. Market Structure: Key Players', 2)
add_data_table(doc,
    ['Player', 'Type', 'Model', 'Agents (Spain)', 'Málaga Presence', 'Notes'],
    [
        ['RE/MAX Spain', 'Franchise', 'Fixed fee + royalty', '3,200+', 'Strong', 'Largest by agent count'],
        ['Engel & Völkers', 'Franchise', 'Commission split', '1,100+', 'Very Strong', 'Luxury leader; fined €22M for misclassification'],
        ['Tecnocasa', 'Franchise', 'High-volume low-ticket', '900+', 'Strong', 'Traditional model'],
        ['Century 21', 'Franchise', 'Commission split', '400+', 'Growing', 'US franchise, expanding'],
        ['Inmobiliarias independientes', 'Independent', 'Variable', '20,000+', 'Dominant', '>60% of market'],
        ['Huspy', 'Tech-enabled', 'Commission + mortgage', '300+', 'Growing', 'Series B; 20× growth 2024'],
        ['Housfy', 'Online/discount', 'Fixed fee', 'N/A', 'Limited', 'Low-cost disruptor'],
        ['Idealista', 'Portal', 'Listing fees', 'N/A', 'Dominant', '#1 portal; 50M+ monthly visits'],
        ['Fotocasa', 'Portal', 'Listing fees', 'N/A', 'Strong', '#2 portal'],
    ],
    col_widths=[3, 3, 3.5, 2.5, 2.5, 4]
)

add_heading(doc, 'B2. Málaga Province Deep-Dive', 2)
add_kpi_table(doc, [
    ('Avg Price/m² Málaga', '€3,180', '+14.2% YoY (2025)'),
    ('Marbella Price/m²', '€4,961', 'Luxury benchmark'),
    ('Transactions 2024', '35,400', '#1 in Andalusia'),
    ('Foreign Buyer Share', '~40%', 'Double the national avg'),
    ('New Build Pipeline', '8,000 units', 'Expected 2025 delivery'),
    ('Rental Yield (gross)', '4.5–6.5%', 'Vs 3.5–4.5% national'),
])

add_data_table(doc,
    ['Sub-Market', 'Avg Price/m²', 'YoY Growth', 'Foreign Buyer %', 'Key Nationality', 'Market Temp'],
    [
        ['Marbella', '€4,961', '+12.8%', '80–90%', 'British, Scandinavian', 'Hot'],
        ['Estepona', '€3,200', '+15.1%', '65%', 'British, Belgian', 'Very Hot'],
        ['Málaga City', '€2,950', '+16.2%', '35%', 'Spanish + intl.', 'Very Hot'],
        ['Benalmádena', '€2,600', '+11.5%', '55%', 'British, Irish', 'Hot'],
        ['Fuengirola', '€2,400', '+10.8%', '50%', 'Finnish, British', 'Hot'],
        ['Torremolinos', '€2,200', '+9.4%', '45%', 'British, Swedish', 'Warm'],
        ['Nerja', '€3,100', '+13.2%', '70%', 'British, Dutch', 'Hot'],
    ],
    col_widths=[3, 3, 2.5, 3, 3.5, 2.5]
)

add_heading(doc, 'B3. Foreign Buyer Analysis', 2)
add_body(doc, 'Foreign buyers represent 15.1% of all Spain transactions (2025) but ~40% of Málaga transactions — a critical differentiator for the Costa del Sol market.')
add_data_table(doc,
    ['Nationality', 'Share of Foreign Buys', 'Avg. Ticket', 'Preferred Zone', 'Key Driver'],
    [
        ['British', '9.7%', '€380k', 'Marbella, Benalmádena, Nerja', 'Lifestyle + weak £ opportunity'],
        ['German', '8.1%', '€350k', 'Mallorca, Costa del Sol', 'Investment + retirement'],
        ['French', '6.8%', '€310k', 'Costa Brava, Barcelona', 'Second home'],
        ['Moroccan', '5.2%', '€180k', 'Málaga city, Melilla', 'Diaspora + investment'],
        ['Romanian', '4.9%', '€150k', 'Madrid, Castellón', 'Residency + work'],
        ['Scandinavian', '4.3%', '€420k', 'Costa del Sol, Canarias', 'Retirement + sun'],
        ['American/Canadian', '3.8%', '€580k', 'Madrid, Barcelona, Marbella', 'Digital nomad + luxury'],
    ],
    col_widths=[3, 3.5, 2.5, 4, 3.5]
)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION C — STRUCTURED PROBLEM SOLVING
# ══════════════════════════════════════════════════════════════════════════════
doc.add_page_break()
add_heading(doc, 'SECTION C — STRUCTURED PROBLEM SOLVING: HUSPY MÁLAGA GROWTH', 1)
add_body(doc, 'Scenario: Huspy has established a Málaga footprint after 20× Spain growth in 2024. The task is to identify the highest-value levers to accelerate the business over the next 12 months.', italic=True, color=MUTED)

add_heading(doc, 'Step 1: Situation Diagnosis (What Do We Know?)', 2)
add_data_table(doc,
    ['Known Strengths', 'Known Gaps', 'Key Risks'],
    [
        ['Integrated mortgage (25+ bank partners)', 'Current mortgage attach rate (~15–20%)', 'E&V fined €22M — freelance model risk'],
        ['UAE playbook: 30% mortgage market share', 'Foreign buyer conversion tools/languages', 'Idealista controls lead flow — dependency'],
        ['Series B ($59M Balderton): growth capital', 'Developer/off-plan pipeline unclear', 'STR moratorium reduces investor returns'],
        ['AI tools + CRM = agent productivity edge', 'Luxury segment (Marbella) penetration', 'Supply shortage limits transaction volume'],
        ['20× Spain growth 2024 = brand momentum', 'Agent NPS and churn rate unknown', 'Golden Visa removed April 2025'],
        ['Málaga: foreign-heavy = digital platform fit', 'Rental vertical unexplored', 'Euribor reversal risk on mortgage product'],
    ],
    col_widths=[5.5, 5.5, 5.5]
)

add_heading(doc, 'Step 2: Value Gap Analysis', 2)
add_data_table(doc,
    ['Gap Area', 'Current State', 'Target State', 'Value Unlock', 'Priority'],
    [
        ['Mortgage Attach Rate', '~15–20% of deals include Huspy mortgage', '40%+ attach rate', 'Each mortgage = €2,000–3,000 revenue; zero incremental agent cost', 'HIGH'],
        ['Foreign Buyer Conversion', 'Foreign buyers (~40% Málaga) underserved', 'Multi-language platform; UK/DE/US marketing', '40% of 35,400 txns = 14,160 addressable/yr', 'HIGH'],
        ['New Build Developer Pipeline', '8,000 new homes expected Málaga 2025', '3–5 developer partnerships', 'Developer pays 5–10% commission on new build', 'MEDIUM-HIGH'],
        ['Luxury Segment (Marbella)', '142 sales >€1M in 2024; avg €4,961/m²', 'Dedicated luxury tier + Marbella team', '5% on €1M = €50k/transaction; 10 deals = €500k', 'MEDIUM-HIGH'],
        ['Rental Vertical', '€15.60/m²/month; +10% YoY', 'Rental management vertical', '€300/month × 200 properties = €60k/month recurring', 'MEDIUM'],
        ['Agent Retention', 'Industry churn ~25%/year (freelance)', '<15% annual churn; agents as partners', 'LTV €45k vs €3k replacement CAC = 15× ROI', 'MEDIUM'],
    ],
    col_widths=[3.5, 3.5, 3.5, 4, 2]
)

add_heading(doc, 'Step 3: 90-Day Action Plan', 2)
add_data_table(doc,
    ['Month', 'Focus', 'Key Actions', 'Output'],
    [
        ['Month 1', 'Diagnose & Baseline',
         '1. Audit agent pipeline: productivity, volume, segment mix\n2. Map current mortgage attach rate per agent\n3. Run P&L baseline: CAC, LTV, contribution margin\n4. Competitive intel: RE/MAX / Tecnocasa recruitment targets\n5. Journey mapping: where do foreign buyers drop off?',
         'Data-backed baseline P&L + agent segmentation matrix'],
        ['Month 2', 'Prioritize & Design',
         '1. Mortgage attach: train agents on pitch; fast-track approval with CaixaBank, BBVA, Santander\n2. Foreign buyer funnel: English/German landing pages; Kyero + A Place in the Sun partnerships\n3. Developer outreach: top 3 developers with 2025 Málaga launches\n4. Luxury pilot: recruit 2 Marbella-specialist agents\n5. Define KPI dashboard: GMV/agent, mortgage attach, NPS, churn',
         'Prioritized growth roadmap with owners, timelines, success metrics'],
        ['Month 3', 'Execute & Measure',
         '1. Agent count: +20% from baseline (Málaga + Marbella)\n2. Mortgage attach: target 40% by Q2\n3. Foreign buyer: 3 international inbound deals via new channels\n4. Developer: 1 signed MOU with new-build developer\n5. Weekly pipeline review: top 10 deals by stage every Monday',
         'Running operational cadence; 90-day KPI scorecard'],
    ],
    col_widths=[2, 3, 9, 4.5]
)

add_heading(doc, 'Step 4: Unit Economics Per Transaction (at Scale)', 2)
add_data_table(doc,
    ['Revenue Item', 'Amount', 'Cost Item', 'Amount'],
    [
        ['Agency commission (4% of €280k)', '€11,200', 'Lead generation (Idealista + paid)', '−€700'],
        ['Agent split (−70% to agent)', '−€7,840', 'Platform & tech (per transaction)', '−€200'],
        ['Huspy net from agency', '€3,360', 'Mortgage processing cost', '−€150'],
        ['Mortgage fee (bank pays, 40% attach)', '+€1,000', 'Compliance / legal', '−€100'],
        ['Premium services (legal, admin)', '+€300', 'TOTAL COST', '−€1,150'],
        ['TOTAL GROSS REVENUE', '€4,660', 'CONTRIBUTION MARGIN (75%)', '€3,510'],
    ],
    col_widths=[5.5, 2.5, 5.5, 2.5]
)
add_body(doc, 'Key lever: Every 10pp increase in mortgage attach rate adds ~€500 per transaction. Getting from 20% to 40% attach = +€1M in annual contribution at 50 txns/month — with zero additional agent cost.', bold=True, color=BLUE)

add_heading(doc, 'Step 5: KPI Dashboard', 2)
add_data_table(doc,
    ['KPI', 'Definition', 'Target (Yr 1)', 'Cadence'],
    [
        ['GMV per Agent', 'Total property value transacted per agent/month', '€350k/month (~1.5 txns)', 'Weekly'],
        ['Mortgage Attach Rate', '% of property sales with Huspy mortgage', '40%+', 'Weekly'],
        ['Lead → Listing Conv.', '% of portal leads becoming active listings', '10–12%', 'Weekly'],
        ['Listing → Sale Time', 'Days from live listing to signed Arras', '<35 days', 'Weekly'],
        ['Avg. Commission (net)', 'Average net fee per completed transaction', '€4,000+', 'Monthly'],
        ['Negotiation Discount', '% achieved below asking price per agent', '<5%', 'Monthly'],
        ['Agent NPS', 'Net Promoter Score from agents', '>55', 'Monthly'],
        ['Agent Churn Rate', '% of agents leaving platform annually', '<15%', 'Monthly'],
        ['Foreign Buyer %', '% of Huspy transactions with non-Spanish buyers', '>30%', 'Monthly'],
        ['Contribution Margin', 'Gross revenue minus direct variable costs', '>70%', 'Monthly'],
    ],
    col_widths=[4, 6, 3.5, 3]
)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION D — CUSTOMER LIFECYCLE PIPELINES
# ══════════════════════════════════════════════════════════════════════════════
doc.add_page_break()
add_heading(doc, 'SECTION D — CUSTOMER LIFECYCLE PIPELINES', 1)
add_body(doc, 'Two distinct journeys converge at the Escritura. Understanding both sides is what separates a reactive agent from a deal architect — and where Huspy\'s dual-sided platform creates compounding value.', italic=True, color=MUTED)

add_kpi_table(doc, [
    ('Seller: Avg Time to Close', '3–6 months', 'Listing mandate → Escritura'),
    ('Buyer: Avg Time to Close', '10–14 weeks', 'Offer accepted → keys in hand'),
    ('Arras Deposit', '10%', 'Arras Penitenciales — binding both sides'),
    ('Agency Commission', '3–6% + IVA', 'Paid by seller; buyer pays zero agency fees'),
    ("Seller's Key Tax", 'Plusvalía', 'Municipal land value gain — seller always pays'),
    ('Buyer Overhead', '10–12%', 'ITP 7% (Andalusia) + legal + notary + registry'),
])

# ── OWNER PIPELINE ────────────────────────────────────────────────────────────
add_heading(doc, 'D1. Property Owner — Selling Pipeline', 2, color=RGBColor(0x0A, 0x16, 0x28))
add_body(doc, 'Huspy opportunity: Winning the listing mandate (Stage 2) is the critical gate — once exclusive, Huspy controls the transaction from valuation to Escritura.', bold=True, color=GOLD)
doc.add_paragraph()

add_pipeline_stage(doc, 1, 'Property Valuation & Offer Preparation', '1–2 weeks',
    'Agent visits in person. Performs Comparative Market Analysis (CMA) using Registradores & Idealista data. Assesses condition, legal status, any illegal additions. Presents pricing proposal and net proceeds simulation (after taxes and commission).',
    ['Nota Simple', 'IBI receipt', 'Cedula de habitabilidad', 'CMA report', 'Realistic pricing = faster sale'])

add_pipeline_stage(doc, 2, 'Listing Agreement (Mandato)', 'Day 1–3',
    'Sign the agency mandate: Exclusive (Exclusiva) or Open (Abierto). Exclusive mandates deliver faster sales and better marketing investment. Agree on commission (3–6% + 21% IVA), timeline (3–6 months), and cancellation terms. Gather all required documents at this stage.',
    ['Escritura de propiedad', 'Energy certificate (CEE)', 'Community fees cert', 'Last 3 IBI receipts', 'Commission: 3–6% + 21% IVA'])

add_pipeline_stage(doc, 3, 'Active Marketing', 'Weeks 1–12+ ongoing',
    'Professional photography, video, floor plan. Publish on Idealista, Fotocasa, Kyero, Rightmove Overseas. For Málaga/Costa del Sol ~40% of buyers are foreign — international reach is essential. Weekly performance reporting to seller. Adjust strategy if no offers in 3–4 weeks.',
    ['Idealista / Fotocasa', 'Kyero / Rightmove Overseas', 'Social media', 'Overpricing = stale listing'])

add_pipeline_stage(doc, 4, 'Viewings & Buyer Communication', 'Ongoing',
    'Agent holds keys and arranges viewings often at short notice. Private viewings are standard in Spain. Collect structured feedback after each visit. Qualify interest level and filter serious prospects from time-wasters.',
    ['Viewing coordination', 'Post-visit feedback loop', '28% of Málaga properties sell within 30 days'])

add_pipeline_stage(doc, 5, 'Offer Negotiation', 'Days 1–7 per offer',
    'Agent presents all offers advising on buyer credibility (proof of funds / pre-approval), price, and conditions. National avg discount 6.2% below asking (2025). In hot Málaga sub-markets, offers often come at or above asking. Agent acts as buffer protecting seller while keeping buyer engaged.',
    ['Proof of funds / pre-approval', 'Counter-offer strategy', 'Avg discount: 6.2% nationally (lower in Málaga)'])

add_pipeline_stage(doc, 6, 'Contrato de Arras — Prepay / Pre-Contract', 'Week 1–2 after offer',
    'CRITICAL LEGAL STEP. Both parties sign Contrato de Arras Penitenciales — the standard binding pre-contract under Art. 1454 Civil Code. Buyer pays 10% deposit (señal). BUYER WITHDRAWS → loses deposit. SELLER WITHDRAWS → returns double (20%). Timeline to Escritura fixed: typically 30–90 days.',
    ['Arras Penitenciales — Art. 1454 Civil Code', 'Buyer pays 10%', 'Seller backs out → returns 20%', '30–90 days to closing fixed'])

add_pipeline_stage(doc, 7, 'Due Diligence & Legal Checks', '2–4 weeks (parallel)',
    "Buyer's lawyer pulls Nota Simple from Land Registry — confirms clean title, no liens, no outstanding debts. Seller clears any existing mortgage. Community checked for arrears. Energy Certificate (CEE) must be valid. Illegal constructions must be declared.",
    ['Nota Simple', 'Community arrears cert', 'CEE energy cert', 'Building permit check', 'Mortgage clearance'])

add_pipeline_stage(doc, 8, 'Escritura de Compraventa — Notary Closing', 'Week 10–14',
    'Both parties at notary office. Escritura Pública signed. Buyer transfers remaining 90% balance via banker\'s cheque — all funds must be traceable (AML). Notary reads deed in full. Keys exchanged. Legal ownership transfers at this moment.',
    ['Notary signing', 'Seller receives 90%', 'Keys exchanged', 'AML: all funds traceable'])

add_pipeline_stage(doc, 9, 'Post-Sale Obligations & Follow-up', 'Weeks 14–20 + next tax year',
    'Seller pays Plusvalía Municipal (land value gain tax) within 30 days to Ayuntamiento. Capital gain (IRPF) declared in following year\'s income tax. Cancel/transfer utility contracts. Agent follows up for testimonial, referral, and reinvestment conversation.',
    ['Plusvalía Municipal — 30 days', 'IRPF capital gain — next tax year', 'Cancel/transfer utilities', 'Referral & testimonial'])

# ── BUYER PIPELINE ────────────────────────────────────────────────────────────
doc.add_page_break()
add_heading(doc, 'D2. Buyer — Purchase Pipeline', 2, color=RGBColor(0x0A, 0x16, 0x28))
add_body(doc, 'Huspy opportunity: Two high-value touchpoints — NIE/banking setup (referral partnerships) and mortgage origination (Stage 3). Getting to the buyer at Stage 2–3 means owning the rest of the journey.', bold=True, color=GOLD)
doc.add_paragraph()

add_pipeline_stage(doc, 1, 'Define Budget & Goals', 'Before anything else',
    'Buyer defines location, property type, intended use (primary / holiday / investment), and budget ceiling. CRITICAL: budget must be +10–12% on top of purchase price for taxes and fees. A €300k property requires ~€333–336k total liquidity. Re-framing the true all-in cost is the agent\'s first job.',
    ['True cost = price + 10–12% taxes & fees', 'Location shortlist', 'Use case: live / holiday / invest'])

add_pipeline_stage(doc, 2, 'NIE & Spanish Bank Account', 'Start immediately — takes 2–3 months',
    'NIE (Número de Identificación de Extranjero) is mandatory for all foreigners buying in Spain. Without it: no bank account, no taxes, no notary signing. Apply at Spanish police station (Comisaría) or Spanish Consulate. Costs €10 (Tasa 790) but wait is 2–3 months in peak season.',
    ['No NIE = cannot complete any purchase', 'Modelo EX-15', 'Open Spanish bank account (all AML payments from here)'])

add_pipeline_stage(doc, 3, 'Mortgage Pre-Approval', '2–6 weeks',
    'Spanish banks lend up to 80% LTV for residents, 70% for non-residents. Bank issues FEIN (Ficha Europea de Información Normalizada) — binding offer, mandatory 10 days before notary signing. Huspy\'s 25+ bank panel = best market rate in one process. Euribor 12m ~2.5% (Mar 2026) — favourable window.',
    ['LTV: 80% residents / 70% non-residents', 'FEIN — mandatory 10 days before signing', 'Huspy: compare 25+ banks simultaneously'])

add_pipeline_stage(doc, 4, 'Property Search & Viewings', 'Ongoing — typically 1–3 months',
    'Search Idealista, Fotocasa, Kyero, Rightmove Overseas. Never buy unseen — always visit in person. Check the neighbourhood at different times of day. Ask specifically about community fees, pending building works, and historical rental performance for investments.',
    ['Idealista / Fotocasa / Kyero', 'Rightmove Overseas (UK buyers)', 'NEVER buy unseen', 'Check community fees + pending works'])

add_pipeline_stage(doc, 5, 'Legal Due Diligence (Before Offer)', '1–2 weeks',
    'Hire independent Spanish property lawyer (abogado). Spain operates on "buyer beware" (caveat emptor). Lawyer pulls Nota Simple (€9) — reveals ownership, mortgages, liens, registered size. Checks Catastro for planning compliance. Queries Ayuntamiento for fines or urban planning issues.',
    ['Nota Simple (€9)', 'Catastro check', 'Ayuntamiento planning query', 'CEE validation', 'Spain = buyer beware'])

add_pipeline_stage(doc, 6, 'Make an Offer', '1–5 days to agree',
    'Written offer via agent. In Málaga\'s competitive sub-markets (Marbella, Málaga city, Benalmádena), low-ball offers rejected quickly. National average discount 6.2% — but Málaga often 2–4% below or at asking. Include pre-approval letter to signal credibility.',
    ['Written offer via agent', 'Proof of funds / pre-approval letter', 'Málaga avg discount: 2–4% (national: 6.2%)'])

add_pipeline_stage(doc, 7, 'Contrato de Arras — Binding Pre-Contract', 'Within 1–2 weeks of offer',
    'Sign Contrato de Arras Penitenciales and pay 10% deposit. Property immediately off market. Point of no return — withdrawing costs the full deposit. Seller backs out → they owe you 20%. Contract fixes final price, completion date, and conditions. Having mortgage pre-approval BEFORE this stage is essential.',
    ['Arras Penitenciales — Art. 1454 Civil Code', 'Pay 10% — non-refundable if you withdraw', 'Get pre-approval BEFORE signing Arras'])

add_pipeline_stage(doc, 8, 'Tasación & Mortgage Finalisation', '3–6 weeks',
    'Bank orders official Tasación (€300–€800). If Tasación < purchase price, bank lends against lower figure — buyer must top up cash. Bank issues formal FEIN. Mandatory: buyer visits notary ALONE (without bank) at least 10 days before signing to confirm understanding — required by 2019 Mortgage Law.',
    ['Tasación — €300–€800', 'FEIN binding mortgage offer', 'Tasación < price = cash top-up', '10-day mandatory cooling-off after FEIN'])

add_pipeline_stage(doc, 9, 'Escritura de Compraventa — Notary Closing', 'Week 10–14',
    'Final signing at notary. Buyer and seller present (or PoA). Full remaining balance paid via banker\'s cheque from Spanish bank. Notary reads full deed. Both parties sign. Keys change hands. Legal ownership transfers at this exact moment. Since 2019: bank pays its own notary and registry costs for the mortgage deed.',
    ['Both parties present', 'Pay remaining 90% via banker\'s cheque', 'Keys received', 'Bank pays its own mortgage notary costs (since 2019)'])

add_pipeline_stage(doc, 10, 'Registration, Taxes & Post-Purchase Setup', 'Weeks 14–20',
    'Pay ITP (7% Andalusia resale) or IVA + AJD (10% + 1.2% new build) within 30 days. Gestoría files Land Registry (15–30 working days). Transfer utilities. Set up community fee direct debit. Investment property: register for tourist rental licence (VFT) if applicable.',
    ['ITP 7% Andalusia — pay within 30 days', 'Land Registry — 15–30 working days', 'Transfer utilities', 'Investment: apply for VFT rental licence'])

# ══════════════════════════════════════════════════════════════════════════════
# SECTION E — AGENCY ECOSYSTEM & OFFICE STRUCTURE
# ══════════════════════════════════════════════════════════════════════════════
doc.add_page_break()
add_heading(doc, 'SECTION E — AGENCY ECOSYSTEM & OFFICE STRUCTURE', 1)
add_body(doc, 'How the real estate agency business actually works: the mechanics of commission flow, every professional involved in a transaction, pricing methodology, and the full organisational structure of an agency office.', italic=True, color=MUTED)

# ── E1. HOW THE AGENCY BUSINESS MODEL WORKS ──────────────────────────────────
add_heading(doc, 'E1. How the Agency Business Model Works', 2)

add_heading(doc, 'Commission Flow: Who Pays What', 3)
add_body(doc, 'In Spain the agency commission is paid exclusively by the seller — buyers pay zero agency fees. The full flow from sale price to net revenue is as follows:')
add_data_table(doc,
    ['Step', 'Actor', 'Pays / Receives', 'Amount (example: €280k sale)', 'Notes'],
    [
        ['1', 'Buyer', 'Pays sale price', '€280,000', 'Plus ~10–12% in taxes & fees (ITP, notary, registry)'],
        ['2', 'Seller → Agency', 'Pays commission', '€11,200 (4%)', 'Commission + 21% IVA on top → seller pays €13,552 total'],
        ['3', 'IVA on commission', 'Seller pays to Hacienda', '€2,352 (21% IVA)', 'Agency collects and remits. Often missed in seller net calculations'],
        ['4', 'Franchise fee / brand royalty', 'Agency pays to franchise', '€1,120 (10% of gross)', 'RE/MAX, E&V, Century 21 all charge this. Independents keep full gross'],
        ['5', 'Agent split', 'Agency pays to agent', '€7,840 (70% split)', 'Typical range: 50–80% to agent depending on seniority and model'],
        ['6', 'Huspy net (agency)', 'Retained after split', '€3,360', 'Before platform costs, lead gen, compliance'],
        ['7', 'Mortgage fee (bank referral)', 'Bank pays Huspy', '+€1,000–1,500', '40% of deals attach a Huspy mortgage — zero extra agent cost'],
        ['8', 'Net contribution margin', 'Huspy retains', '~€3,500–4,500', 'Per transaction before overhead allocation'],
    ],
    col_widths=[0.6, 3, 3.5, 4, 5.4]
)
add_body(doc, 'Critical: IVA (21%) is charged ON TOP of the commission rate — a 4% commission on €280k is €11,200 but the seller actually pays €13,552. Always quote the IVA-inclusive figure to sellers to avoid disputes at closing.', bold=True, color=RED)

add_heading(doc, 'Mandate Types: Exclusive vs. Open', 3)
add_data_table(doc,
    ['Feature', 'Exclusiva (Exclusive)', 'Abierto (Open)'],
    [
        ['Definition', 'One agency has sole right to sell', 'Multiple agencies list simultaneously'],
        ['Portal investment', 'Full professional photos, video, featured listings', 'Minimal — agent won\'t invest if another sells first'],
        ['Agent effort', 'Full — agent knows they\'ll be paid', 'Partial — race-to-close mentality'],
        ['Avg. days on market', '30–45 days (Málaga)', '55–80 days (Málaga)'],
        ['Price achieved', 'Closer to asking (less urgency discounting)', 'More negotiation, more pressure from multiple agents'],
        ['Seller control', 'Single point of contact; coordinated strategy', 'Multiple agents, mixed messages, confused buyers'],
        ['Commission rate', 'Typically 3–4%', 'Typically 5–6% (compensates for lower certainty)'],
        ['Huspy approach', 'TARGET — pitch exclusiva on every listing', 'Accept short-term but convert to exclusiva within 30 days'],
    ],
    col_widths=[4, 6, 6.5]
)

add_heading(doc, 'Revenue Streams: How an Agency Makes Money', 3)
add_data_table(doc,
    ['Revenue Stream', 'Source', 'Typical Amount', 'Frequency', 'Huspy Advantage'],
    [
        ['Agency commission (sales)', 'Seller at Escritura', '3–6% + 21% IVA', 'Per transaction', 'Core — agent split 70%'],
        ['Mortgage origination fee', 'Bank pays on completion', '€1,000–1,500 per mortgage', 'Per mortgage closed', 'Huspy speciality — 25+ bank panel'],
        ['Insurance referral', 'Insurer pays commission', '€200–400 per policy', 'Per transaction', 'Home insurance mandatory for bank mortgage'],
        ['Legal / gestoría referral', 'Partner law firm or gestoría', '€150–300 per referral', 'Per transaction', 'Recommended but not mandatory'],
        ['New build developer fee', 'Developer pays on reservation', '4–8% on new build price', 'Per unit sold', 'Developers pay higher rates — no seller to negotiate with'],
        ['Rental management', 'Landlord pays monthly', '8–12% of monthly rent', 'Monthly recurring', 'Recurring revenue; Huspy opportunity to build vertical'],
        ['Rental finding fee', 'Landlord one-off', '1 month\'s rent', 'Per placement', 'Quick transaction; builds landlord relationship'],
    ],
    col_widths=[4, 3.5, 3.5, 2.8, 4]
)

add_heading(doc, 'Portal Dependency: The Idealista Problem', 3)
add_body(doc, 'The Spanish real estate market has a structural dependency on listing portals — primarily Idealista — that creates both a lead source and a margin risk for agencies.')
add_data_table(doc,
    ['Portal', 'Monthly Visits', 'Market Position', 'Cost to Agency', 'Buyer Profile', 'Risk for Huspy'],
    [
        ['Idealista', '50M+', '#1 — dominant', '€200–600/month per office', 'All buyer types; 60%+ start here', 'HIGH — if Idealista raises prices, Huspy\'s CAC rises'],
        ['Fotocasa', '20M+', '#2 — strong', '€100–300/month', 'Spanish domestic buyers', 'MEDIUM'],
        ['Kyero', '2M+', 'Foreign buyers specialist', '€150–400/month', 'British, Dutch, German buyers', 'LOW — useful for Málaga foreign buyer segment'],
        ['Rightmove Overseas', '1M+', 'UK buyers specifically', '€100–250/month', 'British buyers', 'LOW — niche but high-ticket Málaga buyers'],
        ['Pisos.com', '8M+', '#3 nationally', '€80–200/month', 'Price-sensitive domestic', 'LOW'],
    ],
    col_widths=[3.5, 2.5, 3, 3.5, 4, 3.5]
)
add_body(doc, 'Huspy\'s strategic hedge: mortgage lock-in. Once a buyer uses Huspy for a mortgage pre-approval, they become a Huspy buyer regardless of which portal they found the property on. This partially decouples lead ownership from portal dependency.', italic=True, color=BLUE)

# ── E2. FULL TRANSACTION ECOSYSTEM ───────────────────────────────────────────
add_heading(doc, 'E2. Full Ecosystem: Every Professional in a Transaction', 2)
add_body(doc, 'A Spanish real estate transaction involves up to 10 distinct professional roles. Understanding who each one is, what they do, who pays them, and where Huspy intersects is essential knowledge for any operations director.')
add_data_table(doc,
    ['Role', 'Spanish Title', 'Who Pays', 'Typical Fee', 'At Which Stage', 'Huspy Touchpoint?'],
    [
        ['Real estate agent (seller side)', 'Agente inmobiliario', 'Seller', '3–6% + IVA', 'Valuation → Escritura', 'YES — core Huspy employee/partner'],
        ['Real estate agent (buyer side)', 'Agente del comprador', 'Seller (same pool)', 'Shared from same commission', 'Search → Arras', 'YES — Huspy agent represents buyer'],
        ['Mortgage broker', 'Corredor de crédito hipotecario', 'Bank (on completion)', '€1,000–1,500 per mortgage', 'Pre-approval → FEIN', 'YES — Huspy\'s core differentiator'],
        ['Property lawyer', 'Abogado inmobiliario', 'Buyer (always)', '€1,500–3,000 flat fee', 'Due diligence → Escritura', 'Referral partner opportunity'],
        ['Notary', 'Notario', 'Buyer + Seller (split)', '€800–1,500 (regulated scale)', 'Escritura signing only', 'Required by law — no flexibility'],
        ['Official valuer', 'Tasador homologado', 'Buyer (bank requires it)', '€300–800', 'After Arras, before FEIN', 'Ordered by the bank Huspy places'],
        ['Administrative agency', 'Gestoría', 'Usually buyer', '€300–600 flat fee', 'Post-Escritura', 'Referral partner for post-sale admin'],
        ['Land registrar', 'Registrador de la Propiedad', 'Buyer', '€400–800 (regulated)', 'Post-Escritura registration', 'Required by law — no flexibility'],
        ['Home inspector / surveyor', 'Arquitecto técnico / ITE inspector', 'Buyer (optional)', '€300–600', 'Before Arras (recommended)', 'Referral opportunity — low penetration in Spain vs. UK'],
        ['Insurance broker', 'Corredor de seguros', 'Buyer (annual premium)', '€200–400/year', 'At mortgage closing', 'YES — home insurance mandatory for bank mortgage; Huspy referral fee'],
        ['Community administrator', 'Administrador de fincas', 'Community (owners)', '€20–50/month per unit', 'Ongoing post-purchase', 'Relevant for investor buyers managing multiple units'],
    ],
    col_widths=[3.5, 3.5, 2.5, 3, 3.5, 3.5]
)
add_body(doc, 'Key insight: Huspy directly monetises 3 of these 11 roles (agent, mortgage broker, insurance referral) and has referral relationships with 2 more (lawyer, gestoría). A fully integrated Huspy transaction can generate €5,000–6,500 per deal vs. €3,500 for agency-only.', bold=True, color=GREEN)

# ── E3. PRICING DYNAMICS DEEP-DIVE ───────────────────────────────────────────
add_heading(doc, 'E3. Pricing Dynamics — How Prices Are Set and What Moves Them', 2)

add_heading(doc, 'CMA Methodology: How Agents Set Listing Prices', 3)
add_body(doc, 'A Comparative Market Analysis (Análisis de Mercado Comparativo) is the standard tool for pricing a listing. A rigorous CMA is the single most important thing an agent does — overpricing kills deals, underpricing destroys seller trust.')
add_data_table(doc,
    ['CMA Step', 'What to Do', 'Tool / Source', 'Adjustment Factor'],
    [
        ['1. Pull comparables', 'Find 3–5 SOLD properties: same zone, same type, last 6–12 months', 'Registradores.es, Idealista Valoraciones, Tinsa', 'Sold price only — never listed price'],
        ['2. Calculate base price/m²', 'Sold price ÷ superficie útil (usable area, NOT constructed)', 'Catastro certificate + Nota Simple', 'Spain uses útil, not construida — big difference in apartments'],
        ['3. Condition adjustment', 'Assess reform requirement honestly', 'Agent inspection + photos', 'Full reform needed: −15 to −20%; updated kitchen/bath: −5%'],
        ['4. Floor & light premium', 'Higher floors + natural light + views', 'Market data + agent judgement', 'Penthouse vs. ground floor: +10–20% in Málaga city'],
        ['5. Orientation premium', 'South-facing in Spain = more light = more value', 'Agent site visit', '+5–8% for south-facing in Málaga'],
        ['6. Parking & storage', 'Garage space + trastero', 'Nota Simple / community registry', 'Parking: +€15k–€25k in city; +€8k in coast'],
        ['7. Urgency discount', 'Seller timeline affects price strategy', 'Seller interview at valuation', 'Need to sell in <60 days: −5 to −8% to ensure fast close'],
        ['8. Portal price check', 'Cross-check current active listings', 'Idealista.com live search', 'If zone has 20+ listings, buyer has choice — price sharper'],
    ],
    col_widths=[3, 4.5, 4, 5]
)
add_body(doc, 'Critical distinction: Superficie útil (usable internal area) vs. superficie construida (constructed area including walls, shared areas). Spanish law requires marketing on útil. A 90m² construida apartment is typically 75–80m² útil. Using the wrong figure makes your CMA wrong by 10–15%.', bold=True, color=RED)

add_heading(doc, 'Pricing by Property Type', 3)
add_data_table(doc,
    ['Property Type', 'Spanish Term', 'Avg Price/m² Málaga', 'Avg Price/m² Spain', 'Premium vs. National', 'Key Notes'],
    [
        ['City apartment', 'Piso', '€2,950', '€2,311', '+27.7%', 'Most liquid asset class; fastest to sell'],
        ['Penthouse', 'Ático', '€3,800–5,500', '€2,900', '+30–90%', 'Terrace size is a multiplier; views command 20%+ premium'],
        ['Townhouse', 'Adosado / Casa adosada', '€2,400', '€1,950', '+23%', 'Popular with families; community fees lower than apartments'],
        ['Detached villa', 'Chalet / Villa independiente', '€3,500–8,000+', '€2,600', '+35–200%', 'Wide range — plot size, pool, privacy are key value drivers'],
        ['Ground floor + garden', 'Bajo con jardín', '€2,200–2,600', '€1,900', '+15%', 'Discount vs. upper floors in city; premium on coast (private garden)'],
        ['New build apartment', 'Obra nueva / Piso de nueva construcción', '€3,500–4,500', '€2,800', '+25–60%', 'Pays IVA 10% (not ITP 7%); higher quality; delivery risk'],
        ['Commercial premises', 'Local comercial', '€1,800–3,500', '€1,600', '+12–120%', 'High variance; location is everything; lower liquidity'],
        ['Garage space', 'Plaza de garaje', '€15k–€35k (unit)', '€10k–€20k', '+50–75%', 'Sold separately; high demand in Málaga city centre'],
    ],
    col_widths=[3.5, 3.5, 3, 3, 3, 5.5]
)

add_heading(doc, 'Seasonal Dynamics: When to List, When to Buy', 3)
add_data_table(doc,
    ['Quarter', 'Season', 'Market Activity', 'For Sellers', 'For Buyers', 'Costa del Sol Specifics'],
    [
        ['Q1 (Jan–Mar)', 'Winter warm-up', 'Supply low, motivated buyers return after holidays', 'Good time to list — less competition from other sellers', 'Motivated buyers; less choice drives stronger offers', 'Foreign buyers from UK/Germany research in Jan; view in Feb–Mar'],
        ['Q2 (Apr–Jun)', 'Peak season', 'Highest viewings, most new listings, competitive offers', 'Best time to list — maximum buyer pool', 'Most choice but most competition from other buyers', 'Easter triggers viewing surge; foreign buyers plan summer purchase'],
        ['Q3 (Jul–Aug)', 'Summer divergence', 'Domestic market slows; Costa del Sol accelerates', 'Risky in cities — locals on holiday; less viewings', 'Good for buyers in cities (less competition); difficult on coast', 'INVERSE on Costa: foreign holiday buyers = peak demand Jul–Aug'],
        ['Q4 (Oct–Dec)', 'Year-end push', 'Motivated sellers accept lower prices to close before year-end', 'Price pressure — motivated sellers compete', 'Best negotiating leverage of the year', 'Foreign buyers return post-summer; Oct–Nov strong for international'],
    ],
    col_widths=[2.5, 2.5, 4, 3.5, 3.5, 5]
)
add_body(doc, 'Málaga exceptionalism: The Costa del Sol runs almost inverse to the national seasonal pattern. July–August is peak for foreign buyers (holiday viewings), whereas Madrid and Barcelona nearly shut down. This is why Málaga agents must be active in August when mainland agencies take vacaciones.', italic=True, color=BLUE)

add_heading(doc, 'Off-Plan & New Build: Different Rules, Different Dynamics', 3)
add_data_table(doc,
    ['Feature', 'Off-Plan / New Build (Obra Nueva)', 'Resale (Segunda Mano)', 'Implication for Huspy'],
    [
        ['Transaction tax', 'IVA 10% + AJD 1.2% = 11.2% total', 'ITP 7% (Andalusia)', 'Buyer pays more tax on new build — must budget correctly'],
        ['Typical price premium', '+10–20% above equivalent resale at launch', 'Market price at time of sale', 'Premium justified by quality, guarantee, no reform needed'],
        ['Payment structure', '20–30% deposit off-plan; rest at Escritura (delivery)', '10% Arras; 90% at Escritura', 'Buyer needs cash tied up for 18–36 months during construction'],
        ['Legal guarantee', 'Developer guarantee: 1yr finishing, 3yr habitability, 10yr structure', 'Sold as seen (caveat emptor)', 'New build has stronger consumer protections'],
        ['Commission rate (for agent)', '4–8% — developer pays; no negotiation with owner', '3–6% — negotiated with seller', 'Higher commission but developer relationship is key gate'],
        ['Key risk', 'Delivery delays, developer insolvency, quality shortfall', 'Legal issues (debts, liens, illegal additions)', 'Off-plan bank deposits must be legally protected (Law 20/2015)'],
        ['Mortgage timing', 'Bridge mortgage during construction then full mortgage at delivery', 'Standard mortgage applied at Arras stage', 'More complex mortgage structuring — Huspy advisory value'],
        ['Huspy opportunity', 'Developer partnership = pipeline of future mortgage clients pre-loaded', 'Standard transaction pipeline', '8,000 new build units expected Málaga 2025 = major pipeline opportunity'],
    ],
    col_widths=[4, 5, 4.5, 4]
)

# ── E4. REAL ESTATE OFFICE: WHO DOES WHAT ────────────────────────────────────
add_heading(doc, 'E4. Real Estate Office Organisation: Roles, Structure & Compensation', 2)
add_body(doc, 'Understanding who needs to be inside a real estate office — and how they are compensated — is critical for building a high-performance team and avoiding the legal mistakes that have cost competitors (E&V: €22M fine) significant money.')

add_heading(doc, 'Office Org Chart: Every Role Explained', 3)
add_data_table(doc,
    ['Role / Title', 'Spanish Title', 'Reports To', 'Core Responsibility', 'Comp Model', 'Typical Headcount'],
    [
        ['Branch Director', 'Director / Gerente de Oficina', 'Regional Director', 'P&L ownership; team hiring & coaching; senior deal support; KPI accountability', 'Base €2,500–3,500/month + override on team performance (1–2% of team GMV)', '1 per office'],
        ['Senior Agent', 'Agente Senior / Consultor Senior', 'Branch Director', 'Full-cycle independent deals (valuation → Escritura); mentors juniors; handles luxury segment', 'Commission only: 65–75% split; OTE €60k–€120k/year', '2–4 per office'],
        ['Junior Agent', 'Agente Junior / Consultor', 'Senior Agent or Director', 'Lead follow-up, viewings, buyer qualification; works with mentor for first 6 months', 'Commission only: 50–60% split; OTE €25k–€50k first year; ramp 3–6 months', '3–6 per office'],
        ['Mortgage Advisor', 'Asesor Hipotecario', 'Regional Mortgage Lead', 'Mortgage intake at offer stage; bank liaison; pre-approvals; FEIN processing', 'Base €1,800–2,500/month + €200–400 per mortgage closed', '1 per 2 offices (shared)'],
        ['Transaction Coordinator', 'Coordinadora de Operaciones', 'Branch Director', 'Arras to Escritura paperwork; deadline tracking; notary booking; post-sale admin chase', 'Fixed salary €1,600–2,200/month; no commission', '1 per office'],
        ['Marketing Coordinator', 'Coordinadora de Marketing', 'Regional Marketing or Director', 'Idealista/Fotocasa portal listings; photography coordination; social media; CRM updates', 'Fixed salary €1,500–2,000/month', '1 per 2–3 offices (shared)'],
        ['Administrative / Reception', 'Administrativo/a / Recepción', 'Branch Director', 'Inbound call handling; viewing scheduling; document filing; client welcome', 'Fixed salary €1,400–1,800/month', '1 per office'],
        ['Compliance / Legal Officer', 'Responsable de Cumplimiento', 'Operations Director', 'KYC/AML documentation; contract review; GDPR compliance; agent contract management', 'Fixed salary €2,500–3,500/month (often shared across region)', '1 per region'],
    ],
    col_widths=[3.5, 3.5, 3, 4, 4, 2.5]
)

add_heading(doc, 'Compensation Models: The Freelance vs. Employee Decision', 3)
add_body(doc, 'This is the single most legally sensitive decision in building a real estate team in Spain. Engel & Völkers were fined €22M (TEAC 2023) for misclassifying employees as freelance agents. The distinction has major financial and operational consequences.')
add_data_table(doc,
    ['Feature', 'Autónomo (Freelance)', 'Contrato Laboral (Employee)', 'Hybrid Model'],
    [
        ['Legal basis', 'Self-employed; issues invoices to agency', 'Full employment contract; agency is employer', 'Employee with variable commission component'],
        ['Social security', 'Agent pays own cuota autónomos (€294–€530/month)', 'Agency pays 23.6% SS + agent pays 6.4%', 'Agency pays SS; agent receives base + commission'],
        ['Agency cost (per agent)', 'Commission split only; no SS, no severance', 'SS contribution + holiday pay + severance risk', 'Higher fixed cost but legally safe'],
        ['Agent earnings', '50–80% commission split; no guarantees', 'Base salary €1,200–1,800 + 30–50% commission split', 'Base €1,000–1,500 + 40–60% commission split'],
        ['Control over agent', 'LOW — cannot mandate hours, scripts, exclusivity', 'HIGH — can mandate hours, training, exclusivity', 'MEDIUM — must be genuine contract not fake freelance'],
        ['Risk of misclassification', 'HIGH if agency controls work like an employee', 'ZERO if properly contracted', 'LOW if structured correctly'],
        ['Used by', 'RE/MAX, Tecnocasa, many independents', 'Growing adoption post E&V fine; Huspy policy', 'Emerging best practice in Spain'],
        ['Huspy position', 'Avoid — legal and cultural risk', 'Preferred model for compliance', 'Acceptable for senior agents with genuine autonomy'],
    ],
    col_widths=[4, 5, 5, 5]
)
add_body(doc, 'The E&V case (TEAC resolution 2023): Inspectors found agents were working fixed hours in E&V offices, using E&V systems, wearing E&V uniforms — yet classified as autónomos. The court ruled this was an employment relationship. Fine: €22M back SS contributions + penalties. Any agency replicating this model faces the same exposure.', bold=True, color=RED)

add_heading(doc, 'Agent Economics: What Does a Huspy Agent Actually Earn?', 3)
add_data_table(doc,
    ['Scenario', 'Txns/Month', 'Avg Ticket', 'Commission (4%)', 'Agent Split (70%)', 'Mortgage Attach (40%)', 'Agent Monthly OTE', 'Annual OTE'],
    [
        ['Starting agent (ramp)', '0.5', '€240,000', '€9,600 gross', '€6,720 net', '+€200 avg', '~€6,920', '~€83k'],
        ['Productive agent', '1.0', '€265,000', '€10,600 gross', '€7,420 net', '+€400 avg', '~€7,820', '~€94k'],
        ['Strong performer', '1.5', '€280,000', '€11,200 gross', '€7,840 net', '+€600 avg', '~€8,440', '~€101k'],
        ['Top performer', '2.0', '€300,000', '€12,000 gross', '€8,400 net', '+€800 avg', '~€9,200', '~€110k'],
        ['Luxury specialist (Marbella)', '1.0', '€600,000', '€24,000 gross', '€16,800 net', '+€600 avg', '~€17,400', '~€209k'],
    ],
    col_widths=[4, 2.5, 2.5, 3, 3, 3, 3, 2.5]
)
add_body(doc, 'Mortgage attach is the multiplier. Every deal where the buyer uses a Huspy mortgage adds €400–800 to the agent\'s income at zero extra work. Training every agent to position the Huspy mortgage pre-approval at the first buyer meeting is the highest-ROI activity in the entire operation.', bold=True, color=GREEN)

add_heading(doc, 'Recruiting & Onboarding: What Makes a High-Performing Agent', 3)
add_data_table(doc,
    ['Profile Factor', 'Green Flag', 'Red Flag', 'Huspy Priority Weight'],
    [
        ['Language skills', 'English + Spanish fluent; German/French a bonus', 'Spanish only in a 40% foreign-buyer market', 'HIGH — Málaga foreign buyer segment requires it'],
        ['Track record', '6+ closed deals in last 12 months at prior agency', 'No verifiable closed transactions', 'HIGH — commission-only roles need prior proof'],
        ['Network / book of business', 'Existing seller contacts; lawyer/investor network', 'Starting from zero with no warm contacts', 'MEDIUM — network accelerates first 90 days'],
        ['Tech adoption', 'Uses CRM proactively; comfortable with digital tools', 'Resists CRM; relies only on phone and WhatsApp', 'HIGH — Huspy platform adoption is non-negotiable'],
        ['Mortgage literacy', 'Can explain FEIN, Euribor, LTV to a client', 'Cannot explain the difference between fixed and variable rate', 'HIGH — mortgage attach depends on agent credibility'],
        ['Motivation fit', 'Intrinsically motivated; entrepreneur mindset', 'Expects a salary with no performance accountability', 'HIGH — commission model requires self-driven profile'],
        ['Compliance awareness', 'Understands AML obligations, GDPR, source of funds', 'No awareness of KYC or money laundering red flags', 'MEDIUM-HIGH — Huspy operates to bank compliance standards'],
    ],
    col_widths=[3.5, 5, 4.5, 5.5]
)
doc.add_paragraph()

# ══════════════════════════════════════════════════════════════════════════════
# SPANISH TERMINOLOGY GLOSSARY
# ══════════════════════════════════════════════════════════════════════════════
doc.add_page_break()
add_heading(doc, 'APPENDIX — Essential Spanish Real Estate Terminology', 1)
add_body(doc, 'Key terms every Huspy agent must know fluently — they appear at every stage of both the seller and buyer pipelines.', italic=True, color=MUTED)
doc.add_paragraph()

# header row
hdr_tbl = doc.add_table(rows=1, cols=3)
hdr_tbl.style = 'Table Grid'
for i, h in enumerate(['Spanish Term', 'English', 'Definition']):
    cell_text(hdr_tbl.rows[0].cells[i], h, bold=True, size=9, color=WHITE)
    set_cell_bg(hdr_tbl.rows[0].cells[i], '0A1628')

terms = [
    ('Arras Penitenciales', 'Binding deposit contract',
     'Standard 10% deposit pre-contract. Buyer loses deposit if they withdraw; seller pays double if they withdraw. Art. 1454 Civil Code.'),
    ('Escritura de Compraventa', 'Public deed of sale / Closing',
     'Notarised title deed signed at the notary office. Legal ownership transfers at the exact moment of signing.'),
    ('Nota Simple', 'Land Registry extract',
     'A €9 document from the Registro de la Propiedad showing current owner, mortgages, liens, and registered description. First thing any lawyer checks.'),
    ('Tasación', 'Official bank valuation',
     'Mandatory certified appraisal ordered by the bank (€300–€800). Sets maximum loan amount — if lower than purchase price, buyer must top up in cash.'),
    ('Plusvalía Municipal', 'Municipal capital gains tax (seller)',
     'Tax on the increase in cadastral land value since the last sale. Always paid by the seller — due within 30 days at local Ayuntamiento.'),
    ('NIE', 'Foreigner ID number',
     'Número de Identificación de Extranjero. Mandatory for all foreigners buying property, opening a bank account, or paying taxes in Spain. Takes 2–3 months to obtain.'),
    ('FEIN', 'European mortgage offer sheet',
     'Ficha Europea de Información Normalizada. Binding mortgage offer — must be provided at least 10 days before notary signing (2019 Mortgage Law).'),
    ('ITP', 'Property Transfer Tax (buyer)',
     'Impuesto de Transmisiones Patrimoniales. Paid by buyer on resale properties. 7% in Andalusia — one of Spain\'s lowest rates.'),
    ('Mandato de Exclusividad', 'Exclusive listing agreement',
     'Agreement granting one agency sole rights to sell the property. Delivers faster sales and better marketing vs. open mandates with multiple agencies.'),
    ('Contrato de Reserva', 'Reservation contract',
     'Initial soft reservation before Arras. Buyer pays €3,000–€10,000 to take property off market for 2–4 weeks. Refundable before Arras signed.'),
    ('Catastro', 'Cadastral registry',
     'Government property register showing physical characteristics, location, and fiscal value. Separate from Land Registry. Used for planning compliance checks.'),
    ('Gestoría', 'Administrative agency',
     'Professional administrator who handles tax filings, Land Registry registration, and utility transfers. Essential for post-completion admin.'),
]

data_tbl = doc.add_table(rows=len(terms), cols=3)
data_tbl.style = 'Table Grid'
for ri, (es, en, desc) in enumerate(terms):
    row = data_tbl.rows[ri]
    bg  = 'F0F4FF' if ri % 2 == 0 else 'FFFFFF'
    cell_text(row.cells[0], es, bold=True, size=9, color=NAVY)
    cell_text(row.cells[1], en, size=9, italic=True, color=MUTED)
    cell_text(row.cells[2], desc, size=9)
    for c in row.cells:
        set_cell_bg(c, bg)

for row in data_tbl.rows:
    row.cells[0].width = Cm(4)
    row.cells[1].width = Cm(4)
    row.cells[2].width = Cm(8.5)

doc.add_paragraph()

# ── Footer note ───────────────────────────────────────────────────────────────
add_body(doc, 'Sources: INE · Banco de España · Registradores de España · Tinsa · Idealista Research · CaixaBank Research · Spaincheck.com · Soleada · Terreta Spain · Casa Capitals · Lexidy · Strong Abogados', size=8, color=MUTED, italic=True)
add_body(doc, 'Data reflects best available information as of March 2026. This report is for internal strategic planning purposes.', size=8, color=MUTED, italic=True)

# ── Save ───────────────────────────────────────────────────────────────────────
out = '/home/user/OlgaGomez108/Huspy_Malaga_Market_Analysis_2026.docx'
doc.save(out)
print(f'Saved: {out}')
