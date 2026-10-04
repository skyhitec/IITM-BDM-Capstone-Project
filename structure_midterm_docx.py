"""
Master Structuring & Numbering Script for Mid term.docx
Fixes all Headings, Subheadings, Section Numbers, Table Captions (Table 1, Table 2), 
Figure Captions (Figure 1 to 10), TOC Index Table, Page Breaks, and Formatting.
"""
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml
import os

CHARTS_DIR = 'BDM_Charts'

def set_para_format(para, font_name='Times New Roman', font_size=11, 
                    bold=False, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
                    space_after=4, space_before=0, line_spacing=1.3,
                    italic=False):
    para.alignment = alignment
    pf = para.paragraph_format
    pf.space_after = Pt(space_after)
    pf.space_before = Pt(space_before)
    pf.line_spacing = line_spacing
    
    for run in para.runs:
        run.font.name = font_name
        run.font.size = Pt(font_size)
        run.font.bold = bold
        run.font.italic = italic

def add_heading_styled(doc, text, level=1):
    heading = doc.add_heading(text, level=level)
    for run in heading.runs:
        run.font.name = 'Times New Roman'
        run.font.color.rgb = RGBColor(0, 0, 0)
        if level == 1:
            run.font.size = Pt(14)
        elif level == 2:
            run.font.size = Pt(12)
        else:
            run.font.size = Pt(11)
    heading.paragraph_format.space_before = Pt(10)
    heading.paragraph_format.space_after = Pt(4)
    heading.paragraph_format.line_spacing = 1.3
    return heading

def add_para(doc, text, bold=False, font_size=11, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
             space_after=4, italic=False):
    para = doc.add_paragraph()
    run = para.add_run(text)
    run.font.name = 'Times New Roman'
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    para.alignment = alignment
    para.paragraph_format.space_after = Pt(space_after)
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.line_spacing = 1.3
    return para

def add_figure_compact(doc, image_path, caption, width=4.6):
    if os.path.exists(image_path):
        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        para.paragraph_format.space_before = Pt(2)
        para.paragraph_format.space_after = Pt(2)
        run = para.add_run()
        run.add_picture(image_path, width=Inches(width))
        
        cap = doc.add_paragraph()
        cap_run = cap.add_run(caption)
        cap_run.font.name = 'Times New Roman'
        cap_run.font.size = Pt(9.5)
        cap_run.font.italic = True
        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cap.paragraph_format.space_after = Pt(6)
        cap.paragraph_format.line_spacing = 1.1

def make_table_cell_formatted(cell, text, bold=False, font_size=9, bg_color=None):
    cell.text = ''
    para = cell.paragraphs[0]
    run = para.add_run(str(text))
    run.font.name = 'Times New Roman'
    run.font.size = Pt(font_size)
    run.font.bold = bold
    para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    para.paragraph_format.space_after = Pt(1)
    para.paragraph_format.space_before = Pt(1)
    para.paragraph_format.line_spacing = 1.0
    if bg_color:
        shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{bg_color}"/>')
        cell._tc.get_or_add_tcPr().append(shading)

def add_page_numbers(doc):
    for section in doc.sections:
        footer = section.footer
        footer.is_linked_to_previous = False
        para = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = para.add_run()
        fld_char_begin = parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="begin"/>')
        run._r.append(fld_char_begin)
        run2 = para.add_run()
        instr = parse_xml(f'<w:instrText {nsdecls("w")} xml:space="preserve"> PAGE </w:instrText>')
        run2._r.append(instr)
        run3 = para.add_run()
        fld_char_end = parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="end"/>')
        run3._r.append(fld_char_end)
        for r in [run, run2, run3]:
            r.font.name = 'Times New Roman'
            r.font.size = Pt(9.5)

# Initialize Document
doc = Document()
style = doc.styles['Normal']
style.font.name = 'Times New Roman'
style.font.size = Pt(11)
style.paragraph_format.line_spacing = 1.3
style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

for section in doc.sections:
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(2.0)

# ============================================================
# COVER PAGE (Page 1)
# ============================================================
for _ in range(4): doc.add_paragraph()
t_para = doc.add_paragraph()
t_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = t_para.add_run('Sales Trend and Inventory Optimization Analysis\nof Anand Pharma, Darbhanga')
run.font.name = 'Times New Roman'
run.font.size = Pt(20)
run.font.bold = True

sub = doc.add_paragraph()
sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = sub.add_run('A Mid-Term Project Report | Primary Data Pathway (Jan–June 2026)')
run.font.name = 'Times New Roman'
run.font.size = Pt(12)
run.font.italic = True

for _ in range(4): doc.add_paragraph()

det = doc.add_paragraph()
det.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = det.add_run('Submitted by:\nName: Shudhanshu Kumar Yadav | Roll Number: 23F1000204\nEmail: 23f1000204@ds.study.iitm.ac.in\n\nIITM Online BS Degree Program\nIndian Institute of Technology Madras, Chennai\nAugust 2026')
run.font.name = 'Times New Roman'
run.font.size = Pt(11)

doc.add_page_break()

# ============================================================
# CONTENT / INDEX PAGE (Page 2)
# ============================================================
add_heading_styled(doc, 'Content / Index Page', level=1)
add_para(doc, 'Table of contents indicating major report sections, sub-sections, and page numbers:')

toc_data = [
    ['Section No.', 'Section Heading', 'Page No.'],
    ['1.', 'Executive Summary and Title', '3'],
    ['2.', 'Proof of Data Originality', '4'],
    ['2.5', 'Problem Statement and Business Objectives', '4'],
    ['3.', 'Metadata and Descriptive Statistics', '5'],
    ['3.1', '  Metadata Definition and Dataset Structure', '5'],
    ['3.2', '  Descriptive Statistics and Quantitative Summary', '5'],
    ['4.', 'Detailed Explanation of Analysis Process / Method', '6'],
    ['4.1', '  Data Cleaning and Preprocessing', '6'],
    ['4.2', '  Analytical Methodologies and Mathematical Abstractions', '6'],
    ['5.', 'Results and Findings (Preliminary Insights)', '7'],
    ['5.1', '  ABC Classification & Pareto Analysis Findings', '7'],
    ['5.2', '  FSN Inventory Movement & Expiry Risk Findings', '8'],
    ['5.3', '  Monthly Sales Trends & Seasonal Demand Fluctuations', '9'],
    ['5.4', '  Moving Average Demand Forecast Results', '10'],
    ['6.', 'Interpretation of Results and Recommendations', '11'],
]

t_toc = doc.add_table(rows=len(toc_data), cols=3)
t_toc.style = 'Table Grid'
t_toc.alignment = WD_TABLE_ALIGNMENT.CENTER

for i, row_data in enumerate(toc_data):
    for j, cell_text in enumerate(row_data):
        cell = t_toc.cell(i, j)
        make_table_cell_formatted(cell, cell_text, bold=(i == 0), font_size=9.5,
                                  bg_color='1a73e8' if i == 0 else None)
        if i == 0:
            for run in cell.paragraphs[0].runs: run.font.color.rgb = RGBColor(255, 255, 255)

doc.add_page_break()

# ============================================================
# SECTION 1: EXECUTIVE SUMMARY (Page 3)
# ============================================================
add_heading_styled(doc, '1. Executive Summary and Title', level=1)
add_para(doc, 'Project Title: Sales Trend and Inventory Optimization Analysis of Anand Pharma, Darbhanga', bold=True)

add_para(doc, 
    'Anand Pharma is a retail B2C pharmacy store established in 2018 and operated by Mr. Krishan Mohan Yadav in Donar, Darbhanga, Bihar (GSTIN: 10AEY3656A1Z4), procuring primarily from M/s. Sonu Pharma, Patna. '
    'The store encounters inventory challenges including seasonal demand fluctuations and intuition-based restocking decisions, resulting in stockouts of high-demand items, overstocking, and capital loss from expired medicines.')

add_para(doc,
    'Primary sales, purchase, and stock data spanning 6 consecutive months (January to June 2026) were collected directly from physical vouchers and sales registers via remote digital transmission. '
    'The dataset covers 30 medicine SKUs tracked across 11 structured variables: S.No, Medicine Name, Category, MRP (₹), Purchase Price (₹), Opening Stock, Purchase Qty, Sales Qty, Closing Stock, Total Revenue (₹), and Expiry Date. '
    'Key descriptive statistics show mean revenue of ₹43,521.83 per SKU (median ₹41,108.50), total store revenue of ₹13,05,655, and an average gross margin of ~29.8%.')

add_para(doc,
    'Quantitative techniques — ABC Analysis, FSN Classification, Pareto (80/20) Analysis, and 3-Month Moving Average Demand Forecasting — were executed. '
    'Preliminary findings reveal that 14 Category A medicines (46.7% items) generate 69.0% of revenue; 18 SKUs drive 80% cumulative revenue (Pareto cutoff); 5 Non-Moving SKUs carry high expiry risk; and May 2026 recorded peak revenue (₹2,43,998 — a 23.8% surge over April). '
    'Moving Average forecasts for July 2026 were generated for all SKUs to guide proactive procurement.')

doc.add_page_break()

# ============================================================
# SECTION 2 & 2.5: PROOF OF DATA ORIGINALITY & PROBLEM STATEMENT (Page 4)
# ============================================================
add_heading_styled(doc, '2. Proof of Data Originality', level=1)
add_para(doc,
    'Primary data for 6 consecutive months (January 2026 to June 2026) was collected directly from M/s. Anand Pharma, Donar, Darbhanga, Bihar (GSTIN: 10AEY3656A1Z4; Proprietor: Mr. Krishan Mohan Yadav) through a remote digital pathway (scanned purchase tax invoices and daily sales registers shared via WhatsApp, cross-verified via telephonic calls).')

add_para(doc, '● Primary Sales, Purchase & Inventory Dataset (Jan–June 2026): [Click Here]', bold=True)
add_para(doc, '● Distributor Purchase Invoices & GST Details (M/s. Sonu Pharma): [Click Here]', bold=True)
add_para(doc, '● Pharmacy Firm Images & Servicescape Photos: [Click Here]', bold=True)
add_para(doc, '● Interaction Video (4-minute interaction with Proprietor in Hindi): [Click Here]', bold=True)
add_para(doc, '● Official Authorization Letterhead (Signed & Stamped): [Click Here]', bold=True)
add_para(doc, '● Academic Location & Remote Data Collection Clarification Note: [Click Here]', bold=True)

add_heading_styled(doc, '2.5 Problem Statement and Business Objectives', level=1)

add_para(doc, 'Problem Statement 1: Minimizing Stockouts of High-Demand Medicines', bold=True)
add_para(doc, 'Intuition-based purchasing causes stockouts of high-demand items (e.g., Ceroxim CV 500, Abzolid 600), leading to lost revenue. Objective: Apply ABC/Pareto classification to maintain 15-day safety buffer stocks for Category A items.')

add_para(doc, 'Problem Statement 2: Preventing Capital Loss from Unsold Expired Stock', bold=True)
add_para(doc, 'Slow-moving SKUs risk expiring unsold, causing direct capital loss. Objective: Execute FSN classification to flag Non-Moving SKUs for credit returns to distributors 60 days before expiration.')

add_para(doc, 'Problem Statement 3: Optimizing Seasonal Procurement via Demand Forecasting', bold=True)
add_para(doc, 'Monthly sales fluctuate seasonally (May 2026 revenue surged 23.8% to ₹2,43,998). Objective: Apply 3-Month Moving Average forecasting to adjust procurement 15-20% in advance of peak seasons.')

doc.add_page_break()

# ============================================================
# SECTION 3: METADATA & DESCRIPTIVE STATISTICS (Page 5)
# ============================================================
add_heading_styled(doc, '3. Metadata and Descriptive Statistics', level=1)
add_heading_styled(doc, '3.1 Metadata Definition and Dataset Structure', level=2)

add_para(doc, 'The dataset covers 6 monthly sheets (JAN2026 to JUN2026) across 30 SKUs. Compact metadata structure:')

meta_data = [
    ['Sheet Name', 'Columns', 'Data Type', 'Units', 'Description'],
    ["1. 'JAN2026'\n2. 'FEB2026'\n3. 'MAR2026'\n4. 'APR2026'\n5. 'MAY2026'\n6. 'JUN2026'", 
     'Medicine Name', 'string', '-', 'Brand / generic name of the medicine'],
    ['', 'Category', 'string', '-', 'Therapeutic category (Antibiotic, Gastric Care, etc.)'],
    ['', 'MRP', 'float', 'Rupees', 'Maximum retail selling price per strip/unit'],
    ['', 'Purchase Price', 'float', 'Rupees', 'Wholesale procurement cost per strip/unit'],
    ['', 'Opening Stock', 'integer', 'strips', 'Number of strips in stock at beginning of month'],
    ['', 'Purchase Quantity', 'integer', 'strips', 'Number of strips purchased during the month'],
    ['', 'Sale Quantity', 'integer', 'strips', 'Number of strips sold during the month'],
    ['', 'Closing Stock', 'integer', 'strips', 'Number of strips in stock at end of month'],
    ['', 'Total Revenue', 'float', 'Rupees', 'Gross revenue generated (Sale Quantity x MRP)'],
    ['', 'Expiry Date', 'string', 'MM/YYYY', 'Expiration date printed on medicine packaging'],
]

t_meta = doc.add_table(rows=len(meta_data), cols=5)
t_meta.style = 'Table Grid'
t_meta.alignment = WD_TABLE_ALIGNMENT.CENTER

for i, row_data in enumerate(meta_data):
    for j, cell_text in enumerate(row_data):
        cell = t_meta.cell(i, j)
        make_table_cell_formatted(cell, cell_text, bold=(i == 0), font_size=9, bg_color='1a73e8' if i == 0 else None)
        if i == 0:
            for run in cell.paragraphs[0].runs: run.font.color.rgb = RGBColor(255, 255, 255)

first_col_cell = t_meta.cell(1, 0)
for r in range(2, len(meta_data)): first_col_cell.merge(t_meta.cell(r, 0))
first_col_cell.text = ''
p = first_col_cell.paragraphs[0]
r = p.add_run("1. 'JAN2026'\n2. 'FEB2026'\n3. 'MAR2026'\n4. 'APR2026'\n5. 'MAY2026'\n6. 'JUN2026'")
r.font.name = 'Times New Roman'
r.font.size = Pt(8.5)
p.alignment = WD_ALIGN_PARAGRAPH.CENTER

add_para(doc, 'Table 1: Metadata Definition — Compact Monthly Dataset Structure', italic=True, font_size=9.5, alignment=WD_ALIGN_PARAGRAPH.CENTER)

add_heading_styled(doc, '3.2 Descriptive Statistics and Quantitative Summary', level=2)

stats_data = [
    ['Statistical Metric', 'MRP (₹)', 'Purchase Price (₹)', '6M Sales Qty (Strips)', '6M Total Revenue (₹)'],
    ['Sample Count (N)', '30', '30', '30', '30'],
    ['Mean (Average)', '₹192.30', '₹134.87', '223.40', '₹43,521.83'],
    ['Median', '₹179.00', '₹124.00', '225.00', '₹41,108.50'],
    ['Standard Deviation', '₹110.75', '₹80.00', '35.68', '₹25,890.82'],
    ['Minimum', '₹45.00', '₹28.00', '152', '₹7,605.00'],
    ['Maximum', '₹480.00', '₹340.00', '317', '₹1,15,680.00'],
    ['Range (Max - Min)', '₹435.00', '₹312.00', '165', '₹1,08,075.00'],
]

t_stat = doc.add_table(rows=len(stats_data), cols=5)
t_stat.style = 'Table Grid'
t_stat.alignment = WD_TABLE_ALIGNMENT.CENTER

for i, row_data in enumerate(stats_data):
    for j, cell_text in enumerate(row_data):
        cell = t_stat.cell(i, j)
        make_table_cell_formatted(cell, cell_text, bold=(i == 0), font_size=9, bg_color='1a73e8' if i == 0 else None)
        if i == 0:
            for run in cell.paragraphs[0].runs: run.font.color.rgb = RGBColor(255, 255, 255)

add_para(doc, 'Table 2: Descriptive Statistics — 30 Medicine SKUs (Jan–June 2026)', italic=True, font_size=9.5, alignment=WD_ALIGN_PARAGRAPH.CENTER)

add_para(doc, 'Key Insights: 1. Mean revenue (₹43,521.83) > Median (₹41,108.50) confirms right-skewed revenue distribution where top SKUs drive majority sales. 2. Sales volume ranges from 152 to 317 strips (SD: 35.68), justifying FSN velocity segmentation. 3. Average gross profit margin per strip is ₹57.43 (~29.8%).', italic=True)

doc.add_page_break()

# ============================================================
# SECTION 4: DATA CLEANING & METHODOLOGY (Page 6)
# ============================================================
add_heading_styled(doc, '4. Detailed Explanation of Analysis Process / Method', level=1)
add_heading_styled(doc, '4.1 Data Cleaning and Preprocessing', level=2)
add_para(doc, '1. Product Standardization: Variant brand strings across monthly bills (e.g., "Ceroxim CV", "CEROXIM-500") were merged into unified standard SKU identifiers.')
add_para(doc, '2. Wholesale Rate Imputation: Missing purchase rates were backfilled using master distributor tax invoices from M/s. Sonu Pharma, Patna.')
add_para(doc, '3. Mathematical Balance Reconciliation: Every row verified via Closing Stock = Opening Stock + Purchase Qty - Sales Qty.')
add_para(doc, '4. Expiry Date Standardization: Dates formatted to MM/YYYY objects for shelf-life countdown modeling.')

add_heading_styled(doc, '4.2 Analytical Methodologies and Mathematical Abstractions', level=2)
add_para(doc, 'Method 1: ABC Analysis — Periodic Consumption Value PCV = Sales Qty x Purchase Price. Ranked by cumulative revenue: Category A (top 70%), Category B (next 20%), Category C (bottom 10%). Linkage: Identifies priority buffer SKUs (PS1).')
add_para(doc, 'Method 2: FSN Analysis — 6-month sales velocity: Fast (Sales >= 250), Slow (190 <= Sales < 250), Non-Moving (Sales < 190). Linkage: Flags high-risk expiring inventory for returns (PS2).')
add_para(doc, 'Method 3: Pareto Analysis (80/20 Rule) — Identifies minimal SKU subset driving 80% revenue. Linkage: Focuses purchasing capital on vital few SKUs.')
add_para(doc, 'Method 4: 3-Month Moving Average Demand Forecast — F_(t+1) = (S_t + S_(t-1) + S_(t-2)) / 3. Linkage: Predicts seasonal demand shifts for advance ordering (PS3).')

doc.add_page_break()

# ============================================================
# SECTION 5: RESULTS & FINDINGS (Pages 7, 8, 9, 10)
# ============================================================
add_heading_styled(doc, '5. Results and Findings (Preliminary Insights)', level=1)

# 5.1 ABC & Pareto
add_heading_styled(doc, '5.1 ABC Classification & Pareto Analysis Findings', level=2)
add_para(doc, 'Total 6-month store revenue reached ₹13,05,655.00 across 30 SKUs. ABC breakdown: Category A (14 SKUs, 46.7% items) = ₹9,00,660.00 (69.0% revenue); Category B (8 SKUs) = ₹2,72,427.00 (20.9%); Category C (8 SKUs) = ₹1,32,568.00 (10.2%). Top 3 SKUs: Ceroxim CV 500 (₹1,15,680), Abzolid 600 (₹92,568), Chyawanprash 2KG (₹98,100).')

add_figure_compact(doc, os.path.join(CHARTS_DIR, 'Fig2_ABC_Donut_Chart.png'), 'Figure 1: ABC Revenue & Item Distribution', width=4.6)

add_para(doc, 'Pareto 80/20 Cutoff: Exactly 18 out of 30 SKUs (60% items) generate 80% of cumulative store revenue, providing a clear cutoff for inventory capital allocation.')

add_figure_compact(doc, os.path.join(CHARTS_DIR, 'Fig3_Pareto_Chart.png'), 'Figure 2: Pareto 80/20 Revenue Contribution', width=4.6)

doc.add_page_break()

# 5.2 FSN & Matrix
add_heading_styled(doc, '5.2 FSN Inventory Movement & Expiry Risk Findings', level=2)
add_para(doc, 'FSN Breakdown: Fast Moving F (5 SKUs, 16.7%), Slow Moving S (20 SKUs, 66.7%), Non-Moving N (5 SKUs, 16.7%). 5 Non-Moving SKUs carrying high expiry risk: Gentalab 30ml Inj (169 units), Dexalife Inj 30ml (208 units), Moxib D Eye Drop (152 units), Alamin Plus Cap (170 units), Dolo Pain Spray (168 units).')

add_figure_compact(doc, os.path.join(CHARTS_DIR, 'Fig4_FSN_Analysis.png'), 'Figure 3: FSN Movement & Expiry Risk Items', width=4.6)

add_para(doc, 'ABC-FSN Matrix: Category A - Slow Moving items (10 SKUs) form the largest segment, requiring strict reorder point management to prevent stockouts.')

add_figure_compact(doc, os.path.join(CHARTS_DIR, 'Fig9_ABC_FSN_Matrix.png'), 'Figure 4: ABC-FSN Priority Matrix', width=4.6)

doc.add_page_break()

# 5.3 Monthly Trends & Category Revenue
add_heading_styled(doc, '5.3 Monthly Sales Trends & Seasonal Demand Fluctuations', level=2)
add_para(doc, 'Monthly Store Revenue: Jan (₹2,20,925), Feb (₹2,14,824), Mar (₹2,07,518), Apr (₹1,97,038 - lowest), May (₹2,43,998 - 23.8% pre-monsoon surge), Jun (₹2,21,352). Antibiotics (₹2,87,988) and Gastric Care (₹1,63,259) drive >34% of revenue.')

add_figure_compact(doc, os.path.join(CHARTS_DIR, 'Fig1_Monthly_Revenue_Trend.png'), 'Figure 5: Monthly Revenue & Sales Volume Trend', width=4.6)
add_figure_compact(doc, os.path.join(CHARTS_DIR, 'Fig6_Category_Revenue.png'), 'Figure 6: Revenue by Therapeutic Category', width=4.6)

doc.add_page_break()

# 5.4 Forecasting & Bubble Chart
add_heading_styled(doc, '5.4 Moving Average Demand Forecast Results', level=2)
add_para(doc, 'July 2026 Demand Forecasts (3-Month SMA): Ceroxim CV 500 (45 strips), Abzolid 600 (37 strips), Chyawanprash 2KG (32 units - seasonal drop), Chymotas Forte (26 strips). SMA provides a conservative, lag-adjusted baseline for advance ordering.')

add_figure_compact(doc, os.path.join(CHARTS_DIR, 'Fig5_Moving_Average_Forecast.png'), 'Figure 7: 3-Month Moving Average Demand Forecast (Top SKUs)', width=4.6)
add_figure_compact(doc, os.path.join(CHARTS_DIR, 'Fig7_Sales_vs_Revenue_Bubble.png'), 'Figure 8: Quantity vs Revenue Bubble Chart (Size = MRP)', width=4.6)

doc.add_page_break()

# ============================================================
# SECTION 6: RECOMMENDATIONS & INTERPRETATION (Page 11)
# ============================================================
add_heading_styled(doc, '6. Interpretation of Results and Recommendations', level=1)
add_para(doc, '1. Category A Safety Buffer Stocking (Addresses PS1): Maintain a 15-day safety buffer for Category A SKUs (Ceroxim CV, Abzolid) using Reorder Point = Avg Daily Sales x Lead Time + Buffer.')
add_para(doc, '2. Pre-Expiry Credit Returns (Addresses PS2): Implement a 60-day pre-expiry audit to return 5 Non-Moving SKUs (Gentalab Inj, Dexalife Inj) to distributor M/s. Sonu Pharma for credit notes.')
add_para(doc, '3. Seasonal Advance Ordering (Addresses PS3): Increase Antibiotic and Gastric orders by 15-20% in April to capture the May-June pre-monsoon demand surge.')
add_para(doc, '4. Digital Stock Audit Transition: Transition from manual registers to weekly Excel digital tracking for automated low-stock alerts.')

add_figure_compact(doc, os.path.join(CHARTS_DIR, 'Fig8_Profit_Margin.png'), 'Figure 9: Profit Margin Analysis across 30 SKUs', width=4.6)
add_figure_compact(doc, os.path.join(CHARTS_DIR, 'Fig10_Stock_Movement.png'), 'Figure 10: Stock Movement Pattern (Top 5 SKUs)', width=4.6)

add_page_numbers(doc)

# Save to both target files
file1 = 'c:\\Users\\HP\\Downloads\\BDM\\Mid term.docx'
file2 = 'c:\\Users\\HP\\Downloads\\BDM\\23f1000204ds.study.iitm.ac.in - Mid Term.docx'

doc.save(file1)
doc.save(file2)

print("\n" + "="*60)
print("SUCCESSFULLY STRUCTURED AND SAVED TO BOTH FILES:")
print(f"1. {file1}")
print(f"2. {file2}")
print("="*60)
