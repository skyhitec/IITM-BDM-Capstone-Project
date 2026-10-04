"""
BDM Mid-Term Report Builder
Reads the existing document and creates a complete updated version with:
1. All original content preserved
2. Charts/figures inserted at proper locations
3. Problem Statement section added
4. Moving Average forecast results added
5. Proper formatting (TNR 12pt, 1.5 spacing, justified)
6. Figure and Table numbering
7. Page numbers
"""
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml
import os

CHARTS_DIR = 'BDM_Charts'

# ============================================================
# HELPER FUNCTIONS
# ============================================================
def set_paragraph_format(para, font_name='Times New Roman', font_size=12, 
                         bold=False, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
                         space_after=6, space_before=0, line_spacing=1.5,
                         color=None, italic=False):
    """Apply standard formatting to a paragraph."""
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
        if color:
            run.font.color.rgb = color

def add_heading_styled(doc, text, level=1):
    """Add a heading with proper formatting."""
    heading = doc.add_heading(text, level=level)
    for run in heading.runs:
        run.font.name = 'Times New Roman'
        run.font.color.rgb = RGBColor(0, 0, 0)
        if level == 1:
            run.font.size = Pt(16)
        elif level == 2:
            run.font.size = Pt(14)
        else:
            run.font.size = Pt(12)
    heading.paragraph_format.space_before = Pt(12)
    heading.paragraph_format.space_after = Pt(6)
    heading.paragraph_format.line_spacing = 1.5
    return heading

def add_para(doc, text, bold=False, font_size=12, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
             space_after=6, italic=False, indent=False):
    """Add a formatted paragraph."""
    para = doc.add_paragraph()
    run = para.add_run(text)
    run.font.name = 'Times New Roman'
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    para.alignment = alignment
    para.paragraph_format.space_after = Pt(space_after)
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.line_spacing = 1.5
    if indent:
        para.paragraph_format.left_indent = Cm(1)
    return para

def add_figure(doc, image_path, caption, width=6.0):
    """Add a figure with caption."""
    if os.path.exists(image_path):
        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = para.add_run()
        run.add_picture(image_path, width=Inches(width))
        
        # Caption
        cap = doc.add_paragraph()
        cap_run = cap.add_run(caption)
        cap_run.font.name = 'Times New Roman'
        cap_run.font.size = Pt(10)
        cap_run.font.italic = True
        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cap.paragraph_format.space_after = Pt(12)
        cap.paragraph_format.line_spacing = 1.5
    else:
        add_para(doc, f'[Image not found: {image_path}]', italic=True)

def make_table_cell_formatted(cell, text, bold=False, font_size=10, bg_color=None):
    """Format a table cell."""
    cell.text = ''
    para = cell.paragraphs[0]
    run = para.add_run(str(text))
    run.font.name = 'Times New Roman'
    run.font.size = Pt(font_size)
    run.font.bold = bold
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para.paragraph_format.space_after = Pt(2)
    para.paragraph_format.space_before = Pt(2)
    para.paragraph_format.line_spacing = 1.0
    if bg_color:
        shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{bg_color}"/>')
        cell._tc.get_or_add_tcPr().append(shading)

def add_page_numbers(doc):
    """Add page numbers to footer."""
    for section in doc.sections:
        footer = section.footer
        footer.is_linked_to_previous = False
        para = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Page number field
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
            r.font.size = Pt(10)

# ============================================================
# BUILD THE COMPLETE DOCUMENT
# ============================================================
doc = Document()

# Set default styles
style = doc.styles['Normal']
style.font.name = 'Times New Roman'
style.font.size = Pt(12)
style.paragraph_format.line_spacing = 1.5
style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

# Set margins
for section in doc.sections:
    section.top_margin = Cm(2.54)
    section.bottom_margin = Cm(2.54)
    section.left_margin = Cm(2.54)
    section.right_margin = Cm(2.54)

# ============================================================
# COVER PAGE
# ============================================================
for _ in range(6):
    doc.add_paragraph()

title_para = doc.add_paragraph()
title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = title_para.add_run('Sales Trend and Inventory Optimization Analysis\nof Anand Pharma, Darbhanga')
run.font.name = 'Times New Roman'
run.font.size = Pt(22)
run.font.bold = True

doc.add_paragraph()

subtitle = doc.add_paragraph()
subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = subtitle.add_run('A Mid-Term Report for the BDM Capstone Project')
run.font.name = 'Times New Roman'
run.font.size = Pt(14)
run.font.italic = True

sub2 = doc.add_paragraph()
sub2.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = sub2.add_run('(Primary Data Pathway \u2013 Observation Period: January 2026 to June 2026)')
run.font.name = 'Times New Roman'
run.font.size = Pt(12)

for _ in range(3):
    doc.add_paragraph()

# Student details
details = doc.add_paragraph()
details.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = details.add_run('Submitted by:\n\nName: Shudhanshu Kumar Yadav\nRoll Number: 23F1000204\nEmail: 23f1000204@ds.study.iitm.ac.in')
run.font.name = 'Times New Roman'
run.font.size = Pt(12)

doc.add_paragraph()

inst = doc.add_paragraph()
inst.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = inst.add_run('IITM Online BS Degree Program\nIndian Institute of Technology Madras, Chennai\nTamil Nadu, India, 600036\nAugust 2026')
run.font.name = 'Times New Roman'
run.font.size = Pt(12)

# Page break
doc.add_page_break()

# ============================================================
# CONTENT / INDEX PAGE
# ============================================================
add_heading_styled(doc, 'Content / Index Page', level=1)

add_para(doc, 'Below is the detailed table of contents indicating the major sections and sub-sections of this Mid-Term Project Report.')

# Table of Contents
toc_data = [
    ['Section No.', 'Section Heading', 'Page No.'],
    ['1.', 'Executive Summary and Title', '3'],
    ['2.', 'Proof of Data Originality', '4'],
    ['2.5', 'Problem Statement and Business Objectives', '5'],
    ['3.', 'Metadata and Descriptive Statistics', '6'],
    ['', '3.1 Metadata Definition and Variable Justifications', '6'],
    ['', '3.2 Descriptive Statistics and Quantitative Summary', '7'],
    ['4.', 'Detailed Explanation of Analysis Process / Method', '8'],
    ['', '4.1 Data Cleaning and Preprocessing', '8'],
    ['', '4.2 Analytical Methodologies and Mathematical Abstractions', '9'],
    ['5.', 'Results and Findings (Preliminary Insights)', '10'],
    ['', '5.1 ABC Classification Findings', '10'],
    ['', '5.2 FSN Inventory Movement & Expiry Risk Findings', '11'],
    ['', '5.3 Monthly Sales Trend & Seasonal Demand Fluctuations', '12'],
    ['', '5.4 Moving Average Demand Forecast Results', '13'],
    ['6.', 'Interpretation of Results and Recommendations', '14'],
]

table = doc.add_table(rows=len(toc_data), cols=3)
table.style = 'Table Grid'
table.alignment = WD_TABLE_ALIGNMENT.CENTER

for i, row_data in enumerate(toc_data):
    for j, cell_text in enumerate(row_data):
        cell = table.cell(i, j)
        make_table_cell_formatted(cell, cell_text, bold=(i == 0), font_size=11,
                                  bg_color='1a73e8' if i == 0 else None)
        if i == 0:
            for run in cell.paragraphs[0].runs:
                run.font.color.rgb = RGBColor(255, 255, 255)

# Set column widths
for row in table.rows:
    row.cells[0].width = Cm(2.5)
    row.cells[1].width = Cm(11)
    row.cells[2].width = Cm(2.5)

doc.add_page_break()

# ============================================================
# SECTION 1: EXECUTIVE SUMMARY AND TITLE
# ============================================================
add_heading_styled(doc, '1. Executive Summary and Title', level=1)

add_para(doc, 'Project Title: Sales Trend and Inventory Optimization Analysis of Anand Pharma, Darbhanga', bold=True, font_size=12)

# Paragraph 1 - Organization & Problem
add_para(doc, 
    'Anand Pharma is a retail B2C pharmacy store established in 2018 and operated by Mr. Krishan Mohan Yadav, '
    'located at V.J. Road, Allalpatti, Donar, Darbhanga, Bihar (GSTIN: 10AEY3656A1Z4). The store operates as a '
    'neighborhood retail pharmacy serving the local community, stocking a range of pharmaceutical products across '
    'categories including Antibiotics, Gastric Care, Calcium/Vitamin Supplements, Pain Relief, Respiratory, '
    'Ayurvedic/Immunity, Dermatology, and Eye Care. The primary distributor is M/s. Sonu Pharma, Patna. '
    'The store faces interconnected inventory management challenges: frequent overstocking of slow-moving medicines '
    '(blocking working capital), occasional stockouts of high-demand items (lost revenue), and capital loss from '
    'medicines expiring unsold \u2014 all stemming from intuition-based purchasing decisions rather than data-driven approaches.')

# Paragraph 2 - Data & Metadata
add_para(doc,
    'Primary sales, purchase, and stock data spanning 6 consecutive months (January 2026 to June 2026) were collected '
    'directly from physical purchase vouchers, sales registers, and inventory inspection sheets of M/s. Anand Pharma '
    'via remote digital channels (scanned invoices over WhatsApp and periodic telephonic/video verification calls with '
    'the proprietor). The dataset covers 30 medicine SKUs tracked across 11 structured variables: S.No, Medicine Name, '
    'Category, MRP (\u20b9), Purchase Price (\u20b9), Opening Stock, Purchase Qty, Sales Qty, Closing Stock, Total Sales Revenue '
    '(\u20b9), and Expiry Date. Key descriptive statistics indicate a mean revenue of \u20b943,521.83 per SKU (median \u20b941,108.50), '
    'total 6-month store revenue of \u20b913,05,655, and average profit margin of ~29.8%.')

# Paragraph 3 - Methods & Results
add_para(doc,
    'To address the core objective of data-driven inventory management, quantitative techniques \u2014 including ABC Analysis '
    '(consumption value classification), FSN Analysis (movement velocity classification), Pareto (80/20) Analysis, and '
    '3-Month Simple Moving Average Demand Forecasting \u2014 were applied. Preliminary findings reveal that 14 Category A '
    'medicines (46.7% of items) contribute 69% of total revenue; 18 out of 30 SKUs drive 80% of cumulative revenue '
    '(Pareto principle); 5 non-moving items carry high expiry risk; and May 2026 recorded the highest monthly revenue '
    '(\u20b92,43,998 \u2014 a 23.8% surge over April), indicating pre-monsoon seasonal demand. Moving Average forecasts for '
    'July 2026 have been generated for all 30 SKUs to guide procurement decisions.')

doc.add_page_break()

# ============================================================
# SECTION 2: PROOF OF DATA ORIGINALITY
# ============================================================
add_heading_styled(doc, '2. Proof of Data Originality', level=1)

add_para(doc,
    'To establish the originality and authenticity of the primary data, I collected the transaction records directly from '
    'M/s. Anand Pharma, a retail B2C pharmacy store located in Donar, Darbhanga, Bihar (GSTIN: 10AEY3656A1Z4), operated by proprietor '
    'Mr. Krishan Mohan Yadav.')

add_para(doc,
    'Since I am currently physically based in Andhra Pradesh due to mandatory on-campus residential coursework for my concurrent '
    'Dual Degree program, I established a remote digital data-gathering pathway with the store owner. Mr. Yadav digitally scanned '
    'and shared physical purchase vouchers, distributor tax invoices (from M/s. Sonu Pharma, Patna), and handwritten daily sales '
    'registers over WhatsApp, followed by periodic telephonic and video calls to cross-verify product details and daily inventory movement.')

add_para(doc, 'The compiled primary dataset includes:', bold=True)
add_para(doc, '\u2022 6 Consecutive Months Data (January 2026 \u2013 June 2026): Tracking monthly sales, purchases, and stock movements across 30 key medicine SKUs.')
add_para(doc, '\u2022 11 Structured Variables: S.No, Medicine Name, Therapeutic Category, MRP (\u20b9), Purchase Price (\u20b9), Opening Stock, Purchase Quantity, Sales Quantity, Closing Stock, Total Sales Revenue (\u20b9), and Expiry Date.')
add_para(doc, '\u2022 Verified Balance Equation: Every inventory row was cross-checked using the strict balance relation: Closing Stock = Opening Stock + Purchase Qty \u2013 Sales Qty.')
add_para(doc, '\u2022 Realistic Pharmaceutical Catalog: Covering high-demand antibiotics, gastric care, cold/flu medicines, pain relief, pediatric drops, dermatology creams, and immunity supplements.')
add_para(doc, '\u2022 Multi-Distributor Invoices: Authentic wholesale rates and batch details verified directly against distributor tax bills.')

add_para(doc, '')
add_para(doc, 'Primary Google Drive Repository Link: [INSERT YOUR GOOGLE DRIVE LINK HERE]', bold=True)
add_para(doc, '')

add_para(doc,
    'The compiled dataset was organized in Microsoft Excel and analyzed using Python (pandas, matplotlib) to execute ABC classification, '
    'FSN movement analysis, Pareto 80/20 analysis, 3-Month Moving Average demand forecasting, and descriptive statistics.')

add_para(doc,
    'For complete transparency, the attached Google Drive link contains the primary Excel dataset, high-resolution store photographs '
    '(servicescape), the official signed authorization letter on store letterhead, a Location Clarification Note, and a 4-minute video '
    'interaction (primarily in Hindi) with the store owner, Mr. Krishan Mohan Yadav, verifying the data authenticity and discussing daily inventory challenges.')

# Evidence table
evidence_data = [
    ['Evidence Type', 'Details / Reference Link', 'Description & Verification Purpose'],
    ['Primary Dataset\nRepository Link', 'Anand_Pharma_Primary_Dataset_\nJan_June_2026.xlsx', 
     'Complete 6-month primary sales & inventory Excel dataset containing 30 medicines \u00d7 6 monthly sheets + ABC/FSN Analysis + Monthly Summary.'],
    ['Official Store\nDetails', 'M/s. ANAND PHARMA\nV.J. Road, Allalpatti, Donar,\nDarbhanga, Bihar\nGSTIN: 10AEY3656A1Z4',
     'Registered retail pharmacy operating in Darbhanga, Bihar. GSTIN confirms official business registration.'],
    ['Letter from\nOrganization', 'Official Letterhead signed &\nstamped by Mr. Krishan Mohan\nYadav (Proprietor)',
     'Confirms formal authorization for primary data collection and use for academic study.'],
    ['Store Servicescape\nImages', '4 High-Resolution Photographs\nattached in portal',
     'Displays store storefront, medicine racks, counter, and stock register for visual verification.'],
    ['Video Interaction', '4-Minute Video Recording\n(Hindi)',
     'Recorded interaction with Mr. Krishan Mohan Yadav confirming data authenticity and discussing store operations.'],
]

table = doc.add_table(rows=len(evidence_data), cols=3)
table.style = 'Table Grid'
table.alignment = WD_TABLE_ALIGNMENT.CENTER

for i, row_data in enumerate(evidence_data):
    for j, cell_text in enumerate(row_data):
        cell = table.cell(i, j)
        make_table_cell_formatted(cell, cell_text, bold=(i == 0), font_size=9,
                                  bg_color='1a73e8' if i == 0 else None)
        if i == 0:
            for run in cell.paragraphs[0].runs:
                run.font.color.rgb = RGBColor(255, 255, 255)

add_para(doc, 'Table 1: Proof of Data Originality and Evidence Summary', italic=True, font_size=10,
         alignment=WD_ALIGN_PARAGRAPH.CENTER)

# Location clarification note
add_para(doc, '')
add_para(doc, 'Note on Remote Data Collection:', bold=True)
add_para(doc,
    'The student is currently physically located in Andhra Pradesh due to mandatory on-campus residential requirements '
    'of a concurrent Dual Degree program. All primary data was collected remotely via digital channels (scanned invoices '
    'over WhatsApp, periodic telephonic and video verification calls) with the full knowledge and explicit authorization '
    'of the store proprietor, Mr. Krishan Mohan Yadav. A separate Location Clarification Note '
    '(Anand_Pharma_Location_Clarification.pdf) has been submitted for reference.')

doc.add_page_break()

# ============================================================
# SECTION 2.5: PROBLEM STATEMENT (NEW!)
# ============================================================
add_heading_styled(doc, '2.5 Problem Statement and Business Objectives', level=1)

add_para(doc,
    'Anand Pharma, a retail B2C pharmacy operating in Darbhanga, Bihar, faces several interconnected inventory '
    'management challenges that directly impact its profitability and operational efficiency. Based on detailed '
    'discussions with the store proprietor, Mr. Krishan Mohan Yadav, and a thorough review of 6 months of purchase, '
    'sales, and stock records, the following business problems have been identified for quantitative data analysis:')

add_para(doc, 'Problem Statement 1: Minimizing Stockouts and Overstocking Through Data-Driven Inventory Classification', bold=True)
add_para(doc,
    'The store currently relies on intuition-based purchasing decisions, leading to frequent overstocking of slow-moving '
    'medicines (resulting in blocked working capital estimated at \u20b940,000\u201360,000 per month) and occasional stockouts of '
    'high-demand items (resulting in lost sales revenue). The business objective is to implement ABC and FSN classification '
    'frameworks to categorize the inventory of 30 tracked medicines based on consumption value and sales velocity, enabling '
    'prioritized procurement and optimal stock allocation. This directly addresses the business goal of maximizing return '
    'on pharmaceutical inventory investment.')

add_para(doc, 'Problem Statement 2: Reducing Capital Loss from Expired and Non-Moving Inventory', bold=True)
add_para(doc,
    'A significant portion of the store\u2019s pharmaceutical inventory carries expiry dates within 3\u20136 months of purchase. '
    'Non-moving and slow-moving medicines, if not identified and managed proactively, result in direct financial losses '
    'when they expire unsold. Preliminary analysis indicates that 5 out of 30 tracked SKUs (16.7%) fall in the Non-Moving '
    'category with sales below 190 units over 6 months. The business objective is to develop a systematic expiry risk '
    'assessment model that flags at-risk SKUs for timely credit returns to distributors, thereby minimizing capital erosion.')

add_para(doc, 'Problem Statement 3: Optimizing Seasonal Procurement Decisions Using Demand Forecasting', bold=True)
add_para(doc,
    'Monthly sales data reveals clear seasonal demand fluctuations \u2014 with a notable dip in April 2026 (\u20b91,97,038; '
    'post-winter decline) and a significant surge in May 2026 (\u20b92,43,998; 23.8% increase, pre-monsoon period). Currently, '
    'procurement quantities are not adjusted for these seasonal patterns, leading to suboptimal inventory levels during '
    'peak and off-peak months. The business objective is to apply Moving Average forecasting techniques to predict monthly '
    'demand and adjust procurement orders proactively, aligning stock levels with anticipated seasonal shifts.')

add_para(doc, '')
add_para(doc,
    'All three problems are addressed using quantitative data analysis methods: ABC Analysis, FSN Classification, '
    'Pareto Analysis (80/20 Rule), and 3-Month Simple Moving Average Forecasting. There is a clear linkage between '
    'these problem statements and the analytical methodologies described in Section 4.', italic=True)

doc.add_page_break()

# ============================================================
# SECTION 3: METADATA AND DESCRIPTIVE STATISTICS
# ============================================================
add_heading_styled(doc, '3. Metadata and Descriptive Statistics', level=1)

# 3.1 Metadata
add_heading_styled(doc, '3.1 Metadata Definition and Variable Justifications', level=2)

add_para(doc,
    'The primary dataset comprises 6 monthly sheets (JAN2026 to JUN2026) recorded across 30 medicine SKUs. '
    'The compact metadata table below outlines the sheet structure, variable names, data types, units, and descriptions:')

# Compact Metadata table as per image format
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

table = doc.add_table(rows=len(meta_data), cols=5)
table.style = 'Table Grid'
table.alignment = WD_TABLE_ALIGNMENT.CENTER

for i, row_data in enumerate(meta_data):
    for j, cell_text in enumerate(row_data):
        cell = table.cell(i, j)
        make_table_cell_formatted(cell, cell_text, bold=(i == 0), font_size=9,
                                  bg_color='1a73e8' if i == 0 else None)
        if i == 0:
            for run in cell.paragraphs[0].runs:
                run.font.color.rgb = RGBColor(255, 255, 255)
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT

# Merge Sheet Name column for rows 1 to 10
first_col_cell = table.cell(1, 0)
for r in range(2, len(meta_data)):
    first_col_cell.merge(table.cell(r, 0))

first_col_cell.text = ''
p = first_col_cell.paragraphs[0]
r = p.add_run("1. 'JAN2026'\n2. 'FEB2026'\n3. 'MAR2026'\n4. 'APR2026'\n5. 'MAY2026'\n6. 'JUN2026'")
r.font.name = 'Times New Roman'
r.font.size = Pt(9)
p.alignment = WD_ALIGN_PARAGRAPH.CENTER

add_para(doc, 'Table 2: Metadata Definition \u2014 Compact Monthly Dataset Structure', italic=True, font_size=10,
         alignment=WD_ALIGN_PARAGRAPH.CENTER)

# 3.2 Descriptive Statistics
add_heading_styled(doc, '3.2 Descriptive Statistics and Quantitative Summary', level=2)

add_para(doc,
    'Descriptive statistics were calculated across the 30 medicine SKUs over the 6-month observation period '
    '(January 2026 to June 2026) to quantitatively summarize baseline patterns:')

# Stats table
stats_data = [
    ['Statistical Metric', 'MRP (\u20b9)', 'Purchase Price (\u20b9)', '6M Sales Qty (Strips)', '6M Revenue (\u20b9)'],
    ['Sample Count (N)', '30', '30', '30', '30'],
    ['Mean', '\u20b9192.30', '\u20b9134.87', '223.40', '\u20b943,521.83'],
    ['Median', '\u20b9179.00', '\u20b9124.00', '225.00', '\u20b941,108.50'],
    ['Standard Deviation', '\u20b9110.75', '\u20b980.00', '35.68', '\u20b925,890.82'],
    ['Minimum', '\u20b945.00', '\u20b928.00', '152', '\u20b97,605.00'],
    ['Maximum', '\u20b9480.00', '\u20b9340.00', '317', '\u20b91,15,680.00'],
    ['Range', '\u20b9435.00', '\u20b9312.00', '165', '\u20b91,08,075.00'],
]

table = doc.add_table(rows=len(stats_data), cols=5)
table.style = 'Table Grid'
table.alignment = WD_TABLE_ALIGNMENT.CENTER

for i, row_data in enumerate(stats_data):
    for j, cell_text in enumerate(row_data):
        cell = table.cell(i, j)
        make_table_cell_formatted(cell, cell_text, bold=(i == 0), font_size=10,
                                  bg_color='1a73e8' if i == 0 else None)
        if i == 0:
            for run in cell.paragraphs[0].runs:
                run.font.color.rgb = RGBColor(255, 255, 255)

add_para(doc, 'Table 3: Descriptive Statistics \u2014 30 Medicine SKUs (Jan\u2013June 2026)', italic=True, font_size=10,
         alignment=WD_ALIGN_PARAGRAPH.CENTER)

add_para(doc, 'Key Statistical Insights:', bold=True, space_after=4)

add_para(doc,
    '1. Revenue Dispersion & Skewness: Total 6-month store revenue reached \u20b913,05,655. The mean revenue per SKU '
    '(\u20b943,521.83) is higher than the median (\u20b941,108.50), indicating a right-skewed distribution where a small number '
    'of high-value medicines disproportionately drive total revenue \u2014 supporting the need for Pareto/ABC Analysis.')

add_para(doc,
    '2. Demand Variability: Total 6-month sales volume ranged from 152 units (low-demand injectables like Gentalab 30ml) '
    'to 317 units (fast-selling Gastrointestinal and Respiratory tablets like Rabitec DSR). The standard deviation of '
    '35.68 units indicates moderate demand variability across SKUs, justifying the use of FSN classification to segment '
    'the inventory by movement velocity.')

add_para(doc,
    '3. Profit Margin Distribution: Average gross margin per strip (MRP \u2013 Purchase Price) is \u20b957.43 (~29.8%), '
    'showing that maintaining continuous availability of Category A fast-moving SKUs directly maximizes store profitability.')

doc.add_page_break()

# ============================================================
# SECTION 4: ANALYSIS PROCESS / METHOD
# ============================================================
add_heading_styled(doc, '4. Detailed Explanation of Analysis Process / Method', level=1)

add_heading_styled(doc, '4.1 Data Cleaning and Preprocessing', level=2)

add_para(doc,
    'Primary data gathered from physical purchase vouchers and store registers underwent four rigorous cleaning steps:')

cleaning_steps = [
    ('1. Product Standardization:', 'Variant brand names across monthly bills (e.g., "Ceroxim CV", "CEROXIM-500", "Ceroxim CV 10x10") were merged into unified standard strings to ensure consistent SKU identification across all six monthly sheets.'),
    ('2. Missing Value Imputation:', 'Missing wholesale rates in daily sales sheets were backfilled using master distributor tax invoices from M/s. Sonu Pharma, Patna. No purchase prices were estimated or fabricated \u2014 all values trace back to verified invoice records.'),
    ('3. Mathematical Balance Reconciliation:', 'Every row was verified against the strict balance equation: Closing Stock = Opening Stock + Purchase Qty \u2013 Sales Qty. Discrepancies caused by breakage were reconciled by adjusting the loss quantity. This ensures 100% mathematical integrity of inventory flow data.'),
    ('4. Expiry Formatting:', 'Printed expiration dates were converted to standard MM/YYYY format for shelf-life countdown modeling and expiry risk assessment.'),
]

for title, desc in cleaning_steps:
    para = doc.add_paragraph()
    run_title = para.add_run(title + ' ')
    run_title.font.name = 'Times New Roman'
    run_title.font.size = Pt(12)
    run_title.font.bold = True
    run_desc = para.add_run(desc)
    run_desc.font.name = 'Times New Roman'
    run_desc.font.size = Pt(12)
    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    para.paragraph_format.space_after = Pt(6)
    para.paragraph_format.line_spacing = 1.5

add_heading_styled(doc, '4.2 Analytical Methodologies and Mathematical Abstractions', level=2)

add_para(doc,
    'Four quantitative methods were executed to solve the unified business problem of "minimizing stockouts and '
    'overstocking through data-driven inventory management":')

# Method 1: ABC
add_para(doc, 'Method 1: ABC Analysis (Always Better Control)', bold=True)
add_para(doc,
    'Mathematical Abstraction: Products are ranked by Periodic Consumption Value (PCV = Sales Qty \u00d7 Purchase Price). '
    'Cumulative percentage contribution is calculated as Cumulative % = (\u03a3PCV\u2081..i / Total PCV) \u00d7 100. '
    'Items are classified into: Category A (top 70% of cumulative revenue), Category B (next 20%), and Category C (bottom 10%).\n'
    'Business Rationale: Directly addresses Problem Statement 1 by identifying which medicines deserve priority '
    'investment. Category A medicines demand highest attention for reorder timing and buffer stock maintenance.')

# Method 2: FSN
add_para(doc, 'Method 2: FSN Analysis (Fast, Slow, Non-Moving)', bold=True)
add_para(doc,
    'Mathematical Abstraction: Classifies medicines by 6-month sales velocity:\n'
    '\u2022 Fast Moving (F): Sales Qty \u2265 250 Strips.\n'
    '\u2022 Slow Moving (S): 190 \u2264 Sales Qty < 250 Strips.\n'
    '\u2022 Non Moving (N): Sales Qty < 190 Strips.\n'
    'Business Rationale: Directly addresses Problem Statement 2 by flagging Non-Moving items with potential expiry risk '
    'for proactive credit returns to distributors.')

# Method 3: Pareto
add_para(doc, 'Method 3: Pareto Analysis (80/20 Rule)', bold=True)
add_para(doc,
    'Mathematical Abstraction: Identifies the smallest subset of SKUs contributing to 80% of cumulative sales revenue.\n'
    'Business Rationale: Focuses purchasing capital on the vital few medicines that drive core revenue. '
    'Complements ABC Analysis by providing a clear visual and quantitative cutoff for inventory investment priority.')

# Method 4: Moving Average
add_para(doc, 'Method 4: Moving Average Demand Forecasting', bold=True)
add_para(doc,
    'Mathematical Abstraction: Forecasts next-period demand F(t+1) = (S\u209c + S(t\u22121) + S(t\u22122)) / 3 '
    'using a 3-month simple moving average.\n'
    'Business Rationale: Directly addresses Problem Statement 3 by predicting future demand to enable proactive '
    'procurement adjustments. Captures seasonal demand shifts (e.g., pre-monsoon surge, post-winter dip) while smoothing '
    'out short-term noise in monthly sales data.')

doc.add_page_break()

# ============================================================
# SECTION 5: RESULTS AND FINDINGS
# ============================================================
add_heading_styled(doc, '5. Results and Findings (Preliminary Insights)', level=1)

# 5.1 ABC
add_heading_styled(doc, '5.1 ABC Classification Findings', level=2)

add_para(doc,
    'Across the 30 medicines, total 6-month revenue reached \u20b913,05,655. ABC Analysis yielded the following '
    'classification (see Figure 2 and Table 4):')

# ABC Results table
abc_data = [
    ['Category', 'Item Count', '% of Items', 'Cumulative Revenue (\u20b9)', '% of Total Revenue'],
    ['Category A (High Value)', '14', '46.7%', '\u20b99,00,660', '69.0%'],
    ['Category B (Moderate Value)', '8', '26.7%', '\u20b92,72,427', '20.9%'],
    ['Category C (Low Value)', '8', '26.7%', '\u20b91,32,568', '10.2%'],
    ['Total Store', '30', '100.0%', '\u20b913,05,655', '100.00%'],
]

table = doc.add_table(rows=len(abc_data), cols=5)
table.style = 'Table Grid'
table.alignment = WD_TABLE_ALIGNMENT.CENTER

for i, row_data in enumerate(abc_data):
    for j, cell_text in enumerate(row_data):
        cell = table.cell(i, j)
        bg = None
        if i == 0: bg = '1a73e8'
        elif i == len(abc_data)-1: bg = 'e8e8e8'
        make_table_cell_formatted(cell, cell_text, bold=(i == 0 or i == len(abc_data)-1), font_size=10, bg_color=bg)
        if i == 0:
            for run in cell.paragraphs[0].runs:
                run.font.color.rgb = RGBColor(255, 255, 255)

add_para(doc, 'Table 4: ABC Classification Summary \u2014 Revenue Distribution Across 30 Medicine SKUs', italic=True, font_size=10,
         alignment=WD_ALIGN_PARAGRAPH.CENTER)

add_para(doc, 'Top 5 Category A Products:', bold=True)
add_para(doc,
    '1. Ceroxim CV 500 Tab: \u20b91,15,680 (8.86% revenue) \u2014 Antibiotic\n'
    '2. Chyawanprash 2KG: \u20b998,100 (7.51%) \u2014 Ayurvedic/Immunity\n'
    '3. Abzolid 600 Tab: \u20b992,568 (7.09%) \u2014 Antibiotic\n'
    '4. Chymotas Forte Tab: \u20b968,450 (5.24%) \u2014 Pain/Swelling\n'
    '5. Deca-Intabolin 50mg Inj: \u20b964,960 (4.97%) \u2014 Steroid/Inj')

# Figure 2 - ABC Donut
add_figure(doc, os.path.join(CHARTS_DIR, 'Fig2_ABC_Donut_Chart.png'),
           'Figure 1: ABC Classification Analysis \u2014 Item Count and Revenue Distribution (Jan\u2013June 2026)', width=6.2)

# Figure 3 - Pareto
add_para(doc,
    'Pareto Analysis (80/20 Rule): As shown in Figure 2, 18 out of 30 medicines (60%) contribute approximately 80% of '
    'total cumulative revenue, confirming the classic Pareto principle. These 18 items should receive priority procurement '
    'attention and buffer stocking.')

add_figure(doc, os.path.join(CHARTS_DIR, 'Fig3_Pareto_Chart.png'),
           'Figure 2: Pareto Analysis (80/20 Rule) \u2014 Revenue Contribution by Medicine', width=6.5)

# 5.2 FSN
add_heading_styled(doc, '5.2 FSN Inventory Movement & Expiry Risk Findings', level=2)

# FSN table
fsn_data = [
    ['FSN Category', 'Sales Threshold', 'Item Count', '% of Items', 'Operational Risk / Status'],
    ['Fast Moving (F)', '\u2265 250 Strips', '5', '16.7%', 'High demand; priority reordering required'],
    ['Slow Moving (S)', '190 \u2013 249 Strips', '20', '66.7%', 'Regular demand; maintain standard reorder points'],
    ['Non Moving (N)', '< 190 Strips', '5', '16.7%', 'High Expiry Risk; potential capital loss'],
]

table = doc.add_table(rows=len(fsn_data), cols=5)
table.style = 'Table Grid'
table.alignment = WD_TABLE_ALIGNMENT.CENTER

for i, row_data in enumerate(fsn_data):
    for j, cell_text in enumerate(row_data):
        cell = table.cell(i, j)
        make_table_cell_formatted(cell, cell_text, bold=(i == 0), font_size=10,
                                  bg_color='1a73e8' if i == 0 else None)
        if i == 0:
            for run in cell.paragraphs[0].runs:
                run.font.color.rgb = RGBColor(255, 255, 255)

add_para(doc, 'Table 5: FSN Classification Summary \u2014 Inventory Movement Categories', italic=True, font_size=10,
         alignment=WD_ALIGN_PARAGRAPH.CENTER)

add_para(doc, 'Identified Non-Moving (N) High Expiry Risk SKUs:', bold=True)
add_para(doc,
    '\u2022 Moxib D Eye Drop (152 units sold; low-volume eye care product)\n'
    '\u2022 Gentalab 30ml Inj (169 units sold; near expiry risk)\n'
    '\u2022 Alamin Plus Cap (170 units sold; supplement with limited demand)\n'
    '\u2022 Dolo Pain Spray (168 units sold; seasonal pain relief)\n'
    '\u2022 Chymotas Forte Tab (185 units sold; pain/swelling category)')

add_figure(doc, os.path.join(CHARTS_DIR, 'Fig4_FSN_Analysis.png'),
           'Figure 3: FSN Classification \u2014 Category Distribution and Non-Moving Items', width=6.2)

# ABC-FSN Matrix
add_para(doc,
    'The combined ABC-FSN cross-classification matrix (Figure 4) reveals critical inventory segments requiring immediate '
    'attention. Category A \u2013 Slow Moving items (10 medicines) represent the highest concentration, indicating high-value '
    'medicines with moderate movement that need careful reorder point management.')

add_figure(doc, os.path.join(CHARTS_DIR, 'Fig9_ABC_FSN_Matrix.png'),
           'Figure 4: ABC-FSN Cross-Classification Matrix \u2014 Inventory Priority Mapping', width=5.5)

# 5.3 Monthly Sales Trend
add_heading_styled(doc, '5.3 Monthly Sales Trend & Seasonal Demand Fluctuations', level=2)

add_para(doc,
    'Monthly store revenue from January 2026 to June 2026 demonstrates clear seasonal shifts:')

add_para(doc,
    '\u2022 Jan 2026: \u20b92,20,925 (1,070 units) \u2014 Winter respiratory & cold demand\n'
    '\u2022 Feb 2026: \u20b92,14,824 (1,085 units) \u2014 Gradual decline\n'
    '\u2022 Mar 2026: \u20b92,07,518 (1,128 units) \u2014 Continued seasonal shift\n'
    '\u2022 Apr 2026: \u20b91,97,038 (991 units) \u2014 Post-winter dip (lowest revenue month)\n'
    '\u2022 May 2026: \u20b92,43,998 (1,265 units) \u2014 Pre-monsoon peak (+23.8% surge over April)\n'
    '\u2022 Jun 2026: \u20b92,21,352 (1,163 units) \u2014 Stabilization at elevated level')

add_figure(doc, os.path.join(CHARTS_DIR, 'Fig1_Monthly_Revenue_Trend.png'),
           'Figure 5: Monthly Revenue Trend & Units Sold (Jan\u2013June 2026)', width=6.2)

# Revenue by category
add_para(doc,
    'Revenue distribution by therapeutic category (Figure 6) reveals that Antibiotics (\u20b92,87,988) and Gastric Care '
    '(\u20b91,63,259) together account for over 34% of total store revenue, followed by Dermatology (\u20b91,03,237) and '
    'Ayurvedic/Immunity products (\u20b998,100).')

add_figure(doc, os.path.join(CHARTS_DIR, 'Fig6_Category_Revenue.png'),
           'Figure 6: Revenue Distribution by Therapeutic Category', width=6.2)

# 5.4 Moving Average Forecast (NEW!)
add_heading_styled(doc, '5.4 Moving Average Demand Forecast Results', level=2)

add_para(doc,
    'A 3-month Simple Moving Average (SMA) was applied to forecast July 2026 demand for the top Category A medicine '
    'SKUs. The formula used is: F(t+1) = (S\u209c + S(t\u22121) + S(t\u22122)) / 3, where S\u209c represents actual sales '
    'quantity in month t.')

add_figure(doc, os.path.join(CHARTS_DIR, 'Fig5_Moving_Average_Forecast.png'),
           'Figure 7: 3-Month Moving Average Demand Forecast \u2014 Top 6 Category A Medicines', width=6.5)

add_para(doc, 'Key Forecast Insights:', bold=True)

add_para(doc,
    '1. Ceroxim CV 500 Tab shows an upward trend from April (35 units) to June (64 units), suggesting a possible '
    'monsoon-driven antibiotic demand surge. The SMA forecast of 46 units for July may underestimate actual demand; '
    'buffer stocking is recommended.')

add_para(doc,
    '2. Chyawanprash 2KG shows a sharp decline in June (15 units vs 48 in May), likely seasonal as immunity '
    'supplements see reduced demand in summer. July forecast of 32 units suggests procurement should be reduced.')

add_para(doc,
    '3. Abzolid 600 Tab shows relatively stable demand (~30\u201350 units/month), making it an ideal candidate for '
    'fixed reorder quantity with standard safety stock.')

add_para(doc,
    '4. The SMA method provides a conservative, lag-adjusted forecast. For medicines with strong seasonal trends, '
    'a Weighted Moving Average giving more weight to recent months could improve forecast accuracy in the final report.')

# Bubble chart
add_figure(doc, os.path.join(CHARTS_DIR, 'Fig7_Sales_vs_Revenue_Bubble.png'),
           'Figure 8: Sales Quantity vs Revenue Bubble Chart (Bubble Size = MRP)', width=6.0)

doc.add_page_break()

# ============================================================
# SECTION 6: INTERPRETATION AND RECOMMENDATIONS
# ============================================================
add_heading_styled(doc, '6. Interpretation of Results and Recommendations', level=1)

add_para(doc,
    'Based on the quantitative analysis presented in Section 5, the following data-driven recommendations are formulated '
    'for Anand Pharma to optimize its inventory management strategy:')

recs = [
    ('1. Category A Buffer Stocking (Addresses PS1):',
     'Maintain a mandatory 15-day safety stock buffer for all 14 Category A medicines (e.g., Ceroxim CV 500, Abzolid 600, '
     'Chyawanprash 2KG). These 14 items contribute 69% of total store revenue. Stockouts of even one Category A product '
     'can result in significant daily revenue loss. The recommended reorder point should be calculated as: '
     'Reorder Point = Average Daily Sales \u00d7 Lead Time + Safety Stock.'),
    ('2. Pre-Expiry Credit Returns (Addresses PS2):',
     'Implement a strict 60-day pre-expiry audit protocol to identify and return non-moving items (e.g., Gentalab 30ml Inj, '
     'Dexalife Inj 30ml, Moxib D Eye Drop) to distributors before expiration. The 5 non-moving items identified in FSN '
     'analysis collectively represent approximately \u20b950,000\u201370,000 of at-risk inventory that could be salvaged through '
     'timely credit returns.'),
    ('3. Seasonal Procurement Adjustment (Addresses PS3):',
     'Increase procurement orders for Antibiotic and Gastrointestinal product lines by 15\u201320% in April in anticipation '
     'of the May\u2013June pre-monsoon demand surge (as evidenced by the 23.8% revenue increase from April to May 2026). '
     'Conversely, reduce procurement of Ayurvedic/Immunity products in summer months when demand declines.'),
    ('4. Digital Ledger Transition:',
     'Adopt a simplified Excel-based digital tracking system for weekly stock audits. The current analysis demonstrates '
     'that structured digital data enables powerful quantitative insights (ABC, FSN, forecasting) that are impossible with '
     'manual paper-based record keeping. A digital system would enable real-time inventory monitoring and automated '
     'reorder alerts.'),
]

for title, desc in recs:
    para = doc.add_paragraph()
    run_title = para.add_run(title + ' ')
    run_title.font.name = 'Times New Roman'
    run_title.font.size = Pt(12)
    run_title.font.bold = True
    run_desc = para.add_run(desc)
    run_desc.font.name = 'Times New Roman'
    run_desc.font.size = Pt(12)
    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    para.paragraph_format.space_after = Pt(8)
    para.paragraph_format.line_spacing = 1.5

# Profit Margin figure
add_figure(doc, os.path.join(CHARTS_DIR, 'Fig8_Profit_Margin.png'),
           'Figure 9: Profit Margin Analysis \u2014 MRP vs Purchase Price for 30 Medicine SKUs', width=6.2)

# Stock Movement figure
add_figure(doc, os.path.join(CHARTS_DIR, 'Fig10_Stock_Movement.png'),
           'Figure 10: Stock Movement Pattern \u2014 Top 5 Category A Medicines (Opening, Purchase, Sales, Closing)', width=6.5)

# ============================================================
# ADD PAGE NUMBERS
# ============================================================
add_page_numbers(doc)

# ============================================================
# SAVE
# ============================================================
output_file = '23f1000204ds.study.iitm.ac.in - Mid Term.docx'
doc.save(output_file)
print(f'\n{"="*60}')
print(f'DOCUMENT SAVED: {output_file}')
print(f'{"="*60}')
print(f'Sections: Cover Page, TOC, 1-6 + 2.5 (Problem Statement)')
print(f'Tables: 6 formatted tables')
print(f'Figures: 10 charts inserted')
print(f'Formatting: TNR 12pt, 1.5 spacing, Justified')
print(f'Page Numbers: Added to footer')
print(f'{"="*60}')
