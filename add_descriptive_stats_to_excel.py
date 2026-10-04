import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.load_workbook('Anand_Pharma_Primary_Dataset_Jan_June_2026.xlsx')

# -------------------------------------------------------------
# 1. Add Summary Block in 'ABC & FSN Analysis' Sheet below row 33
# -------------------------------------------------------------
ws = wb['ABC & FSN Analysis']

# Thin border and fill styles
header_fill = PatternFill(start_color="1A73E8", end_color="1A73E8", fill_type="solid")
header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
bold_font = Font(name="Calibri", size=11, bold=True)
regular_font = Font(name="Calibri", size=11)
thin_border = Border(
    left=Side(style='thin', color='D9D9D9'),
    right=Side(style='thin', color='D9D9D9'),
    top=Side(style='thin', color='D9D9D9'),
    bottom=Side(style='thin', color='D9D9D9')
)

start_row = 36
ws.cell(row=start_row, column=1, value="3.2 Descriptive Statistics & Quantitative Summary (30 Medicine SKUs)").font = Font(name="Calibri", size=12, bold=True, color="1A73E8")

# Headers
headers = ["Statistical Metric", "MRP (₹)", "Purchase Price (₹)", "6M Sales Qty (Strips)", "6M Total Revenue (₹)"]
for c, h in enumerate(headers, 1):
    cell = ws.cell(row=start_row+1, column=c, value=h)
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center", vertical="center")

# Rows with Excel Formulas
stats_specs = [
    ("Sample Count (N)", '=COUNT(D4:D33)', '=COUNT(E4:E33)', '=COUNT(F4:F33)', '=COUNT(G4:G33)', '0', '0'),
    ("Mean (Average)", '=AVERAGE(D4:D33)', '=AVERAGE(E4:E33)', '=AVERAGE(F4:F33)', '=AVERAGE(G4:G33)', '#,##0.00', '₹#,##0.00'),
    ("Median", '=MEDIAN(D4:D33)', '=MEDIAN(E4:E33)', '=MEDIAN(F4:F33)', '=MEDIAN(G4:G33)', '#,##0.00', '₹#,##0.00'),
    ("Standard Deviation", '=STDEV(D4:D33)', '=STDEV(E4:E33)', '=STDEV(F4:F33)', '=STDEV(G4:G33)', '#,##0.00', '₹#,##0.00'),
    ("Minimum", '=MIN(D4:D33)', '=MIN(E4:E33)', '=MIN(F4:F33)', '=MIN(G4:G33)', '#,##0', '₹#,##0.00'),
    ("Maximum", '=MAX(D4:D33)', '=MAX(E4:E33)', '=MAX(F4:F33)', '=MAX(G4:G33)', '#,##0', '₹#,##0.00'),
    ("Range (Max - Min)", '=MAX(D4:D33)-MIN(D4:D33)', '=MAX(E4:E33)-MIN(E4:E33)', '=MAX(F4:F33)-MIN(F4:F33)', '=MAX(G4:G33)-MIN(G4:G33)', '#,##0', '₹#,##0.00')
]

for idx, (label, f_mrp, f_pur, f_sales, f_rev, num_fmt, rev_fmt) in enumerate(stats_specs, start_row+2):
    c1 = ws.cell(row=idx, column=1, value=label)
    c2 = ws.cell(row=idx, column=2, value=f_mrp)
    c3 = ws.cell(row=idx, column=3, value=f_pur)
    c4 = ws.cell(row=idx, column=4, value=f_sales)
    c5 = ws.cell(row=idx, column=5, value=f_rev)
    
    c1.font = bold_font
    for c, cell in enumerate([c1, c2, c3, c4, c5], 1):
        cell.border = thin_border
        if c > 1:
            cell.font = regular_font
            cell.alignment = Alignment(horizontal="right")
            if idx == start_row+2: # Sample Count N
                cell.number_format = '0'
            elif c in (2, 3):
                cell.number_format = '₹#,##0.00'
            elif c == 4:
                cell.number_format = num_fmt
            elif c == 5:
                cell.number_format = rev_fmt

# -------------------------------------------------------------
# 2. Add Dedicated Sheet 'Descriptive Statistics'
# -------------------------------------------------------------
if 'Descriptive Statistics' in wb.sheetnames:
    del wb['Descriptive Statistics']

ws_new = wb.create_sheet(title='Descriptive Statistics')

# Title
ws_new.cell(row=1, column=1, value="Anand Pharma, Darbhanga — Primary Dataset Descriptive Statistics").font = Font(name="Calibri", size=14, bold=True, color="1A73E8")
ws_new.cell(row=2, column=1, value="Observation Period: January 2026 to June 2026 (N = 30 Medicine SKUs)").font = Font(name="Calibri", size=11, italic=True)

# Headers
for c, h in enumerate(headers, 1):
    cell = ws_new.cell(row=4, column=c, value=h)
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center", vertical="center")

# Rows linking to 'ABC & FSN Analysis' sheet
stats_specs_linked = [
    ("Sample Count (N)", "=COUNT('ABC & FSN Analysis'!D4:D33)", "=COUNT('ABC & FSN Analysis'!E4:E33)", "=COUNT('ABC & FSN Analysis'!F4:F33)", "=COUNT('ABC & FSN Analysis'!G4:G33)", '0', '0'),
    ("Mean (Average)", "=AVERAGE('ABC & FSN Analysis'!D4:D33)", "=AVERAGE('ABC & FSN Analysis'!E4:E33)", "=AVERAGE('ABC & FSN Analysis'!F4:F33)", "=AVERAGE('ABC & FSN Analysis'!G4:G33)", '#,##0.00', '₹#,##0.00'),
    ("Median", "=MEDIAN('ABC & FSN Analysis'!D4:D33)", "=MEDIAN('ABC & FSN Analysis'!E4:E33)", "=MEDIAN('ABC & FSN Analysis'!F4:F33)", "=MEDIAN('ABC & FSN Analysis'!G4:G33)", '#,##0.00', '₹#,##0.00'),
    ("Standard Deviation", "=STDEV('ABC & FSN Analysis'!D4:D33)", "=STDEV('ABC & FSN Analysis'!E4:E33)", "=STDEV('ABC & FSN Analysis'!F4:F33)", "=STDEV('ABC & FSN Analysis'!G4:G33)", '#,##0.00', '₹#,##0.00'),
    ("Minimum", "=MIN('ABC & FSN Analysis'!D4:D33)", "=MIN('ABC & FSN Analysis'!E4:E33)", "=MIN('ABC & FSN Analysis'!F4:F33)", "=MIN('ABC & FSN Analysis'!G4:G33)", '#,##0', '₹#,##0.00'),
    ("Maximum", "=MAX('ABC & FSN Analysis'!D4:D33)", "=MAX('ABC & FSN Analysis'!E4:E33)", "=MAX('ABC & FSN Analysis'!F4:F33)", "=MAX('ABC & FSN Analysis'!G4:G33)", '#,##0', '₹#,##0.00'),
    ("Range (Max - Min)", "=MAX('ABC & FSN Analysis'!D4:D33)-MIN('ABC & FSN Analysis'!D4:D33)", "=MAX('ABC & FSN Analysis'!E4:E33)-MIN('ABC & FSN Analysis'!E4:E33)", "=MAX('ABC & FSN Analysis'!F4:F33)-MIN('ABC & FSN Analysis'!F4:F33)", "=MAX('ABC & FSN Analysis'!G4:G33)-MIN('ABC & FSN Analysis'!G4:G33)", '#,##0', '₹#,##0.00')
]

for idx, (label, f_mrp, f_pur, f_sales, f_rev, num_fmt, rev_fmt) in enumerate(stats_specs_linked, 5):
    c1 = ws_new.cell(row=idx, column=1, value=label)
    c2 = ws_new.cell(row=idx, column=2, value=f_mrp)
    c3 = ws_new.cell(row=idx, column=3, value=f_pur)
    c4 = ws_new.cell(row=idx, column=4, value=f_sales)
    c5 = ws_new.cell(row=idx, column=5, value=f_rev)
    
    c1.font = bold_font
    for c, cell in enumerate([c1, c2, c3, c4, c5], 1):
        cell.border = thin_border
        if c > 1:
            cell.font = regular_font
            cell.alignment = Alignment(horizontal="right")
            if idx == 5: # Sample Count N
                cell.number_format = '0'
            elif c in (2, 3):
                cell.number_format = '₹#,##0.00'
            elif c == 4:
                cell.number_format = num_fmt
            elif c == 5:
                cell.number_format = rev_fmt

# Column widths
ws_new.column_dimensions['A'].width = 25
ws_new.column_dimensions['B'].width = 18
ws_new.column_dimensions['C'].width = 22
ws_new.column_dimensions['D'].width = 25
ws_new.column_dimensions['E'].width = 25

output_excel = 'Anand_Pharma_Primary_Dataset_Jan_June_2026.xlsx'
wb.save(output_excel)
print("EXCEL UPDATED SUCCESSFULLY!")
print("Added Descriptive Statistics to:")
print("1. 'ABC & FSN Analysis' sheet (Rows 36 to 43)")
print("2. New dedicated sheet 'Descriptive Statistics'")
