import openpyxl
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import numpy as np
import os

# ============================================================
# LOAD DATA
# ============================================================
wb = openpyxl.load_workbook('Anand_Pharma_Primary_Dataset_Jan_June_2026.xlsx', data_only=True)

# --- Monthly Sales Summary ---
ws_summary = wb['Monthly Sales Summary']
months = []
revenues = []
units = []
for row in ws_summary.iter_rows(min_row=2, max_row=ws_summary.max_row):
    m = row[0].value
    r = row[1].value
    u = row[2].value
    if m and r:
        months.append(str(m))
        revenues.append(float(r))
        units.append(int(u))

# --- ABC & FSN Analysis ---
ws_abc = wb['ABC & FSN Analysis']
medicines = []
categories = []
mrps = []
purchase_rates = []
total_sales_qty = []
total_revenue = []
cum_rev_pct = []
abc_cat = []
fsn_cat = []
expiry_risk = []

for row in ws_abc.iter_rows(min_row=4, max_row=ws_abc.max_row):
    sno = row[0].value
    if sno is None:
        break
    med_name = str(row[1].value) if row[1].value else ''
    medicines.append(med_name)
    categories.append(str(row[2].value) if row[2].value else '')
    mrps.append(float(row[3].value) if row[3].value else 0)
    purchase_rates.append(float(row[4].value) if row[4].value else 0)
    total_sales_qty.append(int(row[5].value) if row[5].value else 0)
    total_revenue.append(float(row[6].value) if row[6].value else 0)
    cum_str = str(row[7].value).replace('%', '') if row[7].value else '0'
    cum_rev_pct.append(float(cum_str))
    abc_cat.append(str(row[8].value) if row[8].value else '')
    fsn_cat.append(str(row[9].value) if row[9].value else '')
    expiry_risk.append(str(row[10].value) if row[10].value else '')

# --- Monthly data per medicine ---
month_sheets = ['Jan 2026', 'Feb 2026', 'Mar 2026', 'Apr 2026', 'May 2026', 'June 2026']
monthly_data = {}  # {medicine_name: {month: {sales_qty, revenue, opening, closing, purchase}}}

for ms in month_sheets:
    ws = wb[ms]
    for row in ws.iter_rows(min_row=4, max_row=ws.max_row):
        sno = row[0].value
        if sno is None:
            break
        med = str(row[1].value).strip() if row[1].value else ''
        if med not in monthly_data:
            monthly_data[med] = {}
        monthly_data[med][ms] = {
            'sales_qty': int(row[7].value) if row[7].value else 0,
            'revenue': float(row[9].value) if row[9].value else 0,
            'opening': int(row[5].value) if row[5].value else 0,
            'purchase': int(row[6].value) if row[6].value else 0,
            'closing': int(row[8].value) if row[8].value else 0,
        }

print(f"Loaded {len(medicines)} medicines, {len(months)} months")
print(f"Monthly data for {len(monthly_data)} medicines")

# ============================================================
# OUTPUT DIRECTORY
# ============================================================
out_dir = 'BDM_Charts'
os.makedirs(out_dir, exist_ok=True)

# ============================================================
# COLOR PALETTE
# ============================================================
COLORS = {
    'primary': '#1a73e8',
    'secondary': '#34a853',
    'accent': '#ea4335',
    'warning': '#fbbc04',
    'dark': '#202124',
    'light': '#f8f9fa',
    'cat_a': '#e53935',
    'cat_b': '#fb8c00',
    'cat_c': '#43a047',
    'fast': '#1e88e5',
    'slow': '#fdd835',
    'non': '#e53935',
}

plt.rcParams.update({
    'font.family': 'serif',
    'font.serif': ['Times New Roman'],
    'font.size': 11,
    'axes.titlesize': 14,
    'axes.labelsize': 12,
    'figure.facecolor': 'white',
    'axes.facecolor': '#fafafa',
    'axes.grid': True,
    'grid.alpha': 0.3,
    'grid.color': '#cccccc',
})

# ============================================================
# FIGURE 1: Monthly Revenue Trend Line Chart
# ============================================================
fig, ax1 = plt.subplots(figsize=(10, 6))

color_rev = COLORS['primary']
color_units = COLORS['secondary']

ax1.plot(months, revenues, 'o-', color=color_rev, linewidth=2.5, markersize=8, label='Revenue', zorder=5)
ax1.fill_between(range(len(months)), revenues, alpha=0.1, color=color_rev)
ax1.set_xlabel('Month', fontweight='bold')
ax1.set_ylabel('Total Revenue (INR)', color=color_rev, fontweight='bold')
ax1.tick_params(axis='y', labelcolor=color_rev)

# Annotate values
for i, (m, r) in enumerate(zip(months, revenues)):
    ax1.annotate(f'Rs.{r:,.0f}', (i, r), textcoords="offset points", 
                xytext=(0, 15), ha='center', fontsize=9, fontweight='bold', color=color_rev)

ax2 = ax1.twinx()
ax2.bar(months, units, alpha=0.3, color=color_units, width=0.5, label='Units Sold', zorder=3)
ax2.set_ylabel('Total Units Sold', color=color_units, fontweight='bold')
ax2.tick_params(axis='y', labelcolor=color_units)

# Add seasonal annotations
ax1.annotate('Post-Winter\nDip', xy=(3, revenues[3]), xytext=(3, revenues[3]-20000),
            arrowprops=dict(arrowstyle='->', color=COLORS['accent']),
            fontsize=8, color=COLORS['accent'], ha='center', fontweight='bold')
ax1.annotate('Pre-Monsoon\nPeak (+23.8%)', xy=(4, revenues[4]), xytext=(4, revenues[4]+8000),
            arrowprops=dict(arrowstyle='->', color=COLORS['secondary']),
            fontsize=8, color=COLORS['secondary'], ha='center', fontweight='bold')

ax1.set_title('Figure 1: Monthly Revenue Trend & Units Sold (Jan-June 2026)\nAnand Pharma, Darbhanga', 
              fontweight='bold', pad=15)

lines1, labels1 = ax1.get_legend_handles_labels()
lines2, labels2 = ax2.get_legend_handles_labels()
ax1.legend(lines1 + lines2, labels1 + labels2, loc='upper left', framealpha=0.9)

plt.tight_layout()
plt.savefig(f'{out_dir}/Fig1_Monthly_Revenue_Trend.png', dpi=200, bbox_inches='tight')
plt.close()
print("Figure 1: Monthly Revenue Trend - DONE")

# ============================================================
# FIGURE 2: ABC Classification Donut Chart
# ============================================================
fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6))

# Count items per ABC category
abc_counts = {'A': 0, 'B': 0, 'C': 0}
abc_revenues_total = {'A': 0, 'B': 0, 'C': 0}
for i, cat in enumerate(abc_cat):
    if 'Category A' in cat:
        abc_counts['A'] += 1
        abc_revenues_total['A'] += total_revenue[i]
    elif 'Category B' in cat:
        abc_counts['B'] += 1
        abc_revenues_total['B'] += total_revenue[i]
    elif 'Category C' in cat:
        abc_counts['C'] += 1
        abc_revenues_total['C'] += total_revenue[i]

# Donut 1: Item Count
labels = ['Category A\n(High Value)', 'Category B\n(Moderate)', 'Category C\n(Low Value)']
sizes_items = [abc_counts['A'], abc_counts['B'], abc_counts['C']]
colors_abc = [COLORS['cat_a'], COLORS['cat_b'], COLORS['cat_c']]
explode = (0.05, 0, 0)

wedges1, texts1, autotexts1 = ax1.pie(sizes_items, explode=explode, labels=labels, 
    colors=colors_abc, autopct='%1.1f%%', startangle=90, pctdistance=0.75,
    textprops={'fontsize': 10})
centre_circle = plt.Circle((0, 0), 0.50, fc='white')
ax1.add_artist(centre_circle)
ax1.set_title('By Item Count (N=30)', fontweight='bold', pad=15)
ax1.text(0, 0, f'{sum(sizes_items)}\nItems', ha='center', va='center', fontsize=14, fontweight='bold')

# Donut 2: Revenue Share
sizes_rev = [abc_revenues_total['A'], abc_revenues_total['B'], abc_revenues_total['C']]
total_rev = sum(sizes_rev)
labels_rev = [f'Cat A\nRs.{abc_revenues_total["A"]:,.0f}', 
              f'Cat B\nRs.{abc_revenues_total["B"]:,.0f}', 
              f'Cat C\nRs.{abc_revenues_total["C"]:,.0f}']

wedges2, texts2, autotexts2 = ax2.pie(sizes_rev, explode=explode, labels=labels_rev, 
    colors=colors_abc, autopct='%1.1f%%', startangle=90, pctdistance=0.75,
    textprops={'fontsize': 10})
centre_circle2 = plt.Circle((0, 0), 0.50, fc='white')
ax2.add_artist(centre_circle2)
ax2.set_title('By Revenue Contribution', fontweight='bold', pad=15)
ax2.text(0, 0, f'Rs.{total_rev:,.0f}\nTotal', ha='center', va='center', fontsize=11, fontweight='bold')

fig.suptitle('Figure 2: ABC Classification Analysis\nAnand Pharma (Jan-June 2026)', 
             fontweight='bold', fontsize=14, y=1.02)
plt.tight_layout()
plt.savefig(f'{out_dir}/Fig2_ABC_Donut_Chart.png', dpi=200, bbox_inches='tight')
plt.close()
print("Figure 2: ABC Donut Chart - DONE")

# ============================================================
# FIGURE 3: Pareto Chart (80/20 Rule)
# ============================================================
fig, ax1 = plt.subplots(figsize=(14, 7))

# Sort by revenue descending
sorted_indices = sorted(range(len(total_revenue)), key=lambda i: total_revenue[i], reverse=True)
sorted_meds = [medicines[i][:20] for i in sorted_indices]
sorted_revs = [total_revenue[i] for i in sorted_indices]
cumulative_pct = []
running = 0
total_r = sum(sorted_revs)
for r in sorted_revs:
    running += r
    cumulative_pct.append((running / total_r) * 100)

# Assign colors based on ABC
bar_colors = []
for i in sorted_indices:
    if 'Category A' in abc_cat[i]:
        bar_colors.append(COLORS['cat_a'])
    elif 'Category B' in abc_cat[i]:
        bar_colors.append(COLORS['cat_b'])
    else:
        bar_colors.append(COLORS['cat_c'])

x = np.arange(len(sorted_meds))
bars = ax1.bar(x, sorted_revs, color=bar_colors, width=0.7, alpha=0.85, zorder=3)
ax1.set_ylabel('Revenue (INR)', fontweight='bold', color=COLORS['dark'])
ax1.set_xticks(x)
ax1.set_xticklabels(sorted_meds, rotation=65, ha='right', fontsize=8)

ax2 = ax1.twinx()
ax2.plot(x, cumulative_pct, 'D-', color=COLORS['dark'], linewidth=2, markersize=5, zorder=5)
ax2.set_ylabel('Cumulative Revenue %', fontweight='bold', color=COLORS['dark'])
ax2.set_ylim(0, 105)

# 80% line
ax2.axhline(y=80, color=COLORS['accent'], linestyle='--', linewidth=1.5, alpha=0.7)
ax2.annotate('80% Revenue Line', xy=(len(sorted_meds)*0.6, 81), fontsize=9, 
            color=COLORS['accent'], fontweight='bold')

# Find 80% point
for i, cp in enumerate(cumulative_pct):
    if cp >= 80:
        ax1.axvline(x=i+0.5, color=COLORS['accent'], linestyle='--', linewidth=1, alpha=0.5)
        ax1.annotate(f'{i+1} items = 80%\nof revenue', xy=(i+0.5, max(sorted_revs)*0.9),
                    fontsize=9, color=COLORS['accent'], fontweight='bold', ha='center')
        break

# Legend
from matplotlib.patches import Patch
legend_elements = [Patch(facecolor=COLORS['cat_a'], label=f'Category A ({abc_counts["A"]} items)'),
                   Patch(facecolor=COLORS['cat_b'], label=f'Category B ({abc_counts["B"]} items)'),
                   Patch(facecolor=COLORS['cat_c'], label=f'Category C ({abc_counts["C"]} items)')]
ax1.legend(handles=legend_elements, loc='upper right', framealpha=0.9)

ax1.set_title('Figure 3: Pareto Analysis (80/20 Rule) - Revenue Contribution by Medicine\nAnand Pharma (Jan-June 2026)', 
              fontweight='bold', pad=15)
plt.tight_layout()
plt.savefig(f'{out_dir}/Fig3_Pareto_Chart.png', dpi=200, bbox_inches='tight')
plt.close()
print("Figure 3: Pareto Chart - DONE")

# ============================================================
# FIGURE 4: FSN Distribution Horizontal Bar Chart
# ============================================================
fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 7))

# FSN counts
fsn_counts = {'Fast Moving (F)': 0, 'Slow Moving (S)': 0, 'Non Moving (N)': 0}
fsn_meds = {'Fast Moving (F)': [], 'Slow Moving (S)': [], 'Non Moving (N)': []}
for i, f in enumerate(fsn_cat):
    for key in fsn_counts:
        if key.split('(')[1][0] in f:
            fsn_counts[key] += 1
            fsn_meds[key].append((medicines[i], total_sales_qty[i]))

# Bar chart
fsn_labels = list(fsn_counts.keys())
fsn_values = list(fsn_counts.values())
fsn_colors = [COLORS['fast'], COLORS['slow'], COLORS['non']]

bars = ax1.barh(fsn_labels, fsn_values, color=fsn_colors, height=0.5, edgecolor='white', linewidth=1.5)
for bar, val in zip(bars, fsn_values):
    ax1.text(bar.get_width() + 0.3, bar.get_y() + bar.get_height()/2, 
            f'{val} items ({val/30*100:.1f}%)', va='center', fontweight='bold', fontsize=11)
ax1.set_xlabel('Number of Medicines', fontweight='bold')
ax1.set_title('FSN Category Distribution', fontweight='bold')
ax1.set_xlim(0, max(fsn_values) + 5)

# Non-moving medicines detail with expiry risk
non_moving = [(medicines[i], total_sales_qty[i], expiry_risk[i]) 
              for i in range(len(medicines)) if 'N' in fsn_cat[i] and 'Non' in fsn_cat[i]]
if non_moving:
    nm_names = [nm[0][:18] for nm in non_moving]
    nm_qty = [nm[1] for nm in non_moving]
    nm_colors = [COLORS['accent'] if 'High' in nm[2] else COLORS['warning'] for nm in non_moving]
    
    bars2 = ax2.barh(nm_names, nm_qty, color=nm_colors, height=0.5)
    for bar, val in zip(bars2, nm_qty):
        ax2.text(bar.get_width() + 2, bar.get_y() + bar.get_height()/2, 
                f'{val} units', va='center', fontsize=10)
    ax2.set_xlabel('Total 6-Month Sales (Units)', fontweight='bold')
    ax2.set_title('Non-Moving Items (Expiry Risk)', fontweight='bold')
    ax2.axvline(x=190, color=COLORS['accent'], linestyle='--', alpha=0.7)
    ax2.annotate('Threshold: 190 units', xy=(190, len(nm_names)-0.5), fontsize=8, color=COLORS['accent'])

fig.suptitle('Figure 4: FSN (Fast-Slow-Non Moving) Inventory Classification\nAnand Pharma (Jan-June 2026)', 
             fontweight='bold', fontsize=14, y=1.02)
plt.tight_layout()
plt.savefig(f'{out_dir}/Fig4_FSN_Analysis.png', dpi=200, bbox_inches='tight')
plt.close()
print("Figure 4: FSN Analysis - DONE")

# ============================================================
# FIGURE 5: Moving Average Demand Forecast
# ============================================================
fig, axes = plt.subplots(2, 3, figsize=(16, 10))
axes = axes.flatten()

# Top 6 Category A medicines
top_6_meds = [medicines[i] for i in range(min(6, len(medicines)))]

for idx, med in enumerate(top_6_meds):
    ax = axes[idx]
    if med in monthly_data:
        sales = [monthly_data[med].get(ms, {}).get('sales_qty', 0) for ms in month_sheets]
    else:
        # Try partial match
        matched = None
        for key in monthly_data:
            if med[:10].lower() in key.lower():
                matched = key
                break
        if matched:
            sales = [monthly_data[matched].get(ms, {}).get('sales_qty', 0) for ms in month_sheets]
        else:
            sales = [0]*6
    
    month_labels = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun']
    
    # 3-month moving average
    ma = [None, None, None]
    for i in range(3, len(sales)):
        ma.append((sales[i-1] + sales[i-2] + sales[i-3]) / 3)
    
    # Forecast July
    if len(sales) >= 3:
        forecast_july = (sales[-1] + sales[-2] + sales[-3]) / 3
        extended_months = month_labels + ['Jul*']
        extended_sales = sales + [None]
        extended_ma = ma + [forecast_july]
    else:
        extended_months = month_labels
        extended_sales = sales
        extended_ma = ma
    
    ax.plot(range(len(month_labels)), sales, 'o-', color=COLORS['primary'], linewidth=2, 
            markersize=6, label='Actual Sales', zorder=5)
    
    # Plot MA where available
    ma_x = [i for i, v in enumerate(extended_ma) if v is not None]
    ma_y = [v for v in extended_ma if v is not None]
    ax.plot(ma_x, ma_y, 's--', color=COLORS['accent'], linewidth=1.5, 
            markersize=5, label='3-Month MA', zorder=4)
    
    # Highlight forecast
    if len(extended_months) > 6:
        ax.plot(6, forecast_july, '*', color=COLORS['secondary'], markersize=15, 
                zorder=6, label=f'July Forecast: {forecast_july:.0f}')
        ax.axvline(x=5.5, color='gray', linestyle=':', alpha=0.5)
        ax.text(6, forecast_july + 3, f'{forecast_july:.0f}', ha='center', fontsize=9, 
                fontweight='bold', color=COLORS['secondary'])
    
    ax.set_title(med[:22], fontweight='bold', fontsize=10)
    ax.set_xticks(range(len(extended_months)))
    ax.set_xticklabels(extended_months, fontsize=8)
    ax.legend(fontsize=7, loc='best')
    ax.set_ylabel('Units Sold', fontsize=9)

fig.suptitle('Figure 5: 3-Month Moving Average Demand Forecast (Top 6 Category A Medicines)\nAnand Pharma (Jan-June 2026)', 
             fontweight='bold', fontsize=14, y=1.02)
plt.tight_layout()
plt.savefig(f'{out_dir}/Fig5_Moving_Average_Forecast.png', dpi=200, bbox_inches='tight')
plt.close()
print("Figure 5: Moving Average Forecast - DONE")

# ============================================================
# FIGURE 6: Revenue by Medicine Category (Therapeutic Group)
# ============================================================
fig, ax = plt.subplots(figsize=(12, 6))

# Group by therapeutic category
cat_revenue = {}
cat_count = {}
for i, cat in enumerate(categories):
    cat_clean = cat.strip()
    if cat_clean not in cat_revenue:
        cat_revenue[cat_clean] = 0
        cat_count[cat_clean] = 0
    cat_revenue[cat_clean] += total_revenue[i]
    cat_count[cat_clean] += 1

# Sort by revenue
sorted_cats = sorted(cat_revenue.items(), key=lambda x: x[1], reverse=True)
cat_names = [c[0][:25] for c in sorted_cats]
cat_revs = [c[1] for c in sorted_cats]

# Gradient colors
cmap = plt.cm.RdYlGn_r
colors_gradient = [cmap(i / len(cat_names)) for i in range(len(cat_names))]

bars = ax.barh(range(len(cat_names)), cat_revs, color=colors_gradient, height=0.6, edgecolor='white')
ax.set_yticks(range(len(cat_names)))
ax.set_yticklabels(cat_names, fontsize=10)
ax.invert_yaxis()

for bar, val, cn in zip(bars, cat_revs, [c[0] for c in sorted_cats]):
    count = cat_count[cn]
    ax.text(bar.get_width() + 1000, bar.get_y() + bar.get_height()/2, 
            f'Rs.{val:,.0f} ({count} items)', va='center', fontsize=9, fontweight='bold')

ax.set_xlabel('Total 6-Month Revenue (INR)', fontweight='bold')
ax.set_title('Figure 6: Revenue Distribution by Therapeutic Category\nAnand Pharma (Jan-June 2026)', 
             fontweight='bold', pad=15)
ax.set_xlim(0, max(cat_revs) * 1.25)
plt.tight_layout()
plt.savefig(f'{out_dir}/Fig6_Category_Revenue.png', dpi=200, bbox_inches='tight')
plt.close()
print("Figure 6: Category Revenue - DONE")

# ============================================================
# FIGURE 7: Sales Qty vs Revenue Scatter Plot (Bubble Chart)
# ============================================================
fig, ax = plt.subplots(figsize=(12, 8))

scatter_colors = []
for cat in abc_cat:
    if 'Category A' in cat:
        scatter_colors.append(COLORS['cat_a'])
    elif 'Category B' in cat:
        scatter_colors.append(COLORS['cat_b'])
    else:
        scatter_colors.append(COLORS['cat_c'])

# Bubble size based on MRP
bubble_sizes = [m * 1.5 for m in mrps]

scatter = ax.scatter(total_sales_qty, total_revenue, s=bubble_sizes, c=scatter_colors, 
                     alpha=0.7, edgecolors='white', linewidth=1.5, zorder=5)

# Label top medicines
for i in range(len(medicines)):
    if total_revenue[i] > 60000 or total_sales_qty[i] > 280:
        ax.annotate(medicines[i][:15], (total_sales_qty[i], total_revenue[i]),
                   textcoords="offset points", xytext=(8, 8), fontsize=8,
                   fontweight='bold', alpha=0.8)

ax.set_xlabel('Total 6-Month Sales Quantity (Units/Strips)', fontweight='bold')
ax.set_ylabel('Total 6-Month Revenue (INR)', fontweight='bold')
ax.set_title('Figure 7: Sales Quantity vs Revenue (Bubble Size = MRP)\nAnand Pharma (Jan-June 2026)', 
             fontweight='bold', pad=15)

# Legend
from matplotlib.lines import Line2D
legend_elements = [
    Line2D([0], [0], marker='o', color='w', markerfacecolor=COLORS['cat_a'], markersize=12, label='Category A'),
    Line2D([0], [0], marker='o', color='w', markerfacecolor=COLORS['cat_b'], markersize=12, label='Category B'),
    Line2D([0], [0], marker='o', color='w', markerfacecolor=COLORS['cat_c'], markersize=12, label='Category C'),
]
ax.legend(handles=legend_elements, loc='upper left', framealpha=0.9)
plt.tight_layout()
plt.savefig(f'{out_dir}/Fig7_Sales_vs_Revenue_Bubble.png', dpi=200, bbox_inches='tight')
plt.close()
print("Figure 7: Bubble Chart - DONE")

# ============================================================
# FIGURE 8: Profit Margin Analysis
# ============================================================
fig, ax = plt.subplots(figsize=(14, 7))

margins = [(mrps[i] - purchase_rates[i]) / mrps[i] * 100 for i in range(len(medicines))]
margin_sorted_idx = sorted(range(len(margins)), key=lambda i: margins[i], reverse=True)

med_names_sorted = [medicines[i][:18] for i in margin_sorted_idx]
margins_sorted = [margins[i] for i in margin_sorted_idx]
colors_margin = [COLORS['secondary'] if m > 50 else COLORS['warning'] if m > 30 else COLORS['accent'] for m in margins_sorted]

bars = ax.barh(range(len(med_names_sorted)), margins_sorted, color=colors_margin, height=0.7)
ax.set_yticks(range(len(med_names_sorted)))
ax.set_yticklabels(med_names_sorted, fontsize=8)
ax.invert_yaxis()

for bar, val in zip(bars, margins_sorted):
    ax.text(bar.get_width() + 0.5, bar.get_y() + bar.get_height()/2, 
            f'{val:.1f}%', va='center', fontsize=8, fontweight='bold')

ax.axvline(x=50, color='gray', linestyle='--', alpha=0.5)
ax.text(50.5, -0.5, '50% margin line', fontsize=8, color='gray')
ax.set_xlabel('Profit Margin %', fontweight='bold')
ax.set_title('Figure 8: Profit Margin Analysis (MRP vs Purchase Price)\nAnand Pharma - 30 Medicine SKUs', 
             fontweight='bold', pad=15)

legend_elements = [
    Patch(facecolor=COLORS['secondary'], label='High Margin (>50%)'),
    Patch(facecolor=COLORS['warning'], label='Medium Margin (30-50%)'),
    Patch(facecolor=COLORS['accent'], label='Low Margin (<30%)'),
]
ax.legend(handles=legend_elements, loc='lower right', framealpha=0.9)
plt.tight_layout()
plt.savefig(f'{out_dir}/Fig8_Profit_Margin.png', dpi=200, bbox_inches='tight')
plt.close()
print("Figure 8: Profit Margin - DONE")

# ============================================================
# FIGURE 9: Combined ABC-FSN Matrix Heatmap
# ============================================================
fig, ax = plt.subplots(figsize=(10, 7))

# Create ABC-FSN matrix
matrix_counts = {}
matrix_meds = {}
for abc in ['A', 'B', 'C']:
    for fsn in ['F', 'S', 'N']:
        matrix_counts[(abc, fsn)] = 0
        matrix_meds[(abc, fsn)] = []

for i in range(len(medicines)):
    abc_letter = 'A' if 'Category A' in abc_cat[i] else 'B' if 'Category B' in abc_cat[i] else 'C'
    fsn_letter = 'F' if 'Fast' in fsn_cat[i] else 'N' if 'Non' in fsn_cat[i] else 'S'
    matrix_counts[(abc_letter, fsn_letter)] += 1
    matrix_meds[(abc_letter, fsn_letter)].append(medicines[i][:12])

# Build heatmap data
abc_labels = ['Category A\n(High Value)', 'Category B\n(Moderate)', 'Category C\n(Low Value)']
fsn_labels_hm = ['Fast Moving (F)', 'Slow Moving (S)', 'Non Moving (N)']
heatmap_data = np.zeros((3, 3))
annotations = []

for i, abc in enumerate(['A', 'B', 'C']):
    row_ann = []
    for j, fsn in enumerate(['F', 'S', 'N']):
        count = matrix_counts[(abc, fsn)]
        heatmap_data[i][j] = count
        meds_list = ', '.join(matrix_meds[(abc, fsn)][:3])
        if count > 0:
            row_ann.append(f'{count}\n({meds_list}...)')
        else:
            row_ann.append('0')
    annotations.append(row_ann)

# Risk colors: green (AF) -> red (CN)
risk_cmap = plt.cm.RdYlGn_r
im = ax.imshow(heatmap_data, cmap=risk_cmap, aspect='auto', vmin=0, vmax=max(max(row) for row in heatmap_data))

# Add text annotations
for i in range(3):
    for j in range(3):
        count = int(heatmap_data[i][j])
        meds = matrix_meds[(['A','B','C'][i], ['F','S','N'][j])]
        text = f'{count} items'
        if meds:
            text += f'\n{", ".join([m[:10] for m in meds[:2]])}'
        color = 'white' if heatmap_data[i][j] > 8 else 'black'
        ax.text(j, i, text, ha='center', va='center', fontsize=9, fontweight='bold', color=color)

ax.set_xticks(range(3))
ax.set_xticklabels(fsn_labels_hm, fontweight='bold')
ax.set_yticks(range(3))
ax.set_yticklabels(abc_labels, fontweight='bold')
ax.set_title('Figure 9: ABC-FSN Cross-Classification Matrix\nAnand Pharma (Jan-June 2026)', 
             fontweight='bold', pad=15)

# Priority labels
ax.text(-0.5, -0.7, 'Priority: AF (Reorder Fast) >> AN/CF (Critical) >> CN (Discontinue)', 
        fontsize=9, style='italic', transform=ax.transData)

plt.colorbar(im, label='Number of SKUs', shrink=0.8)
plt.tight_layout()
plt.savefig(f'{out_dir}/Fig9_ABC_FSN_Matrix.png', dpi=200, bbox_inches='tight')
plt.close()
print("Figure 9: ABC-FSN Matrix - DONE")

# ============================================================
# FIGURE 10: Stock Holding Pattern (Top 5 medicines)
# ============================================================
fig, axes = plt.subplots(1, 5, figsize=(20, 5), sharey=False)

top5 = medicines[:5]
for idx, med in enumerate(top5):
    ax = axes[idx]
    if med in monthly_data:
        data = monthly_data[med]
    else:
        matched = None
        for key in monthly_data:
            if med[:10].lower() in key.lower():
                matched = key
                break
        data = monthly_data.get(matched, {}) if matched else {}
    
    opening = [data.get(ms, {}).get('opening', 0) for ms in month_sheets]
    closing = [data.get(ms, {}).get('closing', 0) for ms in month_sheets]
    sales = [data.get(ms, {}).get('sales_qty', 0) for ms in month_sheets]
    purchase = [data.get(ms, {}).get('purchase', 0) for ms in month_sheets]
    
    x = np.arange(6)
    width = 0.2
    ax.bar(x - width, opening, width, label='Opening', color=COLORS['primary'], alpha=0.8)
    ax.bar(x, purchase, width, label='Purchase', color=COLORS['secondary'], alpha=0.8)
    ax.bar(x + width, sales, width, label='Sales', color=COLORS['accent'], alpha=0.8)
    
    ax.plot(x, closing, 'ko-', markersize=4, linewidth=1.5, label='Closing')
    
    ax.set_title(med[:18], fontsize=9, fontweight='bold')
    ax.set_xticks(x)
    ax.set_xticklabels(['J', 'F', 'M', 'A', 'M', 'J'], fontsize=8)
    if idx == 0:
        ax.set_ylabel('Units', fontweight='bold')
        ax.legend(fontsize=6, loc='best')

fig.suptitle('Figure 10: Stock Movement Pattern (Top 5 Category A Medicines)\nOpening Stock, Purchase, Sales & Closing Stock', 
             fontweight='bold', fontsize=13, y=1.05)
plt.tight_layout()
plt.savefig(f'{out_dir}/Fig10_Stock_Movement.png', dpi=200, bbox_inches='tight')
plt.close()
print("Figure 10: Stock Movement - DONE")

print(f"\n{'='*60}")
print(f"ALL 10 CHARTS SAVED IN: {out_dir}/")
print(f"{'='*60}")
for f in sorted(os.listdir(out_dir)):
    print(f"  - {f}")
