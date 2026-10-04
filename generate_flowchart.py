import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import os

out_dir = 'BDM_Charts'
os.makedirs(out_dir, exist_ok=True)

fig, ax = plt.subplots(figsize=(6, 7.5), dpi=300)
ax.set_facecolor('#0b2545')
fig.patch.set_facecolor('#0b2545')

# Draw Flowchart Boxes matching user's screenshot layout
def draw_box(ax, x, y, w, h, text, subtext="", color='#1d4ed8', text_color='white'):
    box = patches.FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.06,rounding_size=0.12",
                                facecolor=color, edgecolor='#38bdf8', linewidth=1, zorder=3)
    ax.add_patch(box)
    
    if subtext:
        ax.text(x + w/2, y + h*0.62, text, ha='center', va='center', color=text_color,
                fontsize=9, fontweight='bold', zorder=4)
        ax.text(x + w/2, y + h*0.32, subtext, ha='center', va='center', color='#cbd5e1',
                fontsize=7.5, fontstyle='italic', zorder=4)
    else:
        ax.text(x + w/2, y + h/2, text, ha='center', va='center', color=text_color,
                fontsize=9.5, fontweight='bold', zorder=4)

def draw_arrow(ax, x1, y1, x2, y2, color='#38bdf8'):
    ax.annotate('', xy=(x2, y2), xytext=(x1, y1),
                arrowprops=dict(arrowstyle='->', color=color, lw=1.8, mutation_scale=14), zorder=2)

# Top Header Node: Primary Operational Data
draw_box(ax, 1.2, 6.8, 3.6, 0.7, "Primary Operational Data", "(Sales, Inventory, Purchase & Expiry)", color='#1e293b')

# Left Column: Problem 1 Flow
draw_box(ax, 0.3, 5.5, 2.5, 0.6, "Inventory Control Problem", color='#1d4ed8')
draw_box(ax, 0.3, 4.3, 2.5, 0.6, "ABC Consumption Value", "(PCV = Sales Qty x Price)", color='#0f766e')
draw_box(ax, 0.3, 3.1, 2.5, 0.6, "EOQ, Safety Stock & ROP", "(Formula-based buffer)", color='#0f766e')
draw_box(ax, 0.3, 1.9, 2.5, 0.6, "15-Day Buffer Stock Policy", color='#065f46')

# Right Column: Problem 2 & 3 Flow
draw_box(ax, 3.2, 5.5, 2.5, 0.6, "Expiry & Demand Problem", color='#065f46')
draw_box(ax, 3.2, 4.3, 2.5, 0.6, "FSN Velocity & 60-Day Return", "(Fast, Slow, Non-Moving)", color='#0f766e')
draw_box(ax, 3.2, 3.1, 2.5, 0.6, "3-Day / 7-Day MA Forecast", "(MAE, RMSE, MAPE Metrics)", color='#0f766e')
draw_box(ax, 3.2, 1.9, 2.5, 0.6, "Optimized Inventory Policy", color='#1d4ed8')

# Arrows
draw_arrow(ax, 2.2, 6.8, 1.55, 6.1)
draw_arrow(ax, 3.8, 6.8, 4.45, 6.1)

draw_arrow(ax, 1.55, 5.5, 1.55, 4.9)
draw_arrow(ax, 1.55, 4.3, 1.55, 3.7)
draw_arrow(ax, 1.55, 3.1, 1.55, 2.5)

draw_arrow(ax, 4.45, 5.5, 4.45, 4.9)
draw_arrow(ax, 4.45, 4.3, 4.45, 3.7)
draw_arrow(ax, 4.45, 3.1, 4.45, 2.5)

ax.set_xlim(0, 6)
ax.set_ylim(1.5, 7.8)
ax.axis('off')

plt.tight_layout()
flowchart_path = os.path.join(out_dir, 'methodology_flowchart.png')
plt.savefig(flowchart_path, dpi=300, bbox_inches='tight', facecolor='#0b2545')
plt.close()

print(f"FLOWCHART SAVED SUCCESSFULLY: {flowchart_path}")
