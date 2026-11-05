import os
import shutil
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import numpy as np

OUTPUT_DIR = os.path.join('project_report', 'figures')
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Copy existing figures
src_figures = [
    ('research_paper/figures/fig1_system_architecture.png', 'fig1_system_architecture.png'),
    ('research_paper/figures/fig2_project_workflow.png', 'fig2_project_workflow.png'),
    ('research_paper/figures/fig3_model_comparison.png', 'fig3_model_comparison.png'),
    ('research_paper/figures/fig4_forecast_vs_actual.png', 'fig4_forecast_vs_actual.png'),
    ('research_paper/figures/fig5_inventory_status_distribution.png', 'fig5_inventory_status_distribution.png'),
    ('research_paper/figures/fig6_demand_scenario_comparison.png', 'fig6_demand_scenario_comparison.png'),
    ('research_paper/figures/fig7_financial_impact_comparison.png', 'fig7_financial_impact_comparison.png'),
    ('demandverse_ai_outputs/visualizations/inventory_status_breakdown.png', 'inventory_status_breakdown.png'),
    ('demandverse_ai_outputs/visualizations/weekday_sales_distribution.png', 'weekday_sales_distribution.png'),
    ('logo.png', 'logo.png')
]

for src, dst in src_figures:
    if os.path.exists(src):
        shutil.copy(src, os.path.join(OUTPUT_DIR, dst))
        print(f"Copied {src} -> {dst}")

# 1. DFD Level-0 (Context Diagram)
fig, ax = plt.subplots(figsize=(8.5, 4.5), dpi=300)
ax.axis('off')

# External Entities
# Left: Retail Category / Replenishment Manager
rect_user = patches.FancyBboxPatch((0.05, 0.35), 0.22, 0.30, boxstyle="round,pad=0.02",
                                   edgecolor="#2C3E50", facecolor="#EBF5FB", linewidth=1.5)
ax.add_patch(rect_user)
ax.text(0.16, 0.50, "Retail Operations &\nReplenishment Planners\n(Store CA_1)",
        ha='center', va='center', fontsize=9.5, weight='bold', color='#1B4F72')

# Center: DemandVerse AI Ecosystem
circ_sys = patches.Circle((0.50, 0.50), 0.18, edgecolor="#1B2631", facecolor="#FEF9E7", linewidth=2.0)
ax.add_patch(circ_sys)
ax.text(0.50, 0.52, "0.0\nDEMANDVERSE AI\nECOSYSTEM", ha='center', va='center',
        fontsize=10.5, weight='bold', color='#7D6608')
ax.text(0.50, 0.40, "(Adaptive Forecasting,\nInventory Optimization,\nFinancial Engine)",
        ha='center', va='center', fontsize=7.5, style='italic', color='#5D6D7E')

# Right: External Suppliers & ERP / Order Execution
rect_erp = patches.FancyBboxPatch((0.73, 0.35), 0.22, 0.30, boxstyle="round,pad=0.02",
                                  edgecolor="#2C3E50", facecolor="#E8F8F5", linewidth=1.5)
ax.add_patch(rect_erp)
ax.text(0.84, 0.50, "Supplier Logistics &\nExecutive Leadership\n(PO & Risk Reports)",
        ha='center', va='center', fontsize=9.5, weight='bold', color='#0E6251')

# Data Flows (User -> System)
ax.annotate('SKU Selection, Horizon (30d),\nSimulation Parameters', xy=(0.33, 0.56), xytext=(0.27, 0.56),
            arrowprops=dict(arrowstyle="->", color="#2980B9", lw=1.5))
# Data Flows (System -> User)
ax.annotate('Forecasts, Stock Health Gauges,\nRevenue-at-Risk Visuals', xy=(0.27, 0.44), xytext=(0.33, 0.44),
            arrowprops=dict(arrowstyle="->", color="#C0392B", lw=1.5))

# Data Flows (System -> ERP/Logistics)
ax.annotate('Recommended Purchase Orders (PO),\nReorder Quantities (RQ)', xy=(0.73, 0.56), xytext=(0.67, 0.56),
            arrowprops=dict(arrowstyle="->", color="#27AE60", lw=1.5))
# Data Flows (ERP -> System)
ax.annotate('M5 Benchmark Historical Data\n(Sales, Calendar, Sell Prices)', xy=(0.67, 0.44), xytext=(0.73, 0.44),
            arrowprops=dict(arrowstyle="->", color="#7D3C98", lw=1.5))

plt.title("Figure 5.1: Data Flow Diagram (DFD) Level-0 — Context Level Architecture",
          fontsize=11, weight='bold', pad=15)
plt.tight_layout()
fig.savefig(os.path.join(OUTPUT_DIR, 'dfd_level_0.png'), dpi=300, bbox_inches='tight')
plt.close(fig)
print("Generated dfd_level_0.png")

# 2. DFD Level-1 (Decomposed Process Flow)
fig, ax = plt.subplots(figsize=(9.5, 6.0), dpi=300)
ax.axis('off')

# Processes (Bubbles/Boxes)
p_boxes = [
    ("1.0 Ingestion &\nStratification", 0.08, 0.70, "#D6EAF8"),
    ("2.0 Dual-Engine\nModel Training", 0.38, 0.70, "#D5F5E3"),
    ("3.0 Adaptive\nModel Selection", 0.68, 0.70, "#FCF3CF"),
    ("4.0 Inventory\nOptimization", 0.20, 0.25, "#E8DAEF"),
    ("5.0 What-If &\nFinancial Engine", 0.55, 0.25, "#FADBD8"),
    ("6.0 REST API &\nDashboard Serving", 0.85, 0.25, "#EDBB99"),
]

for name, x, y, col in p_boxes:
    rect = patches.FancyBboxPatch((x-0.08, y-0.08), 0.16, 0.16, boxstyle="round,pad=0.015",
                                  edgecolor="#2C3E50", facecolor=col, linewidth=1.5)
    ax.add_patch(rect)
    ax.text(x, y, name, ha='center', va='center', fontsize=8.5, weight='bold', color='#1A252F')

# Data Stores (Horizontal bars)
ds1 = patches.Rectangle((0.20, 0.50), 0.25, 0.05, edgecolor="#34495E", facecolor="#EAEDED", linewidth=1.2)
ax.add_patch(ds1)
ax.text(0.325, 0.525, "D1: Clean Sales Dataset (28,695 rows)", ha='center', va='center', fontsize=8, weight='bold')

ds2 = patches.Rectangle((0.55, 0.50), 0.25, 0.05, edgecolor="#34495E", facecolor="#EAEDED", linewidth=1.2)
ax.add_patch(ds2)
ax.text(0.675, 0.525, "D2: 30-Day Final Forecasts (450 rows)", ha='center', va='center', fontsize=8, weight='bold')

# Connectors
ax.annotate('Melt & Key Merge', xy=(0.20, 0.53), xytext=(0.14, 0.65),
            arrowprops=dict(arrowstyle="->", color="#2980B9", lw=1.2))
ax.annotate('80/20 Chrono Split', xy=(0.35, 0.62), xytext=(0.28, 0.55),
            arrowprops=dict(arrowstyle="->", color="#2980B9", lw=1.2))
ax.annotate('ARIMA & Prophet\nPredictions', xy=(0.60, 0.70), xytext=(0.46, 0.70),
            arrowprops=dict(arrowstyle="->", color="#27AE60", lw=1.2))
ax.annotate('Route Optimal\nModel', xy=(0.68, 0.55), xytext=(0.68, 0.62),
            arrowprops=dict(arrowstyle="->", color="#D35400", lw=1.2))
ax.annotate('Forecast Curves', xy=(0.20, 0.33), xytext=(0.58, 0.50),
            arrowprops=dict(arrowstyle="->", color="#8E44AD", lw=1.2))
ax.annotate('Safety Stock (SS) &\nRequired Stock (RS)', xy=(0.47, 0.25), xytext=(0.28, 0.25),
            arrowprops=dict(arrowstyle="->", color="#C0392B", lw=1.2))
ax.annotate('Scenario Demand &\nRisk Exposure', xy=(0.77, 0.25), xytext=(0.63, 0.25),
            arrowprops=dict(arrowstyle="->", color="#16A085", lw=1.2))

plt.title("Figure 5.2: Data Flow Diagram (DFD) Level-1 — Decomposed System Functional Processing",
          fontsize=11, weight='bold', pad=15)
plt.tight_layout()
fig.savefig(os.path.join(OUTPUT_DIR, 'dfd_level_1.png'), dpi=300, bbox_inches='tight')
plt.close(fig)
print("Generated dfd_level_1.png")

# 3. ER Diagram / Data Architecture
fig, ax = plt.subplots(figsize=(9.0, 5.2), dpi=300)
ax.axis('off')

entities = [
    ("PRODUCTS (Catalog)", 0.05, 0.60, [
        "item_id (PK)", "dept_id", "cat_id", "store_id", "state_id"
    ], "#AED6F1"),
    ("CALENDAR", 0.40, 0.60, [
        "date (PK)", "wm_yr_wk", "weekday", "month", "year", "event_name_1", "snap_CA"
    ], "#A9DFBF"),
    ("SELL_PRICES", 0.75, 0.60, [
        "store_id (PK)", "item_id (PK)", "wm_yr_wk (PK)", "sell_price"
    ], "#F9E79F"),
    ("CLEAN_SALES", 0.05, 0.10, [
        "id (PK)", "item_id (FK)", "store_id", "date (FK)", "sales", "sell_price", "snap_CA"
    ], "#D7BDE2"),
    ("MODEL_COMPARISON", 0.40, 0.10, [
        "product_id (PK)", "store_id", "ARIMA_MAE", "Prophet_MAE", "Selected_Model"
    ], "#F5B7B1"),
    ("INVENTORY_LEDGER", 0.75, 0.10, [
        "item_id (PK)", "forecasted_demand", "safety_stock", "simulated_current_stock", "reorder_quantity", "inventory_status"
    ], "#F5CBA7")
]

for name, x, y, attrs, col in entities:
    # Header
    rect_h = patches.Rectangle((x, y+0.22), 0.22, 0.06, facecolor=col, edgecolor="#2C3E50", linewidth=1.2)
    ax.add_patch(rect_h)
    ax.text(x+0.11, y+0.25, name, ha='center', va='center', fontsize=8.5, weight='bold', color='#1A252F')
    # Body
    rect_b = patches.Rectangle((x, y), 0.22, 0.22, facecolor="#FDFEFE", edgecolor="#2C3E50", linewidth=1.2)
    ax.add_patch(rect_b)
    for idx, attr in enumerate(attrs):
        y_txt = y + 0.19 - (idx * 0.035)
        ax.text(x+0.015, y_txt, f"• {attr}", ha='left', va='center', fontsize=7.5, color='#2C3E50')

# Connecting Lines / Relations
ax.annotate('', xy=(0.16, 0.32), xytext=(0.16, 0.60),
            arrowprops=dict(arrowstyle="<->", color="#34495E", lw=1.2))
ax.text(0.17, 0.46, "1 : N (Transactions)", fontsize=7, style='italic', color='#566573')

ax.annotate('', xy=(0.40, 0.21), xytext=(0.27, 0.21),
            arrowprops=dict(arrowstyle="->", color="#34495E", lw=1.2))
ax.text(0.33, 0.23, "evaluates", fontsize=7, style='italic', color='#566573')

ax.annotate('', xy=(0.75, 0.21), xytext=(0.62, 0.21),
            arrowprops=dict(arrowstyle="->", color="#34495E", lw=1.2))
ax.text(0.68, 0.23, "optimizes", fontsize=7, style='italic', color='#566573')

plt.title("Figure 5.3: Entity-Relationship (ER) & Relational Schema Architecture",
          fontsize=11, weight='bold', pad=15)
plt.tight_layout()
fig.savefig(os.path.join(OUTPUT_DIR, 'er_diagram.png'), dpi=300, bbox_inches='tight')
plt.close(fig)
print("Generated er_diagram.png")

print("All report diagrams generated successfully!")

