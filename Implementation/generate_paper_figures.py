import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib.gridspec import GridSpec

# Set style for IEEE publication
plt.rcParams['font.family'] = 'serif'
plt.rcParams['font.size'] = 9
plt.rcParams['axes.labelsize'] = 10
plt.rcParams['axes.titlesize'] = 10
plt.rcParams['xtick.labelsize'] = 8.5
plt.rcParams['ytick.labelsize'] = 8.5
plt.rcParams['legend.fontsize'] = 8.5
plt.rcParams['figure.titlesize'] = 11

OUTPUT_DIR = os.path.join('research_paper', 'figures')
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Load datasets
sales_df = pd.read_csv('demandverse_ai_outputs/data/clean_sales.csv', parse_dates=['date'])
models_df = pd.read_csv('demandverse_ai_outputs/forecasts/model_comparison.csv')
forecasts_df = pd.read_csv('demandverse_ai_outputs/forecasts/final_forecasts.csv', parse_dates=['date'])
inv_df = pd.read_csv('demandverse_ai_outputs/inventory/inventory_recommendations.csv')
whatif_df = pd.read_csv('demandverse_ai_outputs/financial/what_if_analysis.csv')
fin_df = pd.read_csv('demandverse_ai_outputs/financial/financial_impact.csv')

print("Loaded all datasets successfully.")

# ==============================================================================
# FIGURE 1: System Architecture Diagram
# ==============================================================================
fig, ax = plt.subplots(figsize=(7.2, 4.2), dpi=300)
ax.axis('off')

# Tier definitions: (name, y, color, height)
tiers = [
    ("Application & Serving Tier\n(FastAPI ASGI, LRU Cache, CORS Middleware, React 19 SPA, Recharts, Vite)", 0.82, "#EBF5FB", 0.14),
    ("Financial & Scenario Stress-Testing Engine\n(What-If Demand Shocks: -10%, Baseline, +10%; Revenue-at-Risk, Holding Cost)", 0.62, "#E8F8F5", 0.14),
    ("Inventory Intelligence & Digital Twin State Sandbox\n(Safety Stock SS, Reorder Point ROP, Required Stock RS, Status: Healthy / Stockout / Overstock)", 0.42, "#FEF9E7", 0.14),
    ("Predictive Modeling & Adaptive Selection Engine\n(ARIMA(1,1,1), Facebook Prophet, Out-of-Sample MAE/RMSE/MAPE Evaluation, Dynamic Router)", 0.22, "#F4ECF7", 0.14),
    ("Data Ingestion & Chronological Stratification Tier\n(Walmart M5 Dataset: 28,695 records, Store CA_1, Calendar & Sell Prices, 80/20 Chronological Split)", 0.02, "#F2F3F4", 0.14),
]

for title, y, color, height in tiers:
    rect = patches.FancyBboxPatch((0.05, y), 0.90, height, boxstyle="round,pad=0.015",
                                  edgecolor="#34495E", facecolor=color, linewidth=1.2)
    ax.add_patch(rect)
    ax.text(0.50, y + height/2.0, title, ha='center', va='center', fontsize=8.5,
            weight='bold', color='#1C2833', multialignment='center')

# Arrows connecting tiers
for y_arr in [0.16, 0.36, 0.56, 0.76]:
    ax.annotate('', xy=(0.50, y_arr + 0.055), xytext=(0.50, y_arr),
                arrowprops=dict(arrowstyle="->", color="#2C3E50", lw=1.8, mutation_scale=12))

plt.tight_layout()
fig.savefig(os.path.join(OUTPUT_DIR, 'fig1_system_architecture.png'), dpi=300, bbox_inches='tight')
plt.close(fig)
print("Generated Fig 1: System Architecture.")

# ==============================================================================
# FIGURE 2: Complete Project Workflow Flowchart
# ==============================================================================
fig, ax = plt.subplots(figsize=(7.2, 4.0), dpi=300)
ax.axis('off')

# Horizontal pipeline stages
stages = [
    ("Raw M5 Data\n& Filtering\n(Store CA_1)", 0.02, 0.15, "#D6EAF8"),
    ("Feature Merge\n& Preprocessing\n(Wide to Long)", 0.18, 0.15, "#D5F5E3"),
    ("80/20 Temporal\nTrain/Test Split\n(1,530 / 383 d)", 0.34, 0.15, "#FCF3CF"),
    ("Dual Model\nBenchmarking\n(ARIMA & Prophet)", 0.50, 0.15, "#E8DAEF"),
    ("Adaptive Model\nSelection Routing\n(MAPE/MAE/RMSE)", 0.66, 0.15, "#FADBD8"),
    ("Decision Matrix\n& Dashboard\n(Inv, Fin, UI)", 0.82, 0.15, "#EDBB99"),
]

for title, x, width, color in stages:
    rect = patches.FancyBboxPatch((x, 0.35), width, 0.30, boxstyle="round,pad=0.02",
                                  edgecolor="#2C3E50", facecolor=color, linewidth=1.2)
    ax.add_patch(rect)
    ax.text(x + width/2.0, 0.50, title, ha='center', va='center', fontsize=8,
            weight='bold', color='#1A252F', multialignment='center')

# Arrows between horizontal stages
for x_arr in [0.17, 0.33, 0.49, 0.65, 0.81]:
    ax.annotate('', xy=(x_arr + 0.008, 0.50), xytext=(x_arr - 0.002, 0.50),
                arrowprops=dict(arrowstyle="->", color="#1B2631", lw=1.5, mutation_scale=10))

# Sub-annotations below stages
sub_notes = [
    "3,049 SKUs filtered\nto 15 active SKUs",
    "Calendar & Sell Prices\nzero-sales: 0.97%",
    "No data leakage;\nstrict chronological",
    "ARIMA (1,1,1) &\nMeta Prophet (linear)",
    "Prophet: 11 SKUs (73%)\nARIMA: 4 SKUs (27%)",
    "FastAPI endpoints &\nReact 19 Recharts"
]
for i, note in enumerate(sub_notes):
    x_pos = stages[i][1] + stages[i][2]/2.0
    ax.text(x_pos, 0.18, note, ha='center', va='top', fontsize=7.5,
            style='italic', color='#34495E', multialignment='center')
    ax.plot([x_pos, x_pos], [0.35, 0.23], color="#7F8C8D", linestyle=":", lw=1.0)

plt.tight_layout()
fig.savefig(os.path.join(OUTPUT_DIR, 'fig2_project_workflow.png'), dpi=300, bbox_inches='tight')
plt.close(fig)
print("Generated Fig 2: Project Workflow.")

# ==============================================================================
# FIGURE 3: ARIMA versus Prophet Benchmark Comparison
# ==============================================================================
fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(7.2, 3.2), dpi=300)

# Panel A: Mean error metric comparison across the entire portfolio
metrics = ['MAE', 'RMSE', 'MAPE (%)']
arima_means = [models_df['ARIMA_MAE'].mean(), models_df['ARIMA_RMSE'].mean(), models_df['ARIMA_MAPE'].mean()]
prophet_means = [models_df['Prophet_MAE'].mean(), models_df['Prophet_RMSE'].mean(), models_df['Prophet_MAPE'].mean()]
adaptive_means = [
    np.mean([r[f"{r['Selected_Model']}_MAE"] for _, r in models_df.iterrows()]),
    np.mean([r[f"{r['Selected_Model']}_RMSE"] for _, r in models_df.iterrows()]),
    np.mean([r[f"{r['Selected_Model']}_MAPE"] for _, r in models_df.iterrows()]),
]

x = np.arange(len(metrics))
width = 0.26

ax1.bar(x - width, arima_means, width, label='ARIMA (1,1,1)', color='#3498DB', edgecolor='black', linewidth=0.6)
ax1.bar(x, prophet_means, width, label='Prophet', color='#2ECC71', edgecolor='black', linewidth=0.6)
ax1.bar(x + width, adaptive_means, width, label='Adaptive Ensemble', color='#E67E22', edgecolor='black', linewidth=0.6)

ax1.set_ylabel('Metric Value')
ax1.set_title('(a) Portfolio-Level Metric Comparison', weight='bold')
ax1.set_xticks(x)
ax1.set_xticklabels(metrics)
ax1.grid(axis='y', linestyle='--', alpha=0.5)
ax1.legend(frameon=True, facecolor='white', edgecolor='#BDC3C7')

# Values on top of bars
for i in range(len(metrics)):
    ax1.text(i - width, arima_means[i] + 1.0, f'{arima_means[i]:.1f}', ha='center', va='bottom', fontsize=7)
    ax1.text(i, prophet_means[i] + 1.0, f'{prophet_means[i]:.1f}', ha='center', va='bottom', fontsize=7)
    ax1.text(i + width, adaptive_means[i] + 1.0, f'{adaptive_means[i]:.1f}', ha='center', va='bottom', fontsize=7, weight='bold')
ax1.set_ylim(0, 70)

# Panel B: Win-loss distribution pie/donut
wins = [11, 4]
labels = ['Prophet Selected\n(11 SKUs, 73.3%)', 'ARIMA Selected\n(4 SKUs, 26.7%)']
colors = ['#2ECC71', '#3498DB']
explode = (0.04, 0.04)

wedges, texts, autotexts = ax2.pie(wins, labels=labels, autopct='%1.1f%%', startangle=140,
                                  colors=colors, explode=explode,
                                  wedgeprops=dict(width=0.45, edgecolor='black', linewidth=0.7),
                                  pctdistance=0.75)
for t in texts:
    t.set_fontsize(8)
for at in autotexts:
    at.set_fontsize(8)
    at.set_weight('bold')
ax2.set_title('(b) Model Selection Share across 15 SKUs', weight='bold')

plt.tight_layout()
fig.savefig(os.path.join(OUTPUT_DIR, 'fig3_model_comparison.png'), dpi=300, bbox_inches='tight')
plt.close(fig)
print("Generated Fig 3: Model Comparison.")

# ==============================================================================
# FIGURE 4: Actual vs Predicted Demand Curve
# ==============================================================================
# Representative SKU 1: FOODS_3_586 (Selected: Prophet, High Demand Stockout Risk)
# Representative SKU 2: FOODS_1_085 (Selected: ARIMA, Stable Seasonal Pattern)
fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(7.2, 4.4), dpi=300, sharex=False)

sku1 = 'FOODS_3_586'
sku1_hist = sales_df[sales_df['item_id'] == sku1].sort_values('date').tail(60)
sku1_fc = forecasts_df[forecasts_df['product_id'] == sku1].sort_values('date')

ax1.plot(sku1_hist['date'], sku1_hist['sales'], label='Historical Sales (Actual)', color='#2C3E50', lw=1.2)
ax1.plot(sku1_fc['date'], sku1_fc['predicted_demand'], label='30-Day Forecast (Prophet)', color='#E74C3C', lw=1.5)
ax1.fill_between(sku1_fc['date'], sku1_fc['lower_bound'], sku1_fc['upper_bound'],
                 color='#E74C3C', alpha=0.22, label='80% Confidence Interval')
ax1.axvline(sku1_hist['date'].max(), color='#7F8C8D', linestyle='--', lw=1.0, label='Forecast Horizon Boundary')
ax1.set_title(f'(a) SKU {sku1} — High Demand Volatility (Prophet Selected)', weight='bold')
ax1.set_ylabel('Daily Units Sold')
ax1.grid(True, linestyle=':', alpha=0.6)
ax1.legend(loc='upper right', frameon=True, facecolor='white', edgecolor='#BDC3C7', fontsize=7.5)

sku2 = 'FOODS_1_085'
sku2_hist = sales_df[sales_df['item_id'] == sku2].sort_values('date').tail(60)
sku2_fc = forecasts_df[forecasts_df['product_id'] == sku2].sort_values('date')

ax2.plot(sku2_hist['date'], sku2_hist['sales'], label='Historical Sales (Actual)', color='#2C3E50', lw=1.2)
ax2.plot(sku2_fc['date'], sku2_fc['predicted_demand'], label='30-Day Forecast (ARIMA)', color='#2980B9', lw=1.5)
ax2.fill_between(sku2_fc['date'], sku2_fc['lower_bound'], sku2_fc['upper_bound'],
                 color='#2980B9', alpha=0.22, label='80% Confidence Interval')
ax2.axvline(sku2_hist['date'].max(), color='#7F8C8D', linestyle='--', lw=1.0, label='Forecast Horizon Boundary')
ax2.set_title(f'(b) SKU {sku2} — Stable Autoregressive Pattern (ARIMA Selected)', weight='bold')
ax2.set_ylabel('Daily Units Sold')
ax2.set_xlabel('Date (Year-Month)')
ax2.grid(True, linestyle=':', alpha=0.6)
ax2.legend(loc='upper right', frameon=True, facecolor='white', edgecolor='#BDC3C7', fontsize=7.5)

plt.tight_layout()
fig.savefig(os.path.join(OUTPUT_DIR, 'fig4_forecast_vs_actual.png'), dpi=300, bbox_inches='tight')
plt.close(fig)
print("Generated Fig 4: Forecast vs Actual.")

# ==============================================================================
# FIGURE 5: Inventory Status Distribution & Stock Ledger
# ==============================================================================
fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(7.2, 3.2), dpi=300)

# Panel A: Portfolio Status Breakdown
status_counts = inv_df['inventory_status'].value_counts()
colors = {'HEALTHY STOCK': '#27AE60', 'OVERSTOCK': '#F39C12', 'STOCKOUT RISK': '#E74C3C'}
bar_colors = [colors.get(s, '#95A5A6') for s in status_counts.index]

ax1.bar(status_counts.index, status_counts.values, color=bar_colors, edgecolor='black', linewidth=0.7, width=0.55)
ax1.set_ylabel('Number of SKUs')
ax1.set_title('(a) Portfolio Inventory Health Breakdown', weight='bold')
ax1.set_xticklabels(['Healthy Stock\n(60%)', 'Overstock\n(26.7%)', 'Stockout Risk\n(13.3%)'], fontsize=8)
ax1.grid(axis='y', linestyle='--', alpha=0.5)
ax1.set_ylim(0, 11)
for i, v in enumerate(status_counts.values):
    ax1.text(i, v + 0.3, f'{v} SKUs', ha='center', va='bottom', fontsize=8, weight='bold')

# Panel B: Simulated Current Stock vs Required Stock across 15 SKUs
skus = inv_df['item_id'].str.replace('FOODS_', '')
x = np.arange(len(skus))
width = 0.38

ax2.bar(x - width/2, inv_df['simulated_current_stock'], width, label='Current Stock (Simulated)',
        color='#7FB3D5', edgecolor='black', linewidth=0.5)
ax2.bar(x + width/2, inv_df['required_stock'], width, label='Required Stock (Demand + SS)',
        color='#F1948A', edgecolor='black', linewidth=0.5)

ax2.set_ylabel('Units')
ax2.set_title('(b) On-Hand Stock vs. Target Buffer', weight='bold')
ax2.set_xticks(x)
ax2.set_xticklabels(skus, rotation=60, ha='right', fontsize=7)
ax2.grid(axis='y', linestyle='--', alpha=0.5)
ax2.legend(frameon=True, facecolor='white', edgecolor='#BDC3C7', fontsize=7.5)

plt.tight_layout()
fig.savefig(os.path.join(OUTPUT_DIR, 'fig5_inventory_status_distribution.png'), dpi=300, bbox_inches='tight')
plt.close(fig)
print("Generated Fig 5: Inventory Status Distribution.")

# ==============================================================================
# FIGURE 6: What-If Scenario Comparison
# ==============================================================================
fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(7.2, 3.0), dpi=300)

scenarios = ['Low Demand (-10%)', 'Normal Demand', 'High Demand (+10%)']
scen_labels = ['Low (-10%)', 'Baseline', 'High (+10%)']
whatif_summary = whatif_df.groupby('Scenario').agg({
    'Scenario_Forecasted_Demand': 'sum',
    'Required_Stock': 'sum',
    'Shortage_Quantity': 'sum',
    'Reorder_Quantity': 'sum'
}).loc[scenarios]

x = np.arange(len(scenarios))
width = 0.35

# Panel A: Aggregate Demand vs Required Stock
ax1.bar(x - width/2, whatif_summary['Scenario_Forecasted_Demand'], width,
        label='Scenario Demand', color='#3498DB', edgecolor='black', linewidth=0.6)
ax1.bar(x + width/2, whatif_summary['Required_Stock'], width,
        label='Required Buffer', color='#9B59B6', edgecolor='black', linewidth=0.6)
ax1.set_ylabel('Total Units')
ax1.set_title('(a) Projected Demand & Required Buffer', weight='bold')
ax1.set_xticks(x)
ax1.set_xticklabels(scen_labels)
ax1.grid(axis='y', linestyle='--', alpha=0.5)
ax1.legend(frameon=True, facecolor='white', edgecolor='#BDC3C7', fontsize=7.5)
ax1.set_ylim(0, 8500)

for i in range(len(scenarios)):
    dem = whatif_summary['Scenario_Forecasted_Demand'].iloc[i]
    req = whatif_summary['Required_Stock'].iloc[i]
    ax1.text(i - width/2, dem + 150, f'{dem:.0f}', ha='center', va='bottom', fontsize=7.5)
    ax1.text(i + width/2, req + 150, f'{req:.0f}', ha='center', va='bottom', fontsize=7.5)

# Panel B: Reorder Quantity Surge across Scenarios
ax2.plot(scen_labels, whatif_summary['Reorder_Quantity'], marker='o', color='#E74C3C',
         lw=2.0, ms=7, label='Recommended Reorder (PO)')
ax2.plot(scen_labels, whatif_summary['Shortage_Quantity'], marker='s', color='#F39C12',
         lw=1.5, ls='--', ms=6, label='Net Shortage Gap')
ax2.set_ylabel('Units to Procure')
ax2.set_title('(b) Procurement Exposure under Shocks', weight='bold')
ax2.grid(True, linestyle='--', alpha=0.5)
ax2.legend(frameon=True, facecolor='white', edgecolor='#BDC3C7', fontsize=7.5)
ax2.set_ylim(2000, 4200)

for i, txt in enumerate(whatif_summary['Reorder_Quantity']):
    ax2.annotate(f'{txt:.1f} u', (scen_labels[i], txt), textcoords="offset points",
                 xytext=(0, 8), ha='center', fontsize=8, weight='bold')

plt.tight_layout()
fig.savefig(os.path.join(OUTPUT_DIR, 'fig6_demand_scenario_comparison.png'), dpi=300, bbox_inches='tight')
plt.close(fig)
print("Generated Fig 6: Demand Scenario Comparison.")

# ==============================================================================
# FIGURE 7: Financial Impact Comparison
# ==============================================================================
fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(7.2, 3.0), dpi=300)

fin_summary = fin_df.groupby('Scenario').agg({
    'Forecasted_Revenue': 'sum',
    'Revenue_at_Risk': 'sum',
    'Overstock_Holding_Cost': 'sum'
}).loc[scenarios]

# Panel A: Revenue vs Revenue-at-Risk
x = np.arange(len(scenarios))
width = 0.35

ax1.bar(x - width/2, fin_summary['Forecasted_Revenue'], width,
        label='Forecasted Gross Revenue', color='#2ECC71', edgecolor='black', linewidth=0.6)
ax1.bar(x + width/2, fin_summary['Revenue_at_Risk'], width,
        label='Revenue-at-Risk (Stockout)', color='#E74C3C', edgecolor='black', linewidth=0.6)
ax1.set_ylabel('USD ($)')
ax1.set_title('(a) Revenue & Stockout Exposure', weight='bold')
ax1.set_xticks(x)
ax1.set_xticklabels(scen_labels)
ax1.grid(axis='y', linestyle='--', alpha=0.5)
ax1.legend(frameon=True, facecolor='white', edgecolor='#BDC3C7', fontsize=7.5)
ax1.set_ylim(0, 14000)

for i in range(len(scenarios)):
    rev = fin_summary['Forecasted_Revenue'].iloc[i]
    rar = fin_summary['Revenue_at_Risk'].iloc[i]
    ax1.text(i - width/2, rev + 250, f'${rev:,.0f}', ha='center', va='bottom', fontsize=7.5)
    ax1.text(i + width/2, rar + 250, f'${rar:,.0f}', ha='center', va='bottom', fontsize=7.5)

# Panel B: Overstock Holding Costs across Scenarios
ax2.bar(scen_labels, fin_summary['Overstock_Holding_Cost'], color='#F39C12',
        edgecolor='black', linewidth=0.6, width=0.45)
ax2.set_ylabel('Monthly Holding Cost ($)')
ax2.set_title('(b) Capital Tied in Overstock Holding', weight='bold')
ax2.grid(axis='y', linestyle='--', alpha=0.5)
ax2.set_ylim(0, 16)

for i, v in enumerate(fin_summary['Overstock_Holding_Cost']):
    ax2.text(i, v + 0.4, f'${v:.2f}', ha='center', va='bottom', fontsize=8, weight='bold')

plt.tight_layout()
fig.savefig(os.path.join(OUTPUT_DIR, 'fig7_financial_impact_comparison.png'), dpi=300, bbox_inches='tight')
plt.close(fig)
print("Generated Fig 7: Financial Impact Comparison.")

print("All 7 figures generated successfully in publication quality!")

