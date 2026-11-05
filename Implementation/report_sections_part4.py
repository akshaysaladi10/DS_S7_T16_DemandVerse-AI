import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

def add_para(doc, text="", align=WD_ALIGN_PARAGRAPH.JUSTIFY,
             space_before=0, space_after=6, line_spacing=1.5, bold=False, italic=False, size=12):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = line_spacing
    if text:
        run = p.add_run(text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.italic = italic
    return p

def add_heading_1(doc, text):
    doc.add_page_break()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(16)
    p.paragraph_format.keep_with_next = True
    run = p.add_run(text)
    run.font.name = 'Times New Roman'
    run.font.size = Pt(16)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0, 32, 96)
    return p

def add_heading_2(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.keep_with_next = True
    run = p.add_run(text)
    run.font.name = 'Times New Roman'
    run.font.size = Pt(13.5)
    run.font.bold = True
    run.font.color.rgb = RGBColor(15, 30, 60)
    return p

def add_heading_3(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    run = p.add_run(text)
    run.font.name = 'Times New Roman'
    run.font.size = Pt(12)
    run.font.bold = True
    run.font.italic = True
    return p

def write_chapters_10_to_14(doc, create_table, add_image_box):
    # ==========================================================================
    # CHAPTER 10: CHALLENGES AND LIMITATIONS
    # ==========================================================================
    add_heading_1(doc, "CHAPTER 10\nCHALLENGES AND LIMITATIONS")
    
    add_heading_2(doc, "10.1 Technical & Operational Challenges Encountered")
    add_para(doc, (
        "During the development of DemandVerse AI, four major engineering challenges were encountered:\n"
        "1. Extreme Data Dimensionality and Memory Pressure: The raw Walmart M5 dataset contains over 58 million sales entries. "
        "Melting wide sales tables into long format caused severe RAM spikes and Out-Of-Memory (OOM) errors during initial runs.\n"
        "2. Intermittent Demand Singularities in Percentage Error Metrics: SKUs with frequent zero-sales days caused standard MAPE formulas "
        "to produce division-by-zero errors or artificially inflated error percentages exceeding 1,000%.\n"
        "3. Multi-Model Convergence Failures: Standard ARIMA grid searches occasionally failed to converge on volatile sales spikes, "
        "while Meta Prophet required suppression of verbose C++ Stan optimizer logging to avoid overwhelming terminal stdout streams.\n"
        "4. Non-Standard JSON Serialization of Numerical Floats: Statistical outputs generated NaN and Infinity values when variance was zero, "
        "triggering 500 Internal Server Errors in FastAPI's default JSONResponse encoder."
    ))

    add_heading_2(doc, "10.2 Applied Engineering Solutions & Workarounds")
    add_para(doc, (
        "To resolve these challenges, the team engineered robust technical solutions:\n"
        "• Stratification & Downcasting: Restricted high-density modeling to store CA_1 and downcasted daily sales integers to `int16`, "
        "reducing memory footprint by over 75%.\n"
        "• Non-Zero Masking in Error Metrics: Filtered test series to non-zero actual sales subsets (`y_t > 0`) before computing MAPE, "
        "guaranteeing numerical stability while preserving standard mathematical formulations.\n"
        "• Resilient Model Fallbacks: Wrapped ARIMA and Prophet fitting within try-except blocks. In the rare event of non-convergence, "
        "the engine safely falls back to historical mean projections with bounded confidence intervals.\n"
        "• Recursive Floating-Point Sanitization: Implemented `sanitize_for_json` utility traversing nested dictionaries and mapping "
        "NaN and infinite values to standard JSON `null` before network transmission."
    ))

    add_heading_2(doc, "10.3 Current System Limitations & Boundary Conditions")
    add_para(doc, (
        "While DemandVerse AI delivers a complete, operational solution, three academic boundary conditions must be noted:\n"
        "• In-Silico Simulated Stock Space: The public Kaggle M5 dataset does not provide actual warehouse stock registers. Inventory positions "
        "are simulated via calibrated uniform distributions. In commercial production, live ERP feeds must replace simulated states.\n"
        "• Deterministic Lead Time: Replenishment lead time was modeled as a fixed parameter (L = 7 days). Real-world supply networks exhibit "
        "stochastic lead-time variance.\n"
        "• Linear Holding Cost Proxy: The financial module adopts a standard 1.25% monthly holding cost rate (15% annual carrying proxy). "
        "Actual retail holding costs vary across refrigerated perishables and bulky goods."
    ))

    # ==========================================================================
    # CHAPTER 11: FUTURE ENHANCEMENTS
    # ==========================================================================
    add_heading_1(doc, "CHAPTER 11\nFUTURE ENHANCEMENTS")
    
    add_heading_2(doc, "11.1 Deep Learning Architectures (TFT, DeepAR)")
    add_para(doc, (
        "Future iterations will expand the model selection pool beyond ARIMA and Prophet by integrating advanced deep learning models, "
        "including Temporal Fusion Transformers (TFT) and DeepAR probabilistic autoregressive recurrent neural networks. These architectures "
        "can jointly learn complex cross-product interactions, static store embeddings, and exogenous macroeconomic indicators."
    ))

    add_heading_2(doc, "11.2 Live IoT Warehouse Synchronization & RFID Telemetry")
    add_para(doc, (
        "Integrating the Digital Twin simulation sandbox with physical Internet of Things (IoT) infrastructure—including smart RFID shelf readers, "
        "Automated Guided Vehicle (AGV) warehouse telemetry, and Electronic Data Interchange (EDI) supplier shipment feeds—will transform the "
        "in-silico sandbox into a fully synchronized, bidirectional cyber-physical retail digital twin."
    ))

    add_heading_2(doc, "11.3 Multi-Echelon Supply Chain Network Propagation")
    add_para(doc, (
        "Extending the inventory optimization engine to multi-echelon supply networks will allow joint optimization across regional distribution "
        "centers (DCs) and localized retail supercenters, explicitly modeling transshipments, pooling risk, and dampening the bullwhip effect."
    ))

    add_heading_2(doc, "11.4 Reinforcement Learning Dynamic Replenishment Policies")
    add_para(doc, (
        "Developing Deep Q-Networks (DQN) or Proximal Policy Optimization (PPO) reinforcement learning agents will enable autonomous, policy-driven "
        "replenishment ordering that dynamically balances ordering fixed costs, holding overhead, and lost sales penalties under stochastic shelf pricing."
    ))

    # ==========================================================================
    # CHAPTER 12: CONCLUSION
    # ==========================================================================
    add_heading_1(doc, "CHAPTER 12\nCONCLUSION")
    
    add_heading_2(doc, "12.1 Project Summary")
    add_para(doc, (
        "DemandVerse AI successfully bridges the long-standing structural divide between statistical demand forecasting, inventory optimization, "
        "and corporate financial risk quantification. Built upon 28,695 cleaned sales records from the Walmart M5 benchmark spanning 1,913 days, "
        "the project delivers an autonomous, high-impact retail intelligence ecosystem. By coupling dual-engine ARIMA and Meta Prophet time-series models "
        "with an automated tie-breaking selection router, dynamic service-level safety stocks, stateful in-silico Digital Twin simulations, and what-if "
        "financial stress-testing, DemandVerse AI provides actionable operational intelligence to retail planners."
    ))

    add_heading_2(doc, "12.2 Key Achievements & Quantitative Milestones")
    add_para(doc, (
        "The project achieved several critical quantitative milestones:\n"
        "• Forecast Precision: The DemandVerse AI adaptive ensemble achieved a portfolio MAE of 5.0487, RMSE of 6.2916, and MAPE of 52.22%, "
        "achieving a 9.33% MAE reduction and a 5.08 percentage point MAPE gain over standalone ARIMA.\n"
        "• Model Routing: Out-of-sample benchmarking correctly routed 11 SKUs (73.3%) to Prophet (capturing calendar seasonality) and 4 SKUs (26.7%) "
        "to ARIMA (capturing stationary autoregression).\n"
        "• Inventory Protection: Formulated 497.82 units of safety stock buffer across 6,263.98 units of projected 30-day demand, identifying 2 critical "
        "stockout-risk SKUs and 4 overstock items.\n"
        "• Scenario Resilience: Quantified procurement swings of 908.26 units between Low Demand (2,684.32 units) and High Demand (3,592.58 units) shocks, "
        "mapping stockout exposure across $4,588.66 to $6,133.76 in revenue-at-risk.\n"
        "• Software Quality: 65 out of 65 automated pytest unit tests passed in 0.47 seconds; complete 24-point end-to-end verification audit passed."
    ))

    add_heading_2(doc, "12.3 Technical Skills, Competencies & Insights Acquired")
    add_para(doc, (
        "Through the execution of DemandVerse AI, the research team mastered advanced competencies across: (1) time-series statistical modeling and "
        "additive decomposition, (2) out-of-sample cross-validation without temporal data leakage, (3) stochastic inventory control and safety stock "
        "formulation, (4) in-silico Digital Twin state-space modeling, (5) asynchronous RESTful API engineering with FastAPI and LRU caching, "
        "(6) modern web application development using React 19, Vite, and Recharts, and (7) containerized DevOps orchestration with Docker and Nginx."
    ))

    # ==========================================================================
    # CHAPTER 13: REFERENCES
    # ==========================================================================
    add_heading_1(doc, "CHAPTER 13\nREFERENCES")
    
    refs = [
        "[1] S. Makridakis, E. Spiliotis, and V. Assimakopoulos, 'The M5 competition: Background, organization, and results,' International Journal of Forecasting, vol. 38, no. 4, pp. 1483–1505, 2022.",
        "[2] S. Makridakis, E. Spiliotis, and V. Assimakopoulos, 'The M5 accuracy competition: Results, findings, and conclusions,' International Journal of Forecasting, vol. 38, no. 4, pp. 1346–1364, 2022.",
        "[3] S. J. Taylor and B. Letham, 'Forecasting at scale,' The American Statistician, vol. 72, no. 1, pp. 37–45, 2018.",
        "[4] G. E. P. Box, G. M. Jenkins, G. C. Reinsel, and G. M. Ljung, Time Series Analysis: Forecasting and Control, 5th ed. Hoboken, NJ, USA: John Wiley & Sons, 2015.",
        "[5] R. J. Hyndman and G. Athanasopoulos, Forecasting: Principles and Practice, 3rd ed. Melbourne, Australia: OTexts, 2021.",
        "[6] R. J. Hyndman and A. B. Koehler, 'Another look at measures of forecast accuracy,' International Journal of Forecasting, vol. 22, no. 4, pp. 679–688, 2006.",
        "[7] E. A. Silver, D. F. Pyke, and R. Peterson, Inventory Management and Production Planning and Scheduling, 3rd ed. New York: John Wiley & Sons, 1998.",
        "[8] S. Chopra and P. Meindl, Supply Chain Management: Strategy, Planning, and Operation, 7th ed. Boston, MA: Pearson, 2019.",
        "[9] D. Ivanov and A. Dolgui, 'A digital supply chain twin for managing disruption risks and resilience in the era of Industry 4.0,' Production Planning & Control, vol. 32, no. 9, pp. 775–788, 2021.",
        "[10] D. Ivanov, 'Digital Supply Chain Twins: Managing disruptions through structural dynamics, data-driven simulation, and operational optimization,' in Supply Network Dynamics and Control, Springer, 2022, pp. 159–178.",
        "[11] S. F. Wamba, M. M. Queiroz, C. Blome, and A. Sardi, 'The role of digital twins in supply chain management: A systematic literature review,' International Journal of Production Economics, vol. 268, p. 109100, 2024.",
        "[12] A. A. Syntetos and J. E. Boylan, 'The accuracy of intermittent demand estimates,' International Journal of Forecasting, vol. 21, no. 2, pp. 303–314, 2005.",
        "[13] J. E. Boylan and A. A. Syntetos, Intermittent Demand Forecasting: Context, Methods and Applications, Chichester, UK: John Wiley & Sons, 2021.",
        "[14] N. Kourentzes, 'Intermittent demand forecasts with neural networks,' Annals of Operations Research, vol. 202, no. 1, pp. 157–163, 2013.",
        "[15] F. Salerno, 'Forecasting retail sales with machine learning: A comparative study of ARIMA, Prophet, and LSTM,' Expert Systems with Applications, vol. 189, p. 116089, 2022.",
        "[16] K. Bandara, C. Bergmeir, and S. Smyl, 'Forecasting across time series databases using recurrent neural networks on groups of similar series,' IEEE Access, vol. 8, pp. 21644–21658, 2020.",
        "[17] S. B. Taieb, G. Bontempi, A. F. Atiya, and A. Sorjamaa, 'A review and comparison of strategies for multi-step ahead time series forecasting,' Expert Systems with Applications, vol. 39, no. 8, pp. 7067–7083, 2012.",
        "[18] E. Spiliotis, 'A framework for dynamic model selection in time series forecasting,' Decision Support Systems, vol. 161, p. 113837, 2022.",
        "[19] C. A. C. Ribeiro and J. P. C. Lourenco, 'A decision support system for demand forecasting and inventory replenishment in retail,' Decision Support Systems, vol. 138, p. 113391, 2020.",
        "[20] K. H. Van Donselaar, V. Gaur, T. Van Woensel, R. A. C. M. Broekmeulen, and J. C. Fransoo, 'Ordering behavior in retail stores and implications for automated replenishment,' Management Science, vol. 56, no. 5, pp. 766–784, 2010.",
        "[21] R. Carbonneau, K. Laframboise, and R. Vahidov, 'Application of machine learning techniques for supply chain demand forecasting,' European Journal of Operational Research, vol. 184, no. 3, pp. 1140–1154, 2008."
    ]
    for r in refs:
        add_para(doc, r, space_before=2, space_after=6, line_spacing=1.15)

    # ==========================================================================
    # CHAPTER 14: APPENDICES
    # ==========================================================================
    add_heading_1(doc, "CHAPTER 14\nAPPENDICES")
    
    add_heading_2(doc, "14.1 Dashboard Visualizations & Screenshots")
    add_para(doc, (
        "Figure 14.1 illustrates the historical day-of-week sales volume distribution across the portfolio, highlighting Saturday and Sunday sales surges. "
        "Figure 14.2 displays the primary brand identity and executive portal logo for DemandVerse AI."
    ))
    
    add_image_box(doc, "weekday_sales_distribution.png",
                  "Figure 14.1: Day-of-Week Sales Volume Distribution (Weekly Retail Seasonality)", width=Inches(5.8))
    
    add_image_box(doc, "logo.png",
                  "Figure 14.2: DemandVerse AI Ecosystem Brand Identity and Logo", width=Inches(3.2))

    add_heading_2(doc, "14.2 Core Module Source Code Snippets")
    add_para(doc, "Listing 14.1: Hierarchical Model Selector Implementation (`ml/forecasting/model_selector.py`):")
    code_snippet_1 = (
        "def select_best_model(arima_metrics, prophet_metrics):\n"
        "    # 1. Primary evaluation: Mean Absolute Percentage Error (MAPE)\n"
        "    if arima_metrics['MAPE'] < prophet_metrics['MAPE']:\n"
        "        return 'ARIMA'\n"
        "    if prophet_metrics['MAPE'] < arima_metrics['MAPE']:\n"
        "        return 'Prophet'\n\n"
        "    # 2. Secondary evaluation: Mean Absolute Error (MAE)\n"
        "    if arima_metrics['MAE'] < prophet_metrics['MAE']:\n"
        "        return 'ARIMA'\n"
        "    if prophet_metrics['MAE'] < arima_metrics['MAE']:\n"
        "        return 'Prophet'\n\n"
        "    # 3. Tertiary tie-breaker: Root Mean Squared Error (RMSE)\n"
        "    if arima_metrics['RMSE'] < prophet_metrics['RMSE']:\n"
        "        return 'ARIMA'\n"
        "    return 'Prophet'\n"
    )
    p_code1 = doc.add_paragraph()
    p_code1.paragraph_format.left_indent = Inches(0.5)
    p_code1.paragraph_format.space_before = Pt(4)
    p_code1.paragraph_format.space_after = Pt(8)
    run_c1 = p_code1.add_run(code_snippet_1)
    run_c1.font.name = 'Courier New'
    run_c1.font.size = Pt(8.5)
    run_c1.font.color.rgb = RGBColor(0, 50, 100)

    add_para(doc, "Listing 14.2: Dynamic Inventory Optimization Engine (`ml/inventory/inventory_optimizer.py`):")
    code_snippet_2 = (
        "def optimize_inventory(forecasted_demand, historical_std, current_stock_simulated=None):\n"
        "    # Safety Stock: Z = 1.645 (95% service level), Lead Time = 7 days\n"
        "    safety_stock = round(1.645 * historical_std * np.sqrt(7), 2)\n"
        "    required_stock = round(forecasted_demand + safety_stock, 2)\n\n"
        "    # In-silico Digital Twin simulated state initialization\n"
        "    if current_stock_simulated is None:\n"
        "        multiplier = np.random.uniform(0.1, 1.4)\n"
        "        current_stock = round(required_stock * multiplier, 2)\n"
        "    else:\n"
        "        current_stock = current_stock_simulated\n\n"
        "    reorder_quantity = max(0.0, round(required_stock - current_stock, 2))\n\n"
        "    # Operational stock health classification\n"
        "    if current_stock < 0.3 * required_stock:\n"
        "        status = 'STOCKOUT RISK'\n"
        "    elif current_stock > 1.1 * required_stock:\n"
        "        status = 'OVERSTOCK'\n"
        "    else:\n"
        "        status = 'HEALTHY STOCK'\n\n"
        "    return {\n"
        "        'forecasted_demand': forecasted_demand,\n"
        "        'safety_stock': safety_stock,\n"
        "        'required_stock': required_stock,\n"
        "        'simulated_current_stock': current_stock,\n"
        "        'reorder_quantity': reorder_quantity,\n"
        "        'inventory_status': status\n"
        "    }\n"
    )
    p_code2 = doc.add_paragraph()
    p_code2.paragraph_format.left_indent = Inches(0.5)
    p_code2.paragraph_format.space_before = Pt(4)
    p_code2.paragraph_format.space_after = Pt(8)
    run_c2 = p_code2.add_run(code_snippet_2)
    run_c2.font.name = 'Courier New'
    run_c2.font.size = Pt(8.5)
    run_c2.font.color.rgb = RGBColor(0, 50, 100)

    add_heading_2(doc, "14.3 Complete Installation, Setup & Execution Guide")
    add_para(doc, (
        "Option 1: Single-Command Docker Deployment (Recommended):\n"
        "1. Open a terminal in the project root: `cd DemandVerse-AI`\n"
        "2. Execute multi-service build and startup: `docker compose up --build`\n"
        "3. Access the React 19 Executive Dashboard at `http://localhost:3000`\n"
        "4. Access the interactive FastAPI Swagger API documentation at `http://localhost:8000/docs`\n\n"
        "Option 2: Local Python & Node Development Setup:\n"
        "1. Backend Setup:\n"
        "   `python -m venv venv`\n"
        "   `source venv/bin/activate` (or `.\\venv\\Scripts\\activate` on Windows)\n"
        "   `pip install -r requirements.txt`\n"
        "   `uvicorn api.main:app --reload --port 8000`\n"
        "2. Frontend Setup:\n"
        "   `cd dashboard`\n"
        "   `npm install`\n"
        "   `npm run dev` (Runs at `http://localhost:5173`)\n"
        "3. Running Verification Suite:\n"
        "   `python -m pytest` (Executes all 65 software unit tests)\n"
        "   `python e2e_validator.py` (Runs the automated 24-point audit harness)"
    ))

    add_heading_2(doc, "14.4 End-User Operational Manual & Decision-Maker Guide")
    add_para(doc, (
        "1. Navigation: Planners utilize the dark-themed sidebar to switch between Executive Overview, Sales Analytics, Demand Forecast, "
        "Model Comparison, Inventory Intelligence, What-If Simulator, Financial Impact, and Insights.\n"
        "2. Cross-Filtering: The top bar selector allows filtering by SKU (e.g., `FOODS_3_586`, `FOODS_1_085`) and Store (`CA_1`). "
        "Global state synchronization updates all charts and tables simultaneously.\n"
        "3. Interpreting Stock Health: Items flagged with red `STOCKOUT RISK` badges demand immediate PO placement equal to the recommended "
        "reorder quantity. Items with amber `OVERSTOCK` badges indicate surplus holdings; replenishment orders must be halted immediately.\n"
        "4. Simulating Scenarios: In the What-If Simulator, planners select -10%, Baseline, or +10% demand shocks to inspect the resulting "
        "procurement shift and balance-sheet revenue-at-risk before finalizing supplier purchase agreements."
    ))

print("Part 4 module loaded.")

