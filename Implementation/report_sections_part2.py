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

def write_chapters_4_to_6(doc, create_table, add_image_box):
    # ==========================================================================
    # CHAPTER 4: SYSTEM ARCHITECTURE
    # ==========================================================================
    add_heading_1(doc, "CHAPTER 4\nSYSTEM ARCHITECTURE")
    
    add_heading_2(doc, "4.1 High-Level Architecture Diagram")
    add_para(doc, (
        "The DemandVerse AI ecosystem is engineered as a decoupled, multi-tier intelligence pipeline comprising five functional layers: "
        "(1) Data Ingestion and Stratification Tier, (2) Machine Learning & Dual-Engine Forecasting Tier, (3) Inventory Intelligence & "
        "Digital Twin State Sandbox, (4) Financial Quantification & Scenario Stress-Testing Engine, and (5) Application Serving & "
        "Dashboard Presentation Tier. Figure 4.1 depicts the high-level architecture diagram of the system."
    ))
    
    add_image_box(doc, "fig1_system_architecture.png",
                  "Figure 4.1: DemandVerse AI High-Level Multi-Tier System Architecture", width=Inches(6.0))
    
    add_heading_2(doc, "4.2 Description of Each Layer")
    add_para(doc, (
        "1. Data Ingestion & Engineering Layer:\n"
        "Responsible for acquiring raw Walmart M5 data files (sales_train_validation.csv, calendar.csv, sell_prices.csv). "
        "Executes memory-efficient wide-to-long reshaping, merges calendar events and SNAP benefit indicators, joins weekly selling "
        "prices using composite keys (store_id, item_id, wm_yr_wk), imputes missing event flags and inactive prices, audits duplicates, "
        "and partitions data into an 80/20 chronological train/test split.\n\n"
        "2. Machine Learning & Forecasting Layer:\n"
        "Implements statistical ARIMA(1,1,1) and additive Meta Prophet models. Both models are trained on the 80% chronological training "
        "horizon (1,530 days) and tested over the 20% validation split (383 days). The layer outputs out-of-sample MAE, RMSE, and MAPE metrics.\n\n"
        "3. Adaptive Model Selector Router:\n"
        "Inspects out-of-sample error metrics and routes each SKU to its optimal engine using a deterministic tie-breaking hierarchy (MAPE -> MAE -> RMSE). "
        "The winning model is retrained on all 1,913 historical days to project the final 30-day forward demand distribution with 80% confidence bands.\n\n"
        "4. Inventory Intelligence & Digital Twin Sandbox:\n"
        "Translates demand forecasts and historical sales variance into dynamic Safety Stock (SS, Z=1.645), Reorder Points (ROP), and Required Stock (RS). "
        "Maintains an in-silico Digital Twin simulation state representing on-hand inventory position, generating automated Purchase Order (PO) "
        "reorder quantities and classifying SKU health into Healthy Stock, Overstock, and Stockout Risk.\n\n"
        "5. Scenario Sensitivity & Financial Quantification Engine:\n"
        "Stress-tests the inventory buffer across Low Demand (-10%), Baseline, and High Demand (+10%) regimes. Converts unit-level shortages and "
        "surpluses into balance-sheet monetary indicators: Forecasted Gross Revenue, Stockout Revenue-at-Risk, and Monthly Overstock Holding Costs (1.25% monthly rate).\n\n"
        "6. Application & Presentation Layer:\n"
        "Exposes all pre-calculated analytical outputs through an asynchronous FastAPI RESTful service with sub-10ms LRU in-memory caching. "
        "Renders responsive executive analytics across 9 dedicated feature pages in a React 19 single-page dashboard utilizing Recharts."
    ))

    add_heading_2(doc, "4.3 Deployment Architecture")
    add_para(doc, (
        "DemandVerse AI is containerized using Docker and Docker Compose to guarantee platform-independent reproducibility. "
        "The deployment architecture consists of two isolated, communicating container services:\n"
        "• Backend Container (`demandverse-api`): Built on `python:3.11-slim`, running FastAPI and Uvicorn on internal port 8000. Incorporates "
        "CORS middleware allowing authorized cross-origin requests, recursive floating-point sanitizers (replacing NaN/Inf with null), and "
        "dedicated `/health` readiness probes.\n"
        "• Frontend Container (`demandverse-dashboard`): Multi-stage build compiling Vite 8.3 React 19 assets into static production bundles, "
        "which are served via an ultra-lightweight `nginx:alpine` image on external port 3000. Nginx is configured with reverse-proxy rules "
        "forwarding `/api/*` traffic to the backend container and handling client-side SPA routing fallbacks via `try_files $uri /index.html`."
    ))

    # ==========================================================================
    # CHAPTER 5: DESIGN
    # ==========================================================================
    add_heading_1(doc, "CHAPTER 5\nDESIGN")
    
    add_heading_2(doc, "5.1 System Wireframes, UI Mockups & Data Flow Diagrams")
    add_para(doc, (
        "The system design of DemandVerse AI is articulated through Data Flow Diagrams (DFDs) tracing information exchange across boundaries. "
        "Figure 5.1 illustrates the Level-0 Context Diagram, showing primary external entities (Store Planners, Supplier Logistics, and M5 Data Sources) "
        "interacting with the central DemandVerse AI ecosystem. Figure 5.2 provides the decomposed Level-1 DFD detailing internal data stores "
        "and computational transformations."
    ))
    
    add_image_box(doc, "dfd_level_0.png",
                  "Figure 5.1: Data Flow Diagram (DFD) Level-0 — Context Level Architecture", width=Inches(5.8))
    
    add_image_box(doc, "dfd_level_1.png",
                  "Figure 5.2: Data Flow Diagram (DFD) Level-1 — Decomposed Functional Processing Pipeline", width=Inches(6.0))

    add_heading_2(doc, "5.2 Entity-Relationship (ER) Diagram & Schema Architecture")
    add_para(doc, (
        "DemandVerse AI maintains relational integrity across its analytical data lake using natural composite keys. "
        "Figure 5.3 shows the Entity-Relationship (ER) diagram capturing entities: Products Catalog, Calendar Attributes, Sell Prices, "
        "Clean Sales, Model Comparison Benchmarks, and the Inventory Ledger."
    ))
    
    add_image_box(doc, "er_diagram.png",
                  "Figure 5.3: Entity-Relationship (ER) & Relational Schema Architecture", width=Inches(6.0))

    add_heading_2(doc, "5.3 End-to-End Methodological Workflow")
    add_para(doc, (
        "The operational lifecycle of DemandVerse AI follows a rigorous seven-stage methodological pipeline, depicted in Figure 5.4: "
        "Raw Data Ingestion -> Feature Merging -> Chronological 80/20 Splitting -> Dual-Engine Benchmarking -> Adaptive Model Selection -> "
        "Inventory & Scenario Quantification -> Serving & Executive Visualization."
    ))
    
    add_image_box(doc, "fig2_project_workflow.png",
                  "Figure 5.4: DemandVerse AI End-to-End Operational and Methodological Workflow", width=Inches(6.0))

    # ==========================================================================
    # CHAPTER 6: IMPLEMENTATION
    # ==========================================================================
    add_heading_1(doc, "CHAPTER 6\nIMPLEMENTATION")
    
    add_heading_2(doc, "6.1 Module-Wise Implementation Architecture")
    add_para(doc, (
        "The project repository is structured into modular, decoupled directories adhering to modern software engineering best practices:\n"
        "• `api/`: FastAPI web server, route controllers (`routers/`), dependency injection providers (`dependencies.py`), sanitizers (`utils.py`), and test suite (`tests/test_api.py`).\n"
        "• `ml/`: Offline computational modules for ARIMA (`forecasting/arima_model.py`), Prophet (`forecasting/prophet_model.py`), Model Selector (`forecasting/model_selector.py`), "
        "Inventory Optimizer (`inventory/inventory_optimizer.py`), What-If Scenarios (`scenarios/what_if.py`), Financial Impact (`financial/financial_impact.py`), and Business Insights (`insights/business_insights.py`).\n"
        "• `dashboard/`: React 19 SPA source code (`src/`), feature page controllers (`features/`), visualizer components (`components/`), Context state providers (`context/`), and Axios API clients (`services/api.js`).\n"
        "• `demandverse_ai_outputs/`: Pre-computed, clean analytical artifacts (`data/clean_sales.csv`, `forecasts/final_forecasts.csv`, `inventory/inventory_recommendations.csv`, `financial/financial_impact.csv`).\n"
        "• `m5-forecasting-accuracy/`: Raw benchmark data directory containing Walmart sales, calendar, and price tables."
    ))

    add_heading_2(doc, "6.2 Front-End Logic and Component Architecture")
    add_para(doc, (
        "The React 19 frontend is engineered around an atomic component architecture. Key implementation modules include:\n"
        "• `DashboardContext.jsx`: Provides global state synchronization across the application. When a user selects a specific product or store "
        "in the top bar, `DashboardContext` broadcasts the filter state to all nine feature pages simultaneously without prop-drilling.\n"
        "• Responsive Visualizers: Components such as `StockHealthGauge.jsx` render dynamic radial and semi-circular gauges depicting inventory health "
        "(Healthy Stock in emerald green, Overstock in amber, Stockout Risk in crimson red). Chart tooltips (`ChartTooltip.jsx`) format dates, currencies, "
        "and confidence intervals cleanly upon hover.\n"
        "• Nine Analytical Feature Views: Executive Dashboard, Sales Analytics, Demand Forecast, Model Comparison, Inventory Intelligence, "
        "What-If Simulator, Financial Impact, Strategic Insights, and System Observability."
    ))

    add_heading_2(doc, "6.3 Back-End Logic and Route Controllers")
    add_para(doc, (
        "The FastAPI backend implements modular APIRouters mounted onto the primary application (`main.py`). Data ingestion functions in "
        "`dependencies.py` utilize the `@functools.lru_cache(maxsize=1)` decorator. Because the analytical CSV artifacts are static and pre-validated, "
        "in-memory LRU caching eliminates redundant disk I/O, allowing endpoints to serve complex analytical payloads in under 10 milliseconds. "
        "Route controllers validate request query parameters (e.g., `product_id`, `store_id`), filter cached DataFrames, and serialize results into JSON."
    ))

    add_heading_2(doc, "6.4 Data Ingestion, Cleaning & Chronological 80/20 Splitting")
    add_para(doc, (
        "The data preparation pipeline implemented in `DS_Project_Implementation_2.ipynb` executes six rigorous sequential operations:\n"
        "1. Wide-to-Long Reshaping: The horizontal sales matrix ($d_1$ to $d_{1913}$) is melted into a vertical tidy DataFrame, downcasting sales to `int16`.\n"
        "2. Calendar Feature Merge: Calendar attributes, SNAP assistance indicators (`snap_CA`), and holiday event names (`event_name_1`) are merged on day key `d`.\n"
        "3. Sell Price Enactment: Sell prices are joined using composite keys (`store_id`, `item_id`, `wm_yr_wk`).\n"
        "4. Missing Value & Anomaly Audits: Missing event names are imputed with `'None'`. Missing shelf prices during inactive shelf weeks are filled "
        "with the product's historical mean shelf price. Duplicate rows on `(date, store_id, item_id)` are verified at zero.\n"
        "5. Chronological Train/Test Partitioning: To guarantee zero forward lookahead leakage, data is partitioned strictly chronologically per SKU: "
        "Training split = first 1,530 days (80.0%, 2011-01-29 to 2015-04-07); Testing split = remaining 383 days (20.0%, 2015-04-08 to 2016-04-24)."
    ))

    add_heading_2(doc, "6.5 RESTful API Integration & Endpoint Specifications")
    add_para(doc, "Table 6.1 specifies the complete RESTful API contract exposed by the DemandVerse AI FastAPI backend.")
    
    api_headers = ["HTTP Method", "Endpoint URI", "Controller Function", "Payload Description & Responsibilities"]
    api_data = [
        ["GET", "/health", "health_check()", "System health probe returning service status and version for Docker readiness"],
        ["GET", "/api/products", "get_products()", "Returns distinct product IDs, departments, categories, and store metadata"],
        ["GET", "/api/sales/history", "get_sales_history()", "Historical daily sales records filtered by product_id and store_id"],
        ["GET", "/api/sales/weekday-distribution", "get_weekday_distribution()", "Aggregated sales volume grouped by day of week (seasonality)"],
        ["GET", "/api/forecasts", "get_forecasts()", "30-day predicted demand series with upper/lower 80% confidence bounds"],
        ["GET", "/api/forecasts/models", "get_model_comparison()", "Out-of-sample MAE, RMSE, MAPE metrics and selected winning model"],
        ["GET", "/api/inventory/recommendations", "get_inventory_recommendations()", "Forecasted demand, safety stock, ROP, stock position, health status"],
        ["GET", "/api/inventory/what-if", "get_what_if_analysis()", "Scenario demand, required stock, shortages, and reorders (-10%, Normal, +10%)"],
        ["GET", "/api/financial/impact", "get_financial_impact()", "Forecasted revenue, stockout revenue-at-risk, and overstock holding costs"],
        ["GET", "/api/insights/summary", "get_business_insights()", "AI-synthesized priority action cards and procurement recommendations"]
    ]
    col_w_api = [Inches(0.8), Inches(1.8), Inches(1.4), Inches(2.0)]
    create_table(doc, api_headers, api_data, col_w_api)

    add_heading_2(doc, "6.6 Security, Numerical Sanitization & CORS Configuration")
    add_para(doc, (
        "• Numerical Sanitization: Standard JSON serializers crash upon encountering IEEE 754 floating-point `NaN`, `+Infinity`, or `-Infinity`. "
        "The backend implements recursive sanitizers (`api/utils.py:sanitize_for_json`) that recursively traverse dictionaries and lists, replacing "
        "illegal floating-point values with `None` (serialized as standard JSON `null`).\n"
        "• CORS Middleware: Configured via Starlette `CORSMiddleware` in `api/main.py` with explicit allow-origins, methods, and headers to enable "
        "seamless communication between the React 19 dashboard on port 3000 and the FastAPI backend on port 8000.\n"
        "• Health Probes: The `/health` endpoint enables container orchestrators (Docker Compose, Kubernetes) to monitor backend service readiness."
    ))

print("Part 2 module loaded.")
