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

def write_chapters_1_to_3(doc, create_table):
    # ==========================================================================
    # CHAPTER 1: INTRODUCTION
    # ==========================================================================
    add_heading_1(doc, "CHAPTER 1\nINTRODUCTION")
    
    add_heading_2(doc, "1.1 Background of the Project")
    p1 = (
        "Modern omni-channel retail supply chains operate in an environment characterized by extreme demand volatility, "
        "intricate promotional cadences, seasonal demand shifts, and sudden macroeconomic shocks. The fundamental objective "
        "of enterprise retail operations is to satisfy customer demand seamlessly while minimizing total operating costs. "
        "However, retail replenishment managers routinely face the dual operational hazards of understocking and overstocking. "
        "Understocking precipitates stockouts, resulting in unrecoverable immediate revenue loss, diminished consumer satisfaction, "
        "brand degradation, and contractual penalties. According to global retail research, stockouts and overstocks—collectively "
        "termed inventory distortion—account for over $1.1 trillion in annual losses worldwide."
    )
    add_para(doc, p1)
    
    p2 = (
        "Conversely, overstocking sequesters valuable corporate working capital in non-productive inventory assets, severely inflating "
        "warehouse storage, climate control, handling, insurance, and obsolescence costs. Overstocked perishable grocery goods and "
        "seasonal merchandise inevitably suffer severe margin erosion through emergency clearance markdowns, product spoilage, or "
        "total write-offs. Despite substantial investments in modern Enterprise Resource Planning (ERP) systems, the majority of retail "
        "organizations continue to operate with fragmented, siloed analytical workflows. Time-series demand forecasting is frequently "
        "treated as a purely mathematical exercise, executed in isolation from inventory buffer sizing, supplier replenishment lead-time "
        "constraints, and executive financial balance-sheet evaluations."
    )
    add_para(doc, p2)
    
    p3 = (
        "DemandVerse AI was conceived to eliminate this operational fragmentation by engineering an integrated, autonomous, end-to-end "
        "retail intelligence ecosystem. By transforming historical transactional retail logs into forward-looking predictive demand "
        "trajectories, dynamically routing individual stock-keeping units (SKUs) to optimal machine learning algorithms, calculating "
        "statistically grounded safety stock buffers, simulating on-hand inventory states within an in-silico Digital Twin sandbox, "
        "and evaluating the financial exposure of demand shocks, DemandVerse AI provides a complete decision-support system for enterprise retail planners."
    )
    add_para(doc, p3)

    add_heading_2(doc, "1.2 Problem Statement")
    p4 = (
        "Traditional retail demand forecasting and inventory replenishment mechanisms exhibit five critical systemic deficiencies:\n"
        "1. Decoupling of Forecasting from Inventory Operations: Mathematical forecasting tools output unit predictions without translating "
        "them into actionable procurement parameters such as Safety Stock, Reorder Points (ROP), and Purchase Order (PO) quantities.\n"
        "2. Monolithic Single-Model Assignment: Enterprise systems routinely enforce a single global forecasting model across all SKUs, failing "
        "to accommodate the heterogeneous demand characteristics of retail portfolios where seasonal items require additive Fourier models "
        "while fast-moving staples require autoregressive formulations.\n"
        "3. Static and Outdated Buffer Calculations: Conventional ERP software calculates safety stock using annualized static variance, ignoring "
        "forward-looking demand trajectories and supplier lead-time risk exposure.\n"
        "4. Absence of Pre-Shock What-If Stress-Testing: Retail managers typically discover inventory imbalances retrospectively after stockouts occur, "
        "lacking an interactive simulation environment to stress-test stock resilience against hypothetical demand contractions or surges.\n"
        "5. Gap Between Operational Units and Financial Balance Sheets: Forecast errors are reported in abstract statistical metrics (e.g., RMSE) "
        "without quantifying balance-sheet monetary exposure, such as gross revenue-at-risk and surplus carrying cost drag."
    )
    add_para(doc, p4)

    add_heading_2(doc, "1.3 Scope of the Project")
    p5 = (
        "The scope of DemandVerse AI encompasses the end-to-end lifecycle of retail predictive analytics and inventory intelligence:\n"
        "• Data Ingestion & Engineering: Ingestion and processing of multi-year transactional sales, calendar events, SNAP food assistance "
        "flags, and weekly selling prices from the Walmart M5 Forecasting Accuracy benchmark across 1,913 days (2011–2016).\n"
        "• High-Density Stratification: Focused empirical analysis on 15 high-volume, non-sparse SKUs within Walmart store CA_1 (California), "
        "comprising 28,695 cleaned records with an ultra-low zero-sales frequency of 0.97%.\n"
        "• Dual-Engine Predictive Modeling: Implementation and out-of-sample benchmarking of statistical ARIMA(1,1,1) and additive Meta Prophet models.\n"
        "• Dynamic Model Selection: Formulation of a multi-tier tie-breaking routing algorithm evaluating out-of-sample MAE, RMSE, and MAPE.\n"
        "• Dynamic Inventory Control: Automatic calculation of variance-aware Safety Stock (Z=1.645, 95% service level), Reorder Points, and Purchase Orders.\n"
        "• Digital Twin In-Silico State Simulation: Modeling on-hand inventory states and classifying SKU health into Healthy Stock, Overstock, and Stockout Risk.\n"
        "• What-If Sensitivity Engine: Multi-scenario demand shock simulation (-10% Low Demand, Baseline, +10% High Demand).\n"
        "• Financial Risk Quantification: Calculating Forecasted Revenue, Stockout Revenue-at-Risk, and Monthly Overstock Holding Costs using actual shelf prices.\n"
        "• Enterprise Delivery: Development of an asynchronous FastAPI REST backend (with LRU query caching) and an interactive React 19 single-page dashboard."
    )
    add_para(doc, p5)

    add_heading_2(doc, "1.4 Objective(s) of DemandVerse AI")
    p6 = (
        "The project was designed and executed to achieve ten concrete technical and operational objectives:\n"
        "1. Ingest and stratify multi-year transaction data from the Walmart M5 benchmark into memory-optimized temporal data structures.\n"
        "2. Implement, calibrate, and benchmark an Autoregressive Integrated Moving Average (ARIMA(1,1,1)) time-series forecasting model.\n"
        "3. Implement, calibrate, and benchmark an additive Meta Prophet model capturing weekly and annual seasonality.\n"
        "4. Develop an automated, hierarchical out-of-sample model selection router based on MAPE, MAE, and RMSE to assign the optimal engine per SKU.\n"
        "5. Generate robust 30-day forward-looking demand forecasts with 80% confidence bounds for all portfolio products.\n"
        "6. Compute dynamic Safety Stock (SS), Reorder Points (ROP), and recommended Purchase Order Reorder Quantities (RQ).\n"
        "7. Categorize inventory health states into Healthy Stock, Overstock, and Stockout Risk to alert procurement controllers.\n"
        "8. Implement a stateful in-silico Digital Twin simulation sandbox to model stock-state transitions without physical IoT sensors.\n"
        "9. Execute sensitivity what-if simulations across Low Demand (-10%), Baseline, and High Demand (+10%) scenarios.\n"
        "10. Quantify enterprise financial exposure, including Forecasted Gross Revenue, Stockout Revenue-at-Risk, and Monthly Holding Costs."
    )
    add_para(doc, p6)

    add_heading_2(doc, "1.5 Importance of the Application")
    p7 = (
        "DemandVerse AI provides direct, measurable business value to retail enterprises. By shifting demand forecasting from a static "
        "retrospective report to an active, variance-aware decision-support ecosystem, retailers can proactively prevent stockouts before "
        "shelves empty. Simultaneously, automated overstock detection halts unnecessary purchase orders, freeing up substantial working capital "
        "and reducing warehouse holding overhead. Connecting operational shortage quantities directly to monetary revenue-at-risk enables CFOs "
        "and supply chain directors to align procurement budgets with real financial exposure."
    )
    add_para(doc, p7)

    add_heading_2(doc, "1.6 Target Users and Audience")
    p8 = (
        "The primary stakeholders and intended user base of DemandVerse AI include:\n"
        "• Category Managers & Procurement Specialists: Responsible for issuing weekly purchase orders, tracking supplier lead times, and avoiding stockouts.\n"
        "• Inventory Controllers & Warehouse Logistics Managers: Focused on warehouse capacity utilization, buffer management, and minimizing carrying costs.\n"
        "• Chief Financial Officers (CFOs) & Commercial Directors: Focused on capital efficiency, working capital allocation, and revenue protection.\n"
        "• Retail Operations Analysts & Data Scientists: Utilizing model comparison metrics, seasonal decomposition charts, and predictive error logs."
    )
    add_para(doc, p8)

    # ==========================================================================
    # CHAPTER 2: SYSTEM REQUIREMENTS
    # ==========================================================================
    add_heading_1(doc, "CHAPTER 2\nSYSTEM REQUIREMENTS")
    
    add_heading_2(doc, "2.1 Hardware Requirements")
    add_para(doc, "The minimum and recommended hardware configurations required for running the computational machine learning pipeline, FastAPI serving backend, and React 19 dashboard are detailed in Table 2.1.")
    
    hw_headers = ["Parameter", "Minimum Specification", "Recommended Production Specification"]
    hw_data = [
        ["Processor (CPU)", "Quad-Core Intel Core i5 / AMD Ryzen 5 (2.5 GHz+)", "Octa-Core Intel Core i7/i9 or AMD Ryzen 7/9 (3.5 GHz+)"],
        ["System Memory (RAM)", "8 GB DDR4 RAM", "16 GB to 32 GB DDR4/DDR5 RAM"],
        ["Storage Space", "10 GB available SSD space", "50 GB NVMe PCIe SSD (for raw Kaggle M5 data & logs)"],
        ["Display Resolution", "1366 x 768 pixels", "1920 x 1080 (Full HD) or higher for executive dashboard"],
        ["Network Interface", "Standard Broadband (5 Mbps)", "Gigabit Ethernet / Wi-Fi 6 (100 Mbps+ for Docker image builds)"]
    ]
    col_w_hw = [Inches(1.5), Inches(2.2), Inches(2.3)]
    create_table(doc, hw_headers, hw_data, col_w_hw)

    add_heading_2(doc, "2.2 Software Requirements")
    add_para(doc, "The software environment, operating system dependencies, and runtime runtimes are specified in Table 2.2.")
    
    sw_headers = ["Software Component", "Version / Platform", "Purpose in DemandVerse AI"]
    sw_data = [
        ["Operating System", "Windows 10/11 (64-bit) or Ubuntu Linux 22.04 LTS", "Host development, deployment, and testing OS"],
        ["Python Runtime", "Python 3.11+ / Python 3.13 (64-bit)", "Core machine learning pipeline, data science, and FastAPI service"],
        ["Node.js & npm", "Node.js v20.18+ / npm v10+", "Frontend build tools, Vite development server, and React packages"],
        ["Container Runtime", "Docker Engine v24.0+ & Docker Compose v2.20+", "Multi-stage containerization of backend API and frontend Nginx"],
        ["Web Server", "Nginx Alpine 1.25+ (via Docker)", "Production reverse proxy, static asset serving, and SPA routing"],
        ["Web Browser", "Google Chrome 120+, Mozilla Firefox 120+, Edge", "Rendering React 19 single-page dashboard and Recharts visualizers"]
    ]
    col_w_sw = [Inches(1.5), Inches(2.0), Inches(2.5)]
    create_table(doc, sw_headers, sw_data, col_w_sw)

    add_heading_2(doc, "2.3 Development Tools and Frameworks")
    add_para(doc, (
        "• Integrated Development Environments (IDEs): Antigravity IDE, Visual Studio Code (v1.95+), and JupyterLab / Google Colab for interactive EDA.\n"
        "• Machine Learning Frameworks: Statsmodels v0.14.0+ (ARIMA), Meta Prophet v1.1.5+, Scikit-learn v1.3.0+, Pandas v2.1.0+, NumPy v1.26.0+.\n"
        "• API Framework: FastAPI v0.110.0+ with Uvicorn v0.28.0+ ASGI server, Pydantic v2 data validation, and Starlette middleware.\n"
        "• Frontend Framework & Libraries: React 19.2, Vite 8.3 build engine, Recharts 3.10 charting library, Lucide React icons, Axios HTTP client.\n"
        "• Testing Tools: Pytest v9.0.3+, Pytest-Cov (test coverage), Postman / Swagger OpenAPI UI for interactive REST endpoint validation."
    ))

    # ==========================================================================
    # CHAPTER 3: TECHNOLOGY STACK
    # ==========================================================================
    add_heading_1(doc, "CHAPTER 3\nTECHNOLOGY STACK")
    
    add_heading_2(doc, "3.1 Front-End Technologies")
    p_fe = (
        "1. React 19: The core user interface of DemandVerse AI is built on React 19, utilizing functional components, custom hooks, and "
        "React Context. React 19 enables high rendering performance through automatic batching and asynchronous transitions, maintaining smooth "
        "responsiveness across 9 feature-rich analytical pages.\n\n"
        "2. Vite 8.3: Used as the next-generation build tool and local development server. Vite leverages native browser ES modules (ESM) and the "
        "Rolldown bundling engine to deliver sub-second hot module replacement (HMR) and optimized production minification.\n\n"
        "3. Recharts 3.10: A declarative charting library built with React and SVG. Used throughout the dashboard to render responsive, "
        "multi-axis predictive demand curves, confidence interval bands, inventory status distributions, and revenue-at-risk comparisons.\n\n"
        "4. Tailwind CSS: A utility-first CSS framework employed to construct a clean, modern, enterprise-grade dark-themed executive portal "
        "with consistent typographic scales and flexbox layouts.\n\n"
        "5. Lucide React: A library of lightweight, clean vector icons utilized for visual status indicators, navigation bars, and stock health alerts."
    )
    add_para(doc, p_fe)

    add_heading_2(doc, "3.2 Back-End Frameworks")
    p_be = (
        "1. Python 3.11+: The primary programming language powering both offline computational workflows and online web serving, selected for its "
        "mature scientific computing ecosystem and robust asynchronous concurrency support.\n\n"
        "2. FastAPI: A high-performance, asynchronous web framework built on Starlette and Pydantic. FastAPI powers the RESTful serving layer of "
        "DemandVerse AI. Its native async capabilities allow concurrent handling of hundreds of simulation requests with minimal resource overhead.\n\n"
        "3. Uvicorn: A lightning-fast ASGI web server implementation used to run FastAPI applications in development and production container environments.\n\n"
        "4. Python `functools.lru_cache`: High-throughput in-memory caching mechanism wrapped around data loading and metric extraction utilities, "
        "reducing endpoint query response latency to sub-10 milliseconds for pre-computed multi-dimensional analytical matrices."
    )
    add_para(doc, p_be)

    add_heading_2(doc, "3.3 Database & Computational Storage")
    p_db = (
        "In enterprise retail environments where daily forecasting models are evaluated in batch schedules, heavy relational database locks "
        "introduce unnecessary latency. DemandVerse AI implements an optimized file-backed analytical data lake architecture:\n"
        "• Clean Data Storage: Preprocessed, merged, and imputed time series are stored in memory-optimized CSV structures (`clean_sales.csv`, 3.4 MB) "
        "with explicit categorical and integer downcasting to ensure rapid disk I/O.\n"
        "• Analytical Artifact Repositories: Model comparison metrics (`model_comparison.csv`), final 30-day forecasts (`final_forecasts.csv`), "
        "inventory recommendations (`inventory_recommendations.csv`), and what-if financial matrices (`financial_impact.csv`) are persisted in "
        "versioned directory trees (`demandverse_ai_outputs/`) that are loaded on-demand and cached in RAM."
    )
    add_para(doc, p_db)

    add_heading_2(doc, "3.4 Version Control & DevOps")
    p_devops = (
        "• Git & GitHub: Full distributed version control managing feature branches, semantic commits, and team collaboration across research milestones.\n"
        "• Multi-Stage Docker Builds: Dockerfile recipes optimize image sizes: the Python API container utilizes lightweight `python:3.11-slim`, while "
        "the frontend container compiles Vite assets in Node and serves static files using an ultra-compact `nginx:alpine` image.\n"
        "• Docker Compose: Orchestrates multi-container networking (`docker-compose.yml`), binding frontend (port 3000) and backend (port 8000) "
        "into an isolated virtual bridge network for seamless single-command deployment."
    )
    add_para(doc, p_devops)

    add_heading_2(doc, "3.5 APIs, Machine Learning & Numerical Libraries")
    p_ml = (
        "• Statsmodels (`statsmodels.tsa.arima.model.ARIMA`): Powers the statistical Autoregressive Integrated Moving Average (ARIMA(1,1,1)) engine.\n"
        "• Meta Prophet (`prophet.Prophet`): Decomposable Bayesian time-series library modeling piecewise linear trends and Fourier seasonality.\n"
        "• Scikit-learn: Formulates out-of-sample error evaluation, computing Mean Absolute Error (MAE), Root Mean Squared Error (RMSE), and MAPE.\n"
        "• Pandas & NumPy: Core data manipulation and matrix algebra libraries executing wide-to-long reshaping, rolling statistics, and vector math."
    )
    add_para(doc, p_ml)

print("Part 1 module loaded.")
