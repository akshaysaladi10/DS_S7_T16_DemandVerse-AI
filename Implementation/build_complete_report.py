import os
import sys
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import win32com.client

print("Generating DemandVerse AI University Project Report...")

OUTPUT_DIR = os.path.abspath("project_report")
DOCX_PATH = os.path.join(OUTPUT_DIR, "DemandVerse_AI_Project_Report.docx")
PDF_PATH = os.path.join(OUTPUT_DIR, "DemandVerse_AI_Project_Report.pdf")
FIGURES_DIR = os.path.join(OUTPUT_DIR, "figures")
KLH_DIR = os.path.abspath("klh_assets")

doc = docx.Document()

# Page Margins: Left 1.25" for binding, Top/Bottom/Right 1.0"
section = doc.sections[0]
section.page_width = Inches(8.27)
section.page_height = Inches(11.69)
section.top_margin = Inches(1.0)
section.bottom_margin = Inches(1.0)
section.left_margin = Inches(1.25)
section.right_margin = Inches(1.0)
section.header_distance = Inches(0.5)
section.footer_distance = Inches(0.5)

def set_font(run, name='Times New Roman', size=12, bold=False, italic=False, color=None):
    run.font.name = name
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    if color:
        run.font.color.rgb = color

# University Header/Footer Configuration
section.different_first_page_header_footer = True
footer = section.footer
p_f = footer.paragraphs[0]
p_f.alignment = WD_ALIGN_PARAGRAPH.RIGHT
r_ftxt = p_f.add_run("DEPT. OF CSE, KLEF, Bowrampet, Hyderabad          Page ")
set_font(r_ftxt, name='Times New Roman', size=9.5, italic=False, color=RGBColor(100, 100, 100))
fldSimple = parse_xml(r'<w:fldSimple %s w:instr="PAGE"/>' % nsdecls('w'))
p_f._p.append(fldSimple)

def add_para(doc, text="", align=WD_ALIGN_PARAGRAPH.JUSTIFY,
             space_before=0, space_after=6, line_spacing=1.5, bold=False, italic=False, size=12):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = line_spacing
    if text:
        run = p.add_run(text)
        set_font(run, size=size, bold=bold, italic=italic)
    return p

def add_heading_1(doc, text):
    doc.add_page_break()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(16)
    p.paragraph_format.keep_with_next = True
    run = p.add_run(text)
    set_font(run, size=16, bold=True, color=RGBColor(0, 32, 96))
    return p

def add_heading_2(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.keep_with_next = True
    run = p.add_run(text)
    set_font(run, size=13.5, bold=True, color=RGBColor(15, 30, 60))
    return p

def add_heading_3(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    run = p.add_run(text)
    set_font(run, size=12, bold=True, italic=True)
    return p

def add_image_box(doc, img_name, caption, width=Inches(5.8)):
    img_path = os.path.join(FIGURES_DIR, img_name)
    if os.path.exists(img_path):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(10)
        p_img.paragraph_format.space_after = Pt(4)
        run_img = p_img.add_run()
        run_img.add_picture(img_path, width=width)
        
        p_cap = doc.add_paragraph()
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.paragraph_format.space_before = Pt(2)
        p_cap.paragraph_format.space_after = Pt(12)
        run_cap = p_cap.add_run(caption)
        set_font(run_cap, size=10, bold=True, italic=True, color=RGBColor(70, 70, 70))

def set_cell_background(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    tcPr.append(shd)

def create_table(doc, headers, data, col_widths, align_cols=None):
    table = doc.add_table(rows=len(data) + 1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # Headers
    hdr_cells = table.rows[0].cells
    for i, title in enumerate(headers):
        hdr_cells[i].text = title
        set_cell_background(hdr_cells[i], "1F497D")
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(4)
        for run in p.runs:
            set_font(run, size=9.5, bold=True, color=RGBColor(255, 255, 255))
            
    # Data Rows
    for r_idx, row_data in enumerate(data):
        row_cells = table.rows[r_idx + 1].cells
        bg_col = "F2F5F9" if r_idx % 2 == 1 else "FFFFFF"
        for c_idx, val in enumerate(row_data):
            row_cells[c_idx].text = str(val)
            set_cell_background(row_cells[c_idx], bg_col)
            p = row_cells[c_idx].paragraphs[0]
            p.paragraph_format.space_before = Pt(3)
            p.paragraph_format.space_after = Pt(3)
            if align_cols and c_idx in align_cols:
                p.alignment = align_cols[c_idx]
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            for run in p.runs:
                set_font(run, size=9.0)
                
    for row in table.rows:
        for idx, width in enumerate(col_widths):
            row.cells[idx].width = width
            
    doc.add_paragraph().paragraph_format.space_after = Pt(6)
    return table

# ==============================================================================
# 1. TITLE / COVER PAGE
# ==============================================================================
print("Writing Title Page...")
p_t = doc.add_paragraph()
p_t.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_t.paragraph_format.space_before = Pt(18)
p_t.paragraph_format.space_after = Pt(6)
run_t = p_t.add_run("DEMANDVERSE AI")
set_font(run_t, size=24, bold=True, color=RGBColor(0, 32, 96))

p_sub = doc.add_paragraph()
p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_sub.paragraph_format.space_before = Pt(0)
p_sub.paragraph_format.space_after = Pt(18)
run_sub = p_sub.add_run("A Retail Intelligence Ecosystem for Predictive Demand Forecasting, Inventory Optimization, and Revenue Growth Analytics")
set_font(run_sub, size=13, italic=True, bold=True, color=RGBColor(40, 40, 40))

p_subm = doc.add_paragraph()
p_subm.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_subm.paragraph_format.space_before = Pt(12)
p_subm.paragraph_format.space_after = Pt(6)
run_subm = p_subm.add_run("A Project Report Submitted in partial fulfillment of the requirements\nfor the award of the degree of")
set_font(run_subm, size=12, italic=True)

p_deg = doc.add_paragraph()
p_deg.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_deg.paragraph_format.space_before = Pt(4)
p_deg.paragraph_format.space_after = Pt(4)
run_deg = p_deg.add_run("BACHELOR OF TECHNOLOGY\nIN\nDEPARTMENT OF COMPUTER SCIENCE AND ENGINEERING")
set_font(run_deg, size=13.5, bold=True, color=RGBColor(0, 32, 96))

p_by = doc.add_paragraph()
p_by.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_by.paragraph_format.space_before = Pt(12)
p_by.paragraph_format.space_after = Pt(6)
run_by = p_by.add_run("By")
set_font(run_by, size=12, bold=True)

# Authors Table
tbl_auth = doc.add_table(rows=3, cols=2)
tbl_auth.alignment = WD_TABLE_ALIGNMENT.CENTER
auth_data = [
    ("Rachakonda Lakshmi Sai Giridhar", ": 2420030171"),
    ("Saladi Akshay Veer", ": 2420030579"),
    ("Akhil Varma Telka", ": 2420030141")
]
for idx, (name, rno) in enumerate(auth_data):
    r = tbl_auth.rows[idx]
    r.cells[0].text = name
    r.cells[1].text = rno
    r.cells[0].width = Inches(3.2)
    r.cells[1].width = Inches(2.0)
    for c in r.cells:
        p = c.paragraphs[0]
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        for run in p.runs:
            set_font(run, size=11.5, bold=True)

p_gui = doc.add_paragraph()
p_gui.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_gui.paragraph_format.space_before = Pt(16)
p_gui.paragraph_format.space_after = Pt(2)
run_gui = p_gui.add_run("Under the Esteemed Guidance of")
set_font(run_gui, size=11.5, italic=True)

p_gn = doc.add_paragraph()
p_gn.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_gn.paragraph_format.space_before = Pt(2)
p_gn.paragraph_format.space_after = Pt(2)
run_gn = p_gn.add_run("B. Prathusha\nAssistant Professor (Faculty Guide)\nDepartment of Computer Science and Engineering")
set_font(run_gn, size=12, bold=True)

p_ay = doc.add_paragraph()
p_ay.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_ay.paragraph_format.space_before = Pt(6)
p_ay.paragraph_format.space_after = Pt(12)
run_ay = p_ay.add_run("Academic Year: 2026 - 2027  |  Semester: 7 (Course: 24AD3107)")
set_font(run_ay, size=11, bold=True)

# University Logo
logo_path = os.path.join(KLH_DIR, "image2.jpeg")
if os.path.exists(logo_path):
    p_logo = doc.add_paragraph()
    p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_logo.paragraph_format.space_before = Pt(6)
    p_logo.paragraph_format.space_after = Pt(4)
    run_l = p_logo.add_run()
    run_l.add_picture(logo_path, width=Inches(2.5))

p_uni = doc.add_paragraph()
p_uni.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_uni.paragraph_format.space_before = Pt(4)
p_uni.paragraph_format.space_after = Pt(0)
run_u1 = p_uni.add_run("Koneru Lakshmaiah Education Foundation\n")
set_font(run_u1, size=13, bold=True, color=RGBColor(180, 0, 0)) # Red
run_u2 = p_uni.add_run("(Deemed to be University estd. u/s. 3 of the UGC Act, 1956)\nOff-Campus: Bachupally-Gandimaisamma Road, Bowrampet, Hyderabad, Telangana - 500 043.\nPhone: 7815926816 | www.klh.edu.in\nDEPARTMENT OF COMPUTER SCIENCE AND ENGINEERING")
set_font(run_u2, size=9.5, bold=True, color=RGBColor(50, 50, 50))

# ==============================================================================
# 2. DECLARATION PAGE
# ==============================================================================
print("Writing Declaration Page...")
doc.add_page_break()
p_dec_h = doc.add_paragraph()
p_dec_h.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_dec_h.paragraph_format.space_before = Pt(18)
p_dec_h.paragraph_format.space_after = Pt(24)
run_dh = p_dec_h.add_run("DECLARATION")
set_font(run_dh, size=16, bold=True, color=RGBColor(0, 32, 96))

dec_text = (
    "The Project Report entitled “DEMANDVERSE AI: A Retail Intelligence Ecosystem for Predictive Demand "
    "Forecasting, Inventory Optimization, and Revenue Growth Analytics” is a record of Bonafide work of "
    "Rachakonda Lakshmi Sai Giridhar – 2420030171, Saladi Akshay Veer – 2420030579, and Akhil Varma Telka – 2420030141 "
    "submitted in partial fulfillment for the award of Bachelor of Technology in Computer Science and Engineering to the "
    "Koneru Lakshmaiah Education Foundation (KL University). The results embodied in this report have not been copied "
    "from any other departments / University / Institute."
)
add_para(doc, dec_text, space_before=12, space_after=36, line_spacing=1.5)

# Signature blocks
p_sig = doc.add_paragraph()
p_sig.alignment = WD_ALIGN_PARAGRAPH.RIGHT
p_sig.paragraph_format.space_before = Pt(40)
p_sig.paragraph_format.space_after = Pt(4)
run_s1 = p_sig.add_run("Rachakonda Lakshmi Sai Giridhar  (2420030171)\n")
set_font(run_s1, size=11, bold=True)
run_s2 = p_sig.add_run("Saladi Akshay Veer  (2420030579)\n")
set_font(run_s2, size=11, bold=True)
run_s3 = p_sig.add_run("Akhil Varma Telka  (2420030141)\n")
set_font(run_s3, size=11, bold=True)

# ==============================================================================
# 3. CERTIFICATE PAGE
# ==============================================================================
print("Writing Certificate Page...")
doc.add_page_break()
p_cert_h = doc.add_paragraph()
p_cert_h.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_cert_h.paragraph_format.space_before = Pt(18)
p_cert_h.paragraph_format.space_after = Pt(24)
run_ch = p_cert_h.add_run("CERTIFICATE")
set_font(run_ch, size=16, bold=True, color=RGBColor(0, 32, 96))

cert_text = (
    "This is to certify that the project report entitled “DEMANDVERSE AI: A Retail Intelligence Ecosystem "
    "for Predictive Demand Forecasting, Inventory Optimization, and Revenue Growth Analytics” is a bonafide "
    "work done and submitted by Rachakonda Lakshmi Sai Giridhar – 2420030171, Saladi Akshay Veer – 2420030579, "
    "and Akhil Varma Telka – 2420030141 in partial fulfillment of the requirements for the award of the degree of "
    "BACHELOR OF TECHNOLOGY in Department of Computer Science and Engineering, Koneru Lakshmaiah Education Foundation "
    "(K L Deemed to be University), Bowrampet, Hyderabad, during the academic year 2026 - 2027."
)
add_para(doc, cert_text, space_before=12, space_after=60, line_spacing=1.5)

tbl_sigs = doc.add_table(rows=2, cols=3)
tbl_sigs.alignment = WD_TABLE_ALIGNMENT.CENTER
tbl_sigs.rows[0].cells[0].text = "______________________\nSignature of the Guide\n(B. Prathusha)"
tbl_sigs.rows[0].cells[1].text = "______________________\nCourse Coordinator\n(24AD3107)"
tbl_sigs.rows[0].cells[2].text = "______________________\nHead of the Department\n(Department of CSE)"
for c in tbl_sigs.rows[0].cells:
    p = c.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in p.runs:
        set_font(run, size=10, bold=True)

# ==============================================================================
# 4. ACKNOWLEDGEMENT PAGE
# ==============================================================================
print("Writing Acknowledgement Page...")
doc.add_page_break()
p_ack_h = doc.add_paragraph()
p_ack_h.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_ack_h.paragraph_format.space_before = Pt(18)
p_ack_h.paragraph_format.space_after = Pt(20)
run_ah = p_ack_h.add_run("ACKNOWLEDGEMENT")
set_font(run_ah, size=16, bold=True, color=RGBColor(0, 32, 96))

ack_p1 = (
    "The success and timely culmination of this project would not have been possible without the invaluable guidance, "
    "consistent encouragement, and intellectual support rendered by numerous mentors and faculty members. We wish to "
    "express our deepest gratitude to all those who have facilitated our research and engineering journey."
)
add_para(doc, ack_p1)

ack_p2 = (
    "We convey our heartfelt appreciation to our Course Coordinator, and our esteemed Project Guide, B. Prathusha, "
    "Assistant Professor, Department of Computer Science and Engineering, whose insightful critiques, technical guidance, "
    "and rigorous academic standards were instrumental throughout the lifecycle of DemandVerse AI."
)
add_para(doc, ack_p2)

ack_p3 = (
    "We express our sincere gratitude to the Head of the Department, Department of Computer Science and Engineering, "
    "and the leadership of Koneru Lakshmaiah Education Foundation (KL University, Bowrampet Campus) for providing state-of-the-art "
    "computational laboratories, high-performance computing facilities, and an intellectually vibrant environment that enabled "
    "the end-to-end realization of this Data Science Project."
)
add_para(doc, ack_p3)

ack_p4 = (
    "Finally, we extend our sincere thanks to all teaching and non-teaching staff members, laboratory technicians, our beloved "
    "parents, and our classmates for their steadfast encouragement, moral support, and cooperation during our undergraduate studies."
)
add_para(doc, ack_p4, space_after=30)

p_asig = doc.add_paragraph()
p_asig.alignment = WD_ALIGN_PARAGRAPH.RIGHT
run_as1 = p_asig.add_run("Rachakonda Lakshmi Sai Giridhar (2420030171)\nSaladi Akshay Veer (2420030579)\nAkhil Varma Telka (2420030141)\n")
set_font(run_as1, size=11, bold=True)

# ==============================================================================
# 5. ABSTRACT PAGE
# ==============================================================================
print("Writing Abstract Page...")
doc.add_page_break()
p_abs_h = doc.add_paragraph()
p_abs_h.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_abs_h.paragraph_format.space_before = Pt(18)
p_abs_h.paragraph_format.space_after = Pt(18)
run_abh = p_abs_h.add_run("ABSTRACT")
set_font(run_abh, size=16, bold=True, color=RGBColor(0, 32, 96))

abs_text_1 = (
    "Retail enterprise supply chains operate under severe, continuous demand uncertainty. Sub-optimal inventory "
    "replenishment precipitates acute operational and balance-sheet distortions, oscillating between destructive stockouts "
    "— resulting in unrecoverable lost sales, compromised customer goodwill, and retailer penalties — and chronic "
    "overstocking, which traps working capital in non-productive inventory, escalates warehouse holding overhead, and drives "
    "aggressive margin-eroding markdowns. Traditional Enterprise Resource Planning (ERP) forecasting systems typically operate "
    "in computational silos, disconnecting time-series mathematical predictions from downstream buffer sizing, multi-scenario "
    "stress-testing, and executive financial quantification."
)
add_para(doc, abs_text_1)

abs_text_2 = (
    "To address these critical operational deficiencies, this project presents DemandVerse AI: an end-to-end, adaptive retail "
    "intelligence ecosystem that unites statistical and additive machine learning forecasting with dynamic inventory control "
    "and balance-sheet risk analytics. The framework was developed, calibrated, and rigorously validated on 28,695 daily "
    "sales records from the Walmart M5 Forecasting Accuracy benchmark spanning 1,913 historical days across California retail supercenters. "
    "DemandVerse AI implements parallel forecasting engines using Autoregressive Integrated Moving Average (ARIMA(1,1,1)) and Meta Prophet "
    "generalized additive models, governed by an automated, hierarchical out-of-sample tie-breaking model router evaluated across "
    "Mean Absolute Error (MAE), Root Mean Squared Error (RMSE), and Mean Absolute Percentage Error (MAPE)."
)
add_para(doc, abs_text_2)

abs_text_3 = (
    "Empirical benchmarking across an 80/20 chronological holdout split established that the adaptive model selector routed 11 SKUs (73.3%) "
    "characterized by multi-frequency calendar seasonality to Prophet and 4 SKUs (26.7%) exhibiting stationary mean-reverting autoregression "
    "to ARIMA. The resulting adaptive hybrid ensemble achieved a portfolio-average MAE of 5.0487, RMSE of 6.2916, and MAPE of 52.22%, "
    "outperforming isolated monolithic ARIMA (MAE 5.5681, RMSE 6.9153, MAPE 57.30%) and Prophet (MAE 5.2357, RMSE 6.4509, MAPE 57.16%) "
    "across all dimensions. Forward-looking 30-day forecast curves dynamically parameterize a service-level safety stock engine (Z = 1.645, "
    "L = 7 days) and Reorder Points (ROP). Inventory state representation via an in-silico Digital Twin simulation sandbox classified the "
    "active portfolio into 60.0% Healthy Stock, 26.7% Overstock, and 13.3% Stockout Risk. Multi-scenario sensitivity analysis (-10%, Baseline, "
    "+10%) revealed that procurement requirements shift from 2,684.32 to 3,592.58 units, modulating stockout revenue-at-risk between $4,588.66 "
    "and $6,133.76, while surplus carrying costs range from $5.69 to $12.93 per month. The ecosystem is deployed via an asynchronous "
    "FastAPI backend (with LRU caching) and an interactive React 19 single-page dashboard, verified through 65 passing automated software tests "
    "and a comprehensive 24-point end-to-end verification audit."
)
add_para(doc, abs_text_3)

# ==============================================================================
# 6. INDEX TABLE (TABLE OF CONTENTS)
# ==============================================================================
print("Writing Index Table...")
doc.add_page_break()
p_idx_h = doc.add_paragraph()
p_idx_h.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_idx_h.paragraph_format.space_before = Pt(14)
p_idx_h.paragraph_format.space_after = Pt(16)
run_ih = p_idx_h.add_run("INDEX")
set_font(run_ih, size=16, bold=True, color=RGBColor(0, 32, 96))

idx_headers = ["S. No.", "Chapters", "Topics", "Page.no"]
idx_data = [
    ["", "Declaration", "Declaration of Bonafide Work", "ii"],
    ["", "Certificate", "Bonafide Certificate from Department", "iii"],
    ["", "Acknowledgement", "Academic Acknowledgement", "iv"],
    ["", "Abstract", "Executive and Technical Abstract", "v"],
    ["1", "Introduction", "1.1 Background of the Project\n1.2 Problem Statement\n1.3 Scope of the Project\n1.4 Objectives of DemandVerse AI\n1.5 Importance of the Application\n1.6 Target Users and Audience", "9 - 12"],
    ["2", "System Requirements", "2.1 Hardware Requirements\n2.2 Software Requirements\n2.3 Development Tools and Frameworks", "13 - 14"],
    ["3", "Technology Stack", "3.1 Front-End Technologies (React 19, Vite, Recharts)\n3.2 Back-End Frameworks (FastAPI, Uvicorn, LRU Cache)\n3.3 Database & Computational Storage (Walmart M5 Caching)\n3.4 Version Control & DevOps (Git, Docker, Compose, Nginx)\n3.5 APIs, Machine Learning & Numerical Libraries", "15 - 17"],
    ["4", "System Architecture", "4.1 High-Level Architecture Diagram\n4.2 Description of Each Architectural Layer\n4.3 Deployment Architecture & Container Orchestration", "18 - 20"],
    ["5", "Design", "5.1 System Wireframes, UI Mockups & Data Flow Diagrams\n5.2 Entity-Relationship (ER) Diagram & Schema Architecture", "21 - 23"],
    ["6", "Implementation", "6.1 Module-Wise Implementation Architecture\n6.2 Front-End Logic (Component Tree, Context, Recharts)\n6.3 Back-End Logic (FastAPI Endpoints, Route Controllers)\n6.4 Data Ingestion, Cleaning & Chronological 80/20 Splitting\n6.5 RESTful API Integration & Endpoint Specifications\n6.6 Security, Numerical Sanitization & CORS Configuration", "24 - 26"],
    ["7", "Features & Novelties", "7.1 Comprehensive Feature Inventory\n7.2 Novelty 1: Adaptive ARIMA-Prophet Model Selector\n7.3 Novelty 2: Forecast-Based Inventory Optimization\n7.4 Novelty 3: What-If & Financial Impact Quantification", "27 - 33"],
    ["8", "Testing & Validation", "8.1 API & Backend Testing with Pytest (14/14 Passed)\n8.2 Machine Learning & Financial Test Suite (51/51 Passed)\n8.3 Automated End-to-End System Audit (e2e_validator.py)\n8.4 Test Cases, Assertions & Verification Results", "34 - 36"],
    ["9", "Deployment", "9.1 Multi-Stage Containerized Deployment Process\n9.2 Environment Configuration & Dependency Management\n9.3 Production Hosting, Nginx Reverse Proxy & Routing", "37 - 38"],
    ["10", "Challenges & Limitations", "10.1 Technical & Operational Challenges Encountered\n10.2 Applied Engineering Solutions & Workarounds\n10.3 Current System Limitations & Boundary Conditions", "39 - 40"],
    ["11", "Future Enhancements", "11.1 Deep Learning Architectures (TFT, DeepAR)\n11.2 Live IoT Warehouse Synchronization & RFID Telemetry\n11.3 Multi-Echelon Supply Chain Network Propagation\n11.4 Reinforcement Learning Dynamic Replenishment", "41"],
    ["12", "Conclusion", "12.1 Project Summary\n12.2 Key Achievements & Quantitative Milestones\n12.3 Technical Skills, Competencies & Insights Acquired", "42 - 43"],
    ["13", "References", "Academic Papers, Books, Standards & Documentation", "44 - 45"],
    ["14", "Appendices", "14.1 Dashboard Visualizations & Screenshots\n14.2 Core Module Source Code Snippets\n14.3 Complete Installation, Setup & Execution Guide\n14.4 End-User Operational Manual & Decision-Maker Guide", "46 - 48"]
]

col_w = [Inches(0.6), Inches(1.5), Inches(3.2), Inches(0.7)]
align_c = {0: WD_ALIGN_PARAGRAPH.CENTER, 3: WD_ALIGN_PARAGRAPH.CENTER}
create_table(doc, idx_headers, idx_data, col_w, align_c)

print("Index table written.")

# Import chapter writers
from report_sections_part1 import write_chapters_1_to_3
from report_sections_part2 import write_chapters_4_to_6
from report_sections_part3 import write_chapters_7_to_9
from report_sections_part4 import write_chapters_10_to_14

print("Assembling Chapters 1 to 3...")
write_chapters_1_to_3(doc, create_table)

print("Assembling Chapters 4 to 6...")
write_chapters_4_to_6(doc, create_table, add_image_box)

print("Assembling Chapters 7 to 9...")
write_chapters_7_to_9(doc, create_table, add_image_box)

print("Assembling Chapters 10 to 14...")
write_chapters_10_to_14(doc, create_table, add_image_box)

print(f"Saving Word Document to {DOCX_PATH}...")
doc.save(DOCX_PATH)
print("Word document saved successfully!")

# Export to PDF via Word COM Automation
print(f"Converting {DOCX_PATH} to {PDF_PATH} via Microsoft Word COM...")
try:
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    doc_word = word.Documents.Open(DOCX_PATH)
    doc_word.SaveAs(PDF_PATH, FileFormat=17) # 17 = wdFormatPDF
    doc_word.Close()
    word.Quit()
    print(f"PDF exported successfully! File size: {os.path.getsize(PDF_PATH):,} bytes")
except Exception as e:
    print(f"Error during Word COM conversion: {e}")
    sys.exit(1)

print("All tasks completed successfully!")
