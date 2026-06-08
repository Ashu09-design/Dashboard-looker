from fpdf import FPDF

class ShowcasePDF(FPDF):
    def header(self):
        # Header Background
        self.set_fill_color(99, 102, 241)  # Primary color (#6366f1)
        self.rect(0, 0, 210, 45, 'F')
        
        self.set_y(15)
        self.set_font('Helvetica', 'B', 24)
        self.set_text_color(255, 255, 255)
        self.cell(0, 10, 'Digital Enablement Dashboard', 0, 1, 'C')
        
        self.set_font('Helvetica', 'B', 14)
        self.cell(0, 10, 'Project Showcase & Technical Documentation', 0, 1, 'C')
        self.ln(10)

    def footer(self):
        self.set_y(-15)
        self.set_font('Helvetica', 'I', 8)
        self.set_text_color(128, 128, 128)
        self.cell(0, 10, f'Page {self.page_no()} | Digital Enablement Dashboard Showcase', 0, 0, 'C')

    def section_title(self, label):
        self.ln(5)
        self.set_font('Helvetica', 'B', 16)
        self.set_text_color(79, 70, 229) # #4f46e5
        self.cell(0, 10, label, 0, 1, 'L')
        # Line under title
        self.set_draw_color(226, 232, 240)
        self.line(self.get_x(), self.get_y(), 200, self.get_y())
        self.ln(4)

    def sub_section_title(self, label):
        self.set_font('Helvetica', 'B', 13)
        self.set_text_color(15, 23, 42)
        self.cell(0, 8, label, 0, 1, 'L')

    def body_text(self, text):
        self.set_font('Helvetica', '', 11)
        self.set_text_color(71, 85, 105)
        self.multi_cell(0, 6, text)
        self.ln(2)

    def bullet_point(self, title, desc=None):
        self.set_font('Helvetica', 'B', 11)
        self.set_text_color(30, 41, 59)
        # Use a bullet character that Helvetica supports
        if desc:
            self.write(6, f"- {title}: ")
            self.set_font('Helvetica', '', 11)
            self.set_text_color(71, 85, 105)
            self.write(6, f"{desc}\n")
        else:
            self.write(6, f"- {title}\n")
        self.ln(1)

def create_pdf():
    pdf = ShowcasePDF()
    pdf.set_auto_page_break(auto=True, margin=20)
    pdf.add_page()
    
    # 1. Project Overview
    pdf.section_title("Project Overview")
    pdf.body_text("A high-performance, data-driven analytics platform designed to provide real-time visibility into team delivery, project health, and executive-level KPIs. This dashboard transforms fragmented Excel-based tracking into a centralized, interactive visual hub.")
    pdf.body_text("The Digital Enablement Dashboard (also known as Dashboard Looker) is a sophisticated enterprise tool that bridges the gap between manual spreadsheets and automated reporting. It provides two primary layers of visibility:")
    
    pdf.bullet_point("Operational Layer", "Real-time standup tracking, bandwidth monitoring, and blocker management for delivery teams.")
    pdf.bullet_point("Executive Layer", "Secure, JWT-protected analytics covering financial health (SOW/PO), quality metrics (FTR), and resource capacity (Skill Matrix/KPIs).")

    # 2. Core Features
    pdf.section_title("Core Features")
    
    pdf.sub_section_title("Operational Excellence")
    pdf.bullet_point("Smart Standup Panels", "Automatically extracts 'Yesterday' and 'Today' tasks, categorizing them into Deliverables vs. Operational work.")
    pdf.bullet_point("Visual Bandwidth Tracking", "Real-time percentage bars for team capacity, identifying bottlenecks or available resources at a glance.")
    pdf.bullet_point("Live Blocker Management", "Centralized tracking for QBR dates and project-critical dependencies.")
    pdf.bullet_point("Presentation Mode", "A specialized 'TV-optimized' UI for team standups and screen sharing.")
    
    pdf.ln(2)
    pdf.sub_section_title("Executive Intelligence (Manager Portal)")
    pdf.bullet_point("Secure Access", "Role-based access control via JWT (JSON Web Tokens).")
    pdf.bullet_point("Project Health Matrix", "Automated scoring system that computes 'Health' by cross-referencing multiple data sources.")
    pdf.bullet_point("Financial & SOW Tracking", "Interactive doughnut charts and summaries for PO status and project-wise financials.")
    pdf.bullet_point("KPI Scorecards", "Clickable cards for key performance indicators with deep-dive trend analysis.")
    pdf.bullet_point("FTE & Resource Trends", "Time-series visualization of team growth and allocation.")

    pdf.ln(2)
    pdf.sub_section_title("Connectivity & Integration")
    pdf.bullet_point("Multi-Source Support", "Seamlessly connects to Local Excel files, Google Sheets (live export), and SharePoint/OneDrive links.")
    pdf.bullet_point("Zero-Delay Refresh", "Built-in polling mechanism keeps data fresh without requiring manual page reloads.")

    # 3. Technical Stack
    pdf.section_title("Technical Stack")
    
    pdf.set_font('Helvetica', 'B', 12)
    pdf.set_text_color(15, 23, 42)
    pdf.cell(0, 8, "Frontend", 0, 1)
    pdf.bullet_point("Vanilla JS (ES6+)", "High-performance, zero-framework architecture for lightning-fast load times.")
    pdf.bullet_point("Modern CSS3", "Custom grid/flexbox layouts with CSS variables for dynamic theming.")
    pdf.bullet_point("Chart.js", "Interactive, high-fidelity data visualizations.")

    pdf.ln(2)
    pdf.set_font('Helvetica', 'B', 12)
    pdf.set_text_color(15, 23, 42)
    pdf.cell(0, 8, "Backend", 0, 1)
    pdf.bullet_point("Node.js & Express", "Scalable REST API architecture.")
    pdf.bullet_point("JWT Authentication", "Secure session management for executive features.")
    pdf.bullet_point("Multer", "Robust file handling for Excel uploads.")

    pdf.ln(2)
    pdf.set_font('Helvetica', 'B', 12)
    pdf.set_text_color(15, 23, 42)
    pdf.cell(0, 8, "Data Engineering (The 'Secret Sauce')", 0, 1)
    pdf.bullet_point("Python Normalizer", "A specialized bridge using Pandas and OpenPyXL to handle complex Excel logic.")
    pdf.bullet_point("Smart Header Detection", "Scans up to 30 rows deep to find headers, overcoming 'messy' human-edited spreadsheets.")
    pdf.bullet_point("Fuzzy Matching", "Uses string similarity algorithms to map column names (e.g., matching 'PM' to 'Project Manager') automatically.")
    pdf.bullet_point("Hybrid Parsing", "Seamless fallback to Node.js parsers if Python environments are unavailable.")

    # 4. Architecture
    pdf.section_title("Architecture")
    pdf.set_font('Helvetica', '', 11)
    pdf.set_text_color(71, 85, 105)
    pdf.multi_cell(0, 7, "The system architecture is divided into three distinct layers ensuring data integrity and user security:")
    pdf.bullet_point("Data Sources", "Local Excel, Google Sheets, SharePoint / OneDrive")
    pdf.bullet_point("Backend (Node.js)", "Express Server, JWT Auth Middleware, execDataService")
    pdf.bullet_point("Parsing Engine", "Python Normalizer (Smart Parse) with Node.js Fallback Parser")
    pdf.bullet_point("Frontend (Vanilla JS)", "Operational Dashboard, Executive Portal, Data Health Diagnostics")

    # 5. Technical Highlights
    pdf.section_title("Key Technical Highlights (Showcase)")
    
    pdf.sub_section_title("Intelligent Parsing Engine")
    pdf.body_text("Traditional dashboards break when a user adds a row or renames a column in the source Excel. This project implements a Python-powered normalizer that uses fuzzy logic and deep-scanning to ensure 100% uptime regardless of spreadsheet formatting changes.")

    pdf.sub_section_title("Lightweight Performance")
    pdf.body_text("By avoiding heavy frameworks like React or Angular, the dashboard maintains a near-zero footprint, making it ideal for low-bandwidth environments and instant-loading presentation monitors.")

    pdf.sub_section_title("Enterprise Security")
    pdf.body_text("Features a complete authentication flow with encrypted password hashing (Bcrypt) and secure token-based sessions (JWT), ensuring sensitive financial and resource data is protected.")

    pdf.sub_section_title("Data Health Diagnostics")
    pdf.body_text("Includes a dedicated /api/exec/data-health endpoint that provides a 'transparency report' on data quality, identifying missing columns or parsing errors before they impact decision-making.")

    pdf.output("Dashboard_Project_Showcase.pdf")
    print("PDF generated successfully: Dashboard_Project_Showcase.pdf")

if __name__ == "__main__":
    create_pdf()
