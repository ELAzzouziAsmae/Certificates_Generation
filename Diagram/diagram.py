from graphviz import Digraph

dot = Digraph("CertificatesAutomationApp", format="svg")
dot.attr(
    rankdir="LR",  # Garder la disposition horizontale
    size="24,16",   # Ajuste la taille de la page pour occuper toute la fenêtre
    labelloc="t", 
    label="🎓 Certificates Automation Architecture", 
    fontsize="30", 
    fontname="Helvetica-Bold"
)

# === Styles ===
style_ui = {'style': 'filled', 'fillcolor': '#AED6F1', 'shape': 'ellipse', 'fontsize': '18', 'width': '2.5'}
style_worker = {'style': 'filled', 'fillcolor': '#ABEBC6', 'shape': 'box', 'fontsize': '18', 'width': '2.5'}
style_input = {'style': 'filled', 'fillcolor': '#FAD7A0', 'shape': 'cylinder', 'fontsize': '18', 'width': '2.5'}
style_output = {'style': 'filled', 'fillcolor': '#F9E79F', 'shape': 'folder', 'fontsize': '18', 'width': '2.5'}
style_log = {'style': 'filled', 'fillcolor': '#D2B4DE', 'shape': 'folder', 'fontsize': '18', 'width': '2.5'}
style_kpi = {'style': 'filled', 'fillcolor': '#FCF3CF', 'shape': 'note', 'fontsize': '18', 'width': '2.5'}
style_external = {'style': 'filled', 'fillcolor': '#F5B7B1', 'shape': 'box', 'fontsize': '18', 'width': '2.5'}
style_packaging = {'style': 'filled', 'fillcolor': '#D7DBDD', 'shape': 'component', 'fontsize': '18', 'width': '2.5'}

# === Nodes ===
# GUI
dot.node("Main", "main.py\n(Entry Point)", **style_ui)
dot.node("LoginView", "LoginView\n(view/login_view.py)", **style_ui)
dot.node("CertView", "CertView\n(view/cert_view.py)", **style_ui)

# Worker
dot.node("CertWorker", "CertWorker\n(Worker/cert_worker.py)", **style_worker)

# Inputs
dot.node("ExcelData", "Excel Input\nFORM_*.xlsx", **style_input)
dot.node("PPTTemplate", "PPT Template\n(template_certificat.pptx)", **style_input)

# Outputs
dot.node("GeneratedPDF", "PDF Certificates\n(Output/*.pdf)", **style_output)
dot.node("LogFile", "Log File\ncertificates_generation.log", **style_log)
dot.node("KPIExports", "KPI Exports\n(kpi_export*.csv, tmp_*.html)", **style_kpi)

# External Services
dot.node("PowerPointCOM", "PowerPoint COM\n(ppt → pdf)", **style_external)
dot.node("OutlookCOM", "Outlook COM\n(send emails)", **style_external)

# Packaging & Deployment
dot.node("PyInstaller", "PyInstaller\n(version.txt, admin.manifest)", **style_packaging)
dot.node("InnoSetup", "Inno Setup\n(Setup_TrainingProject.iss)", **style_packaging)
dot.node("Executable", "Certificates Generation.exe", **style_packaging)

# === Edges ===
# GUI Flow
dot.edge("Main", "LoginView", label=" Launch App")
dot.edge("LoginView", "CertView", label="✔ Login Success")
dot.edge("CertView", "CertWorker", label="📤 Send Job (Threaded)")

# Data flow from Inputs to Worker
dot.edge("ExcelData", "CertWorker", label="📄 Read via pandas")
dot.edge("PPTTemplate", "CertWorker", label="🧩 Load template")

# Worker to Outputs/Services
dot.edge("CertWorker", "PowerPointCOM", label="🖨 COM Export to PDF")
dot.edge("PowerPointCOM", "GeneratedPDF", label="📂 Save PDF")
dot.edge("CertWorker", "OutlookCOM", label="📧 Send via Outlook COM")
dot.edge("OutlookCOM", "LogFile", label="✍ Email status")
dot.edge("CertWorker", "LogFile", label="📝 Log actions")
dot.edge("CertWorker", "KPIExports", label="📊 Generate KPIs")

# Packaging Flow
dot.edge("PyInstaller", "Executable", label="📦 Bundle to .exe")
dot.edge("InnoSetup", "Executable", label="🧾 Wrap in Installer")

# === Subgraphs ===
with dot.subgraph(name="cluster_ui") as c:
    c.attr(label="🖥 GUI Layer (PyQt5)", style="dashed", ranksep="1.5", nodesep="1.5", fontsize="20")
    c.node("Main")
    c.node("LoginView")
    c.node("CertView")

with dot.subgraph(name="cluster_worker") as c:
    c.attr(label="⚙️ Business Logic (Worker)", style="dashed", ranksep="1.5", nodesep="1.5", fontsize="20")
    c.node("CertWorker")

with dot.subgraph(name="cluster_inputs") as c:
    c.attr(label="📥 Inputs", style="dashed", ranksep="1.5", nodesep="1.5", fontsize="20")
    c.node("ExcelData")
    c.node("PPTTemplate")

with dot.subgraph(name="cluster_outputs") as c:
    c.attr(label="📤 Outputs", style="dashed", ranksep="1.5", nodesep="1.5", fontsize="20")
    c.node("GeneratedPDF")
    c.node("LogFile")
    c.node("KPIExports")

with dot.subgraph(name="cluster_external") as c:
    c.attr(label="🌐 External Services", style="dashed", ranksep="1.5", nodesep="1.5", fontsize="20")
    c.node("PowerPointCOM")
    c.node("OutlookCOM")

# Packaging & Deployment (Maintenant placé après les autres éléments, mais toujours horizontal)
with dot.subgraph(name="cluster_packaging") as c:
    c.attr(label="📦 Packaging & Deployment", style="dashed", ranksep="2", nodesep="2", fontsize="20")
    c.node("PyInstaller")
    c.node("InnoSetup")
    c.node("Executable")

# === Render ===
output_path = r"C:\Users\260001889\Desktop\certificates_architecture_clean"
dot.render(output_path, cleanup=True)
print(f"✅ Diagramme hiérarchique généré : {output_path}.svg")
